// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using OfficeCli.Core;
using Xunit;

namespace OfficeCli.Tests;

public sealed class PinnedForkTests
{
    [Fact]
    public void PinnedForkDoesNotConsumePendingUpdates()
    {
        Assert.True(UpdateChecker.IsPinnedFork);
        using var workspace = new TestWorkspace();
        var executable = Path.Combine(workspace.DirectoryPath, "officecli.exe");
        var pending = executable + ".update";
        byte[] original = "synthetic pinned executable"u8.ToArray();
        byte[] staged = "synthetic pending update, never executable"u8.ToArray();
        File.WriteAllBytes(executable, original);
        File.WriteAllBytes(pending, staged);

        Assert.False(UpdateChecker.TryApplyPendingUpdate(executable));
        Assert.Equal(original, File.ReadAllBytes(executable));
        Assert.True(File.Exists(pending), "Pinned builds must not consume or apply a staged update.");
        Assert.Equal(staged, File.ReadAllBytes(pending));
    }

    [Fact]
    public async Task UpdateCommandsCannotEnableSelfReplacement()
    {
        using var workspace = new TestWorkspace();
        var read = await CliTestDriver.RunAsync(workspace.DirectoryPath, "config", "autoUpdate");
        Assert.Equal(0, read.ExitCode);
        Assert.Equal("false", read.StdOut.Trim());

        var enable = await CliTestDriver.RunAsync(workspace.DirectoryPath, "config", "autoUpdate", "true");
        Assert.Equal(1, enable.ExitCode);
        Assert.Contains("pinned PaperAI", enable.StdErr);

        var disable = await CliTestDriver.RunAsync(workspace.DirectoryPath, "config", "autoUpdate", "false");
        Assert.Equal(0, disable.ExitCode);
        Assert.Equal("autoUpdate = false", disable.StdOut.Trim());

        var refresh = await CliTestDriver.RunAsync(workspace.DirectoryPath, "__update-check__");
        Assert.Equal(0, refresh.ExitCode);
        Assert.Empty(refresh.StdOut);
        Assert.Empty(refresh.StdErr);
    }
}
