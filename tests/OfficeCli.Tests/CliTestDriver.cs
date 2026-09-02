// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.Diagnostics;
using OfficeCli.Core;

namespace OfficeCli.Tests;

internal sealed record CliResult(int ExitCode, string StdOut, string StdErr);

internal static class CliTestDriver
{
    public static async Task<CliResult> RunAsync(string workingDirectory, params string[] arguments)
    {
        var dotnet = Environment.GetEnvironmentVariable("OFFICECLI_TEST_DOTNET") ?? "dotnet";
        var cliAssembly = typeof(ValidationProfiles).Assembly.Location;
        var startInfo = new ProcessStartInfo
        {
            FileName = dotnet,
            WorkingDirectory = workingDirectory,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false,
            CreateNoWindow = true,
        };
        startInfo.ArgumentList.Add(cliAssembly);
        foreach (var argument in arguments) startInfo.ArgumentList.Add(argument);
        startInfo.Environment["OFFICECLI_NO_AUTO_RESIDENT"] = "1";

        using var process = Process.Start(startInfo)
            ?? throw new InvalidOperationException("Could not start the OfficeCLI test process.");
        var stdout = process.StandardOutput.ReadToEndAsync();
        var stderr = process.StandardError.ReadToEndAsync();
        await process.WaitForExitAsync();
        return new CliResult(process.ExitCode, await stdout, await stderr);
    }
}

internal sealed class TestWorkspace : IDisposable
{
    public TestWorkspace()
    {
        DirectoryPath = Path.Combine(
            Path.GetTempPath(),
            "officecli-tests",
            Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(DirectoryPath);
    }

    public string DirectoryPath { get; }

    public void Dispose()
    {
        if (Directory.Exists(DirectoryPath))
            Directory.Delete(DirectoryPath, recursive: true);
    }
}
