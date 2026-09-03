// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.Buffers.Binary;
using OfficeCli.Core;
using Xunit;
using static OfficeCli.Tests.MathTypeTestData;

namespace OfficeCli.Tests;

public class MathTypeStorageTests
{
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void RootEquationStreamDoesNotDependOnPhysicalDirectoryOrder(bool rootFirst)
    {
        var bytes = NestedEquations(rootFirst);
        Assert.Equal("x", MathTypeReader.ReadOle(bytes).Math.InnerText);
    }

    [Fact]
    public void RootEquationCanBeReachedThroughASiblingPointer()
    {
        var bytes = NestedEquations();
        SetPointer(bytes, 0, 76, 1);
        SetPointer(bytes, 1, 72, 3);
        SetPointer(bytes, 3, 68, uint.MaxValue);
        bytes[DirectoryOffset(bytes) + 128 + 67] = 1;
        bytes[DirectoryOffset(bytes) + 384 + 67] = 0;
        Assert.Equal("x", MathTypeReader.ReadOle(bytes).Math.InnerText);
    }

    [Fact]
    public void RootDirectoryIdsSpanMultipleDirectorySectors()
    {
        var bytes = NestedEquations();
        int directory = DirectoryOffset(bytes);
        byte[] rootEntry = bytes.AsSpan(directory + 384, 128).ToArray();
        uint nextDirectorySector = (uint)(bytes.Length / 512 - 1);
        int nextDirectoryOffset = bytes.Length;
        Array.Resize(ref bytes, bytes.Length + 512);
        rootEntry.CopyTo(bytes, nextDirectoryOffset + 384);
        bytes.AsSpan(directory + 384, 128).Clear();
        int fat = 512 * (1 + (int)BinaryPrimitives.ReadUInt32LittleEndian(bytes.AsSpan(76)));
        uint firstDirectorySector = BinaryPrimitives.ReadUInt32LittleEndian(bytes.AsSpan(48));
        BinaryPrimitives.WriteUInt32LittleEndian(bytes.AsSpan(fat + (int)firstDirectorySector * 4), nextDirectorySector);
        BinaryPrimitives.WriteUInt32LittleEndian(bytes.AsSpan(fat + (int)nextDirectorySector * 4), 0xfffffffe);
        SetPointer(bytes, 0, 76, 7);
        Assert.Equal("x", MathTypeReader.ReadOle(bytes).Math.InnerText);
    }

    [Fact]
    public void NestedOrUnreferencedEquationStreamsCannotSubstituteForAMissingRootStream()
    {
        AssertInvalid(NestedEquations(rootName: "RootData"));
        var orphaned = NestedEquations();
        SetPointer(orphaned, 0, 76, uint.MaxValue);
        AssertInvalid(orphaned);
    }

    [Theory]
    [InlineData("root-type")]
    [InlineData("root-sibling")]
    [InlineData("root-cycle")]
    [InlineData("child-out-of-range")]
    [InlineData("sibling-out-of-range")]
    [InlineData("sibling-cycle")]
    [InlineData("reused-sibling")]
    [InlineData("unallocated-entry")]
    [InlineData("invalid-entry-type")]
    [InlineData("stream-child")]
    [InlineData("odd-name-length")]
    [InlineData("missing-name-terminator")]
    [InlineData("duplicate-root-stream")]
    [InlineData("storage-not-stream")]
    public void InvalidRootDirectoryTreesNeverReturnAnEquation(string invalid)
    {
        var bytes = NestedEquations();
        int directory = DirectoryOffset(bytes);
        switch (invalid)
        {
            case "root-type": bytes[directory + 66] = 2; break;
            case "root-sibling": SetPointer(bytes, 0, 68, 3); break;
            case "root-cycle": SetPointer(bytes, 0, 76, 0); break;
            case "child-out-of-range": SetPointer(bytes, 0, 76, 1024); break;
            case "sibling-out-of-range": SetPointer(bytes, 3, 72, 1024); break;
            case "sibling-cycle": SetPointer(bytes, 3, 72, 3); break;
            case "reused-sibling": SetPointer(bytes, 3, 72, 1); break;
            case "unallocated-entry": bytes[directory + 128 + 66] = 0; break;
            case "invalid-entry-type": bytes[directory + 128 + 66] = 7; break;
            case "stream-child": SetPointer(bytes, 3, 76, 2); break;
            case "odd-name-length": BinaryPrimitives.WriteUInt16LittleEndian(bytes.AsSpan(directory + 384 + 64), 31); break;
            case "missing-name-terminator": bytes[directory + 384 + 30] = 1; break;
            case "duplicate-root-stream": SetPointer(bytes, 1, 72, 2); break;
            case "storage-not-stream": bytes[directory + 384 + 66] = 1; break;
        }
        AssertInvalid(bytes);
    }

    [Fact]
    public void RootStreamNamesAreCaseInsensitiveButCaseOnlyDuplicatesAreAmbiguous()
    {
        Assert.Equal("x", MathTypeReader.ReadOle(NestedEquations(rootName: "equation native")).Math.InnerText);
        var duplicate = NestedEquations(rootName: "equation native");
        SetPointer(duplicate, 1, 72, 2);
        AssertInvalid(duplicate);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(64)]
    [InlineData(65)]
    [InlineData(4095)]
    [InlineData(4096)]
    [InlineData(8193)]
    public void SingleRootStreamRoundTripsThroughMiniAndRegularSectors(int length)
    {
        var payload = Enumerable.Range(0, length).Select(i => (byte)i).ToArray();
        foreach (string name in new[] { "Equation Native", "\u0001Ole10Native" })
        {
            var bytes = CompoundFile.WriteSingleStream(name, payload);
            Assert.Equal(payload, CompoundFile.ReadStream(bytes, name));
            Assert.Null(CompoundFile.ReadStream(bytes, "Absent"));
        }
    }

    private static void AssertInvalid(byte[] bytes) => Assert.Equal("invalid_equation_ole",
        Assert.Throws<MathTypeException>(() => MathTypeReader.ReadOle(bytes)).Code);

    internal static int DirectoryOffset(byte[] bytes) =>
        512 * (1 + (int)BinaryPrimitives.ReadUInt32LittleEndian(bytes.AsSpan(48)));

    internal static void SetPointer(byte[] bytes, int entry, int field, uint value) =>
        BinaryPrimitives.WriteUInt32LittleEndian(bytes.AsSpan(DirectoryOffset(bytes) + entry * 128 + field), value);
}
