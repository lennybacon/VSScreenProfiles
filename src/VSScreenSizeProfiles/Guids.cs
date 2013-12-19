// Guids.cs
// MUST match guids.h

using System;

namespace devcoach.Tools.ScreenProfiles
{
    static class GuidList
    {
        public const string guidScreenProfilesPkgString = "878ffc19-4e3f-4992-a55c-70119329f47d";
        public const string guidScreenProfilesCmdSetString = "93b10c3b-6ad2-4074-80fc-19d5abf17c3d";

        public static readonly Guid guidScreenProfilesCmdSet = new Guid(guidScreenProfilesCmdSetString);
    };
}