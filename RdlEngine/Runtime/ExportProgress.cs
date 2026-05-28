/* ====================================================================
   Copyright (C) 2004-2008  fyiReporting Software, LLC
   Copyright (C) 2011  Peter Gill <peter@majorsilence.com>

   This file is part of the fyiReporting RDL project.

   Licensed under the Apache License, Version 2.0 (the "License");
   you may not use this file except in compliance with the License.
   You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0
*/
using System.Threading;

namespace fyiReporting.RDL
{
    /// <summary>
    /// Thread-safe progress counters for long-running renders
    /// </summary>
    public static class ExportProgress
    {
        private static int _current;
        private static int _total;
        private static string _phase;

        public static int Current { get { return Volatile.Read(ref _current); } }
        public static int Total { get { return Volatile.Read(ref _total); } }
        public static string Phase { get { return Volatile.Read(ref _phase); } }

        public static void Reset()
        {
            Volatile.Write(ref _current, 0);
            Volatile.Write(ref _total, 0);
            Volatile.Write(ref _phase, null);
        }

        public static void BeginPhase(string phaseName, int total)
        {
            Volatile.Write(ref _current, 0);
            Volatile.Write(ref _total, total);
            Volatile.Write(ref _phase, phaseName);
        }

        public static void BeginIndeterminate(string phaseName)
        {
            Volatile.Write(ref _current, 0);
            Volatile.Write(ref _total, 0);
            Volatile.Write(ref _phase, phaseName);
        }

        public static void AddToTotal(int rows)
        {
            if (rows > 0) Interlocked.Add(ref _total, rows);
        }

        public static void Increment()
        {
            Interlocked.Increment(ref _current);
        }
    }
}
