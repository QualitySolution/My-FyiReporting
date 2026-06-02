using System;
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
        private static volatile bool _cancelRequested;

        public static int Current { get { return Volatile.Read(ref _current); } }
        public static int Total { get { return Volatile.Read(ref _total); } }
        public static string Phase { get { return Volatile.Read(ref _phase); } }

        /// <summary>True if the UI requested cancellation of the current export</summary>
        public static bool CancelRequested { get { return _cancelRequested; } }

        /// <summary>Cooperative cancel: the next Increment() in the render loop throws</summary>
        public static void RequestCancel() { _cancelRequested = true; }

        /// <summary>Throws OperationCanceledException if a cancel was requested</summary>
        public static void ThrowIfCancellationRequested()
        {
            if (_cancelRequested) throw new OperationCanceledException();
        }

        public static void Reset()
        {
            Volatile.Write(ref _current, 0);
            Volatile.Write(ref _total, 0);
            Volatile.Write(ref _phase, null);
            _cancelRequested = false;
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
            if (_cancelRequested) throw new OperationCanceledException();
            Interlocked.Increment(ref _current);
        }
    }
}
