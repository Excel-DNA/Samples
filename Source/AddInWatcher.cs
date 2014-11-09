using ExcelDna.Integration;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Serialization;

namespace AddInReloader
{
    class AddInWatcher : IDisposable
    {
        // For every directory we watch, keep track of all the add-ins that have files in that directory
        Dictionary<string, WatchedDirectory> _watchedDirectories = new Dictionary<string, WatchedDirectory>();
        HashSet<WatchedAddIn> _dirtyAddIns = new HashSet<WatchedAddIn>();
        object _dirtyLock = new object();

        public AddInWatcher(AddInReloaderConfiguration config)
        {
            foreach (var addIn in config.WatchedAddIns)
            {
                foreach (var file in addIn.WatchedFiles)
                {
                    var directory = Path.GetDirectoryName(file.Path);
                    WatchedDirectory wd;
                    if (!_watchedDirectories.TryGetValue(directory, out wd))
                    {
                        wd = new WatchedDirectory(directory, InvalidateAddIn);
                    }
                    wd.WatchAddIn(addIn);
                }
            }
        }

        // Called in the event handler - don't do slow work here.
        void InvalidateAddIn(WatchedAddIn watchedAddIn)
        {
            lock (_dirtyLock)
            {
                _dirtyAddIns.Add(watchedAddIn);
                ExcelAsyncUtil.QueueAsMacro(ReloadDirtyAddIns);
            }
        }

        // Running in macro context.
        void ReloadDirtyAddIns()
        {
            HashSet<WatchedAddIn> dirtyCopy;
            lock (_dirtyLock)
            {
                dirtyCopy = _dirtyAddIns;
                _dirtyAddIns = new HashSet<WatchedAddIn>();
            }
            foreach (var addIn in dirtyCopy)
            {
                ReloadAddIn(addIn.Path);
            }

            // Force a recalculate on open workbooks.
            XlCall.Excel(XlCall.xlcCalculateNow);
        }

        // Running in macro context.
        static void ReloadAddIn(string xllPath)
        {
            ExcelIntegration.RegisterXLL(xllPath);
        }

        public void Dispose()
        {
            foreach (var wd in _watchedDirectories.Values)
            {
                wd.Dispose();
            }
        }

        class WatchedDirectory : IDisposable
        {
            string _path;
            FileSystemWatcher _directoryWatcher;
            Dictionary<string, WatchedAddIn> _watchedFiles;
            Action<WatchedAddIn> _invalidateAddIn;

            public WatchedDirectory(string path, Action<WatchedAddIn> invalidateAddIn)
            {
                _path = path;
                _directoryWatcher = new FileSystemWatcher(path);
                _directoryWatcher.NotifyFilter = NotifyFilters.LastWrite;
                _directoryWatcher.Changed += DirectoryWatcher_Changed;
                _watchedFiles = new Dictionary<string, WatchedAddIn>(StringComparer.OrdinalIgnoreCase);
                _invalidateAddIn = invalidateAddIn;

                _directoryWatcher.EnableRaisingEvents = true;
            }

            public void WatchAddIn(WatchedAddIn addIn)
            {
                foreach (var file in addIn.WatchedFiles)
                {
                    var fullPath = System.IO.Path.GetFullPath(file.Path);
                    _watchedFiles[fullPath] = addIn; // This only allows one add-in to watch a particular file.
                }
            }

            public void Dispose()
            {
                _directoryWatcher.Dispose();
            }

            void DirectoryWatcher_Changed(object sender, FileSystemEventArgs e)
            {
                Debug.Assert(string.Equals(System.IO.Path.GetFullPath(e.FullPath), e.FullPath, StringComparison.OrdinalIgnoreCase));

                WatchedAddIn addIn;
                if (_watchedFiles.TryGetValue(e.FullPath, out addIn))
                {
                    _invalidateAddIn(addIn);
                }
            }
        }
    }
}
