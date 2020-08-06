using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;
using ExcelDna.Integration;

namespace Uploader
{
    // 1. ExcelMethods contains the single function registered with Excel,
    //    and the two macros that will trigger the upload
    public static class ExcelMethods
    {
        // UploadCreate is the worksheet function
        // It reads the caller, which we need to track for the UploadSelection option, 
        // gathers all the arguments and then starts the RTD tracking with a call to ExccelAsyncUtil.Observe.
        //
        // There can be many more arguments
        // We might experiment with AllowReference=true for the arguments
        // That could improve performance a lot
        // But there are concerns:
        // * Does Upload get called whenever the contents of the reference change?
        // * We need care when we read it (at the Upload-time)
        public static object UploadCreate(object arg1, object arg2, object arg3)
        {
            var caller = XlCall.Excel(XlCall.xlfCaller) as ExcelReference;  // Might be null if called in unusual context, e.g. Application.Run
            var topicFunctionName = nameof(UploadCreate);
            var topicArguments = new object[] { caller, arg1, arg2, arg3 };   // Adding the caller here is important since we need to track it for the RTD topic too
            return ExcelAsyncUtil.Observe(topicFunctionName, topicArguments, () => UploadManager.CreateItem(topicArguments));
        }

        // These two methods just do the upload.
        // It's easy for testing to use the "Menuxxx", but normally these would be called from a Ribbon callback.
        [ExcelCommand(MenuName = "Uploader", MenuText = "Upload All")]
        public static void UploadAll()
        {
            UploadManager.UploadAll();
        }

        // It's easy for testing to use the "Menuxxx", but normally these would be called from a Ribbon callback.
        // But from a Ribbon callback we can't call the C API, which we want to here to get the Selection as an ExcelReference
        // So from a Ribbon callback we would need the QueueAsMacro
        [ExcelCommand(MenuName = "Uploader", MenuText = "Upload Selection")]
        public static void UploadSelection()
        {
            // From a menu we don't need the QueueAsMacro, but from a ribbon button we do
            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                var selection = XlCall.Excel(XlCall.xlfSelection) as ExcelReference;
                if (selection != null)
                    UploadManager.UploadSelection(selection);
                else
                    Console.Beep(); // Like when it's some chart element
            });
        }
    }

    // 2. UploadStatus represents the state of an individual item
    //    We expect it to only change through the steps, else might need some synchronization
    enum UploadStatus
    {
        Waiting,
        InProgress,
        CompletedSuccess,
        CompletedError
    }

    // 3. UploadItem encapsulates the upload job
    //    Through the IExcelObservable mechanism, it will "stream" its status back to the cell, 
    //    as that changes when upload starts, and completes.
    //    The IDisposable directly on this class is a small hack - see the comment at "Subscribe"
    class UploadItem : IExcelObservable, IDisposable
    {
        public UploadStatus Status;
        readonly public object[] Arguments;
        public ExcelReference Caller; // Just a convenience

        // This method will always be called on the main thread
        public UploadItem(ExcelReference caller, object[] arguments)
        {
            Caller = caller;
            Arguments = arguments;
            Status = UploadStatus.Waiting;
        }

        // This method is safe to call from any thread
        // But means that the Status field can change at any time
        public void SetStatus(UploadStatus newStatus)
        {
            Debug.Assert((int)newStatus >= (int)Status, "UpdoadItem.Status should not regress");
            Status = newStatus;
            ReportState();
        }

        // We're implementing against the simpler requirements of the IExcelObservable,
        // where we can assume that Subscribe sill be called at most once,
        // and so we can also implement IDisposable directly here.
        // A more complicated implementation for this (list of IObservers, separate IDisposable object)
        // could be like normal IObservable implementations, hence easier to understand.
        IExcelObserver _observer;
        public IDisposable Subscribe(IExcelObserver observer)
        {
            Debug.Assert(_observer == null);
            _observer = observer;
            ReportState(); // Reporting immediately in the Subscribe means we never return #N/A from into the cell
            return this;
        }

        // This method will always be called on the main thread
        public void Dispose()
        {
            UploadManager.NotifyDispose(this);
        }

        // OnNext is safe to call from any thread
        // (I'm not sure if _observer could really be null in this example... but I put the guard in anyway)
        void ReportState()
        {
            // Whatever we return here will be the result value in the cell
            // I'm showing Caller as a quick diagnostic check
            _observer?.OnNext($"Upload Item at {Caller} - Status: {Status}");
        }
    }

    // 4. UploadManager keeps track of all the items, and manages the transition to start uploading items on demand
    static class UploadManager
    {
        public static List<UploadItem> UploadItems = new List<UploadItem>();

        // We require the first element in arguments to be the caller
        // That simplified the UploadCreate function wrapper, but makes us do a little extra juggling here
        public static IExcelObservable CreateItem(object[] topicArguments)
        {
            var caller = topicArguments[0] as ExcelReference;
            var functionArguments = topicArguments.Skip(1).ToArray();
            var item = new UploadItem(caller, functionArguments);
            UploadItems.Add(item);
            Debug.Print($"CreateItem: {item.Caller}, Arguments: {topicArguments.Length - 1}");
            return item;
        }

        static readonly Random random = new Random(2103);
        public static void UploadAll()
        {
            // Get all the items currently "Waiting", update their status and send for processing
            var waiting = UploadItems.Where(item => item.Status == UploadStatus.Waiting).ToList();
            foreach (var item in waiting)
                item.SetStatus(UploadStatus.InProgress);

            PerformUploads(waiting);
        }

        public static void UploadSelection(ExcelReference selection)
        {
            if (selection == null)
                return;

            // Get all the items currently "Waiting", update their status and send for processing
            var waiting = UploadItems.Where(item => item.Status == UploadStatus.Waiting && IsInsideSelection(item.Caller)).ToList();
            foreach (var item in waiting)
                item.SetStatus(UploadStatus.InProgress);

            PerformUploads(waiting);

            // We just check the top left of the caller
            bool IsInsideSelection(ExcelReference reference)
            {
                if (reference == null)
                    return false;

                if (reference.SheetId != selection.SheetId)
                    return false;

                return reference.RowFirst >= selection.RowFirst &&
                       reference.RowFirst <= selection.RowLast &&
                       reference.ColumnFirst >= selection.ColumnFirst &&
                       reference.ColumnFirst <= selection.ColumnLast;
            }
        }

        internal static void NotifyDispose(UploadItem uploadItem)
        {
            UploadItems.Remove(uploadItem);
            Debug.Print($"RemoveItem: {uploadItem.Caller} - Status: {uploadItem.Status}");
        }

        // 5. This part is where the UploadManager ships off the items for doing the real upload.
        //    The need to then transition the UploadItem.Status to a CompletedXXXX state
        //    The UploadItem.SetStatus can be called from any thread, and will then update through to the sheet.
        // For the example I just put in a random delay and result
        // The real implementation might process items all at once, or in batches too, not a task for every item.
        static void PerformUploads(List<UploadItem> items)
        {
            foreach (var item in items)
            {
                Task.Run(async () =>
                {
                    // Randon delay and result
                    await Task.Delay(random.Next(10000));
                    if (random.Next(3) == 1)
                        item.SetStatus(UploadStatus.CompletedError);
                    else
                        item.SetStatus(UploadStatus.CompletedSuccess);
                });
            }
        }
    }

}
