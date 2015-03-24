using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelDna.Integration;

namespace ObjectHandles
{
    public static class ExcelFunctions
    {
        static DataService _dataService = new DataService();
        static ObjectHandler _objectHandler = new ObjectHandler(_dataService);

        public static object TestHandleSample()
        {
            return "Hello form ObjectHandles!";
        }

        public static object CreateDataObject1(string code)
        {
            return _objectHandler.GetHandle("DataObject1", new object[] { code },
                (objectType, parameters) => _dataService.ProcessRequest(objectType, parameters));
        }

        public static string GetCode(string handle)
        {
            object value;
            // TODO: We might be able to strongly type the GetObject...
            if (_objectHandler.TryGetObject(handle, out value))
            {
                DataObject1 data = (DataObject1)value;
                return data.Code;
            }
            // No object for the handle ...
            return "!!! INVALID HANDLE";
        }

        public static object GetDateTime(string handle)
        {
            object value;
            // TODO: We might be able to strongly type the GetObject...
            if (_objectHandler.TryGetObject(handle, out value))
            {
                DataObject1 data = (DataObject1)value;
                return data.DateTime;
            }
            // No object for the handle ...
            return "!!! INVALID HANDLE";
        }

        public static object GetValue(string handle)
        {
            object value;
            // TODO: We might be able to strongly type the GetObject...
            if (_objectHandler.TryGetObject(handle, out value))
            {
                DataObject1 data = (DataObject1)value;
                return data.Value;
            }
            // No object for the handle ...
            return "!!! INVALID HANDLE";
        }

        // Forces a refresh of all the objects in the handler
        // All objects will be recreated and return a new handle, invalidating all dependencies.
        [ExcelCommand(MenuName="Object Handler", MenuText="Refresh All")]
        public static void RefreshAll()
        {
            _objectHandler.RefreshAll();
        }

        // Does an update of all objects in the handler
        // This is done by a query to the back end, passing in the current rowversions.
        // When the query is done, only the objects with updated rowversions will be refreshed with a new handle.
        [ExcelCommand(MenuName = "Object Handler", MenuText = "Update All")]
        public static void UpdateAll()
        {
            _objectHandler.UpdateAll();
        }
    }
}
