using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using ExcelDna.Integration.Rtd;

namespace RtdPerformance
{
    [ComVisible(true)]                   // Required since the default template puts [assembly:ComVisible(false)] in the AssemblyInfo.cs
    [ProgId(RtdServer.ServerProgId)]     //  If ProgId is not specified, change the XlCall.RTD call in the wrapper to use namespace + type name (the default ProgId)
    public class RtdServer : ExcelRtdServer
    {
        public const string ServerProgId = "RtdPerformance.Server";

        static DataService _dataService;

        protected override bool ServerStart()
        {
            Debug.Print("ServerStart called");
            _dataService = new DataService();
            return true;
        }

        protected override void ServerTerminate()
        {
            if (_dataService == null)
            {
                Debug.Print("ServerStart not called ???");
                return;
            }

            _dataService.Terminate();
        }

        protected override object ConnectData(Topic topic, IList<string> topicInfo, ref bool newValues)
        {
            _dataService.ConnectTopic(topic);
            
            return topic.Value;
        }

        protected override void DisconnectData(Topic topic)
        {
            _dataService.DisconnectTopic(topic);
        }
    }
}