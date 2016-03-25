using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Threading;
using ExcelDna.Integration.Rtd;

namespace RtdClock_ExcelRtdServer
{
    [ComVisible(true)]                   // Required since the default template puts [assembly:ComVisible(false)] in the AssemblyInfo.cs
    [ProgId(RtdClockServer.ServerProgId)]     //  If ProgId is not specified, change the XlCall.RTD call in the wrapper to use namespace + type name (the default ProgId)
    public class RtdClockServer : ExcelRtdServer
    {
        public const string ServerProgId = "RtdClock.ClockServer";

        // Using a System.Threading.Time which invokes the callback on a ThreadPool thread 
        // (normally that would be dangeours for an RTD server, but ExcelRtdServer is thrad-safe)
        Timer _timer;
        List<Topic> _topics;

        protected override bool ServerStart()
        {
            _timer = new Timer(timer_tick, null, 0, 1000);
            _topics = new List<Topic>();
            return true;
        }

        protected override void ServerTerminate()
        {
            _timer.Dispose();
        }

        protected override object ConnectData(Topic topic, IList<string> topicInfo, ref bool newValues)
        {
            _topics.Add(topic);
            return DateTime.Now.ToString("HH:mm:ss") + " (ConnectData)";
        }

        protected override void DisconnectData(Topic topic)
        {
            _topics.Remove(topic);
        }

        void timer_tick(object _unused_state_)
        {
            string now = DateTime.Now.ToString("HH:mm:ss");
            foreach (var topic in _topics)
                topic.UpdateValue(now);
        }
    }
}
