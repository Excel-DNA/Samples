using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Threading;
using ExcelDna.Integration.Rtd;

namespace RtdClock_ExcelRtdServer_PreRegistered
{
    [ComVisible(true)]
    [Guid("296F3518-6374-4F71-B35F-6938346C6076")]
    [ProgId(RtdClockServer.ServerProgId)]
    public class RtdClockServer : ExcelRtdServer
    {
        public const string ServerProgId = "RtdClock.ClockServer";

        readonly object _topicsLock = new object();
        Timer _timer;
        List<Topic> _topics;

        protected override bool ServerStart()
        {
            _topics = new List<Topic>();
            _timer = new Timer(timer_tick, null, 0, 1000);
            return true;
        }

        protected override void ServerTerminate()
        {
            _timer.Dispose();
        }

        protected override object ConnectData(Topic topic, IList<string> topicInfo, ref bool newValues)
        {
            lock (_topicsLock)
            {
                _topics.Add(topic);
            }

            return DateTime.Now.ToString("HH:mm:ss") + " (ConnectData)";
        }

        protected override void DisconnectData(Topic topic)
        {
            lock (_topicsLock)
            {
                _topics.Remove(topic);
            }
        }

        void timer_tick(object _unused_state_)
        {
            string now = DateTime.Now.ToString("HH:mm:ss");

            List<Topic> topics;
            lock (_topicsLock)
            {
                topics = new List<Topic>(_topics);
            }

            foreach (var topic in topics)
            {
                topic.UpdateValue(now);
            }
        }
    }
}
