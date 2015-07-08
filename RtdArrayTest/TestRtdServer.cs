using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using ExcelDna.Integration;
using ExcelDna.Integration.Rtd;

namespace RtdArrayTest
{
    class TestArrayTopic : ExcelRtdServer.Topic
    {
        public string Prefix;

        public TestArrayTopic(ExcelRtdServer server, int topicId) :
            base(server, topicId)
        {
        }
    }

    [ComVisible(true)]
    public class TestRtdServer : ExcelRtdServer
    {
        Random _random = new Random();
        Timer _timer;
        List<TestArrayTopic> _topics;

        protected override bool ServerStart()
        {
            _timer = new Timer(UpdateTopics, null, TimeSpan.FromSeconds(1), TimeSpan.FromSeconds(1));
            _topics = new List<TestArrayTopic>();
            return true;
        }

        protected override void ServerTerminate()
        {
            _timer.Dispose();
            _timer = null;
        }

        protected override Topic CreateTopic(int topicId, IList<string> topicInfo)
        {
            return new TestArrayTopic(this, topicId) { Prefix = topicInfo[0] };
        }

        protected override object ConnectData(Topic topic, IList<string> topicInfo, ref bool newValues)
        {
            TestArrayTopic testArrayTopic = (TestArrayTopic)topic;
            _topics.Add(testArrayTopic);
            Debug.Print("ConnectData - Prefix {0}", testArrayTopic.Prefix);
            return ExcelErrorUtil.ToComError(ExcelError.ExcelErrorNA);
        }

        protected override void DisconnectData(Topic topic)
        {
            TestArrayTopic testArrayTopic = (TestArrayTopic)topic;
            _topics.Remove(testArrayTopic);
            Debug.Print("DisconnectData - Prefix {0}", testArrayTopic.Prefix);
        }

        void UpdateTopics(object _unused_)
        {
            foreach (TestArrayTopic topic in _topics)
            {
                var value = topic.Prefix + ":" + DateTime.Now.ToString("HH:mm:ss.fff") + ";" + _random.NextDouble().ToString("F5");
                topic.UpdateValue(value);
            }
        }
    }
}
