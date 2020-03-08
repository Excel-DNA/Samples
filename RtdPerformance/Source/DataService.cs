using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Threading;
using ExcelDna.Integration.Rtd;

namespace RtdPerformance
{
    class DataService
    {
        List<ExcelRtdServer.Topic> _activeTopics;
        Thread _updateThread;
        Random _random;

        public DataService()
        {
            _updateThread = new Thread(RunUpdates);
            _updateThread.Start();

            _activeTopics = new List<ExcelRtdServer.Topic>();
            _random = new Random(1);
        }

        public void ConnectTopic(ExcelRtdServer.Topic topic)
        {
            lock (_activeTopics)
            {
                _activeTopics.Add(topic);
                topic.UpdateValue($"ConnectData ({DateTime.Now.ToString("HH:mm:ss")}");
            }
        }

        public void DisconnectTopic(ExcelRtdServer.Topic topic)
        {
            lock (_activeTopics)
            {
                _activeTopics.Remove(topic);
            }
        }

        public void Terminate()
        {
            _updateThread.Abort();
        }

        // Runs on update thread        
        void RunUpdates()
        {
            try
            {
                while (true)
                {
                    UpdateSomeTopics();
                    Thread.Sleep(100);
                }
            }
            catch (ThreadAbortException)
            {
                Debug.Print("Update thead aborted");
            }
        }

        // Runs on update thread        
        void UpdateSomeTopics()
        {
            // string updateValue = DateTime.Now.ToString("HH:mm:ss.fff");
            lock (_activeTopics)
            {
                foreach (var topic in _activeTopics)
                {
                    if (_random.Next(10) == 0)
                    {
                        topic.UpdateValue(DateTime.Now);
                    }
                }
            }
        }
    }
}
