using System.Diagnostics;
using System.Runtime.InteropServices;

namespace XLServer
{
    // Step 1: Defines an event sink interface (MessageEvents) to be implemented by the COM sink.
    [ComVisible(true)]
    [InterfaceTypeAttribute(ComInterfaceType.InterfaceIsIDispatch)]
    public interface MessageEvents
    {
        void NewMessage(string s);
    }

    // Step 2: Connects the event sink interface to a class by passing the namespace and event sink interface
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    [ComSourceInterfaces(typeof(MessageEvents))]
    public class MessageHandler
    {
        public delegate void NewMessageDelegate(string s);
        public event NewMessageDelegate NewMessage;
        public MessageHandler() { }

        public void FireNewMessageEvent(string s)
        {
            Debug.Print($"New Message {s}");
            if (NewMessage != null)
            {
                Debug.Print($"Invoke {s}");
                NewMessage.Invoke(s);
            }
        }
    }
}
