using System.Diagnostics;
using System.Runtime.InteropServices;

namespace XLServer
{
    // Step 1: Defines an event sink interface (MessageEvents) to be implemented by the COM sink.
    [ComVisible(true)]
    [Guid("c2b81fbe-8a2d-4e2a-9594-4f3a282a3194")]
    [InterfaceTypeAttribute(ComInterfaceType.InterfaceIsIDispatch)]
    public interface MessageEvents
    {
        [DispId(1)]
        void NewMessage(string s);
    }

    [ComVisible(true)]
    [Guid("f256946f-bb8c-4cd7-aad3-61838a20d8d2")]
    public interface IMessageHandler
    {
        void FireNewMessageEvent(string s);
    }

    // Step 2: Connects the event sink interface to a class by passing the namespace and event sink interface
    [ComVisible(true)]
    [Guid("bec4d3ef-3639-4af6-996e-b837f1d96a24")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComSourceInterfaces(typeof(MessageEvents))]
    public class MessageHandler : IMessageHandler
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
