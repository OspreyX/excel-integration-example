using System;
using System.Collections.Generic;
using System.Threading;
using System.Runtime.InteropServices;

using ExcelDna.Integration;
using ExcelDna.Integration.Rtd;
using Newtonsoft.Json.Linq;

using Openfin.Desktop;



namespace Openfin.RTDAddin
{
    public class AddInRequest
    {
        public AddInRequest(ExcelRtdServer.Topic topic, IList<String> topicInfo) {
            this.topic = topic;
            this.topicInfo = topicInfo;
        }
        private ExcelRtdServer.Topic topic;
        private IList<String> topicInfo;
        public ExcelRtdServer.Topic Topic { get { return topic; } }
        public IList<String> TopicInfo { get { return topicInfo; } }
        public String SubscribeKey { get; set; }
        public String SubscribeTopic { get; set; }  // topic to subscribe in Desktop
    }

    /// <summary>
    ///     Base class for RTD server 
    /// </summary> 
    public abstract class BaseFinServer : ExcelRtdServer, DesktopStateListener
    {
        protected DesktopConnection controller;
        protected Application application;
        protected bool desktopReady = false;
        protected List<AddInRequest> pendingRequests; // store requests sent before desktop is ready
        protected Dictionary<String, List<AddInRequest>> topicMap;

        protected abstract String getFinDesktopAppId();

        [DllImport("user32.dll", SetLastError = true)]
        private static extern void SwitchToThisWindow(IntPtr hWnd, bool turnOn);


        /// <summary>
        /// Invoked by ExcelDNA when loaded
        /// </summary> 
        protected override bool ServerStart()
        {
            pendingRequests = new List<AddInRequest>();
            topicMap     = new Dictionary<string, List<AddInRequest>>();
            controller = new DesktopConnection(getFinDesktopAppId(), "localhost", 9696);
            controller.connect(this);

            return true;
        }

        /// <summary>
        /// Invoked by ExcelDNA when unloaded
        /// </summary> 
        protected override void ServerTerminate()
        {
            this.controller.disconnect();
        }

        /// <summary>
        /// Invoked by ExcelDNA when a cell in Excel is updated with matching function
        /// </summary> 
        protected override object ConnectData(Topic topic, IList<string> topicInfo, ref bool newValues)
        {
            AddInRequest request = new AddInRequest(topic, topicInfo);
            if (desktopReady)
            {
                request.SubscribeKey = getSubscriptionKey(topicInfo);
                request.SubscribeTopic = topicInfo[1];
                addSubscription(request);
            }
            else
            {
                this.pendingRequests.Add(request);
            }
            return getConnectDataDescription(request);
        }

        /// <summary>
        /// Invoked by ExcelDNA when a cell in Excel is updated with removing matching function
        /// </summary> 
        protected override void DisconnectData(Topic topic)
        {
            AddInRequest existing = null;
            foreach (KeyValuePair<String, List<AddInRequest>> pair in this.topicMap)
            {
                foreach (AddInRequest request in pair.Value)
                {
                    if (request.Topic == topic)
                    {
                        existing = request;
                        break;
                    }
                }
                if (existing != null)
                {
                    pair.Value.Remove(existing);
                    if (pair.Value.Count == 0)
                    {
                        unsubscribe(existing);
                    }
                    break;
                }
            }
        }

        private string GetTime()
        {
            return DateTime.Now.ToString("HH:mm:ss.fff");
        }

        /// <summary>
        ///     Callback when OpenFin Desktop is successfully connected and ready to 
        ///     accept commands.
        /// </summary> 
        public void onReady()
        {
            this.desktopReady = true;
            Console.WriteLine("Connection authorized.");
            Console.WriteLine("Creating app...");

            foreach (AddInRequest request in pendingRequests)
            {
                request.SubscribeKey = getSubscriptionKey(request.TopicInfo);
                request.SubscribeTopic = request.TopicInfo[1];
                addSubscription(request);
            }
        }

        public void onError(String reason)
        {
            Console.WriteLine("onError onMessage: {0}", reason);
        }

        /// <summary>
        ///     Callback when a message is sent to this client from OpenFin Desktop.
        /// </summary>
        public void onMessage(String message)
        {
            Console.WriteLine("DemoDesktopStateListener onMessage: {0}", message);
        }

        /// <summary>
        ///     Callback when a message is sent from this client to OpenFin Desktop.
        /// </summary>
        public void onOutgoingMessage(String message)
        {
        }

        /// <summary>
        ///     Callback when the connection with the OpenFin Desktop has closed.
        /// </summary>
        public void onClosed()
        {
        }

        /// <summary>
        ///     add a messge subscription when matching function is enterd in a cell.
        /// </summary>
        protected void addSubscription(AddInRequest request)
        {
            String subKey = request.SubscribeKey;
            List<AddInRequest> list;
            if (!this.topicMap.TryGetValue(subKey, out list))
            {
                list = new List<AddInRequest>();
                this.topicMap.Add(subKey, list);
                subscribe(request);
            }
            list.Add(request);

        }

        /// <summary>
        ///     Get subscription key from cell function.
        /// </summary>
        protected virtual string getSubscriptionKey(IList<string> topicInfo)
        {
            return topicInfo[0] + topicInfo[1];  // should appId + topic
        }

        /// <summary>
        ///     Send subscription request to OpenFin Desktop.
        /// </summary>
        protected virtual void subscribe(AddInRequest request)
        {
            controller.getInterApplicationBus().subscribe(request.TopicInfo[0], request.SubscribeTopic, processMessage);
        }

        /// <summary>
        ///     Send unsubscription request to OpenFin Desktop.
        /// </summary>
        protected virtual void unsubscribe(AddInRequest request)
        {
            controller.getInterApplicationBus().unsubscribe(request.TopicInfo[0], request.TopicInfo[1], processMessage);
        }

        /// <summary>
        ///     process subscribed messages from OpenFin Desktop.
        /// </summary>
        protected void processMessage(String sourceUuid, String topic, object message)
        {
            JObject jMessage = (JObject) message;
            List<AddInRequest> list;
            string realTopic = this.getTopicInMessage(topic, jMessage);
            if (this.topicMap.TryGetValue(sourceUuid + realTopic, out list))
            {
                foreach (AddInRequest request in list)
                {
                    request.Topic.UpdateValue(getDisplayValue(request, jMessage));
                }
            }
            else
            {
                Console.WriteLine("Missing topic {0}", topic);
            }
        }

        protected virtual string getTopicInMessage(string topic, JObject jMessage)
        {
            return topic;
        }

        /// <summary>
        ///     Get display value from message based on cell function.
        /// </summary>
        protected virtual string getDisplayValue(AddInRequest request, JObject message)
        {
            return message.ToString();
        }

        /// <summary>
        ///     Get cell function connect description. The value is shown when a cell function is entered but data is not received yet.
        /// </summary>
        protected virtual string getConnectDataDescription(AddInRequest request)
        {
            if (String.IsNullOrEmpty(request.SubscribeKey))
            {
                return "";
            }
            else
            {
                return "Connecting to: " + request.SubscribeKey;
            }
        }

        protected void bringExcelToFront()
        {
            SwitchToThisWindow(ExcelDnaUtil.WindowHandle, false);
            
        }
    }

    /// <summary>
    ///     Server class for FinDesktop function
    ///     Excel function is in form of 
    ///     =FinDesktop("550e8400-e29b-41d4-a716-4466333333000","EURUSD","bidPrice")
    /// </summary> 
    public class FinDesktopServer : BaseFinServer
    {
        /// <summary>
        ///     Application UUID for connecting to OpenFin Desktop.
        /// </summary>
        protected override String getFinDesktopAppId()
        {
            return "FinDesktopServerCS";
        }

        /// <summary>
        ///     Parse display value from message based on cell function.
        ///     Excel function is in form of 
        ///     =FinDesktop("550e8400-e29b-41d4-a716-4466333333000","EURUSD","bidPrice")
        ///     request.TopicInfo[2] has value of the 3rd parameter, which is a field in the message.
        ///     This method extracts and returns value of the field so it can be displayed
        /// </summary>
        protected override string getDisplayValue(AddInRequest request, JObject message)
        {
            String value = null;
            if (String.IsNullOrEmpty(request.TopicInfo[2]))
            {
                value = message.ToString();
            }
            else
            {
                value = DesktopUtils.getJSONString(message, request.TopicInfo[2]);
            }
            return value;
        }    
    }


}