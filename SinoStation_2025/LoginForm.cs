using Autodesk.Revit.UI;
using Newtonsoft.Json.Linq;
using System;
using System.Net.Http;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SinoStation_2025
{
    public partial class LoginForm : Form
    {
        private UIApplication m_revitUIApp = null;
        private RevitDocument m_connect = null;
        ExternalEvent m_externalEvent_EditRoomOffset; //外部事件處理:讀取其他的cs檔_EditRoomOffset        
        ExternalEvent m_externalEvent_EditRoomPara; //外部事件處理:讀取其他的cs檔_EditRoomOffset
        ExternalEvent m_externalEvent_CreateCrush; //外部事件處理:建立干涉模型
        public static bool external = false;
        public static string external_username = "";
        public LoginForm(UIApplication uiapp, RevitDocument connect, ExternalEvent externalEvent_EditRoomOffset, ExternalEvent externalEvent_EditRoomPara, ExternalEvent externalEvent_CreateCrush)
        {
            m_revitUIApp = uiapp;
            m_connect = connect;
            m_externalEvent_EditRoomOffset = externalEvent_EditRoomOffset;
            m_externalEvent_EditRoomPara = externalEvent_EditRoomPara;
            m_externalEvent_CreateCrush = externalEvent_CreateCrush;
            InitializeComponent();
        }

        private void LoginForm_Load(object sender, EventArgs e)
        {
            external = false;
            webBrowser1.ScriptErrorsSuppressed = true;
            webBrowser1.DocumentCompleted += new WebBrowserDocumentCompletedEventHandler(webBrowser1_DocumentCompleted);
            //string url = "http://127.0.0.1:8000/user/apilogin/";
            string url = "https://bimdata.sinotech.com.tw/user/apilogin/";
            //string url = "http://bimdata.secltd/user/apilogin/";
            webBrowser1.Url = new Uri(url);
        }

        private void webBrowser1_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            string response = webBrowser1.getExternalLoginResponse();
            if (!String.IsNullOrWhiteSpace(response))
            {
                string user_id = JObject.Parse(response).SelectToken("user_id").ToString();
                string new_token = JObject.Parse(response).SelectToken("new_token").ToString();
                external_username = JObject.Parse(response).SelectToken("user_name").ToString();
                HttpClient client = new HttpClient();
                //client.BaseAddress = new Uri("http://127.0.0.1:8000/");
                client.BaseAddress = new Uri("https://bimdata.sinotech.com.tw/");
                //client.BaseAddress = new Uri("http://bimdata.secltd/");
                client.DefaultRequestHeaders.Accept.Clear();
                var headerValue = new System.Net.Http.Headers.MediaTypeWithQualityHeaderValue("application/json");
                client.DefaultRequestHeaders.Accept.Add(headerValue);
                client.DefaultRequestHeaders.ConnectionClose = true;
                var multiForm = new MultipartFormDataContent();
                multiForm.Add(new StringContent("SinoStation-RegulatoryReview"), "RevitAPI");
                Task.WaitAll(client.PostAsync($"/user/apilogin/success/"+user_id+"/"+new_token+"/", multiForm));
                RegulatoryReviewForm.client = client;
                external = true;
                //this.Close();
                this.Dispose();
                if (external == true)
                {
                    RegulatoryReviewForm regulatoryReviewForm = new RegulatoryReviewForm(m_revitUIApp, m_connect, m_externalEvent_EditRoomOffset, m_externalEvent_EditRoomPara, m_externalEvent_CreateCrush);
                    regulatoryReviewForm.Show();
                    m_externalEvent_CreateCrush.Raise();
                }
            }
        }
    }
}
