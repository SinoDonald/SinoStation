using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using Newtonsoft.Json.Linq;

namespace SinoStation_2025
{
    public static class GetWebsiteResponse
    {
        public static string getContentaById(this WebBrowser _webBrowser, string pass_content, bool _resetEle = true)
        {
            var response = _webBrowser.Document.GetElementById(pass_content);
            if (response == null) return "";

            string content = response.InnerText;
            if (_resetEle) response.InnerText = "";

            return content == null ? "" : content;
        }

        public static string getExternalLoginResponse(this WebBrowser _webBrowser)
        {
            string content = _webBrowser.getContentaById("click_user");
            if (content == "") return "";

            return content;
        }
    }
}
