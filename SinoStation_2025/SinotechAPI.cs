using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Windows;
using System.Windows.Media.Imaging;

namespace SinoStation_2025
{
    public class SinotechAPI : IExternalApplication
    {
        public string addinAssmeblyPath = Assembly.GetExecutingAssembly().Location;
        public Result OnStartup(UIControlledApplication a)
        {
            string sinoStationPath = Path.Combine(Directory.GetParent(addinAssmeblyPath).FullName, "SinoStation_2025.dll");
            string checkPlatformPath = Path.Combine(Directory.GetParent(addinAssmeblyPath).FullName, "CheckPlatform.dll");
            string checkPasswayPath = Path.Combine(Directory.GetParent(addinAssmeblyPath).FullName, "CheckPassway.dll");

            RibbonPanel ribbonPanel = null;
            try { a.CreateRibbonTab("捷運規範校核"); } catch { }
            try { ribbonPanel = a.CreateRibbonPanel("捷運規範校核", "空間需求校核"); } 
            catch 
            {
                List<RibbonPanel> panel_list = new List<RibbonPanel>();
                panel_list = a.GetRibbonPanels("捷運規範校核");
                foreach (RibbonPanel rp in panel_list)
                {
                    if (rp.Name == "空間需求校核")
                    {
                        ribbonPanel = rp;
                    }
                }
            }

            PushButton pushbutton1 = ribbonPanel.AddItem(new PushButtonData("SinoStation_2025", "房間", sinoStationPath, "SinoStation_2025.RegulatoryReview")) as PushButton;
            pushbutton1.LargeImage = convertFromBitmap(Properties.Resources.房間檢討);
            PushButton pushbutton2 = ribbonPanel.AddItem(new PushButtonData("Check Platform", "月台", checkPlatformPath, "CheckPlatform.CheckPlatform")) as PushButton;
            pushbutton2.LargeImage = convertFromBitmap(Properties.Resources.CheckPlatform);
            PushButton pushbutton3 = ribbonPanel.AddItem(new PushButtonData("Use Dimension", "通道尺寸標註", checkPasswayPath, "CheckPassway.CreateDimensionType")) as PushButton;
            pushbutton3.LargeImage = convertFromBitmap(Properties.Resources.CreateDimension);
            PushButton pushbutton4 = ribbonPanel.AddItem(new PushButtonData("Check Passway Width", "通道寬度檢核", checkPasswayPath, "CheckPassway.CheckPassway")) as PushButton;
            pushbutton4.LargeImage = convertFromBitmap(Properties.Resources.CheckPassway);

            return Result.Succeeded;
        }
        /// <summary>
        /// 轉換圖片
        /// </summary>
        /// <param name="bitmap"></param>
        /// <returns></returns>
        BitmapSource convertFromBitmap(System.Drawing.Bitmap bitmap)
        {
            return System.Windows.Interop.Imaging.CreateBitmapSourceFromHBitmap(
                bitmap.GetHbitmap(),
                IntPtr.Zero,
                Int32Rect.Empty,
                BitmapSizeOptions.FromEmptyOptions());
        }
        /// <summary>
        /// 關閉
        /// </summary>
        /// <param name="a"></param>
        /// <returns></returns>
        public Result OnShutdown(UIControlledApplication a)
        {
            return Result.Succeeded;
        }
    }
}