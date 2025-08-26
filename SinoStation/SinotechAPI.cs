using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Windows;
using System.Windows.Media.Imaging;

namespace SinoStation
{
    public class SinotechAPI : IExternalApplication
    {
        public string addinAssmeblyPath = Assembly.GetExecutingAssembly().Location;
        public Result OnStartup(UIControlledApplication a)
        {
            string sinoStationPath = Path.Combine(Directory.GetParent(addinAssmeblyPath).FullName, "SinoStation.dll");
            //string sinoStationPath = Path.Combine(Directory.GetParent(addinAssmeblyPath).FullName, "Sino_Station_Old.dll");
            string checkPlatformPath = Path.Combine(Directory.GetParent(addinAssmeblyPath).FullName, "CheckPlatform.dll");
            string checkPasswayPath = Path.Combine(Directory.GetParent(addinAssmeblyPath).FullName, "CheckPassway.dll");
            string autoSignPath = Path.Combine(Directory.GetParent(addinAssmeblyPath).FullName, "AutoSign.dll");
            string createModelPath = Path.Combine(Directory.GetParent(addinAssmeblyPath).FullName, "AutoCreateModel.dll");

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

            PushButton pushbutton1 = ribbonPanel.AddItem(new PushButtonData("SinoStation", "房間", sinoStationPath, "SinoStation.RegulatoryReview")) as PushButton;
            pushbutton1.LargeImage = convertFromBitmap(Properties.Resources.房間檢討);
            PushButton pushbutton2 = ribbonPanel.AddItem(new PushButtonData("Check Platform", "月台", checkPlatformPath, "CheckPlatform.CheckPlatform")) as PushButton;
            pushbutton2.LargeImage = convertFromBitmap(Properties.Resources.CheckPlatform);
            PushButton pushbutton3 = ribbonPanel.AddItem(new PushButtonData("Use Dimension", "通道尺寸標註", checkPasswayPath, "CheckPassway.CreateDimensionType")) as PushButton;
            pushbutton3.LargeImage = convertFromBitmap(Properties.Resources.CreateDimension);
            PushButton pushbutton4 = ribbonPanel.AddItem(new PushButtonData("Check Passway Width", "通道寬度檢核", checkPasswayPath, "CheckPassway.CheckPassway")) as PushButton;
            pushbutton4.LargeImage = convertFromBitmap(Properties.Resources.CheckPassway);

            //try { a.CreateRibbonTab("捷運規範校核"); } catch { }
            //try { ribbonPanel = a.CreateRibbonPanel("捷運規範校核", "自動建模"); }
            //catch
            //{
            //    List<RibbonPanel> panel_list = new List<RibbonPanel>();
            //    panel_list = a.GetRibbonPanels("自動建模");
            //    foreach (RibbonPanel rp in panel_list)
            //    {
            //        if (rp.Name == "建立廁所")
            //        {
            //            ribbonPanel = rp;
            //        }
            //    }
            //}
            //// 在面板上添加一個按鈕, 點擊此按鈕觸動AutoCreateModel.CreateMRT
            //PushButton createToiletBtn = ribbonPanel.AddItem(new PushButtonData("AutoCreateModel", "建立廁所", createModelPath, "AutoCreateModel.CreateMRT")) as PushButton;
            //createToiletBtn.LargeImage = convertFromBitmap(Properties.Resources.Toilet);

            //try { a.CreateRibbonTab("捷運規範校核"); } catch { }
            //try { ribbonPanel = a.CreateRibbonPanel("捷運規範校核", "指標自動化"); }
            //catch
            //{
            //    List<RibbonPanel> panel_list = new List<RibbonPanel>();
            //    panel_list = a.GetRibbonPanels("指標自動化");
            //    foreach (RibbonPanel rp in panel_list)
            //    {
            //        if (rp.Name == "指標校核")
            //        {
            //            ribbonPanel = rp;
            //        }
            //    }
            //}
            //// 在面板上添加一個按鈕, 點擊此按鈕觸動AutoSign.AutoSign
            //PushButton autoSignBtn = ribbonPanel.AddItem(new PushButtonData("AutoSign", "指標校核", autoSignPath, "AutoSign.AutoSign")) as PushButton;
            //autoSignBtn.LargeImage = convertFromBitmap(Properties.Resources.指標自動化);
            //PushButton autoNumberBtn = ribbonPanel.AddItem(new PushButtonData("AutoNumber", "自動編號", autoSignPath, "AutoSign.AutoNumber")) as PushButton;
            //autoNumberBtn.LargeImage = convertFromBitmap(Properties.Resources.自動編號);
            //PushButton autoTagBtn = ribbonPanel.AddItem(new PushButtonData("AutoTag", "自動標籤", autoSignPath, "AutoSign.CreateIndependentTag")) as PushButton;
            //autoTagBtn.LargeImage = convertFromBitmap(Properties.Resources.自動標籤);

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
        ///// <summary>
        ///// 轉換單位
        ///// </summary>
        ///// <param name="number"></param>
        ///// <param name="unit"></param>
        ///// <returns></returns>        
        //public static double ConvertFromInternalUnits(double number, string unit)
        //{
        //    if (unit.Equals("meters"))
        //    {
        //        number = UnitUtils.ConvertFromInternalUnits(number, DisplayUnitType.DUT_METERS); // 2020
        //        //number = UnitUtils.ConvertFromInternalUnits(number, UnitTypeId.Meters); // 2022
        //    }
        //    else if (unit.Equals("millimeters"))
        //    {
        //        number = UnitUtils.ConvertFromInternalUnits(number, DisplayUnitType.DUT_MILLIMETERS); // 2020
        //        //number = UnitUtils.ConvertFromInternalUnits(number, UnitTypeId.Millimeters); // 2022
        //    }
        //    return number;
        //}
        ///// <summary>
        ///// 轉換單位
        ///// </summary>
        ///// <param name="number"></param>
        ///// <param name="unit"></param>
        ///// <returns></returns>
        //public static double ConvertToInternalUnits(double number, string unit)
        //{
        //    if (unit.Equals("meters"))
        //    {
        //        number = UnitUtils.ConvertToInternalUnits(number, DisplayUnitType.DUT_METERS); // 2020
        //        //number = UnitUtils.ConvertToInternalUnits(number, UnitTypeId.Meters); // 2022
        //    }
        //    else if (unit.Equals("millimeters"))
        //    {
        //        number = UnitUtils.ConvertToInternalUnits(number, DisplayUnitType.DUT_MILLIMETERS); // 2020
        //        //number = UnitUtils.ConvertToInternalUnits(number, UnitTypeId.Millimeters); // 2022
        //    }
        //    return number;
        //}
    }
}