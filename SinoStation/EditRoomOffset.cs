using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using static SinoStation.RegulatoryReviewForm;

namespace SinoStation
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class EditRoomOffset : IExternalEventHandler
    {
        List<ElementInfo> elementInfoList = new List<ElementInfo>();
        public void Execute(UIApplication uiapp)
        {
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            elementInfoList = new List<ElementInfo>();
            elementInfoList = RegulatoryReviewForm.elementInfoList;

            using (Transaction trans = new Transaction(doc, "修改Room高度"))
            {
                trans.Start();
                foreach (ElementInfo elementInfo in elementInfoList)
                {
                    try
                    {
                        double offset = elementInfo.roomHeight;
                        double roomLowerOffset = Convert.ToDouble(elementInfo.elem.get_Parameter(BuiltInParameter.ROOM_LOWER_OFFSET).AsValueString()); // 基準偏移
                        if (roomLowerOffset >= 0)
                        {
                            offset = elementInfo.roomHeight - roomLowerOffset;
                        }
                        else
                        {
                            offset = elementInfo.roomHeight + roomLowerOffset;
                        }
                        if (offset > 0)
                        {
                            Parameter roomUpperOffset = elementInfo.elem.get_Parameter(BuiltInParameter.ROOM_UPPER_OFFSET); // 限制偏移
                            roomUpperOffset.SetValueString(offset.ToString());
                        }
                        else
                        {
                            Parameter roomUpperOffset = elementInfo.elem.get_Parameter(BuiltInParameter.ROOM_UPPER_OFFSET); // 限制偏移
                            roomUpperOffset.SetValueString("4.5"); // 如果沒有偵測到上方樓板, 預設限制偏移4.5
                        }
                    }
                    catch (Exception ex)
                    {
                        string error = ex.Message + "\n" + ex.ToString();
                    }
                }
                trans.Commit();
            }
        }
        public string GetName()
        {
            return "Event handler is working now!!";
        }
    }
}
