using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using static SinoStation_2020.RegulatoryReviewForm;

namespace SinoStation_2020
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class EditRoomPara : IExternalEventHandler
    {
        List<DoorData> doorDataList = new List<DoorData>();
        public void Execute(UIApplication uiapp)
        {
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            doorDataList = new List<DoorData>();
            doorDataList = RegulatoryReviewForm.doorDataList;

            using (Transaction trans = new Transaction(doc, "更新歸屬房間"))
            {
                trans.Start();
                foreach (DoorData doorData in doorDataList)
                {
                    // ID的門
                    Element door = doc.GetElement(doorData.id);
                    try
                    {
                        // 找到Room寫入參數
                        Parameter para = door.LookupParameter("Room");
                        if(para == null)
                        {
                            // 找到備註寫入參數
                            para = door.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS);
                        }
                        para.Set(doorData.roomName);
                    }
                    catch (Exception)
                    {
                        
                    }
                }
                TaskDialog.Show("Revit", "更新完成");
                trans.Commit();
            }
        }
        public string GetName()
        {
            return "Event handler is working now!!";
        }
    }
}
