using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;
using System;

namespace SinoStation
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    [Journaling(JournalingMode.NoCommandData)]
    public class RegulatoryReview : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            IExternalEventHandler handler_EditRoomOffset = new EditRoomOffset();
            ExternalEvent externalEvent_EditRoomOffset = ExternalEvent.Create(handler_EditRoomOffset);
            IExternalEventHandler handler_EditRoomPara = new EditRoomPara();
            ExternalEvent externalEvent_EditRoomPara = ExternalEvent.Create(handler_EditRoomPara);
            IExternalEventHandler handler_CreateCrushElems = new CreateCrushElems();
            ExternalEvent externalEvent_CreateCrushElems = ExternalEvent.Create(handler_CreateCrushElems);
            //commandData.Application.Idling += Application_Idling;
            RevitDocument m_connect = new RevitDocument(commandData.Application);
            RegulatoryReviewForm regulatoryReviewform = new RegulatoryReviewForm(commandData.Application, m_connect, externalEvent_EditRoomOffset, externalEvent_EditRoomPara, externalEvent_CreateCrushElems);
            if (regulatoryReviewform.trueOrFalse) { regulatoryReviewform.Show(); externalEvent_CreateCrushElems.Raise(); }

            return Result.Succeeded;
        }
        private void Application_Idling(object sender, IdlingEventArgs e)
        {
            throw new NotImplementedException();
        }
    }
}
