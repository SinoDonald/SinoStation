using Autodesk.Revit;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace SinoStation_2025
{
    public class RevitDocument
    {
        private UIDocument m_revitDoc;
        private Autodesk.Revit.Creation.Application m_appCreator;
        public UIDocument RevitDoc
        {
            get
            {
                return m_revitDoc;
            }
        }
        public RevitDocument(Autodesk.Revit.UI.UIApplication app)
        {
            m_revitDoc = app.ActiveUIDocument;
            m_appCreator = app.Application.Create;
        }
    }
}
