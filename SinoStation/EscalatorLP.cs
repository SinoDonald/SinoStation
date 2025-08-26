using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.DB.Structure;

namespace SinoStation
{
    [Transaction(TransactionMode.Manual)]
    internal class EscalatorLP : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Application app = uiapp.Application;
            Document doc = uidoc.Document;

            // 找到當前專案的Level相關資訊
            FindLevel findLevel = new FindLevel();
            Tuple<List<LevelElevation>, LevelElevation, double> multiValue = findLevel.FindDocViewLevel(doc);
            List<LevelElevation> levelElevList = multiValue.Item1; // 全部樓層

            FamilyInstance escalator = doc.GetElement(uidoc.Selection.PickObject(Autodesk.Revit.UI.Selection.ObjectType.Element).ElementId) as FamilyInstance;
            // 取得Arc
            using (Transaction trans = new Transaction(doc, "新增逃生通過線"))
            {
                trans.Start(); 
                List<Element> familySymbols = new FilteredElementCollector(doc).OfClass(typeof(FamilySymbol)).ToElements().ToList();
                FamilySymbol symbol = familySymbols.Where(x => x.Name.Equals("逃生通過線")).FirstOrDefault() as FamilySymbol;
                if (symbol != null)
                {
                    // 如果FamilySymbol尚未啟動, 必須啟用才能使用使用
                    if (!symbol.IsActive)
                    {
                        symbol.Activate();
                        doc.Regenerate();
                    }
                }
                List<Arc> arcList = GetArc(doc, escalator);
                foreach (Arc arc in arcList)
                {
                    LevelElevation arcLevel = levelElevList.Select(n => new { n, distance = Math.Abs(n.height - arc.Center.Z) }).OrderBy(p => p.distance).First().n;
                    Level level = arcLevel.level;

                    // 放置逃生通過線
                    XYZ arcXYZ1 = new XYZ(arc.Center.X, arc.Center.Y - 2, arc.Center.Z);
                    XYZ arcXYZ2 = new XYZ(arc.Center.X, arc.Center.Y + 2, arc.Center.Z);
                    Curve curve = Line.CreateBound(arcXYZ1, arcXYZ2);
                    FamilyInstance instance = doc.Create.NewFamilyInstance(curve, symbol, level, StructuralType.NonStructural);
                }
                trans.Commit();
            }

            return Result.Succeeded;
        }
        // 取得Arc
        private List<Arc> GetArc(Document doc, FamilyInstance escalator)
        {
            List<Arc> arcList = new List<Arc>();
            try
            {
                // 1.讀取Geometry Option
                Options options = new Options();
                //options.View = doc.GetElement(room.Level.FindAssociatedPlanViewId()) as Autodesk.Revit.DB.View;
                //options.DetailLevel = ViewDetailLevel.Medium;
                options.ComputeReferences = true;
                options.IncludeNonVisibleObjects = true;
                options.View = doc.ActiveView;
                // 得到幾何元素
                GeometryElement geomElem = escalator.get_Geometry(options);
                arcList = GeometryArcs(geomElem);
            }
            catch (Exception ex)
            {
                string error = ex.Message + "\n" + ex.ToString();
            }

            return arcList;
        }
        private static List<Arc> GeometryArcs(GeometryObject geoObj)
        {
            List<Arc> arcs = new List<Arc>();
            if (geoObj is Arc)
            {
                Arc arc = (Arc)geoObj;
                arcs.Add(arc);
            }
            if (geoObj is GeometryInstance)
            {
                GeometryInstance geoIns = geoObj as GeometryInstance;
                GeometryElement geometryElement = (geoObj as GeometryInstance).GetSymbolGeometry(geoIns.Transform); // 座標轉換
                foreach (GeometryObject o in geometryElement)
                {
                    arcs.AddRange(GeometryArcs(o));
                }
            }
            else if (geoObj is GeometryElement)
            {
                GeometryElement geometryElement2 = (GeometryElement)geoObj;
                foreach (GeometryObject o in geometryElement2)
                {
                    arcs.AddRange(GeometryArcs(o));
                }
            }
            return arcs;
        }
    }
}
