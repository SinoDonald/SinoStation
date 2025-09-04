using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using Application = Autodesk.Revit.ApplicationServices.Application;
using Line = Autodesk.Revit.DB.Line;

namespace SinoStation_2020
{
    [Transaction(TransactionMode.Manual)]
    public class FacilityClearance : IExternalCommand
    {
        public class CrushElement
        {
            public long id = new long(); // ID
            public Element elem = null; // 主元件
            public XYZ locationPoint = new XYZ(); // 放置點座標
            public double angle { get; set; } // 旋轉角度
            public double edgeLength { get; set; } // 放置點到邊界的長度
            public List<Element> crushElems = new List<Element>(); // 衝突到的元件
            public List<Curve> modelLineCurves = new List<Curve>(); // 設施淨空模型線
        }
        public class Group
        {
            public List<Element> elemsGroup = new List<Element>();
            public List<CrushElement> crushElements = new List<CrushElement>();
            public List<Curve> curves = new List<Curve>();
            public double maxX { get; set; }
            public double minX { get; set; }
            public double maxY { get; set; }
            public double minY { get; set; }
        }
        public class CompareIdElev
        {
            public ElementId id = null; // ID
            public double elevation = new double(); // 高程
        }
        private List<LevelElevation> levelElevList = new List<LevelElevation>(); // 全部樓層
        private List<ViewPlan> allViewPlans = new List<ViewPlan>(); // 專案中所有的ViewPlans
        private double allowSpace = 0.02; // 允許間距
        private List<CrushElement> crushElements = new List<CrushElement>();
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Application app = uiapp.Application;
            Document doc = uidoc.Document;

            Sino_Station_Form sinoAR_Form = new Sino_Station_Form();
            sinoAR_Form.ShowDialog();
            if(sinoAR_Form.trueOrFalse == true)
            {
                TransactionGroup tranGrp = new TransactionGroup(doc, "設施淨空");
                tranGrp.Start();

                // 找到當前專案的Level相關資訊
                FindLevel findLevel = new FindLevel();
                Tuple<List<LevelElevation>, LevelElevation, double> multiValue = findLevel.FindDocViewLevel(doc);
                levelElevList = new List<LevelElevation>();
                this.levelElevList = multiValue.Item1; // 全部樓層

                // 找到專案中所有的ViewPlan
                List<View> allViews = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Views).WhereElementIsNotElementType().Cast<View>().ToList();
                // 只要預設的ViewPlan
                this.allViewPlans = new List<ViewPlan>();
                foreach (View view in allViews)
                {
                    if (view is ViewPlan)
                    {
                        ViewPlan viewPlan = view as ViewPlan;
                        if(viewPlan.GenLevel != null)
                        {
                            allViewPlans.Add(viewPlan);
                        }
                    }
                }

                allowSpace = sinoAR_Form.allowSpace; // 設定的間距

                // 找到專案中所有的售票機、電扶梯
                List<FamilyInstance> uuCards = new List<FamilyInstance>();
                List<FamilyInstance> storedValues = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_GenericModel).WhereElementIsNotElementType().
                                                     Where(x => x.Name.Equals("儲值卡加值機位置")).Cast<FamilyInstance>().ToList();
                List<FamilyInstance> WIPValues = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_GenericModel).WhereElementIsNotElementType().
                                                     Where(x => x.Name.Equals("自動售票機_WIP")).Cast<FamilyInstance>().ToList();
                List<FamilyInstance> storedTickets = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StructuralColumns).WhereElementIsNotElementType().
                                                     Where(x => x.Name.Equals("儲票機")).Cast<FamilyInstance>().ToList();
                foreach (FamilyInstance storedValue in storedValues)
                {
                    uuCards.Add(storedValue);
                }
                foreach (FamilyInstance WIPValue in WIPValues)
                {
                    uuCards.Add(WIPValue);
                }
                foreach (FamilyInstance storedTicket in storedTickets)
                {
                    uuCards.Add(storedTicket);
                }
                using (Transaction trans = new Transaction(doc, "售票機設施淨空"))
                {
                    // 關閉警示視窗
                    FailureHandlingOptions options = trans.GetFailureHandlingOptions();
                    MyPreProcessor preproccessor = new MyPreProcessor();
                    options.SetClearAfterRollback(true);
                    options.SetFailuresPreprocessor(preproccessor);
                    trans.SetFailureHandlingOptions(options);
                    trans.Start();
                    foreach (FamilyInstance uuCard in uuCards)
                    {
                        // 查詢LocationPoint, 要查詢模型線長的方向
                        LocationPoint lp = uuCard.Location as LocationPoint;
                        double lpX = lp.Point.X;
                        // 族群的旋轉角度
                        double angle = lp.Rotation;
                        // 找到售票機的底面邊線
                        List<Curve> curves = GetBoundarySegment(uuCard, lpX, angle);
                    }
                    // 分析群組
                    foreach (CrushElement crushElement in crushElements)
                    {
                        List<CrushElement> crushElems = crushElements.Where(x => x.elem.Id.IntegerValue != crushElement.elem.Id.IntegerValue).ToList();
                        foreach (CrushElement crushElem in crushElems)
                        {
                            double distance = Math.Sqrt((Math.Pow(crushElement.locationPoint.X - crushElem.locationPoint.X, 2) + Math.Pow(crushElement.locationPoint.Y - crushElem.locationPoint.Y, 2)));
                            double transDistance = Math.Round(UnitUtils.ConvertFromInternalUnits(distance, DisplayUnitType.DUT_METERS), 4, MidpointRounding.AwayFromZero);
                            double allowDistance = Math.Round(crushElement.edgeLength + crushElem.edgeLength + allowSpace, 4, MidpointRounding.AwayFromZero);
                            if (transDistance <= allowDistance)
                            {
                                crushElement.crushElems.Add(crushElem.elem);
                            }
                        }
                    }
                    // 各個高程
                    List<double> elevations = crushElements.Select(x => x.locationPoint.Z).Distinct().OrderBy(x => x).ToList();
                    // 分析群組
                    foreach (double elevation in elevations)
                    {
                        // 先找到同高程的點
                        List<CrushElement> sameElevCrushElems = crushElements.Where(x => x.locationPoint.Z.Equals(elevation)).ToList();
                        AnalysisGroup(doc, sameElevCrushElems);
                    }

                    trans.Commit();
                }

                // 電扶梯
                uuCards = new List<FamilyInstance>();
                IList<Element> genericModels = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_GenericModel).WhereElementIsNotElementType().ToElements();
                List<FamilyInstance> genericModelFIs = genericModels.Where(x => x is FamilyInstance).Cast<FamilyInstance>().ToList();
                List<FamilyInstance> escalators = (from x in genericModelFIs
                                                   where x.Symbol.FamilyName.Contains("電扶梯") && !x.Symbol.FamilyName.Contains("縫封")
                                                   select x).ToList();
                using (Transaction trans = new Transaction(doc, "電扶梯設施淨空"))
                {
                    trans.Start();
                    // View要3D視圖, 找到電扶梯的所有Arc
                    List<View3D> view3Ds = new FilteredElementCollector(doc).OfClass(typeof(View3D)).WhereElementIsNotElementType().Cast<View3D>().ToList();
                    View3D view3D = view3Ds.Where(x => x.Name.Equals("{3D}")).FirstOrDefault();
                    if (view3D != null)
                    {
                        crushElements = new List<CrushElement>(); // 重置要運算的模型線
                        double width = Math.Round(UnitUtils.ConvertToInternalUnits(0.85, DisplayUnitType.DUT_METERS), 4, MidpointRounding.AwayFromZero);
                        double length = Math.Round(UnitUtils.ConvertToInternalUnits(9, DisplayUnitType.DUT_METERS), 4, MidpointRounding.AwayFromZero);
                        foreach (FamilyInstance escalator in escalators)
                        {
                            try
                            {
                                Level level = escalator.Host as Level;
                                if (level == null)
                                {
                                    level = doc.GetElement(escalator.LevelId) as Level;
                                    if (level == null)
                                    {
                                        Floor floor = escalator.Host as Floor;
                                        level = doc.GetElement(floor.LevelId) as Level;
                                    }
                                }
                                // 查詢模型的ViewPlan
                                ViewPlan viewPlan = doc.GetElement(level.FindAssociatedPlanViewId()) as ViewPlan;
                                string viewPlanName = viewPlan.Title.Trim().Split(':')[0];
                                List<ViewPlan> viewPlans = (from x in allViewPlans
                                                            where x.get_Parameter(BuiltInParameter.ELEM_FAMILY_AND_TYPE_PARAM).AsValueString().Trim().Split(':')[0].Equals(viewPlanName)
                                                            select x).ToList();
                                // LocationPoint角度
                                LocationPoint lp = escalator.Location as LocationPoint;
                                double angle = lp.Rotation;
                                double transAngle = angle + (90 * Math.PI / 180);
                                double lpAngle = Math.Round(angle * 180 / Math.PI, 0, MidpointRounding.AwayFromZero);
                                // 找到電扶梯上的Arc
                                List<XYZ> arcXYZs = FindArcXYZ(view3D, level, escalator);
                                foreach (XYZ arcXYZ in arcXYZs)
                                {
                                    List<Curve> curves = new List<Curve>();
                                    // 查詢座標z軸離哪個Level最近
                                    LevelElevation closestLevel = levelElevList.Select(n => new { n, distance = Math.Abs(n.height - arcXYZ.Z) }).OrderBy(p => p.distance).First().n;
                                    Level arczlevel = closestLevel.level;
                                    // 找到同Level的樓板平面圖ViewPlan
                                    viewPlan = (from x in viewPlans
                                                where x.GenLevel.Id.IntegerValue == arczlevel.Id.IntegerValue
                                                select x).FirstOrDefault();
                                    Line axis = Line.CreateBound(arcXYZ, new XYZ(arcXYZ.X, arcXYZ.Y, arcXYZ.Z + 10));
                                    // 求方向向量
                                    double directionAngle = Math.Atan2((arcXYZ.Y - lp.Point.Y), (arcXYZ.X - lp.Point.X)) * 180 / Math.PI;
                                    XYZ start = new XYZ(arcXYZ.X - Math.Cos(angle) * width, arcXYZ.Y - Math.Sin(angle) * width, arcXYZ.Z);
                                    XYZ end = new XYZ(arcXYZ.X + Math.Cos(angle) * width, arcXYZ.Y + Math.Sin(angle) * width, arcXYZ.Z);
                                    XYZ startAdd = new XYZ(start.X + Math.Cos(transAngle) * length, start.Y + Math.Sin(transAngle) * length, start.Z);
                                    XYZ endAdd = new XYZ(end.X + Math.Cos(transAngle) * length, end.Y + Math.Sin(transAngle) * length, end.Z);
                                    XYZ startSub = new XYZ(start.X - Math.Cos(transAngle) * length, start.Y - Math.Sin(transAngle) * length, start.Z);
                                    XYZ endSub = new XYZ(end.X - Math.Cos(transAngle) * length, end.Y - Math.Sin(transAngle) * length, end.Z);
                                    Curve startCurve = Line.CreateBound(start, startSub);
                                    Curve endCurve = Line.CreateBound(end, endSub);
                                    if (lpAngle >= 0 && lpAngle < 90)
                                    {
                                        startCurve = Line.CreateBound(start, startAdd);
                                        endCurve = Line.CreateBound(end, endAdd);
                                        if (directionAngle <= 0)
                                        {
                                            startCurve = Line.CreateBound(start, startSub);
                                            endCurve = Line.CreateBound(end, endSub);
                                        }
                                    }
                                    else if (lpAngle >= 270 && lpAngle <= 360)
                                    {
                                        startCurve = Line.CreateBound(start, startAdd);
                                        endCurve = Line.CreateBound(end, endAdd);
                                        if (directionAngle <= 0)
                                        {
                                            startCurve = Line.CreateBound(start, startSub);
                                            endCurve = Line.CreateBound(end, endSub);
                                        }
                                    }
                                    else
                                    {
                                        if (directionAngle <= 0)
                                        {
                                            startCurve = Line.CreateBound(start, startAdd);
                                            endCurve = Line.CreateBound(end, endAdd);
                                        }
                                    }
                                    curves.Add(startCurve);
                                    curves.Add(endCurve);

                                    // 儲存各電扶梯的模型線
                                    CrushElement crushElem = new CrushElement();
                                    crushElem.id = escalator.Id.IntegerValue; // ID
                                    crushElem.elem = escalator; // 主元件
                                    crushElem.locationPoint = arcXYZ; // 放置點座標
                                    crushElem.angle = angle; // 旋轉角度
                                    crushElem.edgeLength = 0.85; // 放置點到邊界的長度
                                    crushElem.modelLineCurves = curves; // 設施淨空模型線
                                    crushElements.Add(crushElem);
                                }
                            }
                            catch (Exception)
                            {

                            }
                        }
                        // 各個高程
                        List<double> elevations = crushElements.Select(x => x.locationPoint.Z).Distinct().OrderBy(x => x).ToList();
                        // 分析群組
                        foreach(double elevation in elevations)
                        {
                            // 先找到同高程的點
                            List<CrushElement> sameElevCrushElems = crushElements.Where(x => x.locationPoint.Z.Equals(elevation)).ToList();
                            AnalysisGroup(doc, sameElevCrushElems);
                        }

                        doc.Regenerate();
                        uidoc.RefreshActiveView();
                        trans.Commit();
                    }
                    else
                    {
                        TaskDialog.Show("Error", "專案中無'{3D}'視圖。");
                    }
                }
                tranGrp.Assimilate();
            }

            return Result.Succeeded;
        }
        // 讀取儲值卡加值機的正面Edge
        private List<Curve> GetBoundarySegment(FamilyInstance uuCard, double lpX, double angle)
        {
            LocationPoint lp = uuCard.Location as LocationPoint;
            // 先將所有的邊儲存起來
            List<Curve> curveLoop = new List<Curve>();
            // 模型線的Curves
            List<Curve> modelLineCurves = new List<Curve>();
            try
            {
                // 1.讀取Geometry Option
                Options options = new Options();
                //options.View = doc.GetElement(room.Level.FindAssociatedPlanViewId()) as Autodesk.Revit.DB.View;
                options.DetailLevel = ViewDetailLevel.Medium;
                options.ComputeReferences = true;
                options.IncludeNonVisibleObjects = true;
                // 得到幾何元素
                GeometryElement geomElem = uuCard.get_Geometry(options);
                List<Solid> solids = GeometrySolids(geomElem);
                // 找到最大的Solid
                Solid solid = null;
                if (uuCard.Name.Equals("儲值卡加值機位置"))
                {
                    solid = solids[0];
                }
                else if (uuCard.Name.Equals("儲票機") || uuCard.Name.Equals("自動售票機_WIP"))
                {
                    solid = solids.Where(x => x.SurfaceArea.Equals(solids.Max(y => y.SurfaceArea))).FirstOrDefault();
                }
                foreach (Face face in solid.Faces)
                {
                    if (face is PlanarFace)
                    {
                        PlanarFace planarFace = face as PlanarFace;
                        if (Math.Round(planarFace.FaceNormal.Z, 2, MidpointRounding.AwayFromZero).Equals(-1))
                        {
                            foreach (EdgeArray edgeArray in planarFace.EdgeLoops)
                            {
                                foreach (Edge edge in edgeArray)
                                {
                                    Curve curve = edge.AsCurve();
                                    curveLoop.Add(curve);
                                }
                            }
                            // 找到最長的curve
                            List<Curve> maxCurves = curveLoop.Where(x => Math.Round(x.Length, 4, MidpointRounding.AwayFromZero).
                                                    Equals(curveLoop.Max(y => Math.Round(y.Length, 4, MidpointRounding.AwayFromZero)))).ToList();
                            Curve startCurve = maxCurves[0];
                            Curve otherCurve = null;
                            // 找到離LocationPoint近的那個邊
                            double distance = 999;
                            foreach (Curve shortestCurve in maxCurves)
                            {
                                if (shortestCurve.Project(lp.Point).Distance <= distance)
                                {
                                    startCurve = shortestCurve;
                                    distance = shortestCurve.Project(lp.Point).Distance;
                                }
                            }
                            if (uuCard.Name.Equals("儲票機") || uuCard.Name.Equals("自動售票機_WIP"))
                            {
                                otherCurve = maxCurves.Where(x => x != startCurve).FirstOrDefault();
                            }
                            // 找到線段的中心點
                            XYZ curveCenter = new XYZ((startCurve.Tessellate()[0].X + startCurve.Tessellate()[1].X) / 2, 
                                                      (startCurve.Tessellate()[0].Y + startCurve.Tessellate()[1].Y) / 2, 
                                                      (startCurve.Tessellate()[0].Z + startCurve.Tessellate()[1].Z) / 2);
                            // 線短起點的終點座標, 儲值機3米
                            double length = Math.Round(UnitUtils.ConvertToInternalUnits(3, DisplayUnitType.DUT_METERS), 2, MidpointRounding.AwayFromZero);
                            XYZ line1UpEnd = new XYZ(startCurve.Tessellate()[0].X, startCurve.Tessellate()[0].Y + length, startCurve.Tessellate()[0].Z);
                            XYZ line2UpEnd = new XYZ(startCurve.Tessellate()[1].X, startCurve.Tessellate()[1].Y + length, startCurve.Tessellate()[0].Z);
                            XYZ line1DownEnd = new XYZ(startCurve.Tessellate()[0].X, startCurve.Tessellate()[0].Y - length, startCurve.Tessellate()[0].Z);
                            XYZ line2DownEnd = new XYZ(startCurve.Tessellate()[1].X, startCurve.Tessellate()[1].Y - length, startCurve.Tessellate()[0].Z);
                            Line line1 = null;
                            Line line2 = null;
                            if (uuCard.Name.Equals("儲值卡加值機位置"))
                            {
                                if (curveCenter.X > lpX)
                                {
                                    if (0 <= angle && angle <= Math.PI)
                                    {
                                        line1 = Line.CreateBound(startCurve.Tessellate()[0], line1DownEnd);
                                        line2 = Line.CreateBound(startCurve.Tessellate()[1], line2DownEnd);
                                    }
                                    else
                                    {
                                        line1 = Line.CreateBound(startCurve.Tessellate()[0], line1UpEnd);
                                        line2 = Line.CreateBound(startCurve.Tessellate()[1], line2UpEnd);
                                    }
                                    modelLineCurves.Add(line1);
                                    modelLineCurves.Add(line2);
                                }
                                else if (curveCenter.X <= lpX)
                                {
                                    if (0 <= angle && angle <= Math.PI)
                                    {
                                        line1 = Line.CreateBound(startCurve.Tessellate()[0], line1UpEnd);
                                        line2 = Line.CreateBound(startCurve.Tessellate()[1], line2UpEnd);
                                    }
                                    else
                                    {
                                        line1 = Line.CreateBound(startCurve.Tessellate()[0], line1DownEnd);
                                        line2 = Line.CreateBound(startCurve.Tessellate()[1], line2DownEnd);
                                    }
                                    modelLineCurves.Add(line1);
                                    modelLineCurves.Add(line2);
                                }
                            }
                            else if (uuCard.Name.Equals("自動售票機_WIP"))
                            {
                                curveCenter = new XYZ((otherCurve.Tessellate()[0].X + otherCurve.Tessellate()[1].X) / 2,
                                                      (otherCurve.Tessellate()[0].Y + otherCurve.Tessellate()[1].Y) / 2,
                                                      (otherCurve.Tessellate()[0].Z + otherCurve.Tessellate()[1].Z) / 2);
                                if (curveCenter.X >= lpX)
                                {
                                    if (Math.Round(angle / Math.PI, 8, MidpointRounding.AwayFromZero) == 0 ||
                                        Math.Round(angle / Math.PI, 8, MidpointRounding.AwayFromZero) == 2 ||
                                        angle / Math.PI >= 1)
                                    {
                                        line1 = Line.CreateBound(startCurve.Tessellate()[0], line1DownEnd);
                                        line2 = Line.CreateBound(startCurve.Tessellate()[1], line2DownEnd);
                                    }
                                    else
                                    {
                                        line1 = Line.CreateBound(startCurve.Tessellate()[0], line1UpEnd);
                                        line2 = Line.CreateBound(startCurve.Tessellate()[1], line2UpEnd);
                                    }
                                    modelLineCurves.Add(line1);
                                    modelLineCurves.Add(line2);
                                }
                                else if (curveCenter.X < lpX)
                                {
                                    if (Math.Round(angle / Math.PI, 8, MidpointRounding.AwayFromZero) == 0 ||
                                        Math.Round(angle / Math.PI, 8, MidpointRounding.AwayFromZero) == 2 ||
                                        angle / Math.PI >= 1)
                                    {
                                        line1 = Line.CreateBound(startCurve.Tessellate()[0], line1UpEnd);
                                        line2 = Line.CreateBound(startCurve.Tessellate()[1], line2UpEnd);
                                    }
                                    else
                                    {
                                        line1 = Line.CreateBound(startCurve.Tessellate()[0], line1DownEnd);
                                        line2 = Line.CreateBound(startCurve.Tessellate()[1], line2DownEnd);
                                    }
                                    modelLineCurves.Add(line1);
                                    modelLineCurves.Add(line2);
                                }
                            }
                            else if (uuCard.Name.Equals("儲票機"))
                            {
                                curveCenter = new XYZ((otherCurve.Tessellate()[0].X + otherCurve.Tessellate()[1].X) / 2,
                                                      (otherCurve.Tessellate()[0].Y + otherCurve.Tessellate()[1].Y) / 2,
                                                      (otherCurve.Tessellate()[0].Z + otherCurve.Tessellate()[1].Z) / 2);
                                if (curveCenter.X >= lpX)
                                {
                                    if (Math.Round(angle / Math.PI, 8, MidpointRounding.AwayFromZero) == 0 ||
                                        Math.Round(angle / Math.PI, 8, MidpointRounding.AwayFromZero) == 2 ||
                                        angle / Math.PI >= 1)
                                    {
                                        line1 = Line.CreateBound(startCurve.Tessellate()[0], line1DownEnd);
                                        line2 = Line.CreateBound(startCurve.Tessellate()[1], line2DownEnd);
                                    }
                                    else
                                    {
                                        line1 = Line.CreateBound(startCurve.Tessellate()[0], line1UpEnd);
                                        line2 = Line.CreateBound(startCurve.Tessellate()[1], line2UpEnd);
                                    }
                                    modelLineCurves.Add(line1);
                                    modelLineCurves.Add(line2);
                                }
                                else if (curveCenter.X < lpX)
                                {
                                    if (Math.Round(angle / Math.PI, 8, MidpointRounding.AwayFromZero) == 0 ||
                                        Math.Round(angle / Math.PI, 8, MidpointRounding.AwayFromZero) == 2 ||
                                        angle / Math.PI >= 1)
                                    {
                                        line1 = Line.CreateBound(startCurve.Tessellate()[0], line1UpEnd);
                                        line2 = Line.CreateBound(startCurve.Tessellate()[1], line2UpEnd);
                                    }
                                    else
                                    {
                                        line1 = Line.CreateBound(startCurve.Tessellate()[0], line1DownEnd);
                                        line2 = Line.CreateBound(startCurve.Tessellate()[1], line2DownEnd);
                                    }
                                    modelLineCurves.Add(line1);
                                    modelLineCurves.Add(line2);
                                }
                            }
                            CrushElement crushElem = new CrushElement();
                            crushElem.elem = uuCard; // 主元件
                            crushElem.locationPoint = lp.Point; // 放置點座標
                            crushElem.angle = lp.Rotation; // 旋轉角度
                            crushElem.edgeLength = UnitUtils.ConvertFromInternalUnits(startCurve.Length / 2, DisplayUnitType.DUT_METERS); // 放置點到邊界的長度
                            crushElem.modelLineCurves = modelLineCurves; // 設施淨空模型線
                            crushElements.Add(crushElem);
                            break;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                string error = ex.Message + "\n" + ex.ToString();
            }

            return modelLineCurves;
        }
        // 取得Room的Solid
        private static List<Solid> GeometrySolids(GeometryObject geoObj)
        {
            List<Solid> solids = new List<Solid>();
            if (geoObj is Solid)
            {
                Solid solid = (Solid)geoObj;
                if (solid.Faces.Size > 0/* && solid.Volume > 0*/)
                {
                    solids.Add(solid);
                }
            }
            if (geoObj is GeometryInstance)
            {
                GeometryInstance geoIns = geoObj as GeometryInstance;
                GeometryElement geometryElement = (geoObj as GeometryInstance).GetSymbolGeometry(geoIns.Transform); // 座標轉換
                foreach (GeometryObject o in geometryElement)
                {
                    solids.AddRange(GeometrySolids(o));
                }
            }
            else if (geoObj is GeometryElement)
            {
                GeometryElement geometryElement2 = (GeometryElement)geoObj;
                foreach (GeometryObject o in geometryElement2)
                {
                    solids.AddRange(GeometrySolids(o));
                }
            }
            return solids;
        }
        // 分析群組
        private void AnalysisGroup(Document doc, List<CrushElement> crushElements)
        {
            // 已比對過的ElementId與高程
            List<CompareIdElev> compareIdElevList = new List<CompareIdElev>();
            foreach(CrushElement crushElement in crushElements)
            {
                // 如果ElementId、高程尚未比對過的情況下
                int a = compareIdElevList.Where(x => x.id.IntegerValue.Equals(crushElement.elem.Id.IntegerValue)).Where(x => x.elevation.Equals(crushElement.locationPoint.Z)).Count();
                if (a.Equals(0))
                {
                    Group group = new Group();
                    group.elemsGroup.Add(crushElement.elem);
                    foreach (Curve curve in crushElement.modelLineCurves)
                    {
                        group.curves.Add(curve);
                    }
                    // 比對
                    compareIdElevList = CompareElems(compareIdElevList, crushElement, crushElements, group);
                    // 座標連線                    
                    Level level = doc.GetElement(crushElement.elem.LevelId) as Level;
                    if (level == null)
                    {
                        FamilyInstance escalator = crushElement.elem as FamilyInstance;
                        try
                        {
                            Floor floor = escalator.Host as Floor;
                            level = doc.GetElement(floor.LevelId) as Level;
                        }
                        catch (NullReferenceException)
                        {
                            level = escalator.Host as Level;
                        }
                    }
                    // 查詢座標z軸離哪個Level最近
                    LevelElevation closestLevel = levelElevList.Select(n => new { n, distance = Math.Abs(n.height - crushElement.locationPoint.Z) }).OrderBy(p => p.distance).First().n;
                    Level arczlevel = closestLevel.level;
                    // 找到同Level的樓板平面圖ViewPlan
                    ViewPlan viewPlan = (from x in allViewPlans
                                         where x.GenLevel.Id.IntegerValue == arczlevel.Id.IntegerValue
                                         select x).FirstOrDefault();
                    //ViewPlan viewPlan = doc.GetElement(level.FindAssociatedPlanViewId()) as ViewPlan; // 查詢模型的ViewPlan
                    // 找到族群所擺放的視圖與sketchPlane與sketchPlane的z值
                    SketchPlane sketchPlane = viewPlan.SketchPlane;
                    double sketchPlaneZ = sketchPlane.GetPlane().Origin.Z;
                    // 找到距離最遠的兩個邊
                    Curve curve1 = group.curves[0];
                    Curve curve2 = group.curves[1];
                    double maxDistance = 0.0;
                    foreach (Curve firstCurve in group.curves)
                    {
                        XYZ firstCurveCP = new XYZ((firstCurve.Tessellate()[0].X + firstCurve.Tessellate()[1].X) / 2,
                                                   (firstCurve.Tessellate()[0].Y + firstCurve.Tessellate()[1].Y) / 2,
                                                   (firstCurve.Tessellate()[0].Z + firstCurve.Tessellate()[1].Z) / 2);
                        foreach (Curve secondCurve in group.curves)
                        {
                            XYZ secondCurveCP = new XYZ((secondCurve.Tessellate()[0].X + secondCurve.Tessellate()[1].X) / 2,
                                                        (secondCurve.Tessellate()[0].Y + secondCurve.Tessellate()[1].Y) / 2,
                                                        (secondCurve.Tessellate()[0].Z + secondCurve.Tessellate()[1].Z) / 2);
                            double distance = Math.Sqrt((Math.Pow(firstCurveCP.X - secondCurveCP.X, 2) + Math.Pow(firstCurveCP.Y - secondCurveCP.Y, 2)));
                            if(distance > maxDistance)
                            {
                                curve1 = firstCurve;
                                curve2 = secondCurve;
                                maxDistance = distance;
                            }
                        }
                    }
                    List<Curve> maxDisCurves = new List<Curve>();
                    maxDisCurves.Add(curve1);
                    maxDisCurves.Add(curve2);
                    double angle = crushElement.angle;
                    // 畫模型線
                    DrawingModelLine(doc, viewPlan, crushElement, maxDisCurves, angle);
                }
            }
        }
        // 比對
        private List<CompareIdElev> CompareElems(List<CompareIdElev> compareIdElevList, CrushElement crushElement, List<CrushElement> crushElements, Group group)
        {
            List<CrushElement> crushElems = new List<CrushElement>();
            foreach (CrushElement otherCrushElem in crushElements)
            {
                int b = compareIdElevList.Where(x => x.id.IntegerValue.Equals(otherCrushElem.elem.Id.IntegerValue)).Where(x => x.elevation.Equals(crushElement.locationPoint.Z)).Count();
                if (b.Equals(0))
                {
                    crushElems.Add(otherCrushElem);
                }
            }
            foreach (CrushElement crushElem in crushElems)
            {
                double distance = Math.Sqrt((Math.Pow(crushElement.locationPoint.X - crushElem.locationPoint.X, 2) + Math.Pow(crushElement.locationPoint.Y - crushElem.locationPoint.Y, 2)));
                double transDistance = Math.Round(UnitUtils.ConvertFromInternalUnits(distance, DisplayUnitType.DUT_METERS), 4, MidpointRounding.AwayFromZero);
                double allowDistance = Math.Round(crushElement.edgeLength + crushElem.edgeLength + allowSpace, 4, MidpointRounding.AwayFromZero);
                if (transDistance <= allowDistance)
                {
                    int count = compareIdElevList.Where(x => x.id.IntegerValue.Equals(crushElem.elem.Id.IntegerValue)).Where(x => x.elevation.Equals(crushElem.locationPoint.Z)).Count();
                    if (count.Equals(0))
                    {
                        group.elemsGroup.Add(crushElem.elem);
                        foreach (Curve curve in crushElem.modelLineCurves)
                        {
                            group.curves.Add(curve);
                        }
                        CompareIdElev compareIdElev = new CompareIdElev();
                        compareIdElev.id = crushElem.elem.Id;
                        compareIdElev.elevation = crushElem.locationPoint.Z;
                        compareIdElevList.Add(compareIdElev);
                    }
                    CompareElems(compareIdElevList, crushElem, crushElements, group);
                }
            }
            return compareIdElevList;
        }
        // 找到電扶梯的Arc
        private List<XYZ> FindArcXYZ(View3D view3D, Level chooseLevel, Element geoElem)
        {
            List<XYZ> arcXYZs = new List<XYZ>();
            double elevation = Math.Round(chooseLevel.Elevation, 4, MidpointRounding.AwayFromZero); // 高程

            // 1.讀取Geometry Option
            Options options = new Options();
            options.View = view3D;
            options.ComputeReferences = true;
            options.IncludeNonVisibleObjects = true;
            // 得到幾何元素
            GeometryElement geomElem = geoElem.get_Geometry(options);
            List<Arc> geomArcs = GeometryArcs(geomElem);
            foreach(Arc arc in geomArcs)
            {
                arcXYZs.Add(arc.Center);
            }

            return arcXYZs;
        }
        // 取得Room的Arc
        private static List<Arc> GeometryArcs(GeometryObject geoObj)
        {
            List<Arc> arcs = new List<Arc>();
            if (geoObj is Arc)
            {
                Arc solid = (Arc)geoObj;
                arcs.Add(solid);
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
        // 繪製模型線
        private void DrawingModelLine(Document doc, ViewPlan viewPlan, CrushElement crushElement, List<Curve> curves, double angle)
        {
            // 找到族群所擺放的視圖與sketchPlane與sketchPlane的z值
            SketchPlane sketchPlane = viewPlan.SketchPlane;
            double sketchPlaneZ = sketchPlane.GetPlane().Origin.Z;
            try
            {
                List<XYZ> line3Points = new List<XYZ>();
                // 將Curve繪製到平面上
                foreach (Curve curve in curves)
                {
                    XYZ start = new XYZ(curve.Tessellate()[0].X, curve.Tessellate()[0].Y, sketchPlaneZ);
                    XYZ end = new XYZ(curve.Tessellate()[1].X, curve.Tessellate()[1].Y, sketchPlaneZ);
                    Line geomLine = Line.CreateBound(start, end);
                    try
                    {
                        ModelLine modelLine = doc.Create.NewModelCurve(geomLine, sketchPlane) as ModelLine;
                        LocationCurve lc = modelLine.Location as LocationCurve;
                        Line axis = Line.CreateBound(lc.Curve.Tessellate()[0], new XYZ(lc.Curve.Tessellate()[0].X, lc.Curve.Tessellate()[0].Y, lc.Curve.Tessellate()[0].Z + 10));
                        FamilyInstance fi = crushElement.elem as FamilyInstance;
                        if (!fi.Symbol.FamilyName.Contains("電扶梯"))
                        {
                            ElementTransformUtils.RotateElement(doc, modelLine.Id, axis, angle);
                        }
                        line3Points.Add(lc.Curve.Tessellate()[1]);
                    }
                    catch (Autodesk.Revit.Exceptions.ArgumentException ex)
                    {
                        string info = ex.Message + "\n" + ex.ToString();
                    }
                }
                Line line3 = Line.CreateBound(line3Points[0], line3Points[1]);
                ModelLine modelLine3 = doc.Create.NewModelCurve(line3, sketchPlane) as ModelLine;
            }
            catch (Exception ex)
            {
                string error = ex.Message + "\n" + ex.ToString();
            }
        }
        // 關閉警示視窗
        public class MyPreProcessor : IFailuresPreprocessor
        {
            FailureProcessingResult IFailuresPreprocessor.PreprocessFailures(FailuresAccessor failuresAccessor)
            {
                String transactionName = failuresAccessor.GetTransactionName();
                IList<FailureMessageAccessor> fmas = failuresAccessor.GetFailureMessages();
                if (fmas.Count == 0)
                {
                    return FailureProcessingResult.Continue;
                }
                if (transactionName.Equals("CloseWarning"))
                {
                    foreach (FailureMessageAccessor fma in fmas)
                    {
                        if (fma.GetSeverity() == FailureSeverity.Error)
                        {
                            failuresAccessor.DeleteAllWarnings();
                            return FailureProcessingResult.ProceedWithRollBack;
                        }
                        else
                        {
                            failuresAccessor.DeleteWarning(fma);
                        }
                    }
                }
                else
                {
                    foreach (FailureMessageAccessor fma in fmas)
                    {
                        failuresAccessor.DeleteAllWarnings();
                    }
                }

                return FailureProcessingResult.Continue;
            }
        }
    }
}
