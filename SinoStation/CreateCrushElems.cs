using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using static SinoStation.RegulatoryReviewForm;

namespace SinoStation
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class CreateCrushElems : IExternalEventHandler
    {
        UIDocument revitUIDoc = null;
        Document revitDoc = null;
        DataGridView dgv1 = null;
        DataGridView dgv2 = null;
        DataGridView dgv6 = null;
        List<string> createSoildsFailed = new List<string>();
        double extrusionHeight = 5.0; // 擠出干涉元件的高度(m)

        public class CrushElemInfo
        {
            public string id { get; set; } // id
            public string hostName { get; set; } // 房間名稱
            public List<Element> crushElems = new List<Element>(); // 干涉到所有的樓板與天花板
            public List<Floor> crushFloors = new List<Floor>(); // 干涉到所有的樓板
            public List<Ceiling> crushCeilings = new List<Ceiling>(); // 干涉到所有的天花板
            public List<Stairs> crushStairs = new List<Stairs>(); // 干涉到所有的樓梯
            public List<FamilyInstance> crushEscalators = new List<FamilyInstance>(); // 干涉到所有的電扶梯
            public double roomHeight { get; set; } // 淨高
            public double clearSpace { get; set; } // 淨空空間
            public double clearSpaceHeight { get; set; } // 天花內淨高之淨高
            public List<string> crushElemName = new List<string>();
        }

        public void Execute(UIApplication uiapp)
        {
            revitUIDoc = uiapp.ActiveUIDocument;
            revitDoc = revitUIDoc.Document;

            List<Solid> createSolids = new List<Solid>(); // 儲存所有干涉模型

            // 找到當前專案的Level相關資訊
            FindLevel findLevel = new FindLevel();
            Tuple<List<LevelElevation>, LevelElevation, double> multiValue = findLevel.FindDocViewLevel(revitDoc);
            List<LevelElevation> levelElevList = multiValue.Item1; // 全部樓層
            double elevation = Math.Round(UnitUtils.ConvertToInternalUnits(extrusionHeight, UnitTypeId.Meters), 2, MidpointRounding.AwayFromZero);

            // 樓板的材料ID
            Element genericModel = new FilteredElementCollector(revitDoc).OfCategory(BuiltInCategory.OST_GenericModel).WhereElementIsNotElementType().FirstOrDefault();

            TransactionGroup tranGrp = new TransactionGroup(revitDoc, "房間空間校核");
            tranGrp.Start();
            using (Transaction trans = new Transaction(revitDoc, "干涉檢查"))
            {
                trans.Start();
                foreach (ElementInfo elementInfo in elementInfoList)
                {
                    //int index = levelElevList.IndexOf(levelElevList.Where(x => x.name.Equals(elementInfo.level)).Select(x => x).FirstOrDefault()); // 房間樓層
                    //// 查詢上一層與當前樓層差距
                    //double levelElev = levelElevList[index].elevation;
                    //double nextLevelElev = levelElevList[index].elevation + 10;
                    //if (index < levelElevList.Count() - 2) { nextLevelElev = levelElevList[index + 2].elevation; }
                    //else if (index == levelElevList.Count() - 2) { nextLevelElev = levelElevList[index + 1].elevation; }
                    //double elevation = Math.Round(UnitUtils.ConvertToInternalUnits(nextLevelElev - levelElev, DisplayUnitType.DUT_METERS), 2, MidpointRounding.AwayFromZero);
                    //elevation = nextLevelElev - levelElev;

                    //int elementId = Convert.ToInt32(elementInfo.id);
                    ElementId elementId = new ElementId(Convert.ToInt64(elementInfo.id));
                    foreach (Face topFace in elementInfo.bottomFaces)
                    {
                        try
                        {
                            Solid createCrushSolid = CreateCrushSolids(revitDoc, topFace, elevation, genericModel.Category.Id, elementId);
                            if (createCrushSolid != null) { createSolids.Add(createCrushSolid); }
                        }
                        catch(Exception ex) { string error = ex.Message + "\n" + ex.ToString(); createSoildsFailed.Add(elementId.ToString()); }
                    }
                }
                revitDoc.Regenerate();
                revitUIDoc.RefreshActiveView();
                trans.Commit();
            }

            IList<ElementFilter> elementFilters = new List<ElementFilter>(); // 清空過濾器
            ElementCategoryFilter genericModelFilter = new ElementCategoryFilter(BuiltInCategory.OST_GenericModel); // 一般模型
            elementFilters.Add(genericModelFilter);
            LogicalOrFilter logicalOrFilter = new LogicalOrFilter(elementFilters);
            List<CrushElemInfo> crushElemInfos = CrushReport(revitDoc, logicalOrFilter); // 出衝突報告
            RemoveElems(revitDoc, logicalOrFilter); // 移除干涉元件
            tranGrp.Assimilate();

            createSoildsFailed = createSoildsFailed.Distinct().ToList();

            // 房間規範檢討
            dgv1 = RegulatoryReviewForm.dgv1;
            foreach(DataGridViewRow row in dgv1.Rows)
            {
                try
                {
                    string id = row.Cells[row.Cells.Count - 1].Value.ToString();
                    CrushElemInfo crushElemInfo = crushElemInfos.Where(x => x.id.Equals(id)).FirstOrDefault();
                    if (crushElemInfo != null) 
                    {
                        if (crushElemInfo.roomHeight.Equals(0))
                        {
                            if (crushElemInfo.crushElems.Count().Equals(0)) { row.Cells[8].Value = ">" + extrusionHeight + "m。"; }
                            else { row.Cells[8].Value = crushElemInfo.roomHeight.ToString(); }
                        }
                        else { row.Cells[8].Value = crushElemInfo.roomHeight.ToString(); }
                        if (crushElemInfo.crushStairs.Count() > 0 || crushElemInfo.crushEscalators.Count() > 0) { row.Cells[8].Value += " ⭐"; } // 干涉到樓梯、電扶梯
                        try
                        {
                            double unboundedHeight = Convert.ToDouble(row.Cells[7].Value.ToString());
                            if (crushElemInfo.roomHeight < unboundedHeight)
                            {
                                row.Cells[8].Style.ForeColor = System.Drawing.Color.Red;
                                row.DefaultCellStyle.BackColor = System.Drawing.Color.LightYellow;
                            }
                        }
                        catch (Exception ex) { string error = ex.Message + "\n" + ex.ToString(); }
                    }
                    else { row.Cells[8].Value = ">" + extrusionHeight + "m。"; }
                    string roomName = row.Cells[1].Value.ToString();
                    if (roomName.Contains("日用水箱") || roomName.Contains("消防水箱")) { row.Cells[8].Value = ""; }
                    if (createSoildsFailed.Any(x => x.Equals(id))) // 干涉模型生成失敗
                    {
                        row.Cells[8].Value = "-";
                        row.Cells[8].Style.ForeColor = System.Drawing.Color.Red;
                    }
                }
                catch (Exception) { }
            }
            // 房間需求檢討
            dgv2 = RegulatoryReviewForm.dgv2;
            foreach (DataGridViewRow row in dgv2.Rows)
            {
                try
                {
                    string id = row.Cells[row.Cells.Count - 1].Value.ToString();
                    CrushElemInfo crushElemInfo = crushElemInfos.Where(x => x.id.Equals(id)).FirstOrDefault();
                    if (crushElemInfo != null)
                    {
                        if (crushElemInfo.roomHeight.Equals(0))
                        {
                            if (crushElemInfo.crushElems.Count().Equals(0)) { row.Cells[6].Value = ">" + extrusionHeight + "m。"; }
                            else { row.Cells[6].Value = crushElemInfo.roomHeight.ToString(); }
                        }
                        else { row.Cells[6].Value = crushElemInfo.roomHeight.ToString(); }
                        if (crushElemInfo.crushStairs.Count() > 0 || crushElemInfo.crushEscalators.Count() > 0) { row.Cells[6].Value += " ⭐"; } // 干涉到樓梯、電扶梯
                        try
                        {
                            double demandUnboundedHeight = Convert.ToDouble(row.Cells[5].Value.ToString());
                            if (crushElemInfo.roomHeight < demandUnboundedHeight)
                            {
                                row.Cells[6].Style.ForeColor = System.Drawing.Color.Red;
                                row.DefaultCellStyle.BackColor = System.Drawing.Color.LightYellow;
                            }
                        }
                        catch(Exception ex) { string error = ex.Message + "\n" + ex.ToString(); }
                    }
                    else { row.Cells[6].Value = ">" + extrusionHeight + "m。"; }
                    string roomName = row.Cells[1].Value.ToString();
                    if (roomName.Contains("日用水箱") || roomName.Contains("消防水箱")) { row.Cells[6].Value = ""; }
                    if (createSoildsFailed.Any(x => x.Equals(id))) // 干涉模型生成失敗
                    {
                        row.Cells[6].Value = "-";
                        row.Cells[6].Style.ForeColor = System.Drawing.Color.Red;
                    }
                }
                catch (Exception ex) { string error = ex.Message + "\n" + ex.ToString(); }
            }
            // 天花內淨高校核
            dgv6 = RegulatoryReviewForm.dgv6;
            foreach (DataGridViewRow row in dgv6.Rows)
            {
                try
                {
                    string id = row.Cells[row.Cells.Count - 1].Value.ToString();
                    CrushElemInfo crushElemInfo = crushElemInfos.Where(x => x.id.Equals(id)).FirstOrDefault();
                    if (crushElemInfo != null)
                    {  
                        // 天花內淨高
                        if (crushElemInfo.crushCeilings.Count() > 0 && crushElemInfo.crushFloors.Count > 0)
                        {
                            row.Cells[2].Value = "✔";
                            row.Cells[3].Value = crushElemInfo.clearSpace.ToString();
                        }
                        else if (crushElemInfo.crushCeilings.Count() > 0 && crushElemInfo.crushFloors.Count.Equals(0))
                        {
                            row.Cells[2].Value = "✔";
                            row.Cells[3].Value = ">" + crushElemInfo.clearSpace.ToString();
                        }
                        else { row.Cells[3].Value = ""; }
                        
                        if(crushElemInfo.crushElems.Count().Equals(0)) { row.Cells[4].Value = ">" + extrusionHeight + "m。"; }
                        else { row.Cells[4].Value = crushElemInfo.clearSpaceHeight.ToString(); } // 天花內淨高之淨高
                        // 有天花板但淨高小於1.5m; 無天花板但淨高小於4.5m, 紅字顯示
                        if (crushElemInfo.crushCeilings.Count() > 0 && crushElemInfo.clearSpace < 1.5)
                        {
                            row.Cells[3].Style.ForeColor = System.Drawing.Color.Red;
                            row.DefaultCellStyle.BackColor = System.Drawing.Color.LightYellow;
                        }
                        else if (crushElemInfo.crushCeilings.Count() == 0 && crushElemInfo.clearSpaceHeight < 4.5)
                        {
                            row.Cells[4].Style.ForeColor = System.Drawing.Color.Red;
                            row.DefaultCellStyle.BackColor = System.Drawing.Color.LightYellow;
                        }
                        if (crushElemInfo.crushStairs.Count() > 0 || crushElemInfo.crushEscalators.Count() > 0) { row.Cells[4].Value += " ⭐"; } // 干涉到樓梯、電扶梯
                    }
                    else { row.Cells[3].Value = ""; row.Cells[4].Value = ">" + extrusionHeight + "m。"; }
                    string roomName = row.Cells[0].Value.ToString();
                    if (roomName.Contains("日用水箱") || roomName.Contains("消防水箱")) { row.Cells[4].Value = ""; }
                    if (createSoildsFailed.Any(x => x.Equals(id))) // 干涉模型生成失敗
                    {
                        row.Cells[3].Value = "-";
                        row.Cells[3].Style.ForeColor = System.Drawing.Color.Red;
                        row.Cells[4].Value = "-";
                        row.Cells[4].Style.ForeColor = System.Drawing.Color.Red;
                    }
                }
                catch (Exception ex) { string error = ex.Message + "\n" + ex.ToString(); }
            }
        }
        /// <summary>
        /// 使用CreateExtrusionGeometry擠出房間底面Solid
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="topFace"></param>
        /// <param name="height"></param>
        /// <param name="materialId"></param>
        /// <param name="roomId"></param>
        /// <returns></returns>
        private Solid CreateCrushSolids(Document doc, Face topFace, double height, ElementId materialId, ElementId roomId)
        {
            //CurveLoop buttomCurves = new CurveLoop();
            //CurveLoop topCurves = new CurveLoop();
            //// 干涉元件底面, 找到最大邊長
            //double maxLength = 0.0;
            //foreach (EdgeArray edgeArray in topFace.EdgeLoops)
            //{
            //    List<Curve> curves = new List<Curve>();
            //    foreach (Edge edge in edgeArray)
            //    {
            //        Curve curve = edge.AsCurveFollowingFace(topFace);
            //        curves.Add(curve);
            //    }
            //    double curveLoopLength = curves.Select(x => x.Length).Sum();
            //    if (curveLoopLength >= maxLength)
            //    {
            //        buttomCurves = CurveLoop.Create(curves);
            //        maxLength = curveLoopLength;
            //    }
            //}
            //// 干涉元件頂面, 找到最大邊長
            //maxLength = 0.0;
            //foreach (EdgeArray edgeArray in topFace.EdgeLoops)
            //{
            //    List<Curve> curves = new List<Curve>();
            //    foreach (Edge edge in edgeArray)
            //    {
            //        Curve curve = edge.AsCurveFollowingFace(topFace);
            //        IList<XYZ> curveXYZs = new List<XYZ>();
            //        foreach (XYZ curveXYZ in curve.Tessellate())
            //        {
            //            XYZ xyz = new XYZ(curveXYZ.X, curveXYZ.Y, curveXYZ.Z + height);
            //            curveXYZs.Add(xyz);
            //        }
            //        curve = NurbSpline.CreateCurve(HermiteSpline.Create(curveXYZs, false));
            //        curves.Add(curve);
            //    }
            //    double curveLoopLength = curves.Select(x => x.Length).Sum();
            //    if (curveLoopLength >= maxLength)
            //    {
            //        topCurves = CurveLoop.Create(curves);
            //        maxLength = curveLoopLength;
            //    }
            //}

            //IList<CurveLoop> curveLoops = new List<CurveLoop>();
            //curveLoops.Add(buttomCurves);
            //curveLoops.Add(topCurves);

            List<CurveLoop> curveLoops = new List<CurveLoop>();
            Solid solid = null;
            // Area底面輪廓
            foreach (EdgeArray edgeArray in topFace.EdgeLoops)
            {
                CurveArray profile = new CurveArray();
                foreach (Edge edge in edgeArray)
                {
                    Curve curve = edge.AsCurveFollowingFace(topFace);
                    IList<XYZ> curveXYZs = new List<XYZ>();
                    foreach (XYZ curveXYZ in curve.Tessellate()) { curveXYZs.Add(new XYZ(curveXYZ.X, curveXYZ.Y, curveXYZ.Z)); }
                    curve = NurbSpline.CreateCurve(HermiteSpline.Create(curveXYZs, false));
                    profile.Append(curve);
                }
                // 創建 CurveLoop
                CurveLoop curveLoop = new CurveLoop();
                foreach (Curve curve in profile)
                {
                    curveLoop.Append(curve);
                }
                curveLoops.Add(curveLoop);
            }
            try
            {
                SolidOptions options = new SolidOptions(materialId, materialId);
                //solid = GeometryCreationUtilities.CreateLoftGeometry(curveLoops, options);
                solid = GeometryCreationUtilities.CreateExtrusionGeometry(curveLoops, XYZ.BasisZ, height, options);
                DirectShape ds = DirectShape.CreateElement(doc, materialId);
                ds.ApplicationId = "Application id";
                ds.ApplicationDataId = "Geometry object id";
                ds.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set(roomId + "：房間干涉元件");
                ds.SetShape(new List<GeometryObject>() { solid });
            }
            catch (Exception ex) { string error = ex.Message + "\n" + ex.ToString(); createSoildsFailed.Add(roomId.ToString()); }

            return solid;
        }
        /// <summary>
        /// 出衝突報告
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="logicalOrFilter"></param>
        /// <returns></returns>
        private List<CrushElemInfo> CrushReport(Document doc, LogicalOrFilter logicalOrFilter)
        {
            List<View3D> view3Ds = new FilteredElementCollector(doc).OfClass(typeof(View3D)).WhereElementIsNotElementType().Cast<View3D>().ToList();
            View3D view3D = view3Ds.Where(x => x.Name.Equals("{3D}")).FirstOrDefault();

            // 讀取專案中所有的樓板與天花板
            IList<ElementFilter> floorAndCeilingFilters = new List<ElementFilter>(); // 清空過濾器
            ElementCategoryFilter floorFilter = new ElementCategoryFilter(BuiltInCategory.OST_Floors); // 樓板
            ElementCategoryFilter ceilingFilter = new ElementCategoryFilter(BuiltInCategory.OST_Ceilings); // 天花板
            ElementCategoryFilter stairFilter = new ElementCategoryFilter(BuiltInCategory.OST_Stairs); // 樓梯
            floorAndCeilingFilters.Add(floorFilter);
            floorAndCeilingFilters.Add(ceilingFilter);
            floorAndCeilingFilters.Add(stairFilter);

            // 儲存使用中RevitLink的Document
            IList<RevitLinkInstance> revitLinkInss = new FilteredElementCollector(doc).OfClass(typeof(RevitLinkInstance)).WhereElementIsNotElementType().Cast<RevitLinkInstance>().Where(x => x.GetLinkDocument() != null).ToList();

            ICollection<ElementId> elemIds = new List<ElementId>();
            IList<Element> floorsAndCeilings = new FilteredElementCollector(doc).WherePasses(new LogicalOrFilter(floorAndCeilingFilters)).WhereElementIsNotElementType().ToList();
            foreach(Element elem in floorsAndCeilings) { elemIds.Add(elem.Id); }
            foreach (RevitLinkInstance revitLinkIns in revitLinkInss)
            {
                floorsAndCeilings = new FilteredElementCollector(revitLinkIns.GetLinkDocument()).WherePasses(new LogicalOrFilter(floorAndCeilingFilters)).WhereElementIsNotElementType().ToList();
                foreach (Element elem in floorsAndCeilings) { elemIds.Add(elem.Id); }
            }

            // 讀取專案中所有的電扶梯
            IList<Element> escalators = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_GenericModel).WhereElementIsNotElementType().ToList();
            escalators = escalators.Where(x => x.get_Parameter(BuiltInParameter.ELEM_FAMILY_PARAM).AsValueString().Contains("電扶梯"))
                                   .Where(x => !x.get_Parameter(BuiltInParameter.ELEM_FAMILY_PARAM).AsValueString().Contains("間縫")).ToList();
            foreach (Element elem in escalators) { elemIds.Add(elem.Id); }
            foreach (RevitLinkInstance revitLinkIns in revitLinkInss)
            {
                escalators = new FilteredElementCollector(revitLinkIns.GetLinkDocument()).OfCategory(BuiltInCategory.OST_GenericModel).WhereElementIsNotElementType().ToList();
                escalators = escalators.Where(x => x.get_Parameter(BuiltInParameter.ELEM_FAMILY_PARAM).AsValueString().Contains("電扶梯"))
                                       .Where(x => !x.get_Parameter(BuiltInParameter.ELEM_FAMILY_PARAM).AsValueString().Contains("間縫")).ToList();
                foreach (Element elem in escalators) { elemIds.Add(elem.Id); }
            }

            List<CrushElemInfo> crushElemInfos = new List<CrushElemInfo>();
            List<DirectShape> genericModels = new FilteredElementCollector(doc).WherePasses(logicalOrFilter).WhereElementIsNotElementType().Where(x => x is DirectShape)
                                              .Where(x => x.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).AsString() != null)
                                              .Where(x => x.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).AsString().Contains("房間干涉元件")).Cast<DirectShape>().ToList();

            List<string> crushElems = new List<string>();
            ICollection<ElementId> crushSolidIds = new List<ElementId>();
            foreach (DirectShape directShape in genericModels)
            {
                crushSolidIds.Add(directShape.Id);

                CrushElemInfo crushElemInfo = new CrushElemInfo();
                crushElemInfo.hostName = directShape.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).AsString(); // 房間編號
                crushElemInfo.id = crushElemInfo.hostName.Replace("：房間干涉元件", "");

                // 使用BoundingBoxIntersectsFilter、BoundingBoxIsInsideFilter快篩先檢查干涉物件
                Outline outline = new Outline(directShape.get_BoundingBox(view3D).Min, directShape.get_BoundingBox(view3D).Max);
                IList<ElementFilter> elementFilters = new List<ElementFilter>(); // 清空過濾器
                BoundingBoxIntersectsFilter boxIntersectsFilter = new BoundingBoxIntersectsFilter(outline);
                BoundingBoxIsInsideFilter boxIsInsideFilter = new BoundingBoxIsInsideFilter(outline);
                elementFilters.Add(boxIntersectsFilter);
                elementFilters.Add(boxIsInsideFilter);
                LogicalOrFilter bbFilters = new LogicalOrFilter(elementFilters);
                IList<Element> boxFilterElems = new FilteredElementCollector(doc, elemIds).WherePasses(bbFilters).WhereElementIsNotElementType().Excluding(crushSolidIds).ToList();

                try
                {
                    List<Solid> solidList = GetSolids(doc, directShape); // 干涉元件的Solid
                    foreach (Solid solid in solidList)
                    {
                        foreach (Element boxFilterElem in boxFilterElems)
                        {
                            try
                            {
                                IList<Element> elems = new FilteredElementCollector(doc, new List<ElementId>() { boxFilterElem.Id }).WherePasses(new ElementIntersectsSolidFilter(solid))
                                                       .WhereElementIsNotElementType().Excluding(crushSolidIds).Where(x => !String.IsNullOrEmpty(x.Name)).ToList();
                                if (elems.Count > 0)
                                {
                                    foreach (Element elem in elems)
                                    {
                                        if (elem.Category.Id.Value == (long)BuiltInCategory.OST_Stairs) // 樓梯
                                        {
                                            Stairs stairs = elem as Stairs;
                                            if(stairs != null) { crushElemInfo.crushStairs.Add(stairs); }                                            
                                        }
                                        else if (elem.Category.Id.Value == (long)BuiltInCategory.OST_GenericModel) // 一般模型
                                        {
                                            FamilyInstance escalator = elem as FamilyInstance;
                                            if (escalator != null) { crushElemInfo.crushEscalators.Add(escalator); }                                            
                                        }
                                        else
                                        {
                                            if(elem.Category.Id.Value == (long)BuiltInCategory.OST_Floors) // 樓板
                                            {
                                                if(elem is Floor)
                                                {
                                                    Floor floor = elem as Floor;
                                                    if (floor != null)
                                                    {
                                                        crushElemInfo.crushFloors.Add(floor);
                                                        crushElemInfo.crushElems.Add(floor);
                                                    }
                                                }
                                                else if(elem is FamilyInstance)
                                                {
                                                    FamilyInstance floor = elem as FamilyInstance;
                                                    if (floor != null)
                                                    {
                                                        crushElemInfo.crushElems.Add(floor);
                                                    }
                                                }
                                            }
                                            else if (elem.Category.Id.Value == (long)BuiltInCategory.OST_Ceilings) // 天花板
                                            {
                                                if(elem is Ceiling)
                                                {
                                                    Ceiling ceiling = elem as Ceiling;
                                                    if (ceiling != null)
                                                    {
                                                        crushElemInfo.crushCeilings.Add(ceiling);
                                                        crushElemInfo.crushElems.Add(ceiling);
                                                    }
                                                }
                                                else if (elem is FamilyInstance)
                                                {
                                                    FamilyInstance ceiling = elem as FamilyInstance;
                                                    if (ceiling != null)
                                                    {
                                                        crushElemInfo.crushElems.Add(ceiling);
                                                    }
                                                }
                                            }
                                        }
                                        crushElemInfo.crushElemName.Add(elem.Name + "_" + elem.Id); // 干涉到的天花板與樓板
                                    }
                                }
                            }
                            catch(Exception ex) { string error = ex.Message + "\n" + ex.ToString(); }
                        }
                        foreach (RevitLinkInstance revitLinkIns in revitLinkInss)
                        {
                            boxFilterElems = new FilteredElementCollector(revitLinkIns.GetLinkDocument(), elemIds).WherePasses(bbFilters).WhereElementIsNotElementType().Excluding(crushSolidIds).ToList();
                            foreach (Element boxFilterElem in boxFilterElems)
                            {
                                try
                                {
                                    IList<Element> elems = new FilteredElementCollector(revitLinkIns.GetLinkDocument(), new List<ElementId>() { boxFilterElem.Id }).WherePasses(new ElementIntersectsSolidFilter(solid))
                                                           .WhereElementIsNotElementType().Excluding(crushSolidIds).Where(x => !String.IsNullOrEmpty(x.Name)).ToList();
                                    if (elems.Count > 0)
                                    {
                                        foreach (Element elem in elems)
                                        {
                                            if (elem.Category.Id.Value == (long)BuiltInCategory.OST_Stairs) // 樓梯
                                            {
                                                Stairs stairs = elem as Stairs;
                                                if (stairs != null) { crushElemInfo.crushStairs.Add(stairs); }
                                            }
                                            else if (elem.Category.Id.Value == (long)BuiltInCategory.OST_GenericModel) // 一般模型
                                            {
                                                FamilyInstance escalator = elem as FamilyInstance;
                                                if (escalator != null) { crushElemInfo.crushEscalators.Add(escalator); }
                                            }
                                            else
                                            {
                                                if (elem.Category.Id.Value == (long)BuiltInCategory.OST_Floors) // 樓板
                                                {
                                                    if (elem is Floor)
                                                    {
                                                        Floor floor = elem as Floor;
                                                        if (floor != null)
                                                        {
                                                            crushElemInfo.crushFloors.Add(floor);
                                                            crushElemInfo.crushElems.Add(floor);
                                                        }
                                                    }
                                                    else if (elem is FamilyInstance)
                                                    {
                                                        FamilyInstance floor = elem as FamilyInstance;
                                                        if (floor != null)
                                                        {
                                                            crushElemInfo.crushElems.Add(floor);
                                                        }
                                                    }
                                                }
                                                else if (elem.Category.Id.Value == (long)BuiltInCategory.OST_Ceilings) // 天花板
                                                {
                                                    if (elem is Ceiling)
                                                    {
                                                        Ceiling ceiling = elem as Ceiling;
                                                        if (ceiling != null)
                                                        {
                                                            crushElemInfo.crushCeilings.Add(ceiling);
                                                            crushElemInfo.crushElems.Add(ceiling);
                                                        }
                                                    }
                                                    else if (elem is FamilyInstance)
                                                    {
                                                        FamilyInstance ceiling = elem as FamilyInstance;
                                                        if (ceiling != null)
                                                        {
                                                            crushElemInfo.crushElems.Add(ceiling);
                                                        }
                                                    }
                                                }
                                            }
                                            crushElemInfo.crushElemName.Add(elem.Name + "_" + elem.Id); // 干涉到的天花板與樓板
                                        }
                                    }
                                }
                                catch (Exception ex) { string error = ex.Message + "\n" + ex.ToString(); }
                            }
                        }
                    }
                }
                catch (Exception) { }

                try
                {
                    ElementId crushElemId = new ElementId(Convert.ToInt64(crushElemInfo.id));
                    Room room = doc.GetElement(crushElemId) as Room;

                    // 淨空空間
                    if (crushElemInfo.crushCeilings.Count > 0 && crushElemInfo.crushFloors.Count > 0)
                    {
                        double ceilingHeight = GetCeilingTopOrBottomFace(doc, crushElemInfo.crushCeilings, "top"); // 找到最低的天花板頂面                    
                        double floorHeight = GetFloorBottomFace(doc, crushElemInfo.crushFloors, ceilingHeight); // 找到高於天花板頂面的樓板底面
                        double clearSpace = Math.Round(UnitUtils.ConvertFromInternalUnits(floorHeight - ceilingHeight, UnitTypeId.Meters), 2, MidpointRounding.AwayFromZero);
                        if (clearSpace > 0) { crushElemInfo.clearSpace = Math.Round(clearSpace, 2, MidpointRounding.AwayFromZero); }
                    }
                    else if (crushElemInfo.crushCeilings.Count > 0 && crushElemInfo.crushFloors.Count.Equals(0))
                    {
                        double ceilingHeight = GetCeilingTopOrBottomFace(doc, crushElemInfo.crushCeilings, "top"); // 找到最低的天花板頂面
                        LocationPoint lp = room.Location as LocationPoint;
                        double clearSpace = extrusionHeight - Math.Round(UnitUtils.ConvertFromInternalUnits(ceilingHeight - lp.Point.Z, UnitTypeId.Meters), 2, MidpointRounding.AwayFromZero);
                        if (clearSpace > 0) { crushElemInfo.clearSpace = Math.Round(clearSpace, 2, MidpointRounding.AwayFromZero); }
                    }
                    // 淨高
                    if (crushElemInfo.crushElems.Count > 0)
                    {
                        double lowestFaceHeight = GetLowestFace(doc, crushElemInfo.crushElems, room, directShape);
                        crushElemInfo.roomHeight = lowestFaceHeight;
                        // 偵測天花板或樓板的底面
                        if (crushElemInfo.crushCeilings.Count > 0)
                        {
                            double ceilingHeight = GetCeilingTopOrBottomFace(doc, crushElemInfo.crushCeilings, "bottom"); // 找到最低的天花板底面
                            crushElemInfo.clearSpaceHeight = Math.Round(UnitUtils.ConvertFromInternalUnits(ceilingHeight - room.UpperLimit.Elevation, UnitTypeId.Meters), 2, MidpointRounding.AwayFromZero);
                        }
                        else
                        {
                            crushElemInfo.clearSpaceHeight = lowestFaceHeight;
                        }
                    }
                    if (crushElemInfo.crushElems.Count > 0 || crushElemInfo.crushStairs.Count > 0 || crushElemInfo.crushEscalators.Count > 0)
                    {
                        crushElemInfos.Add(crushElemInfo);
                    }
                }
                catch (Exception) { }
            }

            return crushElemInfos;
        }
        /// <summary>
        /// 找到最低的天花板頂面
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="ceilings"></param>
        /// <returns></returns>
        private double GetCeilingTopOrBottomFace(Document doc, List<Ceiling> ceilings, string topOrBottom)
        {
            double height = 0.0;
            foreach(Ceiling ceiling in ceilings)
            {
                List<Solid> floorSolids = GetSolids(doc, ceiling);
                List<Face> floorTopFaces = GetTopOrBottomFaces(floorSolids, topOrBottom);
                foreach(Face floorTopFace in floorTopFaces)
                {
                    if (floorTopFace is PlanarFace)
                    {
                        PlanarFace planarFace = floorTopFace as PlanarFace;
                        double ceilingHeight = planarFace.Origin.Z;
                        if (height == 0.0)
                        {
                            height = ceilingHeight;
                        }
                        else
                        {
                            if (ceilingHeight < height)
                            {
                                height = ceilingHeight;
                            }
                        }
                    }
                }
            }

            return height;
        }
        /// <summary>
        /// 找到高於天花板頂面的樓板底面
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="floors"></param>
        /// <param name="ceilingHeight"></param>
        /// <returns></returns>
        private double GetFloorBottomFace(Document doc, List<Floor> floors, double ceilingHeight)
        {
            double height = 0.0;
            foreach(Floor floor in floors)
            {
                List<Solid> floorSolids = GetSolids(doc, floor);
                List<Face> floorBottomFaces = GetTopOrBottomFaces(floorSolids, "bottom");
                foreach (Face floorBottomFace in floorBottomFaces)
                {
                    if (floorBottomFace is PlanarFace)
                    {
                        PlanarFace planarFace = floorBottomFace as PlanarFace;
                        double floorHeight = planarFace.Origin.Z;
                        if(floorHeight > ceilingHeight)
                        {
                            if (height == 0.0)
                            {
                                height = floorHeight;
                            }
                            else
                            {
                                if (floorHeight < height)
                                {
                                    height = floorHeight;
                                }
                            }
                        }
                    }
                }
            }
            return height;
        }
        /// <summary>
        /// 找到所有樓板、天花板最低的底面
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="elems"></param>
        /// <param name="room"></param>
        /// <param name="directShape"></param>
        /// <returns></returns>
        private double GetLowestFace(Document doc, List<Element> elems, Room room, DirectShape directShape)
        {
            LocationPoint lp = room.Location as LocationPoint;
            double roomHeight = lp.Point.Z;
            double height = 0.0;
            //using (Transaction trans = new Transaction(doc, "取得最低的面"))
            //{
            //    trans.Start();
                foreach (Element elem in elems)
                {
                    List<Solid> solids = GetSolids(doc, elem);
                    List<Face> bottomFaces = GetTopOrBottomFaces(solids, "bottom");
                    foreach (Face bottomFace in bottomFaces)
                    {
                        if (bottomFace is PlanarFace)
                        {
                            PlanarFace planarFace = bottomFace as PlanarFace;
                            double bottomFaceHeight = planarFace.Origin.Z;
                            // 找到干涉元件與面的最低交界點
                            try
                            {
                                List<Curve> dsCurves = new List<Curve>();
                                List<Solid> dsSolids = GetSolids(doc, directShape);
                                foreach (Solid dsSolid in dsSolids)
                                {
                                    foreach (Face dsFace in dsSolid.Faces)
                                    {
                                        foreach (EdgeArray dsEdgeArray in dsFace.EdgeLoops)
                                        {
                                            foreach (Edge dsEdge in dsEdgeArray)
                                            {
                                                Line dsLine = dsEdge.AsCurve() as Line;                                                
                                                try
                                                {
                                                    if (dsLine.Direction.Z != 0) { dsCurves.Add(dsEdge.AsCurveFollowingFace(dsFace)); } // 法線向上
                                                }
                                                catch (Exception ex) { string error = ex.Message + "\n" + ex.ToString(); }                                                
                                            }
                                        }
                                    }
                                }
                                List<XYZ> xyzs = FindFaceCurveIntersections(planarFace, dsCurves);
                                if(xyzs.Count > 0)
                                {
                                    bottomFaceHeight = xyzs.Min(x => x.Z); // 找到xyzs中最低的點

                                    double lowestFaceHeight = Math.Round(UnitUtils.ConvertFromInternalUnits(bottomFaceHeight - roomHeight, UnitTypeId.Meters), 2, MidpointRounding.AwayFromZero);
                                    if (elem.Category.Id.Value == (long)BuiltInCategory.OST_Floors)
                                    {
                                    //if(planarFace.Area >= room.Area) { } // 排除小於房間面積的"樓板"
                                    if (lowestFaceHeight > 0 && height == 0.0) { height = lowestFaceHeight; }
                                        else { if (lowestFaceHeight > 0 && lowestFaceHeight < height) { height = lowestFaceHeight; } }
                                    }
                                    else
                                    {
                                        if (lowestFaceHeight > 0 && height == 0.0) { height = lowestFaceHeight; }
                                        else { if (lowestFaceHeight > 0 && lowestFaceHeight < height) { height = lowestFaceHeight; }}
                                    }
                                }
                            }
                            catch (Exception ex) { string error = ex.Message + "\n" + ex.ToString(); }
                        }
                    }
                }
            //    trans.Commit();
            //}
            return height;
        }
        /// <summary>
        /// 取得頂面或底面
        /// </summary>
        /// <param name="solidList"></param>
        /// <param name="topOrBottom"></param>
        /// <returns></returns>
        private List<Face> GetTopOrBottomFaces(List<Solid> solidList, string topOrBottom)
        {
            List<Face> topOrBottomFaces = new List<Face>();
            foreach (Solid solid in solidList)
            {
                foreach (Face face in solid.Faces)
                {
                    if (topOrBottom.Equals("top")) // 頂面
                    {
                        double faceTZ = face.ComputeNormal(new UV(0.5, 0.5)).Z;
                        if (faceTZ > 0.0) { topOrBottomFaces.Add(face); }
                    }
                    else if(topOrBottom.Equals("bottom")) // 底面
                    {
                        //if (face.ComputeNormal(new UV(0.5, 0.5)).IsAlmostEqualTo(XYZ.BasisZ.Negate())) { topOrBottomFaces.Add(face); }
                        double faceTZ = face.ComputeNormal(new UV(0.5, 0.5)).Z;
                        if (faceTZ < 0.0) { topOrBottomFaces.Add(face); }
                    }
                }
            }
            return topOrBottomFaces;
        }
        /// <summary>
        /// 儲存所有房間的Solid
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="elem"></param>
        /// <returns></returns>
        private List<Solid> GetSolids(Document doc, Element elem)
        {
            List<Solid> solidList = new List<Solid>();

            // 1.讀取Geometry Option
            Options options = new Options();
            //options.View = doc.GetElement(room.Level.FindAssociatedPlanViewId()) as Autodesk.Revit.DB.View;
            options.DetailLevel = ((doc.ActiveView != null) ? doc.ActiveView.DetailLevel : ViewDetailLevel.Medium);
            options.ComputeReferences = true;
            options.IncludeNonVisibleObjects = true;
            // 得到幾何元素
            GeometryElement geomElem = elem.get_Geometry(options);
            List<Solid> solids = GeometrySolids(geomElem);
            foreach (Solid solid in solids)
            {
                solidList.Add(solid);
            }

            return solidList;
        }
        /// <summary>
        /// 取得房間的Solid
        /// </summary>
        /// <param name="geoObj"></param>
        /// <returns></returns>
        private List<Solid> GeometrySolids(GeometryObject geoObj)
        {
            List<Solid> solids = new List<Solid>();
            if (geoObj is Solid)
            {
                Solid solid = (Solid)geoObj;
                if (solid.Faces.Size > 0)
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
        /// <summary>
        /// 找到干涉元件與面的最低交界點
        /// </summary>
        /// <param name="face"></param>
        /// <param name="curves"></param>
        /// <returns></returns>
        public List<XYZ> FindFaceCurveIntersections(Face face, List<Curve> curves)
        {            
            List<XYZ> intersectionPoints = new List<XYZ>(); // 儲存所有交點

            foreach (Curve curve in curves)
            {
                // 創建IntersectionResultArray來儲存交點結果
                IntersectionResultArray intersectionResults;
                // 使用Face的Intersect來查找Curve與Face的交點
                SetComparisonResult result = face.Intersect(curve, out intersectionResults);
                // 如果有交點
                if (result == SetComparisonResult.Overlap && intersectionResults != null)
                {
                    // 儲存所有交點
                    foreach (IntersectionResult intersection in intersectionResults)
                    {
                        intersectionPoints.Add(intersection.XYZPoint);
                    }
                }
            }
            
            return intersectionPoints; // 返回交點座標
        }
        /// <summary>
        /// 移除干涉元件
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="logicalOrFilter"></param>        
        private void RemoveElems(Document doc, LogicalOrFilter logicalOrFilter)
        {
            List<ElementId> crushFloorIds = new List<ElementId>();
            try
            {
                crushFloorIds = new FilteredElementCollector(doc).WherePasses(logicalOrFilter).WhereElementIsNotElementType().Where(x => x is DirectShape).Cast<DirectShape>()
                                .Where(x => x.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).AsString() != null)
                                .Where(x => x.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).AsString().Contains("房間干涉元件")).Select(x => x.Id).ToList();
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Error", "房間參數錯誤！\n" + ex.Message + "\n" + ex.ToString());
            }
            using (Transaction trans = new Transaction(doc, "移除干涉元件"))
            {
                trans.Start();
                doc.Delete(crushFloorIds);
                trans.Commit();
            }
        }
        /// <summary>
        /// 在3D視圖建立模型線
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="planarFace"></param>
        private void DrawModelCurve(Document doc, PlanarFace planarFace)
        {
            foreach (EdgeArray edgeArray in planarFace.EdgeLoops)
            {
                foreach (Edge edge in edgeArray)
                {
                    Curve curve = edge.AsCurveFollowingFace(planarFace);
                    IList<XYZ> curveXYZs = new List<XYZ>();
                    foreach (XYZ curveXYZ in curve.Tessellate()) { curveXYZs.Add(curveXYZ); }
                    curve = NurbSpline.CreateCurve(HermiteSpline.Create(curveXYZs, false));
                    try
                    {
                        Line line = Line.CreateBound(curve.Tessellate()[0], curve.Tessellate()[curve.Tessellate().Count - 1]);
                        XYZ normal = new XYZ(line.Direction.Z - line.Direction.Y, line.Direction.X - line.Direction.Z, line.Direction.Y - line.Direction.X); // 使用與線不平行的任意向量
                        Plane plane = Plane.CreateByNormalAndOrigin(normal, curve.Tessellate()[0]);
                        SketchPlane sketchPlane = SketchPlane.Create(doc, plane);
                        ModelCurve modelCurve = doc.Create.NewModelCurve(line, sketchPlane);
                    }
                    catch (Exception ex) { string error = ex.Message + "\n" + ex.ToString(); }
                }
            }
        }
        public string GetName()
        {
            return "Event handler is working now!!";
        }
    }
}