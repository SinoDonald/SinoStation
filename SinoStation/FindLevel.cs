using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;
using static SinoStation.RegulatoryReview;

namespace SinoStation
{
    public class LevelElevation
    {
        public Level level { get; set; }
        public string name { get; set; }
        public double elevation { get; set; }
        public double height { get; set; }
        public int sort { get; set; }
    }
    public class FindLevel
    {
        // 找到當前視圖的Level相關資訊
        public Tuple<List<LevelElevation>, LevelElevation, double> FindDocViewLevel(Document doc)
        {
            // 查詢所有Level的高程並排序
            List<Level> levels = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Levels).WhereElementIsNotElementType().Cast<Level>().ToList();
            List<LevelElevation> levelElevList = new List<LevelElevation>();
            foreach (Level level in levels)
            {
                LevelElevation levelElevation = new LevelElevation();
                levelElevation.name = level.Name;
                levelElevation.level = level;
                levelElevation.height = level.get_Parameter(BuiltInParameter.LEVEL_ELEV).AsDouble();
                levelElevation.elevation = Convert.ToDouble(level.get_Parameter(BuiltInParameter.LEVEL_ELEV).AsValueString());
                levelElevList.Add(levelElevation);
            }
            levelElevList = (from x in levelElevList
                             select x).OrderBy(x => x.elevation).ToList();
            double startElev = 0.0;
            double endElev = 0.0;
            double floorHeight = 10;
            // 找到當前樓層
            LevelElevation viewLevel = levelElevList[0];
            try
            {
                viewLevel = (from x in levelElevList
                             where x.level.Id.Equals(doc.ActiveView.GenLevel.Id)
                             select x).FirstOrDefault();
                int leCount = levelElevList.IndexOf(viewLevel);
                // 查詢當前樓層與上一樓層的高度, 製作火源高度
                if (levelElevList.Count >= 2)
                {
                    if (leCount < levelElevList.Count)
                    {
                        startElev = levelElevList[leCount].elevation;
                        if ((leCount + 1) < levelElevList.Count)
                        {
                            endElev = levelElevList[leCount + 1].elevation;
                            floorHeight = endElev - startElev;
                        }
                        // 如果當前視圖為最高樓層時
                        else
                        {
                            endElev = levelElevList[leCount].elevation;
                            floorHeight = 10;
                        }
                    }
                    else
                    {
                        startElev = levelElevList[leCount].elevation;
                        if ((leCount + 1) < levelElevList.Count)
                        {
                            endElev = levelElevList[leCount + 1].elevation;
                            floorHeight = endElev - startElev;
                        }
                        // 如果當前視圖為最高樓層時
                        else
                        {
                            endElev = levelElevList[leCount].elevation;
                            floorHeight = 10;
                        }
                    }
                }
            }
            catch (NullReferenceException)
            {

            }

            Tuple<List<LevelElevation>, LevelElevation, double> multiValue = Tuple.Create(levelElevList, viewLevel, floorHeight);

            return multiValue;
        }
        // 計算每層Level的高程差
        public List<LevelElevation> LevelElevationCalcul(List<LevelElevation> levelElevList)
        {
            List<LevelElevation> newlevelElevList = new List<LevelElevation>();
            for (int i = 0; i < levelElevList.Count(); i++)
            {
                LevelElevation newlevelElev = new LevelElevation();
                double startElev = 0.0;
                double endElev = 0.0;
                double floorHeight = 9999;
                // 找到當前樓層
                LevelElevation viewLevel = (from x in levelElevList
                                            where x.level.Id.Equals(levelElevList[i].level.Id)
                                            select x).FirstOrDefault();
                int leCount = levelElevList.IndexOf(viewLevel);
                newlevelElev.level = viewLevel.level; // Level
                newlevelElev.name = viewLevel.name; // 名稱
                newlevelElev.elevation = viewLevel.elevation; // 高程
                newlevelElev.height = floorHeight / 1000; // 與上一樓層高程差
                if (i < levelElevList.Count() - 1)
                {
                    // 查詢當前樓層與上一樓層的高度
                    if (levelElevList.Count >= 2)
                    {
                        if (leCount < levelElevList.Count)
                        {
                            startElev = levelElevList[leCount].elevation;
                            endElev = levelElevList[leCount + 1].elevation;
                            floorHeight = endElev - startElev;
                            newlevelElev.height = floorHeight / 1000; // 與上一樓層高程差
                        }
                        else
                        {
                            startElev = levelElevList[leCount].elevation;
                            endElev = levelElevList[leCount - 1].elevation;
                            floorHeight = startElev - endElev;
                            newlevelElev.height = floorHeight / 1000; // 與上一樓層高程差
                        }
                    }
                }
                newlevelElevList.Add(newlevelElev);
            }

            return newlevelElevList;
        }
    }
}
