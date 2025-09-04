using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using Arc = Autodesk.Revit.DB.Arc;
using Excel = Microsoft.Office.Interop.Excel;
using Floor = Autodesk.Revit.DB.Floor;
using Line = Autodesk.Revit.DB.Line;
using View = Autodesk.Revit.DB.View;

namespace SinoStation_2020
{
    public partial class RegulatoryReviewForm : System.Windows.Forms.Form
    {
        UIApplication revitUIApp = null;
        UIDocument revitUIDoc = null;
        Document revitDoc = null;
        private RevitDocument m_connect = null;
        public static bool login_result = false;
        public static HttpClient client = new HttpClient();
        private string m_filePath = string.Empty;
        public static DataGridView dgv1 = null;
        public static DataGridView dgv2 = null;
        public static DataGridView dgv6 = null;
        ExternalEvent m_externalEvent_EditRoomOffset;
        ExternalEvent m_externalEvent_EditRoomPara;
        ExternalEvent m_externalEvent_CreateCrush;

        public class CodeSort
        {
            public string code { get; set; }
            public int number { get; set; }
            public int sort { get; set; }
        }
        public static List<CodeSort> orderByCode = new List<CodeSort>();
        public static List<LevelElevation> levelElevList = new List<LevelElevation>();
        public class AllBottomFaces
        {
            public Element upElem = null; // 此房間上方的樓板、天花板
            public List<PlanarFace> bottomFaces = new List<PlanarFace>(); // 各樓層樓板的底面
            public double bottomElevation = 0.0; // 底部高程
        }
        public class FloorBottomFaces
        {
            public Floor upFloor = null; // 此房間上方的樓板
            public Ceiling upCeiling = null; // 此房間上方的天花板
            public FamilyInstance upFamilyInstance = null; // 此房間上方的FamilyInstance樓板、天花板
            public double bottomElevation = 0.0; // 底部高程
            public List<PlanarFace> bottomFaces = new List<PlanarFace>(); // 各樓層樓板的底面
        }
        public class LevelFloorBottomFaces
        {
            public string name = string.Empty; // 名稱
            public Level level = null; // 樓層
            public List<FloorBottomFaces> floorBottomFaces = new List<FloorBottomFaces>();
            public List<Element> upFloorCeilings = new List<Element>(); // 此房間上方的所有樓板、天花板
            public List<PlanarFace> bottomFace = new List<PlanarFace>(); // 各樓層樓板的底面
            public double height = 0.0; // 淨高
        }
        public class ExcelInfo
        {
            public string[] titles = new string[] { }; // 標題
            public string[] items = new string[] { }; // 項目
        }
        public class TitalNames
        {
            public int code = 0; // 代碼
            public int classification = 0; // 區域
            public int level = 0; // 樓層
            public int name = 0; // 空間名稱(中文)
            public int engName = 0; // 空間名稱(英文)
            public int otherName = 0; // 其他名稱
            public int category = 0; // 設備/系統
            public int count = 0; // 數量
            public int demandArea = 0; // 需求面積
            public int maxArea = 0; // 最大面積(m2)
            public int minArea = 0; // 最小面積(m2)
            public int permit = 0; // 容許差異(±%)
            public int specificationMinWidth = 0; // 最小規範寬度
            public int demandMinWidth = 0; // 最小需求寬度
            public int unboundedHeight = 0; // 規範淨高
            public int demandUnboundedHeight = 0; // 需求淨高
            public int door = 0; // 門(mm)
        }
        public class ExcelCompare
        {
            public string code = string.Empty; // 代碼
            public string classification = string.Empty; // 區域
            public string level = string.Empty; // 樓層
            public string name = string.Empty; // 空間名稱(中文)
            public string engName = string.Empty; // 空間名稱(英文)
            public List<string> otherNames = new List<string>(); // 其他名稱
            public string system = string.Empty; // 設備/系統
            public string category = string.Empty; // 類別
            public int count = 0; // 數量
            public double demandArea = 0.0; // 需求面積
            public double maxArea = 0.0; // 最大面積(m2)
            public double minArea = 0.0; // 最小面積(m2)
            public double permit = 0.0; // 容許差異(±%)
            public double specificationMinWidth = 0.0; // 最小規範寬度
            public double demandMinWidth = 0.0; // 最小需求寬度
            public double unboundedHeight = 0.0; // 規範淨高
            public double demandUnboundedHeight = 0.0; // 需求淨高
            public double doorWidth = 0.0; // 門寬(mm)
            public double doorHeight = 0.0; // 門高(mm)
        }
        public class ElementInfo
        {
            public string code = string.Empty; // 代碼
            public string sort1 = string.Empty; // 代碼[0]
            public int sort2 = 0; // 代碼[1]
            public string id = string.Empty; // id
            public Element elem = null; // 元件
            public string title = string.Empty; // 標題
            public string name = string.Empty; // 名稱
            public string changeName = string.Empty; // 名稱
            public string engName = string.Empty; // 空間名稱(英文)
            public string type = string.Empty; // 類型
            public string level = string.Empty; // 樓層
            public string levelHeight = string.Empty; // 樓層高度
            public View view = null; // 視圖
            public string material = string.Empty; // 材料
            public double concreteWidth = 0.0; // 混凝土寬度
            public double cost = 0.0; // 成本
            public string unit = string.Empty; // 單位
            public double perimeter = 0.0; // 周長
            public double area = 0.0; // 面積
            public double volume = 0.0; // 體積
            public double topElevation = 0.0; // 頂部高程
            public double count = 0.0; // 數量
            public bool compare = false; // 數量是否與上次一樣
            public List<Solid> solids = new List<Solid>(); // Room的Solid
            public List<Face> bottomFaces = new List<Face>(); // Room的底面
            public List<Element> intersectElems = new List<Element>(); // Solid干涉到的元件
            public double lowerOffset = 0.0; // 基準偏移
            public List<double> boundarySegments = new List<double>(); // 所有邊界長度
            public List<XYZ> boundaryPoints = new List<XYZ>(); // 邊界與Room的座標
            public double maxBoundaryLength = 0.0; // 最長邊界
            public double minBoundaryLength = 0.0; // 最短邊界
            public double roomHeight = 0.0; // 未設邊界的高度
            public string sn = string.Empty; // 編號
        }
        public class ExportExcel
        {
            public string sheetName { get; set; }
            public List<string> titles = new List<string>();
            public List<List<string>> excelDatas = new List<List<string>>();
        }
        public bool trueOrFalse = false; // 是否選擇房間校核Excel檔
        public string prjName = string.Empty; // 工程名稱
        public string prjOwner = string.Empty; // 業主名稱
        public static string filePath = string.Empty; // Excel路徑位置
        bool externalLogin = false; // 使用外部登入驗證
        public class DoorData
        {
            public ElementId id = null;
            public FamilyInstance door = null;
            public Level level = null;
            public Room fromRoom = null;
            public Room toRoom = null;
            public string belong = string.Empty;
            public string roomName = string.Empty; // Room參數
            public double height = 0.0; // 高度
            public double width = 0.0; // 寬度
        }
        List<Document> docList = new List<Document>(); // 專案中所有的Document
        List<ExcelCompare> excelCompareList = new List<ExcelCompare>();
        public static List<ElementInfo> elementInfoList = new List<ElementInfo>();
        public List<DoorData> doorDatas = new List<DoorData>();
        public static List<DoorData> doorDataList = new List<DoorData>();
        long chooseRoomId = 0;
        int dgv1RowIndex = 0;
        int dgv2RowIndex = 0;
        int dgv6RowIndex = 0;
        public SortOrder sortOrder = SortOrder.Ascending; // 預設升冪排序

        public RegulatoryReviewForm(UIApplication uiapp, RevitDocument connect, ExternalEvent externalEvent_EditRoomOffset, ExternalEvent externalEvent_EditRoomPara, ExternalEvent externalEvent_CreateCrush)
        {
            revitUIApp = uiapp;
            revitUIDoc = connect.RevitDoc;
            revitDoc = connect.RevitDoc.Document;
            m_connect = connect;
            m_filePath = filePath;
            m_externalEvent_EditRoomOffset = externalEvent_EditRoomOffset; //建立外部事件_編輯Room偏移
            m_externalEvent_EditRoomPara = externalEvent_EditRoomPara; //建立外部事件_編輯門Room參數
            m_externalEvent_CreateCrush = externalEvent_CreateCrush; //建立外部事件_建立干涉模型
            InitializeComponent();
            // 隱藏
            label6.Hide();
            label7.Hide();
            updateDoorBtn.Hide();
            CenterToParent(); // 置中

            docList.Add(revitDoc); // 原專案Document
            List<RevitLinkInstance> rvtInss = new FilteredElementCollector(revitDoc).OfClass(typeof(RevitLinkInstance)).WhereElementIsNotElementType().Cast<RevitLinkInstance>().ToList();
            foreach (RevitLinkInstance rvtIns in rvtInss)
            {
                // 連結載入的模型在加入
                if (rvtIns.IsMonitoringLinkElement() == true)
                {
                    docList.Add(rvtIns.GetLinkDocument());
                }
            }
            // 先檢視是否有設定好要移除的特殊符號
            List<string> charsToRemove = CreateCharsToRemoveTXT();
            // 選取Excel比對數量差異
            Tuple<bool, string> tuple = ChooseFiles();
            this.trueOrFalse = tuple.Item1;
            filePath = tuple.Item2;

            if (trueOrFalse == true)
            {
                excelCompareList = ReadExcel(filePath, charsToRemove);
                RecordCodeSort(excelCompareList); // 依Code排序
                FindLevel findLevel = new FindLevel();
                Tuple<List<LevelElevation>, LevelElevation, double> multiValue = findLevel.FindDocViewLevel(revitDoc);
                levelElevList = multiValue.Item1.OrderBy(x => x.elevation).ToList(); // 全部樓層
                int i = 0;
                foreach(LevelElevation levelElev in levelElevList)
                {
                    levelElev.sort = i;
                    i++;
                }
                UpdateModel(); // 更新模型資訊
            }
        }

        private void RegulatoryReviewForm_Load(object sender, EventArgs e)
        {
            if (LoginForm.external == false)
            {
                client = client_login();
                if (client == null)
                {
                    this.Text = "SinoStation - 認證失敗，5秒後強制關閉";
                    this.Enabled = false;
                    login_result = false;
                    timer1.Start();
                }
                var result = client.GetAsync($"/user/me").Result;
                if (result.IsSuccessStatusCode)
                {
                    login_result = true;
                    string s = result.Content.ReadAsStringAsync().Result;
                    this.Text = "SinoStation - " + DecodeEncodedNonAsciiCharacters(s.Substring(8, s.Length - 10));
                    revitUIApp.Idling += new EventHandler<IdlingEventArgs>(IdleUpdate);
                }
                else
                {
                    this.Text = "SinoStation - 認證失敗，5秒後強制關閉";
                    this.Enabled = false;
                    login_result = false;
                    timer1.Start();
                }
            }
            else
            {
                this.Text = "SinoStation - " + LoginForm.external_username;
                externalLogin = true;
                revitUIApp.Idling += new EventHandler<IdlingEventArgs>(IdleUpdate);
                //m_externalEvent_ChooseRoomShow.Raise();
            }
        }
        public static HttpClient client_login()
        {
            try
            {
                client = new HttpClient();
                //client.BaseAddress = new Uri("http://127.0.0.1:8000/");
                client.BaseAddress = new Uri("https://bimdata.sinotech.com.tw/");
                //client.BaseAddress = new Uri("http://bimdata.secltd/");
                client.DefaultRequestHeaders.Accept.Clear();
                var headerValue = new System.Net.Http.Headers.MediaTypeWithQualityHeaderValue("application/json");
                client.DefaultRequestHeaders.Accept.Add(headerValue);
                client.DefaultRequestHeaders.ConnectionClose = true;
                Task.WaitAll(client.GetAsync($"/login/?USERNAME={Environment.UserName}&REVITAPI=SinoStation-Passway"));
                //Task.WaitAll(client.GetAsync($"/login/?USERNAME=11111&REVITAPI=SinoPipe"));
                return client;
            }
            catch (Exception)
            {
                login_result = false;
                return null;
            }
        }
        int time_s = 5;
        private void timer1_Tick(object sender, EventArgs e)
        {
            time_s = time_s - 1;
            if (time_s == 0)
            {
                timer1.Stop();
                this.Dispose();
                LoginForm login = new LoginForm(revitUIApp, m_connect, m_externalEvent_EditRoomOffset, m_externalEvent_EditRoomPara, m_externalEvent_CreateCrush);
                login.Show();
            }
            this.Text = "SinoStation - 認證失敗，" + time_s.ToString() + "秒後強制關閉";
        }
        static string DecodeEncodedNonAsciiCharacters(string value)
        {
            return Regex.Replace(
                value,
                @"\\u(?<Value>[a-zA-Z0-9]{4})",
                m => {
                    return ((char)int.Parse(m.Groups["Value"].Value, NumberStyles.HexNumber)).ToString();
                });
        }
        /// <summary>
        /// 匯出Excel
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>

        private void exportBtn_Click(object sender, EventArgs e)
        {
            List<ExportExcel> exportExcels = new List<ExportExcel>();
            // 取得所有DataGridView的Cell資料
            foreach (TabPage tabPage in tabControl1.TabPages)
            {
                foreach (var tabControl in tabPage.Controls)
                {
                    ExportExcel exportExcel = new ExportExcel();
                    DataGridView dgv = tabControl as DataGridView;
                    if (dgv != null)
                    {
                        try
                        {
                            exportExcel.sheetName = dgv.Parent.Text;
                            foreach (DataGridViewColumn dgvCol in dgv.Columns)
                            {
                                exportExcel.titles.Add(dgvCol.HeaderText);
                            }
                            foreach (DataGridViewRow dgvRow in dgv.Rows)
                            {
                                List<string> excelData = new List<string>();
                                foreach (DataGridViewCell cell in dgvRow.Cells)
                                {
                                    excelData.Add(cell.Value.ToString());
                                }
                                exportExcel.excelDatas.Add(excelData);
                            }
                            exportExcels.Add(exportExcel);

                            if (exportExcel.sheetName.Equals("天花內淨高校核"))
                            {
                                exportExcel = new ExportExcel();
                                exportExcel.sheetName = "房間資訊";
                                exportExcel.titles = new List<string>() { "樓層", "樓層高度(mm)", "房間編號", "房間名稱", "房間淨高(mm)", "房間面積(㎡)", "房間寬度(mm)", "房間長度(mm)", "天花板高度(mm)", "備註", "ID" };
                                foreach (DataGridViewRow dgvRow in dgv.Rows)
                                {
                                    try
                                    {
                                        ElementInfo elementInfo = elementInfoList.Where(x => x.id.Equals(dgvRow.Cells[5].Value.ToString())).FirstOrDefault();

                                        List<string> excelData = new List<string>();
                                        excelData.Add(dgvRow.Cells[1].Value.ToString()); // 樓層
                                        excelData.Add(elementInfo.levelHeight); // 樓層高度(mm)
                                        excelData.Add(elementInfo.sn); // 房間編號
                                        excelData.Add(dgvRow.Cells[0].Value.ToString()); // 房間名稱
                                        excelData.Add(dgvRow.Cells[4].Value.ToString()); // 房間淨高(mm)
                                        excelData.Add(elementInfo.area.ToString()); // 房間面積(㎡)
                                        string width = string.Empty;
                                        int count = 4;
                                        if(elementInfo.boundarySegments.Count < 4) { count = elementInfo.boundarySegments.Count; }
                                        for (int i = 0; i < count; i++)
                                        {
                                            width += elementInfo.boundarySegments[i];
                                            if(i != count - 1) { width += "、"; }
                                            else { if(elementInfo.boundarySegments.Count > 4) { width += "..."; } }
                                        }
                                        excelData.Add(width); // 房間寬度(mm)
                                        excelData.Add(""); // 房間長度(mm)
                                        excelData.Add(dgvRow.Cells[3].Value.ToString()); // 天花板高度(mm)
                                        excelData.Add(""); // 備註
                                        excelData.Add(dgvRow.Cells[5].Value.ToString()); // ID
                                        exportExcel.excelDatas.Add(excelData);
                                    }
                                    catch(Exception ex) { string error = ex.Message + "\n" + ex.ToString(); }
                                }
                                exportExcels.Add(exportExcel);
                            }
                        }
                        catch(Exception ex) { string error = ex.Message + "\n" + ex.ToString(); }
                    }
                }
            }
            try { ExportToExcel(exportExcels); }
            catch (Exception ex) { string error = ex.Message + "\n" + ex.ToString(); }
        }
        /// <summary>
        /// 匯出Excel
        /// </summary>
        /// <param name="exportExcels"></param>
        private void ExportToExcel(List<ExportExcel> exportExcels)
        {
            //// 刪除EXCEL程序, 否則每執行一次會持續增加
            //System.Diagnostics.Process[] procs = System.Diagnostics.Process.GetProcessesByName("EXCEL");
            //foreach (System.Diagnostics.Process p in procs) { p.Kill(); }

            Excel.Application excelApp = new Excel.Application(); // 創建Excel
            //excelApp.Visible = true; // 開啟Excel可見
            Workbook workbook = excelApp.Workbooks.Add(); // 創建一個空的workbook
            Sheets sheets = workbook.Sheets; // 獲取當前工作簿的數量
            int sheetCount = 1;

            // 儲存路徑
            string path = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\法規檢核.xlsx";
            FolderBrowserDialog dialog = new FolderBrowserDialog();
            dialog.Description = "請選擇檔案路徑";
            bool sureOrNot = false;
            try
            {
                if (dialog.ShowDialog() == DialogResult.OK) { path = dialog.SelectedPath + "\\"; sureOrNot = true; }
                if (path.EndsWith("\\")) { workbook.SaveCopyAs(path += "法規檢核" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".xlsx"); }
                else { workbook.SaveCopyAs(path + "\\法規檢核" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".xlsx"); }
            }
            catch (Exception) { path = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\法規檢核.xlsx"; }

            if (sureOrNot)
            {
                foreach (ExportExcel exportExcel in exportExcels)
                {
                    string fileName = exportExcel.sheetName;
                    fileName = string.Concat(fileName.Split(Path.GetInvalidFileNameChars())); // Excel不允許某些字符出現在工作表名稱中，例如：:、/、*、[ ] 等。
                    List<string> existingNames = workbook.Worksheets.Cast<Worksheet>().Select(x => x.Name).ToList();
                    Worksheet worksheet = sheets[1];
                    try
                    {
                        if (sheetCount == 1) { if (!existingNames.Contains(fileName)) { worksheet.Name = fileName; } }
                        else
                        {
                            worksheet = sheets.Add(After: sheets[sheets.Count]); // 新增一個工作表
                            try
                            {
                                if (!existingNames.Contains(fileName)) { worksheet.Name = fileName; }
                            }
                            catch (Exception ex) { string error = ex.Message + "\n" + ex.ToString(); }
                        }
                        sheetCount++;

                        worksheet.Cells.Font.Name = "微軟正黑體"; // 設定Excel資料字體字型
                        worksheet.Cells.Font.Size = 10; // 設定Excel資料字體大小
                        worksheet.Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter; // 文字水平置中
                        worksheet.Cells.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter; // 文字垂直置中
                        for (int col = 0; col < exportExcel.titles.Count; col++)
                        {
                            excelApp.Cells[1, col + 1] = exportExcel.titles[col];
                            excelApp.Cells[1, col + 1].Borders.LineStyle = Excel.XlLineStyle.xlContinuous; // 設定框線
                            excelApp.Cells[1, col + 1].Interior.Color = System.Drawing.Color.LightYellow; // 設定樣式與背景色
                        }
                        for (int i = 0; i < exportExcel.excelDatas.Count; i++)
                        {
                            for (int j = 0; j < exportExcel.excelDatas[i].Count; j++)
                            {
                                excelApp.Cells[i + 2, j + 1] = exportExcel.excelDatas[i][j].ToString();
                                excelApp.Cells[i + 2, j + 1].Borders.LineStyle = Excel.XlLineStyle.xlContinuous; // 設定框線
                            }
                        }
                    }
                    catch (Exception ex) { string error = ex.Message + "\n" + ex.ToString(); }
                    ReleaseObject(worksheet);
                }

                // 如果檔案已存在，先刪除
                if (File.Exists(path)) { File.Delete(path); }
                workbook.SaveAs(path);

                // 關閉工作簿和ExcelApp
                workbook.Close();
                excelApp.Quit();

                // 釋放COM
                ReleaseObject(sheets);
                ReleaseObject(workbook);
                ReleaseObject(excelApp);

                this.trueOrFalse = true;
                TaskDialog.Show("Revit", "完成");
            }
            //Close();
        }
        /// <summary>
        /// 釋放COM
        /// </summary>
        /// <param name="obj"></param>
        static void ReleaseObject(object obj)
        {
            try { System.Runtime.InteropServices.Marshal.ReleaseComObject(obj); obj = null; }
            catch (Exception ex) { string error = ex.Message + "\n" + ex.ToString(); obj = null; }
            finally { GC.Collect(); }
        }
        /// <summary>
        /// 更新
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void updateBtn_Click(object sender, EventArgs e)
        {            
            UpdateModel(); // 更新模型資訊
            m_externalEvent_CreateCrush.Raise(); // 干涉檢查淨高與空間淨高
        }
        private void UpdateModel()
        {
            List<string> charsToRemove = CreateCharsToRemoveTXT(); // 先檢視是否有設定好要移除的特殊符號
            this.label1.Text = "";
            foreach (string sign in charsToRemove)
            {
                this.label1.Text += sign + "　";
            }
            this.label1.ForeColor = System.Drawing.Color.Red; // 特殊符號變色
            // 讀取專案與連結模型的Room, 並儲存元件資料
            elementInfoList = new List<ElementInfo>();
            foreach(Document doc in docList)
            {
                foreach(ElementInfo elemInfo in ModelInfo(doc, charsToRemove, excelCompareList))
                {
                    elementInfoList.Add(elemInfo);
                }                
            }
            List<ElementInfo> sortElemInfoList = elementInfoList.OrderBy(x => OrderByCode(x.code)).ThenBy(x => x.name).ThenBy(x => x.level).ToList();            
            CreateRoomReview(excelCompareList, dataGridView1, sortElemInfoList); // 建立「房間規範檢討」欄位資料            
            CreateRoomReview(excelCompareList, dataGridView2, sortElemInfoList); // 建立「房間需求檢討」欄位資料            
            CreateExcelReview(excelCompareList, sortElemInfoList); // 未設置房間            
            CreateRoomReview(excelCompareList, sortElemInfoList); // 未校核房間            
            doorDatas = new List<DoorData>();
            doorDatas = DoorsInfo();
            CreateDoorReview(excelCompareList, dataGridView5, doorDatas); // 門尺寸校核
            CreateMechanicalReview(dataGridView6, sortElemInfoList); // 天花內淨高校核

            dgv1 = dataGridView1;
            dgv2 = dataGridView2;
            dgv6 = dataGridView6;
        }
        /// <summary>
        /// 選擇來源檔案
        /// </summary>
        /// <returns></returns>
        private Tuple<bool, string> ChooseFiles()
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Title = "請選擇房間校核Excel檔";
            ofd.InitialDirectory = ".\\";
            ofd.Filter = "Excel Files (*.xlsx)|*.xlsx|Excel Files (*.xls)|*.xls|All Files (*.*)|*.*";
            ofd.Multiselect = false; // 多選檔案
            bool trueOrFalse = false; // 預設未選取檔案
            string filePath = string.Empty; // 檔案路徑
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                trueOrFalse = true;
                filePath = ofd.FileName;
            }
            else
            {
                trueOrFalse = false;
            }

            return Tuple.Create(trueOrFalse, filePath);
        }
        /// <summary>
        /// 先檢視是否有設定好要移除的特殊符號
        /// </summary>
        /// <returns></returns>
        public static List<string> CreateCharsToRemoveTXT()
        {
            List<string> charsToRemove = new List<string>();
            string charsToRemovePath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\CharsToRemove.txt"; // 取得使用者文件路徑
            // 先檢查是否有此檔案, 沒有的話則新增
            if (!File.Exists(charsToRemovePath))
            {
                string[] signs = new string[] { "@", ",", ".", ";", "'", "(", ")", "_", "-", "\\", "/", " ", "\"" }; // 特殊符號
                foreach (string sign in signs)
                {
                    charsToRemove.Add(sign);
                }
                using (StreamWriter outputFile = new StreamWriter(charsToRemovePath))
                {
                    foreach (string sign in charsToRemove)
                    {
                        outputFile.WriteLine(sign);
                    }
                }
            }
            else
            {
                charsToRemove = new List<string>();
                using (StreamReader sr = new StreamReader(charsToRemovePath))
                {
                    string textContent;
                    while ((textContent = sr.ReadLine()) != null)
                    {
                        charsToRemove.Add(textContent);
                    }
                }
            }

            return charsToRemove;
        }
        /// <summary>
        /// 選取Excel比對數量差異
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="charsToRemove"></param>
        /// <returns></returns>
        public List<ExcelCompare> ReadExcel(string filePath, List<string> charsToRemove)
        {
            List<ExcelCompare> excelCompareList = new List<ExcelCompare>();

            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook workbook = excelApp.Workbooks.Open(filePath);
            Excel._Worksheet workSheet = workbook.Sheets[1];
            Excel.Range Range = workSheet.UsedRange;

            int rowCount = Range.Rows.Count;
            int colCount = Range.Columns.Count;

            // 記錄標頭的欄位數
            TitalNames titleNames = SaveTitleNames(colCount, workSheet);

            // 讀取Excel檔中, 所有物件的名稱、類別、數量
            for (int i = 2; i <= rowCount; i++)
            {
                // 空間名稱(中文)
                if (workSheet.Cells[i, titleNames.name].Value != null)
                {
                    ExcelCompare excelCompare = new ExcelCompare();
                    excelCompare = SaveExcelValue(excelCompare, titleNames, workSheet, charsToRemove, i); // 儲存Excel資料
                    excelCompareList.Add(excelCompare);
                }
            }

            // 清理記憶體
            GC.Collect();
            GC.WaitForPendingFinalizers();
            // 釋放COM對象的經驗法則, 單獨引用與釋放COM對象, 不要使用多"."釋放
            Marshal.ReleaseComObject(Range);
            Marshal.ReleaseComObject(workSheet);
            // 關閉與釋放
            workbook.Close();
            Marshal.ReleaseComObject(workbook);
            excelApp.Quit();
            Marshal.ReleaseComObject(excelApp);

            return excelCompareList;
        }
        /// <summary>
        /// 讀取標頭順序
        /// </summary>
        /// <param name="colCount"></param>
        /// <param name="workSheet"></param>
        /// <returns></returns>
        private TitalNames SaveTitleNames(int colCount, Excel._Worksheet workSheet)
        {
            TitalNames titleNames = new TitalNames();
            for (int i = 1; i <= colCount; i++)
            {
                string titleName = workSheet.Cells[1, i].Value;
                string title = titleName.Replace("\n", "");
                if (title.Equals("代碼"))
                {
                    titleNames.code = i;
                }
                else if (title.Equals("區域"))
                {
                    titleNames.classification = i;
                }
                else if (title.Equals("樓層"))
                {
                    titleNames.level = i;
                }
                else if (title.Equals("空間名稱(中文)"))
                {
                    titleNames.name = i;
                }
                else if (title.Equals("空間名稱(英文)"))
                {
                    titleNames.engName = i;
                }
                else if (title.Equals("其他名稱"))
                {
                    titleNames.otherName = i;
                }
                else if (title.Equals("類別") || title.Equals("設備/系統"))
                {
                    titleNames.category = i;
                }
                else if (title.Equals("數量"))
                {
                    titleNames.count = i;
                }
                else if (title.Equals("最大面積(m2)") || title.Equals("規範最大面積(m2)"))
                {
                    titleNames.maxArea = i;
                }
                else if (title.Equals("最小面積(m2)") || title.Equals("規範最小面積(m2)"))
                {
                    titleNames.minArea = i;
                }
                else if (title.Equals("需求面積(m2)"))
                {
                    titleNames.demandArea = i;
                }
                else if (title.Equals("容許差異(±%)") || title.Equals("面積容許差異(±%)"))
                {
                    titleNames.permit = i;
                }
                else if (title.Equals("規範最小寬度(m)"))
                {
                    titleNames.specificationMinWidth = i;
                }
                else if (title.Equals("實際最小寬度(m)") || title.Equals("需求最小寬度(m)"))
                {
                    titleNames.demandMinWidth = i;
                }
                else if (title.Equals("門(mm)"))
                {
                    titleNames.door = i;
                }

                if (title.Equals("規範淨高(m)"))
                {
                    titleNames.unboundedHeight = i;
                }
                else if (title.Equals("淨高(m)"))
                {
                    titleNames.unboundedHeight = i;
                }

                if (title.Equals("需求淨高(m)"))
                {
                    titleNames.demandUnboundedHeight = i;
                }
                else if (title.Equals("淨高(m)"))
                {
                    titleNames.demandUnboundedHeight = i;
                }
            }
            return titleNames;
        }
        /// <summary>
        /// 儲存Excel資料
        /// </summary>
        /// <param name="excelCompare"></param>
        /// <param name="titleNames"></param>
        /// <param name="workSheet"></param>
        /// <param name="charsToRemove"></param>
        /// <param name="i"></param>
        /// <returns></returns>
        private ExcelCompare SaveExcelValue(ExcelCompare excelCompare, TitalNames titleNames, Excel._Worksheet workSheet, List<string> charsToRemove, int i)
        {
            try
            {
                excelCompare.code = workSheet.Cells[i, titleNames.code].Value; // 代碼
                if (excelCompare.code == null)
                {
                    excelCompare.code = "";
                }
                excelCompare.classification = workSheet.Cells[i, titleNames.classification].Value; // 區域
                excelCompare.level = workSheet.Cells[i, titleNames.level].Value; // 樓層
                // 名稱(設定)
                string editName = workSheet.Cells[i, titleNames.name].Value;
                foreach (string c in charsToRemove)
                {
                    try
                    {
                        editName = editName.Replace(c, string.Empty); // 空間名稱(中文)
                    }
                    catch (Exception ex)
                    {
                        string error = ex.Message + "\n" + ex.ToString();
                    }
                }
                excelCompare.name = editName;
                try
                {
                    excelCompare.engName = workSheet.Cells[i, titleNames.engName].Value; // 空間名稱(英文)
                }
                catch (Exception)
                {
                    excelCompare.engName = "";
                }
                try
                {
                    // 檢查此空間名稱是否有"其他名稱", 有的話則過濾分隔符號後儲存
                    try
                    {
                        string otherFullName = workSheet.Cells[i, titleNames.otherName].Value; // 其他名稱
                        if (otherFullName != "")
                        {
                            otherFullName = otherFullName.Replace(",", "、").Replace("/", "、");
                            string[] otherNames = otherFullName.Split('、');
                            foreach (string otherName in otherNames)
                            {
                                excelCompare.otherNames.Add(otherName);
                            }
                        }
                    }
                    catch (Exception)
                    {
                        excelCompare.engName = "";
                    }
                }
                catch (Exception)
                {

                }
                try
                {
                    excelCompare.permit = workSheet.Cells[i, titleNames.permit].Value; // 容許差異
                }
                catch (Exception)
                {
                    excelCompare.permit = 0.0;
                }
                try
                {
                    excelCompare.unboundedHeight = workSheet.Cells[i, titleNames.unboundedHeight].Value; // 規範淨高
                }
                catch (Exception)
                {
                    excelCompare.unboundedHeight = 0.0;
                }
                try
                {
                    excelCompare.demandUnboundedHeight = workSheet.Cells[i, titleNames.demandUnboundedHeight].Value; // 需求淨高
                }
                catch (Exception)
                {
                    excelCompare.demandUnboundedHeight = 0.0;
                }
                string doorWidthHeight = workSheet.Cells[i, titleNames.door].Value;
                if (doorWidthHeight != null)
                {
                    try
                    {
                        excelCompare.doorWidth = Convert.ToDouble(doorWidthHeight.Split('x')[0]); // 門寬(mm)
                        excelCompare.doorHeight = Convert.ToDouble(doorWidthHeight.Split('x')[1]); // 門高(mm)
                    }
                    catch (Exception)
                    {
                        excelCompare.doorWidth = 0.0;
                        excelCompare.doorHeight = 0.0;
                    }
                }
                try
                {
                    excelCompare.maxArea = workSheet.Cells[i, titleNames.maxArea].Value; // 規範最大面積
                }
                catch (Exception)
                {
                    excelCompare.maxArea = 0;
                }
                try
                {
                    excelCompare.minArea = workSheet.Cells[i, titleNames.minArea].Value; // 規範最小面積
                }
                catch (Exception)
                {
                    excelCompare.minArea = 0;
                }
                try
                {
                    excelCompare.specificationMinWidth = workSheet.Cells[i, titleNames.specificationMinWidth].Value; // 規範最小寬度
                }
                catch (Exception)
                {
                    excelCompare.specificationMinWidth = 0;
                }
                try
                {
                    excelCompare.demandArea = workSheet.Cells[i, titleNames.demandArea].Value; // 需求面積
                }
                catch (Exception)
                {
                    excelCompare.demandArea = 0;
                }
                try
                {
                    excelCompare.demandMinWidth = workSheet.Cells[i, titleNames.demandMinWidth].Value; // 需求最小寬度
                }
                catch (Exception)
                {
                    excelCompare.demandMinWidth = 0;
                }
            }
            catch (Exception)
            {

            }

            return excelCompare;
        }
        /// <summary>
        /// Excel標題項次
        /// </summary>
        /// <returns></returns>
        private List<ExcelInfo> ExcelTitleItems()
        {
            List<ExcelInfo> excelInfoList = new List<ExcelInfo>();
            ExcelInfo excelInfo = new ExcelInfo();

            string[] titles = new string[] { "壹.", "結構體工程" };
            string[] items = new string[] { "項次", "名稱", "類別", "材料", "單位", "數量", "單價", "複價", "備註" };
            excelInfo = new ExcelInfo();
            excelInfo.titles = titles; // 標題
            excelInfo.items = items; // 項目
            excelInfoList.Add(excelInfo);

            titles = new string[] { "貳.", "門窗工程" };
            excelInfo = new ExcelInfo();
            excelInfo.titles = titles; // 標題
            excelInfo.items = items; // 項目
            excelInfoList.Add(excelInfo);

            titles = new string[] { "參.", "假設工程" };
            excelInfo = new ExcelInfo();
            excelInfo.titles = titles; // 標題
            excelInfo.items = items; // 項目
            excelInfoList.Add(excelInfo);

            return excelInfoList;
        }
        /// <summary>
        /// 呼叫Excel
        /// </summary>
        /// <param name="oldElemList"></param>
        /// <param name="sortElemInfoList"></param>
        private void CallExcel(List<ElementInfo> oldElemList, List<ElementInfo> sortElemInfoList)
        {
            // 創建Excel
            Excel.Application excelApp = new Excel.Application();
            excelApp.Visible = true; // 開啟Excel可見            
            Workbook workbook = excelApp.Workbooks.Add(); // 創建一個空的workbook

            // 讀取到當前的Sheet, 並修改名稱為全部顯示
            Excel._Worksheet workSheet = (Excel.Worksheet)excelApp.ActiveSheet;
            workSheet.Name = "(一)標單項目(5%)";

            // Excel標題項次
            List<ExcelInfo> excelInfoList = ExcelTitleItems();

            // 創建Excel
            CreateExcel(excelApp, workSheet, excelInfoList, oldElemList, sortElemInfoList);

            //// 新增Sheet
            //excelApp.Sheets.Add();
            //workSheet = (Excel.Worksheet)excelApp.ActiveSheet;
            //workSheet.Name = "分層統計";
            //// 創建Excel
            //CreateExcel(excelApp, workSheet, sortElemInfoList);

            // 儲存路徑
            string path = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\法規檢核.xlsx";
            FolderBrowserDialog dialog = new FolderBrowserDialog();
            dialog.Description = "請選擇檔案路徑";
            try
            {
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    path = dialog.SelectedPath + "\\";
                }
                if (path.EndsWith("\\"))
                {
                    workbook.SaveCopyAs(path + "法規檢核" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".xlsx");
                }
                else
                {
                    workbook.SaveCopyAs(path + "\\法規檢核" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".xlsx");
                }

                // 刪除EXCEL程序, 否則每執行一次會持續增加
                System.Diagnostics.Process[] procs = System.Diagnostics.Process.GetProcessesByName("EXCEL");
                foreach (System.Diagnostics.Process p in procs)
                {
                    p.Kill();
                }
            }
            catch (Exception)
            {
                path = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\法規檢核.xlsx";
            }

            //把執行的Excel資源釋放
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
            excelApp = null;
            workbook = null;
            workSheet = null;

            TaskDialog.Show("Revit", "完成");
        }
        /// <summary>
        /// 創建Excel
        /// </summary>
        /// <param name="excelApp"></param>
        /// <param name="workSheet"></param>
        /// <param name="excelInfoList"></param>
        /// <param name="oldElemList"></param>
        /// <param name="elementInfoList"></param>
        private void CreateExcel(Excel.Application excelApp, Excel._Worksheet workSheet, List<ExcelInfo> excelInfoList, List<ElementInfo> oldElemList, List<ElementInfo> elementInfoList)
        {
            // 凍結窗格
            // 1.先選定某一個儲存格當作範圍
            Range rangeFreezePoint = workSheet.get_Range("A4", "I4");
            // 2.選定此範圍所在的 Sheet
            rangeFreezePoint.Select();
            // 3.針對被選定的 Sheet 進行凍結視窗
            excelApp.ActiveWindow.FreezePanes = true;

            workSheet.Cells.Font.Name = "微軟正黑體"; // 設定Excel資料字體字型
            workSheet.Cells.Font.Size = 14; // 設定Excel資料字體大小

            // 設定標題欄名稱
            string alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
            int row = 1;
            workSheet.Cells[row, "A"] = "工程名稱：" + prjName;
            string Range = "A" + row + ":I" + row;
            // 設定框線
            workSheet.get_Range(Range).Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            // 合併儲存格
            Range = "A" + row + ":D" + row;
            workSheet.get_Range(Range).Merge(workSheet.get_Range(Range).MergeCells);

            row++;
            workSheet.Cells[row, "A"] = "業主名稱：" + prjOwner;
            Range = "A" + row + ":I" + row;
            // 設定框線
            workSheet.get_Range(Range).Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            // 合併儲存格
            Range = "A" + row + ":D" + row;
            workSheet.get_Range(Range).Merge(workSheet.get_Range(Range).MergeCells);

            row = 4;
            int startRow = row; // 起始Row
            string[] items = new string[] { "項次", "工程項目", "", "", "單位", "數量", "單價", "複價", "備註" };
            for (int numbers = 0; numbers < items.Count(); numbers++)
            {
                string count = alphabet[numbers].ToString();
                workSheet.Cells[row, count] = items[numbers];
            }
            // 設定樣式與背景色
            Range = "A" + startRow + ":I" + row;
            workSheet.get_Range(Range).Interior.Color = System.Drawing.Color.LightGreen;

            foreach (ExcelInfo excelInfo in excelInfoList)
            {
                row++;
                int numbers = 0;
                foreach (string title in excelInfo.titles)
                {
                    string count = alphabet[numbers].ToString();
                    workSheet.Cells[row, count] = title;
                    numbers++;
                    workSheet.Cells[row, "E"] = "式";
                    workSheet.Cells[row, "F"] = 1;
                    workSheet.Cells[row, "H"] = "=F" + row + "*G" + row;
                }
            }
            row++;
            int sumStartRow = startRow + 1;
            int sumEndRow = row - 1;
            workSheet.Cells[row, "B"] = "合計";
            Range = "B" + row + ":B" + row;
            workSheet.get_Range(Range).Font.Color = System.Drawing.Color.Red; //字體顏色
            Range = "H" + row + ":H" + row;
            workSheet.get_Range(Range).Font.Color = System.Drawing.Color.Red; //字體顏色
            workSheet.get_Range(Range).Interior.Color = System.Drawing.Color.Yellow; // 背景色
            workSheet.Cells[row, "H"] = "=SUM(H" + sumStartRow + ":H" + sumEndRow;
            Range = "A" + startRow + ":I" + row;
            SetStyle(workSheet, Range);

            row++;
            foreach (ExcelInfo excelInfo in excelInfoList)
            {
                row++;
                startRow = row; // 起始Row
                int numbers = 0;
                // 標題
                string titleNumber = excelInfo.titles[0];
                foreach (string title in excelInfo.titles)
                {
                    string count = alphabet[numbers].ToString();
                    workSheet.Cells[row, count] = title;
                    numbers++;
                }

                row++;
                numbers = 0;
                // 項目
                foreach (string item in excelInfo.items)
                {
                    string count = alphabet[numbers].ToString();
                    workSheet.Cells[row, count] = item;
                    numbers++;
                }
                // 設定背景色
                Range = "A" + row + ":I" + row;
                //設置儲存格的背景色
                workSheet.get_Range(Range).Interior.Color = System.Drawing.Color.LightGreen;

                // 項目內容
                try
                {
                    List<string> names = elementInfoList.Where(x => !x.name.Equals("")).Select(x => x.name).Distinct().ToList(); // 名稱篩選
                    List<string> typeNames = elementInfoList.Where(x => !x.type.Equals("")).Select(x => x.type).Distinct().ToList(); // 類型篩選

                    int itemNumber = 1;
                    foreach (string name in names)
                    {
                        foreach (string typeName in typeNames)
                        {
                            List<ElementInfo> elemInfoList = (from x in elementInfoList
                                                              where x.title.Equals(titleNumber) && x.name.Equals(name) && x.type.Equals(typeName)
                                                              select x).ToList();
                            ElementInfo defaultElem = elemInfoList.Select(x => x).FirstOrDefault();
                            double areaSum = elemInfoList.Select(x => x.area).Sum();
                            double volumeSum = elemInfoList.Select(x => x.volume).Sum();
                            if (elemInfoList.Count() > 0)
                            {
                                try
                                {
                                    row++;
                                    workSheet.Cells[row, "A"] = defaultElem.title + itemNumber;
                                    workSheet.Cells[row, "B"] = name;
                                    workSheet.Cells[row, "C"] = typeName;
                                    workSheet.Cells[row, "D"] = defaultElem.material;
                                    if (name.Contains("牆"))
                                    {
                                        workSheet.Cells[row, "E"] = " m³";
                                        double volume = elemInfoList.Select(x => x.volume).Sum();
                                        workSheet.Cells[row, "F"] = volume; // 體積加總
                                        if (oldElemList.Count > 0)
                                        {
                                            double oldVolume = oldElemList.Where(x => x.name.Equals(name) && x.type.Equals(typeName)).Select(x => x.count).Sum();
                                            if (!volume.Equals(oldVolume))
                                            {
                                                Range = "A" + row + ":I" + row;
                                                workSheet.get_Range(Range).Font.Color = System.Drawing.Color.DarkBlue; //字體顏色
                                                workSheet.get_Range(Range).Interior.Color = System.Drawing.Color.LightGoldenrodYellow; // 背景色
                                            }
                                        }
                                    }
                                    else if (name.Contains("門") || name.Contains("窗"))
                                    {
                                        workSheet.Cells[row, "E"] = "樘";
                                        workSheet.Cells[row, "F"] = elemInfoList.Count(); // 數量
                                        if (oldElemList.Count > 0)
                                        {
                                            double oldVolume = oldElemList.Where(x => x.name.Equals(name) && x.type.Equals(typeName)).Select(x => x.count).FirstOrDefault();
                                            if (elemInfoList.Count() != oldVolume)
                                            {
                                                Range = "A" + row + ":I" + row;
                                                workSheet.get_Range(Range).Font.Color = System.Drawing.Color.DarkBlue; //字體顏色
                                                workSheet.get_Range(Range).Interior.Color = System.Drawing.Color.LightGoldenrodYellow; // 背景色
                                            }
                                        }
                                    }
                                    else if (name.Contains("假設"))
                                    {
                                        workSheet.Cells[row, "E"] = "式";
                                        workSheet.Cells[row, "F"] = elemInfoList.Count(); // 數量
                                        if (oldElemList.Count > 0)
                                        {
                                            double oldVolume = oldElemList.Where(x => x.name.Equals(name) && x.type.Equals(typeName)).Select(x => x.count).FirstOrDefault();
                                            if (elemInfoList.Count() != oldVolume)
                                            {
                                                Range = "A" + row + ":I" + row;
                                                workSheet.get_Range(Range).Font.Color = System.Drawing.Color.DarkBlue; //字體顏色
                                                workSheet.get_Range(Range).Interior.Color = System.Drawing.Color.LightGoldenrodYellow; // 背景色
                                            }
                                        }
                                    }
                                    workSheet.Cells[row, "G"] = defaultElem.cost;
                                    workSheet.Cells[row, "H"] = "=F" + row + "*G" + row;
                                    workSheet.Cells[row, "I"] = "";
                                    itemNumber++;
                                }
                                catch (Exception)
                                {

                                }
                            }
                        }
                    }

                    // 小計
                    row++;
                    sumStartRow = startRow + 2;
                    sumEndRow = row - 1;
                    workSheet.Cells[row, "B"] = "合計";
                    workSheet.Cells[row, "H"] = "=SUM(H" + sumStartRow + ":H" + sumEndRow;
                    Range = "B" + row + ":B" + row;
                    workSheet.get_Range(Range).Font.Color = System.Drawing.Color.Red; //字體顏色
                    Range = "H" + row + ":H" + row;
                    workSheet.get_Range(Range).Font.Color = System.Drawing.Color.Red; //字體顏色
                    workSheet.get_Range(Range).Interior.Color = System.Drawing.Color.Yellow; // 背景色
                    // 標題對應的小計
                    if (titleNumber.Equals("壹."))
                    {
                        workSheet.Cells[5, "G"] = "=H" + row;
                    }
                    else if (titleNumber.Equals("貳."))
                    {
                        workSheet.Cells[6, "G"] = "=H" + row;
                    }
                    else if (titleNumber.Equals("參."))
                    {
                        workSheet.Cells[7, "G"] = "=H" + row;
                    }
                }
                catch (Exception)
                {

                }

                // 設定樣式
                Range = "A" + startRow + ":I" + row;
                SetStyle(workSheet, Range);

                row++;
            }
        }
        /// <summary>
        /// 樣式調整
        /// </summary>
        /// <param name="workSheet"></param>
        /// <param name="Range"></param>
        private void SetStyle(Excel._Worksheet workSheet, string Range)
        {
            // 設定框線
            workSheet.get_Range(Range).Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            //設定對齊方式, 置中
            workSheet.get_Range(Range).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            // 調整自動寬度
            for (int i = 1; i <= 8; i++)
            {
                workSheet.Columns[i].AutoFit();
            }
        }
        /// <summary>
        /// 儲存元件資料
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="charsToRemove"></param>
        /// <param name="excelCompareList"></param>
        /// <returns></returns>
        private List<ElementInfo> ModelInfo(Document doc, List<string> charsToRemove, List<ExcelCompare> excelCompareList)
        {
            //// 讀取專案中所有的樓板、天花板、地板, 找到它們的底面與高度
            //List<AllBottomFaces> allBottomFacesList = new List<AllBottomFaces>();
            //allBottomFacesList = AllBottomFacesList();

            List <ElementInfo> elementInfoList = new List<ElementInfo>();
            IList<ElementFilter> elementFilters = new List<ElementFilter>(); // 清空過濾器
            ElementCategoryFilter roomFilter = new ElementCategoryFilter(BuiltInCategory.OST_Rooms); // 房間
            elementFilters.Add(roomFilter);
            LogicalOrFilter logicalOrFilter = new LogicalOrFilter(elementFilters);
            List<Element> elems = new FilteredElementCollector(doc).WherePasses(logicalOrFilter).WhereElementIsNotElementType().Where(x => x.get_Parameter(BuiltInParameter.ROOM_AREA).AsDouble() != 0).ToList();
            List<Element> testElems = elems.Where(x => x.Id.IntegerValue.Equals((int)2854095)).Select(x => x).ToList();
            foreach (Element elem in elems)
            {
                if (elem is Room)
                {
                    try
                    {
                        ElementInfo elementInfo = new ElementInfo();
                        elementInfo.elem = elem; // 元件
                        Room room = elem as Room;
                        // 名稱(設定)
                        string editName = elem.get_Parameter(BuiltInParameter.ROOM_NAME).AsString();
                        foreach (string c in charsToRemove)
                        {
                            try { editName = editName.Replace(c, string.Empty); }
                            catch (Exception ex) { string error = ex.Message + "\n" + ex.ToString(); }
                        }
                        elementInfo.id = room.Id.ToString(); // id
                        elementInfo.name = elem.get_Parameter(BuiltInParameter.ROOM_NAME).AsString(); // 房間名稱
                        elementInfo.changeName = editName; // 房間比對後更新名稱
                        try { elementInfo.engName = elem.LookupParameter("房間英文名稱").AsString(); } // 房間英文名稱
                        catch (Exception) { elementInfo.engName = ""; }
                        elementInfo.type = elem.get_Parameter(BuiltInParameter.ELEM_CATEGORY_PARAM).AsValueString(); // 類型
                        elementInfo.level = elem.get_Parameter(BuiltInParameter.LEVEL_NAME).AsString(); // 樓層
                        // 樓層高度
                        try
                        {
                            LevelElevation levelElev = levelElevList.Where(x => x.level.Id.Equals(elem.get_Parameter(BuiltInParameter.ROOM_LEVEL_ID).AsElementId())).FirstOrDefault();
                            LevelElevation nextLevelElev = levelElevList.Where(x => x.sort.Equals(levelElev.sort + 1)).FirstOrDefault();
                            if(levelElev != null && nextLevelElev != null)
                            {
                                double levelHeight = nextLevelElev.elevation - levelElev.elevation;
                                elementInfo.levelHeight = levelHeight.ToString();
                            }
                        }
                        catch(Exception ex) { string error = ex.Message + "\n" + ex.ToString(); }
                        elementInfo.view = doc.GetElement(room.Level.FindAssociatedPlanViewId()) as Autodesk.Revit.DB.View;
                        elementInfo.unit = "㎡";
                        elementInfo.cost = 0.0; // 成本
                        elementInfo.perimeter = Convert.ToDouble(elem.get_Parameter(BuiltInParameter.ROOM_PERIMETER).AsString()); // 周長
                        elementInfo.area = Convert.ToDouble(elem.get_Parameter(BuiltInParameter.ROOM_AREA).AsValueString().Replace(" m²", "")); // 面積
                        elementInfo.sn = elem.get_Parameter(BuiltInParameter.ROOM_NUMBER).AsString(); // 編號
                        ExcelCompare excelCompare = IsHaveExcelData(excelCompareList, editName); // 是否有比對到Excel的資料
                        if (editName.Equals("男廁") || editName.Equals("女廁"))
                        {
                            try
                            {
                                if (elementInfo.engName.Contains("STAFF"))
                                {
                                    excelCompare = (from x in excelCompareList
                                                    where x.name.Equals(elementInfo.changeName) && !elementInfo.area.Equals(0) && x.engName.Contains("STAFF")
                                                    select x).OrderBy(x => x.permit).ThenBy(x => x.maxArea).ThenBy(x => x.minArea).ThenBy(x => x.demandArea).FirstOrDefault();
                                }
                                else
                                {
                                    excelCompare = (from x in excelCompareList
                                                    where x.name.Equals(elementInfo.changeName) && !elementInfo.area.Equals(0) && !x.engName.Contains("STAFF")
                                                    select x).OrderBy(x => x.permit).ThenBy(x => x.maxArea).ThenBy(x => x.minArea).ThenBy(x => x.demandArea).FirstOrDefault();
                                }
                            }
                            catch (Exception)
                            {

                            }
                        }
                        if (excelCompare != null)
                        {
                            elementInfo.code = excelCompare.code; // 代碼
                            string[] sort = excelCompare.code.Split('-');
                            elementInfo.sort1 = sort[0];
                            if (sort.Count() > 1)
                            {
                                try { elementInfo.sort2 = Convert.ToInt32(sort[1]); }
                                catch (Exception) { elementInfo.sort2 = 0; }
                            }
                        }
                        // 基準偏移
                        try
                        {
                            elementInfo.lowerOffset = Math.Round(UnitUtils.ConvertFromInternalUnits(elem.get_Parameter(BuiltInParameter.ROOM_LOWER_OFFSET).AsDouble(), DisplayUnitType.DUT_METERS), 4, MidpointRounding.AwayFromZero);
                        }
                        catch(Exception ex)
                        {
                            string error = ex.Message + "\n" + ex.ToString();
                        }
                        
                        // Room的頂部高程
                        double topElevation = Math.Round(UnitUtils.ConvertFromInternalUnits(room.UpperLimit.Elevation, DisplayUnitType.DUT_METERS), 2, MidpointRounding.AwayFromZero);
                        elementInfo.topElevation = topElevation;
                        // 取得Room所有邊界長度, 還有儲存Solid
                        List<List<Curve>> boundarySegments = GetBoundarySegment(room, elementInfo);
                        List<Curve> maxCurveList = new List<Curve>();
                        double maxPerimeter = 0.0;
                        foreach (List<Curve> boundarySegment in boundarySegments)
                        {
                            if (boundarySegment.Select(x => x.Length).Sum() > maxPerimeter)
                            {
                                maxPerimeter = boundarySegment.Select(x => x.Length).Sum();
                                maxCurveList = boundarySegment;
                            }
                        }
                        // 儲存長度, 並轉換單位為公尺
                        elementInfo.boundarySegments = maxCurveList.Select(x => Math.Round(UnitUtils.ConvertFromInternalUnits(x.Length, DisplayUnitType.DUT_METERS), 2, MidpointRounding.AwayFromZero)).OrderByDescending(x => x).ToList();

                        // 儲存Room的LocationPoint
                        LocationPoint lp = elem.Location as LocationPoint;
                        elementInfo.boundaryPoints.Add(lp.Point);
                        elementInfo.maxBoundaryLength = UnitUtils.ConvertFromInternalUnits(maxCurveList.Select(x => x.Length).Max(), DisplayUnitType.DUT_METERS); // 最長邊界

                        // 取得Room的底面
                        // 1.讀取Geometry Option
                        Options options = new Options();
                        options.DetailLevel = ViewDetailLevel.Medium;
                        options.ComputeReferences = true;
                        options.IncludeNonVisibleObjects = true;
                        // 得到幾何元素
                        GeometryElement geomElem = room.get_Geometry(options);
                        List<Solid> solids = GeometrySolids(geomElem);
                        elementInfo.bottomFaces = GetBottomFaces(solids);

                        //// 與所有上層樓板、天花板比較距離, 取得最短距離
                        //List<double> distances = new List<double>();
                        //distances.Add(0);
                        //foreach (XYZ roomPoint in elementInfo.boundaryPoints)
                        //{
                        //    distances = AllUpBottomFacesDistance(allBottomFacesList, roomPoint);
                        //}
                        //if (distances.Count() > 0)
                        //{
                        //    double minDistance = distances.Min();
                        //    elementInfo.roomHeight = minDistance - elementInfo.lowerOffset; // Room與上層樓板的高度距離
                        //}
                        elementInfoList.Add(elementInfo);
                    }
                    catch (Exception ex)
                    {
                        string error = elem.Id + "\n" + elem.Name + "\n" + ex.Message + "\n" + ex.ToString();
                    }
                }
            }

            return elementInfoList;
        }
        /// <summary>
        /// 記錄依Code排序的順序
        /// </summary>
        /// <param name="excelCompareList"></param>
        /// <returns></returns>
        private void RecordCodeSort(List<ExcelCompare> excelCompareList)
        {
            List<CodeSort> codeSortList = new List<CodeSort>();
            List<string> codeList = excelCompareList.Where(x => !String.IsNullOrEmpty(x.code)).Select(x => x.code).Distinct().OrderBy(x => x).ToList();
            foreach(string code in codeList)
            {
                try
                {
                    string[] codeNumber = code.Split('-');
                    CodeSort codeSort = new CodeSort();
                    codeSort.code = codeNumber[0];
                    if (codeNumber.Length > 1) { codeSort.number = Convert.ToInt32(codeNumber[1]); }
                    else { codeSort.number = 0; }
                    codeSortList.Add(codeSort);
                }
                catch(Exception ex) { string error = ex.Message + "\n" + ex.ToString(); }
            }
            codeSortList = codeSortList.OrderBy(x => x.code).ThenBy(x => x.number).ToList();
            int i = 1;
            foreach (CodeSort item in codeSortList)
            {
                CodeSort codeSort = new CodeSort();
                if (item.number == 0) { codeSort.code = item.code; }
                else { codeSort.code = item.code + "-" + item.number; }
                codeSort.sort = i;
                i++;
                orderByCode.Add(codeSort);
            }
            orderByCode.Add(new CodeSort() { code = "", sort = i }) ;
        }
        /// <summary>
        /// 依Code排序
        /// </summary>
        /// <param name="code"></param>
        /// <returns></returns>
        private int OrderByCode(string code)
        {
            int sort = orderByCode.Where(x => x.code.Equals(code)).Select(x => x.sort).FirstOrDefault();
            return sort;
        }
        /// <summary>
        /// 是否有比對到Excel的資料
        /// </summary>
        /// <param name="excelCompareList"></param>
        /// <param name="roomName"></param>
        /// <returns></returns>
        private ExcelCompare IsHaveExcelData(List<ExcelCompare> excelCompareList, string roomName)
        {
            ExcelCompare excelCompare = excelCompareList.Where(x => x.name.Equals(roomName))
                                        .OrderBy(x => x.permit).ThenBy(x => x.maxArea).ThenBy(x => x.minArea).ThenBy(x => x.demandArea).FirstOrDefault();
            // 名稱無完全一樣的話, 比對一下是否有涵蓋的
            if (excelCompare == null)
            {
                excelCompare = excelCompareList.Where(x => x.name.Contains(roomName))
                               .OrderBy(x => x.permit).ThenBy(x => x.maxArea).ThenBy(x => x.minArea).ThenBy(x => x.demandArea).FirstOrDefault();
                // 如果沒有, 反過來比對一次涵蓋的
                if (excelCompare == null)
                {
                    excelCompare = excelCompareList.Where(x => roomName.Contains(x.name))
                                   .OrderBy(x => x.permit).ThenBy(x => x.maxArea).ThenBy(x => x.minArea).ThenBy(x => x.demandArea).FirstOrDefault();
                    // 如果名稱Equals跟Contains都沒有, 比對一下其他名稱
                    if (excelCompare == null)
                    {
                        excelCompare = excelCompareList.Where(x => x.otherNames.Where(y => y.Equals(roomName)).FirstOrDefault() != null)
                                       .OrderBy(x => x.permit).ThenBy(x => x.maxArea).ThenBy(x => x.minArea).ThenBy(x => x.demandArea).FirstOrDefault();
                    }
                }
            }

            return excelCompare;
        }
        /// <summary>
        /// 是否有比對到模型的資料
        /// </summary>
        /// <param name="sortElemInfoList"></param>
        /// <param name="excelCompare"></param>
        /// <returns></returns>
        private ElementInfo IsHaveElemData(List<ElementInfo> sortElemInfoList, ExcelCompare excelCompare)
        {
            ElementInfo sortElemInfo = sortElemInfoList.Where(x => x.changeName.Equals(excelCompare.name)).FirstOrDefault();

            // 名稱無完全一樣的話, 比對一下是否有涵蓋的
            if (sortElemInfo == null)
            {
                sortElemInfo = sortElemInfoList.Where(x => x.changeName.Contains(excelCompare.name)).FirstOrDefault();
                // 如果沒有, 反過來比對一次涵蓋的
                if(sortElemInfo == null)
                {
                    sortElemInfo = sortElemInfoList.Where(x => !x.changeName.Equals("")).Where(x => excelCompare.name.Contains(x.changeName)).FirstOrDefault();
                    // 如果名稱Equals跟Contains都沒有, 比對一下其他名稱
                    if (sortElemInfo == null)
                    {
                        foreach (string otherName in excelCompare.otherNames)
                        {
                            sortElemInfo = sortElemInfoList.Where(x => x.changeName.Equals(otherName)).FirstOrDefault();
                            if (sortElemInfo != null)
                            {
                                break;
                            }
                        }
                    }
                }
            }

            return sortElemInfo;
        }
        /// <summary>
        /// 儲存門資料
        /// </summary>
        /// <returns></returns>
        private List<DoorData> DoorsInfo()
        {
            foreach(Document linkDoc in docList)
            {
                List<ElementFilter> elementFilters = new List<ElementFilter>(); // 清空過濾器
                ElementCategoryFilter doorFilter = new ElementCategoryFilter(BuiltInCategory.OST_Doors); // 門
                elementFilters.Add(doorFilter);
                LogicalOrFilter logicalOrFilter = new LogicalOrFilter(elementFilters);
                List<FamilyInstance> doors = new FilteredElementCollector(linkDoc).OfCategory(BuiltInCategory.OST_Doors).WhereElementIsNotElementType().Cast<FamilyInstance>().ToList();
                //// 要測試的門
                //doors = (from x in doors
                //         where x.Id.IntegerValue.Equals(590404) || x.Id.IntegerValue.Equals(590490) /*|| x.Id.IntegerValue.Equals(1173561)*/
                //         select x).ToList();

                // 移除鐵捲門
                List<FamilyInstance> ironDoors = (from x in doors
                                                  where x.Symbol.FamilyName.Contains("鐵捲")
                                                  select x).ToList();
                doors = doors.Except(ironDoors).ToList();

                // 所有不需要檢核的房間名稱
                List<string> unReviewRoomName = (from x in excelCompareList
                                                 where x.doorWidth == 0 && x.doorHeight == 0
                                                 select x.name).ToList();
                // 「From Room」和「To Room」判斷，如果門的兩者都不是需要被校核門的房間，就可以篩掉(門)。
                // 移除兩邊都為null的門
                List<FamilyInstance> nullDoors = doors.Where(x => x.FromRoom == null && x.ToRoom == null).ToList();
                doors = doors.Except(nullDoors).ToList();
                // 移除ToRoom不為null, 但不需要檢核的門
                List<FamilyInstance> fromRoomNullDoors = doors.Where(x => x.FromRoom == null).Where(a => unReviewRoomName.Any(t => t.Contains(a.ToRoom.get_Parameter(BuiltInParameter.ROOM_NAME).AsString()))).ToList();
                List<FamilyInstance> publicToilets = (from x in fromRoomNullDoors
                                                      where x.ToRoom.get_Parameter(BuiltInParameter.ROOM_NAME).AsString().Contains("男廁") || x.ToRoom.get_Parameter(BuiltInParameter.ROOM_NAME).AsString().Contains("女廁")
                                                      where x.ToRoom.LookupParameter("房間英文名稱").AsString() == null || !x.ToRoom.LookupParameter("房間英文名稱").AsString().Contains("STAFF")
                                                      select x).ToList();
                fromRoomNullDoors = fromRoomNullDoors.Except(publicToilets).ToList();
                doors = doors.Except(fromRoomNullDoors).ToList();
                // 移除FromRoom不為null, 但不需要檢核的門
                List<FamilyInstance> toRoomNullDoors = doors.Where(x => x.ToRoom == null).Where(a => unReviewRoomName.Any(t => t.Contains(a.FromRoom.get_Parameter(BuiltInParameter.ROOM_NAME).AsString()))).ToList();
                publicToilets = (from x in toRoomNullDoors
                                where x.FromRoom.get_Parameter(BuiltInParameter.ROOM_NAME).AsString().Contains("男廁") || x.FromRoom.get_Parameter(BuiltInParameter.ROOM_NAME).AsString().Contains("女廁")
                                where x.FromRoom.LookupParameter("房間英文名稱").AsString() == null || !x.FromRoom.LookupParameter("房間英文名稱").AsString().Contains("STAFF")
                                select x).ToList();
                toRoomNullDoors = toRoomNullDoors.Except(publicToilets).ToList();
                doors = doors.Except(toRoomNullDoors).ToList();
                // 房間名稱包含「管道、管道間、充氣室、...」等，此門不校核
                List<FamilyInstance> unReviewFromName = doors.Where(x => x.FromRoom != null).Where(x => x.FromRoom.get_Parameter(BuiltInParameter.ROOM_NAME).AsString().Contains("管道") ||
                                                                                                        x.FromRoom.get_Parameter(BuiltInParameter.ROOM_NAME).AsString().Contains("充氣室")).ToList();
                doors = doors.Except(unReviewFromName).ToList();
                List <FamilyInstance> unReviewToName = doors.Where(x => x.ToRoom != null).Where(x => x.ToRoom.get_Parameter(BuiltInParameter.ROOM_NAME).AsString().Contains("管道") ||
                                                                                                     x.ToRoom.get_Parameter(BuiltInParameter.ROOM_NAME).AsString().Contains("充氣室")).ToList();                
                doors = doors.Except(unReviewToName).ToList();

                foreach (FamilyInstance door in doors)
                {
                    DoorData doorData = new DoorData();
                    doorData.door = door;
                    try
                    {
                        doorData.level = linkDoc.GetElement(door.LevelId) as Level;
                    }
                    catch(Exception ex)
                    {
                        string error = ex.Message + "\n" + ex.ToString();
                    }
                    doorData.fromRoom = door.FromRoom;
                    doorData.toRoom = door.ToRoom;
                    try
                    {
                        string belong = "";
                        try
                        {
                            if (door.LookupParameter("Room") != null)
                            {
                                belong = door.LookupParameter("Room").AsString();
                            }
                            else
                            {
                                belong = door.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).AsString();
                            }
                        }
                        catch(Exception)
                        {
                            belong = door.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).AsString();
                        }
                        if(belong != null)
                        {
                            doorData.belong = belong;
                        }
                        else
                        {
                            doorData.belong = "";
                        }
                    }
                    catch(Exception)
                    {
                        doorData.belong = "";
                    }
                    // 高度
                    try
                    {
                        double height = door.Symbol.get_Parameter(BuiltInParameter.GENERIC_HEIGHT).AsDouble();
                        if(door.Symbol.LookupParameter("開門側框寬度") != null)
                        {
                            height = height - door.Symbol.LookupParameter("開門側框寬度").AsDouble();
                        }
                        else if (door.Symbol.LookupParameter("門扇高度(H)") != null)
                        {
                            height = door.Symbol.LookupParameter("門扇高度(H)").AsDouble();
                        }
                        else if (door.Symbol.LookupParameter("門扇高度") != null)
                        {
                            height = door.Symbol.LookupParameter("門扇高度").AsDouble();
                        }
                        doorData.height = Math.Round(UnitUtils.ConvertFromInternalUnits(height, DisplayUnitType.DUT_MILLIMETERS), 0, MidpointRounding.AwayFromZero);
                    }
                    catch (Exception ex)
                    {
                        string error = ex.Message + "\n" + ex.ToString();
                    }
                    // 寬度
                    try
                    {
                        double width = door.Symbol.get_Parameter(BuiltInParameter.GENERIC_WIDTH).AsDouble();
                        if (door.Symbol.LookupParameter("門扇總寬度(W)") != null)
                        {
                            width = door.Symbol.LookupParameter("門扇總寬度(W)").AsDouble();
                        }
                        else if (door.Symbol.LookupParameter("門扇寬度(左)") != null && door.Symbol.LookupParameter("門扇寬度(右)") != null)
                        {
                            width = door.Symbol.LookupParameter("門扇寬度(左)").AsDouble() + door.Symbol.LookupParameter("門扇寬度(右)").AsDouble();
                        }
                        else if (door.Symbol.LookupParameter("門扇寬度(W)") != null)
                        {
                            width = door.Symbol.LookupParameter("門扇寬度(W)").AsDouble();
                        }
                        else if (door.Symbol.LookupParameter("門扇寬度") != null)
                        {
                            width = door.Symbol.LookupParameter("門扇寬度").AsDouble();
                        }
                        doorData.width = Math.Round(UnitUtils.ConvertFromInternalUnits(width, DisplayUnitType.DUT_MILLIMETERS), 0, MidpointRounding.AwayFromZero);
                    }
                    catch (Exception ex)
                    {
                        string error = ex.Message + "\n" + ex.ToString();
                    }

                    try
                    {
                        WallType wallType = linkDoc.GetElement(door.Host.GetTypeId()) as WallType;
                        if (wallType != null)
                        {
                            if (wallType.FamilyName.Equals("Curtain Wall") || wallType.FamilyName.Equals("帷幕牆"))
                            {
                                string wallTypeFN = wallType.FamilyName;
                                if (!doorData.height.Equals(0) && !doorData.width.Equals(0))
                                {
                                    doorDatas.Add(doorData);
                                }
                            }
                            else
                            {
                                doorDatas.Add(doorData);
                            }
                        }
                        else
                        {
                            doorDatas.Add(doorData);
                        }
                    }
                    catch (NullReferenceException ex)
                    {
                        string error = ex.Message + "\n" + ex.ToString();
                    }
                }
            }
            // 排序
            try
            {
                doorDatas = (from x in doorDatas
                             select x).OrderByDescending(x => x.fromRoom != null && x.toRoom != null).ThenByDescending(x => x.fromRoom != null)
                                      .ThenByDescending(x => x.toRoom != null).ThenBy(x => x.door.Symbol.FamilyName)
                                      .ThenBy(x => x.door.Symbol.Name).ThenBy(x => x.level.Elevation).ThenBy(x => x.id).ToList();
            }
            catch (Exception ex)
            {
                string error = ex.Message + "\n" + ex.ToString();
            }

            return doorDatas;
        }
        /// <summary>
        /// 讀取Room所有的Edge
        /// </summary>
        /// <param name="room"></param>
        /// <param name="elementInfo"></param>
        /// <returns></returns>
        private List<List<Curve>> GetBoundarySegment(Room room, ElementInfo elementInfo)
        {
            List<List<Curve>> boundarySegment = new List<List<Curve>>();
            try
            {
                // 1.讀取Geometry Option
                Options options = new Options();
                //options.View = doc.GetElement(room.Level.FindAssociatedPlanViewId()) as Autodesk.Revit.DB.View;
                options.DetailLevel = ViewDetailLevel.Medium;
                options.ComputeReferences = true;
                options.IncludeNonVisibleObjects = true;
                // 得到幾何元素
                GeometryElement geomElem = room.get_Geometry(options);
                List<Solid> solids = GeometrySolids(geomElem);
                elementInfo.solids = solids; // Room的Solids
                Solid solid = solids.FirstOrDefault();
                List<Face> planarFaces = new List<Face>();
                foreach(Face face in solid.Faces)
                {
                    if(face is PlanarFace)
                    {
                        PlanarFace planarFace = (PlanarFace)face;
                        if(planarFace.FaceNormal.Z > 0)
                        {
                            planarFaces.Add(planarFace);
                        }
                    }
                }
                foreach (PlanarFace planarFace in planarFaces)
                {
                    foreach (EdgeArray edgeArray in planarFace.EdgeLoops)
                    {
                        // 先將所有的邊儲存起來
                        List<Curve> curveLoop = new List<Curve>();
                        foreach (Edge edge in edgeArray)
                        {
                            Curve curve = edge.AsCurve();
                            curveLoop.Add(curve);
                        }
                        boundarySegment.Add(SameEdge(curveLoop));
                    }
                }
                // 從Solid查詢干涉到的元件, 找到樓梯、電扶梯
                //foreach (Solid roomSolid in solids)
                //{
                //    FindTheIntersectElems(doc, roomSolid, elementInfo);
                //}
            }
            catch (Exception ex)
            {
                string error = ex.Message + "\n" + ex.ToString();
            }

            return boundarySegment;
        }
        /// <summary>
        /// 取得Room的Solid
        /// </summary>
        /// <param name="geoObj"></param>
        /// <returns></returns>
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
        /// <summary>
        /// 從Solid查詢干涉到的元件, 找到樓梯、電扶梯
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="solid"></param>
        /// <param name="elementInfo"></param>
        private void FindTheIntersectElems(Document doc, Solid solid, ElementInfo elementInfo)
        {
            ElementIntersectsSolidFilter solidFilter = new ElementIntersectsSolidFilter(solid);
            List<Element> intersectElems = new FilteredElementCollector(doc).WherePasses(solidFilter).WhereElementIsNotElementType().ToList();
            List<Element> filterIntersectElems = (from x in intersectElems
                                                  where x is Stairs || x is FamilyInstance
                                                  select x).ToList();
            foreach(Element intersectElem in filterIntersectElems)
            {
                if(intersectElem is Stairs)
                {
                    elementInfo.intersectElems.Add(intersectElem);
                }
                else if(intersectElem is FamilyInstance)
                {
                    FamilyInstance fi = intersectElem as FamilyInstance;
                    if(fi.get_Parameter(BuiltInParameter.ELEM_FAMILY_PARAM) != null)
                    {
                        string paraName = fi.get_Parameter(BuiltInParameter.ELEM_FAMILY_PARAM).AsValueString();
                        if (paraName.Contains("電扶梯") && !paraName.Contains("間縫"))
                        {
                            elementInfo.intersectElems.Add(intersectElem);
                        }
                    }
                }
            }
        }
        /// <summary>
        /// 加總Room同邊長
        /// </summary>
        /// <param name="curveLoop"></param>
        /// <returns></returns>
        private List<Curve> SameEdge(List<Curve> curveLoop)
        {
            List<List<Curve>> curveList = new List<List<Curve>>();
            List<Curve> curves = new List<Curve>();
            // 驗證之後的Curve是否同向量
            foreach (Curve curve in curveLoop)
            {
                if(curve is Line)
                {
                    if(curves.Count == 0)
                    {
                        curves.Add(curve);
                    }
                    else
                    {
                        XYZ listDirection = new XYZ();
                        if (curves[0] is Arc)
                        {
                            // 確認curves中的direction為何
                            Arc listLine = (Arc)curves[0];
                            listDirection = listLine.XDirection;
                        }
                        else
                        {
                            Line listLine = (Line)curves[0];
                            listDirection = listLine.Direction;
                        }
                        // 確認curves中的direction為何
                        try
                        {
                            //Line listLine = (Line)curves[0];
                            //XYZ listDirection = listLine.Direction;
                            Line line = (Line)curve;
                            XYZ direction = line.Direction;
                            if (Math.Round(direction.X, 4, MidpointRounding.AwayFromZero) == Math.Round(listDirection.X, 4, MidpointRounding.AwayFromZero) &&
                               Math.Round(direction.Y, 4, MidpointRounding.AwayFromZero) == Math.Round(listDirection.Y, 4, MidpointRounding.AwayFromZero) &&
                               Math.Round(direction.Z, 4, MidpointRounding.AwayFromZero) == Math.Round(listDirection.Z, 4, MidpointRounding.AwayFromZero))
                            {
                                curves.Add(curve);
                            }
                            else
                            {
                                curveList.Add(curves);
                                curves = new List<Curve>();
                                curves.Add(curve);
                            }
                        }
                        catch (Exception)
                        {

                        }
                    }
                }
                else
                {
                    if(curves.Count > 0)
                    {
                        curveList.Add(curves);
                    }
                    curves = new List<Curve>();
                    curves.Add(curve);
                }
            }
            if(curves.Count > 0)
            {
                curveList.Add(curves);
            }

            // 將邊界加總成一邊
            curves = new List<Curve>();
            foreach(List<Curve> curves1 in curveList)
            {
                if(curves1.Count > 0)
                {
                    XYZ startPoint = curves1[0].Tessellate()[0];
                    XYZ endPoint = curves1[curves1.Count - 1].Tessellate()[1];
                    Line line = null;
                    if (curves1.Count > 1)
                    {
                        List<double> allX = new List<double>();
                        List<double> allY = new List<double>();
                        // 找到x、y的最大最小值
                        foreach (Curve curve in curves1)
                        {
                            allX.Add(curve.Tessellate()[0].X);
                            allX.Add(curve.Tessellate()[1].X);
                            allY.Add(curve.Tessellate()[0].Y);
                            allY.Add(curve.Tessellate()[1].Y);
                        }
                        double maxX = allX.Max();
                        double minX = allX.Min();
                        double maxY = allY.Max();
                        double minY = allY.Min();
                        startPoint = new XYZ(minX, minY, curves1[0].Tessellate()[0].Z);
                        endPoint = new XYZ(maxX, maxY, curves1[0].Tessellate()[0].Z);
                    }
                    try
                    {
                        line = Line.CreateBound(startPoint, endPoint);
                    }
                    catch (Exception ex)
                    {
                        string error = ex.Message + "\n" + ex.ToString();
                        //startPoint = curves1[curves1.Count - 1].Tessellate()[0];
                        //endPoint = curves1[0].Tessellate()[1];
                        //line = Line.CreateBound(startPoint, endPoint);
                    }
                    curves.Add(line);
                }
            }
            // 檢查最後一段curve與開始的是否同一direction, 相同的話即連結
            Line startLine = (Line)curves[0];
            XYZ startDirection = startLine.Direction;
            Line endLine = (Line)curves[curves.Count - 1];
            XYZ endDirection = endLine.Direction;
            if (Math.Round(startDirection.X, 4, MidpointRounding.AwayFromZero) == Math.Round(endDirection.X, 4, MidpointRounding.AwayFromZero) &&
                Math.Round(startDirection.Y, 4, MidpointRounding.AwayFromZero) == Math.Round(endDirection.Y, 4, MidpointRounding.AwayFromZero) &&
                Math.Round(startDirection.Z, 4, MidpointRounding.AwayFromZero) == Math.Round(endDirection.Z, 4, MidpointRounding.AwayFromZero))
            {
                // 查詢正確的線段起終點
                Line line = null;
                try
                {
                    XYZ samePoint = (from x in endLine.Tessellate()
                                     where Math.Round(x.DistanceTo(startLine.Tessellate()[0]), 4, MidpointRounding.AwayFromZero).Equals(0) || 
                                           Math.Round(x.DistanceTo(startLine.Tessellate()[1]), 4, MidpointRounding.AwayFromZero).Equals(0)
                                     select x).FirstOrDefault();
                    XYZ startPoint = (from x in startLine.Tessellate()
                                      where x.DistanceTo(samePoint) > 0
                                      select x).FirstOrDefault();
                    XYZ endPoint = (from x in endLine.Tessellate()
                                    where x.DistanceTo(samePoint) > 0
                                    select x).FirstOrDefault();
                    line = Line.CreateBound(startPoint, endPoint);
                    // 移除curves內舊的Line
                    curves.RemoveAt(0);
                    curves.RemoveAt(curves.Count - 1);
                    // 新增建立的Line
                    curves.Add(line);
                }
                catch(Exception ex)
                {
                    string error = ex.Message + "\n" + ex.ToString();
                }
            }

            return curves;
        }
        /// <summary>
        /// 取得Room的底面
        /// </summary>
        /// <param name="solidList"></param>
        /// <returns></returns>
        private List<Face> GetBottomFaces(List<Solid> solidList)
        {
            List<Face> bottomFaces = new List<Face>();
            foreach (Solid solid in solidList)
            {
                foreach (Face face in solid.Faces)
                {
                    if (face.ComputeNormal(new UV(0.5, 0.5)).IsAlmostEqualTo(XYZ.BasisZ.Negate())) { bottomFaces.Add(face); }
                    //double faceTZ = face.ComputeNormal(new UV(0.5, 0.5)).Z;
                    //if (faceTZ == -1.0) { bottomFaces.Add(face); } // 底面
                }
            }
            return bottomFaces;
        }
        /// <summary>
        /// 旋轉角度
        /// </summary>
        /// <param name="pointA"></param>
        /// <param name="pointB"></param>
        /// <returns></returns>
        public static double PointRotation(XYZ pointA, XYZ pointB)
        {
            XYZ pA = new XYZ(pointA.X, pointA.Y, 0);
            XYZ pB = new XYZ(pointB.X, pointB.Y, 0);
            double Dx = pB.X - pA.X;
            double Dy = pB.Y - pA.Y;
            double DRoation = Math.Atan2(Dy, Dx);
            double WRotation = DRoation / Math.PI * 180;

            return WRotation;
        }
        /// <summary>
        /// 讀取專案中所有的樓板、天花板, 找到它們的底面與高度
        /// </summary>
        /// <returns></returns>
        private List<AllBottomFaces> AllBottomFacesList()
        {
            // 找到各樓層樓板、天花板的底面
            List<AllBottomFaces> allBottomFacesList = new List<AllBottomFaces>();
            foreach (Document allDoc in docList)
            {
                try
                {
                    List<Floor> floors = new FilteredElementCollector(allDoc).OfCategory(BuiltInCategory.OST_Floors).WhereElementIsNotElementType().OfType<Floor>().Cast<Floor>().ToList();
                    //floors = floors.Where(x => x.get_Parameter(BuiltInParameter.FLOOR_PARAM_IS_STRUCTURAL).AsInteger().Equals(1)).ToList(); // 只計算結構樓板
                    List<Ceiling> ceilings = new FilteredElementCollector(allDoc).OfCategory(BuiltInCategory.OST_Ceilings).WhereElementIsNotElementType().OfType<Ceiling>().Cast<Ceiling>().ToList();
                    List<Stairs> stairs = new FilteredElementCollector(allDoc).OfCategory(BuiltInCategory.OST_Stairs).WhereElementIsNotElementType().OfType<Stairs>().Cast<Stairs>().ToList();
                    List<FootPrintRoof> footPrintRoofs = new FilteredElementCollector(allDoc).OfCategory(BuiltInCategory.OST_Roofs).WhereElementIsNotElementType().OfType<FootPrintRoof>().Cast<FootPrintRoof>().ToList();
                    List<ExtrusionRoof> extrusionRoofs = new FilteredElementCollector(allDoc).OfCategory(BuiltInCategory.OST_Roofs).WhereElementIsNotElementType().OfType<ExtrusionRoof>().Cast<ExtrusionRoof>().ToList();
                    AllBottomFaces allBottomFaces = new AllBottomFaces();
                    foreach (Floor floor in floors)
                    {
                        allBottomFaces = new AllBottomFaces();
                        allBottomFaces.upElem = floor; // 樓板
                        // GetBottomFaces返回的是引用列表，而不是元素或面
                        foreach (Reference getBottomFace in HostObjectUtils.GetBottomFaces(floor))
                        {
                            try
                            {
                                // GetElement返回一個元素, (如果傳遞一個引用，這個元素包含被引用的幾何圖形)
                                Element elem = allDoc.GetElement(getBottomFace);
                                // 擁有的引用和它來自的元素中獲取Face
                                PlanarFace planarFace = elem.GetGeometryObjectFromReference(getBottomFace) as PlanarFace;
                                allBottomFaces.bottomFaces.Add(planarFace);
                                // 樓板的底部高程
                                allBottomFaces.bottomElevation = UnitUtils.ConvertFromInternalUnits(floor.get_Parameter(BuiltInParameter.STRUCTURAL_ELEVATION_AT_BOTTOM).AsDouble(), DisplayUnitType.DUT_METERS);
                                allBottomFacesList.Add(allBottomFaces);
                            }
                            catch (Exception ex)
                            {
                                string error = ex.Message + "\n" + ex.ToString();
                            }
                        }
                    }
                    foreach (Ceiling ceiling in ceilings)
                    {
                        allBottomFaces = new AllBottomFaces();
                        allBottomFaces.upElem = ceiling; // 天花板
                                                         // GetBottomFaces返回的是引用列表，而不是元素或面
                        foreach (Reference getBottomFace in HostObjectUtils.GetBottomFaces(ceiling))
                        {
                            try
                            {
                                // GetElement返回一個元素, (如果傳遞一個引用，這個元素包含被引用的幾何圖形)
                                Element elem = allDoc.GetElement(getBottomFace);
                                // 擁有的引用和它來自的元素中獲取Face
                                PlanarFace planarFace = elem.GetGeometryObjectFromReference(getBottomFace) as PlanarFace;
                                allBottomFaces.bottomFaces.Add(planarFace);
                                // 天花板樓層高度
                                Level level = allDoc.GetElement(ceiling.LevelId) as Level;
                                double elevation = level.Elevation;
                                // 天花板的底部高程
                                allBottomFaces.bottomElevation = UnitUtils.ConvertFromInternalUnits(elevation + ceiling.get_Parameter(BuiltInParameter.CEILING_HEIGHTABOVELEVEL_PARAM).AsDouble(), DisplayUnitType.DUT_METERS);
                                allBottomFacesList.Add(allBottomFaces);
                            }
                            catch (Exception ex)
                            {
                                string error = ex.Message + "\n" + ex.ToString();
                            }
                        }
                    }
                    foreach (FootPrintRoof footPrintRoof in footPrintRoofs)
                    {
                        allBottomFaces = new AllBottomFaces();
                        allBottomFaces.upElem = footPrintRoof; // 屋頂
                        // GetBottomFaces返回的是引用列表，而不是元素或面
                        foreach (Reference getBottomFace in HostObjectUtils.GetBottomFaces(footPrintRoof))
                        {
                            try
                            {
                                // GetElement返回一個元素, (如果傳遞一個引用，這個元素包含被引用的幾何圖形)
                                Element elem = allDoc.GetElement(getBottomFace);
                                // 擁有的引用和它來自的元素中獲取Face
                                PlanarFace planarFace = elem.GetGeometryObjectFromReference(getBottomFace) as PlanarFace;
                                if(planarFace == null)
                                {
                                    // 1.讀取Geometry Option
                                    Options options = new Options();
                                    //options.View = doc.GetElement(room.Level.FindAssociatedPlanViewId()) as Autodesk.Revit.DB.View;
                                    options.DetailLevel = ViewDetailLevel.Medium;
                                    options.ComputeReferences = true;
                                    options.IncludeNonVisibleObjects = true;
                                    // 得到幾何元素
                                    GeometryElement geomElem = elem.get_Geometry(options);
                                    List<Solid> solids = GeometrySolids(geomElem);
                                    foreach(Solid solid in solids)
                                    {
                                        FaceArray faces = solid.Faces;
                                        foreach(Face face in faces)
                                        {
                                            PlanarFace pf = face as PlanarFace;
                                            if(pf.FaceNormal.Z < 0)
                                            {
                                                planarFace = pf;
                                                break;
                                            }
                                        }
                                    }
                                }
                                allBottomFaces.bottomFaces.Add(planarFace);
                                // 屋頂樓層高度
                                Level level = allDoc.GetElement(footPrintRoof.LevelId) as Level;
                                double elevation = level.Elevation;
                                // 屋頂的底部高程
                                allBottomFaces.bottomElevation = UnitUtils.ConvertFromInternalUnits(elevation + footPrintRoof.get_Parameter(BuiltInParameter.ROOF_LEVEL_OFFSET_PARAM).AsDouble(), DisplayUnitType.DUT_METERS);
                                allBottomFacesList.Add(allBottomFaces);
                            }
                            catch (Exception ex)
                            {
                                string error = ex.Message + "\n" + ex.ToString();
                            }
                        }
                    }
                    foreach (ExtrusionRoof extrusionRoof in extrusionRoofs)
                    {
                        allBottomFaces = new AllBottomFaces();
                        allBottomFaces.upElem = extrusionRoof; // 屋頂

                        // GetBottomFaces返回的是引用列表，而不是元素或面
                        foreach (Reference getBottomFace in HostObjectUtils.GetBottomFaces(extrusionRoof))
                        {
                            try
                            {
                                // GetElement返回一個元素, (如果傳遞一個引用，這個元素包含被引用的幾何圖形)
                                Element elem = allDoc.GetElement(getBottomFace);

                                // 擁有的引用和它來自的元素中獲取Face
                                PlanarFace planarFace = elem.GetGeometryObjectFromReference(getBottomFace) as PlanarFace;
                                if (planarFace == null)
                                {
                                    // 1.讀取Geometry Option
                                    Options options = new Options();
                                    //options.View = doc.GetElement(room.Level.FindAssociatedPlanViewId()) as Autodesk.Revit.DB.View;
                                    options.DetailLevel = ViewDetailLevel.Medium;
                                    options.ComputeReferences = true;
                                    options.IncludeNonVisibleObjects = true;
                                    // 得到幾何元素
                                    GeometryElement geomElem = elem.get_Geometry(options);
                                    List<Solid> solids = GeometrySolids(geomElem);
                                    foreach (Solid solid in solids)
                                    {
                                        FaceArray faces = solid.Faces;
                                        foreach (Face face in faces)
                                        {
                                            PlanarFace pf = face as PlanarFace;
                                            if (pf.FaceNormal.Z < 0)
                                            {
                                                planarFace = pf;
                                                break;
                                            }
                                        }
                                    }
                                }
                                // 屋頂樓層高度(參考樓層)
                                Level level = allDoc.GetElement(extrusionRoof.get_Parameter(BuiltInParameter.ROOF_CONSTRAINT_LEVEL_PARAM).AsElementId()) as Level;
                                double elevation = level.Elevation;
                                allBottomFaces.bottomFaces.Add(planarFace);
                                // 屋頂的底部高程
                                //allBottomFaces.bottomElevation = UnitUtils.ConvertFromInternalUnits(elevation + extrusionRoof.get_Parameter(BuiltInParameter.ROOF_LEVEL_OFFSET_PARAM).AsDouble(), DisplayUnitType.DUT_METERS);
                                allBottomFaces.bottomElevation = UnitUtils.ConvertFromInternalUnits(elevation, DisplayUnitType.DUT_METERS);
                                allBottomFacesList.Add(allBottomFaces);
                            }
                            catch (Exception ex)
                            {
                                string error = ex.Message + "\n" + ex.ToString();
                            }
                        }
                    }

                    // 找到各樓層FamilyInstance的樓板、天花板、樓梯底面
                    List<Element> docFIFloorCeilings = new List<Element>();
                    IList<ElementFilter> elementFilters = new List<ElementFilter>(); // 清空過濾器
                    ElementCategoryFilter floorFilter = new ElementCategoryFilter(BuiltInCategory.OST_Floors); // FamilyInstance的樓板
                    ElementCategoryFilter ceilingFilter = new ElementCategoryFilter(BuiltInCategory.OST_Ceilings); // FamilyInstance的天花板
                    ElementCategoryFilter stairsFilter = new ElementCategoryFilter(BuiltInCategory.OST_Stairs); // 樓梯            
                    elementFilters.Add(floorFilter);
                    elementFilters.Add(ceilingFilter);
                    elementFilters.Add(stairsFilter);
                    LogicalOrFilter logicalOrFilter = new LogicalOrFilter(elementFilters);
                    List<Element> elems = new FilteredElementCollector(allDoc).WherePasses(logicalOrFilter).WhereElementIsNotElementType().ToList();
                    List<Element> filterElems = elems.Where(x => x is FamilyInstance || x is Stairs).ToList();
                    // 電扶梯
                    List<FamilyInstance> genericModels = new FilteredElementCollector(allDoc).OfCategory(BuiltInCategory.OST_GenericModel).OfType<FamilyInstance>().Cast<FamilyInstance>().ToList();
                    List<FamilyInstance> escalators = (from x in genericModels
                                                       where x.Symbol.FamilyName.Contains("電扶梯") && !x.Symbol.FamilyName.Contains("間縫")
                                                       select x).ToList();
                    foreach (FamilyInstance escalator in escalators)
                    {
                        filterElems.Add(escalator);
                    }

                    // 找到3D視圖
                    List<View3D> view3Ds = new FilteredElementCollector(allDoc).OfClass(typeof(View3D)).WhereElementIsNotElementType().Cast<View3D>().ToList();
                    View3D view3D = view3Ds.Where(x => x.Name.Equals("{3D}")).FirstOrDefault();
                    if (view3D != null)
                    {
                        foreach (Element elem in filterElems)
                        {
                            try
                            {
                                allBottomFaces = new AllBottomFaces();
                                allBottomFaces.upElem = elem; // 樓板、天花板
                                                              // 讀取Geometry Option
                                Options options = new Options();
                                options.View = view3D;
                                options.ComputeReferences = true;
                                options.IncludeNonVisibleObjects = true;
                                // 得到幾何元素
                                GeometryElement geomElem = elem.get_Geometry(options);
                                List<Solid> solids = GeometrySolids(geomElem);
                                //Solid solid = solids.FirstOrDefault();
                                double bottomElevation = 0.0; // 樓板、天花板的底部高程
                                foreach (Solid solid in solids)
                                {
                                    foreach (Face face in solid.Faces)
                                    {
                                        if (face is PlanarFace)
                                        {
                                            // 取得底面或斜板
                                            PlanarFace planarFace = face as PlanarFace;
                                            //if (planarFace.ComputeNormal(new UV(0.5, 0.5)).IsAlmostEqualTo(XYZ.BasisZ.Negate()))
                                            if (planarFace.FaceNormal.Z < 0)
                                            {
                                                try
                                                {
                                                    // 樓板、天花板、樓梯的底部高程
                                                    double faceOriginZ = UnitUtils.ConvertFromInternalUnits(planarFace.Origin.Z, DisplayUnitType.DUT_METERS);
                                                    if (faceOriginZ > bottomElevation)
                                                    {
                                                        bottomElevation = faceOriginZ;
                                                        allBottomFaces.bottomElevation = bottomElevation;
                                                    }
                                                    allBottomFaces.bottomFaces.Add(planarFace);
                                                }
                                                catch (Exception ex)
                                                {
                                                    string error = ex.Message + "\n" + ex.ToString();
                                                }
                                            }
                                        }
                                    }
                                }
                                if (allBottomFaces.bottomFaces.Count > 0)
                                {
                                    allBottomFacesList.Add(allBottomFaces);
                                }
                            }
                            catch (Exception ex)
                            {
                                string error = ex.Message + "\n" + ex.ToString();
                            }
                        }
                    }
                }
                catch (Autodesk.Revit.Exceptions.ArgumentNullException)
                {

                }
            }

            return allBottomFacesList;
        }
        /// <summary>
        /// 與所有上層樓板、天花板比較距離, 取得最短距離
        /// </summary>
        /// <param name="allBottomFacesList"></param>
        /// <param name="roomPoint"></param>
        /// <returns></returns>
        private List<double> AllUpBottomFacesDistance(List<AllBottomFaces> allBottomFacesList, XYZ roomPoint)
        {
            List<double> distances = new List<double>();
            // 找到高於roomPoint, z軸的面
            double roomZ = UnitUtils.ConvertFromInternalUnits(roomPoint.Z, DisplayUnitType.DUT_METERS);
            List<AllBottomFaces> upBottomFaces = allBottomFacesList.Where(x => x.bottomElevation > roomZ).ToList();
            foreach(AllBottomFaces upBottomFace in upBottomFaces)
            {
                try
                {
                    // 樓板、天花板、樓梯的底部高程的面 > Room高度
                    List<PlanarFace> planarFaces = (from x in upBottomFace.bottomFaces
                                                    where UnitUtils.ConvertFromInternalUnits(x.Origin.Z, DisplayUnitType.DUT_METERS) > roomZ
                                                    select x).ToList();
                    foreach (PlanarFace planarFace in planarFaces)
                    {
                        try
                        {
                            double distance = planarFace.Project(roomPoint).Distance;
                            distance = Math.Round(UnitUtils.ConvertFromInternalUnits(distance, DisplayUnitType.DUT_METERS), 2, MidpointRounding.AwayFromZero);
                            if (distance > 0)
                            {
                                distances.Add(distance);
                            }
                        }
                        catch (Exception ex)
                        {
                            string error = ex.Message + "\n" + ex.ToString();
                        }
                    }
                }
                catch(Exception ex)
                {
                    string error = ex.Message + "\n" + ex.ToString();
                }
            }

            return distances;
        }
        /// <summary>
        /// 建立「房間規範檢討」欄位資料
        /// </summary>
        /// <param name="excelCompareList"></param>
        /// <param name="dataGridView"></param>
        /// <param name="sortElemInfoList"></param>
        private void CreateRoomReview(List<ExcelCompare> excelCompareList, DataGridView dataGridView, List<ElementInfo> sortElemInfoList)
        {
            if(dataGridView.Columns.Count == 0)
            {
                DataGridViewTextBoxColumn dgvCol1 = new DataGridViewTextBoxColumn();
                DataGridViewTextBoxColumn dgvCol2 = new DataGridViewTextBoxColumn();
                DataGridViewTextBoxColumn dgvCol3 = new DataGridViewTextBoxColumn();
                DataGridViewTextBoxColumn dgvCol4 = new DataGridViewTextBoxColumn();
                DataGridViewTextBoxColumn dgvCol5 = new DataGridViewTextBoxColumn();
                DataGridViewTextBoxColumn dgvCol6 = new DataGridViewTextBoxColumn();
                DataGridViewTextBoxColumn dgvCol7 = new DataGridViewTextBoxColumn();
                DataGridViewTextBoxColumn dgvCol8 = new DataGridViewTextBoxColumn();
                DataGridViewTextBoxColumn dgvCol9 = new DataGridViewTextBoxColumn();
                DataGridViewTextBoxColumn dgvCol10 = new DataGridViewTextBoxColumn();
                dgvCol1.Name = "dgvCol1";
                dgvCol2.Name = "dgvCol2";
                dgvCol3.Name = "dgvCol3";
                dgvCol4.Name = "dgvCol4";
                dgvCol5.Name = "dgvCol5";
                dgvCol6.Name = "dgvCol6";
                dgvCol7.Name = "dgvCol7";
                dgvCol8.Name = "dgvCol8";
                dgvCol9.Name = "dgvCol9";
                dgvCol10.Name = "dgvCol10";
                if (dataGridView.Name.Equals("dataGridView1"))
                {
                    dgvCol1.HeaderText = "代碼";
                    dgvCol2.HeaderText = "名稱";
                    dgvCol3.HeaderText = "樓層";
                    dgvCol4.HeaderText = "規範面積(㎡)";
                    dgvCol5.HeaderText = "面積(㎡)";
                    dgvCol6.HeaderText = "規範寬度(m)";
                    dgvCol7.HeaderText = "寬度(m)";
                    dgvCol8.HeaderText = "規範淨高(m)";
                    dgvCol9.HeaderText = "淨高(m)";
                    dgvCol10.HeaderText = "ID";
                    dataGridView.Columns.AddRange(new DataGridViewColumn[] { dgvCol1, dgvCol2, dgvCol3, dgvCol4, dgvCol5, dgvCol6, dgvCol7, dgvCol8, dgvCol9, dgvCol10 });
                }
                else
                {
                    dgvCol1.HeaderText = "代碼";
                    dgvCol2.HeaderText = "名稱";
                    dgvCol3.HeaderText = "樓層";
                    dgvCol4.HeaderText = "需求面積(㎡)";
                    dgvCol5.HeaderText = "面積(㎡)";
                    dgvCol6.HeaderText = "需求寬度(m)";
                    dgvCol7.HeaderText = "寬度(m)";
                    dgvCol8.HeaderText = "需求淨高(m)";
                    dgvCol9.HeaderText = "淨高(m)";
                    dgvCol10.HeaderText = "ID";
                    dataGridView.Columns.AddRange(new DataGridViewColumn[] { dgvCol1, dgvCol2, dgvCol3, dgvCol4, dgvCol5,/* dgvCol6, dgvCol7,*/ dgvCol8, dgvCol9, dgvCol10 });
                }
                dataGridView.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            }
            dataGridView.Dock = DockStyle.Fill;
            //是否允許使用者編輯
            dataGridView.ReadOnly = true;
            //是否允許使用者自行新增
            dataGridView.AllowUserToAddRows = false;
            //// 欄位標題自動寬度
            //dataGridView.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells;
            dataGridView.Rows.Clear(); // 清除所有Rows

            int i = 0; // n列
            double maxArea = 0.0; // 最大面積
            double minArea = 0.0; // 最小面積
            foreach (ElementInfo sortElemInfo in sortElemInfoList)
            {
                try
                {
                    ExcelCompare excelCompare = excelCompareList.Where(x => x.code.Equals(sortElemInfo.code)).FirstOrDefault();
                    if (excelCompare == null || excelCompare.code.Equals(""))
                    {
                        if (!sortElemInfo.changeName.Equals("男廁") && !sortElemInfo.changeName.Equals("女廁"))
                        { excelCompare = IsHaveExcelData(excelCompareList, sortElemInfo.changeName); }
                    }
                    if (excelCompare != null)
                    {                        
                        if (dataGridView.Name.Equals("dataGridView1"))
                        {
                            if (!excelCompare.count.Equals(0) || !excelCompare.maxArea.Equals(0) || !excelCompare.minArea.Equals(0) || !excelCompare.specificationMinWidth.Equals(0) || !excelCompare.unboundedHeight.Equals(0))
                            {
                                dataGridView.Rows.Add();
                                for (int cellCount = 0; cellCount <= 9; cellCount++)
                                {
                                    dataGridView.Rows[i].Cells[cellCount].Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                                }
                                dataGridView.Rows[i].Cells[0].Value = excelCompare.code; // 代碼
                                dataGridView.Rows[i].Cells[1].Value = sortElemInfo.name; // 名稱
                                dataGridView.Rows[i].Cells[2].Value = sortElemInfo.level; // 樓層
                                // 1.有最大最小面積限制: 面積<最小面積為NG, 面積> 最大面積*110 % 為NG, 最大面積 < 面積 <= 最大面積 * 110 % 為容許範圍
                                // (舊版規則)有最大面積和最小面積限制: 以最大面積值的正容許差異(%)和最小面積值之間為容許值
                                if (!excelCompare.maxArea.Equals(0) && !excelCompare.minArea.Equals(0))
                                {
                                    maxArea = excelCompare.maxArea + (excelCompare.maxArea * excelCompare.permit);
                                    minArea = excelCompare.minArea/* - (excelCompare.minArea * excelCompare.permit)*/;
                                    dataGridView.Rows[i].Cells[3].Value = minArea + "㎡~" + maxArea + "㎡";
                                    dataGridView.Rows[i].Cells[4].Value = sortElemInfo.area.ToString();
                                    if (sortElemInfo.area > maxArea || sortElemInfo.area < minArea)
                                    {
                                        dataGridView.Rows[i].Cells[4].Style.ForeColor = System.Drawing.Color.Red;
                                        dataGridView.Rows[i].DefaultCellStyle.BackColor = System.Drawing.Color.LightYellow;
                                    }
                                    else if (sortElemInfo.area > excelCompare.maxArea && sortElemInfo.area <= maxArea)
                                    {
                                        double maxPermit = 100 + excelCompare.permit * 100; // 正容許差異
                                                                                            // 面積容許差異不為0的話
                                        if (!excelCompare.permit.Equals(0))
                                        {
                                            dataGridView.Rows[i].Cells[4].Value = sortElemInfo.area + " ≦ " + excelCompare.maxArea + "*" + maxPermit + "% = " + maxArea;
                                        }
                                        dataGridView.Rows[i].Cells[4].Style.ForeColor = System.Drawing.Color.Blue;
                                    }
                                }
                                // 2.有最大面積限制: 面積 > 最大面積 * 110 % 為NG, 最大面積 < 面積 <= 最大面積 * 110 % 為容許範圍，面積 < 最大面積 * 90 % 為容許範圍
                                // (舊版規則)只有最大面積限制: 以最大面積值的正負容許差異(%)範圍檢討
                                else if (!excelCompare.maxArea.Equals(0) && excelCompare.minArea.Equals(0))
                                {
                                    maxArea = excelCompare.maxArea + (excelCompare.maxArea * excelCompare.permit);
                                    minArea = excelCompare.maxArea - (excelCompare.maxArea * excelCompare.permit);
                                    //dataGridView.Rows[i].Cells[3].Value = minArea + "㎡~" + maxArea + "㎡ MAX";
                                    dataGridView.Rows[i].Cells[3].Value = excelCompare.maxArea + "㎡ MAX";
                                    dataGridView.Rows[i].Cells[4].Value = sortElemInfo.area.ToString();

                                    // 面積容許差異不為0的話
                                    if (!excelCompare.permit.Equals(0))
                                    {
                                        // 房間面積 > 最大面積+正容許差異
                                        if (sortElemInfo.area > maxArea)
                                        {
                                            dataGridView.Rows[i].Cells[4].Style.ForeColor = System.Drawing.Color.Red;
                                            dataGridView.Rows[i].DefaultCellStyle.BackColor = System.Drawing.Color.LightYellow;
                                        }
                                        // 最大面積 < 房間面積 <= 最大面積+正容許差異
                                        else if (sortElemInfo.area > excelCompare.maxArea && sortElemInfo.area <= maxArea)
                                        {
                                            double maxPermit = 100 + excelCompare.permit * 100; // 正容許差異
                                            dataGridView.Rows[i].Cells[4].Value = sortElemInfo.area + " ≦ " + excelCompare.maxArea + "*" + maxPermit + "% = " + maxArea;
                                            dataGridView.Rows[i].Cells[4].Style.ForeColor = System.Drawing.Color.Blue;
                                        }
                                        // 房間面積 < 最大面積-負容許差異
                                        else if (sortElemInfo.area < minArea)
                                        {
                                            double maxPermit = 100 - excelCompare.permit * 100; // 負容許差異
                                            dataGridView.Rows[i].Cells[4].Value = sortElemInfo.area + " < " + excelCompare.maxArea + "*" + maxPermit + "% = " + minArea;
                                            dataGridView.Rows[i].Cells[4].Style.ForeColor = System.Drawing.Color.Blue;
                                        }
                                        //// 最大面積+負容許差異 <= 房間面積 < 最大面積
                                        //else if (sortElemInfo.area >= minArea && sortElemInfo.area < excelCompare.maxArea)
                                        //{
                                        //    double maxPermit = 100 - excelCompare.permit * 100; // 負容許差異
                                        //    dataGridView.Rows[i].Cells[4].Value = sortElemInfo.area + " ≧ " + excelCompare.maxArea + "*" + maxPermit + "% = " + minArea;
                                        //}
                                    }
                                    else
                                    {
                                        // 房間面積 > 最大面積+正容許差異
                                        if (sortElemInfo.area > maxArea)
                                        {
                                            dataGridView.Rows[i].Cells[4].Style.ForeColor = System.Drawing.Color.Red;
                                            dataGridView.Rows[i].DefaultCellStyle.BackColor = System.Drawing.Color.LightYellow;
                                        }
                                    }
                                }
                                // 3.有最小面積限制: 面積<最小面積為NG, 面積> 最小面積*110 % 為容許範圍
                                // (舊版規則)只有最小面積限制: 以最小面積值檢討，不考慮容許差異，也沒有上限。
                                else if (excelCompare.maxArea.Equals(0) && !excelCompare.minArea.Equals(0))
                                {
                                    maxArea = excelCompare.minArea + (excelCompare.minArea * excelCompare.permit);
                                    minArea = excelCompare.minArea;
                                    dataGridView.Rows[i].Cells[3].Value = minArea + "㎡ MIN";
                                    dataGridView.Rows[i].Cells[4].Value = sortElemInfo.area.ToString();

                                    // 面積容許差異不為0的話
                                    if (!excelCompare.permit.Equals(0))
                                    {
                                        if (sortElemInfo.area < minArea)
                                        {
                                            dataGridView.Rows[i].Cells[4].Style.ForeColor = System.Drawing.Color.Red;
                                            dataGridView.Rows[i].DefaultCellStyle.BackColor = System.Drawing.Color.LightYellow;
                                        }
                                        else if (sortElemInfo.area > maxArea)
                                        {
                                            double maxPermit = 100 + excelCompare.permit * 100; // 正容許差異
                                            dataGridView.Rows[i].Cells[4].Value = sortElemInfo.area + " > " + excelCompare.minArea + "*" + maxPermit + "% = " + maxArea;
                                            dataGridView.Rows[i].Cells[4].Style.ForeColor = System.Drawing.Color.Blue;
                                        }
                                    }
                                    else
                                    {
                                        if (sortElemInfo.area < minArea)
                                        {
                                            dataGridView.Rows[i].Cells[4].Style.ForeColor = System.Drawing.Color.Red;
                                            dataGridView.Rows[i].DefaultCellStyle.BackColor = System.Drawing.Color.LightYellow;
                                        }
                                    }
                                }
                                // 無最大最小面積限制
                                else
                                {
                                    dataGridView.Rows[i].Cells[3].Value = "NA"; // 需求面積
                                    dataGridView.Rows[i].Cells[4].Value = sortElemInfo.area.ToString();
                                }
                                // 規範寬度 & 最大邊長
                                if (excelCompare.specificationMinWidth.Equals(0))
                                {
                                    dataGridView.Rows[i].Cells[5].Value = "NA";
                                    //dataGridView.Rows[i].Cells[4].Value = Math.Round(sortElemInfo.maxBoundaryLength, 2, MidpointRounding.AwayFromZero);
                                    string lengthList = string.Empty;
                                    if (sortElemInfo.boundarySegments.Count >= 4)
                                    {
                                        for (int a = 0; a < 4; a++)
                                        {
                                            double length = sortElemInfo.boundarySegments[a];
                                            lengthList += length.ToString();
                                            if (a < 3)
                                            {
                                                lengthList += "、";
                                            }
                                            else if (sortElemInfo.boundarySegments.Count > 4)
                                            {
                                                lengthList += "...";
                                            }
                                        }
                                    }
                                    else
                                    {
                                        int a = 0;
                                        foreach (double length in sortElemInfo.boundarySegments)
                                        {
                                            lengthList += length.ToString();
                                            if (a < sortElemInfo.boundarySegments.Count)
                                            {
                                                lengthList += "、";
                                                a++;
                                            }
                                        }
                                    }
                                    dataGridView.Rows[i].Cells[6].Value = lengthList;
                                }
                                else
                                {
                                    dataGridView.Rows[i].Cells[5].Value = excelCompare.specificationMinWidth.ToString();
                                    //dataGridView.Rows[i].Cells[6].Value = Math.Round(sortElemInfo.maxBoundaryLength, 2, MidpointRounding.AwayFromZero);
                                    string lengthList = string.Empty;
                                    if (sortElemInfo.boundarySegments.Count >= 4)
                                    {
                                        for (int a = 0; a < 4; a++)
                                        {
                                            double length = sortElemInfo.boundarySegments[a];
                                            lengthList += length.ToString();
                                            if (a < 3)
                                            {
                                                lengthList += "、";
                                            }
                                            else if (sortElemInfo.boundarySegments.Count > 4)
                                            {
                                                lengthList += "...";
                                            }
                                        }
                                    }
                                    else
                                    {
                                        int a = 0;
                                        foreach (double length in sortElemInfo.boundarySegments)
                                        {
                                            lengthList += length.ToString();
                                            if (a < sortElemInfo.boundarySegments.Count)
                                            {
                                                lengthList += "、";
                                                a++;
                                            }
                                        }
                                    }
                                    dataGridView.Rows[i].Cells[6].Value = lengthList;
                                    //if (sortElemInfo.maxBoundaryLength < excelCompare.specificationMinWidth)
                                    //{
                                    //    dataGridView.Rows[i].Cells[6].Style.ForeColor = System.Drawing.Color.Red;
                                    //    dataGridView.Rows[i].DefaultCellStyle.BackColor = System.Drawing.Color.LightYellow;
                                    //}
                                }
                                // 規範淨高
                                if (excelCompare.unboundedHeight == 0.0)
                                {
                                    dataGridView.Rows[i].Cells[7].Value = "NA";
                                }
                                else
                                {
                                    dataGridView.Rows[i].Cells[7].Value = excelCompare.unboundedHeight.ToString();
                                }
                                try
                                {
                                    // 有干涉到樓梯或電扶梯的提醒
                                    if (sortElemInfo.intersectElems.Count() > 0)
                                    {
                                        dataGridView.Rows[i].Cells[8].Value = sortElemInfo.roomHeight.ToString() + " ⭐"; // 實際淨高
                                    }
                                    else
                                    {
                                        dataGridView.Rows[i].Cells[8].Value = sortElemInfo.roomHeight.ToString(); // 實際淨高
                                    }
                                }
                                catch (Exception ex)
                                {
                                    string error = ex.Message + "\n" + ex.ToString();
                                    dataGridView.Rows[i].Cells[8].Value = "null"; // 實際淨高
                                }
                                //if (sortElemInfo.roomHeight < excelCompare.unboundedHeight)
                                //{
                                //    dataGridView.Rows[i].Cells[8].Style.ForeColor = System.Drawing.Color.Red;
                                //    dataGridView.Rows[i].DefaultCellStyle.BackColor = System.Drawing.Color.LightYellow;
                                //}
                                dataGridView.Rows[i].Cells[9].Value = sortElemInfo.id; // ID
                                i++;
                            }
                        }
                        else
                        {
                            if (!excelCompare.count.Equals(0) || !excelCompare.demandArea.Equals(0) || !excelCompare.demandMinWidth.Equals(0) || !excelCompare.demandUnboundedHeight.Equals(0))
                            {
                                dataGridView.Rows.Add();
                                for (int cellCount = 0; cellCount <= 7; cellCount++)
                                {
                                    dataGridView.Rows[i].Cells[cellCount].Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                                }
                                dataGridView.Rows[i].Cells[0].Value = excelCompare.code; // 代碼
                                dataGridView.Rows[i].Cells[1].Value = sortElemInfo.name; // 名稱
                                dataGridView.Rows[i].Cells[2].Value = sortElemInfo.level; // 樓層

                                maxArea = excelCompare.demandArea + (excelCompare.demandArea * excelCompare.permit);
                                minArea = excelCompare.demandArea - (excelCompare.demandArea * excelCompare.permit);
                                if (excelCompare.demandArea.Equals(0))
                                {
                                    dataGridView.Rows[i].Cells[3].Value = "NA"; // 需求面積
                                    dataGridView.Rows[i].Cells[4].Value = sortElemInfo.area.ToString();
                                }
                                else
                                {
                                    dataGridView.Rows[i].Cells[3].Value = minArea + "㎡~" + maxArea + "㎡"; // 需求面積
                                    dataGridView.Rows[i].Cells[4].Value = sortElemInfo.area.ToString();
                                    if (sortElemInfo.area > maxArea || sortElemInfo.area < minArea)
                                    {
                                        dataGridView.Rows[i].Cells[4].Style.ForeColor = System.Drawing.Color.Red;
                                        dataGridView.Rows[i].DefaultCellStyle.BackColor = System.Drawing.Color.LightYellow;
                                    }
                                }
                                //// 需求寬度 & 最大邊長
                                //if (excelCompare.specificationMinWidth.Equals(0))
                                //{
                                //    dataGridView.Rows[i].Cells[5].Value = "NA";
                                //    //dataGridView.Rows[i].Cells[6].Value = Math.Round(sortElemInfo.minBoundaryLength, 2, MidpointRounding.AwayFromZero);
                                //    string lengthList = string.Empty;
                                //    for (int a = 0; a < 4; a++)
                                //    {
                                //        double length = Math.Round(UnitUtils.ConvertFromInternalUnits(sortElemInfo.boundarySegments[a], DisplayUnitType.DUT_METERS), 2, MidpointRounding.AwayFromZero);
                                //        lengthList += length.ToString() + ", ";
                                //    }
                                //    dataGridView.Rows[i].Cells[6].Value = lengthList;
                                //}
                                //else
                                //{
                                //    dataGridView.Rows[i].Cells[5].Value = excelCompare.demandMinWidth;
                                //    //dataGridView.Rows[i].Cells[6].Value = Math.Round(sortElemInfo.minBoundaryLength, 2, MidpointRounding.AwayFromZero);
                                //    string lengthList = string.Empty;
                                //    for (int a = 0; a < 4; a++)
                                //    {
                                //        double length = Math.Round(UnitUtils.ConvertFromInternalUnits(sortElemInfo.boundarySegments[a], DisplayUnitType.DUT_METERS), 2, MidpointRounding.AwayFromZero);
                                //        lengthList += length.ToString() + ", ";
                                //    }
                                //    dataGridView.Rows[i].Cells[6].Value = lengthList;
                                //    //if (sortElemInfo.minBoundaryLength < excelCompare.demandMinWidth)
                                //    //{
                                //    //    dataGridView.Rows[i].Cells[6].Style.ForeColor = System.Drawing.Color.Red;
                                //    //    dataGridView.Rows[i].DefaultCellStyle.BackColor = System.Drawing.Color.LightYellow;
                                //    //}
                                //}
                                // 需求淨高
                                if (excelCompare.demandUnboundedHeight == 0.0)
                                {
                                    dataGridView.Rows[i].Cells[5].Value = "NA";
                                }
                                else
                                {
                                    dataGridView.Rows[i].Cells[5].Value = excelCompare.demandUnboundedHeight.ToString();
                                }
                                try
                                {
                                    // 有干涉到樓梯或電扶梯的提醒
                                    if (sortElemInfo.intersectElems.Count() > 0)
                                    {
                                        dataGridView.Rows[i].Cells[6].Value = sortElemInfo.roomHeight.ToString() + " ⭐"; // 實際淨高
                                    }
                                    else
                                    {
                                        dataGridView.Rows[i].Cells[6].Value = sortElemInfo.roomHeight.ToString(); // 實際淨高
                                    }
                                }
                                catch (Exception ex)
                                {
                                    string error = ex.Message + "\n" + ex.ToString();
                                    dataGridView.Rows[i].Cells[6].Value = "null"; // 實際淨高
                                }
                                //if (sortElemInfo.roomHeight < excelCompare.demandUnboundedHeight)
                                //{
                                //    dataGridView.Rows[i].Cells[6].Style.ForeColor = System.Drawing.Color.Red;
                                //    dataGridView.Rows[i].DefaultCellStyle.BackColor = System.Drawing.Color.LightYellow;
                                //}
                                dataGridView.Rows[i].Cells[7].Value = sortElemInfo.id; // ID
                                i++;
                            }
                        }
                    }
                }
                catch (Exception)
                {

                }
            }
        }
        /// <summary>
        /// 未設置房間：建立Excel中未建立Room檢核
        /// </summary>
        /// <param name="excelCompareList"></param>
        /// <param name="sortElemInfoList"></param>
        private void CreateExcelReview(List<ExcelCompare> excelCompareList, List<ElementInfo> sortElemInfoList)
        {
            if (dataGridView3.Columns.Count == 0)
            {
                DataGridViewTextBoxColumn dgvCol1 = new DataGridViewTextBoxColumn();
                DataGridViewTextBoxColumn dgvCol2 = new DataGridViewTextBoxColumn();
                DataGridViewTextBoxColumn dgvCol3 = new DataGridViewTextBoxColumn();
                DataGridViewTextBoxColumn dgvCol4 = new DataGridViewTextBoxColumn();
                DataGridViewTextBoxColumn dgvCol5 = new DataGridViewTextBoxColumn();
                dgvCol1.Name = "dgvCol1";
                dgvCol1.HeaderText = "代碼";
                dgvCol2.Name = "dgvCol2";
                dgvCol2.HeaderText = "名稱";
                dgvCol3.Name = "dgvCol3";
                dgvCol3.HeaderText = "樓層";
                dgvCol4.Name = "dgvCol4";
                dgvCol4.HeaderText = "規範面積(㎡)";
                dgvCol5.Name = "dgvCol5";
                dgvCol5.HeaderText = "需求面積(㎡)";
                dataGridView3.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                dataGridView3.Columns.AddRange(new DataGridViewColumn[] { dgvCol1, dgvCol2, dgvCol3, dgvCol4, dgvCol5 });
            }
            dataGridView3.Dock = DockStyle.Fill;
            //是否允許使用者編輯
            dataGridView3.ReadOnly = true;
            //是否允許使用者自行新增
            dataGridView3.AllowUserToAddRows = false;
            dataGridView3.Rows.Clear(); // 清除所有Rows

            int i = 0; // n列
            foreach (ExcelCompare excelCompare in excelCompareList)
            {
                ElementInfo sortElemInfo = IsHaveElemData(sortElemInfoList, excelCompare); // 是否有比對到模型的資料
                if (sortElemInfo == null)
                {
                    dataGridView3.Rows.Add();
                    for (int cellCount = 0; cellCount < 5; cellCount++)
                    {
                        dataGridView3.Rows[i].Cells[cellCount].Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    }
                    dataGridView3.Rows[i].Cells[0].Value = excelCompare.code; // 代碼
                    dataGridView3.Rows[i].Cells[1].Value = excelCompare.name; // 名稱
                    // 樓層
                    if (!String.IsNullOrEmpty(excelCompare.level)) { dataGridView3.Rows[i].Cells[2].Value = excelCompare.level; }
                    else { dataGridView3.Rows[i].Cells[2].Value = ""; }


                    double maxArea = excelCompare.maxArea; // 最大面積
                    double minArea = excelCompare.minArea; // 最小面積
                    double demandArea = excelCompare.demandArea; // 需求面積
                    // 規範面積：最大最小面積皆不為0
                    if (!excelCompare.maxArea.Equals(0) && !excelCompare.minArea.Equals(0))
                    {
                        dataGridView3.Rows[i].Cells[3].Value = minArea + "㎡~" + maxArea + "㎡";
                    }
                    // 規範面積：只有最大面積不為0
                    else if (!excelCompare.maxArea.Equals(0) && excelCompare.minArea.Equals(0))
                    {
                        dataGridView3.Rows[i].Cells[3].Value = maxArea + "㎡ MAX";
                    }
                    // 規範面積：只有最小面積不為0
                    else if (excelCompare.maxArea.Equals(0) && !excelCompare.minArea.Equals(0))
                    {
                        dataGridView3.Rows[i].Cells[3].Value = minArea + "㎡ MIN";
                    }
                    else
                    {
                        dataGridView3.Rows[i].Cells[3].Value = "NA";
                    }
                    // 需求面積
                    if (!excelCompare.demandArea.Equals(0))
                    {
                            dataGridView3.Rows[i].Cells[4].Value = excelCompare.demandArea + "㎡";
                    }
                    else
                    {
                        dataGridView3.Rows[i].Cells[4].Value = "NA";
                    }
                    i++;
                }
            }
        }
        /// <summary>
        /// 未校核房間：建立Revit有放置Room但Excel沒有的名稱
        /// </summary>
        /// <param name="excelCompareList"></param>
        /// <param name="sortElemInfoList"></param>
        private void CreateRoomReview(List<ExcelCompare> excelCompareList, List<ElementInfo> sortElemInfoList)
        {
            if (dataGridView4.Columns.Count == 0)
            {
                DataGridViewTextBoxColumn dgvCol1 = new DataGridViewTextBoxColumn();
                DataGridViewTextBoxColumn dgvCol2 = new DataGridViewTextBoxColumn();
                DataGridViewTextBoxColumn dgvCol3 = new DataGridViewTextBoxColumn();
                DataGridViewTextBoxColumn dgvCol4 = new DataGridViewTextBoxColumn();
                dgvCol1.Name = "dgvCol1";
                dgvCol1.HeaderText = "名稱";
                dgvCol2.Name = "dgvCol2";
                dgvCol2.HeaderText = "樓層";
                dgvCol3.Name = "dgvCol3";
                dgvCol3.HeaderText = "面積(㎡)";
                dgvCol4.Name = "dgvCol4";
                dgvCol4.HeaderText = "ID";
                dataGridView4.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                dataGridView4.Columns.AddRange(new DataGridViewColumn[] { dgvCol1, dgvCol2, dgvCol3, dgvCol4 });
            }
            dataGridView4.Dock = DockStyle.Fill;
            //是否允許使用者編輯
            dataGridView4.ReadOnly = true;
            //是否允許使用者自行新增
            dataGridView4.AllowUserToAddRows = false;
            dataGridView4.Rows.Clear(); // 清除所有Rows

            int i = 0; // n列
            foreach (ElementInfo elemInfo in sortElemInfoList)
            {
                //ExcelCompare excelCompare = excelCompareList.Where(x => x.code.Equals(elemInfo.code)).FirstOrDefault();
                //if (excelCompare == null) { excelCompare = IsHaveExcelData(excelCompareList, elemInfo.changeName); }
                ExcelCompare excelCompare =  excelCompare = IsHaveExcelData(excelCompareList, elemInfo.changeName);
                //ExcelCompare excelCompare = IsHaveExcelData(excelCompareList, elemInfo.changeName); // 是否有比對到Excel的資料
                if (excelCompare == null)
                {
                    dataGridView4.Rows.Add();
                    for (int cellCount = 0; cellCount <= 3; cellCount++)
                    {
                        dataGridView4.Rows[i].Cells[cellCount].Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    }
                    dataGridView4.Rows[i].Cells[0].Value = elemInfo.name;
                    dataGridView4.Rows[i].Cells[1].Value = elemInfo.level;
                    dataGridView4.Rows[i].Cells[2].Value = elemInfo.area;
                    dataGridView4.Rows[i].Cells[3].Value = elemInfo.id;
                    i++;
                }
            }
        }
        /// <summary>
        /// 建立「門尺寸校核」欄位資料
        /// </summary>
        /// <param name="excelCompareList"></param>
        /// <param name="dataGridView"></param>
        /// <param name="doorDatas"></param>
        private void CreateDoorReview(List<ExcelCompare> excelCompareList, DataGridView dataGridView, List<DoorData> doorDatas)
        {
            List<string> charsToRemove = CreateCharsToRemoveTXT(); // 先檢視是否有設定好要移除的特殊符號

            if (dataGridView.Columns.Count == 0)
            {
                this.dataGridView1.AutoGenerateColumns = true;//開始重繪欄位，補足還沒建立的欄位
                DataGridViewTextBoxColumn dgvCol1 = new DataGridViewTextBoxColumn(); // 族群
                DataGridViewTextBoxColumn dgvCol2 = new DataGridViewTextBoxColumn(); // 類型
                DataGridViewTextBoxColumn dgvCol3 = new DataGridViewTextBoxColumn(); // 樓層
                DataGridViewComboBoxColumn dgvCol4 = new DataGridViewComboBoxColumn(); // 歸屬房間
                DataGridViewTextBoxColumn dgvCol5 = new DataGridViewTextBoxColumn(); // 規範門尺寸(mm)
                DataGridViewTextBoxColumn dgvCol6 = new DataGridViewTextBoxColumn(); // 門寬度(mm)
                DataGridViewTextBoxColumn dgvCol7 = new DataGridViewTextBoxColumn(); // 門高度(mm)
                DataGridViewTextBoxColumn dgvCol8 = new DataGridViewTextBoxColumn(); // ID
                //DataGridViewTextBoxColumn dgvCol5 = new DataGridViewTextBoxColumn(); // From Room
                //DataGridViewTextBoxColumn dgvCol6 = new DataGridViewTextBoxColumn(); // To Room
                dgvCol1.Name = "dgvCol1";
                dgvCol2.Name = "dgvCol2";
                dgvCol3.Name = "dgvCol3";
                dgvCol4.Name = "dgvCol4";
                dgvCol5.Name = "dgvCol5";
                dgvCol6.Name = "dgvCol6";
                dgvCol7.Name = "dgvCol7";
                dgvCol8.Name = "dgvCol8";
                dgvCol1.HeaderText = "族群";
                dgvCol2.HeaderText = "類型";
                dgvCol3.HeaderText = "樓層";
                dgvCol4.HeaderText = "歸屬房間";
                dgvCol5.HeaderText = "規範門尺寸(mm)";
                dgvCol6.HeaderText = "門寬度(mm)";
                dgvCol7.HeaderText = "門高度(mm)";
                dgvCol8.HeaderText = "ID";

                dgvCol4.MaxDropDownItems = 5;
                dgvCol4.Visible = true;
                dgvCol4.ReadOnly = false;

                dataGridView.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                dataGridView.Columns.AddRange(new DataGridViewColumn[] { dgvCol1, dgvCol2, dgvCol3, dgvCol4, dgvCol5, dgvCol6, dgvCol7, dgvCol8 });
            }
            dataGridView.Dock = DockStyle.Fill;
            //是否允許使用者編輯
            dataGridView.ReadOnly = false;
            //是否允許使用者自行新增
            dataGridView.AllowUserToAddRows = false;
            //// 依內容自動調整欄寬
            //dataGridView.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells;
            //自動調欄高
            dataGridView.Rows.Clear(); // 清除所有Rows
            int i = 0; // n列
            //int maxWidth = 0; // 最大欄寬

            foreach (DoorData doorData in doorDatas)
            {
                try
                {
                    // 如果門的FromRoom&ToRoom其一需要校核, 則顯示
                    string editFromRoomName = string.Empty;
                    if (doorData.fromRoom != null)
                    {
                        editFromRoomName = doorData.fromRoom.get_Parameter(BuiltInParameter.ROOM_NAME).AsString();
                    }
                    string editToRoomName = string.Empty;
                    if (doorData.toRoom != null)
                    {
                        editToRoomName = doorData.toRoom.get_Parameter(BuiltInParameter.ROOM_NAME).AsString();
                    }
                    foreach (string c in charsToRemove)
                    {
                        try
                        {
                            editFromRoomName = editFromRoomName.Replace(c, string.Empty);
                        }
                        catch (Exception ex)
                        {
                            string error = ex.Message + "\n" + ex.ToString();
                        }
                        try
                        {
                            editToRoomName = editToRoomName.Replace(c, string.Empty);
                        }
                        catch (Exception ex)
                        {
                            string error = ex.Message + "\n" + ex.ToString();
                        }
                    }
                    // 名稱無完全一樣的話, 比對是否有涵蓋的
                    ExcelCompare fromRoomResult = null;
                    if (editFromRoomName != "")
                    {
                        fromRoomResult = (from x in excelCompareList
                                          where x.name.Equals(editFromRoomName) && x.doorWidth != 0 && x.doorHeight != 0
                                          select x).FirstOrDefault();
                        // 走道 != 安全走道
                        if (fromRoomResult == null && !editFromRoomName.Equals("走道"))
                        {
                            fromRoomResult = (from x in excelCompareList
                                              where x.name.Contains(editFromRoomName) && x.doorWidth != 0 && x.doorHeight != 0
                                              select x).FirstOrDefault();
                        }
                    }
                    ExcelCompare toRoomResult = null;
                    if (editToRoomName != "")
                    {
                        toRoomResult = (from x in excelCompareList
                                        where x.name.Equals(editToRoomName) && x.doorWidth != 0 && x.doorHeight != 0
                                        select x).FirstOrDefault();
                        if (toRoomResult == null && !editToRoomName.Equals("走道"))
                        {
                            toRoomResult = (from x in excelCompareList
                                            where x.name.Contains(editToRoomName) && x.doorWidth != 0 && x.doorHeight != 0
                                            select x).FirstOrDefault();
                        }
                    }

                    // 男女廁的公共、員工區, 特殊情況處理
                    if (fromRoomResult != null)
                    {
                        if (doorData.door.FromRoom.get_Parameter(BuiltInParameter.ROOM_NAME).AsString().Contains("男廁") || doorData.door.FromRoom.get_Parameter(BuiltInParameter.ROOM_NAME).AsString().Contains("女廁"))
                        {
                            if (doorData.fromRoom.LookupParameter("房間英文名稱").AsString().Contains("STAFF"))
                            {
                                fromRoomResult = (from x in excelCompareList
                                                  where x.name.Contains(doorData.fromRoom.get_Parameter(BuiltInParameter.ROOM_NAME).AsString()) && x.engName.Contains("STAFF") && x.doorWidth != 0 && x.doorHeight != 0
                                                  select x).FirstOrDefault();
                            }
                            else
                            {
                                fromRoomResult = (from x in excelCompareList
                                                  where x.name.Contains(doorData.fromRoom.get_Parameter(BuiltInParameter.ROOM_NAME).AsString()) && !x.engName.Contains("STAFF") && x.doorWidth != 0 && x.doorHeight != 0
                                                  select x).FirstOrDefault();
                            }
                        }
                    }
                    if (toRoomResult != null)
                    {
                        if (doorData.door.ToRoom.get_Parameter(BuiltInParameter.ROOM_NAME).AsString().Contains("男廁") || doorData.door.ToRoom.get_Parameter(BuiltInParameter.ROOM_NAME).AsString().Contains("女廁"))
                        {
                            if (doorData.toRoom.LookupParameter("房間英文名稱").AsString().Contains("STAFF"))
                            {
                                toRoomResult = (from x in excelCompareList
                                                where x.name.Contains(doorData.toRoom.get_Parameter(BuiltInParameter.ROOM_NAME).AsString()) && x.engName.Contains("STAFF") && x.doorWidth != 0 && x.doorHeight != 0
                                                select x).FirstOrDefault();
                            }
                            else
                            {
                                toRoomResult = (from x in excelCompareList
                                                where x.name.Contains(doorData.toRoom.get_Parameter(BuiltInParameter.ROOM_NAME).AsString()) && !x.engName.Contains("STAFF") && x.doorWidth != 0 && x.doorHeight != 0
                                                select x).FirstOrDefault();
                            }
                        }
                    }
                    if (fromRoomResult != null || toRoomResult != null)
                    {
                        try
                        {
                            dataGridView.Rows.Add();
                            for (int cellCount = 0; cellCount <= 7; cellCount++)
                            {
                                dataGridView.Rows[i].Cells[cellCount].Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                                //int widthCol = dataGridView.Columns[cellCount].Width;
                                //if (widthCol > maxWidth)
                                //{
                                //    maxWidth = widthCol;
                                //}
                                //dataGridView.Columns[cellCount].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
                                //dataGridView.Columns[cellCount].Width = maxWidth;
                            }
                            dataGridView.Rows[i].Cells[0].Value = doorData.door.Symbol.FamilyName; // 族群
                            dataGridView.Rows[i].Cells[1].Value = doorData.door.Symbol.Name; // 類型
                            dataGridView.Rows[i].Cells[2].Value = doorData.level.Name; // 樓層
                            dataGridView.Rows[i].Cells[5].Value = doorData.width; // 門寬度(mm)
                            dataGridView.Rows[i].Cells[6].Value = doorData.height; // 門高度(mm)
                            dataGridView.Rows[i].Cells[7].Value = doorData.door.Id; // ID
                            dataGridView.Rows[i].Cells[7].Style.Alignment = DataGridViewContentAlignment.MiddleCenter;

                            DataGridViewComboBoxCell dgvCol4Cell = new DataGridViewComboBoxCell();
                            dgvCol4Cell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                            dgvCol4Cell.DisplayStyle = DataGridViewComboBoxDisplayStyle.DropDownButton;
                            string colCount = "dgvCol4";

                            // 規範門尺寸(mm)
                            string doorArea = string.Empty;
                            if (fromRoomResult != null && toRoomResult == null)
                            {
                                //將ComboBox下拉式選單的內容儲存到自訂類別中
                                dgvCol4Cell.Items.Add(doorData.fromRoom.Name);
                                dgvCol4Cell.Items.Add("*");
                                if (doorData.belong.Equals(doorData.fromRoom.Name) || doorData.belong.Equals("*"))
                                {
                                    dgvCol4Cell.Value = doorData.belong;
                                    dataGridView[colCount, i] = dgvCol4Cell;
                                }
                                else
                                {
                                    dgvCol4Cell.Value = doorData.fromRoom.Name;
                                    dataGridView[colCount, i] = dgvCol4Cell;
                                }
                                doorArea = fromRoomResult.name;
                            }
                            else if (fromRoomResult == null && toRoomResult != null)
                            {
                                //將ComboBox下拉式選單的內容儲存到自訂類別中
                                dgvCol4Cell.Items.Add(doorData.toRoom.Name);
                                dgvCol4Cell.Items.Add("*");
                                if (doorData.belong.Equals(doorData.toRoom.Name) || doorData.belong.Equals("*"))
                                {
                                    dgvCol4Cell.Value = doorData.belong;
                                    dataGridView[colCount, i] = dgvCol4Cell;
                                }
                                else
                                {
                                    dgvCol4Cell.Value = doorData.toRoom.Name;
                                    dataGridView[colCount, i] = dgvCol4Cell;
                                }
                                doorArea = toRoomResult.name;
                            }
                            else
                            {
                                try
                                {
                                    if (fromRoomResult != null && toRoomResult != null)
                                    {
                                        //將ComboBox下拉式選單的內容儲存到自訂類別中
                                        if (doorData.fromRoom.Name.Equals(doorData.toRoom.Name))
                                        {
                                            dgvCol4Cell.Items.Add(doorData.fromRoom.Name);
                                            dgvCol4Cell.Items.Add("*");
                                        }
                                        else
                                        {
                                            dgvCol4Cell.Items.Add(doorData.fromRoom.Name);
                                            dgvCol4Cell.Items.Add(doorData.toRoom.Name);
                                            dgvCol4Cell.Items.Add("*");
                                        }

                                        // 比對歸屬順序
                                        bool fromRoomName = false;
                                        bool toRoomName = false;
                                        if (doorData.fromRoom.Name.Contains("安全走道"))
                                            fromRoomName = false;
                                        else if (doorData.fromRoom.Name.Contains("走道") || doorData.fromRoom.Name.Contains("通道") || doorData.fromRoom.Name.Contains("月台") ||
                                                 doorData.fromRoom.Name.Contains("付費區") || doorData.fromRoom.Name.Contains("未付費區") || doorData.fromRoom.Name.Contains("非付費區"))
                                            fromRoomName = true;
                                        if (doorData.toRoom.Name.Contains("安全走道"))
                                            toRoomName = false;
                                        else if (doorData.toRoom.Name.Contains("走道") || doorData.toRoom.Name.Contains("通道") || doorData.toRoom.Name.Contains("月台") ||
                                                 doorData.toRoom.Name.Contains("付費區") || doorData.toRoom.Name.Contains("未付費區") || doorData.toRoom.Name.Contains("非付費區"))
                                            toRoomName = true;

                                        if (doorData.fromRoom.Name.Equals(doorData.toRoom.Name))
                                        {
                                            if (doorData.belong.Equals(doorData.fromRoom.Name) || doorData.belong.Equals(doorData.toRoom.Name))
                                            {
                                                dgvCol4Cell.Value = doorData.belong;
                                                dataGridView[colCount, i] = dgvCol4Cell;
                                            }
                                            else
                                            {
                                                dgvCol4Cell.Value = doorData.fromRoom.Name;
                                                dataGridView[colCount, i] = dgvCol4Cell;
                                            }
                                            doorArea = fromRoomResult.name;
                                        }
                                        else if (fromRoomName == true && toRoomName == false)
                                        {
                                            if (doorData.belong.Equals(doorData.fromRoom.Name) || doorData.belong.Equals(doorData.toRoom.Name) || doorData.belong.Equals("*"))
                                            {
                                                dgvCol4Cell.Value = doorData.belong;
                                                dataGridView[colCount, i] = dgvCol4Cell;
                                            }
                                            else
                                            {
                                                dgvCol4Cell.Value = doorData.toRoom.Name;
                                                dataGridView[colCount, i] = dgvCol4Cell;
                                            }
                                            doorArea = toRoomResult.name;
                                        }
                                        else if (fromRoomName == false && toRoomName == true)
                                        {
                                            if (doorData.belong.Equals(doorData.fromRoom.Name) || doorData.belong.Equals(doorData.toRoom.Name) || doorData.belong.Equals("*"))
                                            {
                                                dgvCol4Cell.Value = doorData.belong;
                                                dataGridView[colCount, i] = dgvCol4Cell;
                                            }
                                            else
                                            {
                                                dgvCol4Cell.Value = doorData.fromRoom.Name;
                                                dataGridView[colCount, i] = dgvCol4Cell;
                                            }
                                            doorArea = fromRoomResult.name;
                                        }
                                        else
                                        {
                                            if (doorData.belong.Equals(doorData.fromRoom.Name))
                                            {
                                                dgvCol4Cell.Value = doorData.belong;
                                                dataGridView[colCount, i] = dgvCol4Cell;
                                                doorArea = fromRoomResult.name;
                                            }
                                            else if (doorData.belong.Equals(doorData.toRoom.Name))
                                            {
                                                dgvCol4Cell.Value = doorData.belong;
                                                dataGridView[colCount, i] = dgvCol4Cell;
                                                doorArea = toRoomResult.name;
                                            }
                                            else
                                            {
                                                dgvCol4Cell.Value = "*";
                                                dataGridView[colCount, i] = dgvCol4Cell;
                                                doorArea = "*";
                                            }
                                        }
                                    }
                                }
                                catch (Exception ex)
                                {
                                    string msg = doorData.id + "\n" + ex.Message + "\n" + ex.ToString();
                                }
                            }

                            try
                            {
                                // 規範門尺寸(mm)
                                ExcelCompare result = (from x in excelCompareList
                                                       where x.name.Contains(doorArea) && x.doorWidth != 0 && x.doorHeight != 0
                                                       select x).FirstOrDefault();
                                if (result == null)
                                {
                                    dataGridView.Rows[i].Cells[4].Value = "NA";
                                }
                                else
                                {
                                    // 符合標準
                                    if (doorData.width.Equals(result.doorWidth) && doorData.height.Equals(result.doorHeight))
                                    {
                                        dataGridView.Rows[i].Cells[4].Value = "OK";
                                    }
                                    // NG(紅字)
                                    else if (doorData.width < result.doorWidth || doorData.height < result.doorHeight)
                                    {
                                        dataGridView.Rows[i].Cells[4].Value = result.doorWidth + " x " + result.doorHeight;
                                        for (int j = 4; j <= 6; j++)
                                        {
                                            dataGridView.Rows[i].Cells[j].Style.ForeColor = System.Drawing.Color.Red;
                                        }
                                        dataGridView.Rows[i].DefaultCellStyle.BackColor = System.Drawing.Color.LightYellow;
                                    }
                                    // 不合標準
                                    else
                                    {
                                        dataGridView.Rows[i].Cells[4].Value = result.doorWidth + " x " + result.doorHeight;
                                        for (int j = 4; j <= 6; j++)
                                        {
                                            dataGridView.Rows[i].Cells[j].Style.ForeColor = System.Drawing.Color.Blue;
                                        }
                                        dataGridView.Rows[i].DefaultCellStyle.BackColor = System.Drawing.Color.LightYellow;
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                string error = ex.Message + "\n" + ex.ToString();
                            }

                            i++;
                        }
                        catch (Exception ex)
                        {
                            string error = ex.Message + "\n" + ex.ToString();
                        }
                    }
                }
                catch(Exception ex)
                {
                    string msg = doorData.door.Id + "\n" + ex.Message + "\n" + ex.ToString();
                }
            }
        }
        /// <summary>
        /// 天花內淨高校核
        /// </summary>
        /// <param name="dataGridView"></param>
        /// <param name="sortElemInfoList"></param>
        private void CreateMechanicalReview(DataGridView dataGridView, List<ElementInfo> sortElemInfoList)
        {
            if (dataGridView.Columns.Count == 0)
            {
                DataGridViewTextBoxColumn dgvCol1 = new DataGridViewTextBoxColumn();
                DataGridViewTextBoxColumn dgvCol2 = new DataGridViewTextBoxColumn();
                DataGridViewTextBoxColumn dgvCol3 = new DataGridViewTextBoxColumn();
                DataGridViewTextBoxColumn dgvCol4 = new DataGridViewTextBoxColumn();
                DataGridViewTextBoxColumn dgvCol5 = new DataGridViewTextBoxColumn();
                DataGridViewTextBoxColumn dgvCol6 = new DataGridViewTextBoxColumn();
                dgvCol1.Name = "dgvCol1";
                dgvCol1.HeaderText = "名稱";
                dgvCol2.Name = "dgvCol2";
                dgvCol2.HeaderText = "樓層";
                dgvCol3.Name = "dgvCol3";
                dgvCol3.HeaderText = "天花板";
                dgvCol4.Name = "dgvCol4";
                dgvCol4.HeaderText = "天花內淨高(m)";
                dgvCol5.Name = "dgvCol5";
                dgvCol5.HeaderText = "淨高(m)";
                dgvCol6.Name = "dgvCol6";
                dgvCol6.HeaderText = "ID";
                dataGridView.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                dataGridView.Columns.AddRange(new DataGridViewColumn[] { dgvCol1, dgvCol2, dgvCol3, dgvCol4, dgvCol5, dgvCol6 });
            }
            dataGridView.Dock = DockStyle.Fill;
            //是否允許使用者編輯
            dataGridView.ReadOnly = true;
            //是否允許使用者自行新增
            dataGridView.AllowUserToAddRows = false;
            dataGridView.Rows.Clear(); // 清除所有Rows

            int i = 0; // n列
            foreach (ElementInfo elemInfo in sortElemInfoList)
            {
                dataGridView.Rows.Add();
                for (int cellCount = 0; cellCount <= 5; cellCount++)
                {
                    dataGridView.Rows[i].Cells[cellCount].Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                }
                dataGridView.Rows[i].Cells[0].Value = elemInfo.name; // 名稱
                dataGridView.Rows[i].Cells[1].Value = elemInfo.level; // 樓層
                dataGridView.Rows[i].Cells[2].Value = ""; // 天花板
                dataGridView.Rows[i].Cells[3].Value = 0.0.ToString(); // 天花內淨高
                dataGridView.Rows[i].Cells[4].Value = 0.0.ToString(); // 淨高
                dataGridView.Rows[i].Cells[5].Value = elemInfo.id; // id
                i++;
            }
        }
        /// <summary>
        /// 修改房間高度
        /// </summary>
        private void EditRoomOffset()
        {
            m_externalEvent_EditRoomOffset.Raise();
        }
        private void offsetBtn_Click(object sender, EventArgs e)
        {
            // 修改房間高度
            EditRoomOffset();
        }
        /// <summary>
        /// 關閉
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void closeBtn_Click(object sender, EventArgs e)
        {
            try
            {
                revitUIApp.Idling -= IdleUpdate;
            }
            catch (Exception)
            {

            }
            Close();
        }
        /// <summary>
        /// 讀取要亮顯元素的id
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                DataGridView dgv = sender as DataGridView;
                int rowIndex = e.RowIndex;
                if(rowIndex != -1)
                {
                    int columnIndex = e.ColumnIndex;
                    string rowValue = "";
                    if (!dgv.Name.Equals("dataGridView3"))
                    {
                        rowValue = dgv.Rows[rowIndex].Cells[dgv.ColumnCount - 1].Value.ToString(); // ID
                        if (dgv.Name.Equals("dataGridView5"))
                        {
                            if (!columnIndex.Equals(3))
                            {
                                ElementId id = new ElementId(Convert.ToInt32(rowValue));
                                // 亮顯id元件的FromRoom、ToRoom
                                try
                                {
                                    FamilyInstance door = revitDoc.GetElement(id) as FamilyInstance;
                                    
                                    IList<ElementId> highlightElems = new List<ElementId>();
                                    highlightElems.Add(id);
                                    if(door.FromRoom != null)
                                    {
                                        highlightElems.Add(door.FromRoom.Id);
                                    }
                                    if (door.ToRoom != null)
                                    {
                                        highlightElems.Add(door.ToRoom.Id);
                                    }
                                    revitUIApp.ActiveUIDocument.ShowElements(highlightElems);
                                    revitUIApp.ActiveUIDocument.Selection.SetElementIds(highlightElems);
                                }
                                catch(Exception)
                                {

                                }
                            }
                        }
                        else
                        {
                            ElementId id = new ElementId(Convert.ToInt32(rowValue));
                            // 亮顯id元件
                            IList<ElementId> highlightElems = new List<ElementId>();
                            highlightElems.Add(id);
                            revitUIApp.ActiveUIDocument.ShowElements(id);
                            revitUIApp.ActiveUIDocument.Selection.SetElementIds(highlightElems);
                        }
                    }
                }
            }
            // 升降冪排序
            catch (Exception)
            {
                TaskDialog.Show("Error", "此元件位於其他連結模型中。");
            }
        }
        /// <summary>
        /// 新增特殊符號
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void createBtn_Click(object sender, EventArgs e)
        {
            if(createRemoveTB.Text != "")
            {
                string[] signs = label1.Text.Split('　'); // label1顯示的特殊符號
                bool isRepeat = IsRepeat(signs, createRemoveTB.Text); // 檢查label1中是否有重複字元
                if (isRepeat == false)
                {
                    if (createRemoveTB.Text != "　")
                    {
                        label1.Text += createRemoveTB.Text + "　";
                        UpdateCharsToRemoveTXT(); // 更新特殊符號文字檔
                    }
                }
                createRemoveTB.Text = "";
            }
        }
        /// <summary>
        /// 設定
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void removeBtn_Click(object sender, EventArgs e)
        {
            if (createRemoveTB.Text != "")
            {
                string[] signs = label1.Text.Split('　'); // label1顯示的特殊符號
                bool isRepeat = IsRepeat(signs, createRemoveTB.Text); // 檢查label1中是否有重複字元
                if (isRepeat == true)
                {
                    label1.Text = label1.Text.Replace(createRemoveTB.Text + "　", "");
                    UpdateCharsToRemoveTXT(); // 更新特殊符號文字檔
                }
                createRemoveTB.Text = "";
            }
        }
        /// <summary>
        /// 檢查label1中是否有重複字元
        /// </summary>
        /// <param name="array"></param>
        /// <param name="checkWord"></param>
        /// <returns></returns>
        public static bool IsRepeat(string[] array, string checkWord)
        {
            for (int i = 0; i < array.Length; i++)
            {
                if (array[i].Equals(checkWord))
                {
                    return true;
                }
            }
            return false;
        }
        /// <summary>
        /// 更新特殊符號文字檔
        /// </summary>
        public void UpdateCharsToRemoveTXT()
        {
            string charsToRemovePath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\CharsToRemove.txt"; // 取得使用者文件路徑
            // 先檢查是否有此檔案, 沒有的話則新增
            List<string> charsToRemove = new List<string>();
            if (!File.Exists(charsToRemovePath))
            {
                string[] signs = label1.Text.Split('　'); // label1顯示的特殊符號
                foreach (string sign in signs)
                {
                    charsToRemove.Add(sign);
                }
                using (StreamWriter outputFile = new StreamWriter(charsToRemovePath))
                {
                    foreach (string sign in charsToRemove)
                    {
                        outputFile.WriteLine(sign);
                    }
                }
            }
            else
            {
                charsToRemove = new List<string>();
                string[] signs = label1.Text.Split('　'); // label1顯示的特殊符號
                foreach (string sign in signs)
                {
                    charsToRemove.Add(sign);
                }
                charsToRemove = charsToRemove.Distinct().ToList(); // 移除重複
                File.WriteAllLines(charsToRemovePath, charsToRemove);
            }
        }
        /// <summary>
        /// 干涉檢查淨高與空間淨高
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void IntersectBtn_Click(object sender, EventArgs e)
        {
            if (tabControl1.SelectedIndex.Equals(0))
            {
                m_externalEvent_CreateCrush.Raise();
            }
            else
            {
                try
                {
                    List<ElementInfo> elemInfoList = new List<ElementInfo>();
                    // 讀取DataGridView內的ID, 透過ID取得Area
                    List<int> areaIds = new List<int>();
                    for (int rows = 0; rows < dataGridView1.Rows.Count; rows++)
                    {
                        areaIds.Add(Convert.ToInt32(dataGridView1.Rows[rows].Cells[dataGridView1.ColumnCount - 1].Value.ToString()));
                    }
                    for (int rows = 0; rows < dataGridView2.Rows.Count; rows++)
                    {
                        areaIds.Add(Convert.ToInt32(dataGridView2.Rows[rows].Cells[dataGridView2.ColumnCount - 1].Value.ToString()));
                    }
                    areaIds = areaIds.Distinct().ToList();
                    foreach (int areaId in areaIds)
                    {
                        try
                        {
                            ElementId id = new ElementId((int)areaId);
                            Room room = revitDoc.GetElement(id) as Room;
                            // 1.讀取Geometry Option
                            Options options = new Options();
                            //options.View = doc.GetElement(room.Level.FindAssociatedPlanViewId()) as Autodesk.Revit.DB.View;
                            options.DetailLevel = ViewDetailLevel.Medium;
                            options.ComputeReferences = true;
                            options.IncludeNonVisibleObjects = true;
                            // 得到幾何元素
                            GeometryElement geomElem = room.get_Geometry(options);
                            List<Solid> solids = GeometrySolids(geomElem);

                            ElementInfo elementInfo = new ElementInfo();
                            elementInfo.elem = room;
                            elementInfo.id = room.Id.ToString();
                            elementInfo.solids = solids;
                            // 從Solid查詢干涉到的元件, 找到樓梯、電扶梯
                            foreach (Solid roomSolid in solids)
                            {
                                FindTheIntersectElems(revitDoc, roomSolid, elementInfo);
                            }
                            elemInfoList.Add(elementInfo);
                        }
                        catch (Exception ex)
                        {
                            string error = ex.Message + "\n" + ex.ToString();
                        }
                    }
                    // 搜尋DataGridView的Room ID, 有干涉到樓梯、電扶梯的, 則加上 ⭐
                    for (int rows = 0; rows < dataGridView1.Rows.Count; rows++)
                    {
                        string roomId = dataGridView1.Rows[rows].Cells[dataGridView1.ColumnCount - 1].Value.ToString();
                        // 有干涉到樓梯或電扶梯的元件
                        int intersectCount = elemInfoList.Where(x => x.id.ToString().Equals(roomId)).Select(y => y.intersectElems).FirstOrDefault().Count();
                        if (intersectCount > 0)
                        {
                            // 如果有" ⭐"先移除
                            string unboundedHeight = dataGridView1.Rows[rows].Cells[dataGridView1.ColumnCount - 2].Value.ToString().Replace(" ⭐", "");
                            dataGridView1.Rows[rows].Cells[dataGridView1.ColumnCount - 2].Value = unboundedHeight + " ⭐"; // 實際淨高
                        }
                    }
                    for (int rows = 0; rows < dataGridView2.Rows.Count; rows++)
                    {
                        string roomId = dataGridView2.Rows[rows].Cells[dataGridView2.ColumnCount - 1].Value.ToString();
                        // 有干涉到樓梯或電扶梯的元件
                        int intersectCount = elemInfoList.Where(x => x.id.ToString().Equals(roomId)).Select(y => y.intersectElems).FirstOrDefault().Count();
                        if (intersectCount > 0)
                        {
                            // 如果有" ⭐"先移除
                            string unboundedHeight = dataGridView2.Rows[rows].Cells[dataGridView2.ColumnCount - 2].Value.ToString().Replace(" ⭐", "");
                            dataGridView2.Rows[rows].Cells[dataGridView2.ColumnCount - 2].Value = unboundedHeight + " ⭐"; // 實際淨高
                        }
                    }
                }
                catch (Exception ex)
                {
                    string error = ex.Message + "\n" + ex.ToString();
                }
            }
        }
        /// <summary>
        /// 更新歸屬房間
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void updateDoorBtn_Click(object sender, EventArgs e)
        {
            doorDataList = new List<DoorData>();
            foreach (DataGridViewRow row in dataGridView5.Rows)
            {
                try
                {
                    DoorData doorData = new DoorData();
                    doorData.id = new ElementId(Convert.ToInt32(row.Cells[dataGridView5.ColumnCount - 1].Value.ToString())); // ID
                    doorData.roomName = row.Cells[3].Value.ToString(); // 規範門尺寸(mm)
                    doorDataList.Add(doorData);
                }
                catch (Exception)
                {

                }
            }
            m_externalEvent_EditRoomPara.Raise();
        }
        /// <summary>
        /// 切換page, 顯示、關閉提醒文字
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void tabControl1_Click(object sender, EventArgs e)
        {
            if (tabControl1.SelectedIndex == 0 || tabControl1.SelectedIndex == 1)
            {
                label6.Hide();
                label7.Hide();
                label3.Show();
                label4.Show();
                label5.Show();
                label9.Show();
                label10.Show();
                pictureBox1.Show();
                pictureBox2.Show();
                pictureBox3.Show();
                updateBtn.Show();
                offsetBtn.Show();
                IntersectBtn.Hide();
                updateDoorBtn.Hide();
            }
            else if (tabControl1.SelectedIndex == 2)
            {
                label6.Show();
                label7.Hide();
                label9.Hide();
                label10.Hide();
                updateBtn.Hide();
                offsetBtn.Hide();
                IntersectBtn.Hide();
                updateDoorBtn.Hide();
            }
            else if (tabControl1.SelectedIndex == 3)
            {
                label7.Show();
                label6.Hide();
                label9.Hide();
                label10.Hide();
                updateBtn.Hide();
                offsetBtn.Hide();
                IntersectBtn.Hide();
                updateDoorBtn.Hide();
            }
            else if (tabControl1.SelectedIndex == 4)
            {
                label6.Hide();
                label7.Hide();
                label3.Show();
                label4.Show();
                label5.Show();
                label9.Hide();
                label10.Hide();
                pictureBox1.Show();
                pictureBox2.Show();
                pictureBox3.Show();
                updateBtn.Hide();
                offsetBtn.Hide();
                IntersectBtn.Hide();
                updateDoorBtn.Show();
            }
            else if (tabControl1.SelectedIndex == 5)
            {
                label6.Hide();
                label7.Hide();
                label3.Hide();
                label4.Hide();
                label5.Hide();
                label9.Show();
                label10.Show();
                pictureBox1.Hide();
                pictureBox2.Hide();
                pictureBox3.Hide();
                updateBtn.Hide();
                offsetBtn.Hide();
                IntersectBtn.Hide();
                updateDoorBtn.Hide();
            }
        }
        /// <summary>
        /// 選擇房間時同步顯示datagridview的欄位
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public void IdleUpdate(object sender, IdlingEventArgs e)
        {
            UIApplication uiapp = sender as UIApplication;
            ICollection<ElementId> elems = revitUIDoc.Selection.GetElementIds();
            if (elems.Count > 0)
            {
                Element elem = revitDoc.GetElement(elems.FirstOrDefault());
                if (elem is Room)
                {
                    Room room = (Room)elem;
                    if (room.Id.IntegerValue != chooseRoomId)
                    {
                        // 找到dataGridView1 Room id的列
                        foreach (DataGridViewRow dgv1Row in dataGridView1.Rows)
                        {
                            if (Convert.ToInt64(dgv1Row.Cells[dgv1Row.Cells.Count - 1].Value.ToString()).Equals(room.Id.IntegerValue))
                            {
                                dataGridView1.Rows[dgv1Row.Index].Selected = true;
                                dataGridView1.CurrentCell = dataGridView1.Rows[dgv1Row.Index].Cells[0];
                                dataGridView1.FirstDisplayedScrollingRowIndex = dgv1Row.Index;
                                dgv1RowIndex = dgv1Row.Index;
                                break;
                            }
                            else
                            {
                                dataGridView1.Rows[dgv1RowIndex].Selected = false;
                            }
                        }
                        // 找到dataGridView2 Room id的列
                        foreach (DataGridViewRow dgv2Row in dataGridView2.Rows)
                        {
                            if (Convert.ToInt64(dgv2Row.Cells[dgv2Row.Cells.Count - 1].Value.ToString()).Equals(room.Id.IntegerValue))
                            {
                                dataGridView2.Rows[dgv2Row.Index].Selected = true;
                                dataGridView2.CurrentCell = dataGridView2.Rows[dgv2Row.Index].Cells[0];
                                dataGridView2.FirstDisplayedScrollingRowIndex = dgv2Row.Index;
                                dgv2RowIndex = dgv2Row.Index;
                                break;
                            }
                            else
                            {
                                dataGridView2.Rows[dgv2RowIndex].Selected = false;
                            }
                        }
                        // 找到dataGridView6 Room id的列
                        foreach (DataGridViewRow dgv6Row in dataGridView6.Rows)
                        {
                            if (Convert.ToInt64(dgv6Row.Cells[dgv6Row.Cells.Count - 1].Value.ToString()).Equals(room.Id.IntegerValue))
                            {
                                dataGridView6.Rows[dgv6Row.Index].Selected = true;
                                dataGridView6.CurrentCell = dataGridView6.Rows[dgv6Row.Index].Cells[0];
                                dataGridView6.FirstDisplayedScrollingRowIndex = dgv6Row.Index;
                                dgv6RowIndex = dgv6Row.Index;
                                break;
                            }
                            else
                            {
                                dataGridView6.Rows[dgv6RowIndex].Selected = false;
                            }
                        }
                        chooseRoomId = room.Id.IntegerValue;
                    }
                }
                // 沒有選擇房間則取消顯示
                else
                {
                    dataGridView1.Rows[dgv1RowIndex].Selected = false;
                    dataGridView2.Rows[dgv2RowIndex].Selected = false;
                    dataGridView6.Rows[dgv6RowIndex].Selected = false;
                    chooseRoomId = 0;
                }
            }
            // 沒有選擇房間則取消顯示
            else
            {
                dataGridView1.Rows[dgv1RowIndex].Selected = false;
                dataGridView2.Rows[dgv2RowIndex].Selected = false;
                dataGridView6.Rows[dgv6RowIndex].Selected = false;
                chooseRoomId = 0;
            }
        }
        /// <summary>
        /// 「房間規範檢討」依ID、樓層排序
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dataGridView1_ColumnHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            DataGridView dataGridView = sender as DataGridView;
            string dgvName = dataGridView.Name;
            int columnIndex = e.ColumnIndex;
            if (columnIndex.Equals(0)) // 代碼
            {
                if (sortOrder == SortOrder.Ascending) { sortOrder = SortOrder.Descending; }
                else { sortOrder = SortOrder.Ascending; }
                dataGridView1.Sort(new CustomComparer(dgvName, columnIndex, sortOrder));
            }
            else if (columnIndex.Equals(2)) // 樓層
            {
                if (sortOrder == SortOrder.Ascending) { sortOrder = SortOrder.Descending; }
                else { sortOrder = SortOrder.Ascending; }
                dataGridView1.Sort(new CustomComparer(dgvName, columnIndex, sortOrder));
            }
        }
        /// <summary>
        /// 「房間需求檢討」依ID、樓層排序
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dataGridView2_ColumnHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            DataGridView dataGridView = sender as DataGridView;
            string dgvName = dataGridView.Name;
            int columnIndex = e.ColumnIndex;
            if (columnIndex.Equals(0)) // 代碼
            {
                if (sortOrder == SortOrder.Ascending) { sortOrder = SortOrder.Descending; }
                else { sortOrder = SortOrder.Ascending; }
                dataGridView2.Sort(new CustomComparer(dgvName, columnIndex, sortOrder));
            }
            else if (columnIndex.Equals(2)) // 樓層
            {
                if (sortOrder == SortOrder.Ascending) { sortOrder = SortOrder.Descending; }
                else { sortOrder = SortOrder.Ascending; }
                dataGridView2.Sort(new CustomComparer(dgvName, columnIndex, sortOrder));
            }
        }
        /// <summary>
        /// 「未設置房間」依ID、樓層排序
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dataGridView3_ColumnHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            DataGridView dataGridView = sender as DataGridView;
            string dgvName = dataGridView.Name;
            int columnIndex = e.ColumnIndex;
            if (columnIndex.Equals(0)) // 代碼
            {
                if (sortOrder == SortOrder.Ascending) { sortOrder = SortOrder.Descending; }
                else { sortOrder = SortOrder.Ascending; }
                dataGridView3.Sort(new CustomComparer(dgvName, columnIndex, sortOrder));
            }
            else if (columnIndex.Equals(2)) // 樓層
            {
                if (sortOrder == SortOrder.Ascending) { sortOrder = SortOrder.Descending; }
                else { sortOrder = SortOrder.Ascending; }
                dataGridView3.Sort(new CustomComparer(dgvName, columnIndex, sortOrder));
            }
        }
        /// <summary>
        /// 「未校核房間」依樓層排序
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dataGridView4_ColumnHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            DataGridView dataGridView = sender as DataGridView;
            string dgvName = dataGridView.Name;
            int columnIndex = e.ColumnIndex;
            if (columnIndex.Equals(1)) // 樓層
            {
                if (sortOrder == SortOrder.Ascending) { sortOrder = SortOrder.Descending; }
                else { sortOrder = SortOrder.Ascending; }
                dataGridView4.Sort(new CustomComparer(dgvName, columnIndex, sortOrder));
            }
        }
        /// <summary>
        /// 「門尺寸校核」依樓層排序
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dataGridView5_ColumnHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            DataGridView dataGridView = sender as DataGridView;
            string dgvName = dataGridView.Name;
            int columnIndex = e.ColumnIndex;
            if (columnIndex.Equals(2)) // 樓層
            {
                if (sortOrder == SortOrder.Ascending) { sortOrder = SortOrder.Descending; }
                else { sortOrder = SortOrder.Ascending; }
                dataGridView5.Sort(new CustomComparer(dgvName, columnIndex, sortOrder));
            }
        }
        /// <summary>
        /// 「天花內淨高校核」依樓層排序
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dataGridView6_ColumnHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            DataGridView dataGridView = sender as DataGridView;
            string dgvName = dataGridView.Name;
            int columnIndex = e.ColumnIndex;
            if (columnIndex.Equals(1)) // 樓層
            {
                if (sortOrder == SortOrder.Ascending) { sortOrder = SortOrder.Descending; }
                else { sortOrder = SortOrder.Ascending; }
                dataGridView6.Sort(new CustomComparer(dgvName, columnIndex, sortOrder));
            }
        }
        public class CustomComparer : IComparer
        {
            private string dgvName;
            private int columnIndex;
            private SortOrder sortOrder;

            public CustomComparer(string dgvName, int columnIndex, SortOrder sortOrder)
            {
                this.dgvName = dgvName;
                this.columnIndex = columnIndex;
                this.sortOrder = sortOrder;
            }

            public int Compare(object x, object y)
            {
                DataGridViewRow row1 = x as DataGridViewRow;
                DataGridViewRow row2 = y as DataGridViewRow;
                int row1Sort = 0;
                int row2Sort = 0;
                if (dgvName.Equals("dataGridView4") || dgvName.Equals("dataGridView6"))
                {
                    if (columnIndex.Equals(1)) // 樓層
                    {
                        string row1Value = row1.Cells[columnIndex].Value.ToString();
                        string row2Value = row2.Cells[columnIndex].Value.ToString();
                        LevelElevation row1levelElev = levelElevList.Where(o => o.name.Equals(row1Value)).FirstOrDefault();
                        LevelElevation row2levelElev = levelElevList.Where(o => o.name.Equals(row2Value)).FirstOrDefault();
                        if (String.IsNullOrEmpty(row1Value)) { row1Sort = levelElevList.Count() + 1; }
                        else if (row1levelElev == null) { row1Sort = levelElevList.Count(); }
                        else { row1Sort = levelElevList.Where(o => o.name.Equals(row1.Cells[columnIndex].Value.ToString())).Select(o => o.sort).FirstOrDefault(); }
                        if (String.IsNullOrEmpty(row2Value)) { row2Sort = levelElevList.Count() + 1; }
                        else if (row2levelElev == null) { row2Sort = levelElevList.Count(); }
                        else { row2Sort = levelElevList.Where(o => o.name.Equals(row2.Cells[columnIndex].Value.ToString())).Select(o => o.sort).FirstOrDefault(); }
                    }
                }
                else
                {
                    if (!dgvName.Equals("dataGridView5") && columnIndex.Equals(0)) // 代碼
                    {
                        row1Sort = orderByCode.Where(o => o.code.Equals(row1.Cells[columnIndex].Value.ToString())).Select(o => o.sort).FirstOrDefault();
                        row2Sort = orderByCode.Where(o => o.code.Equals(row2.Cells[columnIndex].Value.ToString())).Select(o => o.sort).FirstOrDefault();
                    }
                    else if (columnIndex.Equals(2)) // 樓層
                    {
                        string row1Value = row1.Cells[columnIndex].Value.ToString();
                        string row2Value = row2.Cells[columnIndex].Value.ToString();
                        LevelElevation row1levelElev = levelElevList.Where(o => o.name.Equals(row1Value)).FirstOrDefault();
                        LevelElevation row2levelElev = levelElevList.Where(o => o.name.Equals(row2Value)).FirstOrDefault();
                        if (String.IsNullOrEmpty(row1Value)) { row1Sort = levelElevList.Count() + 1; }
                        else if (row1levelElev == null) { row1Sort = levelElevList.Count(); }
                        else { row1Sort = levelElevList.Where(o => o.name.Equals(row1.Cells[columnIndex].Value.ToString())).Select(o => o.sort).FirstOrDefault(); }
                        if (String.IsNullOrEmpty(row2Value)) { row2Sort = levelElevList.Count() + 1; }
                        else if (row2levelElev == null) { row2Sort = levelElevList.Count(); }
                        else { row2Sort = levelElevList.Where(o => o.name.Equals(row2.Cells[columnIndex].Value.ToString())).Select(o => o.sort).FirstOrDefault(); }
                    }
                }
                int result = 0;
                if (sortOrder == SortOrder.Descending)
                {
                    if (row1Sort >= row2Sort) { result = 1; }
                    else { result = -1; }
                }
                else
                {
                    if (row1Sort >= row2Sort) { result = -1; }
                    else { result = 1; }
                }

                return result;
            }
        }
    }
}