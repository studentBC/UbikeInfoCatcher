using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Model.StationInfo;
using System.Reflection;
using UbikeInfoCatcher.Model;
using Excel = Microsoft.Office.Interop.Excel;
using System.Timers;
using System.IO;

namespace UbikeInfoCatcher
{
    public partial class Form : System.Windows.Forms.Form
    {
        public Form()
        {
            InitializeComponent();
        }
        //for excel 
        Excel._Workbook wbk;
        Excel._Worksheet wst;
        Excel.Range wrange;
        RetVal db;
        BIKEStation bs;
        public  string savePath = "C:\\Users\\student\\Desktop";
        internal void UIValue(StationInfo Record)
        {
            Record = StationInfo.GetRecord();
            int a = Record.retCode;
            db = Record.retVal;
            savePath = "C:\\Users\\user\\Desktop\\rawdata\\" + db.Station_0001.mday;
            Console.WriteLine(db.Station_0001.mday);
            //board.AppendText("站點:  "+db.Station_0001.sna+"\n"
            //    +"資料更新時間:  "+db.Station_0001.mday+"\n"
            //    +"目前車輛數:  "+db.Station_0001.sbi+"\n"
            //    +"空位數量:  "+db.Station_0001.bemp+"\n"
            //    +"總停車格:  "+db.Station_0001.tot+"\n"
            //    +"*******************"+"\n");
        }
        private void button1_Click(object sender, EventArgs e)
        {
            StationInfo s = new StationInfo();
            UIValue(s);
            Console.Write(savePath);
            
        }

        private void saveTOcsv(object sender, EventArgs e)
        {
            button1_Click(sender, e);
            // 開啟一個新的應用程式
            Excel.Application app = new Excel.Application();
            // 讓Excel文件可見
            app.Visible = true;
            // 停用警告訊息
            app.DisplayAlerts = false;
            // 加入新的活頁簿
            app.Workbooks.Add(Type.Missing);
            // 引用第一個活頁簿
            wbk = app.Workbooks[1];
            // 設定活頁簿焦點
            wbk.Activate();
            try
            {
                // 引用第一個工作表
                wst = (Excel._Worksheet)wbk.Worksheets[1];
                // 命名工作表的名稱
                wst.Name = db.Station_0001.mday;
                // 設定工作表焦點
                wst.Activate();
                app.Cells[1, 1] = "ubike retrieve data";
                // 設定metainfo
                app.Cells[1, 1] = "站點";
                app.Cells[2, 1] = "資料更新時間";
                app.Cells[3, 1] = "目前車輛數";
                app.Cells[4, 1] = "空位數量";
                app.Cells[5, 1] = "總停車格";
                // 設定第1列顏色
                wrange = wst.Range[wst.Cells[1, 1], wst.Cells[5, 1]];
                wrange.Select();
                wrange.Font.Color = ColorTranslator.ToOle(Color.Blue);//word color
                wrange.Interior.Color = ColorTranslator.ToOle(Color.White); //back color
                // 設定第2列資料
                for (int c = 2; c < db.Stations.Count; c++)
                {
                    app.Cells[1, c] = db.Stations[c - 2].sno; //站點
                    app.Cells[2, c] = db.Stations[c - 2].mday; //資料更新時間
                    app.Cells[3, c] = db.Stations[c - 2].sbi;//目前車數
                    app.Cells[4, c] = db.Stations[c - 2].bemp;//空格數
                    app.Cells[5, c] = db.Stations[c - 2].tot;//總停車數
                }
            

                try
                {
                    //另存活頁簿
                    wbk.SaveAs(savePath,Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    wbk.SaveCopyAs(db.Station_0001.mday);
                    Console.WriteLine("儲存文件於 " + Environment.NewLine + savePath);
                }
                catch (Exception ex)
                {
                    Console.WriteLine("儲存檔案出錯，檔案可能正在使用" + Environment.NewLine + ex.Message);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("產生報表時出錯！" + Environment.NewLine + ex.Message);
            }

            //關閉活頁簿
            wbk.Close(false, Type.Missing, Type.Missing);
            //關閉Excel
            app.Quit();
            //釋放Excel資源
            System.Runtime.InteropServices.Marshal.ReleaseComObject(app);
            wbk = null;
            wst = null;
            wrange = null;
            app = null;
            GC.Collect();
        }

        private void saveToExcel(object sender, EventArgs e)
        {
            System.Timers.Timer timer = new System.Timers.Timer(600000000);
            timer.Enabled = true;
            //timer.Interval = 1000*60*60;//1 millisecond , conduct one hour   
            timer.Start();
            timer.Elapsed += new System.Timers.ElapsedEventHandler(saveTOcsv);
            // Have the timer fire repeated events (true is the default)
            timer.AutoReset = false;
            // Start the timer
            timer.Enabled = true;
        }

        private void getDistance(object sender, EventArgs e)
        {
            double d;
            Console.WriteLine("------------------------------");
            //for youbike
            for (int i = 0; i < db.Stations.Count; i++)
            {
                for (int j = 0; j < db.Stations.Count; j++)
                {
                    d = GetDistance.getDistance(Convert.ToDouble(db.Stations.ElementAt(i).lat), Convert.ToDouble(db.Stations.ElementAt(i).lng),
                        Convert.ToDouble(db.Stations.ElementAt(j).lat), Convert.ToDouble(db.Stations.ElementAt(j).lng));
                    if (d == 0 && (i != j))
                    {
                        throw new Exception("need to check");
                    }
                    Console.Write(d + ",");
                }
            }
            //get city bike distance matrix
            //for (int i = 0; i < bs.cbs.Count; i++)
            //{
            //    for (int j = 0; j < bs.cbs.Count; j++)
            //    {
            //        d = GetDistance.getDistance(Convert.ToDouble(bs.cbs.ElementAt(i).StationLat), Convert.ToDouble(bs.cbs.ElementAt(i).StationLon),
            //            Convert.ToDouble(bs.cbs.ElementAt(j).StationLat), Convert.ToDouble(bs.cbs.ElementAt(j).StationLon));
            //        if (d == 0 && (i != j))
            //        {
            //            throw new Exception("need to check");
            //        }
            //        Console.Write(d + ",");
            //    }
            //}
            //列印出所有站點的經緯度
            //for (int j = 0; j < db.Stations.Count; j++)
            //{
            //    //Console.Write(db.Stations.ElementAt(j).lat + "," + db.Stations.ElementAt(j).lng + ",");
            //    board.AppendText(db.Stations.ElementAt(j).lat + "," + db.Stations.ElementAt(j).lng + ",");
            //}
        }

        private void getStartToVertex(object sender, EventArgs e)
        {
            double d;
            Console.WriteLine("------------start to vertex------------------");
            //for (int i = 0; i < db.Stations.Count; i++)
            //{
            //        d = GetDistance.getDistance(Convert.ToDouble(db.Stations.ElementAt(i).lat), Convert.ToDouble(db.Stations.ElementAt(i).lng),
            //            25.049755, 121.58063199999992);

            //    Console.Write(d + ",");
            //    //if (d == 0)
            //    //{
            //    //    throw new Exception("need to check");
            //    //}
            //}
            //for city bike
            for (int i = 0; i < bs.cbs.Count; i++)
            {
                d = GetDistance.getDistance(Convert.ToDouble(bs.cbs.ElementAt(i).StationLat), Convert.ToDouble(bs.cbs.ElementAt(i).StationLon),
                    22.651777, 120.337010);

                Console.Write(d + ",");
                //if (d == 0)
                //{
                //    throw new Exception("need to check");
                //}
            }
        }

        private void open(object sender, EventArgs e)
        {
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                StreamReader sr = new StreamReader(openFileDialog.FileName);

                string[] capacity =  sr.ReadLine().Split(' ');
                int[] stations = new int[capacity.Length];
                for (int i = 0; i < capacity.Length; i++)
                {
                    stations[i] = Convert.ToInt32( capacity[i]);
                }
                Array.Sort(stations);
                for (int j = 0; j < capacity.Length; j++)
                {
                    Console.Write((stations[j]-j ) + " ");
                }
            }
        }

        private void toolStripButton2_Click(object sender, EventArgs e)
        {
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook excelWB ;
            Excel.Worksheet excelWS;
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                StreamReader sr = new StreamReader(openFileDialog.FileName);
                int[] count = new int[256];
                string[] Number = sr.ReadLine().Split(',');
                for(int i =0;i < Number.Length; i++)
                {
                    count[Convert.ToInt32( Number[i])]++;
                }
                for(int a = 0;a < 256; a++)
                {
                    Console.Write(count[a] + " ");
                }
                SaveFileDialog save = new SaveFileDialog();

             
                //Excel.Range oRng;
                saveFileDialog.Filter = "Excel活頁簿(*.xlsx)|*.xlsx";
                if (save.ShowDialog() == DialogResult.OK)
                {
                    string fileName = save.FileName;
                    excelApp.Workbooks.Add();
                    excelWB = excelApp.Workbooks[1];
                    excelWS = excelWB.Worksheets[1];

                    for (int i = 0; i < 256; i++)
                    {
                        excelWS.Cells[2, 2+i] = i;
                        excelWS.Cells[3, 2+i] = count[i];
                    }
                    excelWB.SaveAs(fileName);
                    excelWB.Close();
                }
            }

            excelApp.Quit();
            //釋放Excel資源
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
            GC.Collect();
        }

        private void xml2cs(object sender, EventArgs e)
        {
            CityBikeInfo cs = new CityBikeInfo();
            string json = "";
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                StreamReader sr = new StreamReader(openFileDialog.FileName);
                 json = sr.ReadToEnd();               
            }
            //for (int i = 0; i < 206; i++)
            //{
            //    Console.Write("items.BIKEStation.cbs.Add(items.BIKEStation.CS"+(i+1)+");\n");
            //}
            //string apart = "},";
            //json = json.Replace(apart, apart + "\"CS\":");
            cs  = CityBikeInfo.ReadJson(json);
            bs = cs.BIKEStation;
        }

        private void XMLsave2excel(object sender, EventArgs e)
        {
            // 開啟一個新的應用程式
            Excel.Application app = new Excel.Application();
            // 讓Excel文件可見
            app.Visible = true;
            // 停用警告訊息
            app.DisplayAlerts = false;
            // 加入新的活頁簿
            app.Workbooks.Add(Type.Missing);
            // 引用第一個活頁簿
            wbk = app.Workbooks[1];
            // 設定活頁簿焦點
            wbk.Activate();
            try
            {
                // 引用第一個工作表
                wst = (Excel._Worksheet)wbk.Worksheets[1];
                // 命名工作表的名稱
                wst.Name = "kk citybike";
                // 設定工作表焦點
                wst.Activate();
                app.Cells[1, 1] = "city bike retrieve data";
                // 設定metainfo
                app.Cells[1, 1] = "站點";
                app.Cells[2, 1] = "目前車輛數";
                app.Cells[3, 1] = "空位數量";
                app.Cells[4, 1] = "總停車格";
                app.Cells[5, 1] = "lat,lon";
                // 設定第1列顏色
                wrange = wst.Range[wst.Cells[1, 1], wst.Cells[5, 1]];
                wrange.Select();
                wrange.Font.Color = ColorTranslator.ToOle(Color.Blue);//word color
                wrange.Interior.Color = ColorTranslator.ToOle(Color.White); //back color
                // 設定第2列資料
                for (int c = 2; c < bs.cbs.Count+2; c++)
                {
                    app.Cells[1, c] = c - 1; //站點
                    app.Cells[2, c] = bs.cbs[c - 2].StationNums1; //資料更新時間
                    app.Cells[3, c] = Convert.ToInt32( bs.cbs[c - 2].StationNums2);//目前車數
                    app.Cells[4, c] = Convert.ToInt32(bs.cbs[c - 2].StationNums1)+ Convert.ToInt32(bs.cbs[c - 2].StationNums2);//空格數
                    app.Cells[5, c] = bs.cbs[c - 2].StationLat+","+ bs.cbs[c - 2].StationLon;//總停車數
                }
                DateTime thisDay = DateTime.Today;
                savePath = "C:\\Users\\user\\Desktop\\rawdata\\" + "kk";
                try
                {
                    //另存活頁簿
                    wbk.SaveAs(savePath, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    wbk.SaveCopyAs("kk");
                    Console.WriteLine("儲存文件於 " + Environment.NewLine + savePath);
                }
                catch (Exception ex)
                {
                    Console.WriteLine("儲存檔案出錯，檔案可能正在使用" + Environment.NewLine + ex.Message);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("產生報表時出錯！" + Environment.NewLine + ex.Message);
            }

            //關閉活頁簿
            wbk.Close(false, Type.Missing, Type.Missing);
            //關閉Excel
            app.Quit();
            //釋放Excel資源
            System.Runtime.InteropServices.Marshal.ReleaseComObject(app);
            wbk = null;
            wst = null;
            wrange = null;
            app = null;
            GC.Collect();
        }

        //private static void test(object source, ElapsedEventArgs e)
        //{
        //    MessageBox.Show("OK, test event is fired at: " + DateTime.Now.ToString());
        //}
    }
}
