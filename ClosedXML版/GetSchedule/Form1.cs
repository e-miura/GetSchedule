using System;
using System.Windows.Forms;
using Microsoft.Office.Interop.Outlook;
using System.IO;
using System.Runtime.InteropServices;
using System.Collections.Generic;
using System.Linq;
using GetSchedule.Common;
using System.Text;
using ClosedXML.Excel;

namespace GetSchedule
{
    public partial class Form1 : Form
    {

        public Form1()
        {
            InitializeComponent();
        }

        private void btnGetSchedule_Click(object sender, EventArgs e)
        {
            string yyyymm = dateFrom.Value.Date.ToString("yyyy年MM月度");
            string fdate = dateFrom.Value.Date.ToShortDateString();
            string tdate = dateTo.Value.Date.ToShortDateString();

            // 日付チェック
            if (fdate.CompareTo(tdate).Equals(1))
            {
                MessageBox.Show("To日付はFrom日付より新しい日付でなければなりません。");
                return;
            }

            string FilePath = getValue("Excel", "FilePath");

            // ファイルチェック
            if (File.Exists(FilePath) == false)
            {
                MessageBox.Show("ファイルを設定してください。");
                return;
            }

            if (FilePath.LastIndexOf(".xls") == -1 || FilePath.LastIndexOf(".xlsx") == -1)
            {
                MessageBox.Show("エクセルファイルを指定してください。");
                return;
            }

            // Excelオブジェクトの初期化
            XLWorkbook wb = null;
            IXLWorksheet ws = null;

            try
            {
                // Outlook アクセス
                Microsoft.Office.Interop.Outlook.Application outlook = new Microsoft.Office.Interop.Outlook.Application();
                NameSpace ns = outlook.GetNamespace("MAPI");
                MAPIFolder oFolder = ns.GetDefaultFolder(OlDefaultFolders.olFolderCalendar);
                Items oItems = oFolder.Items;
                AppointmentItem oAppoint = oItems.GetFirst();

                string startDate = oAppoint.Start.ToShortDateString();
                string endDate = oAppoint.End.ToShortDateString();
                DateTime start, end;
                TimeSpan duration;
                TimeSpan sLunchTime, eLunchTime;

                bool schedule = false;

                var arrList = new List<Schedule>();
                var sortList = new List<SortSchedule>();

                // 予定表抽出
                while (oAppoint != null && tdate.ToString().CompareTo(endDate) <= 1)
                {
                    startDate = oAppoint.Start.ToShortDateString();
                    endDate = oAppoint.End.ToShortDateString();

                    // 指定した日付の範囲の予定表を抽出
                    if ((startDate.CompareTo(fdate.ToString()).Equals(0) || startDate.CompareTo(fdate.ToString()).Equals(1))
                        && (tdate.ToString().CompareTo(endDate).Equals(0) || tdate.ToString().CompareTo(endDate).Equals(1)))
                    {
                        if (string.IsNullOrEmpty(oAppoint.Body) == false)
                        {
                            // JOBID
                            string[] arrBody = oAppoint.Body.Split('\n');
                            if (arrBody.Length > 0)
                            {
                                // 時間集計
                                sLunchTime = oAppoint.Start.TimeOfDay;
                                eLunchTime = oAppoint.End.TimeOfDay;

                                // お昼休み考慮
                                if ((sLunchTime.CompareTo(TimeSpan.Parse("12:00:00")) == -1 || sLunchTime.CompareTo(TimeSpan.Parse("12:00:00")) == 0)
                                    && (eLunchTime.CompareTo(TimeSpan.Parse("13:00:00")) == 0 || eLunchTime.CompareTo(TimeSpan.Parse("13:00:00")) == 1))
                                {
                                    start = DateTime.Parse(oAppoint.Start.AddHours(1).ToString("yyyy/MM/dd HH:mm"));
                                }
                                else
                                {

                                    start = DateTime.Parse(oAppoint.Start.ToString("yyyy/MM/dd HH:mm"));
                                }
                                end = DateTime.Parse(oAppoint.End.ToString("yyyy/MM/dd HH:mm"));

                                duration = end - start;


                                // 構造体格納
                                arrList.Add(new Schedule(oAppoint.Start.ToString("yyyy/MM/dd"),
                                                arrBody[0].Replace("\r", "").Replace("#", ""), oAppoint.Body, duration));
                                schedule = true;
                            }
                        }
                    }
                    oAppoint = oItems.GetNext();
                }

                if (schedule)
                {
                    IOrderedEnumerable<Schedule> sortScdl = arrList.OrderBy(a => a.date).ThenBy(a => a.jobid);

                    // 日付、JOBIDごとの時間集計
                    var first = sortScdl.First();
                    string tmpDate = first.date;
                    string tmpJobid = first.jobid;
                    string tmpBody = first.body;
                    TimeSpan total = new TimeSpan(0, 0, 0);

                    foreach (var s in sortScdl)
                    {
                        if (s.date.Equals(tmpDate) && s.jobid.Equals(tmpJobid))
                        {
                            total += s.time;
                        }
                        else
                        {
                            sortList.Add(new SortSchedule(tmpDate, tmpJobid, tmpBody, total));

                            // 初期化
                            total = new TimeSpan(0, 0, 0);
                            total = s.time;
                        }
                        tmpDate = s.date;
                        tmpJobid = s.jobid;
                        tmpBody = s.body;
                    }

                    // 最後の一件
                    sortList.Add(new SortSchedule(tmpDate, tmpJobid, tmpBody, total));

                    /*** エクセル出力 ***/
                    // openメソッド
                    using (wb = new XLWorkbook(Path.GetFullPath(FilePath)))
                    {
                        bool flgSh = false;

                        foreach (var s in wb.Worksheets)
                        {
                            // シート選択
                            if (yyyymm.Equals(s.Name))
                            {
                                ws = wb.Worksheet(s.Name);
                                flgSh = true;
                                break;
                            }
                        }

                        if (!flgSh)
                        {
                            MessageBox.Show("指定した年月のシートがありません。");
                            return;
                        }
                        
                        // 最終行
                        int lastRow = ws.LastRowUsed().RowNumber();

                        int rowDate = int.Parse(getValue("Excel", "rowDate").ToString());
                        int rowStart = int.Parse(getValue("Excel", "rowStart").ToString());
                        int colStart = int.Parse(getValue("Excel", "colStart").ToString());
                        int rowMax = ws.LastRowUsed().RowNumber();
                        int colMax = ws.LastColumnUsed().ColumnNumber();
                        //int rowMax = int.Parse(getValue("Excel", "rowEnd").ToString());
                        //int colMax = int.Parse(getValue("Excel", "colEnd").ToString());
                        double dtime;
                        int colJOBID = int.Parse(getValue("Excel", "colJOBID").ToString());
                        string cellDate;


                        // 指定した日付のエクセル列をクリア
                        foreach (var value in sortList)
                        {
                            for (int col = colStart; col < colMax + 1; col++)
                            {
                                cellDate = ws.Cell(rowDate, col).Value.ToString();

                                if (string.IsNullOrEmpty(cellDate))
                                {
                                    MessageBox.Show("エクセルシートの日付欄に空白があります。\r\nシートをご確認ください。"
                                        + "\r\n\r\n" + rowDate + "行" + col + "列目");
                                    return;
                                }
                                else
                                {
                                    if (value.date.Equals(cellDate.Substring(0,10)))
                                    {
                                        // 該当日の列を初期化
                                        ws.Range(ws.Cell(rowStart, col), ws.Cell(rowMax, col)).Clear();
                                        break;
                                    }
                                }
                            }
                        }

                        // エクセル出力
                        foreach (var value in sortList)
                        {
                            for (int col = colStart; col < colMax + 1; col++)
                            {
                                cellDate = ws.Cell(rowDate, col).Value.ToString();

                                if (value.date.Equals(cellDate.Substring(0, 10)))
                                {
                                    for (int row = rowStart; row < rowMax + 1; row++)
                                    {
                                        if (ws.Cell(row, colJOBID).Value != null)
                                        {
                                            if (value.jobid.Equals(ws.Cell(row, colJOBID).Value.ToString()))
                                            {
                                                dtime = (double)value.time.TotalHours;
                                                ws.Cell(row, col).Value = ToRoundDown(dtime, 2);
                                                goto ExitLoop;
                                            }
                                        }
                                    }
                                }
                            }
                            ExitLoop:;
                        }
                    }
                    
                    wb.SaveAs(FilePath);                        // Excel保存
                    MessageBox.Show("日報を更新しました。");
                }
                else
                {
                    MessageBox.Show("指定した日付の日報はありません。");
                }
            }
            catch (FormatException)
            {
                MessageBox.Show("フォーマット設定に問題があります。");
            }
            catch (IOException)
            {
                MessageBox.Show("エクセルが別のプロセスで使用中です。");
            }
        }

        [DllImport("KERNEL32.DLL")]
        public static extern uint
        GetPrivateProfileString(string lpAppName, string lpKeyName, string lpDefault, StringBuilder lpReturnedString, uint nSize, string lpFileName);
        private static String getValue(String section, String key)
        {
            StringBuilder sb = new StringBuilder(1024);
            GetPrivateProfileString(section, key, "", sb, Convert.ToUInt32(sb.Capacity),
                Path.GetDirectoryName(System.Reflection.Assembly.GetEntryAssembly().Location) + "\\GetSchedule.ini");

            return sb.ToString();
        }

        private static double ToRoundDown(double dValue, int iDigits)
        {
            double dCoef = Math.Pow(10, iDigits);

            return dValue > 0 ? Math.Floor(dValue * dCoef) / dCoef :
                                Math.Ceiling(dValue * dCoef) / dCoef;
        }

        private void dateFrom_ValueChanged(object sender, EventArgs e)
        {
            label1.Text = dateFrom.Value.Date.ToString("yyyy年MM月") + "分の日報を出力します。";
            if (dateFrom.Text.CompareTo(dateTo.Text).Equals(1))
            {
                dateTo.Text = dateFrom.Text;
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            label1.Text = DateTime.Today.ToString("yyyy年MM月") + "分の日報を出力します。";
            dateFrom.Text = DateTime.Today.ToString("yyyy/MM/dd");
            dateTo.Text = DateTime.Today.ToString("yyyy/MM/dd");
        }
    }
}
