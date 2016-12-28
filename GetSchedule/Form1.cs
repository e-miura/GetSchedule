using System;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Outlook;
using System.IO;
using System.Runtime.InteropServices;
using System.Collections.Generic;
using System.Linq;
using GetSchedule.Common;
using System.Text;


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
            Excel.Application xlApp = null;
            Excel.Workbooks xlBooks = null;
            Excel.Workbook xlBook = null;
            Excel.Sheets xlSheets = null;
            Excel.Worksheet xlSheet = null;
            Excel.Range stRange = null;
            Excel.Range endRange = null;
            Excel.Range range = null;

            try
            {
                // Outlook アクセス
                Microsoft.Office.Interop.Outlook.Application outlook = new Microsoft.Office.Interop.Outlook.Application();
                NameSpace ns = outlook.GetNamespace("MAPI");
                MAPIFolder oFolder = ns.GetDefaultFolder(OlDefaultFolders.olFolderCalendar);
                Items oItems = oFolder.Items;
                AppointmentItem oAppoint = oItems.GetFirst();

                string startDate;
                string endDate;
                DateTime start, end;
                TimeSpan duration;
                TimeSpan sLunchTime, eLunchTime;

                bool schedule = false;

                var arrList = new List<Schedule>();
                var sortList = new List<SortSchedule>();

                // 予定表抽出
                while (oAppoint != null)
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
                    xlApp = new Excel.Application();
                    xlApp.Visible = false;

                    // openメソッド
                    xlBooks = xlApp.Workbooks;
                    xlBook = xlBooks.Open(Path.GetFullPath(FilePath));

                    // シート選択
                    xlSheets = xlBook.Worksheets;

                    bool flgSh = false;
                    int num = 1;

                    foreach (Excel.Worksheet sh in xlSheets)
                    {
                        if (yyyymm.Equals(sh.Name))
                        {
                            xlSheet = xlSheets[num] as Excel.Worksheet;
                            flgSh = true;
                            break;
                        }
                        num++;
                    }

                    if (!flgSh)
                    {
                        MessageBox.Show("指定した年月のシートがありません。");
                        return;
                    }

                    int rowDate = int.Parse(getValue("Excel", "rowDate").ToString());
                    int rowStart = int.Parse(getValue("Excel", "rowStart").ToString());
                    int colStart = int.Parse(getValue("Excel", "colStart").ToString());
                    //int rowMax = xlSheet.UsedRange.Rows.Count;
                    //int colMax = xlSheet.UsedRange.Columns.Count;
                    int rowMax = int.Parse(getValue("Excel", "rowEnd").ToString());
                    int colMax = int.Parse(getValue("Excel", "colEnd").ToString());
                    double dtime;
                    int colJOBID = int.Parse(getValue("Excel", "colJOBID").ToString());

                    // 指定した日付のエクセル列をクリア
                    foreach (var value in sortList)
                    {
                        for (int col = colStart; col < colMax + 1; col++)
                        {
                            if (xlSheet.Cells[rowDate, col].Value == null)
                            {
                                MessageBox.Show("エクセルシートの日付欄に空白があります。\r\nシートをご確認ください。"
                                    + "\r\n\r\n" + rowDate + "行" + col + "列目");
                                return;
                            }
                            else
                            {
                                if (value.date.Equals(xlSheet.Cells[rowDate, col].Value.ToString("yyyy/MM/dd")))
                                {
                                    // 該当日の列を初期化
                                    stRange = xlSheet.Cells[rowStart, col];
                                    endRange = xlSheet.Cells[rowMax, col];
                                    range = xlSheet.get_Range(stRange, endRange);
                                    range.Clear();
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
                            if (value.date.Equals(xlSheet.Cells[rowDate, col].Value.ToString("yyyy/MM/dd")))
                            {
                                for (int row = rowStart; row < rowMax + 1; row++)
                                {
                                    if (xlSheet.Cells[row, colJOBID].Value != null)
                                    {
                                        if (value.jobid.Equals(xlSheet.Cells[row, colJOBID].Value.ToString()))
                                        {
                                            dtime = (double)value.time.TotalHours;
                                            xlSheet.Cells[row, col] = ToRoundDown(dtime, 2);
                                            goto ExitLoop;
                                        }
                                    }
                                }
                            }
                        }
                        ExitLoop:;
                    }

                    xlApp.DisplayAlerts = false;
                    xlBook.SaveAs(FilePath);                        // Excel保存
                    MessageBox.Show("日報を更新しました。");
                }
                else
                {
                    MessageBox.Show("指定した日付の日報はありません。");
                }
            }
            catch (FormatException)
            {
                MessageBox.Show("iniファイルのフォーマット設定に問題があります。");
            }
            catch (COMException)
            {
                MessageBox.Show("エクセルが別のプロセスで使用中です。");
            }
            finally
            {
                if (xlBook != null)
                {
                    xlBook.Close();
                }
                if (xlApp != null)
                {
                    xlApp.Quit();
                }

                // Excelオブジェクト開放
                if (stRange != null)
                {
                    Marshal.ReleaseComObject(stRange);
                }
                if (endRange != null)
                {
                    Marshal.ReleaseComObject(endRange);
                }
                if (range != null)
                {
                    Marshal.ReleaseComObject(range);
                }
                if (xlSheet != null)
                {
                    Marshal.ReleaseComObject(xlSheet);
                }
                if (xlSheets != null)
                {
                    Marshal.ReleaseComObject(xlSheets);
                }
                if (xlBook != null)
                {
                    Marshal.ReleaseComObject(xlBook);
                }
                if (xlBooks != null)
                {
                    Marshal.ReleaseComObject(xlBooks);
                }
                if (xlApp != null)
                {
                    Marshal.ReleaseComObject(xlApp);
                }

                stRange = null;
                endRange = null;
                range = null;
                xlSheet = null;
                xlSheets = null;
                xlBook = null;
                xlBooks = null;
                xlApp = null;
                GC.Collect();
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
