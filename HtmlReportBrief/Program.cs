// See https://aka.ms/new-console-template for more information

using System;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Data;

class Program
{
    static void Main(string[] args)
    {
        Console.OutputEncoding = Encoding.UTF8;
        Version version = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;
        Console.WriteLine($"\r \nCurrent test program version : {version}" + "\r\n");

        Application? app = default;
        _Workbook? _wbk = default;
        Sheets _sheets;
        _Worksheet _masterSheet;
        try
        {
            string htmlPath = string.Empty;
            if (args.FirstOrDefault() != null)
            {
                htmlPath = args[0];
                if (!IsValidPath(htmlPath))
                {
                    throw new InvalidOperationException(); 
                }
            }
            else
            {
                htmlPath = System.Environment.CurrentDirectory;
            }
            
            app = new Application();
            _wbk = app.Workbooks.Add();
            _sheets = _wbk.Worksheets;
            _masterSheet = (_Worksheet)_wbk.Sheets.Add();
            _masterSheet.Name = "MasterBrief";
            _masterSheet.Columns[1].ColumnWidth = 50;
            _masterSheet.Range["O1"].Value = "FailReportLink";
            _masterSheet.Range["O1"].Font.Bold = true;
            DirectoryInfo currentPath;
            if (args[0] != null)
            {
                currentPath = new DirectoryInfo(htmlPath);
            }
            else
            {
                currentPath = new DirectoryInfo(Path.Combine(htmlPath, "Results"));
            }
            Dictionary<string, (int Pass, int Fail, int Skip)> data = new Dictionary<string, (int Pass, int Fail, int Skip)>();

            List<FileTimeInfo> fileTimeInfoList = new List<FileTimeInfo>();
            foreach (FileInfo file in currentPath.GetFiles())
            {
                if (file.Extension.ToLower() == ".html")
                {
                    fileTimeInfoList.Add(new FileTimeInfo()
                    {
                        FileName = file.Name,
                        FileCreateTime = file.LastWriteTime
                    });
                }
            }
            var f = from x in fileTimeInfoList
                    orderby x.FileCreateTime
                    select x;

            int workSheetNumber = 2;
            int reportFailCount = 0;
            int reportSuccessCount = 0;
            int failReportCount = 0;

            int warningFileCount = 500;
            Console.WriteLine($"There are {f.Count()} files wait for parse.");
            if (f.Count()> warningFileCount)
            {
                Console.WriteLine($"Warning!!! Excel deal with data and create sheets count more than {warningFileCount}.\r\nIt will spend a lot of time and depend on you system performance.\r\nSuggest separate files count less than 200 running this program\r\n");
            }
            int fileCount = 0;
            foreach (var item in f)
            {
                fileCount++;
                // Add new sheet after exist one.
                _sheets.Add(After: _wbk.Sheets[_wbk.Sheets.Count]);

                Stream myStream = new FileStream(Path.Combine(currentPath.ToString(), item.FileName), FileMode.Open);
                //Read html format by UTF-8
                Encoding encode = System.Text.Encoding.GetEncoding("UTF-8");
                
                StreamReader myStreamReader = new StreamReader(myStream, encode);

                string strhtml = myStreamReader.ReadToEnd();

                //Get the content contained in <div></div> in html through regular expressions
                string divPatten = "(<div(.*?)>)(.|\n)*?(</div>)";
                MatchCollection mcdiv = Regex.Matches(strhtml, divPatten);
                FileTimeInfo fileTimeInfo = new FileTimeInfo();
                Tuple<MatchCollection, MatchCollection, MatchCollection> tuple = fileTimeInfo.InputFormatCheck(item, strhtml);

                string mStrPass = string.Empty;
                string htmlScriPass = string.Empty;

                MatchCollection mcPass = tuple.Item1;
                for (int i = 0; i < mcPass.Count; i++)
                {
                    mStrPass = mcPass[i].ToString().Replace("<div><span class=\"pass\">✔</span><span> ", "").Split('<')[0];
                    htmlScriPass += mStrPass + "\r\n";

                    if (data.ContainsKey(mStrPass))
                    {
                        var value = data[mStrPass];
                        value.Pass += 1;
                        data[mStrPass] = value;
                    }
                    else
                    {
                        data[mStrPass] = (1, 0, 0); // If key not exist, add new one, pass value 1, Fail and Skip value 0
                    }
                }

                string mStrFail = string.Empty;
                string htmlScriFail = string.Empty;
                MatchCollection mcFail = tuple.Item2;
                List<string> errorMessageList = new List<string>();
                string errorMessage = string.Empty;
                for (int i = 0; i < mcFail.Count / 2; i++)
                {
                    mStrFail = mcFail[i].ToString().Replace("<div><span class=\"fail\">✘</span><span> ", "").Split('<')[0];
                    for (int j = 0; j < mcdiv.Count; j++)
                    {
                        if (mcdiv[j].ToString().Contains(mcFail[i].ToString()))
                        {
                            errorMessageList.Add(mcdiv[j + 1].ToString());
                        }
                    }
                    htmlScriFail += mStrFail + "\r\n";

                    if (data.ContainsKey(mStrFail))
                    {
                        var value = data[mStrFail];
                        value.Fail += 1;
                        data[mStrFail] = value;
                    }
                    else
                    {
                        data[mStrFail] = (0, 1, 0);
                    }
                }
                for (int i = 0; i < errorMessageList.Count / 2; i++)
                {
                    string temp = errorMessageList[i].ToString();
                    var starIndex = temp.IndexOf("<pre>");
                    var endIndex = temp.IndexOf("</pre>");
                    errorMessage += temp.Substring(starIndex + 5, endIndex - starIndex - 5) + "\r\n";
                }
                if (mcFail.Count / 2 > 0)
                {
                    reportFailCount++;
                }
                else
                {
                    reportSuccessCount++;
                }

                string mStrSkip = string.Empty;
                string htmlScriSkip = string.Empty;
                MatchCollection mcSkip = tuple.Item3;
                string skipMessage = string.Empty;
                for (int i = 0; i < mcSkip.Count; i++)
                {
                    mStrSkip = mcSkip[i].ToString().Replace("<div><span class=\"skip\">❢</span><span> ", "").Split('<')[0];
                    for (int j = 0; j < mcdiv.Count; j++)
                    {
                        if (mcdiv[j].ToString().Contains(mcSkip[i].ToString()))
                        {
                            string temp = mcdiv[j + 1].ToString();
                            var starIndex = temp.IndexOf("<pre>");
                            var endIndex = temp.IndexOf("</pre>");
                            if (starIndex < 0)
                            {
                                skipMessage += "Skip reason not be defined" + "\r\n";
                            }
                            else
                            {
                                skipMessage += temp.Substring(starIndex + 5, endIndex - starIndex - 5) + "\r\n";
                            }
                        }
                    }
                    htmlScriSkip += mStrSkip + "\r\n";

                    if (data.ContainsKey(mStrSkip))
                    {
                        var value = data[mStrSkip];
                        value.Skip += 1;
                        data[mStrSkip] = value;
                    }
                    else
                    {
                        data[mStrSkip] = (0, 0, 1);
                    }
                }

                _Worksheet _sheet = (_Worksheet)_wbk.Sheets[workSheetNumber];

                if (item.FileName.Length >= 31)
                {
                    _sheet.Name = item.FileName.Substring(0, 31).Replace(" ", "");
                }
                else
                {
                    _sheet.Name = item.FileName.Replace(" ", "");
                }

                if (htmlScriFail != string.Empty)
                {

                    _masterSheet.Hyperlinks.Add(Anchor: _masterSheet.Range[$"O{failReportCount + 2}"],
                                          Address: "",
                                          SubAddress: _sheet.Name + "!A1",
                                          TextToDisplay: $"{item.FileName}"
                                      );
                    failReportCount++;
                }

                // Pass cell input
                // Cells[1,1] First 1 is row 1, Second 1 is column 1
                _sheet.Cells[1, 1] = "Pass" + ":\r\n" + htmlScriPass;
                _sheet.Cells[1, 2] = mcPass.Count;
                _sheet.Cells[1, 1].Font.Color = System.Drawing.Color.Green;

                // Fail cell input
                _sheet.Cells[2, 1] = "Fail" + ":\r\n" + htmlScriFail;
                _sheet.Cells[2, 2] = mcFail.Count / 2;
                _sheet.Cells[2, 3] = "Fail" + ":\r\n" + errorMessage;
                _sheet.Cells[2, 1].Font.Color = System.Drawing.Color.Red;
                _sheet.Cells[2, 3].Font.Color = System.Drawing.Color.Red;

                // Skip cell input
                _sheet.Cells[3, 1] = "Skip" + ":\r\n" + htmlScriSkip;
                _sheet.Cells[3, 2] = mcSkip.Count;
                _sheet.Cells[3, 3] = "Skip" + ":\r\n" + skipMessage;
                _sheet.Cells[3, 1].Font.Color = System.Drawing.Color.Brown;
                _sheet.Cells[3, 3].Font.Color = System.Drawing.Color.Brown;

                // Add link back to the master sheet
                var targetSource = _sheets["MasterBrief"];
                _sheet.Hyperlinks.Add(Anchor: _sheet.Range["D1"],
                                          Address: "",
                                          SubAddress: targetSource.Name + "!A1",
                                          TextToDisplay: "Back to MasterBrief sheet"
                                      );
                _sheet.Range["D1"].Font.Bold = true;
                _sheet.Range["D1"].Font.Size = 12;

                _sheet.Columns[1].ColumnWidth = 80;
                _sheet.Columns[3].ColumnWidth = 80;


                workSheetNumber++;
                Console.WriteLine($"{item.FileName} sheet data generate done, has {f.Count() - fileCount} files left");
            }

            _Worksheet _sheetMaster = (_Worksheet)_wbk.Sheets[1];
            _sheetMaster.Range["A1"].Value = "Category";
            _sheetMaster.Range["A1"].Font.Bold = true;
            _sheetMaster.Range["B1"].Value = "TestingIterationsPassRate";
            _sheetMaster.Range["B1"].Font.Bold = true;
            _sheetMaster.Range["C1"].Value = "Counts";
            _sheetMaster.Range["C1"].Font.Bold = true;
            _sheetMaster.Range["A2"].Value = "PASS";
            _sheetMaster.Range["A2"].Font.Color = System.Drawing.Color.Green.ToArgb();
            _sheetMaster.Range["B2"].NumberFormat = "0.00%";
            _sheetMaster.Range["B2"].Value = double.Parse(reportSuccessCount.ToString()) / double.Parse(currentPath.GetFiles().Count().ToString());
            _sheetMaster.Range["C2"].Value = reportSuccessCount;
            _sheetMaster.Range["A3"].Value = "FAIL";
            _sheetMaster.Range["A3"].Font.Color = System.Drawing.Color.Red.ToKnownColor();
            _sheetMaster.Range["B3"].NumberFormat = "0.00%";
            _sheetMaster.Range["B3"].Value = double.Parse(reportFailCount.ToString()) / double.Parse(currentPath.GetFiles().Count().ToString());
            _sheetMaster.Range["C3"].Value = reportFailCount;

            // Create chart object
            Microsoft.Office.Interop.Excel.ChartObjects chartObjects = (Microsoft.Office.Interop.Excel.ChartObjects)_sheetMaster.ChartObjects(Type.Missing);
            Microsoft.Office.Interop.Excel.ChartObject chartObject = chartObjects.Add(500, 0, 300, 300);
            Microsoft.Office.Interop.Excel.Chart chart = chartObject.Chart;
            // Set chart data range
            Microsoft.Office.Interop.Excel.Range dataRange = _sheetMaster.Range["A1:B3"];
            chart.SetSourceData(dataRange);
            // Set chart type
            chart.ChartType = Microsoft.Office.Interop.Excel.XlChartType.xlPie;
            // Chart visible
            chartObject.Visible = true;
            // Add data label
            chart.SeriesCollection(1).HasDataLabels = true;
            DataLabels dataLabels = chart.SeriesCollection(1).DataLabels;
            // Set data label format
            dataLabels.NumberFormat = "0.0%";
            dataLabels.Position = XlDataLabelPosition.xlLabelPositionCenter;
            // Show value
            dataLabels.ShowValue = true;


            _sheetMaster.Range["A24"].Value = "TestCaseName";
            _sheetMaster.Range["A24"].Font.Bold = true;
            _sheetMaster.Range["B24"].Value = "PASS Rate";
            _sheetMaster.Range["B24"].Font.Bold = true;
            _sheetMaster.Range["C24"].Value = "FAIL Rate";
            _sheetMaster.Range["C24"].Font.Bold = true;
            _sheetMaster.Range["D24"].Value = "SKIP Rate";
            _sheetMaster.Range["D24"].Font.Bold = true;

            int row = 25;
            foreach (var item in data)
            {
                _sheetMaster.Cells[row, 1] = item.Key;
                _sheetMaster.Cells[row, 2].Value = double.Parse(item.Value.Pass.ToString()) / double.Parse((item.Value.Pass + item.Value.Fail + item.Value.Skip).ToString());
                _sheetMaster.Cells[row, 2].NumberFormat = "0.00%";
                _sheetMaster.Cells[row, 3].Value = double.Parse(item.Value.Fail.ToString()) / double.Parse((item.Value.Pass + item.Value.Fail + item.Value.Skip).ToString());
                _sheetMaster.Cells[row, 3].NumberFormat = "0.00%";
                _sheetMaster.Cells[row, 4].Value = double.Parse(item.Value.Skip.ToString()) / double.Parse((item.Value.Pass + item.Value.Fail + item.Value.Skip).ToString());
                _sheetMaster.Cells[row, 4].NumberFormat = "0.00%";
                row++;
            }

            // Create chart object
            Microsoft.Office.Interop.Excel.ChartObjects chartObjects1 = (Microsoft.Office.Interop.Excel.ChartObjects)_sheetMaster.ChartObjects(Type.Missing);
            Microsoft.Office.Interop.Excel.ChartObject chartObject1 = chartObjects1.Add(460, 330, 1000, 550);
            Microsoft.Office.Interop.Excel.Chart chart1 = chartObject1.Chart;

            // Set chart data range
            Microsoft.Office.Interop.Excel.Range dataRange1 = _sheetMaster.Range[$"A24:D{data.Count + 24}"];
            chart1.SetSourceData(dataRange1);
            if (data.Count <= 3)
            {
                // 设置图表类型为簇状柱状图
                chart1.ChartType = Microsoft.Office.Interop.Excel.XlChartType.xlColumnClustered;
            }
            else
            {
                // 设置图表类型为堆积柱状图
                chart1.ChartType = Microsoft.Office.Interop.Excel.XlChartType.xlColumnStacked100;
            }

            chartObject1.Visible = true;

            // Add data label
            foreach (Microsoft.Office.Interop.Excel.Series series in chart1.SeriesCollection())
            {
                series.HasDataLabels = true;
                foreach (DataLabel dataLable in series.DataLabels())
                {
                    dataLable.NumberFormat = "0.0%";
                    dataLable.Position = XlDataLabelPosition.xlLabelPositionCenter;
                    if (dataLable.Text == "100.0%" || dataLable.Text == "0.0%")
                    {
                        dataLable.ShowValue = false;
                    }
                }
            }

            for (int seriesIndex = 1; seriesIndex <= chart1.SeriesCollection().Count; seriesIndex++)
            {
                Series series = (Series)chart1.SeriesCollection(seriesIndex);

                // Set colour
                switch (seriesIndex % 3)
                {
                    case 0: // Grey
                        series.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Gray);
                        break;
                    case 1: // Green
                        series.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Green);
                        break;
                    case 2: // Red
                        series.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                        break;
                }
            }

            chart1.ChartStyle = 10; // Change type
            chart1.HasTitle = true;
            chart1.ChartTitle.Text = "Test Case Pass Rate";

            _sheetMaster.Select();
            _wbk.SaveAs(htmlPath + "\\" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + "_" + "HtmlReportBrief.xlsx"); 
        }
        catch (Exception e)
        {
            throw e;
        }
        finally
        {
            _wbk?.Close();
            app?.Quit(); 

            System.GC.GetGeneration(app);
            IntPtr t = new IntPtr(app.Hwnd);  //Get handle
            int k = 0;

            GetWindowThreadProcessId(t, out k);
            System.Diagnostics.Process p = System.Diagnostics.Process.GetProcessById(k);
            p.Kill();
        }
    }

    [DllImport("User32.dll", CharSet = CharSet.Auto)]
    static extern int GetWindowThreadProcessId(IntPtr t, out int k);

    public static bool IsValidPath(string path)
    {
        return Uri.TryCreate(path, UriKind.Absolute, out _);
    }

    public class FileTimeInfo
    {
        public string? FileName;
        public DateTime FileCreateTime;

        public Tuple<MatchCollection, MatchCollection, MatchCollection> InputFormatCheck(FileTimeInfo fileName, string readFromHtml)
        {
            List<string> format1 = new List<string>();
            format1.Add("(<div><span class=\"pass\"(.*?)>)(.|\n)*?(</span><br></div>)");
            format1.Add("(<div><span class=\"fail\"(.*?)>)(.|\n)*?(</span><br></div>)");
            format1.Add("(<div><span class=\"skip\"(.*?)>)(.|\n)*?(</span><br></div>)");
            List<string> format2 = new List<string>();
            format2.Add("(<div><span class=\"pass\"(.*?)>)(.|\n)*?(</span><br /></div>)");
            format2.Add("(<div><span class=\"fail\"(.*?)>)(.|\n)*?(</span><br /></div>)");
            format2.Add("(<div><span class=\"skip\"(.*?)>)(.|\n)*?(</span><br /></div>)");
            int format1Count = 0;
            int format2Count = 0;
            for (int i = 0; i < format1.Count; i++)
            {
                format1Count += Regex.Matches(readFromHtml, format1[i]).Count;
            }
            for (int i = 0; i < format2.Count; i++)
            {
                format2Count += Regex.Matches(readFromHtml, format2[i]).Count;
            }
            if (format1Count != 0)
            {
                return Tuple.Create(Regex.Matches(readFromHtml, format1[0]), Regex.Matches(readFromHtml, format1[1]), Regex.Matches(readFromHtml, format1[2]));
            }
            else if (format2Count != 0)
            {
                return Tuple.Create(Regex.Matches(readFromHtml, format2[0]), Regex.Matches(readFromHtml, format2[1]), Regex.Matches(readFromHtml, format2[2]));
            }
            else
            {
                throw new Exception($"{fileName.FileName} format incorrect, Please check input html file");
            }
        }
    }
}








