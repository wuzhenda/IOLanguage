using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Collections.Specialized;
using System.ComponentModel.Design;
using System.IO;
using System.Resources;
using System.Text.RegularExpressions;
using ClosedXML.Excel;
using System.Windows.Forms;
using System.Collections;
using System.Diagnostics;

using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Markup;
using System.Windows;
using System.Xml;

namespace IOLanguages
{
    class Program
    {
       
       const string DefaultLang = "__DEFAULT__";
        static string BASE_INPUT_PATH = String.Empty;
        //string BASE_OUTPUT_PATH = System.Environment.CurrentDirectory + @"/output/";

        // exe /o:file /p:path /P:pattern
        static void Main(string[] args)
        {
            BASE_INPUT_PATH = System.Environment.CurrentDirectory + @"/input/";

            //generateXamlResourceFileFromExcel(@"X-Sign basic_6 languages String table_1127_2014.xls"
            //    , @"C:\Users\Daryl.Huang\Desktop\DesignerTrack_0918\六國語言 - 副本\StringResource.en-US.xaml","");

            //generateXmlDocFromExcel(@"Designer_Property_1127_2014.xls");
            //return;



            string path = "./";
            string pattern = "StringResource";
            string output = null;
            bool isExport = true;

            string help = "For exporting: exe /o:output /e:path /P:pattern\r\n" +
                    "OR\r\nFor importing: exe /i:filePath\r\n\r\n" +
                    "/o\t[MUST for exporting]: Specify output file\r\n\r\n" +
                    "/i\t[MUST for importing]: Specify file to be imported\r\n\r\n" +
                    "/e\t[in lower, OPTIONAL]: specify resource file path to be exported, \r\n" +
                    "\tif not present, will be current folder\r\n" +
                    "\tThe file name MUST be in format: FileName.REGION.xaml\r\n\r\n" +
                    "/P\t[in upper, OPTIONAL]: specify resource file name pattern, \r\n" +
                    "\tif not present, default will be: StringResource\r\n\r\n" +
                    "For exporting xaml:exe /ga:translationExcelName /s:srcExcelName[optional]\r\n" +
                    "For exporting xml:exe /gx:translationExcelName";

            Dictionary<string, string> keyVal = ParseArguments(args);

            if (keyVal.ContainsKey("ga") || keyVal.ContainsKey("gx"))
            {
                //if ()
                if (keyVal.ContainsKey("ga"))
                {
                    if (keyVal.ContainsKey("s"))
                    {
                        generateXamlResourceFileFromExcel(keyVal["ga"], String.Empty, keyVal["s"]);
                    }
                    else
                    {
                        generateXamlResourceFileFromExcel(keyVal["ga"], String.Empty, String.Empty);
                    }
                }

                if (keyVal.ContainsKey("gx"))
                {
                    generateXmlDocFromExcel(keyVal["gx"]);
                }
                return;
            }

            if (!keyVal.ContainsKey("import") && !keyVal.ContainsKey("i") &&
                !keyVal.ContainsKey("export") && !keyVal.ContainsKey("e"))
            {
                System.Windows.Forms.MessageBox.Show("Please specify /i or /e!\r\n\r\n" + help, "Warning");
                return;
            }

            if (keyVal.ContainsKey("export"))
            {
                path = keyVal["export"];
            }
            else if (keyVal.ContainsKey("e"))
            {
                path = keyVal["e"];
            }
            else
            {
                isExport = false;
            }

            if (isExport)
            {
                if (!keyVal.ContainsKey("o") && !keyVal.ContainsKey("output"))
                {
                    System.Windows.Forms.MessageBox.Show("Please specify /o for exporting!\r\n\r\n" + help, "Warning");
                    return;
                }

                if (keyVal.ContainsKey("output"))
                {
                    output = keyVal["output"];
                }
                else if (keyVal.ContainsKey("o"))
                {
                    output = keyVal["o"];
                }

                if (keyVal.ContainsKey("pattern"))
                {
                    pattern = keyVal["pattern"];
                }
                else if (keyVal.ContainsKey("P"))
                {
                    pattern = keyVal["p"];
                }

                Dictionary<string, DataSet> dataSetOfAllResourceFiles = loadAllResources(path, pattern);
                //Dictionary<string, DataSet> dataSetOfAllResourceFiles = loadAllResources("./", "*");
                exportToExcel(dataSetOfAllResourceFiles, output + ".xls");
            }
            else
            {
                if (keyVal.ContainsKey("import"))
                {
                    path = keyVal["import"];
                }
                else if (keyVal.ContainsKey("i"))
                {
                    path = keyVal["i"];
                }
            }
        }

        private static Dictionary<string, string> ParseArguments(string[] args)
        {
            Dictionary<string, string> argsKeyVal = new Dictionary<string, string>();

            foreach (string a in args)
            {
                int index = a.IndexOf(':');
                if (index != -1)
                {
                    string key = a.Substring(1, index - 1);
                    string val = a.Substring(index + 1);
                    if (!argsKeyVal.ContainsKey(key))
                    {
                        argsKeyVal.Add(key, val);
                    }
                }
            }
            return argsKeyVal;
        }

        public static Dictionary<string, DataSet> loadAllResources(string path, string pattern)
        {
            Dictionary<string, DataSet> mapOfAllDataSet = new Dictionary<string, DataSet>();

            //string curPath = Directory.GetCurrentDirectory();

            foreach (string file in Directory.GetFiles(path, pattern + ".*xaml", SearchOption.AllDirectories))
            {
                string lang = "";

                try
                {
                    string p = (pattern == "*" ? ".+" : pattern) + "[.](?<l>.+)[.]xaml";

                    var m = Regex.Match(file, p, RegexOptions.IgnoreCase);
                    if (m.Success)
                    {
                        lang = m.Groups["l"].Value;
                    }
                    else
                    {
                        Debug.WriteLine("Warning: " + file + "is not in pattern: " + p);
                        continue;
                    }
                }
                catch (System.Exception ex)
                {
                    System.Windows.Forms.MessageBox.Show(ex.Message, "ERROR");
                    continue;
                }

                DataSet dataSet = new DataSet();
                dataSet.ReadXml(file, XmlReadMode.InferSchema);

                DataTable tableOfAllLang;
                int index = file.LastIndexOf((lang == "" ? lang : ("." + lang)) + ".xaml", StringComparison.OrdinalIgnoreCase);
                string filePathWithPrefixName = file.Substring(0, index);
                if (!mapOfAllDataSet.ContainsKey(filePathWithPrefixName))
                {
                    DataSet dataSetAddToMap = new DataSet();
                    mapOfAllDataSet.Add(filePathWithPrefixName, dataSetAddToMap);

                    // for importing
                    //                     dataSetAddToMap.ReadXmlSchema(new MemoryStream(ASCIIEncoding.Default.GetBytes(
                    //                         dataSet.GetXmlSchema())));

                    tableOfAllLang = new DataTable();
                    tableOfAllLang.TableName = filePathWithPrefixName;

                    DataColumn idColumn = new DataColumn("ID");
                    tableOfAllLang.Columns.Add(idColumn);
                    tableOfAllLang.PrimaryKey = new DataColumn[] { idColumn };

                    DataRow row = tableOfAllLang.NewRow();
                    row["ID"] = filePathWithPrefixName;
                    tableOfAllLang.Rows.Add(row);

                    row = tableOfAllLang.NewRow();
                    //row["ID"] = dataSet.GetXmlSchema().Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;").Replace("\"", "&quot;").Replace("'", "&apos;");
                    row["ID"] = System.Security.SecurityElement.Escape(dataSet.GetXmlSchema());
                    tableOfAllLang.Rows.Add(row);

                    dataSetAddToMap.Tables.Add(tableOfAllLang);
                }
                else
                {
                    tableOfAllLang = mapOfAllDataSet[filePathWithPrefixName].Tables[0];
                }

                if (lang == "")
                {
                    lang = DefaultLang;
                }

                tableOfAllLang.Columns.Add(new DataColumn(lang));

                // Then display informations to test
                // #region test for displaying informations
                //                 foreach (DataTable table in dataSet.Tables)
                //                 {
                //                     Console.WriteLine(table);
                //                     for (int i = 0; i < table.Columns.Count; ++i)
                //                         Console.Write("\t\t" + table.Columns[i].ColumnName);
                //                     Console.WriteLine();
                // 
                //                     if (table.ChildRelations.Count > 0)
                //                     {
                //                         foreach (var row in table.AsEnumerable())
                //                         {
                //                             for (int x = 0; x < table.ChildRelations.Count; x++)
                //                             {
                //                                 for (int i = 0; i < table.Columns.Count; ++i)
                //                                 {
                //                                     Console.Write("\t\t" + row[i]);
                //                                 }
                //                                 Console.WriteLine();
                // 
                //                                 DataRow[] childRows = row.GetChildRows(table.ChildRelations[x]);
                //                                 DataTable tableOfChildRow;
                //                                 for (int y = 0; y < childRows.Length; y++)
                //                                 {
                //                                     tableOfChildRow = childRows[y].Table;
                //                                     for (int i = 0; i < tableOfChildRow.Columns.Count; ++i)
                //                                     {
                //                                         Console.Write("\t\t" + childRows[y][i]);
                //                                     }
                //                                     Console.WriteLine();
                //                                 }
                //                             }
                //                             Console.WriteLine();
                //                         }
                //                     }
                //                     else
                //                     {
                //                         if (table.ParentRelations.Count == 0)
                //                         {
                //                             foreach (var row in table.AsEnumerable())
                //                             {
                //                                 for (int i = 0; i < table.Columns.Count; ++i)
                //                                 {
                //                                     Console.Write("\t\t" + row[i]);
                //                                 }
                //                                 Console.WriteLine();
                //                             }
                //                         }
                //                     }
                //                 }
                // #endregion

                DataTable dataSetTable = dataSet.Tables[0];
                for (int r = 0; r < dataSetTable.Rows.Count; ++r)
                {
                    DataRow row = dataSetTable.Rows[r];
                    string key = row[0].ToString();
                    string val = row[1].ToString();

                    row = tableOfAllLang.Rows.Find(key);
                    if (row == null)
                    {
                        row = tableOfAllLang.NewRow();
                        row["ID"] = key;
                        tableOfAllLang.Rows.Add(row);
                    }
                    Debug.Assert(row != null);
                    row[lang] = val;
                }

                mapOfAllDataSet[filePathWithPrefixName].AcceptChanges();
            }

            return mapOfAllDataSet;
        }

        public static void exportToExcel(Dictionary<string, DataSet> source, string fileName)
        {
            try
            {
                if (File.Exists(fileName))
                {
                    File.Delete(fileName);
                }

                System.IO.StreamWriter excelDoc = new System.IO.StreamWriter(fileName);
                const string startExcelXML =
                    "<xml version>\r\n" +
                    "<Workbook xmlns=\"urn:schemas-microsoft-com:office:spreadsheet\"\r\n" +
                             " xmlns:o=\"urn:schemas-microsoft-com:office:office\"\r\n" +
                             " xmlns:x=\"urn:schemas-microsoft-com:office:excel\"\r\n" +
                             " xmlns:ss=\"urn:schemas-microsoft-com:office:spreadsheet\">\r\n" +
                        "<Styles>\r\n " +
                            "<Style ss:ID=\"Default\" ss:Name=\"Normal\">\r\n " +
                                "<Alignment ss:Vertical=\"Bottom\"/>\r\n" +
                                "<Borders/>\r\n" +
                                "<Font/>\r\n" +
                                "<Interior/>\r\n" +
                                "<NumberFormat/>\r\n" +
                                "<Protection/>\r\n" +
                            "</Style>\r\n" +
                            "<Style ss:ID=\"BoldColumn\">\r\n" +
                                "<Font x:Family=\"Swiss\" ss:Bold=\"1\"/>\r\n" +
                            "</Style>\r\n " +
                            "<Style ss:ID=\"StringLiteral\">\r\n" +
                                "<NumberFormat ss:Format=\"@\"/>\r\n" +
                            "</Style>\r\n" +
                            "<Style ss:ID=\"Decimal\">\r\n" +
                                "<NumberFormat ss:Format=\"0.0000\"/>\r\n" +
                            "</Style>\r\n" +
                            "<Style ss:ID=\"Integer\">\r\n" +
                                "<NumberFormat ss:Format=\"0\"/>\r\n" +
                            "</Style>\r\n" +
                            "<Style ss:ID=\"DateLiteral\">\r\n" +
                                "<NumberFormat ss:Format=\"mm/dd/yyyy;@\"/>\r\n" +
                            "</Style>\r\n" +
                        "</Styles>\r\n";
                const string endExcelXML = "</Workbook>";

                /*
                <xml version>
                <Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"
                    xmlns:o="urn:schemas-microsoft-com:office:office"
                    xmlns:x="urn:schemas-microsoft-com:office:excel"
                    xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet">
                    <Styles>
                        <Style ss:ID="Default" ss:Name="Normal">
                            <Alignment ss:Vertical="Bottom"/>
                            <Borders/>
                            <Font/>
                            <Interior/>
                            <NumberFormat/>
                            <Protection/>
                        </Style>
                        <Style ss:ID="BoldColumn">
                            <Font x:Family="Swiss" ss:Bold="1"/>
                        </Style>
                        <Style ss:ID="StringLiteral">
                            <NumberFormat ss:Format="@"/>
                        </Style>
                        <Style ss:ID="Decimal">
                            <NumberFormat ss:Format="0.0000"/>
                        </Style>
                        <Style ss:ID="Integer">
                            <NumberFormat ss:Format="0"/>
                        </Style>
                        <Style ss:ID="DateLiteral">
                            <NumberFormat ss:Format="mm/dd/yyyy;@"/>
                        </Style>
                    </Styles>

                    <Worksheet ss:Name="Sheet1">
                    </Worksheet>
                </Workbook>
               */
                excelDoc.Write(startExcelXML);

                foreach (KeyValuePair<string, DataSet> kv in source)
                {
                    writeWorkSheet(excelDoc, kv.Value.Tables[0]);
                }

                excelDoc.Write(endExcelXML);
                excelDoc.Close();
            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message, "ERROR");
                return;
            }
        }

        static int sheetCount = 0;
        static void writeWorkSheet(StreamWriter excelDoc, DataTable table)
        {
            int rowCount = 0;

            string sheetName = table.TableName.Substring(table.TableName.LastIndexOf("\\") + 1);

            excelDoc.Write("<Worksheet ss:Name=\"" + sheetName + sheetCount + "\">");
            excelDoc.Write("<Table>");
            excelDoc.Write("<Row>");
            for (int x = 0; x < table.Columns.Count; x++)
            {
                excelDoc.Write("<Cell ss:StyleID=\"BoldColumn\"><Data ss:Type=\"String\">");
                excelDoc.Write(table.Columns[x].ColumnName);
                excelDoc.Write("</Data></Cell>");
            }
            excelDoc.Write("</Row>\r\n");

            foreach (DataRow x in table.Rows)
            {
                rowCount++;
                //                 //if the number of rows is > 64000 create a new page to continue output
                //                 if (rowCount == 64000)
                //                 {
                //                     rowCount = 0;
                //                     sheetCount++;
                //                     excelDoc.Write("</Table>");
                //                     excelDoc.Write(" </Worksheet>");
                //                     excelDoc.Write("<Worksheet ss:Name=\"" + sheetName + "\">");
                //                     excelDoc.Write("<Table>");
                //                 }
                excelDoc.Write("<Row>"); //ID=" + rowCount + "
                for (int y = 0; y < table.Columns.Count; y++)
                {
                    System.Type rowType;
                    rowType = x[y].GetType();
                    switch (rowType.ToString())
                    {
                        case "System.String":
                            string XMLstring = x[y].ToString();
                            XMLstring = XMLstring.Trim();
                            XMLstring = XMLstring.Replace("&", "&");
                            XMLstring = XMLstring.Replace(">", ">");
                            XMLstring = XMLstring.Replace("<", "<");
                            excelDoc.Write("<Cell ss:StyleID=\"StringLiteral\">" +
                                            "<Data ss:Type=\"String\">");
                            excelDoc.Write(XMLstring);
                            excelDoc.Write("</Data></Cell>");
                            break;
                        case "System.DateTime":
                            //Excel has a specific Date Format of YYYY-MM-DD followed by  
                            //the letter 'T' then hh:mm:sss.lll Example 2005-01-31T24:01:21.000
                            //The Following Code puts the date stored in XMLDate 
                            //to the format above
                            DateTime XMLDate = (DateTime)x[y];
                            string XMLDatetoString = ""; //Excel Converted Date
                            XMLDatetoString = XMLDate.Year.ToString() +
                                    "-" +
                                    (XMLDate.Month < 10 ? "0" +
                                    XMLDate.Month.ToString() : XMLDate.Month.ToString()) +
                                    "-" +
                                    (XMLDate.Day < 10 ? "0" +
                                    XMLDate.Day.ToString() : XMLDate.Day.ToString()) +
                                    "T" +
                                    (XMLDate.Hour < 10 ? "0" +
                                    XMLDate.Hour.ToString() : XMLDate.Hour.ToString()) +
                                    ":" +
                                    (XMLDate.Minute < 10 ? "0" +
                                    XMLDate.Minute.ToString() : XMLDate.Minute.ToString()) +
                                    ":" +
                                    (XMLDate.Second < 10 ? "0" +
                                    XMLDate.Second.ToString() : XMLDate.Second.ToString()) +
                                    ".000";
                            excelDoc.Write("<Cell ss:StyleID=\"DateLiteral\">" +
                                            "<Data ss:Type=\"DateTime\">");
                            excelDoc.Write(XMLDatetoString);
                            excelDoc.Write("</Data></Cell>");
                            break;
                        case "System.Boolean":
                            excelDoc.Write("<Cell ss:StyleID=\"StringLiteral\">" +
                                        "<Data ss:Type=\"String\">");
                            excelDoc.Write(x[y].ToString());
                            excelDoc.Write("</Data></Cell>");
                            break;
                        case "System.Int16":
                        case "System.Int32":
                        case "System.Int64":
                        case "System.Byte":
                            excelDoc.Write("<Cell ss:StyleID=\"Integer\">" +
                                    "<Data ss:Type=\"Number\">");
                            excelDoc.Write(x[y].ToString());
                            excelDoc.Write("</Data></Cell>");
                            break;
                        case "System.Decimal":
                        case "System.Double":
                            excelDoc.Write("<Cell ss:StyleID=\"Decimal\">" +
                                    "<Data ss:Type=\"Number\">");
                            excelDoc.Write(x[y].ToString());
                            excelDoc.Write("</Data></Cell>");
                            break;
                        case "System.DBNull":
                            excelDoc.Write("<Cell ss:StyleID=\"StringLiteral\">" +
                                    "<Data ss:Type=\"String\">");
                            excelDoc.Write("");
                            excelDoc.Write("</Data></Cell>");
                            break;
                        default:
                            throw (new Exception(rowType.ToString() + " not handled."));
                    }
                }
                excelDoc.Write("</Row>\r\n");
            }
            excelDoc.Write("</Table>");
            excelDoc.Write(" </Worksheet>");
        }

        /* six language  xaml generate ,xaml path is intended for filling the missing translation part with default English */
        static void generateXamlResourceFileFromExcel(String fileName,String xamlPath,string srcExcelFileName)
        {
            Dictionary<string, string> languageMap = new Dictionary<string, string>();

            String fullPath = BASE_INPUT_PATH + @fileName;

            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook workbook = xlApp.Workbooks.Open(fullPath, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, 1, 0);
            //xlApp.Workbooks.Close();//关闭打开的文档 否则学号会显示科学计数法。 

            Excel.Worksheet workSheet;
            int iRowCnt = 0, iColCnt = 0, iBgnRow, iBgnCol;
            int iEWSCnt = workbook.Worksheets.Count;

            for (int m = 1; m <= iEWSCnt; m++)
            {
                workSheet = (Excel.Worksheet)workbook.Worksheets[m];
                iRowCnt = 0 + workSheet.UsedRange.Cells.Rows.Count;
                iColCnt = 0 + workSheet.UsedRange.Cells.Columns.Count;
                iBgnRow = (workSheet.UsedRange.Cells.Row > 1) ?
                                workSheet.UsedRange.Cells.Row - 1 : workSheet.UsedRange.Cells.Row;
                iBgnCol = (workSheet.UsedRange.Cells.Column > 1) ?
                                workSheet.UsedRange.Cells.Column - 1 : workSheet.UsedRange.Cells.Column;

                //ResourceDictionary res = new ResourceDictionary();
                //String usResDic = File.ReadAllText(xamlPath,Encoding.UTF8);

                /* load xaml document */
                //ResourceDictionary resDic;
                //using (FileStream fs = new FileStream(xamlPath, FileMode.Open, FileAccess.Read))
                //{
                //    resDic = (ResourceDictionary)XamlReader.Load(fs);
                //    DictionaryEntry[] dicHashmap = new DictionaryEntry[resDic.Count];
                //    resDic.CopyTo(dicHashmap, 0);
                //    //Console.Write(XamlWriter.Save(resDic));
                //}
                Dictionary<string, string> orgLanHashMap = null;
                if (!String.IsNullOrEmpty(srcExcelFileName))
                {
                    orgLanHashMap = getDefaultEnTranslationFromExcel(srcExcelFileName);
                }
                if (!Directory.Exists(@"output/"))
                {
                    Directory.CreateDirectory(@"output/");
                }

                int iMaxColumnNum = 1;
                while (!String.IsNullOrEmpty(((Excel.Range)workSheet.UsedRange.Cells[iBgnRow, iBgnCol + iMaxColumnNum]).Text.ToString()))
                {
                    iMaxColumnNum++;
                }
                iMaxColumnNum = iMaxColumnNum - 1;

                for (int k = 0; k < iMaxColumnNum; k++)
                {
                    String xamlFileName = "StringResource" + "."
                        + ((Excel.Range)workSheet.UsedRange.Cells[iBgnRow, iBgnCol + k + 1]).Text.ToString() + ".xaml";
                    StreamWriter smWriter = File.CreateText(@"output/" + xamlFileName);
                    smWriter.WriteLine(@"<ResourceDictionary xmlns=""http://schemas.microsoft.com/winfx/2006/xaml/presentation""");
                    smWriter.WriteLine(@" xmlns:sys=""clr-namespace:System;assembly=mscorlib"" ");
                    smWriter.WriteLine(@" xmlns:x=""http://schemas.microsoft.com/winfx/2006/xaml"">");

                    for (int i = iBgnRow; i < iRowCnt + iBgnRow; i++)
                    {
                        String key = ((Excel.Range)workSheet.UsedRange.Cells[i, iBgnCol]).Text.ToString();
                        String value = ((Excel.Range)workSheet.UsedRange.Cells[i, iBgnCol + k + 1]).Text.ToString();
                        if (!String.IsNullOrEmpty(value) && key != "ID")
                        {
                            languageMap.Add(key, value);
                        }
                    }

                    if (orgLanHashMap != null && orgLanHashMap.Count > 0)
                    {
                        foreach (String item in orgLanHashMap.Keys)
                        {
                            if (languageMap.ContainsKey(item))
                            {
                                String value = String.Empty;
                                languageMap.TryGetValue(item, out value);
                                smWriter.WriteLine(String.Format(@"<sys:String x:Key=""{0}"">{1}</sys:String>", item, value));
                            }
                            else
                            {
                                String value = String.Empty;
                                orgLanHashMap.TryGetValue(item, out value);
                                smWriter.WriteLine(String.Format(@"<sys:String x:Key=""{0}"">{1}</sys:String>", item, value));
                            }
                        }
                    }
                    else
                    {
                        foreach (String item in languageMap.Keys)
                        {
                            String value = String.Empty;
                            languageMap.TryGetValue(item, out value);
                            smWriter.WriteLine(String.Format(@"<sys:String x:Key=""{0}"">{1}</sys:String>", item, value));
                        }
                    }
                    smWriter.WriteLine(@"</ResourceDictionary>");
                    Console.WriteLine("Generating " + xamlFileName + "\t" + "The total key-value count is :" + languageMap.Count);
                    languageMap.Clear();
                    smWriter.Flush();
                    smWriter.Close();
                }

            }// end of parse excel

            workbook.Close();
            Console.WriteLine("Done!");
            //xlApp.Workbooks.Close();
        }

        private static Dictionary<string, string> getDefaultEnTranslationFromExcel(string srcExcelFileName)
        {
            Dictionary<string, string> languageMap = new Dictionary<string, string>();
            Excel.Application xlApp = new Excel.Application();

            string srcExcelPath = BASE_INPUT_PATH + @srcExcelFileName;

            Excel.Workbook standEnworkbook = xlApp.Workbooks.Open(srcExcelPath, 
                0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, 1, 0);

            //xlApp.Workbooks.Close();//关闭打开的文档 否则学号会显示科学计数法。 

            Excel.Worksheet workSheet;
            int iRowCnt = 0, iColCnt = 0, iBgnRow, iBgnCol;
            int iEWSCnt = standEnworkbook.Worksheets.Count;
            for (int m = 1; m <= iEWSCnt; m++)
            {
                workSheet = (Excel.Worksheet)standEnworkbook.Worksheets[m];
                iRowCnt = 0 + workSheet.UsedRange.Cells.Rows.Count;
                iColCnt = 0 + workSheet.UsedRange.Cells.Columns.Count;
                iBgnRow = (workSheet.UsedRange.Cells.Row > 1) ?
                                workSheet.UsedRange.Cells.Row - 1 : workSheet.UsedRange.Cells.Row;
                iBgnCol = (workSheet.UsedRange.Cells.Column > 1) ?
                                workSheet.UsedRange.Cells.Column - 1 : workSheet.UsedRange.Cells.Column;
                for (int i = iBgnRow; i < iRowCnt + iBgnRow; i++)
                {
                    String key = ((Excel.Range)workSheet.UsedRange.Cells[i, iBgnCol]).Text.ToString();
                    String value = ((Excel.Range)workSheet.UsedRange.Cells[i, iBgnCol + 1]).Text.ToString();
                    if (!String.IsNullOrEmpty(value) && key != "ID")
                    {
                        languageMap.Add(key, value);
                    }
                }
                Console.WriteLine("The original file key-value count is:" + languageMap.Count);
            }
            standEnworkbook.Close();
            return languageMap;
        }

        /* for the animation translation part */
        static void generateXmlDocFromExcel(String fileName)
        {
            Dictionary<string, string> languageMap = new Dictionary<string, string>();

            String fullPath = BASE_INPUT_PATH + @fileName;
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook workbook = xlApp.Workbooks.Open(fullPath, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, 1, 0);
            //xlApp.Workbooks.Close();//关闭打开的文档 否则学号会显示科学计数法。 

            Excel.Worksheet workSheet;
            int iRowCnt = 0, iColCnt = 0, iBgnRow, iBgnCol;
            //int iEWSCnt = workbook.Worksheets.Count;

            workSheet = (Excel.Worksheet)workbook.Worksheets[1];
            iRowCnt = 0 + workSheet.UsedRange.Cells.Rows.Count;
            iColCnt = 0 + workSheet.UsedRange.Cells.Columns.Count;
            iBgnRow = (workSheet.UsedRange.Cells.Row > 1) ?
                            workSheet.UsedRange.Cells.Row - 1 : workSheet.UsedRange.Cells.Row;
            iBgnCol = (workSheet.UsedRange.Cells.Column > 1) ?
                            workSheet.UsedRange.Cells.Column - 1 : workSheet.UsedRange.Cells.Column;

            int[] keyLocate = { 3,1,6,1,4,0,0 };

            int iMaxColumnNum = 1;
            while (!String.IsNullOrEmpty(((Excel.Range)workSheet.UsedRange.Cells[iBgnRow, iBgnCol + iMaxColumnNum]).Text.ToString()))
            {
                iMaxColumnNum++;
            }
            iMaxColumnNum = iMaxColumnNum - 1;

            for (int k = 0; k < iMaxColumnNum; k++)
            {
                XmlDocument xmlDoc = new XmlDocument();
                XmlDeclaration xmlDecl;
                XmlElement xmlRoot;
                xmlDecl = xmlDoc.CreateXmlDeclaration("1.0", "utf-8", null);
                xmlDoc.AppendChild(xmlDecl);

                xmlRoot = xmlDoc.CreateElement("configuration");
                xmlDoc.AppendChild(xmlRoot);
                XmlNode root = xmlDoc.SelectSingleNode("configuration");

                XmlElement xAttr = xmlDoc.CreateElement("Node");//创建一个<Node>节点 

                Dictionary<String, String> tempDictionary = new Dictionary<string, string>();
                int iLocation = 0;
                for (int i = iBgnRow + 1; i < iRowCnt + iBgnRow; i++)
                {
                    String key = ((Excel.Range)workSheet.UsedRange.Cells[i, iBgnCol]).Text.ToString();
                    String value = ((Excel.Range)workSheet.UsedRange.Cells[i, iBgnCol + k + 1]).Text.ToString();
                    String sMark = ((Excel.Range)workSheet.UsedRange.Cells[i, iBgnCol]).NoteText();
                    if (String.IsNullOrEmpty(key))
                    {
                        break;
                    }

                    if (String.IsNullOrEmpty(value))
                    {
                        if (i != iBgnRow + 1)
                        {
                            xAttr.SetAttribute("default", tempDictionary.ElementAt(keyLocate[iLocation]).Value);
                            root.AppendChild(xAttr);
                            tempDictionary.Clear();
                            iLocation ++;
                        }
                        xAttr = xmlDoc.CreateElement(key);
                    }
                    else
                    {
                        //xAttr.SetAttribute(key, value);
                        XmlElement childAttr = xmlDoc.CreateElement("add");
                        childAttr.SetAttribute("key", key);
                        childAttr.SetAttribute("value", value);
                        xAttr.AppendChild(childAttr);
                        tempDictionary.Add(key,value);
                    }
                }
                // for the last cycle

                xAttr.SetAttribute("default", tempDictionary.ElementAt(keyLocate[iLocation]).Value);
                root.AppendChild(xAttr);
                tempDictionary.Clear();

                iLocation = 0;
                String xmlFileName = @"DSAnimation_"
                        + ((Excel.Range)workSheet.UsedRange.Cells[iBgnRow, iBgnCol + k + 1]).Text.ToString() + ".xml";
                if (!Directory.Exists(@"output/"))
                {
                    Directory.CreateDirectory(@"output/");
                }
                xmlDoc.Save(@"output/" + xmlFileName);

                Console.WriteLine("Generating " + xmlFileName + "...");
            }

            workbook.Close();
            Console.WriteLine("Done!");
        }
    }
}