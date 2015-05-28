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


namespace Trans
{
    public partial class BackWorker
    {
         const string DefaultLang = "__DEFAULT__";
        
            public static Boolean exportToExcel(string inputDirPath,string outputFileName)
            {

                Dictionary<string, DataSet> source = loadAllResources(inputDirPath);
                string fileName = outputFileName;
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
                    return false;
                }

                return true;
            }

        private static Dictionary<string, DataSet> loadAllResources(string path)
        {
            string pattern = "StringResource";
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

    }
}
