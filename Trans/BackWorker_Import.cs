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
        
        /* six language  xaml generate ,xaml path is intended for filling the missing translation part with default English */
       public static Boolean generateXamlResourceFileFromExcel(String fileName, String outXamlDirPath, string srcExcelFileName)
        {
            Dictionary<string, string> languageMap = new Dictionary<string, string>();

            String fullPath = fileName;

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
                if (!Directory.Exists(@outXamlDirPath))
                {
                    Directory.CreateDirectory(@outXamlDirPath);
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
                    StreamWriter smWriter = File.CreateText(@outXamlDirPath+"\\" + xamlFileName);
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

            return true;
        }


        private static Dictionary<string, string> getDefaultEnTranslationFromExcel(string srcExcelFileName)
        {
            Dictionary<string, string> languageMap = new Dictionary<string, string>();
            Excel.Application xlApp = new Excel.Application();

            string srcExcelPath = srcExcelFileName;

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

            String fullPath = fileName;
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

            int[] keyLocate = { 3, 1, 6, 1, 4, 0, 0 };

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
                            iLocation++;
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
                        tempDictionary.Add(key, value);
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
