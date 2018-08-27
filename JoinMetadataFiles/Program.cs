using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Text.RegularExpressions;
using System.Security.Permissions;
using System.IO;
using System.Configuration;
using System.Collections;
using System.Data;
using System.Net;
using System.Net.Http;
using Microsoft.Office.Interop.Excel;

namespace JoinMetadataFiles
{
    class Program
    {
        public static System.Data.DataTable ConvertCSVtoDataTable(string strFilePath)
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            using (StreamReader sr = new StreamReader(strFilePath))
            {
                string[] headers = sr.ReadLine().Split(',');
                foreach (string header in headers)
                {
                    dt.Columns.Add(header);
                }
                while (!sr.EndOfStream)
                {
                    try
                    {
                        string[] rows = sr.ReadLine().Split(',');
                        DataRow dr = dt.NewRow();
                        for (int i = 0; i < headers.Length - 1; i++)
                        {
                            try
                            {
                                dr[i] = rows[i];
                            }
                            catch { }
                        }
                        dt.Rows.Add(dr);
                    }
                    catch { }
                }

            }


            return dt;
        }
        static void Main(string[] args)
        {
            Application xlApp = null;
            Workbook xlWorkbook;
            Worksheet xlWorksheet;
            Range xlRange;
            DateTime startTime = DateTime.Now;

            System.Data.DataTable dt1 = ConvertCSVtoDataTable(@"C:\Users\byacoube\Documents\tr.csv");
            StringBuilder sb = new StringBuilder();
            System.Data.DataTable dt2 = null;// ConvertCSVtoDataTable(@"C:\Users\byacoube\Documents\om.txt");
            int i = 0;
            //StringBuilder sb = new StringBuilder();
            foreach (DataRow dr1 in dt1.Rows)
            {
                string xlsxFilePath = @"e:" + dr1["File Name"].ToString().ToLower().Replace("/", @"\");
                string prevrowPath = "";
                if (i >= 1)
                    prevrowPath = dt1.Rows[i - 1]["File Name"].ToString().ToLower();
                if (dr1["File Name"].ToString().ToLower() == prevrowPath + ".xlsx")
                {
                    Console.WriteLine(i.ToString() + " - processing file " + xlsxFilePath);

                    xlApp = new Application();
                    xlWorkbook = xlApp.Workbooks.Open(xlsxFilePath);
                    xlWorksheet = xlWorkbook.Sheets[1];

                    xlRange = xlWorksheet.UsedRange;
                    xlRange.Replace(",", "");
                    xlWorkbook.SaveAs(@"c:\csvs\" + i.ToString() + ".csv", XlFileFormat.xlCSV, null, null, null, null, XlSaveAsAccessMode.xlExclusive, null, null, null, null);
                    Console.WriteLine("saved file as csv " + xlsxFilePath + " as " + i.ToString() + ".csv");
                    xlWorkbook.Close();
                    xlApp.Quit();
                    
                    string csvPath = @"c:\csvs\" + i.ToString() + ".csv";


                    prevrowPath = prevrowPath.Substring(0, prevrowPath.LastIndexOf("/"));
                    //prevrowPath = prevrowPath.Substring(0, prevrowPath.LastIndexOf("/"));

                    //dt2 = ConvertCSVtoDataTable(csvPath);
                    sb.Clear();
                    StreamReader sr = new StreamReader(csvPath);
                    string line; int r = 0;
                    while (sr.Peek() >= 0)
                    {
                        r++;
                        //read a line in the CSV file
                        line = sr.ReadLine();
                        if (line.Trim().Length > 0)
                        {
                            line = Encoding.ASCII.GetString(Encoding.ASCII.GetBytes(line));
                            // if it is not a header, i.e not start of a new table, add a new row to target
                            if (r != 1 && !line.StartsWith("SourcePath"))
                            {
                                 line = prevrowPath + "/"+line;
                            }

                            sb.Append(line);
                            sb.Append(@"
");

                        }
                    }

                    StreamWriter sw = new StreamWriter(@"c:\csvs\_csv" + i.ToString() + ".csv", false);
                    sw.Write(sb.ToString());


                    sw.Flush();
                    sw.Close();



                }

                i++;
                /*
                Console.WriteLine("start time " + DateTime.Now.ToShortTimeString());
                foreach (DataRow dr1 in dt1.Rows)
                    foreach (DataRow dr2 in dt2.Rows)
                    {
                        try
                        {
                            Console.WriteLine("in table1, table2: " + dr1[0].ToString() + ", " + dr2[0].ToString());
                            if (!string.IsNullOrEmpty(dr2["SourcePath"].ToString().Trim()))
                            {
                                if (dr1["File Name"].ToString().Contains(dr2["SourcePath"].ToString()))
                                {
                                    sb.Append(dr1[0].ToString());
                                    sb.Append(",");

                                    sb.Append(dr2[0].ToString());
                                    sw.Write(sb.ToString());
                                    sb.Clear();
                                    break;
                                }
                            }
                        }
                        catch { }
                    }



                sw.Flush();
                sw.Close();
                */
            }
        }
    }
}
