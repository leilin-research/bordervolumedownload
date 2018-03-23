using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Net;
using System.Text;
using System.Windows.Forms;
using System.Text.RegularExpressions;
using System.Reflection;
using Excel = Microsoft.Office.Interop.Excel;

namespace borderwaitingvolumedownload
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        public string GetHTML(string url, string encoding)   //得到html；
        {
            WebClient web = new WebClient();
            byte[] buffer = web.DownloadData(url);
            return Encoding.GetEncoding(encoding).GetString(buffer);
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            int d;
            int m;
            int y;
            d = 1;
            m = 1;
            y = 2003;
            string sd;
            string sm;
            string sy;
            string url;
            string result;
            string dd;
            string ndd;


            for (y = 2011; y <= 2013; y++)// which year to start
            {
                for (m = 1; m <= 12; m++) // which month to start
                {
                    if (m == 1 || m == 3 || m == 5 || m == 7 || m == 8 || m == 10 || m == 12)
                    {
                        for (d = 1; d <= 31; d++)
                        {
                            sd = Convert.ToString(d);
                            sm = Convert.ToString(m);
                            sy = Convert.ToString(y);
                            DataTable dtDay = new DataTable(sm + "-" + sd + "-" + sy);
                            dtDay.Columns.Add("mm-dd-yyyy", System.Type.GetType("System.String"));
                            dtDay.Columns.Add("H", System.Type.GetType("System.String"));
                            dtDay.Columns.Add("Auto1", System.Type.GetType("System.String"));
                            dtDay.Columns.Add("Truck1", System.Type.GetType("System.String"));
                            dtDay.Columns.Add("Auto2", System.Type.GetType("System.String"));
                            dtDay.Columns.Add("Truck2", System.Type.GetType("System.String"));
                            dtDay.Columns.Add("Bus", System.Type.GetType("System.String"));
                            dtDay.Columns.Add("Total", System.Type.GetType("System.String"));
                            url = "http://www.peacebridge.com/traffic.php?month=" + sm + "&day=" + sd + "&year=" + sy + "&view=daily&process=View";
                            result = GetHTML(url, "gb2312");

                            Regex check = new Regex("<td\\s+align=right>\\n\\s\\s+<b>\\d\\d*/\\d\\d*/\\d{4}</b>\\n\\s\\s+</td>\\n\\s\\s+<td\\s+align=right>"
                                + "\\n\\s\\s+\\d{1,2}\\n\\s\\s+</td>\\n\\s\\s+<td\\s+align=right>\\n\\s\\s+\\d{1,3}\\n\\s\\s+</td>\\n\\s\\s+<td\\s+align=right>"
                                + "\\n\\s\\s+\\d{1,3}\\n\\s\\s+</td>\\n\\s\\s+<td\\s+align=right>\\n\\s\\s+\\d{1,3}\\n\\s\\s+</td>\\n\\s\\s+<td\\s+align=right>\\n\\s\\s+\\d{1,3}\\n\\s\\s+</td>\\n\\s\\s+<td\\s+align=right>\\n\\s\\s+\\d{1,3}\\n\\s\\s+</td>\\n\\s\\s+<td\\s+align=right>\\n\\s\\s+<b>\\d?,?\\d{1,3}</b>\\n\\s\\s+</td>");
                            MatchCollection marticles = check.Matches(result);
                            foreach (Match mar in marticles)
                            {
                                dd = mar.ToString();
                                ndd = dd.Replace("<td align=right>\n\t\t", " ");
                                ndd = ndd.Replace("</td>", "");
                                ndd = ndd.Replace("<b>", "");
                                ndd = ndd.Replace("</b>", "");
                                ndd = ndd.Replace("\n\t\t", "");
                                char[] charsToTrim = { ' ' };
                                ndd = ndd.Trim();
                                string temp = "";
                                int count = 0;
                                for (int i = 0; i < ndd.Length; i++)  // 只保留一个空格
                                {
                                    if (ndd[i] != ' ')
                                    {
                                        count = 0;
                                        temp += ndd[i];
                                    }
                                    if (ndd[i] == ' ' && count == 0)
                                    {
                                        temp += ndd[i];
                                        count = 1;
                                    }
                                }
                                string[] mdd = temp.Split(' ');
                                DataRow dr = dtDay.NewRow();
                                dr["mm-dd-yyyy"] = mdd[0];
                                dr["H"] = mdd[1];
                                dr["Auto1"] = mdd[2];
                                dr["Truck1"] = mdd[3];
                                dr["Auto2"] = mdd[4];
                                dr["Truck2"] = mdd[5];
                                dr["Bus"] = mdd[6];
                                dr["Total"] = mdd[7];
                                dtDay.Rows.Add(dr);
                            }
                            Excel.Application exData = new Excel.Application();
                            exData.Workbooks.Add(true);
                            int row = 2;
                            for (int i = 0; i < dtDay.Columns.Count; i++)
                            {
                                exData.Cells[1, i + 1] = dtDay.Columns[i].ColumnName.ToString();
                            }
                            for (int i = 0; i < dtDay.Rows.Count; i++)
                            {
                                for (int j = 0; j < dtDay.Columns.Count; j++)
                                {
                                    exData.Cells[row, j + 1] = dtDay.Rows[i][j].ToString();
                                }
                                row++;
                            }
                            // exData.Visible = true;

                            foreach (Excel.Workbook wkb in exData.Workbooks)
                            {
                                wkb.SaveAs(@"D:\study\bordercrossingprediction\data\2003\" + sm + "-" + sd + "-" + sy + ".xlsx", Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                            }
                            exData.Quit();
                            exData = null;
                        }
                    }
                    else if (m == 2 && (y != 2004 && y != 2008))
                    {
                        for (d = 1; d <= 28; d++)
                        {
                            sd = Convert.ToString(d);
                            sm = Convert.ToString(m);
                            sy = Convert.ToString(y);
                            DataTable dtDay = new DataTable(sm + "-" + sd + "-" + sy);
                            dtDay.Columns.Add("mm-dd-yyyy", System.Type.GetType("System.String"));
                            dtDay.Columns.Add("H", System.Type.GetType("System.String"));
                            dtDay.Columns.Add("Auto1", System.Type.GetType("System.String"));
                            dtDay.Columns.Add("Truck1", System.Type.GetType("System.String"));
                            dtDay.Columns.Add("Auto2", System.Type.GetType("System.String"));
                            dtDay.Columns.Add("Truck2", System.Type.GetType("System.String"));
                            dtDay.Columns.Add("Bus", System.Type.GetType("System.String"));
                            dtDay.Columns.Add("Total", System.Type.GetType("System.String"));
                            url = "http://www.peacebridge.com/traffic.php?month=" + sm + "&day=" + sd + "&year=" + sy + "&view=daily&process=View";
                            result = GetHTML(url, "gb2312");

                            Regex check = new Regex("<td\\s+align=right>\\n\\s\\s+<b>\\d\\d*/\\d\\d*/\\d{4}</b>\\n\\s\\s+</td>\\n\\s\\s+<td\\s+align=right>"
                                + "\\n\\s\\s+\\d{1,2}\\n\\s\\s+</td>\\n\\s\\s+<td\\s+align=right>\\n\\s\\s+\\d{1,3}\\n\\s\\s+</td>\\n\\s\\s+<td\\s+align=right>"
                                + "\\n\\s\\s+\\d{1,3}\\n\\s\\s+</td>\\n\\s\\s+<td\\s+align=right>\\n\\s\\s+\\d{1,3}\\n\\s\\s+</td>\\n\\s\\s+<td\\s+align=right>\\n\\s\\s+\\d{1,3}\\n\\s\\s+</td>\\n\\s\\s+<td\\s+align=right>\\n\\s\\s+\\d{1,3}\\n\\s\\s+</td>\\n\\s\\s+<td\\s+align=right>\\n\\s\\s+<b>\\d?,?\\d{1,3}</b>\\n\\s\\s+</td>");
                            MatchCollection marticles = check.Matches(result);
                            foreach (Match mar in marticles)
                            {
                                dd = mar.ToString();
                                ndd = dd.Replace("<td align=right>\n\t\t", " ");
                                ndd = ndd.Replace("</td>", "");
                                ndd = ndd.Replace("<b>", "");
                                ndd = ndd.Replace("</b>", "");
                                ndd = ndd.Replace("\n\t\t", "");
                                char[] charsToTrim = { ' ' };
                                ndd = ndd.Trim();
                                string temp = "";
                                int count = 0;
                                for (int i = 0; i < ndd.Length; i++)  // 只保留一个空格
                                {
                                    if (ndd[i] != ' ')
                                    {
                                        count = 0;
                                        temp += ndd[i];
                                    }
                                    if (ndd[i] == ' ' && count == 0)
                                    {
                                        temp += ndd[i];
                                        count = 1;
                                    }
                                }
                                string[] mdd = temp.Split(' ');
                                DataRow dr = dtDay.NewRow();
                                dr["mm-dd-yyyy"] = mdd[0];
                                dr["H"] = mdd[1];
                                dr["Auto1"] = mdd[2];
                                dr["Truck1"] = mdd[3];
                                dr["Auto2"] = mdd[4];
                                dr["Truck2"] = mdd[5];
                                dr["Bus"] = mdd[6];
                                dr["Total"] = mdd[7];
                                dtDay.Rows.Add(dr);
                            }
                            Excel.Application exData = new Excel.Application();
                            exData.Workbooks.Add(true);
                            int row = 2;
                            for (int i = 0; i < dtDay.Columns.Count; i++)
                            {
                                exData.Cells[1, i + 1] = dtDay.Columns[i].ColumnName.ToString();
                            }
                            for (int i = 0; i < dtDay.Rows.Count; i++)
                            {
                                for (int j = 0; j < dtDay.Columns.Count; j++)
                                {
                                    exData.Cells[row, j + 1] = dtDay.Rows[i][j].ToString();
                                }
                                row++;
                            }
                            // exData.Visible = true;

                            foreach (Excel.Workbook wkb in exData.Workbooks)
                            {
                                wkb.SaveAs(@"D:\study\bordercrossingprediction\data\2003\" + sm + "-" + sd + "-" + sy + ".xlsx", Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                            }
                            exData.Quit();
                            exData = null;
                        }
                    }
                    else if (m == 2 && (y == 2004 || y == 2008))
                    {
                        for (d = 1; d <= 29; d++)
                        {
                            sd = Convert.ToString(d);
                            sm = Convert.ToString(m);
                            sy = Convert.ToString(y);
                            DataTable dtDay = new DataTable(sm + "-" + sd + "-" + sy);
                            dtDay.Columns.Add("mm-dd-yyyy", System.Type.GetType("System.String"));
                            dtDay.Columns.Add("H", System.Type.GetType("System.String"));
                            dtDay.Columns.Add("Auto1", System.Type.GetType("System.String"));
                            dtDay.Columns.Add("Truck1", System.Type.GetType("System.String"));
                            dtDay.Columns.Add("Auto2", System.Type.GetType("System.String"));
                            dtDay.Columns.Add("Truck2", System.Type.GetType("System.String"));
                            dtDay.Columns.Add("Bus", System.Type.GetType("System.String"));
                            dtDay.Columns.Add("Total", System.Type.GetType("System.String"));
                            url = "http://www.peacebridge.com/traffic.php?month=" + sm + "&day=" + sd + "&year=" + sy + "&view=daily&process=View";
                            result = GetHTML(url, "gb2312");

                            Regex check = new Regex("<td\\s+align=right>\\n\\s\\s+<b>\\d\\d*/\\d\\d*/\\d{4}</b>\\n\\s\\s+</td>\\n\\s\\s+<td\\s+align=right>"
                                + "\\n\\s\\s+\\d{1,2}\\n\\s\\s+</td>\\n\\s\\s+<td\\s+align=right>\\n\\s\\s+\\d{1,3}\\n\\s\\s+</td>\\n\\s\\s+<td\\s+align=right>"
                                + "\\n\\s\\s+\\d{1,3}\\n\\s\\s+</td>\\n\\s\\s+<td\\s+align=right>\\n\\s\\s+\\d{1,3}\\n\\s\\s+</td>\\n\\s\\s+<td\\s+align=right>\\n\\s\\s+\\d{1,3}\\n\\s\\s+</td>\\n\\s\\s+<td\\s+align=right>\\n\\s\\s+\\d{1,3}\\n\\s\\s+</td>\\n\\s\\s+<td\\s+align=right>\\n\\s\\s+<b>\\d?,?\\d{1,3}</b>\\n\\s\\s+</td>");
                            MatchCollection marticles = check.Matches(result);
                            foreach (Match mar in marticles)
                            {
                                dd = mar.ToString();
                                ndd = dd.Replace("<td align=right>\n\t\t", " ");
                                ndd = ndd.Replace("</td>", "");
                                ndd = ndd.Replace("<b>", "");
                                ndd = ndd.Replace("</b>", "");
                                ndd = ndd.Replace("\n\t\t", "");
                                char[] charsToTrim = { ' ' };
                                ndd = ndd.Trim();
                                string temp = "";
                                int count = 0;
                                for (int i = 0; i < ndd.Length; i++)  // 只保留一个空格
                                {
                                    if (ndd[i] != ' ')
                                    {
                                        count = 0;
                                        temp += ndd[i];
                                    }
                                    if (ndd[i] == ' ' && count == 0)
                                    {
                                        temp += ndd[i];
                                        count = 1;
                                    }
                                }
                                string[] mdd = temp.Split(' ');
                                DataRow dr = dtDay.NewRow();
                                dr["mm-dd-yyyy"] = mdd[0];
                                dr["H"] = mdd[1];
                                dr["Auto1"] = mdd[2];
                                dr["Truck1"] = mdd[3];
                                dr["Auto2"] = mdd[4];
                                dr["Truck2"] = mdd[5];
                                dr["Bus"] = mdd[6];
                                dr["Total"] = mdd[7];
                                dtDay.Rows.Add(dr);
                            }
                            Excel.Application exData = new Excel.Application();
                            exData.Workbooks.Add(true);
                            int row = 2;
                            for (int i = 0; i < dtDay.Columns.Count; i++)
                            {
                                exData.Cells[1, i + 1] = dtDay.Columns[i].ColumnName.ToString();
                            }
                            for (int i = 0; i < dtDay.Rows.Count; i++)
                            {
                                for (int j = 0; j < dtDay.Columns.Count; j++)
                                {
                                    exData.Cells[row, j + 1] = dtDay.Rows[i][j].ToString();
                                }
                                row++;
                            }
                            // exData.Visible = true;

                            foreach (Excel.Workbook wkb in exData.Workbooks)
                            {
                                wkb.SaveAs(@"D:\study\bordercrossingprediction\data\2003\" + sm + "-" + sd + "-" + sy + ".xlsx", Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                            }
                            exData.Quit();
                            exData = null;
                        }
                    }
                    else if (m == 4 || m == 6 || m == 9 || m == 11)
                    {
                        for (d = 1; d <= 30; d++)
                        {
                            sd = Convert.ToString(d);
                            sm = Convert.ToString(m);
                            sy = Convert.ToString(y);
                            DataTable dtDay = new DataTable(sm + "-" + sd + "-" + sy);
                            dtDay.Columns.Add("mm-dd-yyyy", System.Type.GetType("System.String"));
                            dtDay.Columns.Add("H", System.Type.GetType("System.String"));
                            dtDay.Columns.Add("Auto1", System.Type.GetType("System.String"));
                            dtDay.Columns.Add("Truck1", System.Type.GetType("System.String"));
                            dtDay.Columns.Add("Auto2", System.Type.GetType("System.String"));
                            dtDay.Columns.Add("Truck2", System.Type.GetType("System.String"));
                            dtDay.Columns.Add("Bus", System.Type.GetType("System.String"));
                            dtDay.Columns.Add("Total", System.Type.GetType("System.String"));
                            url = "http://www.peacebridge.com/traffic.php?month=" + sm + "&day=" + sd + "&year=" + sy + "&view=daily&process=View";
                            result = GetHTML(url, "gb2312");

                            Regex check = new Regex("<td\\s+align=right>\\n\\s\\s+<b>\\d\\d*/\\d\\d*/\\d{4}</b>\\n\\s\\s+</td>\\n\\s\\s+<td\\s+align=right>"
                                + "\\n\\s\\s+\\d{1,2}\\n\\s\\s+</td>\\n\\s\\s+<td\\s+align=right>\\n\\s\\s+\\d{1,3}\\n\\s\\s+</td>\\n\\s\\s+<td\\s+align=right>"
                                + "\\n\\s\\s+\\d{1,3}\\n\\s\\s+</td>\\n\\s\\s+<td\\s+align=right>\\n\\s\\s+\\d{1,3}\\n\\s\\s+</td>\\n\\s\\s+<td\\s+align=right>\\n\\s\\s+\\d{1,3}\\n\\s\\s+</td>\\n\\s\\s+<td\\s+align=right>\\n\\s\\s+\\d{1,3}\\n\\s\\s+</td>\\n\\s\\s+<td\\s+align=right>\\n\\s\\s+<b>\\d?,?\\d{1,3}</b>\\n\\s\\s+</td>");
                            MatchCollection marticles = check.Matches(result);
                            foreach (Match mar in marticles)
                            {
                                dd = mar.ToString();
                                ndd = dd.Replace("<td align=right>\n\t\t", " ");
                                ndd = ndd.Replace("</td>", "");
                                ndd = ndd.Replace("<b>", "");
                                ndd = ndd.Replace("</b>", "");
                                ndd = ndd.Replace("\n\t\t", "");
                                char[] charsToTrim = { ' ' };
                                ndd = ndd.Trim();
                                string temp = "";
                                int count = 0;
                                for (int i = 0; i < ndd.Length; i++)  // 只保留一个空格
                                {
                                    if (ndd[i] != ' ')
                                    {
                                        count = 0;
                                        temp += ndd[i];
                                    }
                                    if (ndd[i] == ' ' && count == 0)
                                    {
                                        temp += ndd[i];
                                        count = 1;
                                    }
                                }
                                string[] mdd = temp.Split(' ');
                                DataRow dr = dtDay.NewRow();
                                dr["mm-dd-yyyy"] = mdd[0];
                                dr["H"] = mdd[1];
                                dr["Auto1"] = mdd[2];
                                dr["Truck1"] = mdd[3];
                                dr["Auto2"] = mdd[4];
                                dr["Truck2"] = mdd[5];
                                dr["Bus"] = mdd[6];
                                dr["Total"] = mdd[7];
                                dtDay.Rows.Add(dr);
                            }
                            Excel.Application exData = new Excel.Application();
                            exData.Workbooks.Add(true);
                            int row = 2;
                            for (int i = 0; i < dtDay.Columns.Count; i++)
                            {
                                exData.Cells[1, i + 1] = dtDay.Columns[i].ColumnName.ToString();
                            }
                            for (int i = 0; i < dtDay.Rows.Count; i++)
                            {
                                for (int j = 0; j < dtDay.Columns.Count; j++)
                                {
                                    exData.Cells[row, j + 1] = dtDay.Rows[i][j].ToString();
                                }
                                row++;
                            }
                            // exData.Visible = true;

                            foreach (Excel.Workbook wkb in exData.Workbooks)
                            {
                                wkb.SaveAs(@"D:\study\bordercrossingprediction\data\2003\" + sm + "-" + sd + "-" + sy + ".xlsx", Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                            }
                            exData.Quit();
                            exData = null;
                        }
                    }
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {

        }
    }
}
