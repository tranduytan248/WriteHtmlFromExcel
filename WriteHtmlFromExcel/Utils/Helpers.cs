using Newtonsoft.Json;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using WriteHtmlFromExcel.Models;

namespace CenIT.Report.Utils
{
    public class Helpers
    {
        public static int GetIndexInColumnExcel(string column)
        {
            List<string> character = new List<string> { "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z" };

            List<string> tmp = new List<string>();
            tmp.AddRange(character);

            foreach (string s in tmp)
            {
                foreach (string s1 in tmp)
                {
                    character.Add(s + s1);
                }
            }

            int idx = character.IndexOf(column);
            return idx + 1;
        }

        public static string GetString(string column)
        {
            List<string> character = new List<string> { "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z" };
            string cha = "";
            string num = "";
            for (int i = 0; i < column.Length; i++)
            {
                if (character.Contains(column[i].ToString()))
                {
                    cha += column[i].ToString();
                }
            }

            num = column.Replace(cha, "");

            int i1 = GetIndexInColumnExcel(cha);
            int i2 = Int32.Parse(num);

            return (i1 - 1).ToString() + "," + (i2 - 1).ToString();
        }

        public static string GetStyle(ExcelRange range)
        {
            ExcelStyle styleE = range.Style;
            string styleHTML = string.Empty;
            #region Font
            float fontSize = styleE.Font.Size;
            styleHTML += "font-size:" + fontSize + "px; ";

            string fontFamily = styleE.Font.Name;
            if (fontFamily == "Times New Roman")
            {
                styleHTML += "font-family:Times New Roman, Times, serif; ";
            }
            else
            {
                styleHTML += "font-family:" + fontFamily + ", Helvetica, sans-serif; ";
            }


            bool isItalic = styleE.Font.Italic;
            bool isBold = styleE.Font.Bold;
            bool isUnderline = styleE.Font.UnderLine;

            if (isBold && !isItalic)
            {
                styleHTML += "font-weight: bold; ";
            }
            if (isBold && isItalic)
            {
                styleHTML += "font-style: italic; font-weight: bold; ";
            }
            if (!isBold && isItalic)
            {
                styleHTML += "font-style: italic; ";
            }

            styleHTML += isUnderline ? "text-decoration: underline; " : "";
            #endregion

            #region Background and color
            string backgroundColor = styleE.Fill.BackgroundColor.Rgb;
            if (!string.IsNullOrEmpty(backgroundColor))
            {
                styleHTML += "background-color: #" + backgroundColor.Substring(2, 6) + "; ";
            }

            string color = styleE.Font.Color.Rgb;
            if (!string.IsNullOrEmpty(color))
            {
                styleHTML += "color: #" + color.Substring(2, 6) + "; ";
            }

            #endregion

            #region Align
            string HorizontalAlignment = styleE.HorizontalAlignment.ToString();
            styleHTML += "text-align:" + HorizontalAlignment + "; ";
            string VerticalAlignment = styleE.VerticalAlignment.ToString();
            string dataVertical = "";
            switch (VerticalAlignment)
            {
                case "Center":
                    dataVertical = "middle";
                    break;
                case "Top":
                    dataVertical = "top";
                    break;
                case "Bottom":
                    dataVertical = "bottom";
                    break;
                default:
                    dataVertical = "bottom";
                    break;
            }
            styleHTML += "vertical-align:" + dataVertical + "; ";
            #endregion

            #region Border
            ExcelBorderItem b = styleE.Border.Bottom;
            if (b.Style.ToString() != "None")
            {
                styleHTML += "border-bottom: 1px solid #000000; ";
            }
            ExcelBorderItem t = styleE.Border.Top;
            if (t.Style.ToString() != "None")
            {
                styleHTML += "border-top: 1px solid #000000; ";
            }
            ExcelBorderItem l = styleE.Border.Left;
            if (l.Style.ToString() != "None")
            {
                styleHTML += "border-left: 1px solid #000000; ";
            }
            ExcelBorderItem r = styleE.Border.Right;
            if (r.Style.ToString() != "None")
            {
                styleHTML += "border-right: 1px solid #000000; ";
            }
            #endregion

            return styleHTML;
        }

        public static string GenerateHTML(ExcelPackage p)
        {
            var ws = p.Workbook.Worksheets[1];
            string json = JsonConvert.SerializeObject(ws.Cells.Value);

            // Lấy danh sách từng ô trong excel
            string[,] kvalue = JsonConvert.DeserializeObject<string[,]>(json);

            // Lấy danh sách các ô bị merge
            var megers = ws.MergedCells;
            string json2 = JsonConvert.SerializeObject(megers);
            string[] mergerCells = JsonConvert.DeserializeObject<string[]>(json2);

            List<MergerModel> lsMerge = new List<MergerModel>();

            for (int i = 0; i < mergerCells.Length; i++)
            {
                string a = mergerCells[i];
                string[] sp = a.Split(':');
                string v1 = Helpers.GetString(sp[0]);
                string v2 = Helpers.GetString(sp[1]);

                MergerModel m = new MergerModel();
                m.colStart = Int32.Parse(v1.Split(',')[0]);
                m.rowStart = Int32.Parse(v1.Split(',')[1]);
                m.colEnd = Int32.Parse(v2.Split(',')[0]);
                m.rowEnd = Int32.Parse(v2.Split(',')[1]);
                m.infoMerge = Helpers.GetString(sp[0]) + ":" + Helpers.GetString(sp[1]);

                lsMerge.Add(m);
            }

            lsMerge = lsMerge.OrderBy(m => m.rowStart).ToList();

            string html = string.Empty;
            html += "<table style = 'width:100%'>";

            for (int r = 0; r < kvalue.GetLength(0); r++)
            {
                var heightRow = ws.Row(r + 1).Height;
                html += "<tr style='height: " + heightRow + "px;'>";
                List<MergerModel> lstM = lsMerge.Where(m => m.rowStart == r).ToList();
                if (lstM != null && lstM.Count > 0)
                {
                    lstM = lstM.OrderBy(m => m.rowStart).ToList();
                    for (int c = 0; c < kvalue.GetLength(1); c++)
                    {
                        var obj = lstM.Where(m => m.colStart == c).FirstOrDefault();
                        if (obj == null)
                        {
                            List<MergerModel> lstM2 = lsMerge.Where(m => m.rowStart < r && m.rowEnd >= r && m.colStart <= c && m.colEnd >= c).ToList();
                            if (lstM2 == null || lstM2.Count == 0)
                            {
                                ExcelRange range = ws.Cells[r + 1, c + 1];
                                html += "<td style='" + Helpers.GetStyle(range) + "'>";
                                html += kvalue[r, c];
                                html += "</td>";
                            }
                        }
                        else
                        {
                            if (obj.rowEnd == r)
                            {
                                ExcelRange range = ws.Cells[r + 1, c + 1];
                                html += "<td colspan='" + (obj.colEnd - obj.colStart + 1)
                                    + "' style='" + Helpers.GetStyle(range) + "'>";
                                html += kvalue[r, c];
                                html += "</td>";
                            }
                            else
                            {
                                ExcelRange range = ws.Cells[r + 1, c + 1];
                                html += "<td colspan='" + (obj.colEnd - obj.colStart + 1)
                                    + "' rowspan ='" + (obj.rowEnd - obj.rowStart + 1)
                                    + "' style='" + Helpers.GetStyle(range) + "'>";
                                html += kvalue[r, c];
                                html += "</td>";
                            }
                            c = obj.colEnd;
                        }
                    }
                }
                else
                    for (int c = 0; c < kvalue.GetLength(1); c++)
                    {
                        List<MergerModel> lstM2 = lsMerge.Where(m => m.rowStart < r && m.rowEnd >= r && m.colStart <= c && m.colEnd >= c).ToList();
                        if (lstM2 == null || lstM2.Count == 0)
                        {
                            ExcelRange range = ws.Cells[r + 1, c + 1];
                            html += "<td style='" + Helpers.GetStyle(range) + "'>";
                            html += kvalue[r, c];
                            html += "</td>";
                        }
                    }
                html += "</tr>";
            }
            html += "</table>";
            return html;
        }
    }
}