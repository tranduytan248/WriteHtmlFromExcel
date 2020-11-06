using CenIT.Report.Utils;
using Newtonsoft.Json;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using WriteHtmlFromExcel.Models;

namespace WriteHtmlFromExcel.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            string filename = "2_2_Huyen_Tinh_Import.xlsx";
            string URLServerFile = Server.MapPath("~/TemplateFiles/");
            var filePath = Path.Combine(URLServerFile + "\\" + filename);

            var file = new System.IO.FileInfo(filePath);
            using (ExcelPackage p = new ExcelPackage(file))
            {
                var ws = p.Workbook.Worksheets[1];

                ViewBag.HTMLGenerate = GenerateHTML(p);
            }

            return View();
        }

        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }


        private string GenerateHTML(ExcelPackage p)
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