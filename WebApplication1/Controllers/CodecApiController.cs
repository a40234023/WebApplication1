using Newtonsoft.Json;
using NLog;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using WebApplication1.Models;

namespace WebApplication1.Controllers
{
    public class CodecApiController : Controller
    {
        private static Logger logger = LogManager.GetCurrentClassLogger();

        [HttpPost]
        public ActionResult EncryptString(string template, string Detail)
        {
            List<Dictionary<string, string>> zDetail = JsonConvert.DeserializeObject<List<Dictionary<string, string>>>(Detail);
            return View(JsonConvert.SerializeObject(zDetail));
        }

        [HttpGet]
        public ActionResult GetFile(string name, string data)
        {
            //建立Excel
            logger.Debug("OK; name->" + name);
            logger.Debug("data->" + data);
            ApiResult AResult;
            string filesname = saveExcel(data);
            if (filesname != "")
            {
                logger.Debug("filesname-> " + filesname);
                AResult = new ApiResult
                {
                    code = "200",
                    name = filesname
                };
            }
            else
            {
                AResult = new ApiResult
                {
                    code = "500",
                    name = filesname
                };
            }
            return Json(AResult, JsonRequestBehavior.AllowGet);
            //                        return File(excelDatas.ToArray(), "application/vnd.ms-excel", string.Format($"學生資料.xls"));
            //            return Encoding.Default.GetString((excelDatas.ToArray()));
        }

        private string saveExcel(string Detail)
        {
            try
            {
                List<Dictionary<string, string>> zDetail = JsonConvert.DeserializeObject<List<Dictionary<string, string>>>(Detail);
                logger.Debug("收到的Detail:'" + Detail);

                XSSFWorkbook hssfworkbook = new XSSFWorkbook(); //建立活頁簿
                ISheet sheet = hssfworkbook.CreateSheet("sheet"); //建立sheet

                //設定樣式
                ICellStyle headerStyle = hssfworkbook.CreateCellStyle();
                IFont headerfont = hssfworkbook.CreateFont();
                headerStyle.Alignment = HorizontalAlignment.Center; //水平置中
                headerStyle.VerticalAlignment = VerticalAlignment.Center; //垂直置中
                headerfont.FontName = "微軟正黑體";
                headerfont.FontHeightInPoints = 20;
                headerfont.Boldweight = (short)FontBoldWeight.Bold;
                headerStyle.SetFont(headerfont);

                //新增標題列
                sheet.CreateRow(0); //需先用CreateRow建立,才可通过GetRow取得該欄位
                sheet.AddMergedRegion(new CellRangeAddress(0, 0, 0, 10)); //合併1~1列及A~K欄儲存格
                sheet.GetRow(0).CreateCell(0).SetCellValue("A9000200 全傳(蘇洲)-海運明細");
                sheet.GetRow(0).GetCell(0).CellStyle = headerStyle;
                sheet.CreateRow(1);
                sheet.AddMergedRegion(new CellRangeAddress(1, 1, 0, 10));                
                sheet.GetRow(1).CreateCell(0).SetCellValue("請於06/30下午4點前封箱");
                sheet.GetRow(1).GetCell(0).CellStyle = headerStyle; //套用樣式
                sheet.CreateRow(4);
                sheet.GetRow(4).CreateCell(0).SetCellValue("製表日:");
                sheet.GetRow(4).CreateCell(1).SetCellValue(DateTime.Now.ToString("yyyy/MM/dd"));

                IDataFormat dataformat = hssfworkbook.CreateDataFormat();

                ICellStyle cs_center = hssfworkbook.CreateCellStyle();
                cs_center.BorderBottom = BorderStyle.Thin;
                cs_center.BorderTop = BorderStyle.Hair;
                cs_center.BorderLeft = BorderStyle.Medium;
                cs_center.BorderRight = BorderStyle.Dotted;
                cs_center.Alignment= HorizontalAlignment.Center;

                ICellStyle cs_left = hssfworkbook.CreateCellStyle();
                cs_left.BorderBottom = BorderStyle.Thin;
                cs_left.BorderTop = BorderStyle.Hair;
                cs_left.BorderLeft = BorderStyle.Medium;
                cs_left.BorderRight = BorderStyle.Dotted;
                cs_left.Alignment = HorizontalAlignment.Left;

                ICellStyle cs_right = hssfworkbook.CreateCellStyle();
                cs_right.BorderBottom = BorderStyle.Thin;
                cs_right.BorderTop = BorderStyle.Hair;
                cs_right.BorderLeft = BorderStyle.Medium;
                cs_right.BorderRight = BorderStyle.Dotted;
                cs_right.Alignment = HorizontalAlignment.Right;
                cs_right.DataFormat = dataformat.GetFormat("#,##0");

                ICellStyle cs_right2 = hssfworkbook.CreateCellStyle();
                cs_right2.BorderBottom = BorderStyle.Thin;
                cs_right2.BorderTop = BorderStyle.Hair;
                cs_right2.BorderLeft = BorderStyle.Medium;
                cs_right2.BorderRight = BorderStyle.Dotted;
                cs_right2.Alignment = HorizontalAlignment.Right;
                cs_right2.DataFormat = dataformat.GetFormat("#,##0.00");

                XSSFFont font0 = (XSSFFont)hssfworkbook.CreateFont();
                font0.FontName = "微軟正黑體";
                font0.FontHeightInPoints = 12;
                font0.Boldweight = (short)FontBoldWeight.Bold;

                XSSFFont font1 = (XSSFFont)hssfworkbook.CreateFont();
                font1.FontName = "微軟正黑體";
                font1.FontHeightInPoints = 12;


                sheet.CreateRow(5).CreateCell(0).SetCellValue("序號");
                sheet.GetRow(5).CreateCell(1).SetCellValue("銷售文件");
                sheet.GetRow(5).CreateCell(2).SetCellValue("項目");
                sheet.GetRow(5).CreateCell(3).SetCellValue("物料");
                sheet.GetRow(5).CreateCell(4).SetCellValue("物料說明");
                sheet.GetRow(5).CreateCell(5).SetCellValue("客戶物料號碼(主檔)");
                sheet.GetRow(5).CreateCell(6).SetCellValue("數量");
                sheet.GetRow(5).CreateCell(7).SetCellValue("單價");
                sheet.GetRow(5).CreateCell(8).SetCellValue("小計");
                sheet.GetRow(5).CreateCell(9).SetCellValue("達交日");
                sheet.GetRow(5).CreateCell(10).SetCellValue("備註");

                sheet.GetRow(5).GetCell(0).CellStyle = cs_center;
                sheet.GetRow(5).GetCell(1).CellStyle = cs_center;
                sheet.GetRow(5).GetCell(2).CellStyle = cs_center;
                sheet.GetRow(5).GetCell(3).CellStyle = cs_center;
                sheet.GetRow(5).GetCell(4).CellStyle = cs_center;
                sheet.GetRow(5).GetCell(5).CellStyle = cs_center;
                sheet.GetRow(5).GetCell(6).CellStyle = cs_center;
                sheet.GetRow(5).GetCell(7).CellStyle = cs_center;
                sheet.GetRow(5).GetCell(8).CellStyle = cs_center;
                sheet.GetRow(5).GetCell(9).CellStyle = cs_center;
                sheet.GetRow(5).GetCell(10).CellStyle = cs_center;

                sheet.GetRow(5).GetCell(0).CellStyle.SetFont(font0);
                sheet.GetRow(5).GetCell(1).CellStyle.SetFont(font0);
                sheet.GetRow(5).GetCell(2).CellStyle.SetFont(font0);
                sheet.GetRow(5).GetCell(3).CellStyle.SetFont(font0);
                sheet.GetRow(5).GetCell(4).CellStyle.SetFont(font0);
                sheet.GetRow(5).GetCell(5).CellStyle.SetFont(font0);
                sheet.GetRow(5).GetCell(6).CellStyle.SetFont(font0);
                sheet.GetRow(5).GetCell(7).CellStyle.SetFont(font0);
                sheet.GetRow(5).GetCell(8).CellStyle.SetFont(font0);
                sheet.GetRow(5).GetCell(9).CellStyle.SetFont(font0);
                sheet.GetRow(5).GetCell(10).CellStyle.SetFont(font0);

                sheet.SetColumnWidth(0, 8 * 256);
                sheet.SetColumnWidth(1, 16 * 256);
                sheet.SetColumnWidth(2, 9 * 256);
                sheet.SetColumnWidth(3, 19 * 256);
                sheet.SetColumnWidth(4, 39 * 256);
                sheet.SetColumnWidth(5, 45 * 256);
                sheet.SetColumnWidth(6, 11 * 256);
                sheet.SetColumnWidth(7, 10 * 256);
                sheet.SetColumnWidth(8, 20 * 256);
                sheet.SetColumnWidth(9, 14 * 256);
                sheet.SetColumnWidth(10, 31 * 256);

                int zrow = 6;
                foreach (Dictionary<string, string> item in zDetail)
                {
                    sheet.CreateRow(zrow).CreateCell(0).SetCellValue(zrow - 6);
                    sheet.GetRow(zrow).CreateCell(1).SetCellValue(item["VBELN"].ToString());
                    sheet.GetRow(zrow).CreateCell(2).SetCellValue(item["POSNR"].ToString());
                    sheet.GetRow(zrow).CreateCell(3).SetCellValue(item["MATNR"].ToString());
                    sheet.GetRow(zrow).CreateCell(4).SetCellValue(item["ARKTX"].ToString());
                    sheet.GetRow(zrow).CreateCell(5).SetCellValue(item["POSTX"].ToString());
                    if (item["OCDQTY"].ToString() != "")
                    {
                        sheet.GetRow(zrow).CreateCell(6).SetCellValue(double.Parse(item["OCDQTY"].ToString()));
                    } else
                    {
                        sheet.GetRow(zrow).CreateCell(6).SetCellValue(0);
                    }
                    if (item["KBETR"].ToString() != "")
                    {
                        sheet.GetRow(zrow).CreateCell(7).SetCellValue(double.Parse(item["KBETR"].ToString()));
                    } else
                    {
                        sheet.GetRow(zrow).CreateCell(7).SetCellValue(0.00);
                    }
                    if (item["NETWR"].ToString() != "")
                    {
                        sheet.GetRow(zrow).CreateCell(8).SetCellValue(double.Parse(item["NETWR"].ToString()));
                    }
                    else
                    {
                        sheet.GetRow(zrow).CreateCell(8).SetCellValue(0.00);
                    }
                    sheet.GetRow(zrow).CreateCell(9).SetCellValue(item["ZZPDATE"].ToString());
                    sheet.GetRow(zrow).CreateCell(10).SetCellValue("");

                    sheet.GetRow(zrow).GetCell(0).CellStyle = cs_center;
                    sheet.GetRow(zrow).GetCell(1).CellStyle = cs_center;
                    sheet.GetRow(zrow).GetCell(2).CellStyle = cs_left;
                    sheet.GetRow(zrow).GetCell(3).CellStyle = cs_left;
                    sheet.GetRow(zrow).GetCell(4).CellStyle = cs_left;
                    sheet.GetRow(zrow).GetCell(5).CellStyle = cs_left;
                    sheet.GetRow(zrow).GetCell(6).CellStyle = cs_right;
                    sheet.GetRow(zrow).GetCell(7).CellStyle = cs_right2;
                    sheet.GetRow(zrow).GetCell(8).CellStyle = cs_right2;
                    sheet.GetRow(zrow).GetCell(9).CellStyle = cs_left;
                    sheet.GetRow(zrow).GetCell(10).CellStyle = cs_left;

                    sheet.GetRow(zrow).GetCell(0).CellStyle.SetFont(font1);
                    sheet.GetRow(zrow).GetCell(1).CellStyle.SetFont(font1);
                    sheet.GetRow(zrow).GetCell(2).CellStyle.SetFont(font1);
                    sheet.GetRow(zrow).GetCell(3).CellStyle.SetFont(font1);
                    sheet.GetRow(zrow).GetCell(4).CellStyle.SetFont(font1);
                    sheet.GetRow(zrow).GetCell(5).CellStyle.SetFont(font1);
                    sheet.GetRow(zrow).GetCell(6).CellStyle.SetFont(font1);
                    sheet.GetRow(zrow).GetCell(7).CellStyle.SetFont(font1);
                    sheet.GetRow(zrow).GetCell(8).CellStyle.SetFont(font1);
                    sheet.GetRow(zrow).GetCell(9).CellStyle.SetFont(font1);
                    sheet.GetRow(zrow).GetCell(10).CellStyle.SetFont(font1);

                    zrow += 1;
                }
                sheet.CreateRow(zrow);
                sheet.GetRow(zrow).CreateCell(0).SetCellValue("以上預估  箱");
                sheet.GetRow(zrow).CreateCell(4).SetCellValue("總計:");

                sheet.GetRow(zrow).CreateCell(6).CellFormula = "SUM(G7:G" + zrow + ")";
                sheet.GetRow(zrow).CreateCell(8).CellFormula = "SUM(I7:I" + zrow + ")";

                cs_left = null;
                cs_left = hssfworkbook.CreateCellStyle();
                cs_left.Alignment = HorizontalAlignment.Left;

                cs_right = null;
                cs_right = hssfworkbook.CreateCellStyle();
                cs_right.Alignment = HorizontalAlignment.Right;
                cs_right.DataFormat = dataformat.GetFormat("#,##0.00");

                cs_right2 = null;
                cs_right2 = hssfworkbook.CreateCellStyle();
                cs_right2.Alignment = HorizontalAlignment.Right;
                cs_right2.DataFormat = dataformat.GetFormat("#,##0.00");

                font0 = null;
                font0 = (XSSFFont)hssfworkbook.CreateFont();
                font0.FontName = "微軟正黑體";
                font0.FontHeightInPoints = 16;
                font0.Boldweight = (short)FontBoldWeight.Bold;

                font1 = null;
                font1 = (XSSFFont)hssfworkbook.CreateFont();
                font1.FontName = "微軟正黑體";
                font1.FontHeightInPoints = 14;
                font1.Boldweight = (short)FontBoldWeight.Bold;
                font1.Underline = FontUnderlineType.Single;

                sheet.GetRow(zrow).GetCell(0).CellStyle = cs_left;
                sheet.GetRow(zrow).GetCell(4).CellStyle = cs_right;
                sheet.GetRow(zrow).GetCell(0).CellStyle.SetFont(font0);
                sheet.GetRow(zrow).GetCell(4).CellStyle.SetFont(font0);

                sheet.GetRow(zrow).GetCell(6).CellStyle = cs_right;
                sheet.GetRow(zrow).GetCell(8).CellStyle = cs_right2;
                sheet.GetRow(zrow).GetCell(6).CellStyle.SetFont(font1);
                sheet.GetRow(zrow).GetCell(8).CellStyle.SetFont(font1);

                //            var excelDatas = new MemoryStream();
                //            hssfworkbook.Write(excelDatas);
                string tmpfiles = Request.MapPath("~") + @"\tmpfiles\";
                string filesname = DateTime.Now.ToString("yyyyMMddHHmmssfff") + ".xlsx";
                FileStream file = new FileStream(tmpfiles + filesname, FileMode.Create);//產生檔案
                hssfworkbook.Write(file);
                file.Close();
                return filesname;
            }
            catch (Exception ex)
            {
                logger.Debug("ex=" + ex.Message);
                return "";
            }
        }
    }
}