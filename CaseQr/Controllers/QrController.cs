using CaseQr.Models;
using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using QRCoder;
using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
namespace CaseQr.Controllers
{
    public class QrController : Controller
    {
        [HttpGet]
        public IActionResult Index()
        {
            List<HolgramQr> obj = new List<HolgramQr>();
            return View(obj);
        }

        [HttpPost]
        public IActionResult Index(IFormFile? OtherDocuments)
        {
            try
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                if (OtherDocuments == null || OtherDocuments.Length == 0)
                {
                    throw new Exception ("No file was uploaded or the file is empty.");
                }

                List<HolgramQr> obj = new List<HolgramQr>();

                using (var stream = OtherDocuments.OpenReadStream())
                using (var package = new ExcelPackage(stream))
                {
                    var workbook = package.Workbook;
                    var worksheet = workbook.Worksheets.FirstOrDefault();

                    if (worksheet == null)
                    {
                        throw new Exception("No worksheet found in the uploaded file.");
                    }

                    int rowCount = worksheet.Dimension.Rows;
                    int colCount = worksheet.Dimension.Columns;

                    for (int row = 1; row <= rowCount; row++)
                    {
                        for (int col = 1; col <= colCount; col++)
                        {
                            var cellValue = worksheet.Cells[row, col].Text;

                            QRCodeGenerator QrGenerator = new QRCodeGenerator();
                            QRCodeData QrCodeInfo = QrGenerator.CreateQrCode(cellValue, QRCodeGenerator.ECCLevel.Q);
                            QRCode QrCode = new QRCode(QrCodeInfo);
                            Bitmap QrBitmap = QrCode.GetGraphic(60);

                            string qrdata;
                            using (MemoryStream ms = new MemoryStream())
                            {
                                QrBitmap.Save(ms, ImageFormat.Png);
                                byte[] BitmapArray = ms.ToArray();
                                qrdata = string.Format("data:image/png;base64,{0}", Convert.ToBase64String(BitmapArray));
                            }

                            obj.Add(new HolgramQr
                            {
                                qrdata = qrdata,
                                CaseNo = cellValue
                            });
                        }
                    }
                }

                return View(obj);
            }
            catch (Exception ex)
            {
                TempData["success"] = ex.Message;
                TempData["Valid"] = "0";
                return RedirectToAction("Index");
            }
        }


    }
}
