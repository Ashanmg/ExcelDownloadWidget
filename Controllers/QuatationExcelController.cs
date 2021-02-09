using ExcelDownloadWidget.Models;
using ExcelDownloadWidget.Services;
using GemBox.Spreadsheet;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Style;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using System.Collections.Generic;
using System.Configuration;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Http;
using System.Web.Routing;

namespace ExcelDownloadWidget.Controllers
{
    [RoutePrefix("api/excel")]
    public class QuatationExcelController : ApiController
    {
        private readonly string excelFolderPath = ConfigurationManager.AppSettings["ExcelTemplateFolderPath"].ToString();
        private readonly IMockDataService _mockDataService;

        public QuatationExcelController(IMockDataService mockDataService)
        {
            _mockDataService = mockDataService;
        }
        // GET api/<controller>
        public IEnumerable<string> Get()
        {
            return new string[] { "value1", "value2" };
        }

        // GET api/<controller>/5
        public byte[] Get(int id)
        {
            // Create the excel sheet using the retireived data
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("My Sheet");

				//Add image
				System.Drawing.Image image = System.Drawing.Image.FromFile(HttpContext.Current.Server.MapPath("/Content/Images/logoheader.png"));
				var excelImage = worksheet.Drawings.AddPicture("NZSafety Logo", image);

				//Add the image to row 4, column A	
				excelImage.SetPosition(3, 3, 0, 25);

				//Common Attributes for all cells
				worksheet.Cells.Style.Border.Left.Style = ExcelBorderStyle.Thin;
				worksheet.Cells.Style.Border.Right.Style = ExcelBorderStyle.Thin;
				worksheet.Cells.Style.Border.Top.Style = ExcelBorderStyle.Thin;
				worksheet.Cells.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                worksheet.Cells.Style.Border.Left.Color.SetColor(Color.White);
                worksheet.Cells.Style.Border.Right.Color.SetColor(Color.White);
                worksheet.Cells.Style.Border.Top.Color.SetColor(Color.White);
                worksheet.Cells.Style.Border.Bottom.Color.SetColor(Color.White);

                worksheet.Cells.Style.Font.Name = "Arial";
                worksheet.Cells.Style.Font.Size = 10;

                // Get the mock data from the mock data services
                // IMockDataService mockDataService = new MockDataService();

                var quotationForm = _mockDataService.GetQuotationItemDetailList();

                //Set column with from column A to F
                worksheet.Column(1).Width = 15;
				worksheet.Column(2).Width = 1.8;
				worksheet.Column(3).Width = 56.5;
				worksheet.Column(4).Width = 6.3;
				worksheet.Column(5).Width = 10.5;
				worksheet.Column(6).Width = 11;

				//Add the Title of the Form
				worksheet.Cells["A11:F11"].Merge = true;
				worksheet.Cells["A11"].Style.Font.Bold = true;
				worksheet.Cells["A11"].Style.Font.Size = 14;
				worksheet.Cells["A11"].Style.Font.Name = "Arial";
				worksheet.Cells["A11"].Style.Font.UnderLine = true;
				worksheet.Cells["A11"].Value = "Quotation Form";
				worksheet.Cells["A11"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

				// Client detail section
				worksheet.Cells["A14"].Value = "   Client:";
				worksheet.Cells["A15"].Value = "   Address:";
				worksheet.Cells["A18"].Value = "   Phone No:";
				worksheet.Cells["A19"].Value = "   Fax No:";
				worksheet.Cells["A20"].Value = "   Attention:";

				worksheet.Cells["B14"].Value = quotationForm.ClientName;
				worksheet.Cells["B15"].Value = quotationForm.Address1;
				worksheet.Cells["B16"].Value = quotationForm.Address2;
				worksheet.Cells["B17"].Value = quotationForm.Address3;
				worksheet.Cells["B18"].Value = quotationForm.PhoneNumber;
				worksheet.Cells["B19"].Value = quotationForm.FaxNumber;
				worksheet.Cells["B20"].Value = quotationForm.Attention;

                var clientDetailTextbox01 = worksheet.Drawings.AddShape("Header textbox for client details", eShapeStyle.RoundRect);
                clientDetailTextbox01.SetPosition(12, 14, 0, 2);
                clientDetailTextbox01.SetSize(389, 151);
                clientDetailTextbox01.Fill.Transparancy = 100;
                clientDetailTextbox01.Border.LineStyle = eLineStyle.Solid;
                clientDetailTextbox01.Border.Fill.Color = Color.Black;

                // Quotation detail section
                CreateReactangleTextBox(worksheet, 13, 1, 2, 296, 99, 17, "Quotation No:", "Quotation textbox for quotation number label");

                CreateReactangleTextBox(worksheet, 14, 1, 2, 296, 99, 17, "Date:", "Quotation textbox for Date label");

                CreateReactangleTextBox(worksheet, 15, 1, 2, 296, 99, 17, "Rep No:", "Quotation textbox for Rep No label");

                CreateReactangleTextBox(worksheet, 16, 1, 2, 296, 99, 17, "Rep Name:", "Quotation textbox for Rep Name label");

                CreateReactangleTextBox(worksheet, 17, 1, 2, 296, 99, 17, "Account No:", "Quotation textbox for Account No label");

                CreateReactangleTextBox(worksheet, 18, 1, 2, 296, 99, 17, "Position:", "Quotation textbox for Position label");

                CreateReactangleTextBox(worksheet, 19, 1, 2, 296, 99, 17, "Reference:", "Quotation textbox for Reference label");

                worksheet.Cells["D14"].Value = quotationForm.QuotationNumber;
                worksheet.Cells["D15"].Formula = "=TODAY()";
                worksheet.Cells["D16"].Value = quotationForm.RepNumber;
                worksheet.Cells["D17"].Value = quotationForm.RepName;
                worksheet.Cells["D18"].Value = quotationForm.AccountNumber;
                worksheet.Cells["D19"].Value = quotationForm.Position;
                worksheet.Cells["D20"].Value = quotationForm.Reference;

                var clientDetailTextbox02 = worksheet.Drawings.AddShape("Header textbox for quotation details", eShapeStyle.RoundRect);
                clientDetailTextbox02.SetPosition(12, 14, 2, 290);
                clientDetailTextbox02.SetSize(299, 151);
                clientDetailTextbox02.Fill.Transparancy = 100;
                clientDetailTextbox02.Border.LineStyle = eLineStyle.Solid;
                clientDetailTextbox02.Border.Fill.Color = Color.Black;
                clientDetailTextbox02.Border.Width = 1;

                // Quotation detail section
                worksheet.Cells["A11:F11"].Merge = true;

                //Text Values
                worksheet.Cells["A22"].Value = "Dear Customer,";
				worksheet.Cells["B23"].Value = "Our Company is pleased to submit prices on the following lines as per your request.";
				worksheet.Cells["A24"].Value = "All prices are Nett as shown, subject to G.S.T. and to price fluctuations.";

                //Enter quotation details of items and prices section

                worksheet.Cells[26, 1, 26, 2].Merge = true;

                worksheet.Cells["A26"].Value = "STYLE";
                worksheet.Cells["C26"].Value = "DESCRIPTION";
                worksheet.Cells["D26"].Value = "QTY";
                worksheet.Cells["E26"].Value = "PRICE";
                worksheet.Cells["F26"].Value = "TOTAL";
                worksheet.Cells["A26:F26"].Style.Font.Color.SetColor(Color.Red);
                worksheet.Cells["A26:F26"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["A26:F26"].Style.VerticalAlignment = ExcelVerticalAlignment.Bottom;

                var currentMaxRowNumber = 26;
                var defaultRowCount = 26;

                if (defaultRowCount > quotationForm.QuotationItemDetails.Count)
                {
                    currentMaxRowNumber = CreateQuotationItemDetailTable(defaultRowCount, worksheet, quotationForm.QuotationItemDetails);
                }
                else
                {
                    currentMaxRowNumber = CreateQuotationItemDetailTable(quotationForm.QuotationItemDetails.Count, worksheet, quotationForm.QuotationItemDetails);
                }   

                //sub total row
                var subTotalRow = currentMaxRowNumber + 1;
                worksheet.Cells[subTotalRow, 4, subTotalRow, 5].Merge = true;
                worksheet.Cells[subTotalRow, 4].Value = "Sub Total";
                worksheet.Cells[subTotalRow, 6].Formula = "=SUM(" + worksheet.Cells[26, 6].Address + ":" + worksheet.Cells[currentMaxRowNumber, 6].Address + ")";

                //add Gst to sub total
                var gstRow = currentMaxRowNumber + 2;
                worksheet.Cells[gstRow, 4, gstRow, 5].Merge = true;
                worksheet.Cells[gstRow, 4].Value = "Plus GST";
                worksheet.Cells[gstRow, 6].Formula = "= IF(SUM(" + worksheet.Cells[subTotalRow, 6].Address + ") = 0," + "" + ", SUM(" + worksheet.Cells[subTotalRow, 6].Address + "* 0.15))";

                //total row
                var totalRow = currentMaxRowNumber + 3;
                worksheet.Cells[totalRow, 4, totalRow, 5].Merge = true;
                worksheet.Cells[totalRow, 4].Value = "TOTAL";
                worksheet.Cells[totalRow, 4].Style.Font.UnderLine = true;
                worksheet.Cells[totalRow, 4].Style.Font.Bold = true;
                worksheet.Cells[totalRow, 4].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                worksheet.Cells[totalRow, 4].Style.VerticalAlignment = ExcelVerticalAlignment.Bottom;
                worksheet.Cells[totalRow, 6].Formula = "= IF(SUM("+ worksheet.Cells[subTotalRow, 6].Address + ":" + worksheet.Cells[gstRow, 6].Address + ") = 0," + "" + ", SUM(" + worksheet.Cells[subTotalRow, 6].Address + ":" + worksheet.Cells[gstRow, 6].Address + "))";
                
                
                worksheet.Cells[subTotalRow, 6, totalRow, 6].Style.Numberformat.Format = "$.00";

                // set border of these sub total, gst and total value section
                worksheet.Cells[subTotalRow, 4, gstRow, 6].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                worksheet.Cells[subTotalRow, 4, gstRow, 6].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                worksheet.Cells[subTotalRow, 4, gstRow, 6].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                worksheet.Cells[subTotalRow, 4, gstRow, 6].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                worksheet.Cells[subTotalRow, 4, gstRow, 6].Style.Border.Top.Color.SetColor(Color.Black);
                worksheet.Cells[subTotalRow, 4, gstRow, 6].Style.Border.Right.Color.SetColor(Color.Black);
                worksheet.Cells[subTotalRow, 4, gstRow, 6].Style.Border.Left.Color.SetColor(Color.Black);
                worksheet.Cells[subTotalRow, 4, gstRow, 6].Style.Border.Bottom.Color.SetColor(Color.Black);

                worksheet.Cells[totalRow, 6].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                worksheet.Cells[totalRow, 6].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                worksheet.Cells[totalRow, 6].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                worksheet.Cells[totalRow, 6].Style.Border.Bottom.Style = ExcelBorderStyle.Double;

                worksheet.Cells[totalRow, 6].Style.Border.Top.Color.SetColor(Color.Black);
                worksheet.Cells[totalRow, 6].Style.Border.Right.Color.SetColor(Color.Black);
                worksheet.Cells[totalRow, 6].Style.Border.Left.Color.SetColor(Color.Black);
                worksheet.Cells[totalRow, 6].Style.Border.Bottom.Color.SetColor(Color.Black);

                // footter section
                worksheet.Cells[totalRow + 2, 1].Value = "IMPORTANT: To ensure that the prices quoted herewith are effective against purchases made, please confirm your acceptance of this";
                worksheet.Cells[totalRow + 3, 1].Value = "offer by contacting:";
                worksheet.Cells[totalRow + 2, 1, totalRow + 3, 1].Style.Font.Size = 8;
                worksheet.Cells[totalRow + 4, 1].Value = "This offer is effective until:";

                worksheet.Cells[totalRow + 3, 3, totalRow + 4, 3].Style.Font.Size = 12;
                worksheet.Row(totalRow + 3).Height = 15;
                worksheet.Row(totalRow + 4).Height = 25;
                var blankline01 = worksheet.Drawings.AddShape("Blankline for name", eShapeStyle.Line);
                blankline01.SetPosition(totalRow + 3, 0, 2, 3);
                blankline01.SetSize(286, 0);

                var blankline02 = worksheet.Drawings.AddShape("Blankline for phone", eShapeStyle.Line);
                blankline02.SetPosition(totalRow + 3, 0, 3, 3);
                blankline02.SetSize(143, 0);

                var blankline03 = worksheet.Drawings.AddShape("Blankline for date", eShapeStyle.Line);
                blankline03.SetPosition(totalRow + 4, 0, 2, 56);
                blankline03.SetSize(231, 0);

                var blankline04 = worksheet.Drawings.AddShape("Blankline for sign", eShapeStyle.Line);
                blankline04.SetPosition(totalRow + 4, 0, 3, 3);
                blankline04.SetSize(143, 0);

                CreateReactangleTextBox(worksheet, totalRow + 2, 2, 2, 320, 68, 23, "Phone: ", "Footer textbox for phone label");

                CreateReactangleTextBox(worksheet, totalRow + 3, 12, 2, 320, 68, 23, "Signed:", "Footer textbox for sign label");

                worksheet.Cells[totalRow + 3, 3].Formula = "=$D$17";
                worksheet.Cells[totalRow + 3, 4].Formula = "=$B$18";
                worksheet.Cells[totalRow + 4, 3].Formula = "=SUM(D15+30)";


                //get the workbook as a bytearray
                var excelBytes = package.GetAsByteArray();

                return excelBytes;
            }
        }

        [HttpGet]
        [Route("excel-template")]
        public byte[] GetFromTemplate()
        {
            DeleteFilesFromServer(HttpContext.Current.Server.MapPath("/Content/TempFiles"));

            FileInfo existingFile = new FileInfo(HttpContext.Current.Server.MapPath("/Content/ExcelTemplates/quotation_excel-template.xlsx"));

            using (ExcelPackage package = new ExcelPackage(existingFile))
            {
                var quotationForm = _mockDataService.GetQuotationItemDetailList();

                var worksheet = package.Workbook.Worksheets.First();

                worksheet.Cells["B14"].Value = quotationForm.ClientName;
                worksheet.Cells["B15"].Value = quotationForm.Address1;
                worksheet.Cells["B16"].Value = quotationForm.Address2;
                worksheet.Cells["B17"].Value = quotationForm.Address3;
                worksheet.Cells["B18"].Value = quotationForm.PhoneNumber;
                worksheet.Cells["B19"].Value = quotationForm.FaxNumber;
                worksheet.Cells["B20"].Value = quotationForm.Attention;

                worksheet.Cells["D14"].Value = quotationForm.QuotationNumber;
                worksheet.Cells["D16"].Value = quotationForm.RepNumber;
                worksheet.Cells["D17"].Value = quotationForm.RepName;
                worksheet.Cells["D18"].Value = quotationForm.AccountNumber;
                worksheet.Cells["D19"].Value = quotationForm.Position;
                worksheet.Cells["D20"].Value = quotationForm.Reference;

                var defaultRowCount = 26;

                if (defaultRowCount > quotationForm.QuotationItemDetails.Count)
                {
                    CreateQuotationItemDetailTable(defaultRowCount, worksheet, quotationForm.QuotationItemDetails);
                }
                else
                {
                    CreateQuotationItemDetailTable(quotationForm.QuotationItemDetails.Count, worksheet, quotationForm.QuotationItemDetails);
                }

                AddFileFolderToServer(HttpContext.Current.Server.MapPath("/Content/TempFiles"));

                var destinationFile = new FileInfo(HttpContext.Current.Server.MapPath("/Content/TempFiles/tempquotationexcel.xlsx"));

                package.SaveAs(destinationFile);
            }

            // If using Professional version, put your serial key below.
            SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

            // In order to convert Excel to PDF, we just need to:
            //   1. Load XLS or XLSX file into ExcelFile object.
            //   2. Save ExcelFile object to PDF file.
            ExcelFile workbook = ExcelFile.Load(HttpContext.Current.Server.MapPath("/Content/TempFiles/tempquotationexcel.xlsx"));

            var wwww = workbook.Worksheets.FirstOrDefault();
            workbook.Worksheets.ActiveWorksheet = wwww;

            CellRange range = wwww.Cells.GetSubrange("A1:F62");

            wwww.NamedRanges.SetPrintArea(range);

            workbook.Save(HttpContext.Current.Server.MapPath("/Content/TempFiles/converted.pdf"));

            PdfDocument outputDocument = new PdfDocument();
            PdfDocument inputDocument = PdfReader.Open(HttpContext.Current.Server.MapPath("/Content/TempFiles/converted.pdf"), PdfDocumentOpenMode.Import);

            int count = inputDocument.PageCount;
            for (int idx = 0; idx < count; idx++)
            {
                PdfPage page = inputDocument.Pages[idx];
                outputDocument.AddPage(page);
            }

            byte[] combinedBytes;

            using (MemoryStream rms = new MemoryStream())
            {
                outputDocument.Save(rms, false);
                combinedBytes = rms.ToArray();
            }

            //get the workbook as a bytearray
            return combinedBytes;
        }

        [Route("download-quotation")]
        [HttpGet]
        public byte[] DownloadQuotation()
        {
            bool withDiscount = false;

            FileInfo existingFile;
            if (withDiscount)
            {
                existingFile = new FileInfo(HttpContext.Current.Server.MapPath(string.Format("{0}{1}", excelFolderPath, "quotation_excel-template1.xlsx")));
            }
            else
            {
                existingFile = new FileInfo(HttpContext.Current.Server.MapPath(string.Format("{0}{1}", excelFolderPath, "quotation_excel-template2.xlsx")));
            }

            using (ExcelPackage package = new ExcelPackage(existingFile))
            {
                var quotationForm = _mockDataService.GetQuotationItemDetailList();

                var worksheet = package.Workbook.Worksheets.First();

                worksheet.Cells["B14"].Value = quotationForm.ClientName;
                worksheet.Cells["B15"].Value = quotationForm.Address1;
                worksheet.Cells["B16"].Value = quotationForm.Address2;
                worksheet.Cells["B17"].Value = quotationForm.Address3;
                worksheet.Cells["B18"].Value = quotationForm.PhoneNumber;
                worksheet.Cells["B19"].Value = quotationForm.FaxNumber;
                worksheet.Cells["B20"].Value = quotationForm.Attention;

                worksheet.Cells["D14"].Value = quotationForm.QuotationNumber;
                worksheet.Cells["D16"].Value = quotationForm.RepNumber;
                worksheet.Cells["D17"].Value = quotationForm.RepName;
                worksheet.Cells["D18"].Value = quotationForm.AccountNumber;
                worksheet.Cells["D19"].Value = quotationForm.Position;
                worksheet.Cells["D20"].Value = quotationForm.Reference;

                var defaultRowCount = 26;

                if (defaultRowCount > quotationForm.QuotationItemDetails.Count)
                {
                    InsertQuotationItemDetails(defaultRowCount, worksheet, quotationForm.QuotationItemDetails, withDiscount);
                }
                else
                {
                    InsertQuotationItemDetails(quotationForm.QuotationItemDetails.Count, worksheet, quotationForm.QuotationItemDetails, withDiscount);
                }

                //get the workbook as a bytearray
                return package.GetAsByteArray();
            }
        }

        private int CreateQuotationItemDetailTable(int count, OfficeOpenXml.ExcelWorksheet worksheet, List<QuotationItemDetailDto> quotationItemList)
        {
            var RowNumber = 29;

            for (int i = 0; i < count; i++)
            {
                if (i < quotationItemList.Count)
                {
                    worksheet.Cells[RowNumber, 1, RowNumber, 1].Value = quotationItemList[i].Style;
                    worksheet.Cells[RowNumber, 3, RowNumber, 3].Value = quotationItemList[i].Description;
                    worksheet.Cells[RowNumber, 4, RowNumber, 4].Value = quotationItemList[i].Quantity;
                    worksheet.Cells[RowNumber, 7, RowNumber, 7].Value = quotationItemList[i].UnitPrice;
                }
                RowNumber++;
            }

            var defaultItemrows = 26;
            if (defaultItemrows > quotationItemList.Count)
            {
                worksheet.DeleteRow(56, 973, true);
            }
            else
            {
                var numberOfRowsToDelete = 1029 - (RowNumber + quotationItemList.Count);
                worksheet.DeleteRow(RowNumber + quotationItemList.Count, numberOfRowsToDelete, true);
            }

            return 1;
        }

        private void InsertQuotationItemDetails(int count, OfficeOpenXml.ExcelWorksheet worksheet, List<QuotationItemDetailDto> quotationItemList, bool withDiscount)
        {
            var RowNumber = 29;

            for (int i = 0; i < count; i++)
            {
                if (i < quotationItemList.Count)
                {
                    worksheet.Cells[RowNumber, 1, RowNumber, 1].Value = quotationItemList[i].Style;
                    worksheet.Cells[RowNumber, 3, RowNumber, 3].Value = quotationItemList[i].Description;
                    worksheet.Cells[RowNumber, 4, RowNumber, 4].Value = quotationItemList[i].Quantity;
                    if (withDiscount)
                    {
                        worksheet.Cells[RowNumber, 7, RowNumber, 7].Value = quotationItemList[i].UnitPrice;
                    }
                    else
                    {
                        worksheet.Cells[RowNumber, 5, RowNumber, 5].Value = quotationItemList[i].UnitPrice;
                    }
                }
                RowNumber++;
            }

            var defaultItemrows = 26;
            if (defaultItemrows > quotationItemList.Count)
            {
                var deleteFrom = 56;
                worksheet.DeleteRow(deleteFrom, 973, true);
            }
            else
            {
                var numberOfRowsToDelete = 974 - (quotationItemList.Count - defaultItemrows);
                worksheet.DeleteRow(RowNumber, numberOfRowsToDelete, true);
            }
        }

        private void CreateReactangleTextBox(OfficeOpenXml.ExcelWorksheet worksheet,
        int Row, int RowOffsetPixels, int Column, int ColumnOffsetPixels,
        int PixelWidth, int PixelHeight, string text, string shapeName)
        {
            var textbox = worksheet.Drawings.AddShape(shapeName, eShapeStyle.Rect);
            textbox.SetPosition(Row, RowOffsetPixels, Column, ColumnOffsetPixels);
            textbox.SetSize(PixelWidth, PixelHeight);
            textbox.Text = text;
            textbox.TextAlignment = eTextAlignment.Left;
            textbox.Fill.Color = Color.White;
            textbox.Font.Color = Color.Black;
            textbox.Border.LineStyle = eLineStyle.Solid;
            textbox.Border.Fill.Color = Color.White;
        }

        public void DeleteFilesFromServer(string tempfolderpath)
        {
            if (Directory.Exists(tempfolderpath))
            {
                Directory.Delete(tempfolderpath, true);
            }
        }

        public void AddFileFolderToServer(string tempfolderpath)
        {
            if (!Directory.Exists(tempfolderpath))
            {
                Directory.CreateDirectory(tempfolderpath);
            }
        }
    }
} 