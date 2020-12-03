using ExcelDownloadWidget.Models;
using ExcelDownloadWidget.Services;
using Newtonsoft.Json;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web;
using System.Web.Http;

namespace ExcelDownloadWidget.Controllers
{
    public class QuatationExcelController : ApiController
    {
        //private readonly IMockDataService _mockDataService;

        //public QuatationExcelController(IMockDataService mockDataService)
        //{
        //    _mockDataService = mockDataService;
        //}
        // GET api/<controller>
        public IEnumerable<string> Get()
        {
            return new string[] { "value1", "value2" };
        }

        // GET api/<controller>/5
        public byte[] Get(int id)
        {
            // call the database and get the data set to generate the excel sheet:


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
                IMockDataService mockDataService = new MockDataService();

                var quotationForm = mockDataService.GetQuotationItemDetailList();

                string output = JsonConvert.SerializeObject(quotationForm);

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

				var clientNumber = "FOERRATT C/S 2O";
				var Address1 = "C/- COMPANY MANAGER";
				var Address2 = "16A Lifton Highway";
				var Address3 = "Oceanview";
				var phoneNo = "";
				var faxNumber = "";
				var Attention = "";

				worksheet.Cells["B14"].Value = clientNumber;
				worksheet.Cells["B15"].Value = Address1;
				worksheet.Cells["B16"].Value = Address2;
				worksheet.Cells["B17"].Value = Address3;
				worksheet.Cells["B18"].Value = phoneNo;
				worksheet.Cells["B19"].Value = faxNumber;
				worksheet.Cells["B20"].Value = Attention;

                var clientDetailTextbox02 = worksheet.Drawings.AddShape("Header textbox for client details", eShapeStyle.RoundRect);
                clientDetailTextbox02.SetPosition(12, 14, 0, 2);
                clientDetailTextbox02.SetSize(389, 151);
                clientDetailTextbox02.Fill.Transparancy = 100;
                clientDetailTextbox02.Border.LineStyle = eLineStyle.Solid;
                clientDetailTextbox02.Border.Fill.Color = Color.Black;

                // Quotation detail section
                var quotationDetailtextbox01 = worksheet.Drawings.AddShape("Quotation textbox for quotation number label", eShapeStyle.Rect);
                quotationDetailtextbox01.SetPosition(14, 2, 2, 2);
                quotationDetailtextbox01.SetSize(68, 23);
                quotationDetailtextbox01.Text = "Quotation No:";
                quotationDetailtextbox01.TextAlignment = eTextAlignment.Left;
                quotationDetailtextbox01.Fill.Color = Color.White;
                quotationDetailtextbox01.Font.Color = Color.Black;
                quotationDetailtextbox01.Border.LineStyle = eLineStyle.Solid;
                quotationDetailtextbox01.Border.Fill.Color = Color.White;


                // Quotation detail section
                worksheet.Cells["A11:F11"].Merge = true;


                //Text Values
                worksheet.Cells["A22"].Value = "Dear Customer,";
				worksheet.Cells["B23"].Value = "Our Company is pleased to submit prices on the following lines as per your request.";
				worksheet.Cells["A24"].Value = "All prices are Nett as shown, subject to G.S.T. and to price fluctuations.";

                //Enter quotation details of items and prices section

                var fromRow = 26; 
                var fromCol = 1;
                var toRow = 26; // this is the value we need to calculate using the item list length
                var toCol = 2;

                var defaultRowCount = 26;

                worksheet.Cells[fromRow, fromCol, toRow, toCol].Merge = true;

                worksheet.Cells["A26"].Value = "STYLE";
                worksheet.Cells["C26"].Value = "DESCRIPTION";
                worksheet.Cells["D26"].Value = "QTY";
                worksheet.Cells["E26"].Value = "PRICE";
                worksheet.Cells["F26"].Value = "TOTAL";
                worksheet.Cells["A26:F26"].Style.Font.Color.SetColor(Color.Red);
                worksheet.Cells["A26:F26"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["A26:F26"].Style.VerticalAlignment = ExcelVerticalAlignment.Bottom;

                var currentMaxRowNumber = 26;

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

                var footertextbox01 = worksheet.Drawings.AddShape("Footer textbox for phone label", eShapeStyle.Rect);
                footertextbox01.SetPosition(totalRow + 2, 2, 2, 320);
                footertextbox01.SetSize(68, 23);
                footertextbox01.Text = "Phone: ";
                footertextbox01.TextAlignment = eTextAlignment.Left;
                footertextbox01.Fill.Color = Color.White;
                footertextbox01.Font.Color = Color.Black;
                footertextbox01.Border.LineStyle = eLineStyle.Solid;
                footertextbox01.Border.Fill.Color = Color.White;

                var footertextbox02 = worksheet.Drawings.AddShape("Footer textbox for sign label", eShapeStyle.Rect);
                footertextbox02.SetPosition(totalRow + 3, 12, 2, 320);
                footertextbox02.SetSize(68, 23);
                footertextbox02.Text = "Signed:";
                footertextbox02.TextAlignment = eTextAlignment.Left;
                footertextbox02.Fill.Color = Color.White;
                footertextbox02.Font.Color = Color.Black;
                footertextbox02.Border.LineStyle = eLineStyle.Solid;
                footertextbox02.Border.Fill.Color = Color.White;

                //get the workbook as a bytearray
                var excelBytes = package.GetAsByteArray();

                return excelBytes;
            }
        }

        private int CreateQuotationItemDetailTable(int count, ExcelWorksheet worksheet, List<QuotationItemDetailDto> quotationItemList)
        {
            var RowNumber = 27;
            var isColorRow = false;

            for (int i = 0; i < count; i++)
            {
                worksheet.Cells[RowNumber, 1, RowNumber, 2].Merge = true;

                if (isColorRow)
                {
                    worksheet.Cells[RowNumber, 1, RowNumber, 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    worksheet.Cells[RowNumber, 1, RowNumber, 6].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(1, 255, 255, 153));
                }

                isColorRow = !isColorRow;

                if (i < quotationItemList.Count)
                {
                    worksheet.Cells[RowNumber, 1, RowNumber, 1].Value = quotationItemList[i].Style;
                    worksheet.Cells[RowNumber, 3, RowNumber, 3].Value = quotationItemList[i].Description;
                    worksheet.Cells[RowNumber, 4, RowNumber, 4].Value = quotationItemList[i].Quantity;
                    worksheet.Cells[RowNumber, 5, RowNumber, 5].Value = quotationItemList[i].UnitPrice;
                    worksheet.Cells[RowNumber, 6, RowNumber, 6].Value = quotationItemList[i].Total;
                }

                RowNumber++;
            }

            var fromRow = 26;
            var toRow = fromRow + count;

            //Apply number format
            worksheet.Cells[fromRow + 1, 5, toRow, 6].Style.Numberformat.Format = "$.00";
            worksheet.Cells[fromRow, 1, toRow, 6].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            worksheet.Cells[fromRow, 1, toRow, 6].Style.Border.Right.Style = ExcelBorderStyle.Thin;
            worksheet.Cells[fromRow, 1, toRow, 6].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            worksheet.Cells[fromRow, 1, toRow, 6].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

            worksheet.Cells[fromRow, 1, toRow, 6].Style.Border.Top.Color.SetColor(Color.Black);
            worksheet.Cells[fromRow, 1, toRow, 6].Style.Border.Right.Color.SetColor(Color.Black);
            worksheet.Cells[fromRow, 1, toRow, 6].Style.Border.Left.Color.SetColor(Color.Black);
            worksheet.Cells[fromRow, 1, toRow, 6].Style.Border.Bottom.Color.SetColor(Color.Black);

            return toRow;
        }
    }
}