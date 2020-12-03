using ExcelDownloadWidget.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Web;

namespace ExcelDownloadWidget.Services
{
    public class MockDataService : IMockDataService
    {
        public QuotationFormDto GetQuotationItemDetailList(int quotationId)
        {
            QuotationFormDto quotationForm = new QuotationFormDto();

            if (quotationId == 999999)
            {
                // create mock data here;
                QuotationItemDetailDto detailDto1 = new QuotationItemDetailDto
                {
                    Style = "4194",
                    Description = " WYPALL* X50 (ROAR*) Reinforced Wipers: Roll Wipers (Carton)",
                    Quantity = 2,
                    UnitPrice = 158.98M,
                    Total = 317.96M
                };
                quotationForm.QuotationItemDetails.Add(detailDto1);

                QuotationItemDetailDto detailDto2 = new QuotationItemDetailDto
                {
                    Style = "4198",
                    Description = " WYPALL* X50 (ROAR*) Reinforced Wipers: Roll Wipers (Carton)",
                    Quantity = 1,
                    UnitPrice = 158.98M,
                    Total = 158.98M
                };
                quotationForm.QuotationItemDetails.Add(detailDto2);

                QuotationItemDetailDto detailDto3 = new QuotationItemDetailDto
                {
                    Style = "4917",
                    Description = " Large Roll Dispenser (Each)",
                    Quantity = 1,
                    UnitPrice = 147.22M,
                    Total = 147.22M
                };
                quotationForm.QuotationItemDetails.Add(detailDto3);

                QuotationItemDetailDto detailDto4 = new QuotationItemDetailDto
                {
                    Style = "4926",
                    Description = "Wypall Roll Wipers Dispenser (Each)",
                    Quantity = 1,
                    UnitPrice = 30.07M,
                    Total = 30.07M
                };
                quotationForm.QuotationItemDetails.Add(detailDto4);

                QuotationItemDetailDto detailDto5 = new QuotationItemDetailDto
                {
                    Style = "4940",
                    Description = "Centre Feed Dispenser (Each)",
                    Quantity = 1,
                    UnitPrice = 59.86M,
                    Total = 59.86M
                };
                quotationForm.QuotationItemDetails.Add(detailDto5);

                QuotationItemDetailDto detailDto6 = new QuotationItemDetailDto
                {
                    Style = "H7A",
                    Description = "3M PELTOR Headband Earmuff H7A 290 Green-Each",
                    Quantity = 1,
                    UnitPrice = 39.91M,
                    Total = 39.91M
                };
                quotationForm.QuotationItemDetails.Add(detailDto6);

                QuotationItemDetailDto detailDto7 = new QuotationItemDetailDto
                {
                    Style = "H7HY",
                    Description = "3M PELTOR Earmuff Hygiene Kit HY52 for H7 Series-Pair",
                    Quantity = 1,
                    UnitPrice = 17.3M,
                    Total = 17.3M
                };
                quotationForm.QuotationItemDetails.Add(detailDto7);

                QuotationItemDetailDto detailDto8 = new QuotationItemDetailDto
                {
                    Style = "H7F",
                    Description = "3M PELTOR Folding Headband Earmuff H7F 290 Green-Each",
                    Quantity = 1,
                    UnitPrice = 41.65M,
                    Total = 41.65M
                };
                quotationForm.QuotationItemDetails.Add(detailDto8);

                QuotationItemDetailDto detailDto9 = new QuotationItemDetailDto
                {
                    Style = "H7P3E",
                    Description = "3M PELTOR Helmet Attached Earmuff  H7P3E 290 Green-Each",
                    Quantity = 1,
                    UnitPrice = 42M,
                    Total = 42M
                };
                quotationForm.QuotationItemDetails.Add(detailDto9);

                QuotationItemDetailDto detailDto10 = new QuotationItemDetailDto
                {
                    Style = "2421",
                    Description = "VIKING E/MUFF BIL (Pair)",
                    Quantity = 1,
                    UnitPrice = 29.90M,
                    Total = 29.90M
                };
                quotationForm.QuotationItemDetails.Add(detailDto10);

                // adding duplicates to the validate the some steps
                //quotationForm.Add(detailDto1);
                //quotationForm.Add(detailDto3);
                //quotationForm.Add(detailDto2);
                //quotationForm.Add(detailDto8);
                //quotationForm.Add(detailDto9);
                //quotationForm.Add(detailDto1);
                //quotationForm.Add(detailDto10);
                //quotationForm.Add(detailDto6);
                //quotationForm.Add(detailDto4);
                //quotationForm.Add(detailDto7);
                //quotationForm.Add(detailDto5);
                //quotationForm.Add(detailDto9);
                //quotationForm.Add(detailDto7);
                //quotationForm.Add(detailDto2);
                //quotationForm.Add(detailDto3);
                //quotationForm.Add(detailDto10);
                //quotationForm.Add(detailDto6);
                //quotationForm.Add(detailDto5);
                //quotationForm.Add(detailDto3);
                //quotationForm.Add(detailDto7);

                quotationForm.ClientName = "FOERRATT C/S 2O";
                quotationForm.Address1 = "C/- COMPANY MANAGER";
                quotationForm.Address2 = "16A Lifton Highway";
                quotationForm.Address3 = "Oceanview";

                quotationForm.QuotationNumber = "C656062";
                quotationForm.RepNumber = "2OAN";
                quotationForm.RepName = "Kevin Fernando";
                quotationForm.AccountNumber = "7522WEB";
                quotationForm.Reference = "ANGE RAE";
                quotationForm.Date = DateTime.Today;

                return quotationForm;
            }
            else
            {
                //call the relavant repositories and get the data object list from DB2 side.
                return quotationForm;
            }

        }
    }
}