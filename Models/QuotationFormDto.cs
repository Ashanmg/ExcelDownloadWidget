using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ExcelDownloadWidget.Models
{
    public class QuotationFormDto
    {
        #region Header section
        public string ClientName { get; set; }
        public string Address1 { get; set; }
        public string Address2 { get; set; }
        public string Address3 { get; set; }
        public string PhoneNumber { get; set; }
        public string FaxNumber { get; set; }
        public string Attention { get; set; }
        public string QuotationNumber { get; set; }
        public DateTime Date { get; set; }
        public string RepNumber { get; set; }
        public string RepName { get; set; }
        public string AccountNumber { get; set; }
        public string Position { get; set; }
        public string Reference { get; set; }
        #endregion
        public List<QuotationItemDetailDto> QuotationItemDetails { get; set; }
        public decimal SubTotal { get; set; }
        public decimal PlusGST { get; set; }
        public decimal Total { get; set; }

        public QuotationFormDto()
        {
            QuotationItemDetails = new List<QuotationItemDetailDto>();
        }
    }
}