using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ExcelDownloadWidget.Models
{
    public class QuotationItemDetailDto
    {
        public string Style { get; set; }
        public string Description { get; set; }
        public int Quantity { get; set; }
        public decimal UnitPrice { get; set; }
        public decimal Total { get; set; }
    }
}