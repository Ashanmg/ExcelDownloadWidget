using ExcelDownloadWidget.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelDownloadWidget.Services
{
    public interface IMockDataService : IService
    {
        QuotationFormDto GetQuotationItemDetailList(int quotationId = 999999);
    }
}
