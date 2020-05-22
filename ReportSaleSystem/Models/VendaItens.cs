using System;

namespace ReportSaleSystem.Models
{
    public class VendaItens
    {
        public string SaleID { get; set; }
        public string SalesmanName { get; set; }
        public double ItemID { get; set; }
        public double ItemQuantity { get; set; }
        public double ItemPrice { get; set; }
        public double Total => GetTotal();

        protected virtual double GetTotal()
        {
            return ItemQuantity > 0 ? Math.Round((ItemQuantity * ItemPrice), 2, MidpointRounding.AwayFromZero) : 0;
        }
    }
}
