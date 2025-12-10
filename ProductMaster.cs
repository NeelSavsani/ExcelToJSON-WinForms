using System;

namespace SampleProject1
{
    public class ProductMaster
    {
        public string Code { get; set; }
        public string Name { get; set; }
        public string ShortName { get; set; }
        public string MfgCompany { get; set; }
        public string MfgCode { get; set; }
        public decimal? GST { get; set; }
        public decimal? MRP { get; set; }
        public decimal? SalesRate { get; set; }
        public decimal? PurchaseRate { get; set; }
        public string Packing { get; set; }
        public bool? IsHide { get; set; }
        public decimal? Stock { get; set; }
        public decimal? Discount { get; set; }
        public string Scheme { get; set; }
        public string ExpDate { get; set; }
        public string Generic { get; set; }
    }
}
