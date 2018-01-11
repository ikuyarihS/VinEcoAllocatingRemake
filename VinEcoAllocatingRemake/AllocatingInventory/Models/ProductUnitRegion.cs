using System;

namespace VinEcoAllocatingRemake.AllocatingInventory.Models
{
    public class ProductUnitRegion
    {
        public Guid _id { get; set; }

        public string Region { get; set; }
        public string OrderUnitType { get; set; }
        public double OrderUnitPer { get; set; }
        public string SaleUnitType { get; set; }
        public double SaleUnitPer { get; set; }
    }
}