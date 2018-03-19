using System;

namespace VinEcoAllocatingRemake.AllocatingInventory.Models
{
    #region

    #endregion

    public class ProductUnitRegion
    {
        public Guid Id { get; set; }

        public double OrderUnitPer { get; set; }

        public string OrderUnitType { get; set; }

        public string Region { get; set; }

        public double SaleUnitPer { get; set; }

        public string SaleUnitType { get; set; }
    }
}