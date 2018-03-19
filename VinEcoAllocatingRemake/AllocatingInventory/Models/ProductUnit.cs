using System;
using System.Collections.Generic;

namespace VinEcoAllocatingRemake.AllocatingInventory.Models
{
    #region

    #endregion

    public class ProductUnit
    {
        public Guid Id { get; set; }

        public List<ProductUnitRegion> ListRegion { get; set; }

        public string ProductCode { get; set; }

        public Guid ProductId { get; set; }
    }
}