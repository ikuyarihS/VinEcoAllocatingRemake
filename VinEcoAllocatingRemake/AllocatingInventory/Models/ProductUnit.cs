#region

using System;
using System.Collections.Generic;

#endregion

namespace VinEcoAllocatingRemake.AllocatingInventory.Models
{
    public class ProductUnit
    {
        public Guid _id { get; set; }

        public List<ProductUnitRegion> ListRegion { get; set; }

        public string ProductCode { get; set; }

        public Guid ProductId { get; set; }
    }
}