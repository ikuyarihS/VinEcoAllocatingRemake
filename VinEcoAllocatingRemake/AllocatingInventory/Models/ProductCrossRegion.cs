#region

using System;

#endregion

namespace VinEcoAllocatingRemake.AllocatingInventory.Models
{
    public class ProductCrossRegion
    {
        public ProductCrossRegion()
        {
            ToNorth = true;
            ToSouth = true;
        }

        public Guid _id { get; set; }

        public Guid ProductId { get; set; }

        public bool ToNorth { get; set; }

        public bool ToSouth { get; set; }
    }
}