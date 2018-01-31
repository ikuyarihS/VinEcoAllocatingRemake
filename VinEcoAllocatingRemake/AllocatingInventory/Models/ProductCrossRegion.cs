namespace VinEcoAllocatingRemake.AllocatingInventory.Models
{
    #region

    using System;

    #endregion

    public class ProductCrossRegion
    {
        public ProductCrossRegion()
        {
            this.ToNorth = true;
            this.ToSouth = true;
        }

        public Guid _id { get; set; }

        public Guid ProductId { get; set; }

        public bool ToNorth { get; set; }

        public bool ToSouth { get; set; }
    }
}