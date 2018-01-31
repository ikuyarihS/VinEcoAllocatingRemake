namespace VinEcoAllocatingRemake.AllocatingInventory.Models
{
    #region

    using System;

    #endregion

    public class ProductRate
    {
        public Guid _id { get; set; }

        public string ProductCode { get; set; }

        public Guid ProductId { get; set; }

        public double ToNorth { get; set; }

        public double ToSouth { get; set; }
    }
}