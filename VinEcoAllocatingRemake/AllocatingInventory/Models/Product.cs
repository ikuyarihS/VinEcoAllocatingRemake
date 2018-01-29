#region

using System.Collections.Generic;

#endregion

namespace VinEcoAllocatingRemake.AllocatingInventory.Models
{
    /// <summary>
    ///     The product.
    /// </summary>
    public struct Product
    {
        // public Product()
        // {
        // Id = Guid.NewGuid();
        // ProductId = Id;

        // //ProductClassification = "???";
        // //ProductOrientation = "???";
        // //ProductClimate = "???";
        // //ProductionGroup = "???";

        // //ProductNote = new List<string>();
        // }

        // public Guid Id { get; set; }
        // public Guid ProductId { get; set; }

        /// <summary>
        ///     Gets or sets the product code.
        /// </summary>
        public string ProductCode { get; set; }

        /// <summary>
        ///     Gets or sets the product name.
        /// </summary>
        public string ProductName { get; set; }

        /// <summary>
        ///     Gets or sets the product ve code.
        /// </summary>
        public string ProductVeCode { get; set; }

        /// <summary>
        ///     Gets or sets the product classification.
        /// </summary>
        public string ProductClassification { get; set; }

        /// <summary>
        ///     Gets or sets the product orientation.
        /// </summary>
        public string ProductOrientation { get; set; }

        /// <summary>
        ///     Gets or sets the product climate.
        /// </summary>
        public string ProductClimate { get; set; }

        /// <summary>
        ///     Gets or sets the production group.
        /// </summary>
        public string ProductionGroup { get; set; }

        // Todo - Figure out what the fuck this is.
        /// <summary>
        ///     Product's Notes, like, idk, I actually forgot what this is for
        /// </summary>
        public List<string> ProductNote { get; set; }
    }
}