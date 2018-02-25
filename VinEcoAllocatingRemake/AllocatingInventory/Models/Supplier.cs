namespace VinEcoAllocatingRemake.AllocatingInventory.Models
    {
        /// <summary>
        ///     The supplier.
        /// </summary>
        public struct Supplier
            {
                // public Guid _id { get; set; }

                // public Guid SupplierId { get; set; }

                /// <summary>
                ///     Gets or sets the supplier code.
                /// </summary>
                public string SupplierCode { get; set; }

                /// <summary>
                ///     Gets or sets the supplier name.
                /// </summary>
                public string SupplierName { get; set; }

                /// <summary>
                ///     Gets or sets the supplier type.
                /// </summary>
                public string SupplierType { get; set; }

                /// <summary>
                ///     Gets or sets the supplier region.
                /// </summary>
                public string SupplierRegion { get; set; }
            }
    }