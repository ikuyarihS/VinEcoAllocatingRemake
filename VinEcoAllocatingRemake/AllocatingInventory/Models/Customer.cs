namespace VinEcoAllocatingRemake.AllocatingInventory.Models
    {
        /// <summary>
        ///     The customer.
        /// </summary>
        public struct Customer
            {
                // public Guid _id { get; set; }

                // public Guid CustomerId { get; set; }

                /// <summary>
                ///     Gets or sets the customer key code.
                /// </summary>
                public string CustomerKeyCode { get; set; }

                /// <summary>
                ///     Gets or sets the customer code.
                /// </summary>
                public string CustomerCode { get; set; }

                /// <summary>
                ///     Gets or sets the customer name.
                /// </summary>
                public string CustomerName { get; set; }

                /// <summary>
                ///     Gets or sets the customer type.
                /// </summary>
                public string CustomerType { get; set; }

                /// <summary>
                ///     Gets or sets the customer region.
                /// </summary>
                public string CustomerRegion { get; set; }

                /// <summary>
                ///     Gets or sets the customer big region.
                /// </summary>
                public string CustomerBigRegion { get; set; }

                /// <summary>
                ///     Gets or sets the company.
                /// </summary>
                public string Company { get; set; }
            }
    }