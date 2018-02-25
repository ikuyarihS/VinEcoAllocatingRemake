namespace VinEcoAllocatingRemake.AllocatingInventory.Models
    {
        public struct CustomerOrder
            {
                // public Guid _id { get; set; }
                // public Guid CustomerOrderId { get; set; }
                // public Guid CustomerId { get; set; }
                public string CustomerKeyCode { get; set; }

                public string CustomerCode { get; set; }

                // public string Company { get; set; }
                public double QuantityOrder { get; set; }

                // public double QuantityOrderKg { get; set; }
                public string DesiredRegion { get; set; }

                public string DesiredSource { get; set; }
            }
    }