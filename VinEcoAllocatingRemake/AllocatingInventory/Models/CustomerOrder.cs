using System;

namespace VinEcoAllocatingRemake.AllocatingInventory.Models
{
    public class CustomerOrder
    {
        public Guid _id { get; set; }

        public Guid CustomerOrderId { get; set; }
        public Guid CustomerId { get; set; }
        public string Company { get; set; }
        public string Unit { get; set; }
        public double QuantityOrder { get; set; }
        public double QuantityOrderKg { get; set; }
        public string DesiredRegion { get; set; }
        public string DesiredSource { get; set; }
    }
}