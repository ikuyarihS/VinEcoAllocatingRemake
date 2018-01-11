using System;

namespace VinEcoAllocatingRemake.AllocatingInventory.Models
{
    public class Customer
    {
        public Guid _id { get; set; }

        public Guid CustomerId { get; set; }
        public string CustomerCode { get; set; }
        public string CustomerName { get; set; }
        public string CustomerType { get; set; }
        public string CustomerRegion { get; set; }
        public string CustomerBigRegion { get; set; }
        public string Company { get; set; }
    }
}