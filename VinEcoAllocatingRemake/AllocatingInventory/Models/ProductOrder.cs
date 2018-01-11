using System;
using System.Collections.Generic;

namespace VinEcoAllocatingRemake.AllocatingInventory.Models
{
    public class ProductOrder
    {
        public Guid _id { get; set; }

        public Guid ProductOrderId { get; set; }
        public Guid ProductId { get; set; }
        public List<CustomerOrder> ListCustomerOrder { get; set; }
    }
}