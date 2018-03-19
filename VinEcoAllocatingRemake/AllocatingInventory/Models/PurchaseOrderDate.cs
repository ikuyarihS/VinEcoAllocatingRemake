using System;
using System.Collections.Generic;

namespace VinEcoAllocatingRemake.AllocatingInventory.Models
{
    #region

    #endregion

    public class PurchaseOrderDate
    {
        public Guid Id { get; set; }

        // [BsonDateTimeOptions(Kind = DateTimeKind.Local)]
        public DateTime DateOrder { get; set; }

        public List<ProductOrder> ListProductOrder { get; set; }

        public Guid PurchaseOrderDateId { get; set; }
    }
}