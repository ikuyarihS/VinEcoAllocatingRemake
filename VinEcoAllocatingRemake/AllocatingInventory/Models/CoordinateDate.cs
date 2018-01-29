#region

using System;

#endregion

namespace VinEcoAllocatingRemake.AllocatingInventory.Models
{
    #region

    #endregion

    /// <summary>
    ///     The coordinate date.
    /// </summary>
    public class CoordinateDate
    {
        public Guid _id { get; set; }

        public Guid CoordinateDateId { get; set; }

        public Guid CustomerOrderId { get; set; }

        public DateTime? DateDelier { get; set; }

        public Guid? SupplierOrderId { get; set; }
    }
}