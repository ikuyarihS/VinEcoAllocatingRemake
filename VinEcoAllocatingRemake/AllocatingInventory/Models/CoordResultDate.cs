#region

using System;
using System.Collections.Generic;

#endregion

namespace VinEcoAllocatingRemake.AllocatingInventory.Models
{
    public class CoordResultDate
    {
        public Guid _id { get; set; }

        public Guid CoordResultDateId { get; set; }

        public List<CoordinateDate> ListCoordinateDate { get; set; }

        public Guid ProductId { get; set; }
    }
}