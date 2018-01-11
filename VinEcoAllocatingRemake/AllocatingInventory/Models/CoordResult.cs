using System;
using System.Collections.Generic;

namespace VinEcoAllocatingRemake.AllocatingInventory.Models
{
    public class CoordResult
    {
        public Guid _id { get; set; }

        public Guid CoordResultId { get; set; }

        ////[BsonDateTimeOptions(Kind = DateTimeKind.Local)]
        public DateTime DateOrder { get; set; }

        public List<CoordResultDate> ListCoordResultDate { get; set; }
    }
}