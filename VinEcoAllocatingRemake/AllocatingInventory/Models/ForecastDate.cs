using System;
using System.Collections.Generic;

namespace VinEcoAllocatingRemake.AllocatingInventory.Models
{
    #region

    #endregion

    public class ForecastDate
    {
        public Guid Id { get; set; }

        ////[BsonDateTimeOptions(Kind = DateTimeKind.Local)]
        public DateTime DateForecast { get; set; }

        public Guid ForecastDateId { get; set; }

        public List<ProductForecast> ListProductForecast { get; set; }
    }
}