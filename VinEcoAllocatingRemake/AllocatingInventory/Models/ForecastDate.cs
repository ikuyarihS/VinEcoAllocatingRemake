using System;
using System.Collections.Generic;
using MongoDB.Bson.Serialization.Attributes;

namespace VinEcoAllocatingRemake.AllocatingInventory.Models
{
    public class ForecastDate
    {
        [BsonId] public Guid _id { get; set; }

        public Guid ForecastDateId { get; set; }

        [BsonDateTimeOptions(Kind = DateTimeKind.Local)]
        public DateTime DateForecast { get; set; }

        public List<ProductForecast> ListProductForecast { get; set; }
    }
}