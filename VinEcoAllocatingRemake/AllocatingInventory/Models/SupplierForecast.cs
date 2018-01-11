using System;

namespace VinEcoAllocatingRemake.AllocatingInventory.Models
{
    public class SupplierForecast
    {
        public SupplierForecast()
        {
            Id = Guid.NewGuid();
            SupplierForecastId = Id;

            Level = 1;
            Availability = "1234567";
            Target = "All";
        }

        public Guid Id { get; set; }

        public Guid SupplierForecastId { get; set; }
        public Guid SupplierId { get; set; }
        public string SupplierCode { get; set; }
        public bool LabelVinEco { get; set; }
        public bool FullOrder { get; set; }
        public bool QualityControlPass { get; set; }
        public bool CrossRegion { get; set; }
        public byte Level { get; set; }
        public string Availability { get; set; }
        public string Target { get; set; }
        public double QuantityForecast { get; set; }
        public bool HasKpi { get; set; }
        public double QuantityForecastPlanned { get; set; }
        public double QuantityForecastOriginal { get; set; }
    }
}