using System;

namespace VinEcoAllocatingRemake.AllocatingInventory.Models
{
    public struct Product
    {
        //public Product()
        //{
        //    Id = Guid.NewGuid();
        //    ProductId = Id;

        //    //ProductClassification = "???";
        //    //ProductOrientation = "???";
        //    //ProductClimate = "???";
        //    //ProductionGroup = "???";

        //    //ProductNote = new List<string>();
        //}

        //public Guid Id { get; set; }
        //public Guid ProductId { get; set; }
        public string ProductCode { get; set; }
        public string ProductName { get; set; }
        public string ProductVECode { get; set; }
        public string ProductClassification { get; set; }
        public string ProductOrientation { get; set; }
        public string ProductClimate { get; set; }

        public string ProductionGroup { get; set; }
        //public List<string> ProductNote { get; set; }
    }
}