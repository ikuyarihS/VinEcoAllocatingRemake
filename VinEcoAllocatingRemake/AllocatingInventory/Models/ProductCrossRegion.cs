﻿using System;

namespace VinEcoAllocatingRemake.AllocatingInventory.Models
{
    #region

    #endregion

    public class ProductCrossRegion
    {
        public ProductCrossRegion()
        {
            ToNorth = true;
            ToSouth = true;
        }

        public Guid Id { get; set; }

        public Guid ProductId { get; set; }

        public bool ToNorth { get; set; }

        public bool ToSouth { get; set; }
    }
}