using System;

namespace VinEcoAllocatingRemake.AllocatingInventory
{
    #region

    #endregion

    /// <summary>
    ///     The various Ultilities.
    /// </summary>
    public partial class Utilities
    {
        /// <summary>
        ///     Return 0 instead of a negative for a substraction between two 'convertible double'
        /// </summary>
        /// <param name="value1">Value to be substracted</param>
        /// <param name="value2">Value to substract</param>
        /// <returns>The <see cref="double" />.</returns>
        public double ZeroIfNegative(object value1, object value2)
        {
            double val1 = ObjectToDouble(value1);
            double val2 = ObjectToDouble(value2);
            return Math.Max(val1 - val2, 0);
        }
    }
}