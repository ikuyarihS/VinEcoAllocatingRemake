using System;
using System.Collections.Concurrent;

namespace VinEcoAllocatingRemake.AllocatingInventory
{
    public partial class Utilities
    {
        private readonly ConcurrentDictionary<DateTime, object> _dicDateObject =
            new ConcurrentDictionary<DateTime, object>();
        private readonly ConcurrentDictionary<double, object> _dicDoubleObject =
            new ConcurrentDictionary<double, object>();

        /// <summary>
        ///     Convert string to DateTime.
        ///     Optimization.
        /// </summary>
        public object DateToObject(DateTime suspect)
        {
            // If string has been converted before.
            if (_dicDateObject.TryGetValue(suspect, out object obj))
                return obj;

            obj = suspect;

            // Welp, it's actually a date.
            // Record the string anyway. Dis many importanto.
            _dicDateObject.TryAdd(suspect, obj);
            return obj;
        }

        /// <summary>
        ///     Convert string to DateTime.
        ///     Optimization.
        /// </summary>
        public object DoubleToObject(double suspect)
        {
            // If string has been converted before.
            if (_dicDoubleObject.TryGetValue(suspect, out object obj))
                return obj;

            obj = suspect;

            // Welp, it's actually a date.
            // Record the string anyway. Dis many importanto.
            _dicDoubleObject.TryAdd(suspect, obj);
            return obj;
        }
    }
}