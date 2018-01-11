using System;
using System.Collections.Concurrent;

namespace VinEcoAllocatingRemake.AllocatingInventory
{
    public partial class Utilities
    {
        private readonly ConcurrentDictionary<bool, object> _dicBoolObject =
            new ConcurrentDictionary<bool, object>();

        private readonly ConcurrentDictionary<DateTime, object> _dicDateObject =
            new ConcurrentDictionary<DateTime, object>();

        private readonly ConcurrentDictionary<double, object> _dicDoubleObject =
            new ConcurrentDictionary<double, object>();

        private readonly ConcurrentDictionary<int, object> _dicIntObject =
            new ConcurrentDictionary<int, object>();

        /// <summary>
        ///     Convert DateTime to Object.
        ///     Optimization.
        /// </summary>
        public object DateToObject(DateTime suspect, string dateFormat = "")
        {
            // If string has been converted before.
            if (_dicDateObject.TryGetValue(suspect, out object obj))
                return obj;

            obj = suspect.ToString(dateFormat);

            // Welp, it's actually a date.
            // Record the string anyway. Dis many importanto.
            _dicDateObject.TryAdd(suspect, obj);
            return obj;
        }

        /// <summary>
        ///     Convert double to Object.
        ///     Optimization.
        /// </summary>
        public object DoubleToObject(double suspect)
        {
            // If string has been converted before.
            if (_dicDoubleObject.TryGetValue(suspect, out object obj))
                return obj;

            obj = suspect.ToString(string.Empty);

            // Welp, it's actually a date.
            // Record the string anyway. Dis many importanto.
            _dicDoubleObject.TryAdd(suspect, obj);
            return obj;
        }

        /// <summary>
        ///     Convert Boolean to Object.
        ///     Optimization.
        /// </summary>
        public object BoolToObject(bool suspect)
        {
            // If string has been converted before.
            if (_dicBoolObject.TryGetValue(suspect, out object obj))
                return obj;

            obj = suspect.ToString();

            // Welp, it's actually a date.
            // Record the string anyway. Dis many importanto.
            _dicBoolObject.TryAdd(suspect, obj);
            return obj;
        }

        /// <summary>
        ///     Convert Int to Object.
        ///     Optimization.
        /// </summary>
        public object IntToObject(int suspect)
        {
            // If string has been converted before.
            if (_dicIntObject.TryGetValue(suspect, out object obj))
                return obj;

            obj = suspect.ToString();

            // Welp, it's actually a date.
            // Record the string anyway. Dis many importanto.
            _dicIntObject.TryAdd(suspect, obj);
            return obj;
        }
    }
}