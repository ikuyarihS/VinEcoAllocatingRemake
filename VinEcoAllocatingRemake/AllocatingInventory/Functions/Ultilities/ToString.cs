using System;
using System.Collections.Concurrent;

namespace VinEcoAllocatingRemake.AllocatingInventory
{
    public partial class Utilities
    {
        private readonly ConcurrentDictionary<object, string> _dicObjectString =
            new ConcurrentDictionary<object, string>();
        private readonly ConcurrentDictionary<DateTime, string> _dicDateString =
            new ConcurrentDictionary<DateTime, string>();

        /// <summary>
        ///     Pretty much a cache for converting Object to String.
        /// </summary>
        /// <param name="obj"></param>
        public string ObjectToString(object obj)
        {
            // Check if exists.
            if (_dicObjectString.TryGetValue(obj, out string value))
                return value;

            // If not, well, convert.
            value = obj.ToString();

            // ... and store the result.
            _dicObjectString.TryAdd(obj, value);

            // Then return it.
            return value;
        }

        /// <summary>
        ///     Pretty much a cache for converting Object to String.
        /// </summary>
        public string DateToString(DateTime date, string dateFormat = "")
        {
            // Check if exists.
            if (_dicDateString.TryGetValue(date, out string value))
                return value;

            // If not, well, convert.
            value = date.ToString(dateFormat);

            // ... and store the result.
            _dicDateString.TryAdd(date, value);

            // Then return it.
            return value;
        }
    }
}