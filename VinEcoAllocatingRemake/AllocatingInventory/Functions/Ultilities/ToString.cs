using System;
using System.Collections.Concurrent;

namespace VinEcoAllocatingRemake.AllocatingInventory
{
    public partial class Utilities
    {
        private readonly ConcurrentDictionary<DateTime, string> _dicDateString =
            new ConcurrentDictionary<DateTime, string>();

        private readonly ConcurrentDictionary<object, string> _dicObjectString =
            new ConcurrentDictionary<object, string>();

        private readonly ConcurrentDictionary<string, string> _dicString = 
            new ConcurrentDictionary<string, string>();

        /// <summary>
        ///     Pretty much a cache for converting Object to String.
        /// </summary>
        /// <param name="obj"></param>
        public string ObjectToString(object obj)
        {
            // Check if exists.
            if (_dicObjectString.TryGetValue(obj, out string value))
                return GetString(value);

            // If not, well, convert.
            if (_dicString.TryGetValue(obj.ToString(), out string valueNew))
            {
                _dicObjectString.TryAdd(obj, valueNew);
                return GetString(valueNew);
            }

            value = obj.ToString();
            _dicString.TryAdd(value, value);
            _dicObjectString.TryAdd(obj, value);
            return GetString(value);

            //value =  ? valueNew : obj.ToString();

            // ... and store the result.
            //_dicObjectString.TryAdd(obj, value);

            // Then return it.
            //return value;
        }

        /// <summary>
        ///     Pretty much a cache for converting DateTime to String.
        /// </summary>
        public string DateToString(DateTime date, string dateFormat = "")
        {
            // Check if exists.
            if (_dicDateString.TryGetValue(date, out string value))
                return GetString(value);

            // If not, well, convert.
            value = date.ToString(dateFormat);

            // ... and store the result.
            _dicDateString.TryAdd(date, value);

            // Then return it.
            return GetString(value);
        }

        public string GetString(string suspect)
        {
            if (_dicString.TryGetValue(suspect, out string result))
                return result;

            _dicString.TryAdd(suspect, suspect);

            return suspect;
        }
    }
}