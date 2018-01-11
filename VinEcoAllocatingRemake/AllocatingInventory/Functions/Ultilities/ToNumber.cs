using System;
using System.Collections.Concurrent;
using System.Diagnostics;

namespace VinEcoAllocatingRemake.AllocatingInventory
{
    public partial class Utilities
    {
        private readonly ConcurrentDictionary<object, double> _dicObjectDouble =
            new ConcurrentDictionary<object, double>();

        private readonly ConcurrentDictionary<object, int> _dicObjectInt =
            new ConcurrentDictionary<object, int>();

        private readonly ConcurrentDictionary<object, double> _dicStringDouble =
            new ConcurrentDictionary<object, double>();

        /// <summary>
        ///     Pretty much a cache for converting double.
        /// </summary>
        /// <param name="obj"></param>
        public double ObjectToDouble(object obj)
        {
            
            // Check if exists.
            if (_dicObjectDouble.TryGetValue(obj, out double value))
                return value;

            // Goddamn it.
            if (obj == DBNull.Value)
            {
                _dicObjectDouble.TryAdd(obj, 0);
                return 0;
            }

            try
            {
                // If not, well, convert.
                value = Convert.ToDouble(obj);
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
                value = 0;
            }

            // ... and store the result.
            _dicObjectDouble.TryAdd(obj, value);

            // Then return it.
            return value;
        }

        /// <summary>
        ///     Pretty much a cache for converting double.
        /// </summary>
        /// <param name="key"></param>
        public double StringToDouble(string key)
        {
            // Check if exists.
            if (_dicStringDouble.TryGetValue(key, out double value))
                return value;

            // If not, well, convert.
            value = Convert.ToDouble(key);

            // ... and store the result.
            _dicStringDouble.TryAdd(key, value);

            // Then return it.
            return value;
        }

        /// <summary>
        ///     Pretty much a cache for converting Int.
        /// </summary>
        /// <param name="suspect"></param>
        public int ObjectToInt(object suspect)
        {
            // Check if exists.
            if (_dicObjectInt.TryGetValue(suspect, out int value))
                return value;

            // Goddamn it.
            if (suspect == DBNull.Value)
            {
                _dicObjectDouble.TryAdd(suspect, 0);
                return 0;
            }

            // If not, well, convert.
            value = Convert.ToInt32(suspect);

            // ... and store the result.
            _dicObjectInt.TryAdd(suspect, value);

            // Then return it.
            return value;
        }
    }
}