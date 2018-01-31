﻿namespace VinEcoAllocatingRemake.AllocatingInventory
{
    #region

    using System;
    using System.Collections.Concurrent;

    #endregion

    /// <summary>
    ///     The utilities.
    /// </summary>
    public partial class Utilities
    {
        /// <summary>
        ///     The dic object double.
        /// </summary>
        private readonly ConcurrentDictionary<object, double> _dicObjectDouble =
            new ConcurrentDictionary<object, double>();

        /// <summary>
        ///     The dic object int.
        /// </summary>
        private readonly ConcurrentDictionary<object, int> _dicObjectInt = new ConcurrentDictionary<object, int>();

        /// <summary>
        ///     The dic string double.
        /// </summary>
        private readonly ConcurrentDictionary<object, double> _dicStringDouble =
            new ConcurrentDictionary<object, double>();

        /// <summary>
        ///     Pretty much a cache for converting double.
        /// </summary>
        /// <param name="obj">
        ///     Object to convert to Double.
        /// </param>
        /// <returns>
        ///     The <see cref="double" />.
        /// </returns>
        public double ObjectToDouble(object obj)
        {
            // Check if exists.
            if (this._dicObjectDouble.TryGetValue(obj, out double value)) return value;

            // Goddamn it.
            if (obj == DBNull.Value)
            {
                this._dicObjectDouble.TryAdd(obj, 0);
                return 0;
            }

            try
            {
                // If not, well, convert.
                value = Convert.ToDouble(obj);
            }
            catch (Exception ex)
            {
                // Idk how this is being hit too frequently.
                // Jk I know now.
                // Debug.WriteLine(ex.Message);
                this._dicObjectDouble.TryAdd(obj, 0);
                return 0;
            }

            // ... and store the result.
            this._dicObjectDouble.TryAdd(obj, value);

            // Then return it.
            return value;
        }

        /// <summary>
        ///     Pretty much a cache for converting Int.
        /// </summary>
        /// <param name="suspect">
        ///     Object to convert to Int32.
        /// </param>
        /// <returns>
        ///     The <see cref="int" />.
        /// </returns>
        public int ObjectToInt(object suspect)
        {
            // Check if exists.
            if (this._dicObjectInt.TryGetValue(suspect, out int value)) return value;

            // Goddamn it.
            if (suspect == DBNull.Value)
            {
                this._dicObjectDouble.TryAdd(suspect, 0);
                return 0;
            }

            // If not, well, convert.
            value = Convert.ToInt32(suspect);

            // ... and store the result.
            this._dicObjectInt.TryAdd(suspect, value);

            // Then return it.
            return value;
        }

        /// <summary>
        ///     Pretty much a cache for converting double.
        /// </summary>
        /// <param name="key">
        ///     String to convert to Double.
        /// </param>
        /// <returns>
        ///     The <see cref="double" />.
        /// </returns>
        public double StringToDouble(string key)
        {
            // Check if exists.
            if (this._dicStringDouble.TryGetValue(key, out double value)) return value;

            // If not, well, convert.
            value = Convert.ToDouble(key);

            // ... and store the result.
            this._dicStringDouble.TryAdd(key, value);

            // Then return it.
            return value;
        }
    }
}