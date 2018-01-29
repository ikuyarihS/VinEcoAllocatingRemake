#region

using System;
using System.Collections.Concurrent;

#endregion

namespace VinEcoAllocatingRemake.AllocatingInventory
{
    #region

    #endregion

    /// <summary>
    ///     The utilities.
    /// </summary>
    public partial class Utilities
    {
        /// <summary>
        ///     The dic date string.
        /// </summary>
        private readonly ConcurrentDictionary<DateTime, string> dicDateString = new ConcurrentDictionary<DateTime, string>();

        /// <summary>
        ///     The dic object string.
        /// </summary>
        private readonly ConcurrentDictionary<object, string> dicObjectString = new ConcurrentDictionary<object, string>();

        /// <summary>
        ///     The dic string.
        /// </summary>
        private readonly ConcurrentDictionary<string, string> dicString = new ConcurrentDictionary<string, string>();

        /// <summary>
        ///     Pretty much a cache for converting DateTime to String.
        /// </summary>
        /// <param name="date">
        ///     The date.
        /// </param>
        /// <param name="dateFormat">
        ///     The date Format.
        /// </param>
        /// <returns>
        ///     The <see cref="string" />.
        /// </returns>
        public string DateToString(DateTime date, string dateFormat = "")
        {
            // Check if exists.
            if (dicDateString.TryGetValue(date, out string value)) return GetString(value);

            // If not, well, convert.
            value = date.ToString(dateFormat);

            // ... and store the result.
            dicDateString.TryAdd(date, value);

            // Then return it.
            return GetString(value);
        }

        /// <summary>
        ///     The get string.
        /// </summary>
        /// <param name="suspect">
        ///     The suspect.
        /// </param>
        /// <returns>
        ///     The <see cref="string" />.
        /// </returns>
        public string GetString(string suspect)
        {
            if (dicString.TryGetValue(suspect, out string result)) return result;

            dicString.TryAdd(suspect, suspect);

            return suspect;
        }

        /// <summary>
        ///     The object to string.
        /// </summary>
        /// <param name="obj">
        ///     The obj.
        /// </param>
        /// <returns>
        ///     The <see cref="string" />.
        /// </returns>
        public string ObjectToString(object obj)
        {
            // Check if exists.
            if (dicObjectString.TryGetValue(obj, out string value)) return GetString(value);

            // If not, well, convert.
            if (dicString.TryGetValue(obj.ToString(), out string valueNew))
            {
                dicObjectString.TryAdd(obj, valueNew);
                return GetString(valueNew);
            }

            value = obj.ToString();
            dicString.TryAdd(value, value);
            dicObjectString.TryAdd(obj, value);
            return GetString(value);

            // value =  ? valueNew : obj.ToString();

            // ... and store the result.
            // _dicObjectString.TryAdd(obj, value);

            // Then return it.
            // return value;
        }
    }
}