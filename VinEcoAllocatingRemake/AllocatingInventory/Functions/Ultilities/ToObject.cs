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
        ///     The dic bool object.
        /// </summary>
        private readonly ConcurrentDictionary<bool, object> dicBoolObject = new ConcurrentDictionary<bool, object>();

        /// <summary>
        ///     The dic date object.
        /// </summary>
        private readonly ConcurrentDictionary<DateTime, object> dicDateObject = new ConcurrentDictionary<DateTime, object>();

        /// <summary>
        ///     The dic double object.
        /// </summary>
        private readonly ConcurrentDictionary<double, object> dicDoubleObject = new ConcurrentDictionary<double, object>();

        /// <summary>
        ///     The dic int object.
        /// </summary>
        private readonly ConcurrentDictionary<int, object> dicIntObject = new ConcurrentDictionary<int, object>();

        /// <summary>
        ///     Convert Boolean to Object.
        ///     Optimization.
        /// </summary>
        /// <param name="suspect">
        ///     The suspect.
        /// </param>
        /// <returns>
        ///     The <see cref="object" />.
        /// </returns>
        public object BoolToObject(bool suspect)
        {
            // If string has been converted before.
            if (dicBoolObject.TryGetValue(suspect, out object obj)) return obj;

            obj = GetString(suspect.ToString());

            // Welp, it's actually a bool.
            // Record the string anyway. Dis many importanto.
            dicBoolObject.TryAdd(suspect, obj);
            return obj;
        }

        /// <summary>
        ///     Convert DateTime to Object.
        ///     Optimization.
        /// </summary>
        /// <param name="suspect">
        ///     The suspect.
        /// </param>
        /// <param name="dateFormat">
        ///     The date Format.
        /// </param>
        /// <returns>
        ///     The <see cref="object" />.
        /// </returns>
        public object DateToObject(DateTime suspect, string dateFormat = "")
        {
            // If string has been converted before.
            if (dicDateObject.TryGetValue(suspect, out object obj)) return obj;

            obj = GetString(suspect.ToString(dateFormat));

            // Welp, it's actually a date.
            // Record the string anyway. Dis many importanto.
            dicDateObject.TryAdd(suspect, obj);
            return obj;
        }

        /// <summary>
        ///     Convert double to Object.
        ///     Optimization.
        /// </summary>
        /// <param name="suspect">
        ///     The suspect.
        /// </param>
        /// <returns>
        ///     The <see cref="object" />.
        /// </returns>
        public object DoubleToObject(double suspect)
        {
            // If string has been converted before.
            if (dicDoubleObject.TryGetValue(suspect, out object obj)) return obj;

            // This feels like cheating tbh.
            obj = GetString(suspect.ToString(string.Empty));

            // Welp, it's actually a date.
            // Record the string anyway. Dis many importanto.
            dicDoubleObject.TryAdd(suspect, obj);
            return obj;
        }

        /// <summary>
        ///     Convert Int to Object.
        ///     Optimization.
        /// </summary>
        /// <param name="suspect">
        ///     The suspect.
        /// </param>
        /// <returns>
        ///     The <see cref="object" />.
        /// </returns>
        public object IntToObject(int suspect)
        {
            // If string has been converted before.
            if (dicIntObject.TryGetValue(suspect, out object obj)) return obj;

            // Definitely cheating.
            obj = GetString(suspect.ToString());

            // Welp, it's actually a date.
            // Record the string anyway. Dis many importanto.
            dicIntObject.TryAdd(suspect, obj);
            return obj;
        }
    }
}