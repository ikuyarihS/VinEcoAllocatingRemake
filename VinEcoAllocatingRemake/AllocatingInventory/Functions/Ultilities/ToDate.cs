#region

using System;
using System.Collections.Concurrent;

#endregion

namespace VinEcoAllocatingRemake.AllocatingInventory
{
    /// <summary>
    ///     The utilities.
    /// </summary>
    public partial class Utilities
    {
        /// <summary>
        ///     The dic string date.
        /// </summary>
        private readonly ConcurrentDictionary<string, DateTime> dicStringDate =
            new ConcurrentDictionary<string, DateTime>(StringComparer.OrdinalIgnoreCase);

        /// <summary>
        ///     Convert string to DateTime.
        ///     Optimization.
        /// </summary>
        /// <param name="suspect">String to convert to Date.</param>
        /// <returns>A DateTime value from a string, if convertible.</returns>
        public DateTime? StringToDate(string suspect)
        {
            // If string has been converted before.
            if (dicStringDate.TryGetValue(suspect, out DateTime dateResult)) return dateResult;

            // Otherwise, check if it's even a date.
            if (!DateTime.TryParse(suspect, out DateTime date))
            {
                // Looks like it isn't.
                // Return null, and also record string used.
                dicStringDate.TryAdd(
                    suspect,
                    DateTime.MinValue);
                return null;
            }

            // Welp, it's actually a date.
            // Record the string anyway. Dis many importanto.
            dicStringDate.TryAdd(
                suspect,
                date);
            return date;
        }
    }
}