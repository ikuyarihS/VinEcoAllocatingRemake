using System;
using System.Collections.Concurrent;

namespace VinEcoAllocatingRemake.AllocatingInventory
{
    public partial class Utilities
    {
        private readonly ConcurrentDictionary<string, DateTime?> _dicStringDate =
            new ConcurrentDictionary<string, DateTime?>(StringComparer.OrdinalIgnoreCase);

        /// <summary>
        ///     Convert string to DateTime.
        ///     Optimization.
        /// </summary>
        /// <param name="suspect">String to convert to Date.</param>
        /// <returns>A DateTime value from a string, if convertible.</returns>
        public DateTime? StringToDate(string suspect)
        {
            // If string has been converted before.
            if (_dicStringDate.TryGetValue(suspect, out DateTime? dateResult))
                return dateResult;
            // Otherwise, check if it's even a date.
            if (!DateTime.TryParse(suspect, out DateTime date))
            {
                // Looks like it isn't.
                // Return null, and also record string used.
                _dicStringDate.TryAdd(suspect, null);
                return null;
            }

            // Welp, it's actually a date.
            // Record the string anyway. Dis many importanto.
            _dicStringDate.TryAdd(suspect, date);
            return date;
        }
    }
}