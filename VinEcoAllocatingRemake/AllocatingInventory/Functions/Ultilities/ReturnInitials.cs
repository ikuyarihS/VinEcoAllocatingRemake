﻿#region

using System.Collections.Concurrent;
using System.Text;

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
        ///     The dic string initials.
        /// </summary>
        private readonly ConcurrentDictionary<string, string> dicStringInitials =
            new ConcurrentDictionary<string, string>();

        /// <summary>
        ///     The return initials.
        /// </summary>
        /// <param name="suspect">
        ///     The suspect.
        /// </param>
        /// <returns>
        ///     The <see cref="string" />.
        /// </returns>
        public string ReturnInitials(string suspect)
        {
            // Uhm, have we met before?
            if (dicStringInitials.TryGetValue(suspect, out string result)) return result;

            // Oh ok. Here's a punch.
            var resultToBe = new StringBuilder();
            var yesAppend = true;

            foreach (char c in suspect)
            {
                if (yesAppend) resultToBe.Append(c);

                yesAppend = c == ' ';
            }

            result = resultToBe.ToString();

            // result = string.Join(string.Empty, suspect.Split(' ').Select(x => x.First()));
            // It was super effective.
            dicStringInitials.TryAdd(suspect, result);
            return result;
        }
    }
}