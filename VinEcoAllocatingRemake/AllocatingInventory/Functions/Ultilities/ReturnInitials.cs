﻿using System.Collections.Concurrent;
using System.Text;

namespace VinEcoAllocatingRemake.AllocatingInventory
{
    public partial class Utilities
    {
        private readonly ConcurrentDictionary<string, string> _dicStringInitials =
            new ConcurrentDictionary<string, string>();

        public string ReturnInitials(string suspect)
        {
            // Uhm, have we met before?
            if (_dicStringInitials.TryGetValue(suspect, out string result))
                return result;

            // Oh ok. Here's a punch.

            var resultToBe = new StringBuilder();
            var yesAppend = true;

            foreach (char c in suspect)
            {
                if (yesAppend) resultToBe.Append(c);
                yesAppend = c == ' ';
            }

            result = resultToBe.ToString();

            //result = string.Join(string.Empty, suspect.Split(' ').Select(x => x.First()));
            // It was super effective.
            _dicStringInitials.TryAdd(suspect, result);
            return result;

        }
    }
}