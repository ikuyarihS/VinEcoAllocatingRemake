#region

using System.Collections.Generic;
using System.Globalization;

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
        ///     The dic string proper.
        /// </summary>
        private readonly Dictionary<string, string> dicStringProper = new Dictionary<string, string>();

        /// <summary>
        ///     Proper a string
        /// </summary>
        /// <param name="myString">
        ///     The my String.
        /// </param>
        /// <returns>
        ///     The <see cref="string" />.
        /// </returns>
        public string ProperStr(string myString)
        {
            if (dicStringProper.TryGetValue(myString, out string myProperString)) return myProperString;

            // Creates a TextInfo based on the "en-US" culture.
            TextInfo myTi = new CultureInfo("en-US", false).TextInfo;

            //// Changes a string to lowercase.
            // WriteToRichTextBoxOutput("\"{0}\" to lowercase: {1}", myString, myTI.ToLower(myString));

            //// Changes a string to uppercase.
            // WriteToRichTextBoxOutput("\"{0}\" to uppercase: {1}", myString, myTI.ToUpper(myString));

            //// Changes a string to titlecase.
            // WriteToRichTextBoxOutput("\"{0}\" to titlecase: {1}", myString, myTI.ToTitleCase(myString));
            myProperString = myTi.ToTitleCase(myString.ToLower());
            dicStringProper.Add(myString, myProperString);
            return myProperString;
        }
    }
}