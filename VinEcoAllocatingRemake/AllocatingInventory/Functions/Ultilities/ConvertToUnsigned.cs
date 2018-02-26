using System.Text;
using System.Text.RegularExpressions;

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
        ///     Convert non-ASCII characters in Vietnamese to unsigned, ASCII equivalents.
        /// </summary>
        /// <param name="text">
        ///     The text.
        /// </param>
        /// <returns>
        ///     The <see cref="string" />.
        /// </returns>
        public string ConvertToUnsigned(string text)
        {
            const string excludedChars = "(-)"; // lol.

            for (var i = 33; i < 48; i++)
            {
                if (!excludedChars.Contains(((char) i).ToString()))
                {
                    text = text.Replace(((char) i).ToString(), string.Empty);
                }
            }

            for (var i = 58; i < 65; i++)
            {
                text = text.Replace(((char) i).ToString(), string.Empty);
            }

            for (var i = 91; i < 97; i++)
            {
                text = text.Replace(((char) i).ToString(), string.Empty);
            }

            for (var i = 123; i < 127; i++)
            {
                text = text.Replace(((char) i).ToString(), string.Empty);
            }

            // text = text.Replace(" ", "-"); //Comment lại để không đưa khoảng trắng thành ký tự -
            var regex = new Regex(@"\p{IsCombiningDiacriticalMarks}+");

            string strFormD = text.Normalize(NormalizationForm.FormD);

            return regex.Replace(strFormD, string.Empty).Replace('\u0111', 'd').Replace('\u0110', 'D');
        }
    }
}