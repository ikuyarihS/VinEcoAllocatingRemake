using System.Diagnostics.CodeAnalysis;

namespace VinEcoAllocatingRemake.AllocatingInventory
{
    #region

    using System;
    using System.Windows;
    using System.Windows.Controls;
    using System.Windows.Documents;
    using System.Windows.Media;

    #endregion

    #region

    #endregion

    /// <summary>
    ///     The allocating inventory.
    /// </summary>
    [SuppressMessage("ReSharper", "ArrangeThisQualifier")]
    public partial class AllocatingInventory
    {
        /// <summary>
        ///     The extra time stamp.
        /// </summary>
        private void ExtraTimeStamp()
        {
            var textRange = new TextRange(
                                    this.RichTextBoxOutput.Document.ContentEnd,
                                    this.RichTextBoxOutput.Document.ContentEnd)
                                    {
                                        Text = DateTime.Now.ToString(
                                            "[HH:mm] ")
                                    };

            textRange.ApplyPropertyValue(TextElement.ForegroundProperty, new BrushConverter().ConvertFromString("ForestGreen") ?? throw new InvalidOperationException("What the heck?"));
        }

        /// <summary>
        ///     The rich text box output text changed.
        /// </summary>
        /// <param name="sender"> The sender. </param>
        /// <param name="e"> The e. </param>
        private void RichTextBoxOutputTextChanged(object sender, TextChangedEventArgs e)
        {
            this.RichTextBoxOutput.ScrollToEnd();
        }

        /// <summary>
        ///     The try clear.
        /// </summary>
        private void TryClear()
        {
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        /// <summary>
        ///     Appending desired text into Output RichTextBox.
        ///     By default, it will be in a new line.
        /// </summary>
        /// <param name="message">Message for RichTextBox to Append.</param>
        /// <param name="importanceLevel">
        ///     Level of Importance.
        ///     0 = Default.
        ///     1 = Very importanto.
        ///     2 = Meh.
        /// </param>
        /// <param name="newLine">A seperated new line?</param>
        /// <param name="hasTimeStamp">Include Time Stamp</param>
        private void WriteToRichTextBoxOutput(
            object message = null,
            byte importanceLevel = 0,
            bool newLine = true,
            bool hasTimeStamp = true)
        {
            void Action()
            {
                var textRange = new TextRange(
                    this.RichTextBoxOutput.Document.ContentEnd,
                    this.RichTextBoxOutput.Document.ContentEnd);
                var brushConverter = new BrushConverter();

                string extraMessage = string.Empty;
                if (message == null || message.ToString() == string.Empty) message = string.Empty;

                switch (importanceLevel)
                {
                    case 0:
                        break;
                    case 1:
                        extraMessage = "!!! - ";
                        break;
                    case 2:
                        extraMessage = "      ";
                        break;
                    default:
                        extraMessage = string.Empty;
                        break;
                }

                if (hasTimeStamp && message != null && message.ToString() != string.Empty) this.ExtraTimeStamp();

                textRange.Text = $"{(hasTimeStamp ? extraMessage : string.Empty)}{message}{(newLine ? "\r" : " ")}";
                textRange.ApplyPropertyValue(TextElement.ForegroundProperty, brushConverter.ConvertFromString("Cornflowerblue") ?? throw new InvalidOperationException("What the heck?"));
            }

            Application.Current.Dispatcher.BeginInvoke((Action)Action);
        }
    }
}