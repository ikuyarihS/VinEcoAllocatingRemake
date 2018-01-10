using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Media;

namespace VinEcoAllocatingRemake.AllocatingInventory
{
    public partial class AllocatingInventory
    {
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
        private void WriteToRichTextBoxOutput(object message = null, byte importanceLevel = 0, bool newLine = true,
            bool hasTimeStamp = true)
        {
            void Action()
            {
                try
                {
                    var textRange = new TextRange(RichTextBoxOutput.Document.ContentEnd,
                        RichTextBoxOutput.Document.ContentEnd);
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

                    if (hasTimeStamp && message != null && message.ToString() != string.Empty)
                        ExtraTimeStamp();

                    textRange.Text =
                        $"{(hasTimeStamp ? extraMessage : string.Empty)}{message}{(newLine ? "\r" : " ")}";
                    textRange.ApplyPropertyValue(TextElement.ForegroundProperty,
                        brushConverter.ConvertFromString("Cornflowerblue") ??
                        throw new InvalidOperationException("What the heck?"));
                }

                catch (Exception ex)
                {
                    WriteToRichTextBoxOutput(ex.Message);
                    throw;
                }
            }

            Application.Current.Dispatcher.BeginInvoke((Action) Action);
        }

        private void RichTextBoxOutput_TextChanged(object sender, TextChangedEventArgs e)
        {
            RichTextBoxOutput.ScrollToEnd();
        }

        private void ExtraTimeStamp()
        {
            try
            {
                var textRange = new TextRange(RichTextBoxOutput.Document.ContentEnd,
                    RichTextBoxOutput.Document.ContentEnd) {Text = DateTime.Now.ToString("[HH:mm] ")};

                textRange.ApplyPropertyValue(TextElement.ForegroundProperty,
                    new BrushConverter().ConvertFromString("ForestGreen") ??
                    throw new InvalidOperationException("What the heck?"));
            }

            catch (Exception ex)
            {
                WriteToRichTextBoxOutput(ex.Message);
                throw;
            }
        }
    }
}