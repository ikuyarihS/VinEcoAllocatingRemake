using System;
using System.ComponentModel;

namespace VinEcoAllocatingRemake.AllocatingInventory
{
    public partial class AllocatingInventory
    {
        private void ReadPurchaseOrder(object sender, DoWorkEventArgs e)
        {
            try
            {
            }
            catch (Exception ex)
            {
                WriteToRichTextBoxOutput(ex.Message);
                throw;
            }
        }
    }
}