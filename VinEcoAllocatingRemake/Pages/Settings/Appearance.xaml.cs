namespace VinEcoAllocatingRemake.Pages.Settings
{
    #region

    using System.Windows.Controls;

    #endregion

    /// <summary>
    ///     Interaction logic for Appearance.xaml
    /// </summary>
    public partial class Appearance : UserControl
    {
        public Appearance()
        {
            this.InitializeComponent();

            // create and assign the appearance view model
            this.DataContext = new AppearanceViewModel();
        }
    }
}