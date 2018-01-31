namespace VinEcoAllocatingRemake
{
    #region

    using System.ComponentModel;
    using System.Windows.Shell;

    using VinEcoAllocatingRemake.Properties;

    #endregion

    /// <summary>
    ///     Interaction logic for MainWindow.xaml
    /// </summary>
    // ReSharper disable once InheritdocConsiderUsage
    public partial class MainWindow
    {
        /// <summary>
        ///     Initializes a new instance of the <see cref="MainWindow" /> class.
        /// </summary>
        public MainWindow()
        {
            this.InitializeComponent();

            this.MyTaskBarInfo.ProgressState = TaskbarItemProgressState.Normal;
        }

        /// <summary>
        ///     The on closing.
        /// </summary>
        /// <param name="e"> The e. </param>
        protected override void OnClosing(CancelEventArgs e)
        {
            Settings.Default.Save();
            base.OnClosing(e);
        }
    }
}