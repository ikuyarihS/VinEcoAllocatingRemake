#region

using System.ComponentModel;
using System.Windows.Shell;
using VinEcoAllocatingRemake.Properties;

#endregion

namespace VinEcoAllocatingRemake
{
    /// <summary>
    ///     Interaction logic for MainWindow.xaml
    /// </summary>
    // ReSharper disable once InheritdocConsiderUsage
    public partial class MainWindow
    {
        public MainWindow()
        {
            InitializeComponent();

            MyTaskBarInfo.ProgressState = TaskbarItemProgressState.Normal;
        }

        protected override void OnClosing(CancelEventArgs e)
        {
            Settings.Default.Save();
            base.OnClosing(e);
        }
    }
}