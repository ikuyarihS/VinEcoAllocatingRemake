#region

using System;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Windows;
using System.Windows.Controls.Primitives;
using System.Windows.Media.Animation;
using Aspose.Cells;
using VinEcoAllocatingRemake.Properties;

#endregion

namespace VinEcoAllocatingRemake.AllocatingInventory
{
    #region

    #endregion

    /// <summary>
    ///     The allocating inventory.
    /// </summary>
    public partial class AllocatingInventory
    {
        /// <summary>
        ///     The application path.
        /// </summary>
        private readonly string _applicationPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);

        /// <summary>
        ///     The bgw.
        /// </summary>
        private readonly BackgroundWorker _bgw =
            new BackgroundWorker {WorkerReportsProgress = true, WorkerSupportsCancellation = true};

        /// <summary>
        ///     The global export table options opts.
        /// </summary>
        private readonly ExportTableOptions _globalExportTableOptionsOpts =
            new ExportTableOptions
            {
                CheckMixedValueType = true,
                ExportAsString = false,
                FormatStrategy = CellValueFormatStrategy.None,
                ExportColumnName = true
            };

        /// <summary>
        ///     The ulti.
        /// </summary>
        private readonly Utilities _ulti = new Utilities();

        /// <summary>
        ///     The is backgroundworker idle.
        /// </summary>
        private bool _isBackgroundworkerIdle = true;

        /// <summary>
        ///     The background worker process changed.
        /// </summary>
        /// <param name="sender">
        ///     The sender.
        /// </param>
        /// <param name="e">
        ///     The e.
        /// </param>
        private void BackgroundWorkerProcessChanged(object sender, ProgressChangedEventArgs e)
        {
            if (ProgressStatusBar.Value >= 1) ProgressStatusBar.Value = 0;

            ProgressStatusBar.BeginAnimation(
                RangeBase.ValueProperty,
                new DoubleAnimation(e.ProgressPercentage, new Duration(TimeSpan.FromSeconds(1))));

            var mainWindow = Application.Current.MainWindow as MainWindow;

            switch (e.ProgressPercentage)
            {
                case 0:
                    ProgressStatusBarLabel.Text = string.Empty;
                    break;
                case 100:
                    ProgressStatusBarLabel.Text = "Done!";
                    if (mainWindow != null) mainWindow.MyTaskBarInfo.ProgressValue = 0;

                    break;
                default:
                    ProgressStatusBarLabel.Text = $"{e.ProgressPercentage.ToString(string.Empty)}%";
                    if (mainWindow != null) mainWindow.MyTaskBarInfo.ProgressValue = e.ProgressPercentage / 100d;

                    break;
            }
        }

        /// <summary>
        ///     The background worker run worker completed.
        /// </summary>
        /// <param name="sender">
        ///     The sender.
        /// </param>
        /// <param name="e">
        ///     The e.
        /// </param>
        private void BackgroundWorkerRunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            _bgw.DoWork -= ReadForecast;
            _bgw.DoWork -= ReadPurchaseOrder;
            _isBackgroundworkerIdle = true;
            WriteToRichTextBoxOutput("Done!");
            WriteToRichTextBoxOutput();
        }

        /// <summary>
        ///     The cancel_ on click.
        /// </summary>
        /// <param name="sender">
        ///     The sender.
        /// </param>
        /// <param name="e">
        ///     The e.
        /// </param>
        private void Cancel_OnClick(object sender, RoutedEventArgs e)
        {
            if (!_bgw.IsBusy)
            {
                WriteToRichTextBoxOutput("Ủa còn chưa kịp làm gì mà :<");
                return;
            }

            _bgw.CancelAsync();

            if (Application.Current.MainWindow is MainWindow mainWindow) mainWindow.MyTaskBarInfo.ProgressValue = 0;

            ProgressStatusBar.Value = 0;
            ProgressStatusBarLabel.Text = "Canceled!";

            WriteToRichTextBoxOutput();
            WriteToRichTextBoxOutput("Hoãn! Hoãn ngay! Không có chơi bời gì hết nữa!");
            WriteToRichTextBoxOutput();
        }

        /// <summary>
        ///     The fite moi handler.
        /// </summary>
        /// <param name="sender">
        ///     The sender.
        /// </param>
        /// <param name="e">
        ///     The e.
        /// </param>
        private void FiteMoiHandler(object sender, RoutedEventArgs e)
        {
            if (!_bgw.IsBusy && _isBackgroundworkerIdle)
            {
                if (_isBackgroundworkerIdle)
                {
                    _bgw.DoWork += FiteMoi;
                    _isBackgroundworkerIdle = false;
                }

                _bgw.RunWorkerAsync();
            }
            else
            {
                MessageBox.Show("Đang uýnh nhau, đợi xíu!");
            }
        }

        /// <summary>
        ///     The forecast handler.
        /// </summary>
        /// <param name="sender">
        ///     The sender.
        /// </param>
        /// <param name="e">
        ///     The e.
        /// </param>
        private void ForecastHandler(object sender, RoutedEventArgs e)
        {
            if (!_bgw.IsBusy && _isBackgroundworkerIdle)
            {
                if (_isBackgroundworkerIdle)
                {
                    _bgw.DoWork += ReadForecast;
                    _isBackgroundworkerIdle = false;
                }

                _bgw.RunWorkerAsync();
            }
            else
            {
                MessageBox.Show("Đang uýnh nhau, đợi xíu!");
            }
        }

        /// <summary>
        ///     The initializer.
        /// </summary>
        private void Initializer()
        {
            _bgw.ProgressChanged += BackgroundWorkerProcessChanged;

            _bgw.RunWorkerCompleted += BackgroundWorkerRunWorkerCompleted;

            ProgressStatusBarLabel.Text = string.Empty;

            WriteToRichTextBoxOutput();
            WriteToRichTextBoxOutput("Sẵn sàng oánh nhau!", 1);
        }

        /// <summary>
        ///     A simple Function to open the folder where the program is.
        ///     Quality of life.
        /// </summary>
        /// <param name="sender">
        ///     The sender.
        /// </param>
        /// <param name="e">
        ///     The e.
        /// </param>
        private void OpenApplicationPath(object sender, RoutedEventArgs e)
        {
            Process[] processExcel = Process.GetProcessesByName("excel");

            foreach (Process process in processExcel) process.Kill();

            WriteToRichTextBoxOutput("Vừng ơi mở ra!!!");

            Process.Start(_applicationPath);
        }

        /// <summary>
        ///     The order handler.
        /// </summary>
        /// <param name="sender">
        ///     The sender.
        /// </param>
        /// <param name="e">
        ///     The e.
        /// </param>
        private void OrderHandler(object sender, RoutedEventArgs e)
        {
            if (!_bgw.IsBusy && _isBackgroundworkerIdle)
            {
                if (_isBackgroundworkerIdle)
                {
                    _bgw.DoWork += ReadPurchaseOrder;
                    _isBackgroundworkerIdle = false;
                }

                _bgw.RunWorkerAsync();
            }
            else
            {
                MessageBox.Show("Đang uýnh nhau, đợi xíu!");
            }
        }

        /// <summary>
        ///     Memorize last opened page. Quality of Life.
        /// </summary>
        /// <param name="sender">
        ///     The sender.
        /// </param>
        /// <param name="e">
        ///     The e.
        /// </param>
        private void ScoutingPrice_OnLoaded(object sender, RoutedEventArgs e)
        {
            Settings.Default.LastPage = "AllocatingInventory/Pages/AllocatingPage.xaml";
        }

        // private void ProcessData(object sender, RoutedEventArgs e)
        // {
        // if (!_bgw.IsBusy)
        // {
        // if (_isBackgroundworkerIdle) _isBackgroundworkerIdle = false;

        // _bgw.RunWorkerAsync();
        // }
        // else
        // {
        // MessageBox.Show("Đang uýnh nhau, đợi xíu!");
        // }
        // }
    }
}