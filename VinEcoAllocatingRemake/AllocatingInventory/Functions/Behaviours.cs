using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Windows;
using System.Windows.Controls.Primitives;
using System.Windows.Media.Animation;
using Aspose.Cells;
using MongoDB.Driver;
using VinEcoAllocatingRemake.Properties;

//using System.Runtime.Caching;

namespace VinEcoAllocatingRemake.AllocatingInventory
{
    public partial class AllocatingInventory
    {
        private readonly string _applicationPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);

        private readonly BackgroundWorker _bgw = new BackgroundWorker
        {
            WorkerReportsProgress = true,
            WorkerSupportsCancellation = true
        };

        // Optimization stuff.
        private readonly Dictionary<string, string> _dicCoreName =
            new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        private readonly Dictionary<string, (string Region, string Type)> _dicLocation =
            new Dictionary<string, (string Region, string Type)>(StringComparer.OrdinalIgnoreCase);

        private readonly Dictionary<(string coreName, string Region, string Type), string> _dicNewName =
            new Dictionary<(string coreName, string Region, string Type), string>();

        private readonly ExportTableOptions _globalExportTableOptionsopts = new ExportTableOptions
        {
            CheckMixedValueType = true,
            ExportAsString = false,
            FormatStrategy = CellValueFormatStrategy.None,
            ExportColumnName = true
        };

        private readonly Utilities _ulti = new Utilities();

        //private ObjectCache _cache = MemoryCache.Default;

        private bool _isBackgroundworkerIdle = true;

        /// <summary>
        ///     A simple Function to open the folder where the program is.
        ///     Quality of life.
        /// </summary>
        private void OpenApplicationPath(object sender, RoutedEventArgs e)
        {
            try
            {
                Process[] processExcel = Process.GetProcessesByName("excel");

                foreach (Process process in processExcel)
                    process.Kill();

                WriteToRichTextBoxOutput("Vừng ơi mở ra!!!");

                //if (_applicationPath == null)
                //{
                //    WriteToRichTextBoxOutput("Có lỗi xảy ra, không mở được thư mục!");
                //    return;
                //}

                Process.Start(_applicationPath);
            }
            catch (Exception ex)
            {
                WriteToRichTextBoxOutput(ex.Message);
                throw;
            }
        }

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
        ///     Memorize last openned page. Quality of Life.
        /// </summary>
        private void ScoutingPrice_OnLoaded(object sender, RoutedEventArgs e)
        {
            Settings.Default.LastPage = "AllocatingInventory/Pages/AllocatingPage.xaml";
        }

        private void Initializer()
        {
            try
            {
                _bgw.ProgressChanged += BackgroundWorker_ProcessChanged;

                _bgw.RunWorkerCompleted += BackgroundWorker_RunWorkerCompleted;

                ProgressStatusBarLabel.Text = string.Empty;

                //WriteToRichTextBoxOutput("Khởi động MongoDB", 1);

                ////starting the mongod server (when app starts)
                //var start = new ProcessStartInfo
                //{
                //    FileName = $@"{_applicationPath}\mongod.exe",
                //    WindowStyle = ProcessWindowStyle.Hidden,
                //    UseShellExecute = false,
                //    Arguments = $@"--dbpath {_applicationPath}\MongoDB"
                //};
                //// set UseShellExecute to false

                //Process mongod = Process.Start(start);

                //// Mongo CSharp Driver Code (see Mongo docs)
                //var client = new MongoClient();
                //IMongoDatabase database = client.GetDatabase("ChiaHangRemake");

                WriteToRichTextBoxOutput();
                WriteToRichTextBoxOutput("Sẵn sàng oánh nhau!", 1);

                //MouseDown += Window_MouseDown;
            }
            catch (Exception ex)
            {
                WriteToRichTextBoxOutput(ex.Message);
                throw;
            }
        }

        private void BackgroundWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            //_bgw.DoWork -= DataHandler;
            //_bgw.DoWork -= ProcessData;
            _isBackgroundworkerIdle = true;
            WriteToRichTextBoxOutput("Done!");
            WriteToRichTextBoxOutput();
        }

        private void BackgroundWorker_ProcessChanged(object sender, ProgressChangedEventArgs e)
        {
            if (ProgressStatusBar.Value >= 1) ProgressStatusBar.Value = 0;
            ProgressStatusBar.BeginAnimation(RangeBase.ValueProperty,
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
                    ProgressStatusBarLabel.Text = $"{e.ProgressPercentage}%";
                    if (mainWindow != null) mainWindow.MyTaskBarInfo.ProgressValue = e.ProgressPercentage / 100d;
                    break;
            }
        }

        private void HereWeGo(object sender, RoutedEventArgs e)
        {
            if (!_bgw.IsBusy)
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

        private void ProcessData(object sender, RoutedEventArgs e)
        {
            if (!_bgw.IsBusy)
            {
                if (_isBackgroundworkerIdle) _isBackgroundworkerIdle = false;

                _bgw.RunWorkerAsync();
            }
            else
            {
                MessageBox.Show("Đang uýnh nhau, đợi xíu!");
            }
        }
    }
}