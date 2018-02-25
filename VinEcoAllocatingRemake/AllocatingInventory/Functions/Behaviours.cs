// --------------------------------------------------------------------------------------------------------------------
// <copyright file="Behaviours.cs" company="VinEco">
//   Shirayuki 2018.
// </copyright>
// <summary>
//   The allocating inventory.
// </summary>
// --------------------------------------------------------------------------------------------------------------------

using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Reflection;
using System.Windows;
using System.Windows.Controls.Primitives;
using System.Windows.Media.Animation;
using Aspose.Cells;
using VinEcoAllocatingRemake.Properties;

namespace VinEcoAllocatingRemake.AllocatingInventory
    {
        #region

        #endregion

        /// <summary>
        ///     The allocating inventory.
        /// </summary>
        // ReSharper disable once StyleCop.SA1404
        [SuppressMessage("ReSharper", "ArrangeThisQualifier")]
        public partial class AllocatingInventory
            {
                /// <summary>
                ///     The application path.
                /// </summary>
                private readonly string applicationPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);

                /// <summary>
                ///     The bgw.
                /// </summary>
                private readonly BackgroundWorker bgw            =
                    new BackgroundWorker { WorkerReportsProgress = true, WorkerSupportsCancellation = true };

                /// <summary>
                ///     The global export table options opts.
                /// </summary>
                private readonly ExportTableOptions globalExportTableOptionsOpts =
                    new ExportTableOptions
                        {
                            CheckMixedValueType = true,
                            ExportAsString      = false,
                            FormatStrategy      = CellValueFormatStrategy.None,
                            ExportColumnName    = true
                        };

                /// <summary>
                ///     The ulti.
                /// </summary>
                private readonly Utilities ulti = new Utilities();

                /// <summary>
                ///     The is backgroundworker idle.
                /// </summary>
                private bool isBackgroundworkerIdle = true;

                /// <summary>
                ///     The background worker process changed.
                /// </summary>
                /// <param name="sender"> The sender. </param>
                /// <param name="e"> The e. </param>
                private void BackgroundWorkerProcessChanged(object sender, ProgressChangedEventArgs e)
                    {
                        if (this.ProgressStatusBar.Value >= 1)
                            {
                                this.ProgressStatusBar.Value = 0;
                            }

                        this.ProgressStatusBar.BeginAnimation(
                            RangeBase.ValueProperty,
                            new DoubleAnimation(e.ProgressPercentage, new Duration(TimeSpan.FromSeconds(1))));

                        var mainWindow = Application.Current.MainWindow as MainWindow;

                        switch (e.ProgressPercentage)
                            {
                                case 0:
                                    this.ProgressStatusBarLabel.Text = string.Empty;
                                    break;
                                case 100:
                                    this.ProgressStatusBarLabel.Text = "Done!";
                                    if (mainWindow != null)
                                        {
                                            mainWindow.MyTaskBarInfo.ProgressValue = 0;
                                        }

                                    break;
                                default:
                                    this.ProgressStatusBarLabel.Text = $"{e.ProgressPercentage.ToString(string.Empty)}%";
                                    if (mainWindow != null)
                                        {
                                            mainWindow.MyTaskBarInfo.ProgressValue = e.ProgressPercentage / 100d;
                                        }

                                    break;
                            }
                    }

                /// <summary>
                ///     The background worker run worker completed.
                /// </summary>
                /// <param name="sender"> The sender. </param>
                /// <param name="e"> The e. </param>
                private void BackgroundWorkerRunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
                    {
                        this.bgw.DoWork             -= this.FiteMoi;
                        this.bgw.DoWork             -= this.ReadForecast;
                        this.bgw.DoWork             -= this.ReadPurchaseOrder;
                        this.isBackgroundworkerIdle =  true;
                        this.WriteToRichTextBoxOutput("Done!");
                        this.WriteToRichTextBoxOutput();
                        this.TryClear();
                    }

                /// <summary>
                ///     The cancel_ on click.
                /// </summary>
                /// <param name="sender"> The sender. </param>
                /// <param name="e"> The e. </param>
                private void Cancel_OnClick(object sender, RoutedEventArgs e)
                    {
                        if (!this.bgw.IsBusy)
                            {
                                this.WriteToRichTextBoxOutput("Ủa còn chưa kịp làm gì mà :<");
                                return;
                            }

                        this.bgw.CancelAsync();

                        if (Application.Current.MainWindow is MainWindow mainWindow)
                            {
                                mainWindow.MyTaskBarInfo.ProgressValue = 0;
                            }

                        this.ProgressStatusBar.Value     = 0;
                        this.ProgressStatusBarLabel.Text = "Canceled!";

                        this.WriteToRichTextBoxOutput();
                        this.WriteToRichTextBoxOutput("Hoãn! Hoãn ngay! Không có chơi bời gì hết nữa!");
                        this.WriteToRichTextBoxOutput();
                    }

                /// <summary>
                ///     The fite moi handler.
                /// </summary>
                /// <param name="sender"> The sender. </param>
                /// <param name="e"> The e. </param>
                private void FiteMoiHandler(object sender, RoutedEventArgs e)
                    {
                        if (this.bgw.IsBusy || !this.isBackgroundworkerIdle)
                            {
                                MessageBox.Show("Đang uýnh nhau, đợi xíu!");
                                return;
                            }

                        if (this.isBackgroundworkerIdle)
                            {
                                this.bgw.DoWork += this.FiteMoi;
                            }

                        this.isBackgroundworkerIdle = false;
                        this.bgw.RunWorkerAsync();
                    }

                /// <summary>
                ///     The forecast handler.
                /// </summary>
                /// <param name="sender"> The sender. </param>
                /// <param name="e"> The e. </param>
                private void ForecastHandler(object sender, RoutedEventArgs e)
                    {
                        if (!this.bgw.IsBusy && this.isBackgroundworkerIdle)
                            {
                                if (this.isBackgroundworkerIdle)
                                    {
                                        this.bgw.DoWork             += this.ReadForecast;
                                        this.isBackgroundworkerIdle =  false;
                                    }

                                this.bgw.RunWorkerAsync();
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
                        this.bgw.ProgressChanged += this.BackgroundWorkerProcessChanged;

                        this.bgw.RunWorkerCompleted += this.BackgroundWorkerRunWorkerCompleted;

                        this.ProgressStatusBarLabel.Text = string.Empty;

                        this.WriteToRichTextBoxOutput();
                        this.WriteToRichTextBoxOutput("Sẵn sàng oánh nhau!", 1);
                    }

                /// <summary>
                ///     A simple Function to open the folder where the program is.
                ///     Quality of life.
                /// </summary>
                /// <param name="sender"> The sender. </param>
                /// <param name="e"> The e. </param>
                private void OpenApplicationPath(object sender, RoutedEventArgs e)
                    {
                        Process[] processExcel = Process.GetProcessesByName("excel");

                        foreach (Process process in processExcel)
                            {
                                process.Kill();
                            }

                        this.WriteToRichTextBoxOutput("Vừng ơi mở ra!!!");

                        Process.Start(this.applicationPath);
                    }

                /// <summary>
                ///     The order handler.
                /// </summary>
                /// <param name="sender"> The sender. </param>
                /// <param name="e"> The e. </param>
                private void OrderHandler(object sender, RoutedEventArgs e)
                    {
                        if (!this.bgw.IsBusy && this.isBackgroundworkerIdle)
                            {
                                if (this.isBackgroundworkerIdle)
                                    {
                                        this.bgw.DoWork             += this.ReadPurchaseOrder;
                                        this.isBackgroundworkerIdle =  false;
                                    }

                                this.bgw.RunWorkerAsync();
                            }
                        else
                            {
                                MessageBox.Show("Đang uýnh nhau, đợi xíu!");
                            }
                    }

                /// <summary>
                ///     Memorize last opened page. Quality of Life.
                /// </summary>
                /// <param name="sender"> The sender. </param>
                /// <param name="e"> The e. </param>
                private void ScoutingPrice_OnLoaded(object sender, RoutedEventArgs e)
                    {
                        Settings.Default.LastPage = "AllocatingInventory/Pages/AllocatingPage.xaml";
                    }
            }
    }