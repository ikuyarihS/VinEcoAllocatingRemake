﻿// --------------------------------------------------------------------------------------------------------------------
// <copyright file="AppearanceViewModel.cs" company="VinEco">
//   Shirayuki 2018.
// </copyright>
// <summary>
//   A simple view model for configuring theme, font and accent colors.
// </summary>
// --------------------------------------------------------------------------------------------------------------------

namespace VinEcoAllocatingRemake.Pages.Settings
{
    #region

    using System.ComponentModel;
    using System.Linq;
    using System.Windows.Media;

    using FirstFloor.ModernUI.Presentation;

    #endregion

    /// <summary>
    ///     A simple view model for configuring theme, font and accent colors.
    /// </summary>
    public class AppearanceViewModel : NotifyPropertyChanged
    {
        /// <summary>
        ///     The font large.
        /// </summary>
        private const string FontLarge = "large";

        /// <summary>
        ///     The font small.
        /// </summary>
        private const string FontSmall = "small";

        // 9 accent colors from metro design principles
        /*private Color[] accentColors = new Color[]{
            Color.FromRgb(0x33, 0x99, 0xff),   // blue
            Color.FromRgb(0x00, 0xab, 0xa9),   // teal
            Color.FromRgb(0x33, 0x99, 0x33),   // green
            Color.FromRgb(0x8c, 0xbf, 0x26),   // lime
            Color.FromRgb(0xf0, 0x96, 0x09),   // orange
            Color.FromRgb(0xff, 0x45, 0x00),   // orange red
            Color.FromRgb(0xe5, 0x14, 0x00),   // red
            Color.FromRgb(0xff, 0x00, 0x97),   // magenta
            Color.FromRgb(0xa2, 0x00, 0xff),   // purple            
        };*/
        
        /// <summary>
        ///     20 accent colors from Windows Phone 8
        /// </summary>
        private Color selectedAccentColor;

        /// <summary>
        ///     The selected font size.
        /// </summary>
        private string selectedFontSize;

        /// <summary>
        ///     The selected theme.
        /// </summary>
        private Link selectedTheme;

        /// <summary>
        ///     Initializes a new instance of the <see cref="AppearanceViewModel"/> class.
        /// </summary>
        public AppearanceViewModel()
        {
            // add the default themes
            this.Themes.Add(new Link { DisplayName = "dark", Source = AppearanceManager.DarkThemeSource });
            this.Themes.Add(new Link { DisplayName = "light", Source = AppearanceManager.LightThemeSource });

            this.SelectedFontSize = AppearanceManager.Current.FontSize == FontSize.Large ? FontLarge : FontSmall;
            this.SyncThemeAndColor();

            AppearanceManager.Current.PropertyChanged += this.OnAppearanceManagerPropertyChanged;
        }

        /// <summary>
        ///     Gets the accent colors.
        /// </summary>
        public Color[] AccentColors
        {
            get; 
        } 
            =
                {
                    Color.FromRgb(0xa4, 0xc4, 0x00), // lime
                    Color.FromRgb(0x60, 0xa9, 0x17), // green
                    Color.FromRgb(0x00, 0x8a, 0x00), // emerald
                    Color.FromRgb(0x00, 0xab, 0xa9), // teal
                    Color.FromRgb(0x1b, 0xa1, 0xe2), // cyan
                    Color.FromRgb(0x00, 0x50, 0xef), // cobalt
                    Color.FromRgb(0x6a, 0x00, 0xff), // indigo
                    Color.FromRgb(0xaa, 0x00, 0xff), // violet
                    Color.FromRgb(0xf4, 0x72, 0xd0), // pink
                    Color.FromRgb(0xd8, 0x00, 0x73), // magenta
                    Color.FromRgb(0xa2, 0x00, 0x25), // crimson
                    Color.FromRgb(0xe5, 0x14, 0x00), // red
                    Color.FromRgb(0xfa, 0x68, 0x00), // orange
                    Color.FromRgb(0xf0, 0xa3, 0x0a), // amber
                    Color.FromRgb(0xe3, 0xc8, 0x00), // yellow
                    Color.FromRgb(0x82, 0x5a, 0x2c), // brown
                    Color.FromRgb(0x6d, 0x87, 0x64), // olive
                    Color.FromRgb(0x64, 0x76, 0x87), // steel
                    Color.FromRgb(0x76, 0x60, 0x8a), // mauve
                    Color.FromRgb(0x87, 0x79, 0x4e) // taupe
                };

        /// <summary>
        ///     The font sizes.
        /// </summary>
        public string[] FontSizes => new[] { FontSmall, FontLarge };

        /// <summary>
        ///     Gets or sets the selected accent color.
        /// </summary>
        public Color SelectedAccentColor
        {
            get => this.selectedAccentColor;
            set
            {
                if (this.selectedAccentColor == value)
                {
                    return;
                }

                this.selectedAccentColor = value;
                this.OnPropertyChanged("SelectedAccentColor");

                AppearanceManager.Current.AccentColor = value;
            }
        }

        /// <summary>
        ///     Gets or sets the selected font size.
        /// </summary>
        public string SelectedFontSize
        {
            get => this.selectedFontSize;
            set
            {
                if (this.selectedFontSize == value)
                {
                    return;
                }

                this.selectedFontSize = value;
                this.OnPropertyChanged("SelectedFontSize");

                AppearanceManager.Current.FontSize = value == FontLarge ? FontSize.Large : FontSize.Small;
            }
        }

        /// <summary>
        ///     Gets or sets the selected theme.
        /// </summary>
        public Link SelectedTheme
        {
            get => this.selectedTheme;
            set
            {
                if (this.selectedTheme == value)
                {
                    return;
                }

                this.selectedTheme = value;
                this.OnPropertyChanged("SelectedTheme");

                // and update the actual theme
                AppearanceManager.Current.ThemeSource = value.Source;
            }
        }

        /// <summary>
        ///     Gets the themes.
        /// </summary>
        public LinkCollection Themes { get; } = new LinkCollection();

        /// <summary>
        ///     The on appearance manager property changed.
        /// </summary>
        /// <param name="sender"> The sender. </param>
        /// <param name="e"> The e. </param>
        private void OnAppearanceManagerPropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName == "ThemeSource" || e.PropertyName == "AccentColor")
            {
                this.SyncThemeAndColor();
            }
        }

        /// <summary>
        ///     The sync theme and color.
        /// </summary>
        private void SyncThemeAndColor()
        {
            // synchronizes the selected viewmodel theme with the actual theme used by the appearance manager.
            this.SelectedTheme = this.Themes.FirstOrDefault(l => l.Source.Equals(AppearanceManager.Current.ThemeSource));

            // and make sure accent color is up-to-date
            this.SelectedAccentColor = AppearanceManager.Current.AccentColor;
        }
    }
}