#region

using System;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Windows;

#endregion

namespace VinEcoAllocatingRemake
{
    #region

    #endregion

    /// <summary>
    ///     The program.
    /// </summary>
    public static class Program
    {
        /// <summary>
        ///     The main.
        /// </summary>
        [STAThread]
        public static void Main()
        {
            AppDomain.CurrentDomain.AssemblyResolve += OnResolveAssembly;
            App.Main();
        }

        /// <summary>
        ///     The on resolve assembly.
        /// </summary>
        /// <param name="sender">
        ///     The sender.
        /// </param>
        /// <param name="args">
        ///     The args.
        /// </param>
        /// <returns>
        ///     The <see cref="Assembly" />.
        /// </returns>
        private static Assembly OnResolveAssembly(object sender, ResolveEventArgs args)
        {
            Assembly executingAssembly = Assembly.GetExecutingAssembly();
            var assemblyName = new AssemblyName(args.Name);

            string path = $"{assemblyName.Name}.dll";
            if (assemblyName.CultureInfo.Equals(CultureInfo.InvariantCulture) == false) path = $@"{assemblyName.CultureInfo}\{path}";

            using (Stream stream = executingAssembly.GetManifestResourceStream(path))
            {
                if (stream == null) return null;

                var assemblyRawBytes = new byte[stream.Length];
                stream.Read(assemblyRawBytes, 0, assemblyRawBytes.Length);
                return Assembly.Load(assemblyRawBytes);
            }
        }
    }

    /// <summary>
    ///     Interaction logic for App.xaml
    /// </summary>
    [SuppressMessage("StyleCop.CSharp.MaintainabilityRules", "SA1402:FileMayOnlyContainASingleClass", Justification = "Reviewed. Suppression is OK here.")]
    // ReSharper disable once InheritdocConsiderUsage
    public partial class App
    {
        /// <summary>
        ///     The application startup.
        /// </summary>
        /// <param name="sender">
        ///     The sender.
        /// </param>
        /// <param name="e">
        ///     The e.
        /// </param>
        private void ApplicationStartup(object sender, StartupEventArgs e)
        {
            // Create a Startup Window.
            var newWindow = new MainWindow {Title = "Phần mềm của KHSX VinEco"};

            // Show the window.
            newWindow.Show();
        }
    }
}