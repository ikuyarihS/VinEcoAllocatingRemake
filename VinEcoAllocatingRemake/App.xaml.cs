using System;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Windows;

namespace VinEcoAllocatingRemake
{
    /// <summary>
    ///     Interaction logic for App.xaml
    /// </summary>
    // ReSharper disable once InheritdocConsiderUsage
    public partial class App
    {
        private void Application_Startup(object sender, StartupEventArgs e)
        {
            // Create a Startup Window.
            var newWindow = new MainWindow {Title = "Phần mềm của KHSX VinEco"};

            // Show the window.
            newWindow.Show();
        }
    }

    public static class Program
    {
        [STAThread]
        public static void Main()
        {
            AppDomain.CurrentDomain.AssemblyResolve += OnResolveAssembly;
            App.Main();
        }

        private static Assembly OnResolveAssembly(object sender, ResolveEventArgs args)
        {
            Assembly executingAssembly = Assembly.GetExecutingAssembly();
            var assemblyName = new AssemblyName(args.Name);

            string path = $"{assemblyName.Name}.dll";
            if (assemblyName.CultureInfo.Equals(CultureInfo.InvariantCulture) == false)
                path = $@"{assemblyName.CultureInfo}\{path}";

            using (Stream stream = executingAssembly.GetManifestResourceStream(path))
            {
                if (stream == null) return null;

                var assemblyRawBytes = new byte[stream.Length];
                stream.Read(assemblyRawBytes, 0, assemblyRawBytes.Length);
                return Assembly.Load(assemblyRawBytes);
            }
        }
    }
}