using System.Configuration;
using System.Data;
using System.Windows;
using SAP2000v1;
using System;
using System.Runtime.InteropServices;
using System.Text;
using OfficeOpenXml; // Add this

namespace AESPerformansApp
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {
        protected override void OnStartup(StartupEventArgs e)
        {
            try
            {
                // Set EPPlus license for non-commercial use (EPPlus 8.0+ method)
                ExcelPackage.License.SetNonCommercialPersonal("AES Performance App User");
                
                // Console window'u aç (sadece debug için)
                #if DEBUG
                AllocConsole();
                #endif
                
                base.OnStartup(e);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Startup Error: {ex.Message}\n\nStack Trace:\n{ex.StackTrace}", 
                               "Application Startup Error", 
                               MessageBoxButton.OK, 
                               MessageBoxImage.Error);
                Environment.Exit(1);
            }
        }

        [System.Runtime.InteropServices.DllImport("kernel32.dll")]
        private static extern bool AllocConsole();
    }
}
