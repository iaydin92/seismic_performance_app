using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using Microsoft.Win32;
using SAP2000v1;
using System.IO;
using System.Collections.ObjectModel;
using System.Linq;
using System.Collections.ObjectModel;
using AESPerformansApp.Calculations;

namespace AESPerformansApp
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {

        //Fields to hold sap2000 objects, 
        private cOAPI _sapObject;
        private cSapModel _sapModel;
        private string _currentSapFilePath; // Store the current SAP2000 file path

        private Dictionary<string, ObservableCollection<Variant>> _sectionVariants = new();        
        private readonly ObservableCollection<string> _frameSections = new ObservableCollection<string>();
        private readonly ObservableCollection<string> _allSections = new ObservableCollection<string>(); // NEW - stores all sections
        public MainWindow()
        {
            InitializeComponent();
            SectionsGrid.ItemsSource = _frameSections;   

        }

        private void BrowseButton_Click(object sender, RoutedEventArgs e) 
        {
            var dlg = new OpenFileDialog
            {
                Filter = "SAP2000 Model (*.sdb)|*.sdb|SAP2000 Text (*.s2k)|*.s2k|All files (*.*)|*.*",
                Title = "Select a SAP2000 Model"
            };
            if (dlg.ShowDialog(this) == true)
            {
                _currentSapFilePath = dlg.FileName; // Store the file path
                StatusText.Text = $"Selected: {System.IO.Path.GetFileName(dlg.FileName)}";
                OpenSapModel(dlg.FileName);
            }
        }

        private void OpenSapModel(string filePath)
        {

            try
            {
                // (1) Get or create SAP2000
                cHelper helper = new Helper();


                _sapObject = helper.CreateObjectProgID("CSI.SAP2000.API.SapObject");
                    if (_sapObject == null)
                    {
                        MessageBox.Show("SAP2000 API object could not be created. Is SAP2000 installed?",
                                        "API Error", MessageBoxButton.OK, MessageBoxImage.Error);
                        StatusText.Text = "SAP2000 API could not be started.";
                        return;
                    }

                    int retStart = _sapObject.ApplicationStart();
                    if (retStart != 0)
                    {
                        MessageBox.Show($"SAP2000 could not be started. Error code: {retStart}",
                                        "Startup Error", MessageBoxButton.OK, MessageBoxImage.Error);
                        StatusText.Text = "SAP2000 startup failed.";
                        ReleaseSapObjects();
                        return;
                    }


                _sapModel = _sapObject.SapModel;
                if (_sapModel == null)
                {
                    MessageBox.Show("Unable to get SapModel object.", "Model Error",
                                    MessageBoxButton.OK, MessageBoxImage.Error);
                    StatusText.Text = "SapModel object unavailable.";
                    ReleaseSapObjects();
                    return;
                }

                // (2) Open the model
                StatusText.Text = $"Opening '{Path.GetFileName(filePath)}'…";
                int retOpen = _sapModel.File.OpenFile(filePath);

                if (retOpen == 0)
                {
                    StatusText.Text = $"Connected: {Path.GetFileName(filePath)}";
                    MessageBox.Show($"Successfully opened '{Path.GetFileName(filePath)}'.",
                                    "Connected", MessageBoxButton.OK, MessageBoxImage.Information);

                    // (3) Optional: set working units
                    int retUnits = _sapModel.SetPresentUnits(eUnits.kN_mm_C);
                    Debug.WriteLine(retUnits == 0
                        ? "Units set to kN-mm-C."
                        : $"Unit set failed – code {retUnits}");
                    LoadFrameSections();
                }
                else
                {
                    MessageBox.Show($"Could not open model. Error code: {retOpen}",
                                    "Open Error", MessageBoxButton.OK, MessageBoxImage.Error);
                    StatusText.Text = "Model could not be opened.";
                    ReleaseSapObjects();
                }
            }
            catch (COMException comEx)
            {
                MessageBox.Show($"COM error while talking to SAP2000:\n{comEx.Message}",
                                "COM Error", MessageBoxButton.OK, MessageBoxImage.Error);
                StatusText.Text = "SAP2000 communication error.";
                ReleaseSapObjects();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Unexpected error:\n{ex.Message}",
                                "General Error", MessageBoxButton.OK, MessageBoxImage.Error);
                StatusText.Text = "An unexpected error occurred.";
                ReleaseSapObjects();
            }
        }

        private void ReleaseSapObjects()              // NEW
        {
            try
            {
                if (_sapObject != null)
                {
                    _sapObject.ApplicationExit(false);
                    Marshal.ReleaseComObject(_sapObject);
                }
            }
            catch { /* swallow cleanup exceptions */ }
            finally
            {
                _sapObject = null;
                _sapModel = null;
            }
        }


        //Load Frame Sections 
        private void LoadFrameSections()
        {
            _frameSections.Clear();
            _allSections.Clear(); // NEW - clear both collections

            if (_sapModel == null) return;

            int count = 0;
            string[] names = null;
            int ret = _sapModel.PropFrame.GetNameList(ref count, ref names);

            if (ret == 0 && names != null)
            {
                foreach (var n in names)
                {
                    _allSections.Add(n); // NEW - add to master list
                    _frameSections.Add(n); // add to display list
                }
                StatusText.Text += $"  |  {count} sections loaded.";
            }
            else
            {
                StatusText.Text += "  |  Could not retrieve section list.";
            }
        }

        // NEW ▾ search filter method
        private void SearchTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            string searchText = SearchTextBox.Text.ToLower();
            
            _frameSections.Clear();
            
            if (string.IsNullOrWhiteSpace(searchText))
            {
                // If search is empty, show all sections
                foreach (var section in _allSections)
                {
                    _frameSections.Add(section);
                }
            }
            else
            {
                // Filter sections that contain the search text
                var filteredSections = _allSections.Where(s => 
                    s.ToLower().Contains(searchText));
                
                foreach (var section in filteredSections)
                {
                    _frameSections.Add(section);
                }
            }
        }

        private void SectionsGrid_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (SectionsGrid.SelectedItem is string sectionName)
            {

                            // Diyelim ki sectionName değişkenin var:
            if (!_sectionVariants.ContainsKey(sectionName))
            {
                _sectionVariants[sectionName] = new ObservableCollection<Variant>();
            }
            var variants = _sectionVariants[sectionName];
            // SectionDetailsWindow'u açarken:
            var dlg = new SectionDetailsWindow(sectionName, _allSections, _sapModel, _currentSapFilePath, variants);
            dlg.Owner = this;
            dlg.ShowDialog();
            }
        }

    }
}