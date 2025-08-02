using AESPerformansApp.Calculations;
using SAP2000v1;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace AESPerformansApp
{
    /// <summary>
    /// Interaction logic for SectionDetailsWindow.xaml
    /// </summary>
    public partial class SectionDetailsWindow : Window
    {
        private readonly string _sectionName;
        private readonly cSapModel _sapModel;
        private readonly string _sapFilePath; // Store SAP2000 file path for $2k writer
        private readonly ObservableCollection<string> _allColumnSections = new();
        private readonly ObservableCollection<Variant> _variants;
        
        // Store references to input controls
        private TextBox _txtL, _txtL2, _txtL3, _txtFy, _txtColumnLength, _txtColumnFy;
        private ComboBox _columnSectionComboBox;

        public SectionDetailsWindow(
            string sectionName,
            ObservableCollection<string> allSections,
            cSapModel sapModel,
            string sapFilePath,
            ObservableCollection<Variant> variants)
        {
            InitializeComponent();
            _sectionName = sectionName;
            _sapModel = sapModel;
            _sapFilePath = sapFilePath;
            _variants = variants;
            
            // Set the section name label
            SectionLabel.Text = $"Section: {sectionName}";
            
            // Manually populate the ComboBox with all specific categories
            CategoryComboBox.Items.Add("Beam-I-Section");
            CategoryComboBox.Items.Add("Beam-Channel-Section");
            CategoryComboBox.Items.Add("Beam-Box-Section");
            CategoryComboBox.Items.Add("Beam-Angle-Section");
            CategoryComboBox.Items.Add("Beam-Tube-Section");
            CategoryComboBox.Items.Add("Beam-UserDefined-Section");
            
            CategoryComboBox.Items.Add("Brace-I-Section");
            CategoryComboBox.Items.Add("Brace-Channel-Section");
            CategoryComboBox.Items.Add("Brace-Box-Section");
            CategoryComboBox.Items.Add("Brace-Angle-Section");
            CategoryComboBox.Items.Add("Brace-Tube-Section");
            CategoryComboBox.Items.Add("Brace-UserDefined-Section");
            
            CategoryComboBox.Items.Add("Column-I-Section");
            CategoryComboBox.Items.Add("Column-Channel-Section");
            CategoryComboBox.Items.Add("Column-Box-Section");
            CategoryComboBox.Items.Add("Column-Angle-Section");
            CategoryComboBox.Items.Add("Column-Tube-Section");
            CategoryComboBox.Items.Add("Column-UserDefined-Section");
            
            CategoryComboBox.SelectedIndex = 0; // Select first item
            
            // Load column sections from main window
            LoadColumnSections(allSections);
            
            // Create initial inputs based on default selection
            CreateInputsForCategory("Beam-I-Section");
            
            // Initialize DataGrid binding
            VariantsDataGrid.ItemsSource = _variants;
        }

        private void LoadColumnSections(ObservableCollection<string> allSections)
        {
            // Copy sections from main window
            foreach (var section in allSections)
            {
                _allColumnSections.Add(section);
            }
        }

        private void CategoryComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (CategoryComboBox.SelectedItem is string selectedCategory)
            {
                CreateInputsForCategory(selectedCategory);
            }
        }

        private void CreateInputsForCategory(string category)
        {
            // Clear existing inputs
            DynamicInputsPanel.Children.Clear();
            
            if (category.StartsWith("Brace"))
            {
                CreateBraceInputs();
            }
            else if (category.StartsWith("Beam"))
            {
                CreateBeamInputs();
            }
            else if (category.StartsWith("Column"))
            {
                CreateColumnInputs();
            }
            else
            {
                // CreateOtherInputs();
            }
        }

        private void CreateBraceInputs()
        {
            // L input
            var lblL = new TextBlock { Text = "L:", FontWeight = FontWeights.Bold, Margin = new Thickness(0, 10, 0, 5) };
            _txtL = new TextBox { Margin = new Thickness(0, 0, 0, 10) };
            
            // L1 input
            var lblL2 = new TextBlock { Text = "L2:", FontWeight = FontWeights.Bold, Margin = new Thickness(0, 10, 0, 5) };
            _txtL2 = new TextBox { Margin = new Thickness(0, 0, 0, 10) };
            
            // L2 input
            var lblL3 = new TextBlock { Text = "L3:", FontWeight = FontWeights.Bold, Margin = new Thickness(0, 10, 0, 5) };
            _txtL3 = new TextBox { Margin = new Thickness(0, 0, 0, 10) };
            
            // fy input
            var lblFy = new TextBlock { Text = "fy:", FontWeight = FontWeights.Bold, Margin = new Thickness(0, 10, 0, 5) };
            _txtFy = new TextBox { Margin = new Thickness(0, 0, 0, 10) };
            
            DynamicInputsPanel.Children.Add(lblL);
            DynamicInputsPanel.Children.Add(_txtL);
            DynamicInputsPanel.Children.Add(lblL2);
            DynamicInputsPanel.Children.Add(_txtL2);
            DynamicInputsPanel.Children.Add(lblL3);
            DynamicInputsPanel.Children.Add(_txtL3);
            DynamicInputsPanel.Children.Add(lblFy);
            DynamicInputsPanel.Children.Add(_txtFy);
        }

        private void CreateBeamInputs()
        {
            // L input
            var lblL = new TextBlock { Text = "L:", FontWeight = FontWeights.Bold, Margin = new Thickness(0, 10, 0, 5) };
            _txtL = new TextBox { Margin = new Thickness(0, 0, 0, 10) };
            
            // Connected Column Section (ComboBox only)
            var lblColumnSection = new TextBlock { Text = "Bağlı Kolon Kesiti:", FontWeight = FontWeights.Bold, Margin = new Thickness(0, 10, 0, 5) };
            
            // ComboBox for column sections
            _columnSectionComboBox = new ComboBox 
            { 
                Margin = new Thickness(0, 0, 0, 10),
                ItemsSource = _allColumnSections
            };
            
            // Connected Column Length
            var lblColumnLength = new TextBlock { Text = "Bağlı Kolon Uzunluğu L:", FontWeight = FontWeights.Bold, Margin = new Thickness(0, 10, 0, 5) };
            _txtColumnLength = new TextBox { Margin = new Thickness(0, 0, 0, 10) };

            // fy input
            var lblFy = new TextBlock { Text = "fy:", FontWeight = FontWeights.Bold, Margin = new Thickness(0, 10, 0, 5) };
            _txtFy = new TextBox { Margin = new Thickness(0, 0, 0, 10) };
            
            // Column fy input
            var lblColumnFy = new TextBlock { Text = "Kolon fy:", FontWeight = FontWeights.Bold, Margin = new Thickness(0, 10, 0, 5) };
            _txtColumnFy = new TextBox { Margin = new Thickness(0, 0, 0, 10) };
            
            DynamicInputsPanel.Children.Add(lblL);
            DynamicInputsPanel.Children.Add(_txtL);
            DynamicInputsPanel.Children.Add(lblColumnSection);
            DynamicInputsPanel.Children.Add(_columnSectionComboBox);
            DynamicInputsPanel.Children.Add(lblColumnLength);
            DynamicInputsPanel.Children.Add(_txtColumnLength);
            DynamicInputsPanel.Children.Add(lblFy);
            DynamicInputsPanel.Children.Add(_txtFy);
            DynamicInputsPanel.Children.Add(lblColumnFy);
            DynamicInputsPanel.Children.Add(_txtColumnFy);
        }

        private void CreateColumnInputs()
        {
            // L input
            var lblL = new TextBlock { Text = "L:", FontWeight = FontWeights.Bold, Margin = new Thickness(0, 10, 0, 5) };
            _txtL = new TextBox { Margin = new Thickness(0, 0, 0, 10) };
            
            // fy input
            var lblFy = new TextBlock { Text = "fy:", FontWeight = FontWeights.Bold, Margin = new Thickness(0, 10, 0, 5) };
            _txtFy = new TextBox { Margin = new Thickness(0, 0, 0, 10) };
            
            DynamicInputsPanel.Children.Add(lblL);
            DynamicInputsPanel.Children.Add(_txtL);
            DynamicInputsPanel.Children.Add(lblFy);
            DynamicInputsPanel.Children.Add(_txtFy);
        }

        private void Ok_Click(object sender, RoutedEventArgs e)
        {
            // Get variant name
            string variantName = VariantNameTextBox.Text.Trim();
            if (string.IsNullOrEmpty(variantName))
            {
                MessageBox.Show("Lütfen varyant adı girin.", "Uyarı", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            string selectedCategory = CategoryComboBox.SelectedItem as string;
            if (string.IsNullOrEmpty(selectedCategory)) return;

            // 1. Create the appropriate input data model (NO CALCULATION)
            CalculationInputBase inputData = CreateInputData(selectedCategory);
            if (inputData == null) return;

            // 2. Check if variant already exists
            var existingVariant = _variants.FirstOrDefault(v => v.Name == variantName);
            if (existingVariant != null)
            {
                // Update existing variant's input data
                existingVariant.Category = selectedCategory;
                existingVariant.InputData = inputData;
            }
            else
            {
                // Add new variant
                var newVariant = new Variant(variantName, selectedCategory, inputData);
                _variants.Add(newVariant);
            }

            // 3. Clear variant name textbox for next entry
            VariantNameTextBox.Clear();
            VariantNameTextBox.Focus();
        }

        private CalculationInputBase CreateInputData(string category)
        {
            try
            {
                if (category.StartsWith("Beam"))
                {
                    return new BeamCalculationInput
                    {
                        SectionName = _sectionName,
                        Category = category,  // NEW: Kategori bilgisini ekle
                        L = double.Parse(_txtL.Text),
                        Fy = double.Parse(_txtFy.Text),
                        SelectedColumnSection = _columnSectionComboBox.SelectedItem as string,
                        ConnectedColumnLength = double.Parse(_txtColumnLength.Text),
                        ColumnFy = double.Parse(_txtColumnFy.Text)
                    };
                }
                else if (category.StartsWith("Brace"))
                {
                    return new BraceCalculationInput
                    {
                        SectionName = _sectionName,
                        Category = category,  // NEW: Kategori bilgisini ekle
                        L = double.Parse(_txtL.Text),
                        L2 = double.Parse(_txtL2.Text),
                        L3 = double.Parse(_txtL3.Text),
                        Fy = double.Parse(_txtFy.Text)
                    };
                }
                else if (category.StartsWith("Column"))
                {
                    return new ColumnCalculationInput
                    {
                        SectionName = _sectionName,
                        Category = category,  // NEW: Kategori bilgisini ekle
                        L = double.Parse(_txtL.Text),
                        Fy = double.Parse(_txtFy.Text)
                    };
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error parsing input values: {ex.Message}", "Input Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return null;
            }
            return null;
        }

        private Calculator CreateCalculator(string category)
        {
            if (category.StartsWith("Beam"))
            {
                return new BeamCalculator();
            }
            else if (category.StartsWith("Brace"))
            {
                return new BraceCalculator();
            }
            else if (category.StartsWith("Column"))
            {
                return new ColumnCalculator();
            }
            return null;
        }

        private void ShowResults(CalculationInputBase resultData)
        {
            StringBuilder resultText = new StringBuilder();
            resultText.AppendLine($"Results for section: {_sectionName}");

            if (resultData is BeamCalculationInput beamResult)
            {
              
            }
            else if (resultData is BraceCalculationInput braceResult)
            {
            }
            else if (resultData is ColumnCalculationInput columnResult)
            {
                resultText.AppendLine($"Axial Capacity: {columnResult.AxialCapacity_Result:F2}");
            }

        }


        private void Cancel_Click(object sender, RoutedEventArgs e) =>
            DialogResult = false;

        // YENİ METOT
        private void HesaplaVeYaz_Click(object sender, RoutedEventArgs e)
        {
            if (_variants.Count == 0)
            {
                MessageBox.Show("Hesaplanacak veya yazılacak varyant bulunmuyor.", "Uyarı", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            try
            {
                // Her bir varyant için hesaplama yap ve yaz
                foreach (var variant in _variants)
                {
                    if (variant.InputData == null) continue;

                    // 1. Doğru hesaplayıcıyı oluştur
                    Calculator calculator = CreateCalculator(variant.Category);
                    if (calculator == null) continue;

                    // 2. Hesaplamaları yap (sonuçlar variant.InputData içine yazılır)
                    calculator.PerformCalculation(variant.InputData, _sapModel, _sapFilePath);

                    
                    // HingeData oluştur (hesaplama sonuçlarından)
                    var hingeData = CreateHingeDataFromCalculation(variant.InputData);
                    if (hingeData == null) continue; // Bu satırı da düzelt!
                    switch (variant.Category)
                    {
                        case "Beam-I-Section":
                        case "Beam-Channel-Section":
                        case "Beam-Box-Section":
                        case "Beam-Angle-Section":
                        case "Beam-Tube-Section":
                        case "Beam-UserDefined-Section":
                            var beamWriter = new Sap2kBeamWriter(_sapFilePath);
                            beamWriter.WriteBeamHingeDefinitions(variant.Name, (HingeData)hingeData);
                            break;

                        case "Brace-I-Section":
                        case "Brace-Channel-Section":
                        case "Brace-Box-Section":
                        case "Brace-Angle-Section":
                        case "Brace-Tube-Section":
                        case "Brace-UserDefined-Section":
                            var braceWriter = new Sap2kBraceWriter(_sapFilePath);
                            braceWriter.WriteBraceHingeDefinitions(variant.Name, (BraceHingeData)hingeData);
                            break;

                        case "Column-I-Section":
                        case "Column-Channel-Section":
                        case "Column-Box-Section":
                        case "Column-Angle-Section":
                        case "Column-Tube-Section":
                        case "Column-UserDefined-Section":
                            var columnWriter = new Sap2kColumnWriter(_sapFilePath);
                            columnWriter.WriteColumnHingeDefinitions(variant.Name, (ColumnHingeData)hingeData);
                            break;

                        default:
                            // Gerekirse hata veya uyarı verin
                            break;
                    }
                }

                MessageBox.Show($"{_variants.Count} adet varyant başarıyla hesaplandı ve dosyaya yazıldı.", "İşlem Tamamlandı", MessageBoxButton.OK, MessageBoxImage.Information);

                // Pencereyi kapat
                DialogResult = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Hesaplama veya yazma sırasında bir hata oluştu: {ex.Message}", "Hata", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void EditVariant_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button button && button.Tag is Variant variant)
            {
                // Load variant data into the form
                VariantNameTextBox.Text = variant.Name;
                
                // Set category
                CategoryComboBox.SelectedItem = variant.Category;
                
                // Load input values if available
                if (variant.InputData != null)
                {
                    LoadVariantInputs(variant.InputData);
                }
                
                MessageBox.Show($"Varyant '{variant.Name}' düzenleme için yüklendi. Değişikliklerinizi yapın ve 'Varyant Ekle/Güncelle' butonuna basın.", 
                    "Bilgi", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        private void DeleteVariant_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button button && button.Tag is Variant variant)
            {
                var result = MessageBox.Show($"'{variant.Name}' varyantını silmek istediğinizden emin misiniz?", 
                    "Silme Onayı", MessageBoxButton.YesNo, MessageBoxImage.Question);
                
                if (result == MessageBoxResult.Yes)
                {
                    _variants.Remove(variant);
                    MessageBox.Show($"Varyant '{variant.Name}' silindi.", "Bilgi", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
        }

        private void LoadVariantInputs(CalculationInputBase inputData)
        {
            try
            {
                if (inputData is BeamCalculationInput beamInput)
                {
                    _txtL.Text = beamInput.L.ToString();
                    _txtFy.Text = beamInput.Fy.ToString();
                    _txtColumnLength.Text = beamInput.ConnectedColumnLength.ToString();
                    _txtColumnFy.Text = beamInput.ColumnFy.ToString();
                    
                    if (_columnSectionComboBox != null && !string.IsNullOrEmpty(beamInput.SelectedColumnSection))
                    {
                        _columnSectionComboBox.SelectedItem = beamInput.SelectedColumnSection;
                    }
                }
                else if (inputData is BraceCalculationInput braceInput)
                {
                    _txtL.Text = braceInput.L.ToString();
                    _txtL2.Text = braceInput.L2.ToString();
                    _txtL3.Text = braceInput.L3.ToString();
                    _txtFy.Text = braceInput.Fy.ToString();
                }
                else if (inputData is ColumnCalculationInput columnInput)
                {
                    _txtL.Text = columnInput.L.ToString();
                    _txtFy.Text = columnInput.Fy.ToString();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Varyant verileri yüklenirken hata oluştu: {ex.Message}", "Hata", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// Hesaplama sonuçlarından HingeData objesi oluşturur
        /// </summary>
        private object? CreateHingeDataFromCalculation(CalculationInputBase inputData)
        {
            try
            {
                if (inputData is BeamCalculationInput beamInput)
                {
                    // BeamCalculator'dan gelen GERÇEK hesaplama sonuçlarını kullan
                    return new HingeData
                    {
                        // Original values (before control factors) - BeamCalculator'dan gelen gerçek değerler
                        profile_a = beamInput.profile_a,
                        profile_b = beamInput.profile_b,
                        profile_c = beamInput.profile_c,
                        profile_IO = beamInput.profile_IO,
                        profile_LS = beamInput.profile_LS,
                        profile_CP = beamInput.profile_CP,
                        
                        // Final values (after control factors) - BeamCalculator'dan gelen gerçek değerler
                        profile_a_prime = beamInput.profile_a_prime,
                        profile_b_prime = beamInput.profile_b_prime,
                        profile_c_prime = beamInput.profile_c_prime,
                        profile_IO_prime = beamInput.profile_IO_prime,
                        profile_LS_prime = beamInput.profile_LS_prime,
                        profile_CP_prime = beamInput.profile_CP_prime
                    };
                }
                else if (inputData is BraceCalculationInput braceInput)
                {
                    // BraceCalculator sonuçları için HingeData (geçici değerler - BraceCalculator'da henüz implement edilmemiş)
                    return new BraceHingeData
                    {
                        // Original values (before control factors) - BeamCalculator'dan gelen gerçek değerler
                        compression_a = braceInput.compression_a,
                        compression_b = braceInput.compression_b,
                        compression_c = braceInput.compression_c,
                        compression_IO = braceInput.compression_IO,
                        compression_LS = braceInput.compression_LS,
                        compression_CP = braceInput.compression_CP,

                        // Final values (after control factors) - BeamCalculator'dan gelen gerçek değerler

                        tension_a = braceInput.tension_a,
                        tension_b = braceInput.tension_b,
                        tension_c = braceInput.tension_c,
                        tension_IO = braceInput.tension_IO,
                        tension_LS = braceInput.tension_LS,
                        tension_CP = braceInput.tension_CP
                    };
                }
                else if (inputData is ColumnCalculationInput columnInput)
                {
                    return new ColumnHingeData
                    {
                         // Original values (before control factors) - BeamCalculator'dan gelen gerçek değerler
                        compression_a = columnInput.compression_a,
                        compression_b = columnInput.compression_b,
                        compression_c = columnInput.compression_c,
                        compression_IO = columnInput.compression_IO,
                        compression_LS = columnInput.compression_LS,
                        compression_CP = columnInput.compression_CP,

                        // Final values (after control factors) - BeamCalculator'dan gelen gerçek değerler

                        tension_a = columnInput.tension_a,
                        tension_b = columnInput.tension_b,
                        tension_c = columnInput.tension_c,
                        tension_IO = columnInput.tension_IO,
                        tension_LS = columnInput.tension_LS,
                        tension_CP = columnInput.tension_CP
                    };
                }
                
                return null;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"HingeData oluşturulurken hata: {ex.Message}", "Hata", MessageBoxButton.OK, MessageBoxImage.Error);
                return null;
            }
        }
    }
}

 