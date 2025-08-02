using System;
using System.ComponentModel;

namespace AESPerformansApp.Calculations
{
    /// <summary>
    /// Represents a calculation variant with its name, category, and results
    /// </summary>
    public class Variant : INotifyPropertyChanged
    {
        private string _name = string.Empty;
        private string _category = string.Empty;
        private CalculationInputBase? _inputData;
        private DateTime _createdDate;

        public string Name
        {
            get => _name;
            set
            {
                _name = value;
                OnPropertyChanged(nameof(Name));
            }
        }

        public string Category
        {
            get => _category;
            set
            {
                _category = value;
                OnPropertyChanged(nameof(Category));
            }
        }

        public CalculationInputBase? InputData
        {
            get => _inputData;
            set
            {
                _inputData = value;
                OnPropertyChanged(nameof(InputData));
            }
        }

        public DateTime CreatedDate
        {
            get => _createdDate;
            set
            {
                _createdDate = value;
                OnPropertyChanged(nameof(CreatedDate));
            }
        }

        /// <summary>
        /// Gets a summary of the calculation results for display
        /// </summary>
        public string ResultSummary
        {
            get
            {
                if (InputData == null) return "Hesaplanmamış";

                if (InputData is BeamCalculationInput beamResult)
                {
                }
                else if (InputData is BraceCalculationInput braceResult)
                {
                }
                else if (InputData is ColumnCalculationInput columnResult)
                {
                    return $"Axial: {columnResult.AxialCapacity_Result:F2}";
                }

                return "Bilinmeyen sonuç";
            }
        }

        public Variant()
        {
            CreatedDate = DateTime.Now;
        }

        public Variant(string name, string category, CalculationInputBase inputData)
        {
            Name = name;
            Category = category;
            InputData = inputData;
            CreatedDate = DateTime.Now;
        }

        public event PropertyChangedEventHandler? PropertyChanged;

        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}