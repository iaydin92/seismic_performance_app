using System;

namespace AESPerformansApp.Calculations
{
    public abstract class CalculationInputBase
    {
        public string SectionName { get; set; }
        public string Category { get; set; }  // NEW: Category bilgisi (örn: "Beam-I-Section")
        public double Fy { get; set; }
        public double L { get; set; }
    }
} 