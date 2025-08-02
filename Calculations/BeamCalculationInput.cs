using System;

namespace AESPerformansApp.Calculations
{
    public class BeamCalculationInput : CalculationInputBase
    {
        // User inputs
        public string SelectedColumnSection { get; set; }
        public double ConnectedColumnLength { get; set; }
        public double ColumnFy { get; set; } // Column yield strength (user input)

        // SAP2000 I-Section geometric properties
        public double t3 { get; set; }  // Total depth (d)
        public double t2 { get; set; }  // Flange width (b)
        public double tf { get; set; }  // Flange thickness (tf)
        public double tw { get; set; }  // Web thickness (tw)
        public double t2b { get; set; } // Bottom flange width (usually same as t2 for I-sections)
        public double tfb { get; set; } // Bottom flange thickness (usually same as tf for I-sections)
        public double radius { get; set; } // Fillet radius
        
        // Additional calculated properties
        public double WebHeight { get; set; } // h = t3 - 2*tf (clear web height)
        public string MaterialProperty { get; set; } // Material from SAP2000

        // Additional section properties
        public double I33 { get; set; } // Major-axis moment of inertia (mm^4)
        public double Z33 { get; set; } // Major-axis section modulus (mm^3)

        // Calculation results - Profile values (before control factors)
        public double profile_a { get; set; }
        public double profile_b { get; set; }
        public double profile_c { get; set; }
        public double profile_IO { get; set; }
        public double profile_LS { get; set; }
        public double profile_CP { get; set; }
        
        // Final values (after applying control factors)
        public double profile_a_prime { get; set; }
        public double profile_b_prime { get; set; }
        public double profile_c_prime { get; set; }
        public double profile_IO_prime { get; set; }
        public double profile_LS_prime { get; set; }
        public double profile_CP_prime { get; set; }
        
        // Control factors
        public double control_1 { get; set; }
        public double control_2 { get; set; }
        public double control_3 { get; set; }
        public double control_4 { get; set; }
        public double adjustment_factor { get; set; }
    }
} 