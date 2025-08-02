namespace AESPerformansApp.Calculations
{
    // Inherits common properties from CalculationInputBase
    public class ColumnCalculationInput : CalculationInputBase
    {
        // This class might have specific inputs or results for columns in the future.
        // For now, it mainly serves to distinguish column calculations.

        // Example of a column-specific result
        public double AxialCapacity_Result { get; set; }
         // User Inputs specific to Braces
        public double L2 { get; set; }
        public double L3 { get; set; }

        // Data to be fetched from SAP2000
        public double Area {get;set;}
        
        // Calculation Result
        public double t3 { get; set; }  // Total depth (d)
        public double t2 { get; set; }  // Flange width (b)
        public double tf { get; set; }  // Flange thickness (tf)
        public double tw { get; set; }  // Web thickness (tw)
        public double t2b { get; set; } // Bottom flange width (usually same as t2 for I-sections)
        public double tfb { get; set; } // Bottom flange thickness (usually same as tf for I-sections)
        public double radius { get; set; } // Fillet radius
        
        public double r2 {get;set;}
        public double r3 {get;set;}
        // Additional calculated properties
        public double WebHeight { get; set; } // h = t3 - 2*tf (clear web height)
        public string MaterialProperty { get; set; } // Material from SAP2000

        // Additional section properties
        public double I33 { get; set; } // Major-axis moment of inertia (mm^4)
        public double Z33 { get; set; } // Major-axis section modulus (mm^3)

    
        //reults 
        public double compression_a { get; set; }
        public double compression_b { get; set; }
        public double compression_c { get; set; }
        public double compression_IO { get; set; }
        public double compression_LS { get; set; }
        public double compression_CP { get; set; }
        
        public double tension_a { get; set; }
        public double tension_b { get; set; }
        public double tension_c { get; set; }
        public double tension_IO { get; set; }
        public double tension_LS { get; set; }
        public double tension_CP { get; set; }
        
    }
} 