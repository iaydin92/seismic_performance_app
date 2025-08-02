using SAP2000v1;
using System;
using System.Runtime.InteropServices;
using System.Windows; // MessageBox için ekliyoruz
using System.Text;

namespace AESPerformansApp.Calculations
{
    /// <summary>
    /// Calculator for all beam-related section types.
    /// </summary>
    public class BeamCalculator : Calculator
    {
        public void PerformCalculation(CalculationInputBase inputs, cSapModel sapModel, string sapFilePath)
        {
            if (inputs is not BeamCalculationInput beamData)
            {
                // Invalid input type for this calculator
                return;
            }
            // --- Step 1: Fetch data from SAP2000 based on Category ---
            FetchSectionDataFromSAP(beamData, sapModel);

            // --- Step 2: Calculate additional geometric properties ---
            // --- Step 3: Perform the main calculation ---
            // Use user inputs (beamData.L, beamData.Fy) and SAP2000 data.
            // Moment and shear calculations intentionally left blank; will be implemented with final design formulas.
            // Example calculations using actual geometric properties
            double bf_tf = (beamData.t2/2) / beamData.tf;  // bf/tf ratio uses half flange width
            double h_tw = (beamData.t3 - 2*beamData.tf-2*beamData.radius)/beamData.tw;  // h/tw ratio
            
            // Slenderness check formulas
            double limit1 = 52.0 / Math.Sqrt(beamData.Fy / 6.895);  // bf/tf <= 52/sqrt(fy/6.895)
            double limit2 = 418.0 / Math.Sqrt(beamData.Fy / 6.895); // h/tw <= 418/sqrt(fy/6.895)
            double limit3 = 65.0 / Math.Sqrt(beamData.Fy / 6.895);  // bf/tf >= 65/sqrt(fy/6.895)
            double limit4 = 640.0 / Math.Sqrt(beamData.Fy / 6.895); // h/tw >= 640/sqrt(fy/6.895)
            
            // Result checks
            bool result1 = bf_tf <= limit1;  // bf/tf <= 52/sqrt(fy/6.895)
            bool result2 = h_tw <= limit2;   // h/tw <= 418/sqrt(fy/6.895)
            bool result3 = bf_tf >= limit3;  // bf/tf >= 65/sqrt(fy/6.895)
            bool result4 = h_tw >= limit4;   // h/tw >= 640/sqrt(fy/6.895)
            double theta_y = beamData.Z33*beamData.Fy*beamData.L/(6*200000*beamData.I33);
            double M_y = (beamData.Z33*beamData.Fy)/1000000;

            // Calculate profile values and IO, LS, CP values
            double profile_a = CalculateProfileParameter(
                result1, result2, result3, result4,
                bf_tf, h_tw, limit1, limit2, limit3, limit4,
                ductileValue: 9.0, notDuctileValue: 4.0);
            double profile_b = CalculateProfileParameter(
                result1, result2, result3, result4,
                bf_tf, h_tw, limit1, limit2, limit3, limit4,
                ductileValue: 11.0, notDuctileValue: 6.0); // L6 = 11.0, M6 = 6.0
            double profile_c = CalculateProfileParameter(
                result1, result2, result3, result4,
                bf_tf, h_tw, limit1, limit2, limit3, limit4,
                ductileValue: 0.6, notDuctileValue: 0.2); 
            double profile_IO = CalculateProfileParameter(
                result1, result2, result3, result4,
                bf_tf, h_tw, limit1, limit2, limit3, limit4,
                ductileValue: 1.0, notDuctileValue: 0.25); 
            double profile_LS = CalculateProfileParameter(
                result1, result2, result3, result4,
                bf_tf, h_tw, limit1, limit2, limit3, limit4,
                ductileValue: 9.0, notDuctileValue: 3.0); 
            double profile_CP = CalculateProfileParameter(
                result1, result2, result3, result4,
                bf_tf, h_tw, limit1, limit2, limit3, limit4,
                ductileValue: 11.0, notDuctileValue: 4.0);

            // Calculate all control factors using the new Excel formulas
            var columnData = GetColumnProperties(beamData.SelectedColumnSection, sapModel);
            
            var (control_1, control_2, control_3, control_4) = CalculateAllControls(
                columnData.tf,                      // C22 - column tcf
                beamData.t2,                        // C7 - beam bbf  
                0.0,                                // C27
                beamData.tf,                        // C8 - beam tbf
                beamData.ColumnFy,                  // C19 - column_fy
                columnData.t3,                      // C20 - column_dc
                0.0,                                // C15 - beam_n
                beamData.t3,                        // C6 - beam_db
                M_y,                                // F10 - M_y
                beamData.L,                         // C14 - beam_length
                beamData.ConnectedColumnLength,     // C26 - column_length
                columnData.tf,                      // C24 - column_tbw
                // Additional parameters for control_4
                result1, result2, result3, result4, // I3, I4, I6, I7
                bf_tf, h_tw,                        // F3, F4
                limit1, limit2, limit3, limit4);    // H3, H4, H6, H7
            double adjustment_factor = control_1*control_2*control_3*control_4;
            
            double profile_a_prime = profile_a*adjustment_factor;
            double profile_b_prime = profile_b*adjustment_factor;
            double profile_c_prime = profile_c*adjustment_factor;
            double profile_IO_prime = profile_IO*adjustment_factor;
            double profile_LS_prime = profile_LS*adjustment_factor;
            double profile_CP_prime = profile_CP*adjustment_factor;

            // Store results in BeamCalculationInput for later use
            beamData.profile_a = profile_a;
            beamData.profile_b = profile_b;
            beamData.profile_c = profile_c;
            beamData.profile_IO = profile_IO;
            beamData.profile_LS = profile_LS;
            beamData.profile_CP = profile_CP;
            
            beamData.profile_a_prime = profile_a_prime;
            beamData.profile_b_prime = profile_b_prime;
            beamData.profile_c_prime = profile_c_prime;
            beamData.profile_IO_prime = profile_IO_prime;
            beamData.profile_LS_prime = profile_LS_prime;
            beamData.profile_CP_prime = profile_CP_prime;
            
            beamData.control_1 = control_1;
            beamData.control_2 = control_2;
            beamData.control_3 = control_3;
            beamData.control_4 = control_4;
            beamData.adjustment_factor = adjustment_factor;

            Console.WriteLine($"Control results: control_1={control_1:F2}, control_2={control_2:F2}, control_3={control_3:F2}, control_4={control_4:F2}");

            // --- Step 4: Write hinge definitions to $2k file ---
            try
            {
                var sap2kWriter = new Sap2kBeamWriter(sapFilePath);
                var hingeData = new HingeData
                {
                    // Original values (before control factors)
                    profile_a = profile_a,
                    profile_b = profile_b,
                    profile_c = profile_c,
                    profile_IO = profile_IO,
                    profile_LS = profile_LS,
                    profile_CP = profile_CP,
                    
                    // Final values (after control factors) - used in $2k file
                    profile_a_prime = profile_a_prime,
                    profile_b_prime = profile_b_prime,
                    profile_c_prime = profile_c_prime,
                    profile_IO_prime = profile_IO_prime,
                    profile_LS_prime = profile_LS_prime,
                    profile_CP_prime = profile_CP_prime
                };

                string hingeName = beamData.SectionName + "_M3";
                sap2kWriter.WriteBeamHingeDefinitions(hingeName, hingeData);
                
                Console.WriteLine($"Successfully wrote hinge definitions for {hingeName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error writing hinge definitions: {ex.Message}");
                // Don't throw - calculation was successful even if hinge writing failed
            }

            // --- Step 5: Export results to Excel ---
            try
            {
                var excelExporter = new ExcelExporter(sapFilePath);
                excelExporter.ExportBeamResults(beamData,
                    profile_a, profile_b, profile_c, profile_IO, profile_LS, profile_CP,
                    profile_a_prime, profile_b_prime, profile_c_prime, profile_IO_prime, profile_LS_prime, profile_CP_prime,
                    control_1, control_2, control_3, control_4, adjustment_factor);
                
                Console.WriteLine($"Successfully exported results to Excel for {beamData.SectionName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error exporting to Excel: {ex.Message}");
                // Don't throw - calculation was successful even if Excel export failed
            }
   
        }

        private void FetchSectionDataFromSAP(BeamCalculationInput beamData, cSapModel sapModel)
        {
            try
            {
                Console.WriteLine($"\n=== FETCHING DATA FOR {beamData.SectionName} ===");
                Console.WriteLine($"Category: {beamData.Category}");

                // Initialize variables for SAP2000 API calls
                string fileName = "";
                string matProp = "";
                double t3 = 0, t2 = 0, tf = 0, tw = 0, t2b = 0, tfb = 0, filletRadius = 0;
                int color = 0;
                string notes = "";
                string guid = "";
                int result = -1;

                // Category'ye göre farklı SAP2000 API çağrıları
                if (beamData.Category.Contains("I-Section"))
                {
                    // GetISection_1 metodunu kullanarak FilletRadius'u da çekiyoruz
                    result = sapModel.PropFrame.GetISection_1(beamData.SectionName, ref fileName, ref matProp, 
                        ref t3, ref t2, ref tf, ref tw, ref t2b, ref tfb, ref filletRadius, ref color, ref notes, ref guid);
                    
                    Console.WriteLine($"I-Section API Result: {result}");
                    if (result == 0)
                    {
                        // Store all geometric properties
                        beamData.t3 = t3;                    // Total depth (d)
                        beamData.t2 = t2;                    // Flange width (b)
                        beamData.tf = tf;                    // Flange thickness
                        beamData.tw = tw;                    // Web thickness
                        beamData.t2b = t2b;                  // Bottom flange width
                        beamData.tfb = tfb;                  // Bottom flange thickness
                        beamData.radius = filletRadius;      // Fillet radius - SAP2000'den direkt!
                        beamData.MaterialProperty = matProp;
                    }
                }
                else if (beamData.Category.Contains("Channel"))
                {
                    result = sapModel.PropFrame.GetChannel(beamData.SectionName, ref fileName, ref matProp, 
                        ref t3, ref t2, ref tf, ref tw, ref color, ref notes, ref guid);
                    
                    if (result == 0)
                    {
                        beamData.t3 = t3;
                        beamData.t2 = t2;
                        beamData.tf = tf;
                        beamData.tw = tw;
                        beamData.radius = 0; // Channel sections may not have fillet radius in the same way
                        beamData.MaterialProperty = matProp;
                    }
                }
                else if (beamData.Category.Contains("Box"))
                {
                    result = sapModel.PropFrame.GetTube(beamData.SectionName, ref fileName, ref matProp, 
                        ref t3, ref t2, ref tf, ref tw, ref color, ref notes, ref guid);
                    
                    if (result == 0)
                    {
                        beamData.t3 = t3;
                        beamData.t2 = t2;
                        beamData.tf = tf;
                        beamData.tw = tw;
                        beamData.radius = 0; // Box sections may not have fillet radius in the same way
                        beamData.MaterialProperty = matProp;
                    }
                }
                else if (beamData.Category.Contains("Angle"))
                {
                    result = sapModel.PropFrame.GetAngle(beamData.SectionName, ref fileName, ref matProp, 
                        ref t3, ref t2, ref tf, ref tw, ref color, ref notes, ref guid);
                    
                    if (result == 0)
                    {
                        beamData.t3 = t3;
                        beamData.t2 = t2;
                        beamData.tf = tf;
                        beamData.tw = tw;
                        beamData.radius = 0; // Angle sections may not have fillet radius in the same way
                        beamData.MaterialProperty = matProp;
                    }
                }
                else if (beamData.Category.Contains("Tube"))
                {
                    result = sapModel.PropFrame.GetTube(beamData.SectionName, ref fileName, ref matProp, 
                        ref t3, ref t2, ref tf, ref tw, ref color, ref notes, ref guid);
                    
                    if (result == 0)
                    {
                        beamData.t3 = t3;
                        beamData.t2 = t2;
                        beamData.tf = tf;
                        beamData.tw = tw;
                        beamData.radius = 0; // Tube sections may not have fillet radius in the same way
                        beamData.MaterialProperty = matProp;
                    }
                }
                else if (beamData.Category.Contains("UserDefined"))
                {
                    Console.WriteLine($"UserDefined section - using default values");
                    result = 0; // Assume success and use default values
                }

                if (result != 0)
                {
                    Console.WriteLine($"Warning: Could not fetch section data. Error code: {result}");
                }

                // After geometry is fetched, get inertia and modulus values
                {
                    double area = 0, As2 = 0, As3 = 0, torsion = 0;
                    double I22 = 0, I33 = 0, I23 = 0, S22 = 0, S33 = 0, Z22 = 0, Z33 = 0;
                    double R22 = 0, R33 = 0, EccV2 = 0, EccV3 = 0;

                    int resGen = sapModel.PropFrame.GetGeneral_1(beamData.SectionName, ref fileName, ref matProp,
                        ref t3, ref t2, ref area, ref As2, ref As3, ref torsion, ref I22, ref I33, ref I23,
                        ref S22, ref S33, ref Z22, ref Z33, ref R22, ref R33, ref EccV2, ref EccV3,
                        ref color, ref notes, ref guid);

                    if (resGen == 0)
                    {
                        beamData.I33 = I33; // mm^4
                        beamData.Z33 = Z33; // mm^3
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error fetching section data: {ex.Message}");
            }
        }

        private double CalculateProfileParameter(
            bool result1, bool result2, bool result3, bool result4,
            double bf_tf, double h_tw, 
            double limit1, double limit2, double limit3, double limit4,
            double ductileValue, double notDuctileValue)
        {
            // Excel formula: IF(AND(I3="OK",I4="OK"),L5,IF(OR(I6="OK",I7="OK"),M5,MIN(L5+(M5-L5)/(H6-H3)*(F3-H3),L5+(M5-L5)/(H7-H4)*(F4-H4))))
            if (result1 && result2) // AND condition - both result1 and result2 are OK
            {
                return ductileValue;
            }
            else if (result3 || result4) // OR condition - either result3 or result4 is OK
            {
                return notDuctileValue;
            }
            else
            {
                // MIN calculation with linear interpolation
                double term1 = ductileValue + (notDuctileValue - ductileValue) / (limit3 - limit1) * (bf_tf - limit1);
                double term2 = ductileValue + (notDuctileValue - ductileValue) / (limit4 - limit2) * (h_tw - limit2);
                return Math.Min(term1, term2);
            }
        }

        private (double control_1, double control_2, double control_3, double control_4) CalculateAllControls(
            double C22, double C7, double C27, double C8,           // Original control_1 parameters
            double C19, double C20, double C15, double C6,          // Additional parameters
            double F10, double C14, double C26, double C24,         // More parameters for control_2
            bool I3, bool I4, bool I6, bool I7,                    // Result flags for control_4
            double F3, double F4,                                   // bf_tf, h_tw for control_4
            double H3, double H4, double H6, double H7)             // Limits for control_4
        {
            // === CONTROL_1 ===
            // Excel formula: =IF(C22>=C7/5.2,1,IF(AND(C22>=C7/7,C27<=C7/5.2,C27>=C8/2),1,IF(AND(C22<C7/7,C27>=C8),1,0.8)))
            double control_1;
            if (C22 >= C7 / 5.2)
            {
                control_1 = 1.0;
            }
            else if (C22 >= C7 / 7.0 && C27 <= C7 / 5.2 && C27 >= C8 / 2.0)
            {
                control_1 = 1.0;
            }
            else if (C22 < C7 / 7.0 && C27 >= C8)
            {
                control_1 = 1.0;
            }
            else
            {
                control_1 = 0.8;
            }

            // === CONTROL_2 ===
            // Excel formula: =IF(AND(((C15*F10*1000000/C6*(C14/(C14-C20))*((C26-C6)/C26)/1000)/(0.55*C19*C20*C24/1000))>=0.6,((C15*F10*1000000/C6*(C14/(C14-C20))*((C26-C6)/C26)/1000)/(0.55*C19*C20*C24/1000))<=0.9),1,0.8)
            
            double numerator = (C15 * F10 * 1000000 / C6 * (C14 / (C14 - C20)) * ((C26 - C6) / C26) / 1000);
            double denominator = (0.55 * C19 * C20 * C24 / 1000);
            double ratio = numerator / denominator;
            
            double control_2;
            if (ratio >= 0.6 && ratio <= 0.9)
            {
                control_2 = 1.0;
            }
            else
            {
                control_2 = 0.8;
            }

            // === CONTROL_3 ===
            // Excel formula: =IF((C14-IF(C15=1,C20,2*C20))/C6>=8,1,0.5^((8-(C14-IF(C15=1,C20,2*C20))/C6)/3))
            
            double adjustedLength = C14 - (C15 == 1 ? C20 : 2 * C20);
            double lengthRatio = adjustedLength / C6;
            
            double control_3;
            if (lengthRatio >= 8)
            {
                control_3 = 1.0;
            }
            else
            {
                control_3 = Math.Pow(0.5, (8 - lengthRatio) / 3);
            }

            // === CONTROL_4 ===
            // Excel formula: =MIN(IF(I3="OK",1,IF(I6="OK",0.5,((0.5-1)/(H6-H3)*(F3-H3)+1))),IF(I4="OK",1,IF(I7="OK",0.5,((0.5-1)/(H7-H4)*(F4-H4)+1))))
            
            // First part: IF(I3="OK",1,IF(I6="OK",0.5,((0.5-1)/(H6-H3)*(F3-H3)+1)))
            double part1;
            if (I3) // I3="OK"
            {
                part1 = 1.0;
            }
            else if (I6) // I6="OK"
            {
                part1 = 0.5;
            }
            else
            {
                part1 = ((0.5 - 1) / (H6 - H3) * (F3 - H3) + 1);
            }
            
            // Second part: IF(I4="OK",1,IF(I7="OK",0.5,((0.5-1)/(H7-H4)*(F4-H4)+1)))
            double part2;
            if (I4) // I4="OK"
            {
                part2 = 1.0;
            }
            else if (I7) // I7="OK"
            {
                part2 = 0.5;
            }
            else
            {
                part2 = ((0.5 - 1) / (H7 - H4) * (F4 - H4) + 1);
            }
            
            double control_4 = Math.Min(part1, part2);

            return (control_1, control_2, control_3, control_4);
        }

        private BeamCalculationInput GetColumnProperties(string columnSectionName, cSapModel sapModel)
        {
            var columnData = new BeamCalculationInput { SectionName = columnSectionName };

            if (string.IsNullOrEmpty(columnSectionName))
            {
                Console.WriteLine("No column section selected, using default values");
                return columnData;
            }

            try
            {
                Console.WriteLine($"\n=== FETCHING COLUMN DATA FOR {columnSectionName} ===");

                // Initialize variables for SAP2000 API calls
                string fileName = "";
                string matProp = "";
                double t3 = 0, t2 = 0, tf = 0, tw = 0, t2b = 0, tfb = 0, filletRadius = 0;
                int color = 0;
                string notes = "";
                string guid = "";
                int result = -1;



                result = sapModel.PropFrame.GetISection_1(columnSectionName, ref fileName, ref matProp, 
                    ref t3, ref t2, ref tf, ref tw, ref t2b, ref tfb, ref filletRadius, ref color, ref notes, ref guid);
                
                if (result == 0)
                {
                    Console.WriteLine($"Column I-Section API Result: {result}");
                    columnData.t3 = t3;                    // Total depth
                    columnData.t2 = t2;                    // Flange width
                    columnData.tf = tf;                    // Flange thickness
                    columnData.tw = tw;                    // Web thickness
                    columnData.t2b = t2b;                  // Bottom flange width
                    columnData.tfb = tfb;                  // Bottom flange thickness
                    columnData.radius = filletRadius;      // Fillet radius
                    columnData.MaterialProperty = matProp;
                    return columnData;
                }

                // Try Channel section
                result = sapModel.PropFrame.GetChannel(columnSectionName, ref fileName, ref matProp, 
                    ref t3, ref t2, ref tf, ref tw, ref color, ref notes, ref guid);
                
                if (result == 0)
                {
                    columnData.t3 = t3;
                    columnData.t2 = t2;
                    columnData.tf = tf;
                    columnData.tw = tw;
                    columnData.radius = 0;
                    columnData.MaterialProperty = matProp;
                    return columnData;
                }

                // Try Box/Tube section
                result = sapModel.PropFrame.GetTube(columnSectionName, ref fileName, ref matProp, 
                    ref t3, ref t2, ref tf, ref tw, ref color, ref notes, ref guid);
                
                if (result == 0)
                {
                    columnData.t3 = t3;
                    columnData.t2 = t2;
                    columnData.tf = tf;
                    columnData.tw = tw;
                    columnData.radius = 0;
                    columnData.MaterialProperty = matProp;
                    return columnData;
                }

                // Try Angle section
                result = sapModel.PropFrame.GetAngle(columnSectionName, ref fileName, ref matProp, 
                    ref t3, ref t2, ref tf, ref tw, ref color, ref notes, ref guid);
                
                if (result == 0)
                {
                    columnData.t3 = t3;
                    columnData.t2 = t2;
                    columnData.tf = tf;
                    columnData.tw = tw;
                    columnData.radius = 0;
                    columnData.MaterialProperty = matProp;
                    return columnData;
                }

                Console.WriteLine($"Warning: Could not fetch column section properties for {columnSectionName}");
                return columnData; // Return empty data
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error fetching column section data: {ex.Message}");
                return columnData; // Return empty data
            }
        }
    }
} 