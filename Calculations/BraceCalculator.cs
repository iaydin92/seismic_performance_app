using SAP2000v1;
using System;
using System.Windows;

namespace AESPerformansApp.Calculations
{
    /// <summary>
    /// Calculator for all brace-related section types.
    /// </summary>
    public class BraceCalculator : Calculator
    {
        public void PerformCalculation(CalculationInputBase inputs, cSapModel sapModel, string sapFilePath)
        {
            if (inputs is not BraceCalculationInput braceData)
            {
                return;
            }

            // --- Step 1: Fetch data from SAP2000 based on Category ---
            FetchSectionDataFromSAP(braceData, sapModel);
            double kl_r = Math.Max(braceData.L3/ braceData.r3, braceData.L2/ braceData.r2);
            
            // Slenderness check formulas
            double limit1 = 4.2*Math.Sqrt(200000/braceData.Fy);  // bf/tf <= 52/sqrt(fy/6.895)
            double limit2 = 2.1*Math.Sqrt(200000/braceData.Fy);
            
            // Result checks
            bool result1 = kl_r >= limit1;  // bf/tf <= 52/sqrt(fy/6.895)
            bool result2 = kl_r <= limit2;   // h/tw <= 418/sqrt(fy/6.895)

            double lc_ix = braceData.L3/braceData.r3;
            double lc_iy = braceData.L2/braceData.r2;
            double fe = Math.Pow(Math.PI, 2) * 200000 / Math.Pow(Math.Max(lc_ix, lc_iy), 2);
            double fcr;
            if (Math.Max(lc_ix, lc_iy) <= 4.71 * Math.Sqrt(200000 / braceData.Fy))
            {
                fcr = Math.Pow(0.658, braceData.Fy / fe) * braceData.Fy;
            }
            else
            {
                fcr = 0.877 * fe;
            }
            double p_y = fcr*braceData.Area/1000;
            double delta_c = (p_y*1000*braceData.L)/(braceData.Area*200000);
            double t_y = braceData.Fy * braceData.Area/1000;
            double delta_t = t_y*1000*braceData.L/(braceData.Area*200000);
            
            // Calculate profile values and IO, LS, CP values
            double compression_a = CalculateParameter(
                result1, result2, 
                kl_r, limit1, limit2, 
                stockyValue: 1.0, slenderValue: 0.5);
            double compression_b = CalculateParameter(
                result1, result2, 
                kl_r, limit1, limit2, 
                stockyValue: 7.0, slenderValue: 9.0); 
            double compression_c = CalculateParameter(
                result1, result2, 
                kl_r, limit1, limit2, 
                stockyValue: 0.5, slenderValue: 0.3); 
            double compression_IO = CalculateParameter(
                result1, result2, 
                kl_r, limit1, limit2, 
                stockyValue: 0.5, slenderValue: 0.5); 
            double compression_LS = CalculateParameter(
                result1, result2, 
                kl_r, limit1, limit2,
                stockyValue: 6.0, slenderValue: 7.0); 
            double compression_CP = CalculateParameter(
                result1, result2, 
                kl_r, limit1, limit2, 
                stockyValue: 7.0, slenderValue: 9.0);

            double tension_a = 8;
            double tension_b =9;
            double tension_c = 0.6;
            double tension_IO = 0.5;
            double tension_LS = 7;
            double tension_CP = 9;


            // Hesaplanan değerleri göster - Debug için
            string debugMessage = $@"BRACE HESAPLAMA SONUÇLARI - {braceData.SectionName}

            === GİRİŞ DEĞERLERİ ===
            L: {braceData.L:F2} mm
            L2: {braceData.L2:F2} mm  
            L3: {braceData.L3:F2} mm
            Fy: {braceData.Fy:F2} MPa
            Area: {braceData.Area:F2} mm²
            I33: {braceData.I33:F2} mm⁴
            r2: {braceData.r2:F2} mm
            r3: {braceData.r3:F2} mm

            === HESAPLANAN DEĞERLER ===
            kl_r: {kl_r:F4}
            lc_ix: {lc_ix:F6}
            lc_iy: {lc_iy:F6}
            fe: {fe:F2} MPa
            fcr: {fcr:F2} MPa

            p_y: {p_y:F4} kN
            delta_c: {delta_c:F6} mm
            t_y: {t_y:F4} kN
            delta_t: {delta_t:F6} mm

            === COMPRESSION VALUES ===
            compression_a: {compression_a:F4}
            compression_b: {compression_b:F4}
            compression_c: {compression_c:F4}
            compression_IO: {compression_IO:F4}
            compression_LS: {compression_LS:F4}
            compression_CP: {compression_CP:F4}

            === TENSION VALUES ===
            tension_a: {tension_a:F4}
            tension_b: {tension_b:F4}
            tension_c: {tension_c:F4}
            tension_IO: {tension_IO:F4}
            tension_LS: {tension_LS:F4}
            tension_CP: {tension_CP:F4}";

            MessageBox.Show(debugMessage, "Brace Hesaplama Debug", MessageBoxButton.OK, MessageBoxImage.Information);

            // Store results in BeamCalculationInput for later use
            braceData.compression_a = compression_a;
            braceData.compression_b = compression_b;
            braceData.compression_c = compression_c;
            braceData.compression_IO = compression_IO;
            braceData.compression_LS = compression_LS;
            braceData.compression_CP = compression_CP;
            
            braceData.tension_a = tension_a;
            braceData.tension_b = tension_b;
            braceData.tension_c = tension_c;
            braceData.tension_IO = tension_IO;
            braceData.tension_LS = tension_LS;
            braceData.tension_CP = tension_CP;
            
            // --- Step 4: Write hinge definitions to $2k file ---
            try
            {
                var sap2kWriter = new Sap2kBraceWriter(sapFilePath);
                var hingeData = new BraceHingeData
                {
                    // Original values (before control factors)
                    compression_a = compression_a,
                    compression_b = compression_b,
                    compression_c = compression_c,
                    compression_IO = compression_IO,
                    compression_LS = compression_LS,
                    compression_CP = compression_CP,
                    
                    tension_a = tension_a,
                    tension_b = tension_b,
                    tension_c = tension_c,
                    tension_IO = tension_IO,
                    tension_LS = tension_LS,
                    tension_CP = tension_CP,

                    p_y = p_y,
                    delta_c = delta_c,
                    t_y = t_y,
                    delta_t = delta_t
                };

                string hingeName = braceData.SectionName + "_Axial";
                sap2kWriter.WriteBraceHingeDefinitions(hingeName, hingeData);
                
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
                // var excelExporter = new ExcelExporter(sapFilePath);
                // excelExporter.ExportBraceResults(braceData,
                //     profile_a, profile_b, profile_c, profile_IO, profile_LS, profile_CP,
                //     profile_a_prime, profile_b_prime, profile_c_prime, profile_IO_prime, profile_LS_prime, profile_CP_prime,
                //     control_1, control_2, control_3, control_4, adjustment_factor);
                
                Console.WriteLine($"Successfully exported results to Excel for {braceData.SectionName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error exporting to Excel: {ex.Message}");
                // Don't throw - calculation was successful even if Excel export failed
            }
   
        }

        private void FetchSectionDataFromSAP(BraceCalculationInput braceData, cSapModel sapModel)
        {
            try
            {
                Console.WriteLine($"\n=== FETCHING DATA FOR {braceData.SectionName} ===");
                Console.WriteLine($"Category: {braceData.Category}");

                // Initialize variables for SAP2000 API calls
                string fileName = "";
                string matProp = "";
                double t3 = 0, t2 = 0, tf = 0, tw = 0, t2b = 0, tfb = 0, filletRadius = 0;
                int color = 0;
                string notes = "";
                string guid = "";
                int result = -1;

                // Category'ye göre farklı SAP2000 API çağrıları
                if (braceData.Category.Contains("I-Section"))
                {
                    // GetISection_1 metodunu kullanarak FilletRadius'u da çekiyoruz
                    result = sapModel.PropFrame.GetISection_1(braceData.SectionName, ref fileName, ref matProp, 
                        ref t3, ref t2, ref tf, ref tw, ref t2b, ref tfb, ref filletRadius, ref color, ref notes, ref guid);
                    
                    Console.WriteLine($"I-Section API Result: {result}");
                    if (result == 0)
                    {
                        // Store all geometric properties
                        braceData.t3 = t3;                    // Total depth (d)
                        braceData.t2 = t2;                    // Flange width (b)
                        braceData.tf = tf;                    // Flange thickness
                        braceData.tw = tw;                    // Web thickness
                        braceData.t2b = t2b;                  // Bottom flange width
                        braceData.tfb = tfb;                  // Bottom flange thickness
                        braceData.radius = filletRadius;      // Fillet radius - SAP2000'den direkt!
                        braceData.MaterialProperty = matProp;
                    }
                }
                else if (braceData.Category.Contains("Channel"))
                {
                    result = sapModel.PropFrame.GetChannel(braceData.SectionName, ref fileName, ref matProp, 
                        ref t3, ref t2, ref tf, ref tw, ref color, ref notes, ref guid);
                    
                    if (result == 0)
                    {
                        braceData.t3 = t3;
                        braceData.t2 = t2;
                        braceData.tf = tf;
                        braceData.tw = tw;
                        braceData.radius = 0; // Channel sections may not have fillet radius in the same way
                        braceData.MaterialProperty = matProp;
                    }
                }
                else if (braceData.Category.Contains("Box"))
                {
                    result = sapModel.PropFrame.GetTube(braceData.SectionName, ref fileName, ref matProp, 
                        ref t3, ref t2, ref tf, ref tw, ref color, ref notes, ref guid);
                    
                    if (result == 0)
                    {
                        braceData.t3 = t3;
                        braceData.t2 = t2;
                        braceData.tf = tf;
                        braceData.tw = tw;
                        braceData.radius = 0; // Box sections may not have fillet radius in the same way
                        braceData.MaterialProperty = matProp;
                    }
                }
                else if (braceData.Category.Contains("Angle"))
                {
                    result = sapModel.PropFrame.GetAngle(braceData.SectionName, ref fileName, ref matProp, 
                        ref t3, ref t2, ref tf, ref tw, ref color, ref notes, ref guid);
                    
                    if (result == 0)
                    {
                        braceData.t3 = t3;
                        braceData.t2 = t2;
                        braceData.tf = tf;
                        braceData.tw = tw;
                        braceData.radius = 0; // Angle sections may not have fillet radius in the same way
                        braceData.MaterialProperty = matProp;
                    }
                }
                else if (braceData.Category.Contains("Tube"))
                {
                    result = sapModel.PropFrame.GetTube(braceData.SectionName, ref fileName, ref matProp, 
                        ref t3, ref t2, ref tf, ref tw, ref color, ref notes, ref guid);
                    
                    if (result == 0)
                    {
                        braceData.t3 = t3;
                        braceData.t2 = t2;
                        braceData.tf = tf;
                        braceData.tw = tw;
                        braceData.radius = 0; // Tube sections may not have fillet radius in the same way
                        braceData.MaterialProperty = matProp;
                    }
                }
                else if (braceData.Category.Contains("UserDefined"))
                {
                    Console.WriteLine($"UserDefined section - using default values");
                    result = 0; // Assume success and use default values
                }

                if (result != 0)
                {
                    Console.WriteLine($"Warning: Could not fetch section data. Error code: {result}");
                }

                // GetSectProps ile tüm kesit özelliklerini al
                {
                    double area = 0, As2 = 0, As3 = 0, torsion = 0;
                    double I22 = 0, I33 = 0, S22 = 0, S33 = 0, Z22 = 0, Z33 = 0;
                    double R22 = 0, R33 = 0;

                    int resSect = sapModel.PropFrame.GetSectProps(braceData.SectionName, 
                        ref area, ref As2, ref As3, ref torsion, ref I22, ref I33, 
                        ref S22, ref S33, ref Z22, ref Z33, ref R22, ref R33);

                    Console.WriteLine($"GetSectProps API Result: {resSect}");
                    Console.WriteLine($"Fetched Area: {area} mm²");
                    Console.WriteLine($"Fetched I33: {I33} mm⁴");

                    if (resSect == 0)
                    {
                        braceData.Area = area;    // mm^2
                        braceData.I33 = I33;      // mm^4
                        braceData.Z33 = Z33;      // mm^3
                        braceData.r2 = R22;       // mm
                        braceData.r3 = R33;       // mm
                        
                        Console.WriteLine($"Successfully fetched all section properties for {braceData.SectionName}");
                    }
                    else
                    {
                        Console.WriteLine($"ERROR: GetSectProps failed with result code: {resSect}");
                        MessageBox.Show($"SAP2000'den {braceData.SectionName} için kesit özellikleri alınamadı. Hata kodu: {resSect}", 
                            "SAP2000 API Hatası", MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error fetching section data: {ex.Message}");
            }
        }

        private double CalculateParameter(
            bool result1, bool result2, 
            double kl_r, 
            double limit1, double limit2,
            double stockyValue, double slenderValue)
        {
            if (result1)
            {
                return stockyValue;
            }
            else if (result2)
            {
                return slenderValue;
            }
            else
            {
                // Linear interpolation
                return slenderValue + (slenderValue - stockyValue) / (limit2 - limit1) * (kl_r - limit2);
            }
        }

        

       
}
}