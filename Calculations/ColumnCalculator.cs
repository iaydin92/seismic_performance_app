using SAP2000v1;
using System;

namespace AESPerformansApp.Calculations
{
    /// <summary>
    /// Calculator for all column-related section types.
    /// </summary>
    public class ColumnCalculator : Calculator
    {
        public void PerformCalculation(CalculationInputBase inputs, cSapModel sapModel, string sapFilePath)
        {
            if (inputs is not ColumnCalculationInput columnData)
            {
                return;
            }

            // --- Step 1: Fetch data from SAP2000 based on Category ---
            double sectionArea = FetchSectionDataFromSAP(columnData, sapModel);

            // --- Step 2: Perform the main calculation ---
            // Placeholder for axial capacity calculation using section area.
            columnData.AxialCapacity_Result = sectionArea * columnData.Fy / 1000.0; // Convert to kN

            // --- Step 3: Calculation is complete ---
            Console.WriteLine($"Calculation complete for {columnData.SectionName} ({columnData.Category}). Axial Capacity = {columnData.AxialCapacity_Result:F2} kN");
        }

        private double FetchSectionDataFromSAP(ColumnCalculationInput columnData, cSapModel sapModel)
        {
            try
            {
                Console.WriteLine($"Fetching data for column section: {columnData.SectionName} - Category: {columnData.Category}");

                // Initialize variables for SAP2000 API calls
                string fileName = "";
                string matProp = "";
                double t3 = 0, t2 = 0, tf = 0, tw = 0, t2b = 0, tfb = 0;
                int color = 0;
                string notes = "";
                string guid = "";

                int result = -1;
                double sectionArea = 0.0;

                // Category'ye göre farklı SAP2000 API çağrıları
                if (columnData.Category.Contains("I-Section"))
                {
                    result = sapModel.PropFrame.GetISection(columnData.SectionName, ref fileName, ref matProp, 
                        ref t3, ref t2, ref tf, ref tw, ref t2b, ref tfb, ref color, ref notes, ref guid);
                    // I-Section area approximation: A = 2*tf*t2 + (t3-2*tf)*tw
                    sectionArea = 2 * tf * t2 + (t3 - 2 * tf) * tw;
                    Console.WriteLine($"  → I-Section API called. Result: {result}, Area ≈ {sectionArea:F2} mm²");
                }
                else if (columnData.Category.Contains("Channel"))
                {
                    result = sapModel.PropFrame.GetChannel(columnData.SectionName, ref fileName, ref matProp, 
                        ref t3, ref t2, ref tf, ref tw, ref color, ref notes, ref guid);
                    // Channel area approximation: A = tf*t2 + (t3-tf)*tw
                    sectionArea = tf * t2 + (t3 - tf) * tw;
                    Console.WriteLine($"  → Channel API called. Result: {result}, Area ≈ {sectionArea:F2} mm²");
                }
                else if (columnData.Category.Contains("Box"))
                {
                    result = sapModel.PropFrame.GetTube(columnData.SectionName, ref fileName, ref matProp, 
                        ref t3, ref t2, ref tf, ref tw, ref color, ref notes, ref guid);
                    // Box area approximation: A = 2*tf*(t2-2*tw) + 2*tw*t3
                    sectionArea = 2 * tf * (t2 - 2 * tw) + 2 * tw * t3;
                    Console.WriteLine($"  → Box/Tube API called. Result: {result}, Area ≈ {sectionArea:F2} mm²");
                }
                else if (columnData.Category.Contains("Angle"))
                {
                    result = sapModel.PropFrame.GetAngle(columnData.SectionName, ref fileName, ref matProp, 
                        ref t3, ref t2, ref tf, ref tw, ref color, ref notes, ref guid);
                    // Angle area approximation: A = tf*t2 + tw*t3 - tf*tw
                    sectionArea = tf * t2 + tw * t3 - tf * tw;
                    Console.WriteLine($"  → Angle API called. Result: {result}, Area ≈ {sectionArea:F2} mm²");
                }
                else if (columnData.Category.Contains("Tube"))
                {
                    result = sapModel.PropFrame.GetTube(columnData.SectionName, ref fileName, ref matProp, 
                        ref t3, ref t2, ref tf, ref tw, ref color, ref notes, ref guid);
                    // Circular tube area: A = π*(D²-(D-2*t)²)/4 where D=t3, t=tw
                    double outerD = t3;
                    double innerD = outerD - 2 * tw;
                    sectionArea = Math.PI * (outerD * outerD - innerD * innerD) / 4.0;
                    Console.WriteLine($"  → Tube API called. Result: {result}, Area ≈ {sectionArea:F2} mm²");
                }
                else if (columnData.Category.Contains("UserDefined"))
                {
                    Console.WriteLine($"  → UserDefined section - using default area");
                    sectionArea = 5000.0; // Default area in mm²
                    result = 0;
                }

                if (result != 0)
                {
                    Console.WriteLine($"  → Warning: Could not fetch section data. Using default area.");
                    sectionArea = 5000.0; // Default area in mm²
                }

                return sectionArea;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"  → Error fetching section data: {ex.Message}");
                return 5000.0; // Default fallback area
            }
        }
    }
} 