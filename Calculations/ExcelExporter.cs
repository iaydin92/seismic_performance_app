using System;
using System.IO;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Drawing;

namespace AESPerformansApp.Calculations
{
    /// <summary>
    /// Handles exporting calculation results to Excel files
    /// </summary>
    public class ExcelExporter
    {
        private readonly string _sapFilePath;
        private readonly string _excelFilePath;

        public ExcelExporter(string sapFilePath)
        {
            _sapFilePath = sapFilePath;
            _excelFilePath = GetExcelFilePath(sapFilePath);
        }

        /// <summary>
        /// Gets the Excel file path based on SAP2000 model file path
        /// </summary>
        private string GetExcelFilePath(string sapFilePath)
        {
            string directory = Path.GetDirectoryName(sapFilePath);
            string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(sapFilePath);
            return Path.Combine(directory, fileNameWithoutExtension + "_Results.xlsx");
        }

        /// <summary>
        /// Exports beam calculation results to Excel
        /// </summary>
        public void ExportBeamResults(BeamCalculationInput beamData, 
                                     double profile_a, double profile_b, double profile_c,
                                     double profile_IO, double profile_LS, double profile_CP,
                                     double profile_a_prime, double profile_b_prime, double profile_c_prime,
                                     double profile_IO_prime, double profile_LS_prime, double profile_CP_prime,
                                     double control_1, double control_2, double control_3, double control_4,
                                     double adjustment_factor)
        {
            try
            {
                // Remove this line - license is now set globally in App.xaml.cs
                // ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                
                using (var package = File.Exists(_excelFilePath) ? new ExcelPackage(new FileInfo(_excelFilePath)) : new ExcelPackage())
                {
                    // Get or create worksheet
                    var worksheet = package.Workbook.Worksheets["Beam Results"] ?? package.Workbook.Worksheets.Add("Beam Results");

                    // Create headers if this is the first entry
                    if (worksheet.Dimension == null)
                    {
                        CreateHeaders(worksheet);
                    }

                    // Find the next empty row
                    int row = worksheet.Dimension?.End.Row + 1 ?? 2;

                    // Fill data
                    FillBeamData(worksheet, row, beamData, 
                                profile_a, profile_b, profile_c, profile_IO, profile_LS, profile_CP,
                                profile_a_prime, profile_b_prime, profile_c_prime, profile_IO_prime, profile_LS_prime, profile_CP_prime,
                                control_1, control_2, control_3, control_4, adjustment_factor);

                    // Auto-fit columns
                    worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();

                    // Save file
                    package.SaveAs(new FileInfo(_excelFilePath));

                    Console.WriteLine($"Successfully exported beam results to {_excelFilePath}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error exporting to Excel: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// Creates headers for the Excel table
        /// </summary>
        private void CreateHeaders(ExcelWorksheet worksheet)
        {
            var headers = new[]
            {
                // Basic Info
                "Beam Name", "Category", "L (mm)", "Fy (MPa)",
                
                // SAP2000 Properties
                "t3 (mm)", "t2 (mm)", "tf (mm)", "tw (mm)", "t2b (mm)", "tfb (mm)", "Radius (mm)",
                "I33 (mm4)", "Z33 (mm3)", "Material",
                
                // Column Info
                "Selected Column", "Column Length (mm)", "Column Fy (MPa)",
                
                // Original Calculated Values
                "profile_a", "profile_b", "profile_c", "profile_IO", "profile_LS", "profile_CP",
                
                // Control Factors
                "control_1", "control_2", "control_3", "control_4", "Adjustment Factor",
                
                // Final Values (After Control Factors)
                "profile_a_prime", "profile_b_prime", "profile_c_prime", 
                "profile_IO_prime", "profile_LS_prime", "profile_CP_prime",
                
                // Calculation Date
                "Calculation Date"
            };

            // Set headers
            for (int i = 0; i < headers.Length; i++)
            {
                var cell = worksheet.Cells[1, i + 1];
                cell.Value = headers[i];
                cell.Style.Font.Bold = true;
                cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                cell.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
                cell.Style.Border.BorderAround(ExcelBorderStyle.Thin);
            }
        }

        /// <summary>
        /// Fills beam data into the Excel row
        /// </summary>
        private void FillBeamData(ExcelWorksheet worksheet, int row, BeamCalculationInput beamData,
                                 double profile_a, double profile_b, double profile_c,
                                 double profile_IO, double profile_LS, double profile_CP,
                                 double profile_a_prime, double profile_b_prime, double profile_c_prime,
                                 double profile_IO_prime, double profile_LS_prime, double profile_CP_prime,
                                 double control_1, double control_2, double control_3, double control_4,
                                 double adjustment_factor)
        {
            int col = 1;

            // Basic Info
            worksheet.Cells[row, col++].Value = beamData.SectionName;
            worksheet.Cells[row, col++].Value = beamData.Category;
            worksheet.Cells[row, col++].Value = beamData.L;
            worksheet.Cells[row, col++].Value = beamData.Fy;

            // SAP2000 Properties
            worksheet.Cells[row, col++].Value = beamData.t3;
            worksheet.Cells[row, col++].Value = beamData.t2;
            worksheet.Cells[row, col++].Value = beamData.tf;
            worksheet.Cells[row, col++].Value = beamData.tw;
            worksheet.Cells[row, col++].Value = beamData.t2b;
            worksheet.Cells[row, col++].Value = beamData.tfb;
            worksheet.Cells[row, col++].Value = beamData.radius;
            worksheet.Cells[row, col++].Value = beamData.I33;
            worksheet.Cells[row, col++].Value = beamData.Z33;
            worksheet.Cells[row, col++].Value = beamData.MaterialProperty;

            // Column Info
            worksheet.Cells[row, col++].Value = beamData.SelectedColumnSection;
            worksheet.Cells[row, col++].Value = beamData.ConnectedColumnLength;
            worksheet.Cells[row, col++].Value = beamData.ColumnFy;

            // Original Calculated Values
            worksheet.Cells[row, col++].Value = Math.Round(profile_a, 3);
            worksheet.Cells[row, col++].Value = Math.Round(profile_b, 3);
            worksheet.Cells[row, col++].Value = Math.Round(profile_c, 3);
            worksheet.Cells[row, col++].Value = Math.Round(profile_IO, 3);
            worksheet.Cells[row, col++].Value = Math.Round(profile_LS, 3);
            worksheet.Cells[row, col++].Value = Math.Round(profile_CP, 3);

            // Control Factors
            worksheet.Cells[row, col++].Value = Math.Round(control_1, 3);
            worksheet.Cells[row, col++].Value = Math.Round(control_2, 3);
            worksheet.Cells[row, col++].Value = Math.Round(control_3, 3);
            worksheet.Cells[row, col++].Value = Math.Round(control_4, 3);
            worksheet.Cells[row, col++].Value = Math.Round(adjustment_factor, 3);

            // Final Values (After Control Factors)
            worksheet.Cells[row, col++].Value = Math.Round(profile_a_prime, 3);
            worksheet.Cells[row, col++].Value = Math.Round(profile_b_prime, 3);
            worksheet.Cells[row, col++].Value = Math.Round(profile_c_prime, 3);
            worksheet.Cells[row, col++].Value = Math.Round(profile_IO_prime, 3);
            worksheet.Cells[row, col++].Value = Math.Round(profile_LS_prime, 3);
            worksheet.Cells[row, col++].Value = Math.Round(profile_CP_prime, 3);

            // Calculation Date
            worksheet.Cells[row, col++].Value = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

            // Add borders to all cells in this row
            for (int i = 1; i < col; i++)
            {
                worksheet.Cells[row, i].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            }
        }
    }
} 