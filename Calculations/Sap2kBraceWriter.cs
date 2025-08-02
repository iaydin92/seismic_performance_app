using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;

namespace AESPerformansApp.Calculations
{
    /// <summary>
    /// Handles writing hinge definitions to SAP2000 $2k text files
    /// </summary>
    public class Sap2kBraceWriter 
    {
        private readonly string _sapFilePath;
        private readonly string _s2kFilePath;

        public Sap2kBraceWriter (string sapFilePath)
        {
            _sapFilePath = sapFilePath;
            _s2kFilePath = GetS2kFilePath(sapFilePath);
        }

        /// <summary>
        /// Gets the corresponding .$2k file path from the SAP2000 model file path
        /// </summary>
        private string GetS2kFilePath(string sapFilePath)
        {
            string directory = Path.GetDirectoryName(sapFilePath);
            string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(sapFilePath);
            return Path.Combine(directory, fileNameWithoutExtension + ".$2k");
        }

        /// <summary>
        /// Writes hinge definitions to the $2k file
        /// </summary>
        public void WriteBraceHingeDefinitions (string hingeName, BraceHingeData hingeData)
        {
            try
            {
                if (!File.Exists(_s2kFilePath))
                {
                    throw new FileNotFoundException($"$2K file not found: {_s2kFilePath}");
                }

                // Read existing file content
                var lines = File.ReadAllLines(_s2kFilePath).ToList();

                // Remove existing hinge definitions for this hinge name
                RemoveExistingHingeDefinitions(lines, hingeName);

                // Add new hinge definitions
                AddHingeDefinitions(lines, hingeName, hingeData);

                // Write back to file
                File.WriteAllLines(_s2kFilePath, lines);

                Console.WriteLine($"Successfully wrote hinge definitions for {hingeName} to {_s2kFilePath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error writing hinge definitions: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// Removes existing hinge definitions for the specified hinge name
        /// </summary>
        private void RemoveExistingHingeDefinitions(List<string> lines, string hingeName)
        {
            // Remove lines containing this hinge name from hinge tables
            for (int i = lines.Count - 1; i >= 0; i--)
            {
                if (lines[i].Contains($"HingeName={hingeName}"))
                {
                    lines.RemoveAt(i);
                }
            }
        }

        /// <summary>
        /// Adds new hinge definitions to the file
        /// </summary>
        private void AddHingeDefinitions(List<string> lines, string hingeName, BraceHingeData hingeData)
        {
            // Find or create table sections
            AddHingeDef02Table(lines, hingeName, hingeData);
            AddHingeDef03Table(lines, hingeName, hingeData);
            AddHingeDef04Table(lines, hingeName, hingeData);
        }

        /// <summary>
        /// Adds HINGES DEF 02 table entry
        /// </summary>
        private void AddHingeDef02Table(List<string> lines, string hingeName, BraceHingeData hingeData)
        {
            string tableHeader = "TABLE:  \"HINGES DEF 02 - NONINTERACTING - DEFORM CONTROL - GENERAL\"";
            int tableIndex = FindOrCreateTable(lines, tableHeader);


            




            string hingeDefLine = $"   HingeName={hingeName}   DOFType=\"Axial P\"   Symmetric=No   " +
                              $"BeyondE=Extrapolated   FDType=Force-Displ   UseYldForce=No   UseYldDispl=No   " +
                              $"FDPosForSF = {hingeData.t_y:F2} FDPosDisSF = {hingeData.delta_t:F2} FDNegForSF = {hingeData.p_y:F2} FDNegDisSF = {hingeData.delta_c:F2} " +
                              $"LengthType=Absolute   SSAbsLen=1   HysType=Kinematic";

            lines.Insert(tableIndex + 1, hingeDefLine);
        }

        /// <summary>
        /// Adds HINGES DEF 03 table entries (Force-Deform points)
        /// </summary>
        private void AddHingeDef03Table(List<string> lines, string hingeName, BraceHingeData hingeData)
        {
            string tableHeader = "TABLE:  \"HINGES DEF 03 - NONINTERACTING - DEFORM CONTROL - FORCE-DEFORM\"";
            int tableIndex = FindOrCreateTable(lines, tableHeader);

            var forceDeformLines = new List<string>
            {
                $"   HingeName={hingeName}   FDPoint=-E   Force={-hingeData.compression_c:F2}   Displ={-hingeData.compression_b:F4}",
                $"   HingeName={hingeName}   FDPoint=-D   Force={-hingeData.compression_c:F2}   Displ={(-hingeData.compression_a + (-1))-(-hingeData.compression_c):F4}",
                $"   HingeName={hingeName}   FDPoint=-C   Force={-1:F2}   Displ={-hingeData.compression_a:F4}",
                $"   HingeName={hingeName}   FDPoint=-B   Force={-1:F2}   Displ={0:F1}",
                $"   HingeName={hingeName}   FDPoint=A   Force={0:F1}   Displ={0:F1}",
                $"   HingeName={hingeName}   FDPoint=B   Force={1:F2}   Displ={0:F1}",
                $"   HingeName={hingeName}   FDPoint=C   Force={(1+0.03*hingeData.tension_a):F2}   Displ={hingeData.tension_a:F4}",
                $"   HingeName={hingeName}   FDPoint=D   Force={hingeData.tension_c:F2}   Displ={(1+0.03*hingeData.tension_a) - hingeData.tension_c + hingeData.tension_a:F4}",
                $"   HingeName={hingeName}   FDPoint=E   Force={hingeData.tension_c:F2}   Displ={hingeData.tension_b:F4}"
            };

            // Insert all force-deform lines
            for (int i = 0; i < forceDeformLines.Count; i++)
            {
                lines.Insert(tableIndex + 1 + i, forceDeformLines[i]);
            }
        }

        /// <summary>
        /// Adds HINGES DEF 04 table entries (Acceptance criteria)
        /// </summary>
        private void AddHingeDef04Table(List<string> lines, string hingeName, BraceHingeData hingeData)
        {
            string tableHeader = "TABLE:  \"HINGES DEF 04 - NONINTERACTING - DEFORM CONTROL - ACCEPTANCE\"";
            int tableIndex = FindOrCreateTable(lines, tableHeader);

            var acceptanceLines = new List<string>
            {
            $"   HingeName={hingeName}   ACPoint=IO   ACPos={hingeData.tension_IO:F2}   ACNeg={-hingeData.compression_IO:F2}",
            $"   HingeName={hingeName}   ACPoint=LS   ACPos={hingeData.tension_LS:F2}   ACNeg={-hingeData.compression_LS:F2}",
            $"   HingeName={hingeName}   ACPoint=CP   ACPos={hingeData.tension_CP:F2}   ACNeg={-hingeData.compression_CP:F2}"
            };

            // Insert all acceptance lines
            for (int i = 0; i < acceptanceLines.Count; i++)
            {
                lines.Insert(tableIndex + 1 + i, acceptanceLines[i]);
            }
        }

        /// <summary>
        /// Finds existing table or creates a new one
        /// </summary>
        private int FindOrCreateTable(List<string> lines, string tableHeader)
        {
            // Look for existing table
            for (int i = 0; i < lines.Count; i++)
            {
                if (lines[i].Trim() == tableHeader)
                {
                    return i;
                }
            }

            // Table doesn't exist, create it at the end
            lines.Add("");
            lines.Add(tableHeader);
            return lines.Count - 1;
        }
    }

    /// <summary>
    /// Data structure to hold hinge calculation results
    /// </summary>
    public class BraceHingeData
    {
        // Original calculated values (before control factors)
        public double compression_a { get; set; }
        public double compression_b { get; set; }
        public double compression_c { get; set; }
        public double compression_IO { get; set; }
        public double compression_LS { get; set; }
        public double compression_CP { get; set; }
        
        // Final values (after applying control factors) - these are used in $2k file
        public double tension_a { get; set; }
        public double tension_b { get; set; }
        public double tension_c { get; set; }
        public double tension_IO { get; set; }
        public double tension_LS { get; set; }
        public double tension_CP { get; set; }

        public double t_y { get; set; }
        public double delta_t { get; set; }
        public double p_y { get; set; }
        public double delta_c { get; set; }
    }
} 