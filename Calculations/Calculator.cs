using SAP2000v1;

namespace AESPerformansApp.Calculations
{
    /// <summary>
    /// Defines a standard for all calculator classes.
    /// Each calculator must implement this interface to ensure it can perform a calculation.
    /// </summary>
    public interface Calculator
    {
        /// <summary>
        /// Performs the specific calculation for an element.
        /// </summary>
        /// <param name="inputs">The data model containing all necessary inputs and where results will be stored.</param>
        /// <param name="sapModel">The active SAP2000 model object to fetch data from.</param>
        /// <param name="sapFilePath">The path to the SAP2000 file for $2k writing.</param>
        void PerformCalculation(CalculationInputBase inputs, cSapModel sapModel, string sapFilePath);
    }
} 