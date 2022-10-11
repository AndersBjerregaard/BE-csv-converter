using CSV_Converter.Contracts.Responses;
using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Windows;

namespace CSV_Converter.Infrastructure
{
    public class Converter
    {
        private const string _header = "SN,Slave Address,FullName,"; // The header for every CSV file

        public string ConverterDirectoryPath { get; private set; }

        /// <summary>
        /// Throws FileIO Exceptions.
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        public bool PeakFile(string filePath)
        {
            return File.Exists(filePath);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="targetIterations">The wanted number of cells to iterate through. E.g. if 6, csv files will be produced with 6 lines excluding the header.</param>
        /// <returns></returns>
        public ConvertResponse Convert(string filePath, int targetIterations)
        {
            ConvertResponse convertResponse = new ConvertResponse(); // Holding object for the return value of this method

            using (FileStream stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
            {
                // Read from an OpenXml Excel file (2007 format; *.xlsx)
                using (IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream))
                {
                    System.Diagnostics.Debug.WriteLine("\nExcel File Output:\n");

                    string[] cellEntries = new string[6]; // Backing array for storing the device data
                    int iterations = 0; // Variable storing how many cells of data have been iterated through
                    int numberOfFilesCreated = 0; // Variable storing how many csv files have been produced

                    do
                    {
                        while (excelReader.Read())
                        {
                            try
                            {
                                string cellData = excelReader.GetString(9); // Will throw an out of bounds exception if it attempts to read the Device Data spreadsheet

#if DEBUG
                                System.Diagnostics.Debug.WriteLine(cellData); // Write to debug for testing purposes
#endif

                                if (cellData.Contains("SN")) // This should return true if it reads the column header
                                {
                                    continue; // Skip this cell
                                }

                                // Check if this is the end of elligible cell data
                                if (cellData.StartsWith(',') || String.IsNullOrWhiteSpace(cellData))
                                {
                                    // Finish reading
                                    break;
                                }

                                cellEntries[iterations] = cellData; // Add the cell data to the backing array

                                if (iterations == targetIterations - 1)
                                {
                                    CreateCSVFile(cellEntries); // Create a csv file with the current 6 stored cell entries
                                    numberOfFilesCreated++;
                                    iterations = 0; // Reset the iteration variable
                                } else
                                {
                                    iterations++;
                                }
                            }
                            catch (Exception) 
                            {
                                excelReader.NextResult();
                            }
                        }
                    } while (excelReader.NextResult());

                    convertResponse.Success = true;
                    convertResponse.NumberOfFilesProduced = numberOfFilesCreated;
                    convertResponse.DirectoryPath = ConverterDirectoryPath;
                }

                return convertResponse;
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="targetIterations">The wanted number of cells to iterate through. E.g. if 6, csv files will be produced with 6 lines excluding the header.</param>
        /// <param name="wantedRepetitions">The wanted number of devices to include in the same csv file. E.g. if 3, then include 3 devices at a time.</param>
        /// <returns></returns>
        public ConvertResponse Convert(string filePath, int targetIterations, int wantedRepetitions)
        {
            ConvertResponse convertResponse = new ConvertResponse(); // Holding object for the return value of this method

            using (FileStream stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
            {
                // Read from an OpenXml Excel file (2007 format; *.xlsx)
                using (IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream))
                {
                    System.Diagnostics.Debug.WriteLine("\nExcel File Output:\n");

                    int boundary = targetIterations * wantedRepetitions;

                    string[] cellEntries = new string[boundary]; // Backing array for storing the device data
                    int iterations = 0; // Variable storing how many cells of data have been iterated through
                    int numberOfFilesCreated = 0; // Variable storing how many csv files have been produced

                    //bool lastDataSalvaged = false;

                    do
                    {
                        while (excelReader.Read())
                        {
                            try
                            {
                                string cellData = excelReader.GetString(9); // Will throw an out of bounds exception if it attempts to read the Device Data spreadsheet

#if DEBUG
                                System.Diagnostics.Debug.WriteLine(cellData); // Write to debug for testing purposes
#endif

                                if (cellData.Contains("SN")) // This should return true if it reads the column header
                                {
                                    continue; // Skip this cell
                                }

                                // Check if this is the end of elligible cell data
                                if (cellData.StartsWith(',') || String.IsNullOrWhiteSpace(cellData))
                                {
                                    //lastDataSalvaged = true;
                                    var lastBitOfData = new string[iterations];
                                    for (int i = 0; i < iterations; i++)
                                    {
                                        lastBitOfData[i] = cellEntries[i];
                                    }
                                    CreateCSVFile(lastBitOfData);
                                    numberOfFilesCreated++;
                                    // Finish reading
                                    break;
                                }

                                cellEntries[iterations] = cellData; // Add the cell data to the backing array

                                if (iterations == (boundary) - 1)
                                {
                                    CreateCSVFile(cellEntries); // Create a csv file with the current 6 stored cell entries
                                    numberOfFilesCreated++;
                                    iterations = 0; // Reset the iteration variable
                                }
                                else
                                {
                                    iterations++;
                                }
                            }
                            catch (Exception)
                            {
                                excelReader.NextResult();
                            }
                        }
                    } while (excelReader.NextResult());

                    convertResponse.Success = true;
                    convertResponse.NumberOfFilesProduced = numberOfFilesCreated;
                    convertResponse.DirectoryPath = ConverterDirectoryPath;
                    //if (lastDataSalvaged)
                    //{
                    //    convertResponse.SalvagedData = "The last csv file was not produced with the amount of data corresponding to the wanted amount of css";
                    //}
                }

                return convertResponse;
            }
        }

        private void CreateCSVFile(string[] data)
        {
            // If directory storing the csv files does not exist, create it
            if (!Directory.Exists("CSV files"))
            {
                ConverterDirectoryPath = Directory.CreateDirectory("CSV files").FullName;
            } else if (ConverterDirectoryPath is null)
            {
                // Set the directory path to the existing one
                ConverterDirectoryPath = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase);
            }

            // Since the name of the file should be that of the device name. We just read the first cell of data, split it. Look at the last bit and cut off the number associated with it.
            string fileName = data[0].Split(',')[2].Split('_')[0];

            var filePath = $"CSV files\\{fileName}.csv"; // Name and path of the csv file

            // Check if a file with the same name exists. If so, override it
            if (File.Exists(filePath))
            {
                File.Delete(filePath);
            }

            // Create the csv file and write to it
            using (StreamWriter writer = new StreamWriter(new FileStream(filePath, FileMode.Create, FileAccess.Write)))
            {
                // Write the header
                writer.WriteLine(_header);

                foreach (string cell in data)
                {
                    // Write the 6 cells of data
                    writer.WriteLine(cell);
                }
            }
        }
    }
}
