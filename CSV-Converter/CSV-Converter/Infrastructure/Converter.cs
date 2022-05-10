﻿using CSV_Converter.Contracts.Responses;
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
        private int _fileIterationName = 1; // The placeholding name for the produced CSV files
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

        public ConvertResponse Convert(string filePath)
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

                                System.Diagnostics.Debug.WriteLine(cellData); // Write to debug for testing purposes

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

                                if (iterations == 5)
                                {
                                    CreateCSVFile(cellEntries); // Create a csv file with the current 6 stored cell entries
                                    numberOfFilesCreated++;
                                    _fileIterationName++;
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

            // Add a zero in front of the iterator if it's below 10
            string iteration = (_fileIterationName >= 10) ? _fileIterationName.ToString() : $"0{_fileIterationName}";

            var filePath = $"CSV files\\SL_{iteration}.csv"; // Name and path of the csv file

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
