using System;
using System.Collections.Generic;
using System.Text;

namespace CSV_Converter.Contracts.Responses
{
   public class ConvertResponse
    {
        public bool Success { get; set; }
        public IEnumerable<string> Errors { get; set; }
        public int NumberOfFilesProduced { get; set; }
        public string DirectoryPath { get; set; }
        public string SalvagedData { get; set; }
    }
}
