using System;
using System.Collections.Generic;
using System.Text;

namespace CSV_Converter.Contracts.Responses
{
    public class CanConvertResponse
    {
        public bool Success { get; set; }
        public IEnumerable<string> Errors { get; set; }
    }
}
