using CSV_Converter.Contracts.Responses;
using System;
using System.Collections.Generic;
using System.Text;

namespace CSV_Converter.Infrastructure
{
    public class ConverterController
    {
        private readonly Converter _converter;

        public ConverterController(Converter converter)
        {
            _converter = converter;
        }

        public CanConvertResponse CanConvert(string filePath)
        {
            bool result = _converter.PeakFile(filePath);

            return (result) ? new CanConvertResponse { Success = true } : new CanConvertResponse { Success = false, Errors = new[] { "Could not peak file" } };
        }

        public ConvertResponse Convert(string filePath, int numberOfInverters)
        {
            return _converter.Convert(filePath, numberOfInverters);
        }

        public ConvertResponse Convert(string filePath, int numberOfInverters, int numberOfCss)
        {
            return _converter.Convert(filePath, numberOfInverters, numberOfCss);
        }
    }
}
