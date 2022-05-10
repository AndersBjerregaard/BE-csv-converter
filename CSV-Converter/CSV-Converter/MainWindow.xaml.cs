using CSV_Converter.Infrastructure;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace CSV_Converter
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private readonly ConverterController _controller;

        public MainWindow()
        {
            // his is required to parse strings in binary BIFF2-5 Excel documents encoded with DOS-era code pages. These encodings are registered by default in the full .NET Framework, but not on .NET Core.
            // See https://github.com/ExcelDataReader/ExcelDataReader for reference
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            Converter converter = new Converter();

            _controller = new ConverterController(converter);

            InitializeComponent();
        }

        private void FileSelectionButton_Click(object sender, RoutedEventArgs e)
        {
            // Create OpenFileDialog
            OpenFileDialog fileDialog = new OpenFileDialog
            {
                // Set filter for file extension and default file extension
                DefaultExt = "xlsm",
                Filter = "Excel Macro-Enabled Workbook (*.xlsm)|*.xlsm|Excel Workbook (*.xlsx)|*.xlsx|Excel Binary Workbook (*.xlsb)|*.xlsb|Excel 97-2003 Workbook (*.xls)|*.xls"
            };

            // Display OpenFileDialog by calling ShowDialog method
            Nullable<bool> fileDialogResult = fileDialog.ShowDialog();

            // Get the selected file name and display in a TextBox
            if (fileDialogResult.GetValueOrDefault())
            {
                // Open document
                string filename = fileDialog.FileName;
                FilePathTextBox.Text = filename;
            }
        }

        private void ConvertButton_Click(object sender, RoutedEventArgs e)
        {
            // Check if there is any file selected
            if (String.IsNullOrWhiteSpace(FilePathTextBox.Text))
            {
                // Define a message box
                string messageBoxText = "No file was selected.";
                string caption = "Error";
                MessageBoxButton button = MessageBoxButton.OK;
                MessageBoxImage icon = MessageBoxImage.Warning;

                // Show the message box
                MessageBoxResult messageBoxResult = MessageBox.Show(messageBoxText, caption, button, icon, MessageBoxResult.OK);
            } else
            {
                // Check if filepath and file are elligible
                Contracts.Responses.CanConvertResponse canConvertResult = _controller.CanConvert(FilePathTextBox.Text);

                if (!canConvertResult.Success)
                {
                    // Do some stuff
                    string errors = String.Join("\n", canConvertResult.Errors);

                    MessageBox.Show(errors);

                }
                else // File peak was succesful
                {
                    if (!TryParseInvertersTextBox())
                    {
                        MessageBox.Show("Could not recognize input in 'Numbers of inverters' as only numbers", "Error", MessageBoxButton.OK, MessageBoxImage.Hand);
                    } else
                    {
                        int numberOfInverters = int.Parse(InvertersTextBox.Text);

                        // Attempt to convert to csv files
                        Contracts.Responses.ConvertResponse result = _controller.Convert(FilePathTextBox.Text, numberOfInverters);

                        if (result.Success)
                        {
                            MessageBox.Show($"{result.NumberOfFilesProduced} csv files were produced at the following path: {result.DirectoryPath}", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
                        }
                        else
                        {
                            string errors = String.Join("\n", result.Errors);

                            MessageBox.Show($"Oops, something went wrong. Here are the errors:\n{errors}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                        }
                    }
                }
            }
        }

        private void InvertersTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            bool parsed = int.TryParse(InvertersTextBox.Text, out int numberOfInverters);

            if (!parsed)
            {
                InvertersTextBox.Text = 0.ToString();
                MessageBox.Show("Please only enter numbers", "Error", MessageBoxButton.OK, MessageBoxImage.Hand);
            }
        }

        private bool TryParseInvertersTextBox()
        {
            bool parsed = int.TryParse(InvertersTextBox.Text, out int numberOfInverters);

            if (!parsed)
            {
                InvertersTextBox.Text = 0.ToString();
                return false;
            }

            return true;
        }
    }
}
