using System;
using System.Collections.Generic;
using System.Data;
using Microsoft.Win32;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Media;
using Aspose.Cells;
using Newtonsoft.Json;

namespace Localizer
{
    public partial class MainWindow : Window
    {
        public ResourceModel[] LanguageResources = {
            new ResourceModel { FileName = "de", ColumnExcelName = "German"},
            new ResourceModel { FileName = "en", ColumnExcelName = "English"},
            new ResourceModel { FileName = "es", ColumnExcelName = "Spanish"},
            new ResourceModel { FileName = "fr", ColumnExcelName = "French"},
            new ResourceModel { FileName = "it", ColumnExcelName = "Italian"},
            new ResourceModel { FileName = "ja", ColumnExcelName = "Japanese"},
            new ResourceModel { FileName = "ko", ColumnExcelName = "Korean"},
            new ResourceModel { FileName = "pt", ColumnExcelName = "Portuguese"},
            new ResourceModel { FileName = "ru", ColumnExcelName = "Russian"},
            new ResourceModel { FileName = "zh-hans", ColumnExcelName = "Chinese (Simplified)"}
        };

        public bool IsAllTranslatesLoaded = false;
        public bool IsXlsLoaded = false;

        public MainWindow()
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            InitializeComponent();
        }

        private void ExcelFileSelect_Click(object sender, RoutedEventArgs e)
        {

            var openFileDialog = new OpenFileDialog { Filter = "Excel Files|*.xls;*.xlsx;*.xlsm" };

            if (openFileDialog.ShowDialog() == true)
            {
                string fileName = openFileDialog.FileName;
                xlsPathTextBox.Text = fileName;
                logRichText.AppendText("Excel File", "Black");
                try
                {
                    FileStream fstream = new FileStream(fileName, FileMode.Open);

                    // Instantiate a Workbook object that represents the existing Excel file
                    Workbook workbook = new Workbook(fstream);

                    // Get the reference of "A1" cell from the cells collection of a worksheet
                    Cell cell = workbook.Worksheets[0].Cells["A1"];

                    // Put the "Hello World!" text into the "A1" cell
                    cell.PutValue("Hello World!");
                    fstream.Close();

                    // Save the Excel file
                    workbook.Save(fileName);

                    // Closing the file stream to free all resources
                }
                catch (FileNotFoundException)
                {
                    logRichText.AppendText(" : NOT FOUND", "Red");
                    xlsPathTextBox.Background = Brushes.IndianRed;
                }

                logRichText.AppendText(" : OK", "Green");
                xlsPathTextBox.Background = Brushes.LightGreen;
                IsXlsLoaded = true;

                ConsoleNewLine();
                //logRichText.Document.Blocks.Clear();
                //logRichText.Document.Blocks.Add(new Paragraph(new Run(text)));
            }
        }

        private void AddNewKey_Click(object sender, RoutedEventArgs e)
        {
            var newKeyValue = newKeyNameTextBox.Text;
            var defaultValue = defaultNewKeyValue.Text;
            if (string.IsNullOrWhiteSpace(newKeyValue))
            {
                _ = MessageBox.Show("Key Name cannot be empty.", "Warning", MessageBoxButton.OK);
            }
            else if (string.IsNullOrWhiteSpace(xlsPathTextBox.Text) || string.IsNullOrWhiteSpace(filePathTextBox.Text))
            {
                _ = MessageBox.Show("Please select xls file path and folder with translates.", "Warning", MessageBoxButton.OK);
            }
            else if (!IsXlsLoaded || !IsAllTranslatesLoaded)
            {
                _ = MessageBox.Show("Excel file or translate files was not loaded!", "Warning", MessageBoxButton.OK);
            }
            else
            {
                foreach (var languageResource in LanguageResources)
                {
                    languageResource.ResourceParsedText.Add(newKeyValue, defaultValue);
                    var serializedObject = JsonConvert.SerializeObject(languageResource.ResourceParsedText, Formatting.Indented);
                    File.WriteAllText(languageResource.FullPath, serializedObject);



                    logRichText.AppendText($"New key for {languageResource.FileName}", "Black");
                    logRichText.AppendText(" : OK", "Green");
                    ConsoleNewLine();
                }

            }
        }

        private void LoadTranlsatesFromXLS_Click(object sender, RoutedEventArgs e)
        {
            //var test = ExcelFileContent.AsDataSet();
            //var mySheet = test.Tables[0];
            //var mySheetHeader = mySheet.Rows[0];
            //foreach (var languageResource in LanguageResources)
            //{
            //    DataRow workRow = mySheet.NewRow();
            //    workRow[0] = "asdasdasdas1123";
            //    workRow[4] = "xzcs2131231newKeyValuesdaasdasdasdas123";
            //    mySheet.Rows.Add(workRow);
            //    ExcelFileContent.Close();
            //}
        }

        private void SelectFolder_Click(object sender, RoutedEventArgs e)
        {
            var folderBrowser = new OpenFileDialog
            {
                ValidateNames = false,
                CheckFileExists = false,
                CheckPathExists = true,
                FileName = "Folder Selection."
            };

            if (folderBrowser.ShowDialog() == true)
            {
                string folderPath = System.IO.Path.GetDirectoryName(folderBrowser.FileName);
                filePathTextBox.Text = folderPath;
                bool isSmthMissed = false;
                foreach (var languageResource in LanguageResources)
                {
                    logRichText.AppendText(languageResource.FileName, "Black");

                    try
                    {
                        var fullPath = $@"{folderPath}\{languageResource.FileName}.json";
                        var resText = File.ReadAllText(fullPath);
                        languageResource.FullPath = fullPath;
                        languageResource.ResourceText = resText;
                        languageResource.ResourceParsedText =
                            JsonConvert.DeserializeObject<Dictionary<string, string>>(resText);

                        logRichText.AppendText(": OK", "Green");

                    }
                    catch (FileNotFoundException)
                    {
                        isSmthMissed = true;
                        logRichText.AppendText(" : NOT FOUND", "Red");
                    }
                    catch (Exception ex)
                    {
                        isSmthMissed = true;
                        logRichText.AppendText(ex.Message, "Red");
                    }

                    logRichText.AppendText(Environment.NewLine);
                }

                filePathTextBox.IsReadOnly = IsAllTranslatesLoaded = !isSmthMissed;

                ConsoleNewLine();
                filePathTextBox.Background = IsAllTranslatesLoaded ? Brushes.LightGreen : Brushes.PaleVioletRed;
            }
        }

        private void ConsoleNewLine()
        {
            logRichText.AppendText("---------------------------------------------", "Black");
            logRichText.AppendText(Environment.NewLine);
        }
    }

    public static class Ext
    {
        public static void AppendText(this RichTextBox box, string text, string color)
        {
            BrushConverter bc = new BrushConverter();
            TextRange tr = new TextRange(box.Document.ContentEnd, box.Document.ContentEnd);
            tr.Text = text;
            try
            {
                tr.ApplyPropertyValue(TextElement.ForegroundProperty, bc.ConvertFromString(color));
            }
            catch (FormatException) { }
        }
    }

    public class ResourceModel
    {
        public string FileName { get; set; }
        public string ColumnExcelName { get; set; }
        public string ResourceText { get; set; }
        public string FullPath { get; set; }
        public Dictionary<string, string> ResourceParsedText { get; set; }
    }
}
