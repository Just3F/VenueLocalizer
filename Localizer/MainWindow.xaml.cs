using System;
using System.Collections.Generic;
using Microsoft.Win32;
using System.IO;
using System.Linq;
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
        public Workbook workbook;
        public int NameExcelIndex;

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

                    workbook = new Workbook(fstream);
                    fstream.Close();

                    var worksheet = workbook.Worksheets[0];
                    var columns = worksheet.Cells.Columns.OfType<Column>();
                    var rows = worksheet.Cells.Rows.OfType<Row>();

                    var lastRowIndex = rows.FirstOrDefault(x => x.FirstDataCell == null).Index;

                    foreach (var languageResource in LanguageResources)
                    {
                        for (int i = 0; i < columns.Count(); i++)
                        {
                            var langColumn = rows.FirstOrDefault()[i];
                            if (langColumn.StringValue == languageResource.ColumnExcelName)
                            {
                                languageResource.ColumnExcelIndex = i;
                            }
                        }
                    }
                    for (int i = 0; i < columns.Count(); i++)
                    {
                        var langColumn = rows.FirstOrDefault()[i];
                        if (langColumn.StringValue == "Name")
                        {
                            NameExcelIndex = i;
                        }
                    }

                    logRichText.AppendText(" : OK", "Green");
                    xlsPathTextBox.Background = Brushes.LightGreen;
                    IsXlsLoaded = true;
                }
                catch (FileNotFoundException)
                {
                    logRichText.AppendText(" : NOT FOUND", "Red");
                    xlsPathTextBox.Background = Brushes.IndianRed;
                }
                catch (IOException)
                {
                    logRichText.AppendText(" : EXCEL Already opened in another application", "Red");
                    xlsPathTextBox.Background = Brushes.IndianRed;
                }

                ConsoleNewLine();
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
                var worksheet = workbook.Worksheets[0];
                var columns = worksheet.Cells.Columns.OfType<Column>();
                var rows = worksheet.Cells.Rows.OfType<Row>();

                var lastRowIndex = rows.FirstOrDefault(x => x.FirstDataCell == null).Index;
                worksheet.Cells[lastRowIndex, NameExcelIndex].PutValue(newKeyValue);

                foreach (var languageResource in LanguageResources)
                {
                    languageResource.ResourceParsedText.Add(newKeyValue, defaultValue);
                    var serializedObject = JsonConvert.SerializeObject(languageResource.ResourceParsedText, Formatting.Indented);
                    File.WriteAllText(languageResource.FullPath, serializedObject);


                    var cellValue = languageResource.ColumnExcelName == "English" ? defaultValue : "";
                    worksheet.Cells[lastRowIndex, languageResource.ColumnExcelIndex].PutValue(cellValue);
                    workbook.Save(xlsPathTextBox.Text);


                    logRichText.AppendText($"New key for {languageResource.FileName}", "Black");
                    logRichText.AppendText(" : OK", "Green");
                }

                ConsoleNewLine();
            }
        }

        private void LoadTranlsatesFromXLS_Click(object sender, RoutedEventArgs e)
        {


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
        public int ColumnExcelIndex { get; set; }
        public string ResourceText { get; set; }
        public string FullPath { get; set; }
        public Dictionary<string, string> ResourceParsedText { get; set; }
    }
}
