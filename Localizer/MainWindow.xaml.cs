using System;
using System.Collections.Generic;
using Microsoft.Win32;
using System.IO;
using System.Text.Json.Serialization;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Media;
using Newtonsoft.Json;

namespace Localizer
{
    public partial class MainWindow : Window
    {
        public ResourceModel[] LanguageResources = {
            new ResourceModel { FileName = "de" },
            new ResourceModel { FileName = "en" },
            new ResourceModel { FileName = "es" },
            new ResourceModel { FileName = "fr" },
            new ResourceModel { FileName = "it" },
            new ResourceModel { FileName = "ja" },
            new ResourceModel { FileName = "ko" },
            new ResourceModel { FileName = "pt" },
            new ResourceModel { FileName = "ru" },
            new ResourceModel { FileName = "zh-hans" }
        };

        public bool IsAllTranslatesLoaded = false;
        public bool IsXlsLoaded = false;
        public string ExcelFileContent = "";

        public MainWindow()
        {
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog { Filter = "Excel Files|*.xls;*.xlsx;*.xlsm" };

            if (openFileDialog.ShowDialog() == true)
            {
                string fileName = openFileDialog.FileName;
                xlsPathTextBox.Text = fileName;
                logRichText.AppendText("Excel File", "Black");
                try
                {
                    ExcelFileContent = File.ReadAllText(fileName);
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
                    var serializedObject = JsonConvert.SerializeObject(languageResource.ResourceParsedText);
                    File.WriteAllText(languageResource.FullPath, serializedObject);

                    logRichText.AppendText($"New key for {languageResource.FileName}", "Black");
                    logRichText.AppendText(" : OK", "Green");
                    ConsoleNewLine();
                }

            }
        }

        private void LoadTranlsatesFromXLS_Click(object sender, RoutedEventArgs e)
        {
            //TODO
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
        public string ResourceText { get; set; }
        public string FullPath { get; set; }
        public Dictionary<string, string> ResourceParsedText { get; set; }
    }
}
