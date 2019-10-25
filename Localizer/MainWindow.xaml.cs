using System;
using Microsoft.Win32;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Media;

namespace Localizer
{
    public partial class MainWindow : Window
    {
        public readonly ResourceModel[] LanguageResources = new ResourceModel[]
        {
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

        public MainWindow()
        {
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.DefaultExt = ".txt";
            openFileDialog.Filter = "Text Document (.txt)|*.txt";
            if (openFileDialog.ShowDialog() == true)
            {
                string fileName = openFileDialog.FileName;
                filePathTextBox.Text = fileName;
                var text = File.ReadAllText(fileName);
                //logRichText.Document.Blocks.Clear();
                //logRichText.Document.Blocks.Add(new Paragraph(new Run(text)));
            }
        }

        private void AddNewKey_Click(object sender, RoutedEventArgs e)
        {
            var newKeyValue = newKeyNameTextBox.Text;
            if (string.IsNullOrWhiteSpace(newKeyValue))
            {
                _ = MessageBox.Show("Key Name cannot be empty.", "Warning", MessageBoxButton.OK);
            }
            else if (string.IsNullOrWhiteSpace(xlsPathTextBox.Text) || string.IsNullOrWhiteSpace(filePathTextBox.Text))
            {
                _ = MessageBox.Show("Please select xls file path and folder with translates.", "Warning", MessageBoxButton.OK);
            }
            else
            {
                //TODO add all keys to files
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
                        var textResource = File.ReadAllText($@"{folderPath}\{languageResource.FileName}.json");
                        logRichText.AppendText(": OK", "Green");

                    }
                    catch (FileNotFoundException ex)
                    {
                        isSmthMissed = true;
                        logRichText.AppendText(" : NOT FOUND", "Red");
                    }

                    logRichText.AppendText(Environment.NewLine);
                }

                filePathTextBox.IsReadOnly = !isSmthMissed;

                logRichText.AppendText("---------------------------------------------", "Black");
                logRichText.AppendText(Environment.NewLine);
            }
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
    }
}
