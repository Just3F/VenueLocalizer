using Microsoft.Win32;
using System.IO;
using System.Windows;
using System.Windows.Documents;

namespace Localizer
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public readonly ResourceModel[] resources = new ResourceModel[]
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
                mainTextContent.Document.Blocks.Clear();
                mainTextContent.Document.Blocks.Add(new Paragraph(new Run(text)));
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
            OpenFileDialog folderBrowser = new OpenFileDialog();
            // Set validate names and check file exists to false otherwise windows will
            // not let you select "Folder Selection."
            folderBrowser.ValidateNames = false;
            folderBrowser.CheckFileExists = false;
            folderBrowser.CheckPathExists = true;
            // Always default to Folder Selection.
            folderBrowser.FileName = "Folder Selection.";
            if (folderBrowser.ShowDialog() == true)
            {
                string folderPath = System.IO.Path.GetDirectoryName(folderBrowser.FileName);
                filePathTextBox.Text = folderPath;
                // ...
            }
        }
    }

    public class ResourceModel
    {
        public string FileName { get; set; }
    }
}
