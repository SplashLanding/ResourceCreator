using FileHelpers;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Xml;
using Excel = Microsoft.Office.Interop.Excel;
using System.Resources;
using System.Threading;

namespace Resource_Creator
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        public MainWindow()
        {
            InitializeComponent();
            tbHowToUse.Text = "To use this magical thing, you will need to have a csv file and a blink resx file that you created through VS. Plug in the file locations and press 'Create' and let the magic happen. :)";
        }

        public void DoMagic()
        {
            try
            {
                XmlDocument doc = new XmlDocument();
                string filePath = txtResxLocation.Text;
                doc.Load(filePath);
                XmlElement root = doc.DocumentElement;

                XmlElement datum = null;
                XmlElement value = null;
                XmlAttribute datumName = null;
                XmlAttribute datumSpace = doc.CreateAttribute("xml:space");
                datumSpace.Value = "preserve";

                var excelTranlationSheet = ParseCSV().ToList();
                //var excelTranlationSheet = ReadExcelFile();

                foreach (var pair in excelTranlationSheet)
                {
                    datum = doc.CreateElement("data");
                    datumName = doc.CreateAttribute("name");
                    datumName.Value = pair.Name;
                    value = doc.CreateElement("value");
                    value.InnerText = pair.Value;

                    datum.Attributes.Append(datumName);
                    datum.Attributes.Append(datumSpace);
                    datum.AppendChild(value);
                    root.AppendChild(datum);
                }

                doc.Save(filePath);
                tbOutput.Text = "Success";
                MessageBox.Show("Look here (" + txtResxLocation + ") Did I get it right? If not you should really talk to my creator, after all I am only a magic novice.", "Shazam!!!" );
            }
            catch (Exception ex)
            {
                MessageBox.Show("Oh boy the magic trick didn't work. :( Maybe you can help me with the trick look to the output box. If you don't understand the output, like me, you should talk to my creator.", "Oh now!");
                tbOutput.Text = ex.GetBaseException().ToString();
            }
        }

        public IEnumerable<TranslationModel> ParseCSV()
        {
            var engine = new FileHelperEngine<TranslationModel>();
            var result = engine.ReadFile(txtExcelFile.Text);
            return result;
        }

        private void btnCreate_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("Press 'Ok' and I will perform some magic for you...", "Prepare yourself :)");
            DoMagic();
            //CreateResourceFile();
            //ReadExcelFile();
        }

        private void btnBrowseResx_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();

            bool? userClickedOK = openFileDialog.ShowDialog();

            // Process input if the user clicked OK.
            if (userClickedOK == true)
            {
                txtResxLocation.Text = openFileDialog.FileName;
            }
        }

        public List<TranslationModel> ReadExcelFile()
        {
            try
            {
                Excel.Application xlApp = new Excel.Application();
                Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(txtExcelFile.Text);
                Excel.Worksheet xlWorksheet = xlWorkbook.Sheets[1];
                Excel.Range xlRange = xlWorksheet.UsedRange;

                int rowCount = xlRange.Rows.Count;
                int colCount = xlRange.Columns.Count;

                List<TranslationModel> translationModelList = new List<TranslationModel>();

                //Gets Names
                for (int i = 1; i < rowCount; i++)
                {
                    TranslationModel model = new TranslationModel();

                    //Gets Names
                    for (int j = 1; j < colCount; j++)
                    {
                        model.Name = xlRange.Cells[i, j].Value2.ToString();
                        //MessageBox.Show(xlRange.Cells[i, j].Value2.ToString());
                    }

                    //Gets Values
                    for (int j = 2; j <= colCount; j++)
                    {
                        model.Value = xlRange.Cells[i, j].Value2.ToString();
                        //MessageBox.Show(xlRange.Cells[i, j].Value2.ToString());
                    }

                    translationModelList.Add(model);
                }

                return translationModelList;
            }
            catch (Exception ex)
            {
                ex.GetBaseException();
                throw;
            }
        }

        private void btnBrowseExcel_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();

            bool? userClickedOK = openFileDialog.ShowDialog();

            // Process input if the user clicked OK.
            if (userClickedOK == true)
            {
                txtExcelFile.Text = openFileDialog.FileName;
            }
        }

        private void CreateResourceFile()
        {
            using (ResourceWriter resxFile = new ResourceWriter("Test.resx"))
            {
                var excelFile = ReadExcelFile();
                foreach (var item in excelFile)
                {
                    resxFile.AddResource(item.Name, item.Value);
                }
                resxFile.Generate();
                resxFile.Close();
            }         
        }
    }
}