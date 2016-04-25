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
using System.Xml;
using System.IO;
using FileHelpers;
using Microsoft.Win32;

namespace Resource_Creator
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
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


                var csv = ParseCSV().ToList();

                foreach (var pair in csv)
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
            }
            catch (Exception ex)
            {
                tbOutput.Text = ex.GetBaseException().ToString();
            }
        }

        public IEnumerable<TranslationModel> ParseCSV()
        {
            var engine = new FileHelperEngine<TranslationModel>();
            var result = engine.ReadFile(txtCSVLocation.Text);
            return result;
        }

        private void btnCreate_Click(object sender, RoutedEventArgs e)
        {
            DoMagic();
        }

        private void btnBrowseCSV_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "CSV Files (.csv)|*.txt|All Files (*.*)|*.*";
            openFileDialog.FilterIndex = 1;

            bool? userClickedOK = openFileDialog.ShowDialog();

            // Process input if the user clicked OK.
            if (userClickedOK == true)
            {
                txtCSVLocation.Text = openFileDialog.FileName;
            }

        }

        private void btnBrowseResx_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "RESX Files (.resx)|*.txt|All Files (*.*)|*.*";
            openFileDialog.FilterIndex = 1;

            bool? userClickedOK = openFileDialog.ShowDialog();

            // Process input if the user clicked OK.
            if (userClickedOK == true)
            {
                txtResxLocation.Text = openFileDialog.FileName;
            }
        }
    }
}
