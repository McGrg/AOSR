using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using Microsoft.Win32;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Excel = Microsoft.Office.Interop.Excel;
//using Spire.Xls;
using PdfSharp;

namespace AOSR
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private string filename, docNumber, date, florNumber, project, kindofwork;
        private string filenameSource;
        private Excel.Application app = null;
        private Excel.Workbook wBook = null;
        private Excel.Worksheet wSheet = null;
        private Excel.Range range = null;
        private Excel.Application appSource = null;
        private Excel.Workbook wBookSource = null;
        private Excel.Worksheet wSheetSource = null;
        //private List<String> kindofworks = new List<String>();
        //private List<String> projects = new List<String>();

        public MainWindow()
        {
            InitializeComponent();
            JobComboBox.MaxDropDownHeight = 100;
            ProjectComboBox.MaxDropDownHeight = 100;
        }

        private void InsertBtn_Click(object sender, RoutedEventArgs e)
        {
            wSheet.Cells[11,2].Value = DocNumberTextBox.Text;
            wSheet.Cells[11, 9].Value = DateTextBox.Text;
        }

        private void OpenBtn_Click(object sender, RoutedEventArgs e)
        {
            Cleaning();
            try
            {
                OpenFileDialog opf = new OpenFileDialog();
                opf.Filter = "Документы Excel(*.xls;*.xlsx)|*.xls;*.xlsx";
                opf.ShowDialog();
                filename = opf.FileName;
                FileNameTextBox.Text = filename;
                app = new Excel.Application();
                wBook = app.Workbooks.Open(filename);
                wSheet = (Excel.Worksheet)wBook.Sheets[1];

                docNumber = wSheet.Cells[11, 2].Text;
                date = wSheet.Cells[11, 9].Text;
                florNumber = wSheet.Cells[11, 9].Text;
                FillRanges();


            }
            catch (Exception ex)
            {
                wBook.Close();
                app.Quit();
                MessageBox.Show(ex.Message.ToString(), "Error happend: ");
            }
        }

        private void SaveBtn_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Filter = "Документы Excel(*.xls;*.xlsx)|*.xls;*.xlsx";
            sfd.ShowDialog();
            filename = sfd.FileName;
            wBook.SaveAs(filename);
            wBook.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, filename + ".pdf");
            
        }
        //
        // Заполнение Combobox значениями из файла-источника
        //
        private void OpenSourceBtn_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                OpenFileDialog opf = new OpenFileDialog();
                opf.Filter = "Документы Excel(*.xls;*.xlsx)|*.xls;*.xlsx";
                opf.ShowDialog();
                filename = opf.FileName;
                FileNameTextBox.Text = filename;
                app = new Excel.Application();
                wBook = app.Workbooks.Open(filename);
                wSheet = (Excel.Worksheet)wBook.Sheets[1];
                string temp= "test";
                int i = 1;
                while (temp != "")
                {
                    temp = wSheet.Cells[i, 19].Text;
                    //kindofworks.Add(temp);
                    JobComboBox.Items.Add(temp);
                    temp = wSheet.Cells[i, 20].Text;
                    ProjectComboBox.Items.Add(temp);
                    i++;
                }
                
                //docNumber = wSheet.Cells[11, 2].Text;
                //MessageBox.Show(docNumber.ToString(), "Message: ");
                //date = wSheet.Cells[11, 9].Text;
                //florNumber = wSheet.Cells[11, 9].Text;
                //FillRanges();
            }
            catch (Exception ex)
            {
                wBook.Close();
                app.Quit();
                MessageBox.Show(ex.Message.ToString(), "Error happend: ");
            }
        }
                    
        private void ExitBtn_Click(object sender, RoutedEventArgs e)
        {
            Cleaning();
            this.Close();
        }
        private void CancelBtn_Click(object sender, RoutedEventArgs e)
        {
            Cleaning();
        }

        private void Cleaning()
        {
            FileNameTextBox.Text = "";
            DocNumberTextBox.Text = "";
            DateTextBox.Text = "";
            FlorNumberTextBox.Text = "";
            JobComboBox.Items.Clear();
            JobComboBox.SelectedIndex = 0;
            ProjectComboBox.Items.Clear();
            ProjectComboBox.SelectedIndex = 0;
        }
        //
        //заполнение текстовых полей
        //
        private void FillRanges()
        {
            DocNumberTextBox.Text = docNumber;
            DateTextBox.Text = date;
            FlorNumberTextBox.Text = florNumber;
        }
        
    }
}
