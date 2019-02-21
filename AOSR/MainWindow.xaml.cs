﻿using System;
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
        private string filename, docNumber, date, florNumber, project, kindofwork, height, designer, material, nextWork, documents;
        private const string filenameSource= "D:\\Работа\\Интеллект Про\\Корпус 3\\АОСР\\Source.xlsx";
        private const string axes = "в осях 1/А3-32/М3";
        private Excel.Application app = null;
        private Excel.Workbook wBook = null;
        private Excel.Worksheet wSheet = null;

        public MainWindow()
        {
            InitializeComponent();
            LoadSource();
            JobComboBox.MaxDropDownHeight = 100;
            ProjectComboBox.MaxDropDownHeight = 100;
            MaterialsComboBox.MaxDropDownHeight = 100;
            NextWorkComboBox.MaxDropDownHeight = 100;
            ApplicationsComboBox.MaxDropDownHeight = 100;
        }

        private void InsertBtn_Click(object sender, RoutedEventArgs e)
        {
            if (CheckFilled())
                {
                designer = ProjectComboBox.SelectedValue.ToString();
                kindofwork = JobComboBox.SelectedValue.ToString();
                material = MaterialsComboBox.SelectedValue.ToString();
                nextWork = NextWorkComboBox.SelectedValue.ToString();
                documents = ApplicationsComboBox.SelectedValue.ToString();
                florNumber = FlorNumberTextBox.Text;
                height = HeightChoose(florNumber);
                kindofwork = Phrase(kindofwork, florNumber, height);
                wSheet.Cells[11, 2].Value = DocNumberTextBox.Text;
                wSheet.Cells[11, 9].Value = DateTextBox.Text;
                wSheet.Cells[24, 1].Value = kindofwork;
                wSheet.Cells[26, 1].Value = designer;
                wSheet.Cells[29, 1].Value = material;
                wSheet.Cells[40, 1].Value = nextWork;
                wSheet.Cells[47, 1].Value = documents;
            }
            else
            {
                MessageBox.Show("CheckBoxes should be filled first", "Error happend: ");
            }
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
                FillRanges();
            }
            catch (Exception ex)
            {
                wBook.Close();
                app.Quit();
                MessageBox.Show(ex.Message.ToString(), "Error happend: ");
            }
        }

        private bool CheckFilled ()
        {
            if (JobComboBox.SelectedValue !=null && ProjectComboBox.SelectedValue!=null && MaterialsComboBox.SelectedValue !=null && NextWorkComboBox.SelectedValue !=null && ApplicationsComboBox.SelectedValue !=null)
            {
                return true;
            }
            else return false;
        }

        private void SaveBtn_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Filter = "Документы Excel(*.xls;*.xlsx)|*.xls;*.xlsx";
            sfd.ShowDialog();
            filename = sfd.FileName;
            wBook.SaveAs(filename);
            wBook.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, filename + ".pdf");
            wBook.Close();
            
        }
        //
        // Заполнение Combobox значениями из файла-источника
        // Адрес файла источника D:\\Работа\\Интеллект Про\\Корпус 3\\АОСР\\АОСР.xlsx
        //
        private void LoadSource()
        {
          Excel.Application appSource = null;
          Excel.Workbook wBookSource = null;
          Excel.Worksheet wSheetSource = null;
            try
            {
                appSource = new Excel.Application();
                wBookSource = appSource.Workbooks.Open(filenameSource);
                wSheetSource = (Excel.Worksheet)wBookSource.Sheets[1];
                string temp= "test";
                int i = 1;
                while (temp != "")
                {
                    temp = wSheetSource.Cells[i, 4].Text;
                    JobComboBox.Items.Add(temp);
                    temp = wSheetSource.Cells[i, 5].Text;
                    ProjectComboBox.Items.Add(temp);
                    temp = wSheetSource.Cells[i, 6].Text;
                    MaterialsComboBox.Items.Add(temp);
                    temp = wSheetSource.Cells[i, 7].Text;
                    NextWorkComboBox.Items.Add(temp);
                    temp = wSheetSource.Cells[i, 8].Text;
                    ApplicationsComboBox.Items.Add(temp);
                    i++;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), "Error happend: ");
            }
            finally
            {
                wBookSource.Close();
                appSource.Quit();
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
        }

        //
        //заполнение текстовых полей
        //
        private void FillRanges()
        {
            DocNumberTextBox.Text = docNumber;
            DateTextBox.Text = date;
        }

        private string HeightChoose(string florNumb)
        {
            double h = 0;
            string height = "";
            try
            {
                h = double.Parse(florNumb);
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), "Error happend in parsing: ");
            }
            if (h<4)
                {
                   switch (h)
                   {
                    case 1:
                        height = "0.000";
                        break;
                    case 2:
                        height = "3.200";
                        break;
                    case 3:
                        height = "5.600";
                        break;
                    default:
                        height = "0";
                        break;
                   }
                }
            else height = (8.4 + (h - 4) * 2.8).ToString()+"00";
            return height;
        }
        //
        // возвращает фразу с данными по работам, этажу, осями, отметками
        //
        private string Phrase (string kindOfWork, string florNumber, string height)
        {
            string text = kindOfWork + " " + florNumber + " этаж, на отм." + height + ", " + axes;
            return text;
        }

    }
}
