using System;
using System.Collections;
using System.Windows;
using Microsoft.Win32;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using System.Collections.Generic;





namespace AOSR
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private string filename, docNumber, date, florNumber, project, kindofwork, height, designer, material, nextWork, documents;
        private const string filenameSource= "D:\\Работа\\Интеллект Про\\Корпус 3\\АОСР\\Source.xlsx"; //расположение файла источника для полей формы
        private const string appdir = "D:\\Работа\\Интеллект Про\\Корпус 3\\АОСР\\Сертификаты"; // расположение папки с сертификатами и приложениями
        private const string toSaveDir = "D:\\Работа\\Интеллект Про\\Корпус 3\\АОСР\\";
        private const string axes = "в осях 1/А3-32/М3";
        private Excel.Application app = null;
        private Excel.Workbook wBook = null;
        private Excel.Worksheet wSheet = null;
        private List<string> appList = new List<string>(); //список всех файлов в папке с приложениями
        private List<string> fileToAdd = new List<string>(); //список файлов с приложениями для добавления в один PDF файл

        public MainWindow()
        {
            InitializeComponent();
            JobComboBox.MaxDropDownHeight = 100;
            ProjectComboBox.MaxDropDownHeight = 100;
            MaterialsComboBox.MaxDropDownHeight = 100;
            NextWorkComboBox.MaxDropDownHeight = 100;
            ApplicationsComboBox.MaxDropDownHeight = 100;
            LoadSource(); //загрузка данных в приложение за файла источника для полей формы
            GetAppList(); //загрузка в appList списка файлов из папки с приложениями
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
                string temp = "test";
                int i = 1;
                while (temp != "")
                {
                    temp = wSheetSource.Cells[i, 5].Text;
                    ProjectComboBox.Items.Add(temp);
                    temp = wSheetSource.Cells[i, 6].Text;
                    MaterialsComboBox.Items.Add(temp);
                    temp = wSheetSource.Cells[i, 7].Text;
                    NextWorkComboBox.Items.Add(temp);
                    temp = wSheetSource.Cells[i, 8].Text;
                    ApplicationsComboBox.Items.Add(temp);
                    temp = wSheetSource.Cells[i, 4].Text;
                    JobComboBox.Items.Add(temp);
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

        //
        //заполнение выходного файла Excel данными на основании заполненной формы
        //
        private void InsertBtn_Click(object sender, RoutedEventArgs e)
        {
           
            if (CheckFilled())
                {
                designer = ProjectComboBox.SelectedValue.ToString();
                kindofwork = JobComboBox.SelectedValue.ToString();
                material = MaterialsComboBox.SelectedValue.ToString();
                nextWork = NextWorkComboBox.SelectedValue.ToString();
                //documents = ApplicationsComboBox.SelectedValue.ToString();
                florNumber = FlorNumberTextBox.Text;
                height = HeightChoose(florNumber);
                kindofwork = Phrase(kindofwork, florNumber, height); //добавление к работам этажа, отметки, осей
                nextWork = Phrase(nextWork, florNumber, height); //добавление к работам этажа, отметки, осей
                wSheet.Cells[11, 2].Value = DocNumberTextBox.Text;
                wSheet.Cells[11, 9].Value = DateTextBox.Text;
                wSheet.Cells[24, 1].Value = kindofwork;
                wSheet.Cells[26, 1].Value = designer;
                wSheet.Cells[29, 1].Value = material;
                wSheet.Cells[40, 1].Value = nextWork;
                documents = "";// заглушка!!!
                //
                // добавление имен файлов в список для создания pdf в соответствии с указанными материалами
                //
                string dataSource = material.ToLower().Replace("-", "/"); 
                foreach (string item in appList)
                {
                    if (SearchText(item, dataSource))
                    {
                        string itemFullName = appdir + "\\" + item;
                        fileToAdd.Add(itemFullName);
                        documents = documents + item.Replace(".pdf","") + ", ";
                    }
                }
                wSheet.Cells[47, 1].Value = documents;
            }
            else
            {
                MessageBox.Show("CheckBoxes should be filled first", "Error happend: ");
            }
        }
        //
        // сохранение файла XLS и многостраничного PDF с приложениями
        //
        private void SaveBtn_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Filter = "Документы Excel(*.xls;*.xlsx)|*.xls;*.xlsx";
            sfd.ShowDialog();
            filename = sfd.FileName;
            wBook.SaveAs(filename);
            filename = filename + ".pdf";
            wBook.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, filename);
            wBook.Close();
            MakePDF(filename, fileToAdd);
        }

        private void PropertyButton_Click(object sender, RoutedEventArgs e)
        {
            Prop propWindow = new Prop();
            propWindow.Owner = this;
            propWindow.Show();
        }

        private void CancelBtn_Click(object sender, RoutedEventArgs e)
        {
            Cleaning();
        }

        private void ExitBtn_Click(object sender, RoutedEventArgs e)
        {
            Cleaning();
            this.Close();
        }

        private bool CheckFilled ()
        {
            if (JobComboBox.SelectedValue !=null && ProjectComboBox.SelectedValue!=null && MaterialsComboBox.SelectedValue !=null && NextWorkComboBox.SelectedValue !=null)
            {
                return true;
            }
            else return false;
        }

        private void Cleaning()
        {
            FileNameTextBox.Text = "";
            DocNumberTextBox.Text = "";
            DateTextBox.Text = "";
            FlorNumberTextBox.Text = "";
        }

        //
        //заполнение текстовых полей формы
        //
        private void FillRanges()
        {
            DocNumberTextBox.Text = docNumber;
            DateTextBox.Text = date;
        }

        //
        // высотная отметка проведения работ в зависимости от этажа
        //
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
            else height = (8.4 + (h - 4) * 2.8).ToString()+"00";//добавить вариант, если с десятичными знаками и без.
            return height;
        }
        //
        // возвращает фразу с данными по работам, этажу, осями, отметками
        //
        private string Phrase (string textToForm, string florNumber, string height)
        {
            string text = textToForm + " " + florNumber + " этаж, на отм." + height + ", " + axes;
            return text;
        }

        //
        //создание многостраничного PDF с приложениями, на входе исходный PDF файл и список файлов для приложения
        //

        private void MakePDF(string pdfFileName, List<string> fileArray)
        {

            try
            {
                //MessageBox.Show(pdfFileName, "Opening first PDF");
                // Open the output document
                PdfDocument outputDocument = PdfReader.Open(pdfFileName);

                // Iterate files
                foreach (string file in fileArray)
                {
                    //MessageBox.Show(file, "Opening PDF");
                    // Open the document to import pages from it.
                    PdfDocument inputDocument = PdfReader.Open(file, PdfDocumentOpenMode.Import);

                    // Iterate pages
                    int count = inputDocument.PageCount;
                    for (int idx = 0; idx < count; idx++)
                    {
                        // Get the page from the external document...
                        PdfPage page = inputDocument.Pages[idx];
                        // ...and add it to the output document.
                        outputDocument.AddPage(page);
                    }

                    outputDocument.Save(pdfFileName);
                }
            }
            catch (Exception ex)
            {
                //
                //сделать обработчик исключения, создать отдельную папку и сохранить в нее файл PDF и все приложения PDF
                //
                MessageBox.Show(ex.ToString() + " all applications have been saved in separate directory", "PDF saving file error!");
            }

        }

        //
        //заполнение appList файлами из директории с приложениями
        //
        private void GetAppList()
        {
            try
            {
                DirectoryInfo dir = new DirectoryInfo(appdir);
                if (dir.Exists)
                foreach (var item in dir.GetFiles())
                {
                    appList.Add(item.Name);
                }
                else
                    MessageBox.Show("Such directory doesn't exist!", "Error happend in opening application directory: ");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), "Error happend in opening application directory: ");
            }
        }

        //
        //поиск в материалах данных для списка приложений
        //
        private static bool SearchText(string example, string source)
        {
            example = example.ToLower().Trim().Remove(example.Length - 4, 4).Replace("-", "/");
            int index = example.IndexOf('№');
            if ((index == 0) || (index > 0))
            {
                example = example.Substring(index);
                index = example.IndexOf("от");
                if ((index == 0) || (index > 0))
                    example = example.Substring(0, index);
            }
            if (source.Contains(example))
            {
                return true;
            }
            else return false;
        }

    }
}
