using System;
using System.Collections;
using Win = System.Windows;
using Win32 = Microsoft.Win32;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using System.Collections.Generic;
using Forms = System.Windows.Forms;


namespace AOSR
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Win.Window
    {
        private string docNumber, date, florNumber, project, kindofwork, height, designer, material, nextWork, documents, pathName;
        private const string filenameSource= "D:\\Работа\\Интеллект Про\\Корпус 3\\АОСР\\Source.xlsx"; //расположение файла источника для полей формы
        private const string appdir = "D:\\Работа\\Интеллект Про\\Корпус 3\\АОСР\\Сертификаты"; // расположение папки с сертификатами и приложениями
        private const string toSaveDir = "D:\\Работа\\Интеллект Про\\Корпус 3\\АОСР\\";
        private const string listOfWorks = "D:\\Работа\\Интеллект Про\\Корпус 3\\АОСР\\Акты поэтажка без работ по АППОР - копия.xlsx"; // расположение файла с данными по всем актам
        private const string filename = "D:\\Работа\\Интеллект Про\\Корпус 3\\АОСР\\АОСР2 - копия.xls"; // расположение файла образца для актов
        private const string fileAppl = "D:\\Работа\\Интеллект Про\\Корпус 3\\АОСР\\Сопроводиловка.xlsx"; // расположение файла описи
        private const string axes = "в осях 1/А3-32/М3";
        private Excel.Application app = null;
        private Excel.Workbook wBook = null;
        private Excel.Worksheet wSheet = null;
        private List<string> appList = new List<string>(); //список всех файлов в папке с приложениями
        private List<string> fileToAdd=null; //список файлов с приложениями для добавления в один PDF файл

        public MainWindow()
        {
            InitializeComponent();
            GetAppList(); //загрузка в appList списка файлов из папки с приложениями
        }

        private void CancelBtn_Click(object sender, Win.RoutedEventArgs e)
        {
            Cleaning();
        }

        private void ExitBtn_Click(object sender, Win.RoutedEventArgs e)
        {
            Cleaning();
            this.Close();
        }

        private void Cleaning()
        {
            FileNameTextBox.Text = "";
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
                Win.MessageBox.Show(ex.Message.ToString(), "Error happend in parsing: ");
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
                Win.MessageBox.Show(ex.ToString() + " all applications have been saved in separate directory", "PDF saving file error!");
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
                    Win.MessageBox.Show("Such directory doesn't exist!", "Error happend in opening application directory: ");
            }
            catch (Exception ex)
            {
                Win.MessageBox.Show(ex.Message.ToString(), "Error happend in opening application directory: ");
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

        //
        // загрузка данных для создания всех актов ао этажам
        //
        private void LoadListOfWorks()
        {
            Excel.Application appSource = null;
            Excel.Workbook wBookSource = null;
            Excel.Worksheet wSheetSource = null;
            try
            {
                appSource = new Excel.Application();
                wBookSource = appSource.Workbooks.Open(listOfWorks);
                wSheetSource = (Excel.Worksheet)wBookSource.Sheets[1];
                app = new Excel.Application();
                wBook = app.Workbooks.Open(filename);
                wSheet = (Excel.Worksheet)wBook.Sheets[1];
                string temp = "test";
                for (int row =9; row<11; row++ )//счетчик по строкам временно на 1 строку
                { 
                    florNumber = wSheetSource.Cells[row, 1].Text;
                    int act = 0;
                    //Win.MessageBox.Show("florNumber-" + florNumber, "Cell name: ");
                    for (int column =2; column<84; column++)
                    {
                        temp = wSheetSource.Cells[row, column].Text;
                        if (temp!= "")
                        {
                            //Win.MessageBox.Show("row-"+row + "colunn-" + column, "Cell name: ");
                            act++;
                            height = HeightChoose(florNumber);
                            docNumber = "№ 3." + florNumber + "." + act.ToString();
                            date = "«15» ноября 2018 г.";
                            kindofwork = wSheetSource.Cells[7, column].Text + " " + wSheetSource.Cells[6, column].Text + " " + florNumber + " этаж, на отм. +" + height + ", " + axes;
                            //Win.MessageBox.Show("kindofwork-" + kindofwork, "Input: ");
                            designer = wSheetSource.Cells[4, column].Text;
                            //Win.MessageBox.Show("designer-" + designer, "Input: ");
                            material = wSheetSource.Cells[5, column].Text;
                            //Win.MessageBox.Show("material-" + material, "Input: ");
                            nextWork = wSheetSource.Cells[3, column].Text;
                            //Win.MessageBox.Show("nextWork-" + nextWork, "Input: ");
                            documents = "";
                            fileToAdd = new List<string>();
                            string dataSource = material.ToLower().Replace("-", "/");
                            foreach (string item in appList)
                            {
                                if (SearchText(item, dataSource))
                                {
                                    string itemFullName = appdir + "\\" + item;
                                    fileToAdd.Add(itemFullName);
                                    documents = documents + item.Replace(".pdf", "") + ", ";
                                }
                            }
                            Insert();
                            string name = pathName + "\\" + "Акт" + docNumber;
                            Save(name);
                        }
                    }
                }
                Win.MessageBox.Show("Task completed!", "System: ");
            }
            catch (Exception ex)
            {
                Win.MessageBox.Show(ex.Message.ToString(), "Error happend in opening: ");
            }
            finally
            {
                wBookSource.Close();
                appSource.Quit();
                wBook.Close();
                app.Quit();
            }
        }

        private void Insert()
        {
            wSheet.Cells[11, 2].Value = docNumber;
            wSheet.Cells[11, 9].Value = date;
            wSheet.Cells[24, 1].Value = kindofwork;
            wSheet.Cells[26, 1].Value = designer;
            wSheet.Cells[29, 1].Value = material;
            wSheet.Cells[40, 1].Value = nextWork;
            //
            // заполнение описи в случае более 5 приложений к акту
            //
            if (fileToAdd.Count>5)
            {
                Excel.Application appInventory = null;
                Excel.Workbook wBookInventory = null;
                Excel.Worksheet wSheetInventory = null;
                wSheet.Cells[47, 1].Value = "в соответствием с листом описи";
                try
                {
                    appInventory = new Excel.Application();
                    wBookInventory = appInventory.Workbooks.Open(fileAppl);
                    wSheetInventory = (Excel.Worksheet)wBookInventory.Sheets[1];
                    wSheetInventory.Cells[7, 2].Value = "Опись передаваемых документов к акту " + docNumber;
                    int i = 1;
                    int n = 11;
                    string[] invent = documents.Split(new char[] {','}, StringSplitOptions.RemoveEmptyEntries);
                    foreach (string paper in invent)
                    {
                        if (paper != " ")
                        {
                            wSheetInventory.Cells[n, 2].Value = i;
                            wSheetInventory.Cells[n, 3].Value = paper;
                            wSheetInventory.Cells[n, 10].Value = "1 экз.";
                            Excel.Range line = (Excel.Range)wSheetInventory.Rows[n + 1];
                            line.Insert();
                            string column1Numb = "C" + n.ToString();
                            string column2Numb = "I" + n.ToString();
                            line = (Excel.Range)wSheetInventory.Range[column1Numb, column2Numb];
                            line.Merge();
                            i++;
                            n++;
                        }
                    }
                    string name = pathName + "\\" + "Акт" + docNumber + " опись";
                    wBookInventory.SaveAs(name + ".xls");
                    wBookInventory.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, name + ".pdf");
                }
                catch (Exception ex)
                {
                    Win.MessageBox.Show(ex.Message.ToString(), "Error happend in opening Inventory file: ");
                }
                finally
                {
                    wBookInventory.Close();
                    appInventory.Quit();
                }
            }
            else wSheet.Cells[47, 1].Value = documents;
        }

        private void DirectoryBtn_Click(object sender, Win.RoutedEventArgs e)
        {
            Forms.FolderBrowserDialog folder = new Forms.FolderBrowserDialog();
            folder.SelectedPath = toSaveDir;
            Forms.DialogResult res = folder.ShowDialog();
            pathName = folder.SelectedPath;
            if (pathName != "")
                LoadListOfWorks();
        }

        private void Save(string filetosave)
        {
            wBook.SaveAs(filetosave+".xls");
            wBook.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, filetosave+".pdf");
            MakePDF(filetosave + ".pdf", fileToAdd);
        }
    }
}
