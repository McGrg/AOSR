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
        private string docNumber, startdate, finishdate, florNumber, kindofwork, height, designer, material, nextWork, documents, pathName, subcontractor;
        private const string filenameSource= "D:\\Работа\\Интеллект Про\\Корпус 3\\АОСР\\Source.xlsx"; //расположение файла источника для полей формы
        private const string appdir = "D:\\Работа\\Интеллект Про\\Корпус 3\\АОСР\\Сертификаты"; // расположение папки с сертификатами и приложениями
        private const string toSaveDir = "D:\\Работа\\Интеллект Про\\Корпус 3\\АОСР\\";
        private const string listOfWorks = "D:\\Работа\\Интеллект Про\\Корпус 3\\АОСР\\Акты поэтажка без работ по АППОР.xlsx"; // расположение файла с данными по всем актам
        private const string filename = "D:\\Работа\\Интеллект Про\\Корпус 3\\АОСР\\Образец АОСР.xlsx"; // расположение файла образца для актов
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
                wSheetSource = (Excel.Worksheet)wBookSource.Sheets[1];// 1-этажи 2-26, 6 - лестница Н1
                app = new Excel.Application();
                wBook = app.Workbooks.Open(filename);
                wSheet = (Excel.Worksheet)wBook.Sheets[1];
                string temp = "test";
                for (int row =9; row<10; row++ )//счетчик по строкам временно на 1 строку
                { 
                    florNumber = wSheetSource.Cells[row, 1].Text;//колонка 1 - номер этажа
                    startdate = wSheetSource.Cells[row, 2].Text;
                    finishdate = wSheetSource.Cells[row, 3].Text;
                    subcontractor = wSheetSource.Cells[row, 109].Text;
                    int act = 0;
                    for (int column =4; column<55; column++)// колонки с данными 4-108 для этажей, 4-27 для лестницы Н1
                    {
                        temp = wSheetSource.Cells[row, column].Text;
                        if (temp!= "")
                        {
                            act++;
                            if (florNumber != "")
                            {
                                int res = 0;
                                if (int.TryParse(florNumber, out res))
                                height = HeightChoose(res);
                            }
                            else
                            {
                                height = null;
                            }
                            docNumber = "№ 3." + florNumber + "." + act.ToString();
                            string text="";
                            if (height != null)
                            {
                                text = " на отм. +" + height + ", ";
                            }
                            kindofwork = wSheetSource.Cells[7, column].Text + " " + wSheetSource.Cells[6, column].Text + " " + florNumber + " этаж," + text + axes;
                            designer = wSheetSource.Cells[4, column].Text;
                            material = wSheetSource.Cells[5, column].Text;
                            nextWork = wSheetSource.Cells[3, column].Text;
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
            wSheet.Cells[11, 9].Value = finishdate;
            wSheet.Cells[42, 5].Value = finishdate;
            wSheet.Cells[41, 5].Value = startdate;
            wSheet.Cells[31, 1].Value = kindofwork;
            wSheet.Cells[33, 1].Value = designer;
            int rowheight = 0;
            if (material.Length < 100) rowheight = 20;
            else if (material.Length < 300) rowheight = 50;
            else if (material.Length < 500) rowheight = 80;
            else rowheight = 120;
            wSheet.Cells[36, 1].RowHeight = rowheight;
            wSheet.Cells[36, 1].Value = material;
            wSheet.Cells[47, 1].Value = nextWork;
            //
            // заполнение описи в случае более 5 приложений к акту
            //
            if (fileToAdd.Count>5)
            {
                Excel.Application appInventory = null;
                Excel.Workbook wBookInventory = null;
                Excel.Worksheet wSheetInventory = null;
                wSheet.Cells[54, 1].RowHeight = 15;
                wSheet.Cells[54, 1].Value = "в соответствием с листом описи";
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
                            wSheetInventory.Cells[n, 3].WrapText = false;
                            wSheetInventory.Cells[n, 2].Value = i;
                            if (paper.Length>60)
                            {
                                wSheetInventory.Cells[n, 3].WrapText = true;
                                wSheetInventory.Cells[n, 3].RowHeight = 40;
                            }
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
            else wSheet.Cells[54, 1].Value = documents;
            if (subcontractor == "")
            {
                wSheet.Cells[25, 1].Value = "";
                wSheet.Cells[25, 1].RowHeight = 10;
                wSheet.Cells[26, 1].Value = "";
                wSheet.Cells[26, 1].RowHeight = 10;
                wSheet.Cells[71, 1].Value = "";
                wSheet.Cells[71, 1].RowHeight = 10;
                wSheet.Cells[72, 1].Value = "";
                wSheet.Cells[72, 1].RowHeight = 10;
                wSheet.Cells[73, 1].Value = "";
                wSheet.Cells[73, 1].RowHeight = 10;
            }
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

        //
        // высотная отметка проведения работ в зависимости от этажа
        //
        private string HeightChoose(int flor)
        {
            string height = "";
            if (flor < 4)
            {
                switch (flor)
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
            else
            {
                double res;
                res = (8.4 + (flor - 4) * 2.8);
                height = string.Format("{0:0.000}", res);
            }
            return height;
        }

        //
        //создание многостраничного PDF с приложениями, на входе исходный PDF файл и список файлов для приложения
        //
        private void MakePDF(string pdfFileName, List<string> fileArray)
        {

            try
            {
                // Open the output document
                PdfDocument outputDocument = PdfReader.Open(pdfFileName);

                // Iterate files
                foreach (string file in fileArray)
                {
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
    }
}
