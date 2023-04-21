using Microsoft.Office.Interop.Excel;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using Excel = Microsoft.Office.Interop.Excel;

namespace Template_4333
{
    /// <summary>
    /// Логика взаимодействия для _4333_Sagitova.xaml
    /// </summary>
    public partial class _4333_Sagitova : System.Windows.Window
    {
        public _4333_Sagitova()
        {
            InitializeComponent();
        }
        //импорт
        private void Button_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog()
            {
                DefaultExt = "*.xls;*.xlsx",
                Filter = "файл Excel (Spisok.xlsx)|*.xlsx",
                Title = "Выберите файл базы данных"
            };
            if (!(ofd.ShowDialog() == true))
                return;
            string[,] list;

            Excel.Application ObjWorkExcel = new
            Excel.Application();

            Excel.Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Open(ofd.FileName);
            Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[1];

            var lastCell = ObjWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);
            int _columns = (int)lastCell.Column;
            int _rows = (int)lastCell.Row;
            list = new string[_rows, _columns];
            for (int j = 0; j < _columns; j++)
            {
                for (int i = 0; i < _rows; i++)
                {
                    list[i, j] = ObjWorkSheet.Cells[i + 1, j + 1].Text;
                }
            }
            ObjWorkBook.Close(false, Type.Missing, Type.Missing);
            ObjWorkExcel.Quit();
            GC.Collect();

            using (isrpo2Entities usersEntities = new isrpo2Entities())
            {
                for (int i = 0; i < _rows; i++)
                {
                    usersEntities.Post_Work.Add(new Post_Work()
                    {
                        Id = list[i, 0],
                        Post = list[i, 1],
                        FIO = list[i, 2],
                        Login = list[i, 3],
                        Password = list[i, 4],
                        LastEntry = list[i, 5],
                        EntryType = list[i, 6]
                    });
                }
                usersEntities.SaveChanges();
            }
        }
        //экспорт
        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            List<Post_Work> allWorkers;
            using (isrpo2Entities usersEntities = new isrpo2Entities())
            {
                allWorkers = usersEntities.Post_Work.ToList();
            }

            var app = new Excel.Application();
            Excel.Workbook workbook = app.Workbooks.Add(Type.Missing);
            app.Visible = true;

            Excel.Worksheet worksheet1 = app.Worksheets.Add();
            worksheet1.Name = "Успешно";
            Excel.Worksheet worksheet2 = app.Worksheets.Add();
            worksheet2.Name = "Неуспешно";

            worksheet1.Cells[1, 1] = "Код клиента";
            worksheet1.Cells[1, 2] = "Должность";
            worksheet1.Cells[1, 3] = "Логин";

            worksheet2.Cells[1, 1] = "Код клиента";
            worksheet2.Cells[1, 2] = "Должность";
            worksheet2.Cells[1, 3] = "Логин";
            int rowIndex1 = 2;
            int rowIndex2 = 2;
            foreach (var workers in allWorkers)
            {
                if (workers.EntryType == "Успешно")
                {
                    worksheet1.Cells[rowIndex1, 1] = workers.Id;
                    worksheet1.Cells[rowIndex1, 2] = workers.Post;
                    worksheet1.Cells[rowIndex1, 3] = workers.Login;
                    rowIndex1++;
                }
                else if (workers.EntryType == "Неуспешно")
                {
                    worksheet2.Cells[rowIndex2, 1] = workers.Id;
                    worksheet2.Cells[rowIndex2, 2] = workers.Post;
                    worksheet2.Cells[rowIndex2, 3] = workers.Login;
                    rowIndex2++;
                }

            }

        }
    }

}