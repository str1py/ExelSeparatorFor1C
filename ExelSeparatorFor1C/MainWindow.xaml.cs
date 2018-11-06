using Microsoft.Win32;
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
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Text.RegularExpressions;
//КЛОП ХОРОШИЙ!! olyashUtya olyashUtya
namespace ExelSeparatorFor1C
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        _Excel.Workbook wb;
        _Excel.Worksheet ws;
        string pathtodit = Directory.GetCurrentDirectory();
        _Application excel = new _Excel.Application();
        string path;

        List<string> names = new List<string>();
        List<string> namesNotNull = new List<string>();
        List<dynamic> numbers = new List<dynamic>();
        List<dynamic> length = new List<dynamic>();
        List<dynamic> width = new List<dynamic>();
        List<dynamic> count = new List<dynamic>();


        int _numbernext;
        int _lengthnext;
        int _widthnext;
        int _countnext;



        public MainWindow()
        {
            InitializeComponent();
        }

        private void SelectButton_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog myDialog = new OpenFileDialog();
            myDialog.Filter = "xls файлы(*.xls)|*.xls";
            myDialog.CheckFileExists = true;
            myDialog.Multiselect = true;
            if (myDialog.ShowDialog() == true)
            {
                FilePathTB.Text = myDialog.FileName;
                path = myDialog.FileName;
            }
        }

        private void  DevideButton_Click(object sender, RoutedEventArgs e)
        {
            CopyNames(path, 1);
        }

        protected void ClearLists()
        {
            names.Clear(); namesNotNull.Clear(); numbers.Clear();
            length.Clear(); width.Clear(); count.Clear();
        }

        public async void CopyNames(string filepath, int Sheet)
        {
            ClearLists(); // Очистка всех списков
            await AddToLog("Очистка списокв завершена!");

            try
            {
                wb = excel.Workbooks.Open(path);
                ws = wb.Worksheets[Sheet];
            }
            catch(Exception e)
            {
                await AddToLog(e.ToString());
            }

            int usedRowsNum = ws.UsedRange.Rows.Count;
            int usedColumsNum = ws.UsedRange.Columns.Count;

            int rowIndex = 12;
            int columnIndex = 1;

            try
            {
                for (int i = 1; i < usedRowsNum; i++)//Считывание всех значений из 1го столбца
                {
                    names.Add(ws.Cells[rowIndex, columnIndex].Value2);
                    rowIndex++;
                }
            }catch(Exception e)
            {
                await AddToLog(e.ToString());
            }

            try
            {
                for (int j = 0; j < names.Count; j++)//Удаление пустых
                {
                    if (names[j] != null)
                        namesNotNull.Add(names[j]);
                }
            }
            catch (Exception e)
            {
                await AddToLog(e.ToString());
            }

            namesNotNull.Remove("Заказчик___________________________");
            namesNotNull.Remove("Менеджер _______________________________");

            for (int q = 0; q < namesNotNull.Count; q++)//Вывод в лог
            {
                await AddToLog("Найден материал : " + namesNotNull[q]);
            }
            wb.Close();
            CreateNewXLS();
            CopyColums(path, 1);
            AddColomsToXLS();
        }

        public async void CreateNewXLS()
        {
            for (int i = 0; i < namesNotNull.Count; i++)
            {
                string str;
                str = Regex.Replace(namesNotNull[i], @"[^\w\.@-]", "", RegexOptions.None, TimeSpan.FromSeconds(1.5));//удаление неприемлевых знаков
                str = str.Replace(Environment.NewLine, string.Empty);
                str = str.Replace("\n", "br");

                this.wb = excel.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
                wb.SaveAs(pathtodit + @"\Worksbooks\" + str);

                await AddToLog("Создан " + str + ".xls");

            
                //namesNotNull[i] = str;
            }
            wb.Close();
        }

        public async void CopyColums(string filepath, int Sheet)
        {

            wb = excel.Workbooks.Open(path);
            ws = wb.Worksheets[Sheet];

            int usedRowsNum = ws.UsedRange.Rows.Count;
            int usedColumsNum = ws.UsedRange.Columns.Count;
            int rowIndex = 13;
            int columnIndex = 2;


            //ВЫТАСКИВАЕМ ПОЗИЦИИ
            try {
                for (int i = 1; i < usedRowsNum; i++) {
                    if (ws.Cells[rowIndex, columnIndex].Value2 == null && ws.Cells[rowIndex + 1, columnIndex].Value2 == null) {
                        await AddToLog("Позиция - значения успешно получены");
                        break;
                    }
                    else
                    {
                        var num = ws.Cells[rowIndex, columnIndex].Value2;
                        numbers.Add(num);
                        rowIndex++;
                    }
                }
            }
            catch (Exception e) {
                await AddToLog(e.ToString());
            }

            //ВЫТАСКИВАЕМ ДЛИННУ
            rowIndex = 13;
            columnIndex = 3;
            try {
                for (int i = 1; i < usedRowsNum; i++) {
                    if (ws.Cells[rowIndex, columnIndex].Value2 == null && ws.Cells[rowIndex + 1, columnIndex].Value2 == null)  {
                        await AddToLog("Длина - значения успешно получены");
                        break;
                    }
                    else {
                        var num = ws.Cells[rowIndex, columnIndex].Value2;
                        length.Add(num);
                        rowIndex++;
                    }
                }
            }
            catch (Exception e) {
                await AddToLog(e.ToString());
            }

            //ВЫТАСКИВАЕМ ШИРИНУ
            rowIndex = 13;
            columnIndex = 6;
            try {
                for (int i = 1; i < usedRowsNum; i++) {
                    if (ws.Cells[rowIndex, columnIndex].Value2 == null && ws.Cells[rowIndex + 1, columnIndex].Value2 == null) {
                        await AddToLog("Ширина - значения успешно получены");
                        break;
                    }
                    else {
                        var num = ws.Cells[rowIndex, columnIndex].Value2;
                        width.Add(num);
                        rowIndex++;
                    }
                }
            }
            catch (Exception e) {
                await AddToLog(e.ToString());
            }

            //ВЫТАСКИВАЕМ КОЛИЧЕСТВО
            rowIndex = 13;
            columnIndex = 9;
            try  {
                for (int i = 1; i < usedRowsNum; i++) {
                    if (ws.Cells[rowIndex, columnIndex].Value2 == null && ws.Cells[rowIndex + 1, columnIndex].Value2 == null)  {
                        await AddToLog("Количество - значения успешно получены");
                        break;
                    }
                    else  {
                        var num = ws.Cells[rowIndex, columnIndex].Value2;
                        count.Add(num);
                        rowIndex++;
                    }
                }
            }
            catch (Exception e)  {
                await AddToLog(e.ToString());
            }

  
            wb.Close();
        }

        public async void AddColomsToXLS()
        {      
            switch (namesNotNull.Count)
            {
                case 1:
                    break;
                case 2:
                    wb = excel.Workbooks.Open(pathtodit + @"\Worksbooks\" + namesNotNull[0] + ".xls");
                    ws = wb.Worksheets[1];

                    ///////ПЕРВАЯ ТАБЛИЦА/////////
                    //ПЕРЕНОС ПОЗИЦИЙ
                    int columnIndex = 1;
                    try  {
                        for (int i = 0; i < numbers.Count; i++)  {
                            if (numbers[i] == null)  {
                                _numbernext = i;
                                break;
                            }
                            else  {
                                ws.Cells[i + 1, columnIndex].Value2 = numbers[i];
                            }
                        }
                    }
                    catch (Exception e) {
                        await AddToLog(e.ToString());
                    }
                    //ПЕРЕНОС ДЛИННЫ
                    columnIndex = 2;
                    try  {
                        for (int i = 0; i < length.Count; i++) {
                            if (length[i] == null) {
                                _lengthnext = i;
                                break;
                            }
                            else   {
                                ws.Cells[i + 1, columnIndex].Value2 = length[i];
                            }
                        }
                    }
                    catch (Exception e)
                    {
                        await AddToLog(e.ToString());
                    }
                    //ПЕРЕНОС ШИРИНЫ
                    columnIndex = 3;
                    try {
                        for (int i = 0; i < width.Count; i++)  {
                            if (numbers[i] == null)  {
                                _widthnext = i;
                                break;
                            }
                            else   {
                                ws.Cells[i + 1, columnIndex].Value2 = width[i];
                            }
                        }
                    }
                    catch (Exception e)
                    {
                        await AddToLog(e.ToString());
                    }
                    //ПЕРЕНОС КОЛИЧЕСТВА
                    columnIndex = 4;
                    try {
                        for (int i = 0; i < count.Count; i++) {
                            if (count[i] == null){
                                _countnext = i;
                                break;
                            }
                            else {
                                ws.Cells[i + 1, columnIndex].Value2 = count[i];
                            }
                        }
                    }
                    catch (Exception e) {
                        await AddToLog(e.ToString());
                    }

                    wb.Close();


                    ///////ВТОРАЯ ТАБЛИЦА/////////
                    wb = excel.Workbooks.Open(pathtodit + @"\Worksbooks\" + namesNotNull[1] +".xls");
                    ws = wb.Worksheets[1];
                    //ПЕРЕНОС ПОЗИЦИЙ
                    columnIndex = 1;
                    try {
                        for (int i = _numbernext; i < numbers.Count; i++) {
                            if (numbers[i] == null) {
                                _numbernext = i;
                                break;
                            }
                            else {
                                ws.Cells[i + 1, columnIndex].Value2 = numbers[i];
                            }
                        }
                    }
                    catch (Exception e)
                    {
                        await AddToLog(e.ToString());
                    }
                    //ПЕРЕНОС ДЛИННЫ
                    columnIndex = 2;
                    try  {
                        for (int i = _lengthnext; i < length.Count; i++)  {
                            if (length[i] == null) {
                                _lengthnext = i;
                                break;
                            }
                            else {
                                ws.Cells[i + 1, columnIndex].Value2 = length[i];
                            }
                        }
                    }
                    catch (Exception e)
                    {
                        await AddToLog(e.ToString());
                    }
                    //ПЕРЕНОС ШИРИНЫ
                    columnIndex = 3;
                    try
                    {
                        for (int i = _widthnext; i < width.Count; i++) {
                            if (numbers[i] == null) {
                                _widthnext = i;
                                break;
                            }
                            else {
                                ws.Cells[i + 1, columnIndex].Value2 = width[i];
                            }
                        }
                    }
                    catch (Exception e)
                    {
                        await AddToLog(e.ToString());
                    }
                    //ПЕРЕНОС КОЛИЧЕСТВА
                    columnIndex = 4;
                    try  {
                        for (int i = _countnext; i < count.Count; i++)  {
                            if (count[i] == null){
                                _countnext = i;
                                break;
                            }
                            else  {
                                ws.Cells[i + 1, columnIndex].Value2 = count[i];
                            }
                        }
                    }
                    catch (Exception e)
                    {
                        await AddToLog(e.ToString());
                    }
                    break;
                case 3:
                    break;
                case 4:
                    break;
                case 5:
                    break;
                case 6:
                    break;
                case 7:
                    break;
                case 8:
                    break;
            }




            wb = excel.Workbooks.Open(path);
            ws = wb.Worksheets[1];

  


        }

        async Task AddToLog(string message)
        {
            await LogListBox.Dispatcher.BeginInvoke(new System.Action(delegate ()
            {
                LogListBox.Items.Add(message);
            }));
        }

        private void ExcelClose()
        {
            wb.Close(false);
            excel.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excel);
            excel = null;
            wb = null;
            ws = null;
            System.GC.Collect();
        }

    }
}
