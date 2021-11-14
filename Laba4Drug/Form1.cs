using Google.Apis.Sheets.v4;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Application = Microsoft.Office.Interop.Excel.Application;
using System.IO;
using Google.Apis.Services;
using Google.Apis.Auth.OAuth2;
using Action = System.Action;
using System.Threading;
using System.Diagnostics;

namespace Laba4Drug
{
    public partial class Form1 : Form
    {
        public int[] startArray;
        public Form1()
        {
            InitializeComponent();
        }
        private async void FactorialAsync()
        {
            await Task.Run(() =>  startArray = BubbleSort(startArray));                // выполняется асинхронно
        }
        private void button1_Click(object sender, EventArgs e)
        {
            if (bubbleBox.Checked == true)
            {
/*                int[] bubbleArray = BubbleSort(startArray);*/
                FactorialAsync();
            }
            if (vstavBox.Checked == true)
            {
                int[] vstavArray = InsertionSort(startArray);
                chart2.Series[0].Points.DataBindXY(null, vstavArray);
            }
            if (shakeBox.Checked == true)
            {
                int[] shakeArray = ShakerSort(startArray);
                chart3.Series[0].Points.DataBindXY(null, shakeArray);
            }
            if (fastBox.Checked == true)
            {
                int[] fastArray = QuickSort(startArray);
                chart4.Series[0].Points.DataBindXY(null, fastArray);
            }
            if (bogoBox.Checked == true)
            {
                int[] bogoArray = BogoSort(startArray);
                chart5.Series[0].Points.DataBindXY(null, bogoArray);
            }
        }
        private void Form1_Load(object sender, EventArgs e)
        {

        }
        private void закрытьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Close();
        }
        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }
        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }
        private void button1_Click_1(object sender, EventArgs e)
        {
            dataGridView1.Rows.Clear();
            int m = Convert.ToInt32(textBox1.Text);
            Random rnd = new Random();
            startArray = new int[m];
            for (int j = 0; j < m; j++)
            {
                dataGridView1.Rows.Add();
                startArray[j] = rnd.Next(100);
                dataGridView1[0, j].Value = startArray[j];   
            }
        }
        //Пузырьковая сортировка
        private int[] BubbleSort(int[] array)
        {
            for (int i = 0; i < startArray.Count(); i++)
            {
                Action action2 = () => chart1.Series[0].Points.Add(startArray[i]);
                Invoke(action2);
            }
            Action action6 = () => chart1.Update();
            Invoke(action6);

            Stopwatch sw = new Stopwatch();
            sw.Start();
            for (int i = 1; i < array.Length; i++)
            {
                for (int j = 0; j < array.Length - 1; j++)
                {
                    if (array[j + 1] < array[j])
                    {
                        int n = array[j];
                        array[j] = array[j + 1];
                        array[j + 1] = n;

                        Action action3 = () => chart1.Series[0].Points.Clear();
                        Invoke(action3);
                        for (int k = 0; k < startArray.Count(); k++) 
                        { 
                            Action action4 = () => chart1.Series[0].Points.Add(startArray[k]); 
                            Invoke(action4); 
                        }
                        Thread.Sleep(50);
                    }
                }
            }
            sw.Stop();
            double aaa = sw.ElapsedMilliseconds / 1000;
            Action action5 = () => vrem1.Text = aaa.ToString();
            Invoke(action5);
            return array;
        }
        //Быстрая сортировка
        private void Swap(ref int value1, ref int value2)
        {
            int value3 = value1;
            value1 = value2;
            value2 = value3;
        }
        private int PivotIndex(int[] array, int minId, int maxId)
        {
            int index = minId - 1;
            for (int i = minId; i < maxId; i++)
            {
                if (array[i] < array[maxId])
                {
                    index++;
                    Swap(ref array[index], ref array[i]);
                }
            }
            index++;
            Swap(ref array[index], ref array[maxId]);
            return index;
        }
        private int[] QuickSort(int[] array, int minId, int maxId)
        {
            if (minId >= maxId)
                return array;
            int pivotPoint = PivotIndex(array, minId, maxId);
            QuickSort(array, minId, pivotPoint - 1);
            QuickSort(array, pivotPoint + 1, maxId);
            return array;
        }
        public int[] QuickSort(int[] array)
        {
            return QuickSort(array, 0, array.Length - 1);
        }
        //Сортировка вставками
        private int[] InsertionSort(int[] array)
        {
            for (int i = 1; i < array.Length; i++)
            {
                int key = array[i];
                int j = i;
                while ((j > 0) && (array[j - 1] > key))
                {
                    Swap(ref array[j - 1], ref array[j]);
                    j--;
                }

                array[j] = key;
            }
            return array;
        }
        //Шейкерная сортировка
        private int[] ShakerSort(int[] array)
        {
            for (int i = 0; i < array.Length / 2; i++)
            {
                var swapFlag = false;
                //проход слева направо
                for (int j = i; j < array.Length - i - 1; j++)
                {
                    if (array[j] > array[j + 1])
                    {
                        Swap(ref array[j], ref array[j + 1]);
                        swapFlag = true;
                    }
                }
                //проход справа налево
                for (int j = array.Length - 2 - i; j > i; j--)
                {
                    if (array[j - 1] > array[j])
                    {
                        Swap(ref array[j - 1], ref array[j]);
                        swapFlag = true;
                    }
                }
                //если обменов не было выходим
                if (!swapFlag)
                {
                    break;
                }
            }
            return array;
        }
        static bool IsSorted(int[] a)
        {
            for (int i = 0; i < a.Length - 1; i++)
            {
                if (a[i] > a[i + 1])
                    return false;
            }

            return true;
        }
        //перемешивание элементов массива
        private int[] RandomPermutation(int[] array)
        {
            Random random = new Random();
            var n = array.Length;
            while (n > 1)
            {
                n--;
                var i = random.Next(n + 1);
                var temp = array[i];
                array[i] = array[n];
                array[n] = temp;
            }
            return array;
        }
        //случайная сортировка
        private int[] BogoSort(int[] array)
        {
            while (!IsSorted(array))
            {
                array = RandomPermutation(array);
            }
            return array;
        }
        private void excelToolStripMenuItem_Click(object sender, EventArgs e)
        {
            openFileDialog1.FileName = String.Empty;
            DialogResult li = openFileDialog1.ShowDialog();
            if (li != DialogResult.OK) return;
            try
            {
                dataGridView1.Rows.Clear();
                Application ObjWorkExcel = new Application();
                Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Open(openFileDialog1.FileName, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing); //открыть файл
                Worksheet ObjWorkSheet = (Worksheet)ObjWorkBook.Sheets[1];
                var lastCell = ObjWorkSheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell);
                string sellx = String.Empty;
                string selly = String.Empty;
                for (int i = 0; i < lastCell.Row; i++)
                {
                    sellx = ObjWorkSheet.Cells[i + 1, 1].Text.ToString();
                    selly = ObjWorkSheet.Cells[i + 1, 2].Text.ToString();
                    if (sellx.Trim() != String.Empty && selly.Trim() != String.Empty)
                        dataGridView1.Rows.Add(sellx, selly);
                }
                ObjWorkBook.Close(false, Type.Missing, Type.Missing);
                ObjWorkExcel.Quit();
                GC.Collect();
            }
            catch (Exception exception)
            {
                MessageBox.Show("При попытке загрузки из Excel произошла обшика!", "Ошибка!");
            }
        }
        private static readonly string[] Scopes = { SheetsService.Scope.Spreadsheets };
        private const string SpreadsheetId = "1SHINt3UKLZ4p_tyTRnk7DSgmnyG-3Sk9lnxY73fKrhg";
        private const string GoogleCredentialsFileName = "laba4-332110-1ebf8c6441c1.json";
        private const string ReadRange = "Лист1!A:B";

        private static SheetsService GetSheetsService()
        {
            using (var stream = new FileStream(GoogleCredentialsFileName, FileMode.Open, FileAccess.Read))
            {
                var serviceInitializer = new BaseClientService.Initializer
                {
                    HttpClientInitializer = GoogleCredential.FromStream(stream).CreateScoped(Scopes)
                };
                return new SheetsService(serviceInitializer);
            }
        }
        async Task readAsy()
        {
            try
            {
                var serviceValues = GetSheetsService().Spreadsheets.Values;
                await ReadAsync(serviceValues);
            }
            catch (Exception)
            {
                MessageBox.Show("Количество ячеек Х не соответствует количеству Y. Не все данные будут занесены в таблицу!", "Предупреждение.");
            }
        }
        private async Task ReadAsync(SpreadsheetsResource.ValuesResource valuesResource)
        {
            var response = await valuesResource.Get(SpreadsheetId, ReadRange).ExecuteAsync();
            var values = response.Values;
            if (values == null || !values.Any())
            {
                MessageBox.Show("Документ пустой!");
                return;
            }
            var header = string.Join(" ", values.First().Select(r => r.ToString()));
            Console.WriteLine($"Header: {header}");

            List<string> baza = new List<string>();
            for (int i = 0; i < values.Count; i++)
            {
                string pern1 = values[i][0].ToString();

                baza.Add($"{pern1}");
                dataGridView1.Rows.Clear();

                foreach (string s in baza)
                {
                    var result = s.Split(';');
                    dataGridView1.Rows.Add(result[0]);
                }
            }
        }
        private async void googleToolStripMenuItem_Click(object sender, EventArgs e)
        {
            await readAsy();
        }

        private void vrem1_Click(object sender, EventArgs e)
        {

        }
    }
}
