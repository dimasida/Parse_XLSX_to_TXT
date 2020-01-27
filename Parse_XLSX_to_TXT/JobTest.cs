//Тестовое задание: форматированный вывод данных из xlsx в txt

using System;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelParseToTxt
{
    //Структура для вывода строки
    internal struct DataRow
    {
        public int Code;
        public string Articul;
        public string NameDevice;
        public string NameManufacturer;
        public string UnitDevice;
        public string Price;
    }

    /* Реализация связного списка
	 * 
	 * 
	 * public class Node<T>
	{
		public Node(T DataRow)
		{
			Data = DataRow;
		}
		public T Data { get; set; }
		public Node<T> Next { get; set; }
	}

	public class LinkedList<T> : IEnumerable<T>  // односвязный список
	{
		Node<T> head; // головной/первый элемент
		Node<T> tail; // последний/хвостовой элемент

		// добавление элемента
		public void Add(T data)
		{
			Node<T> node = new Node<T>(data);

			if (head == null)
				head = node;
			else
				tail.Next = node;
			tail = node;
		}

		// реализация интерфейса IEnumerable для foreach
		IEnumerator IEnumerable.GetEnumerator()
		{
			return ((IEnumerable)this).GetEnumerator();
		}

		IEnumerator<T> IEnumerable<T>.GetEnumerator()
		{
			Node<T> current = head;
			while (current != null)
			{
				yield return current.Data;
				current = current.Next;
			}
		}
	}
	*/

    internal class JobTest
    {

        //Метод добавления данных в структуру
        public static DataRow AddData(int LocNumber, int LocRows, Excel.Worksheet mySheet)
        {
            DataRow myDataRow;
            Excel.Range LocalRange;

            myDataRow.Code = LocNumber;

            LocalRange = (Excel.Range)mySheet.Cells[LocRows, 2];
            if ((LocalRange.Value2 != null) && (LocalRange.Value2.ToString() != ""))
            {
                myDataRow.Articul = LocalRange.Value2.ToString();
            }
            else
            {
                myDataRow.Articul = "Нет";
            }

            LocalRange = (Excel.Range)mySheet.Cells[LocRows, 3];
            myDataRow.NameDevice = LocalRange.Value2.ToString();

            LocalRange = (Excel.Range)mySheet.Cells[LocRows, 4];
            if ((LocalRange.Value2 != null) && (LocalRange.Value2.ToString() != ""))
            {
                myDataRow.NameManufacturer = LocalRange.Value2.ToString();
            }
            else
            {
                myDataRow.NameManufacturer = "Нет";
            }

            LocalRange = (Excel.Range)mySheet.Cells[LocRows, 5];
            myDataRow.UnitDevice = LocalRange.Value2.ToString();

            LocalRange = (Excel.Range)mySheet.Cells[LocRows, 6];
            myDataRow.Price = LocalRange.Value2.ToString();

            return myDataRow;
        }

        public static string GetFileExtension(string fileName)
        {
            return fileName.Substring(fileName.LastIndexOf(".") + 1);
        }

        private static void Main()
        {
            Excel.Application ExcelApp;
            Excel.Workbook ExcelWorkbook;
            Excel.Worksheet ExcelWorksheet;

            string pathXls = "";
            string pathTxt = "";
            Boolean PathBool = true;

            Console.WriteLine("Программа начала работать.");

            do
            {
                Console.Write("Введите верный путь для файла.xlsx: ");
                pathXls = Console.ReadLine();

                if ((pathXls.EndsWith(".xlsx")) && (File.Exists(pathXls)))
                {
                    PathBool = false;
                }
            } while (PathBool);

            Console.WriteLine();
            PathBool = true;

            do
            {
                Console.Write("Введите верный путь для файла.txt: ");
                pathTxt = Console.ReadLine();

                if ((pathTxt.EndsWith(".txt")) && (File.Exists(pathTxt)))
                {
                    PathBool = false;
                }
            } while (PathBool);

            Console.WriteLine("Были указаны файлы с верными расширениями. Теперь подождите, программа обрабатывает информацию.");
            //pathXls = @"C:\\Users\\Alya\\source\\repos\\Parse_XLSX_to_TXT\\Parse_XLSX_to_TXT\\lib\\Price_Kompjuternaja_perifеrija_2018_07_10.xlsx";
            //pathTxt = @"C:\\Users\\Alya\\source\\repos\\Parse_XLSX_to_TXT\\Parse_XLSX_to_TXT\\lib\\Parse.txt";

            ExcelApp = new Excel.Application();
            if (ExcelApp == null)
            {
                Console.WriteLine("Excel is not Insatalled.");

            }
            else
            {
                ExcelApp.Workbooks.Open(pathXls);
                ExcelWorkbook = ExcelApp.ActiveWorkbook;
                ExcelWorksheet = (Excel.Worksheet)ExcelWorkbook.Worksheets[1];

                int Rows = 0;
                int Number = 0;
                Excel.Range Range;

                StreamWriter TxtWriter = new StreamWriter(pathTxt, false, System.Text.Encoding.Default);

                do
                {
                    string Temp;
                    Rows++;

                    Range = (Excel.Range)ExcelWorksheet.Cells[Rows, 1];

                    if (Range.Value2 != null)
                    {
                        Temp = Range.Value2.ToString();
                        Int32.TryParse(Temp, out Number);

                        if (Number > 0)
                        {
                            DataRow dr;
                            //Вызов метода добавления данных
                            dr = AddData(Number, Rows, ExcelWorksheet);

                            //Форматированной запись в txt
                            TxtWriter.WriteLine($"Код: {dr.Code}, Артикул: {dr.Articul}, Наименование: {dr.NameDevice}, " +
                                $"Производитель: {dr.NameManufacturer}, Единица измерения: {dr.UnitDevice}, " +
                                    "Цена: {0:N}р.", dr.Price);
                        }
                    }
                }
                while (Rows != 911);

                //Закрытие файлов
                TxtWriter.Close();
                ExcelWorkbook.Close(false);
                ExcelApp.Quit();

                //Обнуляю
                ExcelWorksheet = null;
                ExcelWorkbook = null;
                ExcelApp = null;

                Console.WriteLine("Конец программы.");
                Console.ReadKey();

                GC.Collect();
            }
        }
    }
}