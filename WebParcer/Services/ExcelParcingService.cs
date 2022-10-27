using Microsoft.AspNetCore.Hosting;
using Microsoft.EntityFrameworkCore;
using Microsoft.IdentityModel.Tokens;
using System.ComponentModel.DataAnnotations;
using System.Runtime.InteropServices;
using WebParcer.DBContext;
using WebParcer.Models.TableModels;
using Excel = Microsoft.Office.Interop.Excel;

namespace WebParcer.Services
{
    public class ExcelParcingService
    {
        private IWebHostEnvironment _webHostEnvironment;
        private ApplicationDBContext _dBContext;

        public ExcelParcingService(IWebHostEnvironment webHostEnvironment, ApplicationDBContext dBContext)
        {
            _webHostEnvironment = webHostEnvironment;
            _dBContext = dBContext;
        }

        public void Parce(string filename)
        {
            string filePath = Path.Combine(_webHostEnvironment.WebRootPath + "\\Files\\" + filename);
            object rOnly = true;
            object SaveChanges = false;
            object MissingObj = System.Reflection.Missing.Value;

            Excel.Application app = new Excel.Application();
            Excel.Workbooks workbooks = app.Workbooks;
            Excel.Workbook workbook = workbooks.Open(filePath, MissingObj, rOnly, MissingObj, MissingObj,
                                        MissingObj, MissingObj, MissingObj, MissingObj, MissingObj,
                                        MissingObj, MissingObj, MissingObj, MissingObj, MissingObj);

            Excel.Sheets sheets = workbook.Sheets;            // Получение всех страниц докуента

            foreach (Excel.Worksheet worksheet in sheets)
            {
                bool firstTime = true;

                Excel.Range UsedRange = worksheet.UsedRange;    // Получаем диапазон используемых на странице ячеек

                Excel.Range urRows = UsedRange.Rows;            // Получаем строки в используемом диапазоне

                Excel.Range urColums = UsedRange.Columns;       // Получаем столбцы в используемом диапазоне

                int RowsCount = urRows.Count;
                int ColumnsCount = urColums.Count;


                List<double> listOfNumbers = new List<double>(7);
                int counter = 0;

                List<List<double>> listOfRecords = new List<List<double>>();

                for (int i = 1; i <= RowsCount; i++)        //Перебор каждой ячейки в таблице
                {
                    for (int j = 1; j <= ColumnsCount; j++)
                    {
                        Excel.Range CellRange = UsedRange.Cells[i, j];
                        string? CellText = (CellRange == null || CellRange.Value2 == null) ? null :
                                            (CellRange as Excel.Range).Value2.ToString();          //Получение содержимого ячейки

                        if (CellText is null)
                            continue;
                        double number = 0;
                        if (!Double.TryParse(CellText, out number))     
                        {
                            if (CellText.Contains("ПО КЛАССУ") || CellText.Contains("БАЛАНС"))
                            {
                                j += 7;      //Если конец таблицы класса или всей таблицы то не читаем сумарные значения
                                counter++;   //Подсчёт номера таблицы
                                if (!CellText.Contains("БАЛАНС"))
                                {
                                    WriteToDb(listOfRecords, counter);      //Запись в базу данных
                                    listOfRecords.Clear();
                                }
                            }       
                        }
                        else
                        {
                            if (firstTime)      //Пропускаем дату
                            {
                                firstTime = false;
                                continue;
                            }
                            if (listOfNumbers.Count == 7)
                            {
                                listOfRecords.Add(new List<double>(listOfNumbers));     //Заполняем коллекцию данными из одной строки
                                listOfNumbers.Clear();
                            }
                            listOfNumbers.Add(number);
                        }
                    }
                }
                _dBContext.SaveChanges();   //Сохранение изменений в базе данных

                //Освобождение ресурсов
                if (urRows != null) Marshal.ReleaseComObject(urRows);
                if (urColums != null) Marshal.ReleaseComObject(urColums);
                if (UsedRange != null) Marshal.ReleaseComObject(UsedRange);
                if (worksheet != null) Marshal.ReleaseComObject(worksheet);
            };
        }
        private void WriteToDb(List<List<double>> records, int counter)     //Функция добавление в бд
        {
            //Исходя из номера обработанной таблицы коллекция строк приводится к типу нужной таблицы и добавляется в базу данных
            switch (counter)
            {
                case 1:
                    {
                        var listOfClass = Create<Class1>(records);
                        Save(_dBContext.Class1s, listOfClass);
                        break;
                    }
                case 2:
                    {
                        var listOfClass = Create<Class2>(records);
                        Save(_dBContext.Class2s, listOfClass);
                        break;
                    }
                case 3:
                    {
                        var listOfClass = Create<Class3>(records);
                        Save(_dBContext.Class3s, listOfClass);
                        break;
                    }
                case 4:
                    {
                        var listOfClass = Create<Class4>(records);
                        Save(_dBContext.Class4s, listOfClass);
                        break;
                    }
                case 5:
                    {
                        var listOfClass = Create<Class5>(records);
                        Save(_dBContext.Class5s, listOfClass);
                        break;
                    }
                case 6:
                    {
                        var listOfClass = Create<Class6>(records);
                        Save(_dBContext.Class6s, listOfClass);
                        break;
                    }
                case 7:
                    {
                        var listOfClass = Create<Class7>(records);
                        Save(_dBContext.Class7s, listOfClass);
                        break;
                    }
                case 8:
                    {
                        var listOfClass = Create<Class8>(records);
                        Save(_dBContext.Class8s, listOfClass);
                        break;
                    }
                case 9:
                    {
                        var listOfClass = Create<Class9>(records);
                        Save(_dBContext.Class9s, listOfClass);
                        break;
                    }
                default:
                    throw new Exception();
            }
        }
        private List<T> Create<T>(List<List<double>> records) where T : TableModelBase, new()       //Приведение коллекции строк к коллекции типа определённой таблицы
        {
            List<T> recordsToAdd = new List<T>();
            foreach (var record in records)
            {
                T classItem = new T()
                {
                    B_sch = (int)record[0],
                    InBalanceActive = record[1],
                    InBalancePassive = record[2],
                    Debit = record[3],
                    Credit = record[4],
                    OutBalanceActive = record[5],
                    OutBalancePassive = record[6],
                };
                recordsToAdd.Add(classItem);
            }

            return recordsToAdd;
        }
        private void Save<T>(DbSet<T> database, List<T> obj) where T : class       //Добавление коллекции данных в базу данных
        {
            foreach (var record in obj)
            {
                database.AddRange(record);
            }
        }
    }
}
