using System;
using Aspose.Cells;

// Учебная программа для расчёта пробного давления по ГОСТ 34347-2017

namespace ConsoleApp2
{
    internal class Program
    {

        private static int sigmaMax; //Допускаемое напряжение при максимальной или равной расчётной температуре
        private static int sigmaMin; //Допускаемое напряжение при минимальной расчётной температуре
        private static int sigma20; //Допускаемое напряжение при 20 градусов
        private static double sigma; // Допускаемое напряжение при расчётной температуре
        private static double temperatyre; //расчётная температура
        private static double desingPressure; //расчётное давление
        private static double testPressure; //пробное давление
        private static double ratio; //коэффициент (при конкретных случаях смотрим ГОСТ 34233.1 примечания в Приложениях
        private static int material; //номер материала
        private static bool interpolation = false; //интерполяция
        private static string condition;

        static void Main(string[] args)
        {

            // Загрузить файл Excel
            Workbook wb = new Workbook("test.xlsx");

            // Получить рабочий лист, используя его индекс
            Worksheet worksheet = wb.Worksheets[0];

            // Печать имени рабочего листа
            Console.WriteLine("Worksheet: " + worksheet.Name);

            // Определяем кол-во заполненых столбцов
            int cols = worksheet.Cells.MaxDataColumn;

            // Определяем кол-во заполненых столбцов
            int rows = worksheet.Cells.MaxDataRow;

            //Выводим список доступных материалов
            for (int j = 0; j < rows; j++)
            {
                Console.WriteLine("Материал: {0}, номер {1}", worksheet.Cells[j, 0].Value, worksheet.Cells[j, 0].Row);
            }

            //Выбор материала

            //Обрабатываем с помощью цикла возможные ошибки (ввод различных символов, кроме чисел)
            do
            {
                Console.WriteLine("Введите номер материала");
            }
            //Вводим материал
            //var material = Convert.ToInt16(Console.ReadLine());
            while (!int.TryParse(Console.ReadLine(), out material));

            Console.WriteLine("Выбран материал " + worksheet.Cells[material, 0].Value);

            //Указываем температуру

            //Обрабатываем с помощью цикла возможные ошибки (ввод различных символов, кроме чисел)
            do
            {
                Console.WriteLine("Введите температуру");
            }
            //Вводим температуру
            //temperatyre = Convert.ToDouble(Console.ReadLine());
            while (!double.TryParse(Console.ReadLine(), out temperatyre));

            //Указываем расчётное давление

            //Обрабатываем с помощью цикла возможные ошибки (ввод различных символов, кроме чисел)
            do
            {
                Console.WriteLine("Введите расчётное давление");
            }
            //Вводим расчётную температуру
            while (!double.TryParse(Console.ReadLine(), out desingPressure));

            try
            {
                // Поиск столбца по температуре
                int i = 0; //Ищем температуру по первой строке
            for (int j = 0; j < cols; j++)
            {
                //Находим равную или большую температуру по базе
                if (Convert.ToInt16(worksheet.Cells[i, j].Value) >= temperatyre)
                {
                    sigmaMax = worksheet.Cells[i, j].Column;
                    // Если температуры нет в базе включаем интерполяцию 
                    if (Convert.ToInt16(worksheet.Cells[i, sigmaMax].Value) != temperatyre && temperatyre >=20)
                    {
                        interpolation = true;
                    }
                    break;
                }
            }
                // Интерполяция
                if (interpolation == true)
                {
                    // Условие, если температура ниже 20
                    if (temperatyre >= 20)
                    {
                        sigmaMin = sigmaMax - 1;
                    }
                    else
                    {
                        // Если температура меньше 20
                        sigmaMin = Convert.ToInt16(worksheet.Cells[i, 2].Column);
                    }

                    Console.WriteLine("Максимальная температура X1 " + worksheet.Cells[i, sigmaMax].Value);
                    Console.WriteLine("Минимальная температура X2 " + worksheet.Cells[i, sigmaMin].Value);
                    Console.WriteLine("Максимальная температура Y1 " + worksheet.Cells[material, sigmaMax].Value);
                    Console.WriteLine("Минимальная температура Y2 " + worksheet.Cells[material, sigmaMin].Value);
                    var X1 = Convert.ToInt16(worksheet.Cells[i, sigmaMax].Value);
                    var X2 = Convert.ToInt16(worksheet.Cells[i, sigmaMin].Value);
                    var Y1 = Convert.ToInt16(worksheet.Cells[material, sigmaMax].Value);
                    var Y2 = Convert.ToInt16(worksheet.Cells[material, sigmaMin].Value);
                    var X = temperatyre;
                    var Y = Y1 + (Y2 - Y1) * ((X - X1) / (X2 - X1));
                    sigma = Math.Round(Y,2);
                }
                else
                {
                    Console.WriteLine("Не включаем интерполяцию ({0})", interpolation);
                    sigma = Convert.ToInt16(worksheet.Cells[material, sigmaMax].Value);
                }
                Console.WriteLine("Значение допускаемого напряжения: {0} МПа", sigma);

                // Вычисляем допускаемое напряжение при 20 градусов

                sigma20 = Convert.ToInt16(worksheet.Cells[material, 1].Value);

                //Вводим коэффициент 

                Console.WriteLine("Нужен дополнительный коээфициент? ( + )");

                condition = Convert.ToString(Console.ReadLine());

                if (condition=="+")
                {
                    //Обрабатываем с помощью цикла возможные ошибки (ввод различных символов, кроме чисел)
                    do
                    {
                        Console.WriteLine("Введите коэффициент");
                    }
                    //Вводим коэффициент

                    while (!double.TryParse(Console.ReadLine(), out ratio));
                }
                else
                {
                    ratio = 1;
                }

                // Вычисляем пробное гидравлическое давление
                testPressure = Math.Round(1.25 * desingPressure * (sigma20 * ratio / sigma * ratio),2);

                if (condition == "+")
                {
                    Console.WriteLine("Pпр= 1,25 * {0} * ( ({1} * {3}) / ({2} * {3}) )", desingPressure, sigma20, sigma, ratio);
                }
                else
                {
                    Console.WriteLine("Pпр= 1,25 * {0} * ( {1} / {2} )", desingPressure, sigma20, sigma);
                }

                Console.WriteLine("Пробное давление: {0} МПа", testPressure);
            }
            catch
            {
                Console.WriteLine("В базе нет такой температуры");
            }
            Console.ReadLine();
        }
    }
}