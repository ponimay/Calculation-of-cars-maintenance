using NServiceBus.Unicast;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing.Design;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Telegram.Bot;
using Telegram.Bot.Args;
using Telegram.Bot.Types;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using Application = Microsoft.Office.Interop.Excel.Application;                 
using Microsoft.Office.Interop.Word;
using Word = Microsoft.Office.Interop.Word;
using Application1 = Microsoft.Office.Interop.Word.Application;
using System.Security.AccessControl;
using Document = Microsoft.Office.Interop.Word.Document;
using static System.Windows.Forms.DataFormats;
using System.Globalization;
using static System.Net.Mime.MediaTypeNames;
using Range = Microsoft.Office.Interop.Excel.Range;

namespace WinFormsApp4
{
    public partial class Form1 : Form
    {
        private double coefWinter = 20; // коофицент по зиме
        private double coefSpring = 25; // коофицент по весне
        private double coefSummer = 30; // коофицент по лету
        private double coefAutumn = 22; // коофицент по осени
        private const double Probeg = 15000; // максимальный пробег
        
        public Form1()
        {
            InitializeComponent();
        }
        
        
        TelegramBotClient bot = new TelegramBotClient("5313774371:AAHfKFg1ylcdLT-PiF8MyBXLJG4K9i2mAbM");

        private void button1_Click(object sender, EventArgs e)
        {
            
            // Создание объекта Word
            Application1 Word = new Word.Application();
            Word.Document doc = null;
            doc = Word.Documents.Open(@"E:\User\Мой Excel файл.docx"); // initialize the doc object
                                                                       // Открытие файла Word

            string text = doc.Content.Text.Split(' ')[0];// индекс пробега
            string text1 = doc.Content.Text.Split('@')[1];// индекс даты
                                                           //после зациклить индексы

            DateTime date = DateTime.MinValue; // инициализация переменной даты
            string[] words = text1.Split(' '); // разделение текста на слова
            foreach (string word1 in words)
            {
                if (DateTime.TryParseExact(word1, "dd.MM.yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out date))
                {
                    // если удалось успешно преобразовать строку в дату, выходим из цикла
                    MessageBox.Show (date.ToString()); //  выводит дату
                    break;
                }
            }
            doc.Close(true);
           
            Excel.Application excel = new Excel.Application();
            Excel.Workbook workbook = excel.Workbooks.Open(@"E:\User\Мой Excel файл.xlsx");
            Excel.Worksheet worksheet = workbook.Sheets[1];
            worksheet.Name = "Данные";
            worksheet.Cells[1, 1] = "Дата";
            worksheet.Cells[1, 2] = "Средний пробег за сезон";
            worksheet.Cells[1, 3] = "Сезон";
            worksheet.Cells[2, 3] = "Зима";
            worksheet.Cells[3, 3] = "Весна";
            worksheet.Cells[4, 3] = "Лето";
            worksheet.Cells[5, 3] = "Осень";
            worksheet.Cells[1, 4] = "Максимальный пробег за сезон";
            worksheet.Cells[1, 5] = "Средний пробег по сезонам";
            workbook.Save();
            double currentMileage = double.Parse(text); // пробег
            double coef = date.Month; // коофицент
            workbook.Save();
            
            if (text == "")
            {
                MessageBox.Show("Напишите количеств пройденных километров.");
                return;
            }
            double winterMileage = 0, springMileage = 0, summerMileage = 0, autumnMileage = 0; // по пробегу месяца
            int winterDays = 0, springDays = 0, summerDays = 0, autumnDays = 0; // по дням

            // Заполняем ячейки A2 и B2 начальными значениями
            worksheet.Cells[2, 1] = date.ToShortDateString();
            worksheet.Cells[2, 2] = currentMileage.ToString();
            int row = 2; // номер строки для заполнения

            while (currentMileage < Probeg)
            {
                double distance = coef * 31; // 31 - коофицент каждого цесяца
                currentMileage += distance;

                worksheet.Cells[row, 1] = date.ToShortDateString();
                worksheet.Cells[row, 2] = currentMileage.ToString();
                row++;

                switch (date.Month)
                {
                    case 12:
                    case 1:
                    case 2:
                        coef = coefWinter;
                        winterMileage += distance;
                        winterDays += 31;
                        break;
                    case 3:
                    case 4:
                    case 5:
                        coef = coefSpring;
                        springMileage += distance;
                        springDays += 31;
                        break;
                    case 6:
                    case 7:
                    case 8:
                        coef = coefSummer;
                        summerMileage += distance;
                        summerDays += 31;
                        break;
                    case 9:
                    case 10:
                    case 11:
                        coef = coefAutumn;
                        autumnMileage += distance;
                        autumnDays += 31;
                        break;
                    default:
                        break;
                }

                date = date.AddMonths(1);

            }
            double winterAverage = winterMileage / winterDays;
            double winterMax;
            if (currentMileage - Probeg < winterAverage * 31)
            {
                winterMax = currentMileage - Probeg;
            }
            else
            {
                winterMax = winterAverage * 31;
            }

            double springAverage = springMileage / springDays;
            double springMax;
            if (currentMileage - Probeg < springAverage * 31)
            {
                springMax = currentMileage - Probeg;
            }
            else
            {
                springMax = springAverage * 31;
            }

            double summerAverage = summerMileage / summerDays;
            double summerMax;
            if (currentMileage - Probeg < summerAverage * 31)
            {
                summerMax = currentMileage - Probeg;
            }
            else
            {
                summerMax = summerAverage * 31;
            }

            double autumnAverage = autumnMileage / autumnDays;
            double autumnMax;
            if (currentMileage - Probeg < autumnAverage * 31)
            {
                autumnMax = currentMileage - Probeg;
            }
            else
            {
                autumnMax = autumnAverage * 31;
            }
            int columnCount = worksheet.UsedRange.Columns.Count;
            MessageBox.Show(columnCount.ToString());

            worksheet.Cells[2, 4] = winterAverage;
            worksheet.Cells[2, 5] = winterMax;

            worksheet.Cells[3, 4] = springAverage;
            worksheet.Cells[3, 5] = springMax;

            worksheet.Cells[4,4] = summerAverage;
            worksheet.Cells[4,5] = summerMax;

            worksheet.Cells[5,4] = autumnAverage;
            worksheet.Cells[5,5] = autumnMax;


            //double lastValue = double.Parse(worksheet.Rows[worksheet.Rows.Count - 2].Cells[1].Value.ToString());
            //double prevValue = double.Parse(worksheet.Rows[worksheet.Rows.Count - 3].Cells[1].Value.ToString());
            //double difference = lastValue - prevValue;

            double lastValue = double.Parse(worksheet.Cells[worksheet.Rows.Count, 2].Value.ToString());
            double prevValue = double.Parse(worksheet.Cells[worksheet.Rows.Count - 1, 2].Value.ToString());
            double difference = lastValue - prevValue;

            MessageBox.Show(difference.ToString());
            excel.Visible = true;
            string path = @"E:\User\Мой Excel файл.xlsx";
            workbook.SaveAs(path);
            MessageBox.Show("Файл сохранён!", "Сохранение файла", MessageBoxButtons.OK, MessageBoxIcon.Information);
            workbook.Close(true);
            excel.Quit();

            if (currentMileage > Probeg)
            {
                string message = $"Внимание! Пора сделать тех. обслуживание машины! Пробег превысил {Probeg}, за последнее время было пройдено {600} км.";
                bot.SendTextMessageAsync("1083342768", message);

            }
            else
            {
                string message = $"Скоро нужно будет сделать техническое обслуживание машины! Пробег машины достигает {Probeg}, за последний месяц пройдено {600} км.";
                bot.SendTextMessageAsync("1083342768", message);

            }
        }
    }

}
