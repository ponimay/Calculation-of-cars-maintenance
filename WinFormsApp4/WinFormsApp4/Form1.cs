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

namespace WinFormsApp4
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        private double coefWinter = 20; // коофицент по зиме
        private double coefSpring = 25; // коофицент по весне
        private double coefSummer = 30; // коофицент по лету
        private double coefAutumn = 22; // коофицент по осени
        private const double Probeg = 15000; // максимальный пробег

        private void ExportToExcel()
        {

            Application excelApp = new Application();
            // новая книга Excel
            Workbook excelWorkbook = excelApp.Workbooks.Add();
            // новый лист
            Worksheet excelWorksheet = excelWorkbook.Sheets.Add();
            // название листа
            excelWorksheet.Name = "Данные";

            // Копируем данные из dataGridView1 в Excel
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                for (int j = 0; j < dataGridView1.Columns.Count; j++)
                {
                    excelWorksheet.Cells[i+1 , j+1] = dataGridView1.Rows[i].Cells[j].Value;
                }
            }
            excelApp.Visible = false;
            excelApp.UserControl = true;
            string path = (@"E:\User\");
            string filename = "Мой Excel файл.xlsx";
            string fullfilename = System.IO.Path.Combine(path, filename);
            if (System.IO.File.Exists(fullfilename)) System.IO.File.Delete(fullfilename);
            excelWorkbook.SaveAs(fullfilename, Excel.XlFileFormat.xlWorkbookDefault); //формат Excel 2007

            excelWorkbook.Close(false); //false - закрыть рабочую книгу не сохраняя изменения
            excelApp.Quit(); 
            MessageBox.Show("Файл сохранён!", "Сохранение файла", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        TelegramBotClient bot = new TelegramBotClient("5313774371:AAHfKFg1ylcdLT-PiF8MyBXLJG4K9i2mAbM");
        private void button1_Click(object sender, EventArgs e)
        {
            dataGridView1.Rows.Clear();
            dataGridView1.Columns.Clear();

            dataGridView1.Columns.Add("Дата", "Дата");
            dataGridView1.Columns.Add("Средний пробег за сезон", "Средний пробег за сезон");
            dataGridView1.Columns.Add("Сезон", "Сезон");
            dataGridView1.Columns.Add("Максимальный пробег за сезон", "Максимальный пробег за сезон");
            dataGridView1.Columns.Add("Средний пробег по сезонам", "Средний пробег по сезонам");
            
            if(textBox1.Text == "")
            {
                MessageBox.Show("Напишите количеств пройденных киллометров.");
                return;
            }
            DateTime startDate = dateTimePicker1.Value; // выбрать дату
            double currentMileage = double.Parse(textBox1.Text); // пробег
            DateTime selectedDate = dateTimePicker1.Value;
            double coef = selectedDate.Month; // коофицент


            double winterMileage = 0, springMileage = 0, summerMileage = 0, autumnMileage = 0; // по пробегу месяца
            int winterDays = 0, springDays = 0, summerDays = 0, autumnDays = 0; // по дням

            while (currentMileage < Probeg)
            {
                double distance = coef * 31; // 31 - коофицент каждого цесяца
                currentMileage += distance;

                dataGridView1.Rows.Add(startDate.ToShortDateString(), currentMileage.ToString());

                switch (startDate.Month)
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

                startDate = startDate.AddMonths(1);

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

            dataGridView1.Rows[0].Cells[2].Value = "Зима".ToString();
            dataGridView1.Rows[1].Cells[2].Value = "Весна".ToString();
            dataGridView1.Rows[2].Cells[2].Value = "Лето".ToString();
            dataGridView1.Rows[3].Cells[2].Value = "осень".ToString();
            
            dataGridView1.Rows[0].Cells[3].Value = winterAverage.ToString();
            dataGridView1.Rows[0].Cells[4].Value = winterMax.ToString();

            dataGridView1.Rows[1].Cells[3].Value = springAverage.ToString();
            dataGridView1.Rows[1].Cells[4].Value = springMax.ToString();

            dataGridView1.Rows[2].Cells[3].Value = summerAverage.ToString();
            dataGridView1.Rows[2].Cells[4].Value = summerMax.ToString();

            dataGridView1.Rows[3].Cells[3].Value = autumnAverage.ToString();
            dataGridView1.Rows[3].Cells[4].Value = autumnMax.ToString();
            

            double lastValue = double.Parse(dataGridView1.Rows[dataGridView1.Rows.Count - 2].Cells[1].Value.ToString());
            double prevValue = double.Parse(dataGridView1.Rows[dataGridView1.Rows.Count - 3].Cells[1].Value.ToString());
            double difference = lastValue - prevValue;
            ExportToExcel();
            if (currentMileage > Probeg)
            {
                string message = $"Внимание! Пора сделать тех. обслуживание машины! Пробег достигает {Probeg}, за последнее время было пройдено {difference} км.";
                bot.SendTextMessageAsync("1083342768", message);
                
            }
            else
            {
                string message = $"Скоро нужно будет сделать техническое обслуживание машины! Пробег машины достигает {Probeg}, за последний месяц пройдено {difference} км.";
                bot.SendTextMessageAsync("1083342768", message);
                
            }
        }

    }

}
