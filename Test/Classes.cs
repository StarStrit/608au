﻿using System;
using System.Reflection;
using System.IO;
using System.Drawing;
using System.Windows.Forms;

namespace Test
{
    public partial class Classes : Form
    {
        private Button graphic_Form608; // создание глобальной закрытой ссылки на кнопку
        private int nedely_table; // для номера недели
        #region Включение двойной буфферизации для быстрой перерисовки таблицы
        void SetDoubleBuffered2(Control c, bool value)
        {
            PropertyInfo pi = typeof(Control).GetProperty("DoubleBuffered", BindingFlags.SetProperty | BindingFlags.Instance | BindingFlags.NonPublic);
            if (pi != null)
                pi.SetValue(c, value, null);
        }
        #endregion
        /////////////////////////////////////////////////////////////////////////////////////////////
        public Classes(Button graphic_Main, int nedelytable) // получаем исходную ссылку на кнопку
        {
            graphic_Form608 = graphic_Main; // принимаем ссылку на кнопку
            nedely_table = nedelytable; // получаем номер недели
            InitializeComponent();
        }
        private void Form608_Load(object sender, EventArgs e)
        {
            // заполнение таблицы графика базовыми строками и столбцами при запуске
            dataGridView2.RowCount = 18; dataGridView2.ColumnCount = 7;
            dataGridView2.Columns[0].Width = 50;
            #region Красим заголовки недель таблицы
            for (int i = 0; i < 7; i++)
                dataGridView2.Rows[0].Cells[i].Style.BackColor = Color.SteelBlue;
            for (int i = 0; i < 7; i++)
                dataGridView2.Rows[9].Cells[i].Style.BackColor = Color.SteelBlue;
            for (int i = 0; i < 7; i++)
                dataGridView2.Rows[2].Cells[i].Style.BackColor = Color.LightSteelBlue;
            for (int i = 3; i < 9; i++)
                dataGridView2.Rows[i].Cells[0].Style.BackColor = Color.LightSteelBlue;
            for (int i = 0; i < 7; i++)
                dataGridView2.Rows[11].Cells[i].Style.BackColor = Color.LightSteelBlue;
            for (int i = 12; i < 18; i++)
                dataGridView2.Rows[i].Cells[0].Style.BackColor = Color.LightSteelBlue;
            #endregion
            #region Неделя 1
            dataGridView2.Rows[0].Cells[3].Style.ForeColor = Color.White;
            dataGridView2.Rows[0].Cells[3].Value = "1 неделя";
            dataGridView2.Rows[1].Cells[0].Value = "Пары";
            dataGridView2.Rows[1].Cells[1].Value = "1-я";
            dataGridView2.Rows[1].Cells[2].Value = "2-я";
            dataGridView2.Rows[1].Cells[3].Value = "3-я";
            dataGridView2.Rows[1].Cells[4].Value = "4-я";
            dataGridView2.Rows[1].Cells[5].Value = "5-я";
            dataGridView2.Rows[1].Cells[6].Value = "6-я";
            dataGridView2.Rows[2].Cells[0].Value = "Время";
            dataGridView2.Rows[2].Cells[1].Value = "09:00-10:35";
            dataGridView2.Rows[2].Cells[2].Value = "10:45-12:20";
            dataGridView2.Rows[2].Cells[3].Value = "13:00-14:35";
            dataGridView2.Rows[2].Cells[4].Value = "14:45-16:20";
            dataGridView2.Rows[2].Cells[5].Value = "16:25-18:00";
            dataGridView2.Rows[2].Cells[6].Value = "18:05-19:40";
            dataGridView2.Rows[3].Cells[0].Value = "Пнд";
            dataGridView2.Rows[4].Cells[0].Value = "Втр";
            dataGridView2.Rows[5].Cells[0].Value = "Срд";
            dataGridView2.Rows[6].Cells[0].Value = "Чтв";
            dataGridView2.Rows[7].Cells[0].Value = "Птн";
            dataGridView2.Rows[8].Cells[0].Value = "Сбт";
            #endregion
            #region Неделя 2
            dataGridView2.Rows[9].Cells[3].Style.ForeColor = Color.White;
            dataGridView2.Rows[9].Cells[3].Value = "2 неделя";
            dataGridView2.Rows[10].Cells[0].Value = "Пары";
            dataGridView2.Rows[10].Cells[1].Value = "1-я";
            dataGridView2.Rows[10].Cells[2].Value = "2-я";
            dataGridView2.Rows[10].Cells[3].Value = "3-я";
            dataGridView2.Rows[10].Cells[4].Value = "4-я";
            dataGridView2.Rows[10].Cells[5].Value = "5-я";
            dataGridView2.Rows[10].Cells[6].Value = "6-я";
            dataGridView2.Rows[11].Cells[0].Value = "Время";
            dataGridView2.Rows[11].Cells[1].Value = "09:00-10:35";
            dataGridView2.Rows[11].Cells[2].Value = "10:45-12:20";
            dataGridView2.Rows[11].Cells[3].Value = "13:00-14:35";
            dataGridView2.Rows[11].Cells[4].Value = "14:45-16:20";
            dataGridView2.Rows[11].Cells[5].Value = "16:25-18:00";
            dataGridView2.Rows[11].Cells[6].Value = "18:05-19:40";
            dataGridView2.Rows[12].Cells[0].Value = "Пнд";
            dataGridView2.Rows[13].Cells[0].Value = "Втр";
            dataGridView2.Rows[14].Cells[0].Value = "Срд";
            dataGridView2.Rows[15].Cells[0].Value = "Чтв";
            dataGridView2.Rows[16].Cells[0].Value = "Птн";
            dataGridView2.Rows[17].Cells[0].Value = "Сбт";
            #endregion
            HighlightColor(); // выделяем цветом активную пару недели
            #region Заполняем график занятий 608 ау.
            string path_file = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Расписание кафедры");
            string file_rasp = Path.Combine(path_file, Properties.Settings.Default.date_rasp.ToString("dd.MM.yyyy") + ".txt");
            StreamReader rasp_reader = new StreamReader(file_rasp);

            string bufferfile = "";
            string NamePrep;
            while (bufferfile != null)
            {
                bufferfile = rasp_reader.ReadLine();
                if ((bufferfile != null) && (bufferfile.IndexOf("#ff00ff\"> ") >= 0))
                {
                    // сокращаем имя препода
                    NamePrep = bufferfile.Substring(bufferfile.IndexOf("COLOR=\"#ff00ff\"> ") + 17, bufferfile.Length - bufferfile.IndexOf("COLOR=\"#ff00ff\"> ") - 28);
                    NamePrep = NamePrep.Remove(1, NamePrep.IndexOf(" ")); // удаляем лишние символы в строке до пробела (без учета первого символа)
                    NamePrep = NamePrep.Replace(".", ""); // удаление оставшегося лишнего символа .
                    #region 1 неделя
                    for (int i = 3; i <= 8; i++)
                    {
                        while (bufferfile.IndexOf("SIZE=2><P ALIGN=\"CENTER\">") == -1)
                            bufferfile = rasp_reader.ReadLine();
                        bufferfile = rasp_reader.ReadLine();
                        for (int j = 1; j < 7; j++)
                        {
                            bufferfile = rasp_reader.ReadLine();
                            // если пара в 608б
                            if (bufferfile.IndexOf("а.608б") >= 0)
                                if (dataGridView2.Rows[i].Cells[j].Value == null)
                                    dataGridView2.Rows[i].Cells[j].Value = NamePrep + " - " + bufferfile.Substring(bufferfile.IndexOf("ALIGN=\"CENTER\">") + 15, bufferfile.Length - bufferfile.IndexOf("</F") - 2) + "\n608б";
                                else
                                    dataGridView2.Rows[i].Cells[j].Value += "\n-------------------\n" + NamePrep + " - " + bufferfile.Substring(bufferfile.IndexOf("ALIGN=\"CENTER\">") + 15, bufferfile.Length - bufferfile.IndexOf("</F") - 2) + "\n608б";
                            else
                            // если пар нет или они в 608а
                            if ((bufferfile.IndexOf("\">_") == -1) && (bufferfile.IndexOf("а.608а") >= 0))
                                if (dataGridView2.Rows[i].Cells[j].Value == null)
                                    dataGridView2.Rows[i].Cells[j].Value = NamePrep + " - " + bufferfile.Substring(bufferfile.IndexOf("ALIGN=\"CENTER\">") + 15, bufferfile.Length - bufferfile.IndexOf("</F") - 2) + "\n608а";
                                else
                                    dataGridView2.Rows[i].Cells[j].Value += "\n-------------------\n" + NamePrep + " - " + bufferfile.Substring(bufferfile.IndexOf("ALIGN=\"CENTER\">") + 15, bufferfile.Length - bufferfile.IndexOf("</F") - 2) + "\n608а";
                            bufferfile = rasp_reader.ReadLine();
                        }
                    }
                    #endregion
                    #region 2 неделя
                    bufferfile = rasp_reader.ReadLine();
                    for (int i = 12; i <= 17; i++)
                    {
                        while (bufferfile.IndexOf("SIZE=2 COLOR=\"#0000ff\"><P ALIGN=\"CENTER\">") == -1)
                            bufferfile = rasp_reader.ReadLine();
                        bufferfile = rasp_reader.ReadLine();
                        for (int j = 1; j < 7; j++)
                        {
                            bufferfile = rasp_reader.ReadLine();
                            // если пара в 608б
                            if (bufferfile.IndexOf("а.608б") >= 0)
                                if (dataGridView2.Rows[i].Cells[j].Value == null)
                                    dataGridView2.Rows[i].Cells[j].Value = NamePrep + " - " + bufferfile.Substring(bufferfile.IndexOf("ALIGN=\"CENTER\">") + 15, bufferfile.Length - bufferfile.IndexOf("</F") - 2) + "\n608б";
                                else
                                    dataGridView2.Rows[i].Cells[j].Value += "\n-------------------\n" + NamePrep + " - " + bufferfile.Substring(bufferfile.IndexOf("ALIGN=\"CENTER\">") + 15, bufferfile.Length - bufferfile.IndexOf("</F") - 2) + "\n608б";
                            else
                            // если пар нет или они в 608а
                            if ((bufferfile.IndexOf("\"> _") == -1) && (bufferfile.IndexOf("а.608а") >= 0))
                                if (dataGridView2.Rows[i].Cells[j].Value == null)
                                    dataGridView2.Rows[i].Cells[j].Value = NamePrep + " - " + bufferfile.Substring(bufferfile.IndexOf("ALIGN=\"CENTER\">") + 15, bufferfile.Length - bufferfile.IndexOf("</F") - 2) + "\n608а";
                                else
                                    dataGridView2.Rows[i].Cells[j].Value += "\n-------------------\n" + NamePrep + " - " + bufferfile.Substring(bufferfile.IndexOf("ALIGN=\"CENTER\">") + 15, bufferfile.Length - bufferfile.IndexOf("</F") - 2) + "\n608а";
                            bufferfile = rasp_reader.ReadLine();
                        }
                    }
                    #endregion
                }
            }
            dataGridView2.Enabled = true;
            SetDoubleBuffered2(dataGridView2, true);
            rasp_reader.Dispose();
            #endregion
        }
        private void Form608_FormClosed(object sender, FormClosedEventArgs e)
        {
            graphic_Form608.Enabled = true; // включаем кнопку снова, после закрытия графика расписаний
        }
        private void dataGridView2_SelectionChanged(object sender, EventArgs e)
        {
            // отменяет выделение верхней левой ячейки при начальном запуске
            if (MouseButtons != MouseButtons.Left)
                ((DataGridView)sender).CurrentCell = null;
        }
        private void dataGridView2_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            // Убираем выделение ячейки
            ((DataGridView)sender).CurrentCell = null;
        }
        private void dataGridView2_CellMouseMove(object sender, DataGridViewCellMouseEventArgs e)
        {
            // перерисовываем таблицу с выделением строки
            dataGridView2.Invalidate();
        }
        private void dataGridView2_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            // рисуем границу ячейки при наведении на неё указателя
            if (e.RowIndex != 0 && e.RowIndex != 1 && e.RowIndex != 2 && e.RowIndex != 9 && e.RowIndex != 10 && e.RowIndex != 11) // не выделяем строки с заголовками
                if (dataGridView2.RectangleToScreen(e.RowBounds).Contains(MousePosition))
                {
                    var boundDataGrid = e.RowBounds; boundDataGrid.Width -= 2; boundDataGrid.Height -= 2; // изменяем выделение, чтобы оно влезло в ячейку полностью
                    e.Graphics.FillRectangle(new SolidBrush(Color.FromArgb(35, Color.SteelBlue)), boundDataGrid);
                    e.Graphics.DrawRectangle(new Pen(Color.SteelBlue), boundDataGrid);
                }
        }
        private void dataGridView2_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            // запрещаем выделение строки с заголовком недель
            if (e.RowIndex == 0 || e.RowIndex == 9)
                dataGridView2.ClearSelection();
        }
        /////////////////////////////////////////////////////////////////////////////////////////////
        private void HighlightColor()
        {
            var date_time = "";
            int nedelyA = 0, nedelyB = 0, nedelyC = 0;
            #region Проверяем неделю расписания
            if (nedely_table == 0) // первая неделя расписания
            {
                nedelyA = 3;
                nedelyB = 9;
                nedelyC = 2;
            }
            else // втораяя неделя расписания
            {
                nedelyA = 12;
                nedelyB = 18;
                nedelyC = 11;
            }
            #endregion
            #region Выделяем текущий день недели
            for (int i = nedelyA; i < nedelyB; i++)
            {
                date_time = dataGridView2.Rows[i].Cells[0].Value.ToString();
                if (DateTime.Today.ToString("ddd") == date_time.Remove(2, 1))
                {
                    dataGridView2.Rows[i].Cells[0].Style.BackColor = Color.Yellow;
                    // выделяем текущее время дня и пары
                    for (int j = 1; j < 7; j++)
                    {
                        date_time = dataGridView2.Rows[nedelyC].Cells[j].Value.ToString(); date_time = date_time.Remove(0, 6); // получаем нужный формат времени для TimeSpan
                        if (DateTime.Now.TimeOfDay <= new TimeSpan(Convert.ToInt32(date_time.Remove(2, 3)), Convert.ToInt32(date_time.Remove(0, 3)), 0))
                        {
                            // если время от 0.00 и до 8.59, то пару не выделяем
                            if (DateTime.Now.TimeOfDay < new TimeSpan(9, 0, 0))
                            {
                                i = nedelyB; // глушилка цикла 1
                                break; // глушилка цикла 2
                            }
                            // иначе, пару выделяем по текущему времени в таблице
                            dataGridView2.Rows[nedelyC].Cells[j].Style.BackColor = Color.Yellow;
                            dataGridView2.Rows[i].Cells[j].Style.BackColor = Color.Yellow;
                            i = nedelyB; // глушилка цикла 1
                            break; // глушилка цикла 2
                        }
                    }
                }
            }
            #endregion
        }
    }
}