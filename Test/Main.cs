using System;
using System.IO;
using System.Net;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using System.Drawing;
using System.Reflection;

namespace Test
{
    public partial class ESPPForm : Form
    {
        private string web_esstu = ""; // закрытая строковая переменная для чтения и записи html страниц
        private int nedely_rasp = 0; // закрытая строковая переменная для хранения номера недели
        #region Грузим статусы сети: в сети; нет сети; обновление расписания, проверка
        private Bitmap status_con = new Bitmap(Properties.Resources.status_connected);
        private Bitmap status_dis = new Bitmap(Properties.Resources.status_disconnected);
        #endregion
        #region Унаследование для DoubleBuffered - включение двойной буфферизации для быстрой перерисовки dataGridView1
        void SetDoubleBuffered(Control c, bool value)
        {
            PropertyInfo pi = typeof(Control).GetProperty("DoubleBuffered", BindingFlags.SetProperty | BindingFlags.Instance | BindingFlags.NonPublic);
            if (pi != null)
                pi.SetValue(c, value, null);
        }
        #endregion
        public ESPPForm()
        {
            InitializeComponent();
            timer1.Start(); // запускаем таймер для постоянной проверки соединения с сайтом
            #region Выделяем ячейку прямоугольником под указателем
            Rectangle rect = new Rectangle(0, 0, 0, 0);
            dataGridView1.CellPainting += (s, e) =>
            {
                // рисуем границу ячейки при наведении на неё указателя
                if (dataGridView1.RectangleToScreen(e.CellBounds).Contains(MousePosition))
                    rect = e.CellBounds;
            };
            dataGridView1.Paint += (s, e) =>
            {
                e.Graphics.FillRectangle(new SolidBrush(Color.FromArgb(35, Color.SteelBlue)), rect);
                e.Graphics.DrawRectangle(new Pen(Color.SteelBlue), rect);
            };
            #endregion
        }
        private void Main_FormClosed(object sender, FormClosedEventArgs e)
        {
            comboBox1.Items.Clear();
            comboBox2.Items.Clear();
            web_esstu = ""; // очистка строки данных веб-страницы
            // File.Delete("caf39.txt");
            /* Работает с настройками программы
             * Properties.Settings.Default.starter++;
             * Properties.Settings.Default.Save();
             * MessageBox.Show(Properties.Settings.Default.starter.ToString(), "Уведомление", MessageBoxButtons.OK, MessageBoxIcon.Warning);*/
        }
        private void Main_Load(object sender, EventArgs e)
        {
            timer1.Start(); // запускаем таймер для постоянной проверки соединения с сайтом
            try
            {
                HttpWebRequest test_link = (HttpWebRequest)WebRequest.Create("https://portal.esstu.ru/raspisan.htm");
                HttpWebResponse test_response = (HttpWebResponse)test_link.GetResponse(); // проверка соединения с сайтом ВСГУТУ
                if (HttpStatusCode.OK == test_response.StatusCode)
                {
                    #region Узнаем номер кафедры для ссылки
                    WebClient CafClient = new WebClient();
                    Stream CafStream = CafClient.OpenRead("https://portal.esstu.ru/bakalavriat/craspisanEdt.htm");
                    StreamReader CafReader = new StreamReader(CafStream, System.Text.Encoding.Default);
                    string Cafstr = "";
                    string NomerCaf = ""; // нужня для хранения продолжения ссылки (номера кафедры)
                    while (!CafReader.EndOfStream)
                    {
                        Cafstr = CafReader.ReadLine();
                        if (Cafstr.IndexOf("Электроснабжение промышленных предприятий") >= 0)
                        {
                            NomerCaf = Cafstr.Substring(Cafstr.IndexOf("<a href=\"") + 9, Cafstr.IndexOf("<a href=\""));
                            break;
                        }
                        else if (CafReader.EndOfStream)
                        {
                            MessageBox.Show("Не обнаружена кафедра ЭСППиСХ.\nРасписание не доступно", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            return;
                        }
                    }
                    CafClient.Dispose(); CafStream.Close(); CafReader.Close();
                    #endregion
                    #region Загрузка рассписания кафедры
                    HttpWebRequest RaspReq = (HttpWebRequest)WebRequest.Create("https://portal.esstu.ru/bakalavriat/" + NomerCaf + ".htm"); // передача адреса строки
                    HttpWebResponse RaspRes = (HttpWebResponse)RaspReq.GetResponse(); // отправка запроса и получение ответа
                    Stream RaspStr = RaspRes.GetResponseStream(); // получение ответа ввиде потока информации
                    StreamReader RaspRead = new StreamReader(RaspStr, System.Text.Encoding.Default); // создание потоковой переменной для принятия данных
                    // StreamWriter RaspWrite = new StreamWriter("caf39.txt"); // создание текстового файла для хранения данных веб-страницы
                    web_esstu = RaspRead.ReadToEnd(); // сохранение загруженных данных в глобальную строку
                    // RaspWrite.WriteLine(response); // запись считанных данных в файл
                    RaspStr.Dispose();
                    RaspRead.Dispose();
                    // RaspWrite.Dispose();
                    RaspRes.Close();
                    #endregion
                    #region Узнаем дату обновления расписания
                    var request_caf39 = (HttpWebRequest)WebRequest.Create("https://portal.esstu.ru/bakalavriat/Caf39.htm");
                    request_caf39.UserAgent = "Rasp. 608";
                    using (var get_date = (HttpWebResponse)request_caf39.GetResponse())
                    {
                        DateTime date_time = get_date.LastModified;
                        label1.Text = "Дата обновления: " + date_time.ToString("dd.MM.yyyy");
                    }
                    #endregion
                    #region Узнаем текущую неделя расписания
                    WebClient client = new WebClient();
                    Stream data = client.OpenRead("https://esstu.ru/index.htm");
                    StreamReader reader = new StreamReader(data);
                    string response = "";
                    while (!reader.EndOfStream)
                    {
                        response = reader.ReadLine();
                        if (response.IndexOf("\"header-date\"") >= 0)
                        {
                            response = response.Substring(response.IndexOf("\"header-date\"") + 14, response.Length - 47);
                            daterasp.Text = "Общее расписание кафедры: " + response;
                            // выбираем текущую неделю в comboBox'е
                            comboBox2.SelectedIndex = 0;
                            if (response.IndexOf("II н") >= 0)
                            {
                                comboBox2.SelectedIndex = 1;
                                nedely_rasp = 1; // для передачи на форму "Form608"
                            }
                            break;
                        }
                        else if (reader.EndOfStream)
                            MessageBox.Show("Не могу определить номер недели.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                    client.Dispose();
                    data.Dispose();
                    reader.Dispose();
                    comboBox2.Enabled = true;
                    #endregion
                    #region Чтение файла, загрузка преподавателей и недель учебы
                    // StreamReader stremer = new StreamReader("caf39.txt");
                    TextReader skan_web = new StringReader(web_esstu); // создание считывающего потока строковых данных
                    response = "";
                    while (response != null)
                    {
                        response = skan_web.ReadLine();
                        if ((response != null) && (response.IndexOf("COLOR=\"#ff00ff\"> ") >= 0))
                        {
                            comboBox1.Items.Add(response.Substring(response.IndexOf("COLOR=\"#ff00ff\"> ") + 17, response.Length - response.IndexOf("COLOR=\"#ff00ff\"> ") - 28));
                        }
                    }
                    расписаниеToolStripMenuItem.Enabled = true;
                    comboBox1.SelectedIndex = 0;
                    comboBox1.Enabled = true;
                    dataGridView1.Enabled = true;
                    SetDoubleBuffered(dataGridView1, true); // включаем быструю перерисовку таблицы
                    graphic608.Enabled = true;
                    #endregion
                    #region Ставаим статус соединения - онлайн
                    pictureBox1.Size = status_con.Size;
                    pictureBox1.Image = status_con;
                    pictureBox1.Invalidate(); // перерисовка бокса
                    #endregion
                }
            }
            catch (WebException)
            {
                MessageBox.Show("Нет соединения с расписанием ВСГУТУ.\nРасписание не доступно.", "Ошибка соединения", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            #region Внесение базового рассписания в dataGridView
            dataGridView1.Rows.Clear();
            dataGridView1.RowCount = 0; dataGridView1.RowCount = 7;
            for (int i = 0; i < 7; i++)
                dataGridView1.Rows[0].Cells[i].Style.BackColor = Color.Gainsboro;
            for (int i = 1; i < 7; i++)
                dataGridView1.Rows[i].Cells[0].Style.BackColor = Color.Gainsboro;
            dataGridView1.Rows[0].Cells[0].Value = "Время";
            dataGridView1.Rows[0].Cells[1].Value = "09:00-10:35";
            dataGridView1.Rows[0].Cells[2].Value = "10:45-12:20";
            dataGridView1.Rows[0].Cells[3].Value = "13:00-14:35";
            dataGridView1.Rows[0].Cells[4].Value = "14:45-16:20";
            dataGridView1.Rows[0].Cells[5].Value = "16:25-18:00";
            dataGridView1.Rows[0].Cells[6].Value = "18:05-19:40";
            dataGridView1.Rows[1].Cells[0].Value = "Пнд";
            dataGridView1.Rows[2].Cells[0].Value = "Втр";
            dataGridView1.Rows[3].Cells[0].Value = "Срд";
            dataGridView1.Rows[4].Cells[0].Value = "Чтв";
            dataGridView1.Rows[5].Cells[0].Value = "Птн";
            dataGridView1.Rows[6].Cells[0].Value = "Сбт";
            #endregion
            #region Выделяем цветом активную пару/время/день, которая идёт сейчас
            var date_time = ""; // хранит время и день недели, взятое из таблицы
            // проверяем, соответствует ли выбранная неделя с текущей: нет - выделение отменяем; да - выделение начинаем
            if ((comboBox2.SelectedIndex == 0 && nedely_rasp == 0) || (comboBox2.SelectedIndex == 1 && nedely_rasp == 1))
                // выделяем текущую пару
                for (int i = 1; i < 7; i++)
                {
                    date_time = dataGridView1.Rows[i].Cells[0].Value.ToString();
                    if (DateTime.Today.ToString("ddd") == date_time.Remove(2, 1))
                    {
                        dataGridView1.Rows[i].Cells[0].Style.BackColor = Color.Yellow;
                        // выделяем текущее время дня и пары
                        for (int j = 1; j < 7; j++)
                        {
                            date_time = dataGridView1.Rows[0].Cells[j].Value.ToString(); date_time = date_time.Remove(0, 6); // получаем нужный формат времени для TimeSpan
                            if (DateTime.Now.TimeOfDay <= new TimeSpan(Convert.ToInt32(date_time.Remove(2, 3)), Convert.ToInt32(date_time.Remove(0, 3)), 0))
                            {
                                // если время от 0.00 и до 8.59, то пару не выделем
                                if (DateTime.Now.TimeOfDay < new TimeSpan(9, 0, 0))
                                {
                                    i = 7; // глушилка цикла 1
                                    break; // глушилка цикла 2
                                }
                                // иначе, пару выделяем по текущему времени в таблице
                                dataGridView1.Rows[0].Cells[j].Style.BackColor = Color.Yellow;
                                dataGridView1.Rows[i].Cells[j].Style.BackColor = Color.Yellow;
                                i = 7; // глушилка цикла 1
                                break; // глушилка цикла 2
                            }
                        }
                    }
                }
            #endregion
            #region Составление расписания при выборе преподавателя
            // StreamReader raspstream = new StreamReader("caf39.txt");
            TextReader skan_web = new StringReader(web_esstu);
            string raspstr = "";
            while (raspstr != null)
            {
                raspstr = skan_web.ReadLine();
                // ищем тег преподавателя в файле
                if ((raspstr != null) && (raspstr.IndexOf("COLOR=\"#ff00ff\"> ") >= 0))
                {
                    // сверяем препода из файла с преподом из comboBox 
                    if ((raspstr.Substring(raspstr.IndexOf("COLOR=\"#ff00ff\"> ") + 17, raspstr.Length - raspstr.IndexOf("COLOR=\"#ff00ff\"> ") - 28) == comboBox1.Text))
                    {
                        // составляем рассписание препода, пока не закончилась нужная неделя
                        while ((raspstr != null) && (raspstr.IndexOf("SIZE=2 COLOR=\"#0000ff\"") == -1))
                        {
                            raspstr = skan_web.ReadLine();
                            #region 1 неделя
                            // если выбрана 1 неделя
                            if ((raspstr.IndexOf("SIZE=2><P ALIGN=\"CENTER\">") >= 0) && (comboBox2.SelectedItem.ToString() == "1 неделя"))
                            {
                                while ((raspstr != null) && (raspstr != "</TR>"))
                                {
                                    for (int i = 1; i <= 6; i++)
                                    {
                                        while (raspstr.IndexOf("SIZE=2><P ALIGN=\"CENTER\">") == -1) // тег для поиска дня недели
                                            raspstr = skan_web.ReadLine();
                                        raspstr = skan_web.ReadLine();
                                        for (int j = 1; j != 7; j++)
                                        {
                                            raspstr = skan_web.ReadLine();
                                            // меняем цвет ячейки пары, если она идет в 608 ау.
                                            if ((raspstr.IndexOf("а.608") >= 0) && (dataGridView1.Rows[i].Cells[j].Style.BackColor != Color.Yellow))
                                                dataGridView1.Rows[i].Cells[j].Style.BackColor = Color.PowderBlue;
                                            // если пар нет
                                            if (raspstr.IndexOf("\">_") >= 0)
                                                dataGridView1.Rows[i].Cells[j].Value = "-";
                                            else // или они есть
                                            {
                                                #region Форматирование текста
                                                raspstr = raspstr.Substring(raspstr.IndexOf("ALIGN=\"CENTER\">") + 15, raspstr.Length - raspstr.IndexOf("ALIGN=\"CENTER\">") - 27);
                                                raspstr = raspstr.Replace("    ","\n"); // заменя пробелов на переход новой строки
                                                raspstr = raspstr.Replace("...", ""); // замена троеточия на пустой символ
                                                raspstr = raspstr.Trim(); // обрезка строки
                                                raspstr = raspstr.Replace("пр.", "\n-пр.-\n");
                                                raspstr = raspstr.Replace("лаб.", "\n-лаб.-\n");
                                                raspstr = raspstr.Replace("лек.", "\n-лек.-\n");
                                                #endregion
                                                dataGridView1.Rows[i].Cells[j].Value = raspstr;
                                            }
                                            raspstr = skan_web.ReadLine();
                                        }
                                    }
                                    skan_web.Dispose();
                                    return;
                                }
                            }
                            #endregion
                            #region 2 неделя
                            // если выбрана 2 неделя
                            if ((raspstr.IndexOf("SIZE=2 COLOR=\"#0000ff\"><P ALIGN=\"CENTER\">") >= 0) && (comboBox2.SelectedItem.ToString() == "2 неделя"))
                            {
                                while ((raspstr != null) && (raspstr != "</TR>"))
                                {
                                    for (int i = 1; i <= 6; i++)
                                    {
                                        while (raspstr.IndexOf("SIZE=2 COLOR=\"#0000ff\"><P ALIGN=\"CENTER\">") == -1)
                                            raspstr = skan_web.ReadLine();
                                        raspstr = skan_web.ReadLine();
                                        for (int j = 1; j != 7; j++)
                                        {
                                            raspstr = skan_web.ReadLine();
                                            if ((raspstr.IndexOf("а.608") >= 0) && (dataGridView1.Rows[i].Cells[j].Style.BackColor != Color.Yellow))
                                                dataGridView1.Rows[i].Cells[j].Style.BackColor = System.Drawing.Color.PowderBlue;
                                            if (raspstr.IndexOf("\"> _") >= 0)
                                                dataGridView1.Rows[i].Cells[j].Value = "-";
                                            else
                                            {
                                                #region Форматирование текста
                                                raspstr = raspstr.Substring(raspstr.IndexOf("ALIGN=\"CENTER\"> ") + 16, raspstr.Length - raspstr.IndexOf("ALIGN=\"CENTER\"> ") - 28);
                                                raspstr = raspstr.Replace("    ", "\n");
                                                raspstr = raspstr.Replace("...", "");
                                                raspstr = raspstr.Trim();
                                                raspstr = raspstr.Replace("пр.", "\n-пр.-\n");
                                                raspstr = raspstr.Replace("лаб.", "\n-лаб.-\n");
                                                raspstr = raspstr.Replace("лек.", "\n-лек.-\n");
                                                #endregion
                                                dataGridView1.Rows[i].Cells[j].Value = raspstr;
                                            }
                                            raspstr = skan_web.ReadLine();
                                        }
                                    }
                                    skan_web.Dispose();
                                    return;
                                }
                            }
                            #endregion
                        }
                    }
                }
            }
            #endregion
        }
        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            #region Внесение базового рассписания в dataGridView
            dataGridView1.Rows.Clear();
            dataGridView1.RowCount = 0; dataGridView1.RowCount = 7;
            for (int i = 0; i < 7; i++)
                dataGridView1.Rows[0].Cells[i].Style.BackColor = Color.Gainsboro;
            for (int i = 1; i < 7; i++)
                dataGridView1.Rows[i].Cells[0].Style.BackColor = Color.Gainsboro;
            dataGridView1.Rows[0].Cells[0].Value = "Время";
            dataGridView1.Rows[0].Cells[1].Value = "09:00-10:35";
            dataGridView1.Rows[0].Cells[2].Value = "10:45-12:20";
            dataGridView1.Rows[0].Cells[3].Value = "13:00-14:35";
            dataGridView1.Rows[0].Cells[4].Value = "14:45-16:20";
            dataGridView1.Rows[0].Cells[5].Value = "16:25-18:00";
            dataGridView1.Rows[0].Cells[6].Value = "18:05-19:40";
            dataGridView1.Rows[1].Cells[0].Value = "Пнд";
            dataGridView1.Rows[2].Cells[0].Value = "Втр";
            dataGridView1.Rows[3].Cells[0].Value = "Срд";
            dataGridView1.Rows[4].Cells[0].Value = "Чтв";
            dataGridView1.Rows[5].Cells[0].Value = "Птн";
            dataGridView1.Rows[6].Cells[0].Value = "Сбт";
            #endregion
            #region Выделяем цветом активную пару/время/день, которая идёт сейчас
            var date_time = ""; // хранит время и день недели, взятое из таблицы
            // проверяем, соответствует ли выбранная неделя с текущей: нет - выделение отменяем; да - выделение начинаем
            if ((comboBox2.SelectedIndex == 0 && nedely_rasp == 0) || (comboBox2.SelectedIndex == 1 && nedely_rasp == 1))
                // выделяем текущую пару
                for (int i = 1; i < 7; i++)
                {
                    date_time = dataGridView1.Rows[i].Cells[0].Value.ToString();
                    if (DateTime.Today.ToString("ddd") == date_time.Remove(2, 1))
                    {
                        dataGridView1.Rows[i].Cells[0].Style.BackColor = Color.Yellow;
                        // выделяем текущее время дня и пары
                        for (int j = 1; j < 7; j++)
                        {
                            date_time = dataGridView1.Rows[0].Cells[j].Value.ToString(); date_time = date_time.Remove(0, 6); // получаем нужный формат времени для TimeSpan
                            if (DateTime.Now.TimeOfDay <= new TimeSpan(Convert.ToInt32(date_time.Remove(2, 3)), Convert.ToInt32(date_time.Remove(0, 3)), 0))
                            {
                                // если время от 0.00 и до 8.59, то пару не выделем
                                if (DateTime.Now.TimeOfDay < new TimeSpan(9, 0, 0))
                                {
                                    i = 7; // глушилка цикла 1
                                    break; // глушилка цикла 2
                                }
                                // иначе, пару выделяем по текущему времени в таблице
                                dataGridView1.Rows[0].Cells[j].Style.BackColor = Color.Yellow;
                                dataGridView1.Rows[i].Cells[j].Style.BackColor = Color.Yellow;
                                i = 7; // глушилка цикла 1
                                break; // глушилка цикла 2
                            }
                        }
                    }
                }
            #endregion
            #region Составление расписания при выборе преподавателя
            // StreamReader raspstream = new StreamReader("caf39.txt");
            TextReader skan_web = new StringReader(web_esstu);
            string raspstr = "";
            while (raspstr != null)
            {
                raspstr = skan_web.ReadLine();
                // ищем тег преподавателя в файле
                if ((raspstr != null) && (raspstr.IndexOf("COLOR=\"#ff00ff\"> ") >= 0))
                {
                    // сверяем препода из файла с преподом из comboBox 
                    if ((raspstr.Substring(raspstr.IndexOf("COLOR=\"#ff00ff\"> ") + 17, raspstr.Length - raspstr.IndexOf("COLOR=\"#ff00ff\"> ") - 28) == comboBox1.Text))
                    {
                        // составляем рассписание препода, пока не закончилась нужная неделя
                        while ((raspstr != null) && (raspstr.IndexOf("SIZE=2 COLOR=\"#0000ff\"") == -1))
                        {
                            raspstr = skan_web.ReadLine();
                            #region 1 неделя
                            // если выбрана 1 неделя
                            if ((raspstr.IndexOf("SIZE=2><P ALIGN=\"CENTER\">") >= 0) && (comboBox2.SelectedItem.ToString() == "1 неделя"))
                            {
                                while ((raspstr != null) && (raspstr != "</TR>"))
                                {
                                    for (int i = 1; i <= 6; i++)
                                    {
                                        while (raspstr.IndexOf("SIZE=2><P ALIGN=\"CENTER\">") == -1) // тег для поиска дня недели
                                            raspstr = skan_web.ReadLine();
                                        raspstr = skan_web.ReadLine();
                                        for (int j = 1; j != 7; j++)
                                        {
                                            raspstr = skan_web.ReadLine();
                                            // меняем цвет ячейки пары, если она идет в 608 ау.
                                            if ((raspstr.IndexOf("а.608") >= 0) && (dataGridView1.Rows[i].Cells[j].Style.BackColor != Color.Yellow))
                                                dataGridView1.Rows[i].Cells[j].Style.BackColor = Color.PowderBlue;
                                            // если пар нет
                                            if (raspstr.IndexOf("\">_") >= 0)
                                                dataGridView1.Rows[i].Cells[j].Value = "-";
                                            else // или они есть
                                            {
                                                #region Форматирование текста
                                                raspstr = raspstr.Substring(raspstr.IndexOf("ALIGN=\"CENTER\">") + 15, raspstr.Length - raspstr.IndexOf("ALIGN=\"CENTER\">") - 27);
                                                raspstr = raspstr.Replace("    ", "\n");
                                                raspstr = raspstr.Replace("...", "");
                                                raspstr = raspstr.Trim();
                                                raspstr = raspstr.Replace("пр.", "\n-пр.-\n");
                                                raspstr = raspstr.Replace("лаб.", "\n-лаб.-\n");
                                                raspstr = raspstr.Replace("лек.", "\n-лек.-\n");
                                                #endregion
                                                dataGridView1.Rows[i].Cells[j].Value = raspstr;
                                            }
                                            raspstr = skan_web.ReadLine();
                                        }
                                    }
                                    skan_web.Dispose();
                                    return;
                                }
                            }
                            #endregion
                            #region 2 неделя
                            // если выбрана 2 неделя
                            if ((raspstr.IndexOf("SIZE=2 COLOR=\"#0000ff\"><P ALIGN=\"CENTER\">") >= 0) && (comboBox2.SelectedItem.ToString() == "2 неделя"))
                            {
                                while ((raspstr != null) && (raspstr != "</TR>"))
                                {
                                    for (int i = 1; i <= 6; i++)
                                    {
                                        while (raspstr.IndexOf("SIZE=2 COLOR=\"#0000ff\"><P ALIGN=\"CENTER\">") == -1)
                                            raspstr = skan_web.ReadLine();
                                        raspstr = skan_web.ReadLine();
                                        for (int j = 1; j != 7; j++)
                                        {
                                            raspstr = skan_web.ReadLine();
                                            if ((raspstr.IndexOf("а.608") >= 0) && (dataGridView1.Rows[i].Cells[j].Style.BackColor != Color.Yellow))
                                                dataGridView1.Rows[i].Cells[j].Style.BackColor = System.Drawing.Color.PowderBlue;
                                            if (raspstr.IndexOf("\"> _") >= 0)
                                                dataGridView1.Rows[i].Cells[j].Value = "-";
                                            else
                                            {
                                                #region Форматирование текста
                                                raspstr = raspstr.Substring(raspstr.IndexOf("ALIGN=\"CENTER\"> ") + 16, raspstr.Length - raspstr.IndexOf("ALIGN=\"CENTER\"> ") - 28);
                                                raspstr = raspstr.Replace("    ", "\n");
                                                raspstr = raspstr.Replace("...", "");
                                                raspstr = raspstr.Trim();
                                                raspstr = raspstr.Replace("пр.", "\n-пр.-\n");
                                                raspstr = raspstr.Replace("лаб.", "\n-лаб.-\n");
                                                raspstr = raspstr.Replace("лек.", "\n-лек.-\n");
                                                #endregion
                                                dataGridView1.Rows[i].Cells[j].Value = raspstr;
                                            }
                                            raspstr = skan_web.ReadLine();
                                        }
                                    }
                                    skan_web.Dispose();
                                    return;
                                }
                            }
                            #endregion
                        }
                    }
                }
            }
            #endregion
        }
        private void сформироватьГрафикЗанятийДляПечатиToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Word.Document document = null;
            try
            {
                #region Открываем doc-файл Word из ресурса
                this.Cursor = Cursors.WaitCursor; // показываем ожидающий курсор, мол - "жди, идет работа" :)
                Word.Application docApp = new Word.Application();
                var fullName = Path.Combine(Environment.ExpandEnvironmentVariables("%USERPROFILE%/Documents"), "608au.doc"); // составляем строку для хранения документа во временной папке
                if (File.Exists(fullName) == true) // проверка на наличие уже существующего файла и его удаление
                    File.Delete(fullName);
                byte[] resource_doc = Properties.Resources.word_test; // находим и считываем документ из ресурса программы
                File.WriteAllBytes(fullName, resource_doc); // копируем данные из документа байтового масиисва в ресурсах, в документ временной папки
                document = docApp.Documents.Open(fullName, ReadOnly: false, Visible: false);
                document.Activate();
                #endregion
                #region Вставляем текущую дату рассписания
                Word.Bookmarks date_book = document.Bookmarks;
                Word.Range date_range;
                foreach (Word.Bookmark mark in date_book)
                {
                    date_range = mark.Range;
                    date_range.Text = label1.Text.ToString();
                }
                #endregion
                #region Записываем расписание пар в таблицу
                Word.Table table = docApp.ActiveDocument.Tables[1];
                // StreamReader filehtm = new StreamReader("caf39.txt");
                TextReader skan_web = new StringReader(web_esstu);
                string bufferfile = "";
                string NamePrep;
                int i = 4, q = 1;
                while (bufferfile != null)
                {
                    bufferfile = skan_web.ReadLine();
                    if ((bufferfile != null) && (bufferfile.IndexOf("#ff00ff\"> ") >= 0))
                    {
                        // сохраняем сокращенное имя препода в хранилище
                        NamePrep = bufferfile.Substring(bufferfile.IndexOf("COLOR=\"#ff00ff\"> ") + 17, bufferfile.Length - bufferfile.IndexOf("COLOR=\"#ff00ff\"> ") - 28);
                        NamePrep = NamePrep.Remove(1, NamePrep.IndexOf(" ")); // удаляем лишние символы в строке до пробела (без учета первого символа)
                        NamePrep = NamePrep.Replace(".", ""); // удаление оставшегося лишнего символа .
                        bufferfile = skan_web.ReadLine();
                        #region 1 неделя
                        while (i <= 44) // пока не достигли конца таблицы - конца недели
                        {
                            while (bufferfile.IndexOf("SIZE=2><P ALIGN=\"CENTER\">") == -1) // тег для поиска дня недели
                                bufferfile = skan_web.ReadLine();
                            bufferfile = skan_web.ReadLine();
                            while (q <= 6) // счетчик пар
                            {
                                bufferfile = skan_web.ReadLine();
                                if ((bufferfile.IndexOf("\">_") == -1) && (bufferfile.IndexOf("а.608а") >= 0))
                                    table.Rows[i].Cells[3].Range.Text = NamePrep + " - " + bufferfile.Substring(bufferfile.IndexOf("ALIGN=\"CENTER\">") + 15, bufferfile.Length - bufferfile.IndexOf("</F") - 2);
                                else
                                if ((bufferfile.IndexOf("\">_") == -1) && (bufferfile.IndexOf("а.608б") >= 0))
                                    table.Rows[i].Cells[4].Range.Text = NamePrep + " - " + bufferfile.Substring(bufferfile.IndexOf("ALIGN=\"CENTER\">") + 15, bufferfile.Length - bufferfile.IndexOf("</F") - 2);
                                bufferfile = skan_web.ReadLine();
                                q++; i++;
                            }
                            i++; q = 1;
                            bufferfile = skan_web.ReadLine();
                        }
                        i = 4;
                        #endregion
                        #region 2 неделя
                        while (i <= 44)
                        {
                            while (bufferfile.IndexOf("SIZE=2 COLOR=\"#0000ff\"><P ALIGN=\"CENTER\">") == -1)
                                bufferfile = skan_web.ReadLine();
                            bufferfile = skan_web.ReadLine();
                            while (q <= 6)
                            {
                                bufferfile = skan_web.ReadLine();
                                if ((bufferfile.IndexOf("\"> _") == -1) && (bufferfile.IndexOf("а.608а") >= 0))
                                    table.Rows[i].Cells[7].Range.Text = NamePrep + " - " + bufferfile.Substring(bufferfile.IndexOf("ALIGN=\"CENTER\">") + 15, bufferfile.Length - bufferfile.IndexOf("</F") - 2);
                                else
                                if ((bufferfile.IndexOf("\"> _") == -1) && (bufferfile.IndexOf("а.608б") >= 0))
                                    table.Rows[i].Cells[8].Range.Text = NamePrep + " - " + bufferfile.Substring(bufferfile.IndexOf("ALIGN=\"CENTER\">") + 15, bufferfile.Length - bufferfile.IndexOf("</F") - 2);
                                bufferfile = skan_web.ReadLine();
                                q++; i++;
                            }
                            i++; q = 1;
                            bufferfile = skan_web.ReadLine();
                        }
                        i = 4;
                        #endregion
                    }
                }
                skan_web.Dispose();
                #endregion
                this.Cursor = Cursors.Default;
                docApp.Visible = true;
                document = null;
            }
            catch (Exception ex)
            {
                // ошибки выводим
                MessageBox.Show(ex.Message, "Ошибка чтения", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                document = null;
            }
        }
        private void graphic608_Click(object sender, EventArgs e)
        {
            Form608 f608 = new Form608(this.graphic608, web_esstu, this.nedely_rasp); // создаем объект второй формы и передаем ей состояние кнопки формы, html сайт и день недели
            f608.Show();
            graphic608.Enabled = false; // блокируем кнопку, чтобы не открывать форму по несколько раз
        }
        private void оПрограммеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            About about_Form = new About();
            about_Form.Show();
        }
        private void pictureBox1_MouseEnter(object sender, EventArgs e)
        {
            toolTip1.SetToolTip(pictureBox1, "Отображает состояние подключения к интернету, где:\n     * Зеленый индикатор - интернет подключен;\n     * Красный индикатор - нет подключение к интернету (загружено резервное расписание)."); // всплывающая подсказка
        }
        private void timer1_Tick(object sender, EventArgs e)
        {
            // Проверяем соединение с сайтом и выводим статус сети
            try
            {
                HttpWebRequest request = (HttpWebRequest)HttpWebRequest.Create("https://portal.esstu.ru/raspisan.htm");
                HttpWebResponse response = (HttpWebResponse)request.GetResponse();
                if (HttpStatusCode.OK == response.StatusCode)
                {
                    response.Close();
                    #region Ставаим статус соединения - онлайн
                    pictureBox1.Size = status_con.Size;
                    pictureBox1.Image = status_con;
                    pictureBox1.Invalidate();
                    #endregion
                }
            }
            catch (WebException)
            {
                #region Ставаим статус соединения - не в сети
                pictureBox1.Size = status_dis.Size;
                pictureBox1.Image = status_dis;
                pictureBox1.Invalidate();
                #endregion
            }
        }
        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            // Отменяет выделение верхней левой ячейки при начальном запуске
            if (MouseButtons != MouseButtons.Left)
                ((DataGridView)sender).CurrentCell = null;
        }
        private void dataGridView1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            // Убираем выделение ячейки
            ((DataGridView)sender).CurrentCell = null;
        }
        private void dataGridView1_CellMouseEnter(object sender, DataGridViewCellEventArgs e)
        {
            // Перерисовываем таблицу с выделением ячейки
            dataGridView1.Invalidate();
        }
    }
}
// !Функции!
// Подумать над локальном сохранением расписаний по последним датам изменения
// Сделать справку - описание работы программы
// Сделать уведомление о наличии нового расписания
// Сделать автомат. подстройку высоты окна под таблицуы
// Добавить Артефакт и отзыв о программе (форма для заполнения со звездами, пожеланиями и отправкой мне на почту)
// Сделать загрузку программы красивее, подумать о Flat стиле
// Сделать кнопку - "Проверка обновлений"
// Сделать строки одинаковой ширины
// "Неизвестный издатель" при установке

// !Оптимизация!
// Программа виснет при проверки соединения