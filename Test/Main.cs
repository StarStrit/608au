using System;
using System.IO;
using System.Net;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using System.Drawing;
using System.Reflection;
using System.Diagnostics;

namespace Test
{
    public partial class ESPPForm : Form
    {
        private int nedely_rasp = 0; // для хранения номера недели
        #region Статусы сети: онлайн, оффлайн
        private Bitmap status_con = new Bitmap(Properties.Resources.status_connected);
        private Bitmap status_dis = new Bitmap(Properties.Resources.status_disconnected);
        #endregion
        #region Включение двойной буфферизации для быстрой перерисовки таблицы
        void SetDoubleBuffered(Control c, bool value)
        {
            PropertyInfo pi = typeof(Control).GetProperty("DoubleBuffered", BindingFlags.SetProperty | BindingFlags.Instance | BindingFlags.NonPublic);
            if (pi != null)
                pi.SetValue(c, value, null);
        }
        #endregion
        ///////////////////////////////////////////////////////////////////////////////////////////
        public ESPPForm()
        {
            InitializeComponent();
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
        private void ESPPForm_Shown(object sender, EventArgs e)
        {
            timer_updateStatus.Start(); // таймер для проверки соединения с сайтом и обновления расписания
            if (Properties.Settings.Default.starter == 0 | Properties.Settings.Default.number_nedel == "0") // если запускаем программу первый раз или настройки стандартные
            {
                try
                {
                    HttpWebRequest test_link = (HttpWebRequest)WebRequest.Create("https://portal.esstu.ru/raspisan.htm");
                    HttpWebResponse test_response = (HttpWebResponse)test_link.GetResponse(); // проверка соединения с сайтом ВСГУТУ
                    if (HttpStatusCode.OK == test_response.StatusCode) // если соединение установлено
                    {
                        #region Узнаем номер кафедры
                        WebClient CafClient = new WebClient();
                        Stream CafStream = CafClient.OpenRead("https://portal.esstu.ru/bakalavriat/craspisanEdt.htm");
                        StreamReader CafReader = new StreamReader(CafStream, System.Text.Encoding.Default);
                        string Cafstr = "";
                        string NomerCaf = ""; // нужна для хранения продолжения ссылки (номера кафедры)
                        while (!CafReader.EndOfStream)
                        {
                            Cafstr = CafReader.ReadLine();
                            if (Cafstr.IndexOf("Электроснабжение промышленных предприятий") >= 0)
                            {
                                NomerCaf = Cafstr.Substring(Cafstr.IndexOf("<a href=\"") + 9, Cafstr.IndexOf("<a href=\""));
                                Properties.Settings.Default.number_caf = NomerCaf;
                                Properties.Settings.Default.Save();
                                break;
                            }
                            else if (CafReader.EndOfStream)
                            {
                                MessageBox.Show("Не обнаружена кафедра ЭСППиСХ.\nРасписание не доступно!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                return;
                            }
                        }
                        CafClient.Dispose(); CafStream.Close(); CafReader.Close();
                        #endregion
                        #region Узнаем дату обновления расписания
                        var request_caf39 = (HttpWebRequest)WebRequest.Create("https://portal.esstu.ru/bakalavriat/Caf39.htm");
                        request_caf39.UserAgent = "Rasp. 608";
                        using (var get_date = (HttpWebResponse)request_caf39.GetResponse())
                        {
                            DateTime date_time = get_date.LastModified;
                            Properties.Settings.Default.date_rasp = date_time;
                            Properties.Settings.Default.Save();
                            label1.Text = "Дата обновления: " + date_time.ToString("dd.MM.yyyy");
                        }
                        #endregion
                        #region Загрузка рассписания кафедры
                        // отправка запроса на подключение к сайту 
                        HttpWebRequest RaspReq = (HttpWebRequest)WebRequest.Create("https://portal.esstu.ru/bakalavriat/" + NomerCaf + ".htm"); // передача адреса строки
                        HttpWebResponse RaspRes = (HttpWebResponse)RaspReq.GetResponse(); // отправка запроса и получение ответа
                        Stream RaspStr = RaspRes.GetResponseStream(); // получение ответа ввиде потока информации
                        StreamReader RaspRead = new StreamReader(RaspStr, System.Text.Encoding.Default); // создание потоковой переменной для принятия данных

                        // создание каталога в папке программы и файла расписания по датам
                        string path_file = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Расписание кафедры"); // путь к имени каталога рабочей папки
                        if (!Directory.Exists(path_file))
                            Directory.CreateDirectory(path_file); // создание каталога, если его нет в рабочей папке
                        string file_rasp = Path.Combine(path_file, Properties.Settings.Default.date_rasp.ToString("dd.MM.yyyy") + ".txt"); // создание текстового файла в каталоге расписания
                        StreamWriter raspfile_write = new StreamWriter(file_rasp); // создаём поток для записи данных в файл

                        // загрузка данных с сайта и сохранение
                        string web_esstu = ""; // для временного хранения html страницы
                        web_esstu = RaspRead.ReadToEnd(); // сохранение загруженных данных
                        raspfile_write.WriteLine(web_esstu); // запись считанных данных в файл
                        // чистка ресурсов
                        raspfile_write.Dispose();
                        RaspRead.Dispose();
                        RaspStr.Dispose();
                        RaspRes.Close();
                        #endregion
                        #region Узнаем текущую неделю расписания
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
                                daterasp.Text = "Текущая неделя: " + response;
                                Properties.Settings.Default.number_nedel = response;
                                Properties.Settings.Default.Save();
                                comboBox2.SelectedIndex = 0; // выбираем текущую неделю в comboBox'е
                                if (Properties.Settings.Default.number_nedel.IndexOf("II н") >= 0)
                                {
                                    comboBox2.SelectedIndex = 1;
                                    nedely_rasp = 1; // для передачи на форму "Form608"
                                }
                                break;
                            }
                            else if (reader.EndOfStream)
                                MessageBox.Show("Не могу определить номер недели!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                        reader.Dispose();
                        data.Dispose();
                        client.Dispose();
                        comboBox2.Enabled = true;
                        #endregion
                        #region Загрузка преподавателей
                        TextReader skan_web = new StringReader(web_esstu); // создание считывающего потока
                        response = "";
                        while (response != null)
                        {
                            response = skan_web.ReadLine();
                            if ((response != null) && (response.IndexOf("COLOR=\"#ff00ff\"> ") >= 0))
                                comboBox1.Items.Add(response.Substring(response.IndexOf("COLOR=\"#ff00ff\"> ") + 17, response.Length - response.IndexOf("COLOR=\"#ff00ff\"> ") - 28));
                        }
                        comboBox1.SelectedIndex = 0;
                        comboBox1.Enabled = true;
                        dataGridView1.Enabled = true;
                        SetDoubleBuffered(dataGridView1, true); // включаем быструю перерисовку таблицы
                        graphic608.Enabled = true;
                        web_esstu = ""; // очистка загруженных данных с сайта
                        #endregion
                        #region Ставаим статус соединения - онлайн
                        pictureBox1.Size = status_con.Size;
                        pictureBox1.Image = status_con;
                        pictureBox1.Invalidate(); // перерисовка бокса
                        menu_checkUpdate.Enabled = true;

                        Properties.Settings.Default.starter++; // увеличиваем счетчик запуска программы
                        Properties.Settings.Default.Save(); // обязательно сохраняем настройки в файл (C:\Users\aleks\AppData\Local\Zelenkov_A.V)
                        startCount.Text = "Запусков программы: " + Properties.Settings.Default.starter; // указываем кол-во запусков
                        #endregion
                    }
                }
                catch (WebException) // соединение с сайтом не установлено
                {
                    MessageBox.Show("Нет соединения с расписанием ВСГУТУ.\nРасписание не доступно.", "Ошибка соединения", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            else // если программа уже была запущена ранее и если настройки были изменены
            {
                label1.Text = "Дата обновления: " + Properties.Settings.Default.date_rasp.ToString("dd.MM.yyyy"); // указываем дату обновления расписания
                #region Проверка наличия интернета
                try
                {
                    HttpWebRequest test_link = (HttpWebRequest)WebRequest.Create("https://portal.esstu.ru/raspisan.htm");
                    HttpWebResponse test_response = (HttpWebResponse)test_link.GetResponse();
                    if (HttpStatusCode.OK == test_response.StatusCode)
                    {
                        pictureBox1.Size = status_con.Size;
                        pictureBox1.Image = status_con;
                        pictureBox1.Invalidate();
                        menu_checkUpdate.Enabled = true;
                    }
                }
                catch
                {
                    pictureBox1.Size = status_dis.Size;
                    pictureBox1.Image = status_dis;
                    pictureBox1.Invalidate();
                    menu_checkUpdate.Enabled = false;
                }
                Properties.Settings.Default.starter++;
                Properties.Settings.Default.Save();
                startCount.Text = "Запусков программы: " + Properties.Settings.Default.starter;
                #endregion
                #region Проверяем наличие каталога и файла с данными
                string path_file = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Расписание кафедры");
                string file_rasp = Path.Combine(path_file, Properties.Settings.Default.date_rasp.ToString("dd.MM.yyyy") + ".txt");
                if (!Directory.Exists(path_file) | !File.Exists(file_rasp))
                {
                    MessageBox.Show("Внимание, данные пропали!\n\n" +
                        "Скорей всего они были случайно удалены/повреждены или Вы открываете программу первый раз, но не переживайте.\n" +
                        "Будет произведён автоматический сброс параметров и программа перезапустится.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    Properties.Settings.Default.starter = 0;
                    Properties.Settings.Default.Save();
                    // перезапуск приложения
                    Application.Restart();
                    Environment.Exit(1);
                }
                #endregion
                #region Указываем текущую неделю расписания
                daterasp.Text = "Текущая неделя: " + Properties.Settings.Default.number_nedel;
                comboBox2.SelectedIndex = 0;
                if (Properties.Settings.Default.number_nedel.IndexOf("II н") >= 0)
                {
                    comboBox2.SelectedIndex = 1;
                    nedely_rasp = 1;
                }
                comboBox2.Enabled = true;
                #endregion
                #region Загрузка преподавателей
                string path_file2 = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Расписание кафедры");
                string file_rasp2 = Path.Combine(path_file2, Properties.Settings.Default.date_rasp.ToString("dd.MM.yyyy") + ".txt");
                StreamReader rasp_reader2 = new StreamReader(file_rasp2);

                string response = "";
                while (response != null)
                {
                    response = rasp_reader2.ReadLine();
                    if ((response != null) && (response.IndexOf("COLOR=\"#ff00ff\"> ") >= 0))
                        comboBox1.Items.Add(response.Substring(response.IndexOf("COLOR=\"#ff00ff\"> ") + 17, response.Length - response.IndexOf("COLOR=\"#ff00ff\"> ") - 28));
                }
                comboBox1.SelectedIndex = 0;
                comboBox1.Enabled = true;
                dataGridView1.Enabled = true;
                SetDoubleBuffered(dataGridView1, true); // включаем быструю перерисовку таблицы
                graphic608.Enabled = true;
                #endregion
            }
        }
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            RedrawTable();
        }
        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            RedrawTable();
        }
        private void сформироватьГрафикЗанятийДляПечатиToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Word.Document document = null;
            try
            {
                #region Открываем doc-файл Word из ресурса
                this.Cursor = Cursors.WaitCursor; // показываем ожидающий курсор, мол - "жди, идет запуск" :)
                Word.Application docApp = new Word.Application();
                var fullName = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "608au.doc"); // составляем строку для хранения документа во временной папке
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
                string path_file = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Расписание кафедры");
                string file_rasp = Path.Combine(path_file, Properties.Settings.Default.date_rasp.ToString("dd.MM.yyyy") + ".txt");
                StreamReader rasp_reader = new StreamReader(file_rasp);

                string bufferfile = "";
                string NamePrep;
                int i = 4, q = 1;
                while (bufferfile != null)
                {
                    bufferfile = rasp_reader.ReadLine();
                    if ((bufferfile != null) && (bufferfile.IndexOf("#ff00ff\"> ") >= 0))
                    {
                        // сохраняем сокращенное имя препода в хранилище
                        NamePrep = bufferfile.Substring(bufferfile.IndexOf("COLOR=\"#ff00ff\"> ") + 17, bufferfile.Length - bufferfile.IndexOf("COLOR=\"#ff00ff\"> ") - 28);
                        NamePrep = NamePrep.Remove(1, NamePrep.IndexOf(" ")); // удаляем лишние символы в строке до пробела (без учета первого символа)
                        NamePrep = NamePrep.Replace(".", ""); // удаление оставшегося лишнего символа .
                        bufferfile = rasp_reader.ReadLine();
                        #region 1 неделя
                        while (i <= 44) // пока не достигли конца таблицы - конца недели
                        {
                            while (bufferfile.IndexOf("SIZE=2><P ALIGN=\"CENTER\">") == -1) // тег для поиска дня недели
                                bufferfile = rasp_reader.ReadLine();
                            bufferfile = rasp_reader.ReadLine();
                            while (q <= 6) // счетчик пар
                            {
                                bufferfile = rasp_reader.ReadLine();
                                if ((bufferfile.IndexOf("\">_") == -1) && (bufferfile.IndexOf("а.608а") >= 0))
                                    table.Rows[i].Cells[3].Range.Text = NamePrep + " - " + bufferfile.Substring(bufferfile.IndexOf("ALIGN=\"CENTER\">") + 15, bufferfile.Length - bufferfile.IndexOf("</F") - 2);
                                else
                                if ((bufferfile.IndexOf("\">_") == -1) && (bufferfile.IndexOf("а.608б") >= 0))
                                    table.Rows[i].Cells[4].Range.Text = NamePrep + " - " + bufferfile.Substring(bufferfile.IndexOf("ALIGN=\"CENTER\">") + 15, bufferfile.Length - bufferfile.IndexOf("</F") - 2);
                                bufferfile = rasp_reader.ReadLine();
                                q++; i++;
                            }
                            i++; q = 1;
                            bufferfile = rasp_reader.ReadLine();
                        }
                        i = 4;
                        #endregion
                        #region 2 неделя
                        while (i <= 44)
                        {
                            while (bufferfile.IndexOf("SIZE=2 COLOR=\"#0000ff\"><P ALIGN=\"CENTER\">") == -1)
                                bufferfile = rasp_reader.ReadLine();
                            bufferfile = rasp_reader.ReadLine();
                            while (q <= 6)
                            {
                                bufferfile = rasp_reader.ReadLine();
                                if ((bufferfile.IndexOf("\"> _") == -1) && (bufferfile.IndexOf("а.608а") >= 0))
                                    table.Rows[i].Cells[7].Range.Text = NamePrep + " - " + bufferfile.Substring(bufferfile.IndexOf("ALIGN=\"CENTER\">") + 15, bufferfile.Length - bufferfile.IndexOf("</F") - 2);
                                else
                                if ((bufferfile.IndexOf("\"> _") == -1) && (bufferfile.IndexOf("а.608б") >= 0))
                                    table.Rows[i].Cells[8].Range.Text = NamePrep + " - " + bufferfile.Substring(bufferfile.IndexOf("ALIGN=\"CENTER\">") + 15, bufferfile.Length - bufferfile.IndexOf("</F") - 2);
                                bufferfile = rasp_reader.ReadLine();
                                q++; i++;
                            }
                            i++; q = 1;
                            bufferfile = rasp_reader.ReadLine();
                        }
                        i = 4;
                        #endregion
                    }
                }
                rasp_reader.Dispose();
                #endregion
                this.Cursor = Cursors.Default;
                docApp.Visible = true;
                document = null;
            }
            catch (Exception ex)
            {
                // ошибки выводим
                MessageBox.Show("Ошибка при создании документа: "+ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                document = null;
            }
        }
        private void graphic608_Click(object sender, EventArgs e)
        {
            Classes f608 = new Classes(this.graphic608, this.nedely_rasp); // создаем объект второй формы и передаем ей состояние кнопки формы и день недели
            f608.Show();
            graphic608.Enabled = false; // блокируем кнопку, избавляемся от избыточных нажатий
        }
        private void оПрограммеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Author author_Form = new Author();
            author_Form.ShowDialog();
        }
        private void pictureBox1_MouseEnter(object sender, EventArgs e)
        {
            toolTip1.SetToolTip(pictureBox1, "Отображает состояние подключения, где:\n     * Обычный индикатор - подключен к сайту ВСГУТУ;\n     * Зачёркнутый индикатор - нет подключения к сайту ВСГУТУ."); // всплывающая подсказка
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
            // перерисовываем таблицу с выделением ячейки
            dataGridView1.Invalidate();
        }
        private void menu_updateRaspMenuItem_Click(object sender, EventArgs e)
        {
            if (pictureBox1.Image == status_con)
            {
                var request_caf = (HttpWebRequest)WebRequest.Create("https://portal.esstu.ru/bakalavriat/Caf39.htm");
                request_caf.UserAgent = "Rasp. 608";
                DateTime date;
                using (var get_date = (HttpWebResponse)request_caf.GetResponse())
                    date = get_date.LastModified;
                if (date != Properties.Settings.Default.date_rasp && menu_checkUpdate.Enabled != false)
                {
                    request_caf.Abort();
                    UpdateTable(date);
                }
                else MessageBox.Show("У вас актуальная версия расписания.\nОбновление не требуется.", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
        private void УдалитьРасписаниеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var result_dialog = MessageBox.Show("Вы точно хотите удалить все данные?\n\n" +
                "В этом случае все загруженные расписания будут удалены, выполнется сброс настроек и перезапуск программы.\n\n" +
                "P.S.: старайтесь делать сброс ТОЛЬКО при неправильном отображении расписания в таблицах или при неправильной работе программы.", "Внимание", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            if (result_dialog == DialogResult.Yes)
            {
                // удаление каталога с данными
                string path_file = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Расписание кафедры");
                Directory.Delete(path_file, true);
                dataGridView1.Rows.Clear();
                dataGridView1.Enabled = false;
                // сброс параметров
                Properties.Settings.Default.starter = 0;
                Properties.Settings.Default.Save();
                MessageBox.Show("Сброс программы выполнен.\nВсе данные удалены. Выполняю перезапуск...", "Успешно", MessageBoxButtons.OK, MessageBoxIcon.Information);
                Application.Restart();
                Environment.Exit(1);
            }
        }
        private void timer_updateStatus_Tick(object sender, EventArgs e)
        {
            // Проверяем соединение с сайтом и выводим статус сети
            try
            {
                HttpWebRequest request = (HttpWebRequest)HttpWebRequest.Create("https://portal.esstu.ru/raspisan.htm");
                HttpWebResponse response = (HttpWebResponse)request.GetResponse();
                if (HttpStatusCode.OK == response.StatusCode)
                {
                    response.Close();
                    request.Abort();
                    #region Статус соединения - онлайн
                    if (pictureBox1.Image == status_dis)
                    {
                        pictureBox1.Size = status_con.Size;
                        pictureBox1.Image = status_con;
                        pictureBox1.Invalidate();
                        menu_checkUpdate.Enabled = true; // включаем пункт в меню для работы с обновлением расписания
                    }
                    #endregion
                    #region Обновление расписания
                    var request_caf = (HttpWebRequest)WebRequest.Create("https://portal.esstu.ru/bakalavriat/Caf39.htm");
                    request_caf.UserAgent = "Rasp. 608";
                    DateTime date;
                    using (var get_date = (HttpWebResponse)request_caf.GetResponse())
                        date = get_date.LastModified;
                    if (date != Properties.Settings.Default.date_rasp && menu_checkUpdate.Enabled != false)
                    {
                        request_caf.Abort();
                        UpdateTable(date);
                    }
                    #endregion
                }
            }
            catch (WebException)
            {
                #region Статус соединения - оффлайн
                if (pictureBox1.Image == status_con)
                {
                    pictureBox1.Size = status_dis.Size;
                    pictureBox1.Image = status_dis;
                    pictureBox1.Invalidate();
                    menu_checkUpdate.Enabled = false;
                }
                #endregion
            }
        }
        private void ОПрограммеToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            try
            {
                File.WriteAllBytes("Справка 608au.chm", Properties.Resources.reference);
                Process.Start("Справка 608au.chm");
            }
            catch
            {
                MessageBox.Show("Упс...\nСправка уже открыта!?\nСначала закройте предыдущую справку, затем открывайте новую. ;)", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
        private void Panel5_MouseEnter(object sender, EventArgs e)
        {
            toolTip1.SetToolTip(panel5, "Будет доступно в следующей версии.\nПозволит выбирать сохраненное расписание по датам.");
        }
        private void ОткрытьПапкуСРасписаниемToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string path_file = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Расписание кафедры");
            Process.Start(path_file);
        }
        ///////////////////////////////////////////////////////////////////////////////////////////
        private void UpdateTable(DateTime date)
        {
            menu_checkUpdate.Enabled = false;
            var question = MessageBox.Show("Обнаружено новое расписание.\nХотите обновить текущее расписание?", "Внимание", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            if (question == DialogResult.Yes)
            {
                #region Переводим программу в режим ожидания
                расписаниеToolStripMenuItem.Enabled = false;
                this.Cursor = Cursors.WaitCursor;
                graphic608.Enabled = false;
                comboBox1.Enabled = false;
                comboBox2.Enabled = false;
                comboBox3.Enabled = false;
                dataGridView1.Rows.Clear();
                dataGridView1.Enabled = false;
                // сохранение новой даты расписания
                Properties.Settings.Default.date_rasp = date;
                Properties.Settings.Default.Save();
                #endregion
                try
                {
                    HttpWebRequest test_link_ = (HttpWebRequest)WebRequest.Create("https://portal.esstu.ru/raspisan.htm");
                    HttpWebResponse test_response_ = (HttpWebResponse)test_link_.GetResponse();
                    if (HttpStatusCode.OK == test_response_.StatusCode)
                    {
                        test_response_.Close();
                        test_link_.Abort();
                        #region Проверяем наличие кафедры ЭСППиСХ
                        WebClient CafClient_ = new WebClient();
                        Stream CafStream_ = CafClient_.OpenRead("https://portal.esstu.ru/bakalavriat/craspisanEdt.htm");
                        StreamReader CafReader_ = new StreamReader(CafStream_, System.Text.Encoding.Default);
                        string Caf_ = "";
                        string NomerCaf = "";
                        while (!CafReader_.EndOfStream)
                        {
                            Caf_ = CafReader_.ReadLine();
                            if (Caf_.IndexOf("Электроснабжение промышленных предприятий") >= 0)
                            {
                                NomerCaf = Caf_.Substring(Caf_.IndexOf("<a href=\"") + 9, Caf_.IndexOf("<a href=\""));
                                if (NomerCaf != Properties.Settings.Default.number_caf)
                                {
                                    Properties.Settings.Default.number_caf = NomerCaf;
                                    Properties.Settings.Default.Save();
                                }
                                break;
                            }
                            else if (CafReader_.EndOfStream)
                            {
                                MessageBox.Show("Не обнаружена кафедра ЭСППиСХ.\nОшибка обновления!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }
                        }
                        CafReader_.Close();
                        CafStream_.Close();
                        CafClient_.Dispose();
                        #endregion
                        #region Загрузка нового расписания
                        // отправка запроса на подключение к сайту 
                        HttpWebRequest RaspReq = (HttpWebRequest)WebRequest.Create("https://portal.esstu.ru/bakalavriat/" + Properties.Settings.Default.number_caf + ".htm");
                        HttpWebResponse RaspRes = (HttpWebResponse)RaspReq.GetResponse();
                        Stream RaspStr = RaspRes.GetResponseStream();
                        StreamReader RaspRead = new StreamReader(RaspStr, System.Text.Encoding.Default);
                        // создание каталога и файла расписания
                        string path_file = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Расписание кафедры");
                        if (!Directory.Exists(path_file))
                            Directory.CreateDirectory(path_file);
                        string file_rasp = Path.Combine(path_file, Properties.Settings.Default.date_rasp.ToString("dd.MM.yyyy") + ".txt");
                        StreamWriter raspfile_write = new StreamWriter(file_rasp);
                        // загрузка данных с сайта и сохранение
                        string web_esstu = "";
                        web_esstu = RaspRead.ReadToEnd();
                        raspfile_write.WriteLine(web_esstu);
                        // чистка ресурсов
                        raspfile_write.Dispose();
                        RaspRead.Dispose();
                        RaspStr.Dispose();
                        RaspRes.Close();
                        RaspReq.Abort();
                        #endregion
                        #region Узнаем текущую неделю расписания
                        WebClient client = new WebClient();
                        Stream data = client.OpenRead("https://esstu.ru/index.htm");
                        StreamReader reader = new StreamReader(data);
                        string response2 = "";
                        while (!reader.EndOfStream)
                        {
                            response2 = reader.ReadLine();
                            if (response2.IndexOf("\"header-date\"") >= 0)
                            {
                                response2 = response2.Substring(response2.IndexOf("\"header-date\"") + 14, response2.Length - 47);
                                Properties.Settings.Default.number_nedel = response2;
                                Properties.Settings.Default.Save();
                                break;
                            }
                        }
                        reader.Dispose();
                        data.Dispose();
                        client.Dispose();
                        #endregion
                        #region Завершаем обновление
                        MessageBox.Show("Расписание обновлено.\nПрограмма будет перезапущена.", "Успешно", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        Application.Restart();
                        Environment.Exit(1);
                        #endregion
                    }
                }
                catch (WebException)
                {
                    MessageBox.Show("Пропало соединение с расписанием ВСГУТУ.\nОбновление не завершено, попробуйте позже.", "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            menu_checkUpdate.Enabled = true;
        } // обновление расписания
        private void RedrawTable()
        {
            #region Внесение базового расписания в dataGridView
            dataGridView1.Rows.Clear();
            dataGridView1.RowCount = 0; dataGridView1.RowCount = 7;
            for (int i = 0; i < 7; i++)
                dataGridView1.Rows[0].Cells[i].Style.BackColor = Color.LightSteelBlue;
            for (int i = 1; i < 7; i++)
            {
                dataGridView1.Rows[i].Cells[0].Style.BackColor = Color.LightSteelBlue;
                dataGridView1.Rows[i].MinimumHeight = 60; // минимальная высота строки
            }
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
            var date_time = ""; // хранит время и день недели
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
                                // если время от 0.00 и до 8.59, то пару не выделяем
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
            string path_file = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Расписание кафедры");
            string file_rasp = Path.Combine(path_file, Properties.Settings.Default.date_rasp.ToString("dd.MM.yyyy") + ".txt");
            StreamReader rasp_reader = new StreamReader(file_rasp);

            string raspstr = "";
            while (raspstr != null)
            {
                raspstr = rasp_reader.ReadLine();
                // ищем тег преподавателя в файле
                if ((raspstr != null) && (raspstr.IndexOf("COLOR=\"#ff00ff\"> ") >= 0))
                {
                    // сверяем препода из файла с преподом из comboBox 
                    if ((raspstr.Substring(raspstr.IndexOf("COLOR=\"#ff00ff\"> ") + 17, raspstr.Length - raspstr.IndexOf("COLOR=\"#ff00ff\"> ") - 28) == comboBox1.Text))
                    {
                        // составляем расписание препода, пока не закончилась нужная неделя
                        while ((raspstr != null) && (raspstr.IndexOf("SIZE=2 COLOR=\"#0000ff\"") == -1))
                        {
                            raspstr = rasp_reader.ReadLine();
                            #region 1 неделя
                            // если выбрана 1 неделя
                            if ((raspstr.IndexOf("SIZE=2><P ALIGN=\"CENTER\">") >= 0) && (comboBox2.SelectedItem.ToString() == "1 неделя"))
                            {
                                while ((raspstr != null) && (raspstr != "</TR>"))
                                {
                                    for (int i = 1; i <= 6; i++)
                                    {
                                        while (raspstr.IndexOf("SIZE=2><P ALIGN=\"CENTER\">") == -1) // тег для поиска дня недели
                                            raspstr = rasp_reader.ReadLine();
                                        raspstr = rasp_reader.ReadLine();
                                        for (int j = 1; j != 7; j++)
                                        {
                                            raspstr = rasp_reader.ReadLine();
                                            // меняем цвет ячейки пары, если она идет в 608 ау.
                                            if ((raspstr.IndexOf("а.608") >= 0) && (dataGridView1.Rows[i].Cells[j].Style.BackColor != Color.Yellow))
                                                dataGridView1.Rows[i].Cells[j].Style.BackColor = Color.LightSkyBlue;
                                            // если пар нет
                                            if (raspstr.IndexOf("\">_") >= 0)
                                                dataGridView1.Rows[i].Cells[j].Value = "-";
                                            else // или они есть
                                            {
                                                #region Форматирование текста
                                                raspstr = raspstr.Substring(raspstr.IndexOf("ALIGN=\"CENTER\">") + 15, raspstr.Length - raspstr.IndexOf("ALIGN=\"CENTER\">") - 27);
                                                raspstr = raspstr.Replace("    ", "\n"); // замена пробелов на новую строку
                                                raspstr = raspstr.Replace("...", ""); // замена троеточия на пустой символ
                                                raspstr = raspstr.Trim(); // убираем лишние пробелы в строке
                                                raspstr = raspstr.Replace("пр.", "\n-практика-\n");
                                                raspstr = raspstr.Replace("лаб.", "\n-лабораторная-\n");
                                                raspstr = raspstr.Replace("лек.", "\n-лекция-\n");
                                                if (raspstr.IndexOf("-\n") >= 0) // отсекаем лишний текст после типа пары
                                                    raspstr = raspstr.Substring(0, raspstr.Length - (raspstr.Length - raspstr.IndexOf("-\n") - 1));
                                                #endregion
                                                dataGridView1.Rows[i].Cells[j].Value = raspstr;
                                            }
                                            raspstr = rasp_reader.ReadLine();
                                        }
                                    }
                                    rasp_reader.Dispose();
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
                                            raspstr = rasp_reader.ReadLine();
                                        raspstr = rasp_reader.ReadLine();
                                        for (int j = 1; j != 7; j++)
                                        {
                                            raspstr = rasp_reader.ReadLine();
                                            if ((raspstr.IndexOf("а.608") >= 0) && (dataGridView1.Rows[i].Cells[j].Style.BackColor != Color.Yellow))
                                                dataGridView1.Rows[i].Cells[j].Style.BackColor = Color.LightSkyBlue;
                                            if (raspstr.IndexOf("\"> _") >= 0)
                                                dataGridView1.Rows[i].Cells[j].Value = "-";
                                            else
                                            {
                                                #region Форматирование текста
                                                raspstr = raspstr.Substring(raspstr.IndexOf("ALIGN=\"CENTER\"> ") + 16, raspstr.Length - raspstr.IndexOf("ALIGN=\"CENTER\"> ") - 28);
                                                raspstr = raspstr.Replace("    ", "\n");
                                                raspstr = raspstr.Replace("...", "");
                                                raspstr = raspstr.Trim(); // убираем лишние пробелы в строке
                                                raspstr = raspstr.Replace("пр.", "\n-практика-\n");
                                                raspstr = raspstr.Replace("лаб.", "\n-лабораторная-\n");
                                                raspstr = raspstr.Replace("лек.", "\n-лекция-\n");
                                                if (raspstr.IndexOf("-\n") >= 0) // отсекаем лишний текст после типа пары
                                                    raspstr = raspstr.Substring(0, raspstr.Length - (raspstr.Length - raspstr.IndexOf("-\n") - 1));
                                                #endregion
                                                dataGridView1.Rows[i].Cells[j].Value = raspstr;
                                            }
                                            raspstr = rasp_reader.ReadLine();
                                        }
                                    }
                                    rasp_reader.Dispose();
                                    return;
                                }
                            }
                            #endregion
                        }
                    }
                }
            }
            #endregion
        } // перерисовка таблицы расписания
    }
}

// !Функции!
// Добавить Артефакт и отзыв о программе (форма для заполнения со звездами, пожеланиями и отправкой мне на почту)
// Сделать загрузку программы красивее, подумать о Flat стиле
// Сделать сканирования колледжа правильным (нет типа пары - пр., лек., лаб.)
// Сделать отображение последних изменений в расписании. Чтобы видеть разницу ДО и ПОСЛЕ
// Поправить запуск ворда (место его запуска)
// Подумать над обновление расписания без перезапуска проги
// Прога запускается первый раз с ошибкой