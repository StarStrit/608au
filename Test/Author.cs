using System;
using System.Diagnostics;
using System.Windows.Forms;

namespace Test
{
    public partial class Author : Form
    {
        public Author()
        {
            InitializeComponent();
        }
        private void About_Load(object sender, EventArgs e)
        {
            label3.Text += System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;
        }
        private void linkLabel1_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                // Копируем текст в буфер и выводим подсказку
                Clipboard.SetText(linkLabel1.Text);
                prompt.Show("Текст скопирован.", linkLabel1, e.X * 0, e.Y * 0 + 19);
            }
            if (e.Button == MouseButtons.Left)
            {
                Process.Start("mailto:lito1395@mail.ru");
                linkLabel1.LinkVisited = true;
            }
        }
        private void linkLabel2_MouseClick(object sender, MouseEventArgs e)
        {
            // открываем офиц. страницу Visual Studio 2017
           Process.Start("https://visualstudio.microsoft.com/ru");
            linkLabel2.LinkVisited = true;
        }
        private void linkLabel3_MouseClick(object sender, MouseEventArgs e)
        {
            // открываем офиц. страницу кафедры ЭСППиСХ "ВСГУТУ"
            Process.Start("https://www.esstu.ru/uportal/connector/employee/list.htm?departmentCode=2601");
            linkLabel3.LinkVisited = true;
        }
        private void linkLabel4_MouseClick(object sender, MouseEventArgs e)
        {
            // открываем офиц. страницу "ВСГУТУ"
            Process.Start("https://www.esstu.ru");
            linkLabel4.LinkVisited = true;
        }
        private void LinkLabel5_MouseClick(object sender, MouseEventArgs e)
        {
            // открываем офиц. страницу меня (ВК) :) 
            Process.Start("https://vk.com/anonym011");
            linkLabel5.LinkVisited = true;
        }
        private void PictureBox1_MouseClick(object sender, MouseEventArgs e)
        {
            // открываем офиц. страницу моего сайта :)
            Process.Start("http://testlabsoft.ru.xsph.ru/");
        }
        private void PictureBox1_MouseEnter(object sender, EventArgs e)
        {
            toolTip1.SetToolTip(pictureBox1, "Нажми на меня, чтобы перейти на сайт разработчика :)");
        }
        private void LinkLabel1_MouseEnter(object sender, EventArgs e)
        {
            toolTip1.SetToolTip(linkLabel1, "Нажми правой кнопкой мыши, чтобы скопировать адрес почты");
        }
    }
}