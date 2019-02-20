using System;
using System.Windows.Forms;

namespace Test
{
    public partial class About : Form
    {
        public About()
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
                prompt.Show("Текст помещен в буфер обмена.", linkLabel1, e.X * 0, e.Y * 0 + 19);
            }
            if (e.Button == MouseButtons.Left)
            {
                System.Diagnostics.Process.Start("mailto:lito1395@mail.ru");
                linkLabel1.LinkVisited = true;
            }
        }
        private void linkLabel2_MouseClick(object sender, MouseEventArgs e)
        {
            // открываем офиц. страницу Visual Studio 2017
            System.Diagnostics.Process.Start("https://visualstudio.microsoft.com/ru");
            linkLabel2.LinkVisited = true;
        }
        private void linkLabel3_MouseClick(object sender, MouseEventArgs e)
        {
            // открываем офиц. страницу кафедры ЭСППиСХ "ВСГУТУ"
            System.Diagnostics.Process.Start("https://www.esstu.ru/uportal/connector/employee/list.htm?departmentCode=2601");
            linkLabel3.LinkVisited = true;
        }
        private void linkLabel4_MouseClick(object sender, MouseEventArgs e)
        {
            // открываем офиц. страницу "ВСГУТУ"
            System.Diagnostics.Process.Start("https://www.esstu.ru");
            linkLabel4.LinkVisited = true;
        }
    }
}