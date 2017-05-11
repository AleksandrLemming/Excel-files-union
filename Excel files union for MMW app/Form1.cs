using System;
using System.IO;
using System.Windows.Forms;

namespace Excel_files_union_for_MMW_app
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        Excel excel = null;

        private void Form1_Load(object sender, EventArgs e)
        {
            this.excel = new Excel();
            listBox1.HorizontalScrollbar = true;
            listBox1.ScrollAlwaysVisible = true;
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            listBox1.Items.Clear(); //очищаем листбокс

            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK) //если папка выбрана, то выводим список файлов в листбокс и загоняем их в массив
            {
                excel.PathFolder(folderBrowserDialog1.SelectedPath); //запоминаем путь к папке

                excel.filelists = Directory.GetFiles(folderBrowserDialog1.SelectedPath); //список загоняем в массив

                foreach (var item in excel.filelists) //заполняем листбокс списком файлов из выбранной папки
                {
                    listBox1.Items.Add(item);
                }
            }

        }

        private void Button2_Click(object sender, EventArgs e)
        {
            if (listBox1.Items.Count > 0) //если папка пуста или не выбрана, то выдаст сообщение
            {
                WindowState = FormWindowState.Minimized;

                excel.UnionFiles();

                WindowState = FormWindowState.Normal;
                MessageBox.Show("Все готово!");
                listBox1.Items.Clear();
            }
            else
            {
                MessageBox.Show("Вы не выбрали папку или в папке нету файлов!");
            }
        }

        private void NumericUpDown1_ValueChanged(object sender, EventArgs e)
        {
            excel.rowsinheader = Convert.ToInt32(numericUpDown1.Value);
        }

        private void NumericUpDown2_ValueChanged(object sender, EventArgs e)
        {
            excel.sheetnumber = Convert.ToInt32(numericUpDown1.Value);
        }
    }
}
