using ParserConsole_2_;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ParserApp
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog fbd = new FolderBrowserDialog();

            //Выводим диалоговое окно для выбора каталога.
            // Данный класс возвращает следующие значения:
            // 1) Объект System.Windows.Forms.DialogResult.OK, 
            //    если пользователь нажимает кнопку 
            //    ОК в диалоговом окне;
            // 2) Объект System.Windows.Forms.DialogResult.Cancel,  
            //    если пользователь закрывает диалоговое окно.
            DialogResult result = fbd.ShowDialog();

            //Если пользователь выбрал директорию
            //и нажал ОК, то выводим путь в textBox1
            if (result == DialogResult.OK)
            {
                //Вывод пути к  
                //выбранной директории 
                textBoxPath.Text = fbd.SelectedPath;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {

            label8.Visible = true;
            label4.Visible = true;
            //button1.UseWaitCursor = true;
            button1.Enabled = false;
            button2.Enabled = false;
            richTextBoxId.Enabled = false;
            richTextBoxIdProducts.Enabled = false;
            naturalIds.Enabled = false;
            textBoxPathLinks.Enabled = false;
            button3.Enabled = false;
            var path = textBoxPath.Text;
            MainClassWithLists.Jewelries = new List<Jewelry>();
            //MainClassWithLists.AddTemp = new List<Jewelry>();
            //Stopwatch stopwatch = new Stopwatch();
            //stopwatch.Start();
            ParserConsole_2_.Parser.artsNatural = ParserConsole_2_.Parser.GetJewelriesWithNatural();
            ParserConsole_2_.Parser.links = ParserConsole_2_.Parser.GetLinks(textBoxPathLinks.Text);
            

            //stopwatch.Stop();
            //stopwatch.Reset();
            //label8.Text = "links + artsNatural" + stopwatch.Elapsed.Minutes.ToString() + " - " + stopwatch.Elapsed.Seconds.ToString();
            //stopwatch.Start();

            ParserConsole_2_.Parser.CreateOrUpdateExcel(path);

            ParserConsole_2_.Parser.Parse();

            var listId = richTextBoxId.Text.Split(',').ToArray();
            if (listId != null)
                MainClassWithLists.DeleteById(listId);

            listId = richTextBoxIdProducts.Text.Split(',').ToArray();

            if (listId != null)
                MainClassWithLists.DeleteByIdProduct(listId);

            var result = Parallel.ForEach(MainClassWithLists.Jewelries.Select(p => p.IdProduct), ParserConsole_2_.Parser.GetMoreInformation);

            listId = naturalIds.Text.Split(',').ToArray();
            if (listId != null)
                ParserConsole_2_.Parser.AddNaturalByStone(listId);

            ParserConsole_2_.Parser.AddNaturalWord();
            ParserConsole_2_.Parser.AddCells(MainClassWithLists.Jewelries);
            ParserConsole_2_.Parser.EndExcel();
            //stopwatch.Stop();

            //label8.Text += "\nExcel done" + stopwatch.Elapsed.Hours.ToString() + stopwatch.Elapsed.Minutes.ToString() + " - " + stopwatch.Elapsed.Seconds.ToString();
            //button1.UseWaitCursor = false;
            button1.Enabled = true;
            button2.Enabled = true;
            richTextBoxId.Enabled = true;
            richTextBoxIdProducts.Enabled = true;
            naturalIds.Enabled = true;
            textBoxPathLinks.Enabled = true;
            button3.Enabled = true;
            label4.Visible = false;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            //Stopwatch stopwatch = new Stopwatch();
            //stopwatch.Start();

            label8.Visible = true;
            label4.Visible = true;
            //button1.UseWaitCursor = true;
            button1.Enabled = false;
            button2.Enabled = false;
            richTextBoxId.Enabled = false;
            richTextBoxIdProducts.Enabled = false;
            naturalIds.Enabled = false;
            textBoxPathLinks.Enabled = false;
            button3.Enabled = false;
            ParserConsole_2_.Parser.links = ParserConsole_2_.Parser.GetLinks(textBoxPathLinks.Text);
            var result = Parallel.ForEach(MainClassWithLists.Jewelries.Select(p => p.IdProduct), ParserConsole_2_.Parser.GetDescription);

            //button1.UseWaitCursor = false;
            button1.Enabled = true;
            button2.Enabled = true;
            richTextBoxId.Enabled = true;
            richTextBoxIdProducts.Enabled = true;
            naturalIds.Enabled = true;
            textBoxPathLinks.Enabled = true;
            button3.Enabled = true;
            label4.Visible = false;
            label8.Visible = false;
            //stopwatch.Stop();
            //label8.Text += "\nAdd descriptions" + stopwatch.Elapsed.Hours.ToString() + stopwatch.Elapsed.Minutes.ToString() + " - " + stopwatch.Elapsed.Seconds.ToString();
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
           
        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            Dispose(true);
            Environment.Exit(Environment.ExitCode);
            Application.Exit();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            Dispose(true);
            Environment.Exit(Environment.ExitCode);
            MainClassWithLists.Jewelries = new List<Jewelry>();
        }
    }
}
