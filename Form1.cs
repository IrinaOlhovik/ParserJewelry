using ParserConsole_2_;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ParserApp
{
    public partial class Form1 : Form
    {
        public static int index = 0;
        public static Label label10 = new Label();
        public Form1()
        {
            InitializeComponent();
            label10.Size = new Size(59, 13);
            label10.Location = new Point(422, 397);
            label10.TabIndex = 23;
            label10.Margin = new Padding(3, 0, 3, 0);
            label10.Text = "Text";
            label10.Visible = true;
            Controls.Add(label10);
            backgroundWorker1 = new BackgroundWorker();
            backgroundWorker1.WorkerReportsProgress = true;
            backgroundWorker1.WorkerSupportsCancellation = true;
            backgroundWorker1.DoWork += new DoWorkEventHandler(backgroundWorker1_DoWork);
            backgroundWorker1.ProgressChanged += new ProgressChangedEventHandler(backgroundWorker1_ProgressChanged);
            backgroundWorker1.RunWorkerCompleted += new RunWorkerCompletedEventHandler(backgroundWorker1_RunWorkerCompleted);
        }
        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {

        }
        private void backgroundWorker1_DoWork(object sender, System.ComponentModel.DoWorkEventArgs e)
        {
            var bw = sender as BackgroundWorker;
            bool work = true;
            while(work)
            {
                Thread.Sleep(1000);
                var i = ParserConsole_2_.Parser.indexAdd;
                var total = MainClassWithLists.Jewelries.Count;
                if (i == total)
                {
                    work = false;
                    bw.CancelAsync();
                }
                lock(label10)
                {
                    bw.ReportProgress(1, i);
                }

            }
        }
        private void backgroundWorker1_ProgressChanged(object sender, System.ComponentModel.ProgressChangedEventArgs e)
        {
            progressBar1.Value = int.Parse(e.UserState.ToString());
        }
        private void button1_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            DialogResult result = fbd.ShowDialog();
            if (result == DialogResult.OK)
            {
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
           // MainClassWithLists.AddTemp = new List<Jewelry>();
            Stopwatch stopwatch = new Stopwatch();
            stopwatch.Start();
            ParserConsole_2_.Parser.artsNatural = new List<string>(); //ParserConsole_2_.Parser.GetJewelriesWithNatural();
            ParserConsole_2_.Parser.links = ParserConsole_2_.Parser.GetLinks(textBoxPathLinks.Text);
            stopwatch.Stop();


            label8.Text = "links + artsNatural" + stopwatch.Elapsed.Minutes.ToString() + " - " + stopwatch.Elapsed.Seconds.ToString();
            stopwatch.Reset();
            stopwatch.Start();

            ParserConsole_2_.Parser.CreateOrUpdateExcel(path);

            ParserConsole_2_.Parser.Parse();

            var listId = richTextBoxId.Text.Split(',').ToArray();
            if (listId != null)
                MainClassWithLists.DeleteById(listId);

            listId = richTextBoxIdProducts.Text.Split(',').ToArray();

            if (listId != null)
                MainClassWithLists.DeleteByIdProduct(listId);

            progressBar1.Maximum = MainClassWithLists.Jewelries.Count;
            backgroundWorker1.RunWorkerAsync();

            var result = Parallel.ForEach(MainClassWithLists.Jewelries.Select(p => p.IdProduct),
                    ParserConsole_2_.Parser.GetMoreInformation);

           
            listId = naturalIds.Text.Split(',').ToArray();
            if (listId != null)
                ParserConsole_2_.Parser.AddNaturalByStone(listId);

            ParserConsole_2_.Parser.AddNaturalWord();
            ParserConsole_2_.Parser.AddCells(MainClassWithLists.Jewelries);
            ParserConsole_2_.Parser.EndExcel();
            //stopwatch.Stop();

            label8.Text += "\nExcel done" + stopwatch.Elapsed.Hours.ToString() + stopwatch.Elapsed.Minutes.ToString() + " - " + stopwatch.Elapsed.Seconds.ToString();
            //button1.UseWaitCursor = false;
            button1.Enabled = true;
            naturalIds.Enabled = true;
            richTextBoxIdProducts.Enabled = true;
            textBoxPathLinks.Enabled = true;
            richTextBoxId.Enabled = true;
            button2.Enabled = true;
            button3.Enabled = true;
            label4.Visible = false;
        }
        private void button3_Click(object sender, EventArgs e)
        {
            Stopwatch stopwatch = new Stopwatch();
            
            stopwatch.Start();
            label10.Visible = true;
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
            
            var result = Parallel.ForEach(MainClassWithLists.Jewelries.Select(p => p.IdProduct),i => {
  
                ParserConsole_2_.Parser.GetDescription(i);
                });
            
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
            stopwatch.Stop();
            label8.Text += "\nAdd descriptions" + stopwatch.Elapsed.Hours.ToString() + stopwatch.Elapsed.Minutes.ToString() + " - " + stopwatch.Elapsed.Seconds.ToString();
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
        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            var txtNumbers = richTextBoxIdProducts.Text;
            using (TextWriter textWriter = new StreamWriter("fileaidProducts.txt"))
            {
                textWriter.WriteLine(txtNumbers);

            }
            txtNumbers = richTextBoxId.Text;
            using (TextWriter textWriter = new StreamWriter("fileId.txt"))
            {
                textWriter.WriteLine(txtNumbers);
            }
            txtNumbers = naturalIds.Text;
            using (TextWriter textWriter = new StreamWriter("fileStones.txt", false, Encoding.ASCII))
            {
                textWriter.WriteLine(txtNumbers);
            }
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            string text = "";
            using (TextReader textReader = new StreamReader("fileId.txt"))
            {
                text = textReader.ReadToEnd();
            }
            richTextBoxId.Text = text;
            using (TextReader textReader = new StreamReader("fileaidProducts.txt"))
            {
                text = textReader.ReadToEnd();
            }
            richTextBoxIdProducts.Text = text;
            using (TextReader textReader = new StreamReader("fileStones.txt",Encoding.ASCII))
            {
                text = textReader.ReadToEnd();
            }
            naturalIds.Text = text;
        }
    }
}
