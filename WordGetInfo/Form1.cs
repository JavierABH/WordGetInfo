using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Text.RegularExpressions;
using Xceed.Words.NET;



namespace WordGetInfo
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button_ImportFile_Click(object sender, EventArgs e)
        {
            // Select file
            OpenFileDialog dlFileCSV = new OpenFileDialog();
            dlFileCSV.Title = "Seleccionar el archivo Word...";
            dlFileCSV.Filter = "Archivos Word (*.docx,*.doc)|*.docx;*.doc|" +
                "Todos los archivos (*.*)|*.*";
            dlFileCSV.FilterIndex = 1;
            dlFileCSV.RestoreDirectory = true;
            if (dlFileCSV.ShowDialog() == DialogResult.OK)
            {
                txt_FilePath.Text = dlFileCSV.FileName;
            }
        }

        private void button_Search_Click(object sender, EventArgs e)
        {
            string filePath = txt_FilePath.Text;
            string text = ExtractTextFromWord(filePath);
            string search = ExtractNumber(text, ref 2);
            label_deednumber.Text = search;

            label_deednumber.Visible = true;
            label_price.Visible = true;

        }

        static string ExtractTextFromWord(string filePath)
        {
            using (DocX document = DocX.Load(filePath))
            {
                string text = string.Join("\n", document.Paragraphs.Select(p => p.Text));
                return text;
            }
        }

        static string ExtractNumber(string text, ref string numero ,ref string precio)
        {
            Match match = Regex.Match(text, "Escritura Pública Número (.*?)-");
            if (match.Success)
            {

                return match.Groups[1].Value;
            }
            else
            {
                return null;
            }
        }
    }
}
