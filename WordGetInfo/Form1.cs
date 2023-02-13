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
using System.Configuration;
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
                button_Search.Enabled = true;
            }
        }

        private void button_Search_Click(object sender, EventArgs e)
        {
            string filePath = txt_FilePath.Text;
            string strnumero = "";
            string strprecio = "";
            float precio = 0;

            try
            {
                string text = ExtractTextFromWord(filePath);
                string dataobtain = ExtractData(text, ref strnumero, ref strprecio);
                label_deednumber.Text = strnumero;
                label_price.Text = strprecio;

                label_deednumber.Visible = true;
                label_price.Visible = true;


                precio = float.Parse(strprecio.Replace(",", ""));
                if (precio > 1659840)
                    MessageBox.Show("Dar aviso al SAT", "SAT aviso",
                        MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                else
                    MessageBox.Show("No es necesario dar aviso", "SAT aviso?",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message, "Error al leer el archivo",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        static string ExtractTextFromWord(string filePath)
        {
            try
            {
                using (DocX document = DocX.Load(filePath))
                {
                    string text = string.Join("\n", document.Paragraphs.Select(p => p.Text));
                    return text;
                }
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }

        static string ExtractData(string text, ref string numero ,ref string precio)
        {
            Match matchNumero = Regex.Match(text, "Escritura Pública Número (.*?)-");
            if (matchNumero.Success)
            {
                numero = matchNumero.Groups[1].Value;
            }
            Match matchPrecio = Regex.Match(text, "Precio de la Operación.*?\\$(.*?) ");
            if (matchPrecio.Success)
            {
                precio = matchPrecio.Groups[1].Value;
            }
            if (matchNumero.Success && matchPrecio.Success)
            {
                return "True";
            }
            else
            {
                return null;
            }
        }
    }
}
