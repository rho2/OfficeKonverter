using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;



namespace ConvertereForms
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            //Einfügen des aktuellen Datum zu Beginn des Logs
            textBox2.Text += DateTime.Now.ToLongDateString() + " " + DateTime.Now.ToLongTimeString() + "\n";
            textBox2.Text += "---------------------------------------------------------------------------\n";
        }
        #region Deklarationen
        //String für den Pfad
        string srcPfad = "";
        //Listen in denen die Pfade der Dokumente absgespeichert werden
        List<string> dateienWord = new List<string>();
        List<string> dateienExcel = new List<string>();
        //Count, Anzahl der konvertieren Dokumente 
        int count = 0;
        #endregion


        /// <summary>
        /// Erzwingt ein updaten der Textbox (nicht schön, aber funktioniert)
        /// </summary>
        private void upText()
        {
            textBox2.Invalidate();
            textBox2.Update();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            //Eingabe des Pfades
            folderBrowserDialog1.ShowDialog();
            srcPfad = folderBrowserDialog1.SelectedPath;
            textBox1.Text = srcPfad;

            //abspeichern alles Dateien
            string[] FileNames = Directory.GetFiles(srcPfad, "*.*", SearchOption.AllDirectories);

            //geht alle Dateien durch
            foreach (string s in FileNames)
            {
                //Wählt die Dateiendung
                switch (Path.GetExtension(s))
                {
                    case ".doc":
                        dateienWord.Add(s);
                        textBox2.Text += s + " gefunden \n";
                        upText();
                        break;
                    case ".xls":
                        dateienExcel.Add(s);
                        textBox2.Text += s + " gefunden \n";
                        upText();
                        break;
 
                    default:
                        break;
                }
            }
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            //Word
            if (checkBox1.Checked)
            {
                convertWordFiles();
                //Remove
                if (checkBox3.Checked)
                {
                    foreach (string s in dateienWord)
                    {
                        File.Delete(s);
                        textBox2.Text += s + " gelöscht\n";
                        upText();
                    }
                }
            }
            //Excel
            if (checkBox2.Checked)
            {
                convertExcelFiles();
                //Remove
                if (checkBox3.Checked)
                {
                    foreach (string s in dateienExcel)
                    {
                        File.Delete(s);
                        textBox2.Text += s + " gelöscht\n";
                        upText();
                    }
                }
            }
            //Abspeichern der Log-Datei
            textBox2.Text += count + " Datei(en) erfolgreich konvertiert";
            File.WriteAllLines(srcPfad + @"\log.txt", textBox2.Lines);
        }

        /// <summary>
        /// Konvertiert alle Dateien in dateienWord
        /// </summary>
        private void convertWordFiles()
        {
            //Erzeugt eine neue Word-Anwendung
            Microsoft.Office.Interop.Word.Application word = new Microsoft.Office.Interop.Word.Application();
            foreach (string s in dateienWord)
            {
                //Einlesen des Dokumentes
                var sourceFile = new FileInfo(s);
                var doc = word.Documents.Open(sourceFile.FullName);
                string newFile = "";

                //Prüfen auf Makro
                if (doc.HasVBProject)
                {
                    //Abspeichern des Dokumentes
                    newFile = sourceFile.FullName.Replace(".doc", ".docm");
                    doc.SaveAs2(FileName: newFile, FileFormat: WdSaveFormat.wdFormatXMLDocumentMacroEnabled, CompatibilityMode: Microsoft.Office.Interop.Word.WdCompatibilityMode.wdWord2003);

                }
                else
                {
                    //Abspeichern des Dokumentes
                    newFile = sourceFile.FullName.Replace(".doc", ".docx");
                    doc.SaveAs2(FileName: newFile, FileFormat: WdSaveFormat.wdFormatXMLDocument, CompatibilityMode: Microsoft.Office.Interop.Word.WdCompatibilityMode.wdWord2003);
                }
                
                //schließt das Dokument
                word.ActiveDocument.Close();

                //Zeit anpassen
                File.SetCreationTime(newFile, File.GetCreationTime(s));
                File.SetLastWriteTime(newFile, File.GetLastWriteTime(s));
                File.SetLastAccessTime(newFile, File.GetLastAccessTime(s));

                //Log
                textBox2.Text += s + " konvertiert\n";
                upText();
                count++;
            }
            //Beendet Word
            word.Quit();
        }

        /// <summary>
        /// Konvertiert alle Dokumente in dateinExcel
        /// </summary>
        private void convertExcelFiles()
        {
            //Erzeugt eine neue Excel-Anwendung
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();

            foreach (string s in dateienExcel)
            {
                //Einlesen des Dokumentes
                var sourceFile = new FileInfo(s);
                var doc = excel.Workbooks.Open(sourceFile.FullName);

                string newFile = "";

                //Prüfen auf Makro
                if (doc.HasVBProject)
                {
                    //Abspeichern des Dokumentes
                    newFile = sourceFile.FullName.Replace(".xls", ".xlsm");
                    doc.SaveAs(Filename: newFile, FileFormat: Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbookMacroEnabled);

                }
                else
                {
                    //Abspeichern des Dokumentes
                    newFile = sourceFile.FullName.Replace(".xls", ".xlsx");
                    doc.SaveAs(Filename: newFile, FileFormat: Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook);
                }
                //schließt das Dokument
                excel.ActiveWorkbook.Close();

                //Zeit anpassen
                File.SetCreationTime(newFile, File.GetCreationTime(s));
                File.SetLastWriteTime(newFile, File.GetLastWriteTime(s));
                File.SetLastAccessTime(newFile, File.GetLastAccessTime(s));

                //Log
                textBox2.Text += s + " konvertiert\n";
                upText();
                count++;
            }
            //Beendet Excel
            excel.Quit();
        }
    }
}
