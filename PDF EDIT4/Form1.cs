using DevExpress.Emf;
using DevExpress.Utils.CommonDialogs;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.codec;
using Image = iTextSharp.text.Image;


namespace PDF_EDIT4
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        string path_pdf_input;
        string path_pdf_output;
        string path_pdf_input_Merge;
        string path_pdf_output_merge;
        private static iText.Kernel.Pdf.PdfObject filterArray;

        private void simpleButton1_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.InitialDirectory = @"C:\";
            openFileDialog1.RestoreDirectory = true;
            openFileDialog1.Title = "Browse PDF Files";
            openFileDialog1.DefaultExt = "PDF";
            openFileDialog1.Filter = "pdf files (*.pdf)|*.pdf|All files (*.*)|*.*";
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                path_pdf_input = openFileDialog1.FileName;
                txt_Path_Input_pdf.Text = path_pdf_input;
            }
        }

        private void ExtractPages(string inputFile, string outputFile, int startPage, int endPage)
        {
            using (FileStream stream = new FileStream(outputFile, FileMode.Create))
            {
                Document document = new Document();
                PdfCopy pdf = new PdfCopy(document, stream);
                document.Open();

                using (iTextSharp.text.pdf.PdfReader reader = new iTextSharp.text.pdf.PdfReader(inputFile))
                {
                    for (int i = startPage; i <= endPage; i++)
                    {
                        pdf.AddPage(pdf.GetImportedPage(reader, i));
                    }
                }

                document.Close();
            }
        }

        private void simpleButton2_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog saveFileDialog = new FolderBrowserDialog();
            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                path_pdf_output = saveFileDialog.SelectedPath;
                txt_Path_output_pdf.Text = path_pdf_output;
            }
        }

        private void simpleButton3_Click(object sender, EventArgs e)
        {
            string[] araQB = RichTxt_QB.Text.Split('\n');
            string[] araname = RichTxt_Name.Text.Split('\n');
            string[] arapagecount = RichTxt_Page_Count.Text.Split('\n');
            if (txt_Path_Input_pdf.Text == "") { MessageBox.Show("يعني ما مبين عليك انو المسار تبع الملف فاضي"); }
            else if (txt_Path_output_pdf.Text == "")
            {
                MessageBox.Show("إذا ما فيا ازعاج يعني المسار تبع المجلد تعبلينا ياه");
            }
            else
            {
                try
                {
                    string inputFile = txt_Path_Input_pdf.Text;
                    int startpage = 0;
                    int endPage = 0;
                    for (int i = 0; i < araQB.Length - 1; i++)
                    {
                        string outputFile = txt_Path_output_pdf.Text + @"\" + araQB[i] + " " + araname[i] + ".pdf";
                        startpage = 1 + endPage;
                        endPage = startpage + Convert.ToInt32(arapagecount[i]) - 1;
                        ExtractPages(inputFile, outputFile, startpage, endPage);
                    }
                }
                catch (Exception)
                {
                }

            }



        }

        private void simpleButton5_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog saveFileDialog = new FolderBrowserDialog();
            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                path_pdf_input_Merge = saveFileDialog.SelectedPath;
                textEdit1.Text = path_pdf_input_Merge;
            }
        }

        private void simpleButton4_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog saveFileDialog = new FolderBrowserDialog();
            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                path_pdf_output_merge = saveFileDialog.SelectedPath;
                textEdit2.Text = path_pdf_output_merge;
            }
        }

        private void simpleButton6_Click(object sender, EventArgs e)
        {
            string[] araQB = txt_QB_Mereg.Text.Split('\n');
            string[] araname = txt_Name_Mereg.Text.Split('\n');
            string[] araZoho = txt_Zoho_Mereg.Text.Split('\n');
            string[] Fin;
            string[] HR;
            if (textEdit1.Text == "") { MessageBox.Show("يعني ما مبين عليك انو المسار تبع الملف فاضي"); }
            else if (textEdit2.Text == "")
            {
                MessageBox.Show("إذا ما فيا ازعاج يعني المسار تبع المجلد تعبلينا ياه");
            }
            else
            {
                try
                {
                    string path = textEdit1.Text;
                    string excelFilePath = textEdit2.Text;
                    for (int i = 0; i < araQB.Length - 1; i++)
                    {
                        
                        Fin = FoundFile(path, araQB[i]);
                        HR = FoundFile(path, "Z"+araZoho[i]+" ");
                        int totalLength = Fin.Length + HR.Length;
                        string[] mergedArray = new string[totalLength];
                        mergedArray = Fin.Concat(HR).ToArray();
                        string outputFile = textEdit2.Text + @"\" + araQB[i] + " " + araname[i] + ".pdf";
                        MergePDFs(mergedArray, outputFile);
                        
                    }
                }

                catch (Exception)
                {
                }
                MessageBox.Show("Done");
            }
        }

        private string[] FoundFile(string folderPath, string fileNameToFind)
        {

            if (Directory.Exists(folderPath))
            {
                // البحث عن الملفات داخل المجلد
                string[] files = Directory.GetFiles(folderPath)
                    .Where(filePath => Path.GetFileName(filePath).Contains(fileNameToFind))
                    .ToArray();
                if (files.Length > 0)
                {
                    Console.WriteLine("The following files were found:");
                    foreach (string file in files)
                    {
                        Console.WriteLine(file);
                        return files;
                    }
                }
                else
                {
                    Console.WriteLine("file not found.");
                    return files;
                }
                return files;
            }
            else
            {
                Console.WriteLine("folder not found.");
                return null;
            }
        }

        private void FileToMereg(string[] pdfFile, string name)
        {
            // قائمة بأسماء ملفات PDF التي تريد دمجها
            string[] pdfFiles = pdfFile;

            // اسم الملف المخرج النهائي
            string outputPdfFile = "merged.pdf";

            // دمج الملفات
            MergePDFs(pdfFiles, outputPdfFile);

            Console.WriteLine("تم دمج الملفات بنجاح.");
        }

        private void MergePDFs(string[] inputFiles, string outputFile)
        {
            // إعداد مستند PDF جديد للدمج
            Document doc = new Document();
            PdfCopy copy = new PdfCopy(doc, new FileStream(outputFile, FileMode.Create));
            doc.Open();

            foreach (string inputFile in inputFiles)
            {
                // فتح الملف الحالي للقراءة
                iTextSharp.text.pdf.PdfReader reader = new iTextSharp.text.pdf.PdfReader(inputFile);

                for (int i = 1; i <= reader.NumberOfPages; i++)
                {
                    // استخراج صفحة واحدة من الملف الحالي
                    PdfImportedPage page = copy.GetImportedPage(reader, i);
                    copy.AddPage(page);
                }

                reader.Close();
            }

            // إغلاق المستند والحفظ
            doc.Close();
        }

        private void simpleButton8_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.InitialDirectory = @"C:\";
            openFileDialog1.RestoreDirectory = true;
            openFileDialog1.Title = "Browse PDF Files";
            openFileDialog1.DefaultExt = "PDF";
            openFileDialog1.Filter = "pdf files (*.pdf)|*.pdf|All files (*.*)|*.*";
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                path_pdf_input = openFileDialog1.FileName;
                textEdit3.Text = path_pdf_input;
            }
        }

        private void simpleButton7_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog saveFileDialog = new FolderBrowserDialog();
            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                path_pdf_output = saveFileDialog.SelectedPath;
                textEdit4.Text = path_pdf_output + @"\compression.Pdf";
            }
        }


        private void simpleButton9_Click(object sender, EventArgs e)
        {
                
        
        }
    }
} 

