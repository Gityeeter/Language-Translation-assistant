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
using System.Collections;
using Microsoft.Office.Interop.Excel;
using System.IO;
namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        Dictionary<string, string> names =
        new Dictionary<string, string>();
        public Form1()
        {
            //Example Common japanese names in a game
            InitializeComponent();
            names.Add("瑞羽", "Mizuha");
            names.Add("佐野コーチ", "Coach Sano");
            names.Add("アリサ", "Alyssa");
            names.Add("雛多", "Hinata");
            names.Add("ベスリー","Bethly");
            names.Add("椛", "Momiji");
            names.Add("雪月", "Yuzuki");
            names.Add("優子", "Yuko");
            names.Add("恒一", "Koichi");
            names.Add("百々花", "Momoka");
            names.Add("まりあ", "Maria");
        }
        private  void button3_Click(object sender, EventArgs e)
        {
            System.IO.Stream myStream = null;
            OpenFileDialog theDialog = new OpenFileDialog();
            theDialog.Title = "Open Excel File";
            theDialog.Filter = "xlsx files|*.xlsx";
            theDialog.InitialDirectory = @"C:\";
            if (theDialog.ShowDialog() == DialogResult.OK)
            {
                string sFileName = theDialog.FileName;

                if ((myStream = theDialog.OpenFile()) != null)
                {
                    var excel = new Microsoft.Office.Interop.Excel.Application();
                    Workbook xlWorkBook = excel.Workbooks.Open(sFileName);
                    var xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
                    Microsoft.Office.Interop.Excel.Range xlRange = xlWorkSheet.UsedRange;
                  //  ArrayList arlist = new ArrayList();
                    //MessageBox.Show(temp);
                    var rCnt = xlRange.Rows.Count;
                    var cCnt = xlRange.Columns.Count;
                  //  string text = "C:\\Users\\The birb king\\Documents\\actors.txt";
                  //  TextWriter writer = File.CreateText(text);
                    System.Console.WriteLine(rCnt);                  
                    using (System.IO.StreamWriter file = new System.IO.StreamWriter("C:\\Users\\The birb king\\Documents\\output_textfile.txt"))
                    {
                        for (int i = 0; i < rCnt; i++)
                        {

                            string temp = (string)(xlRange.Cells[i + 1, 2] as Microsoft.Office.Interop.Excel.Range).Value2;
                            System.Console.WriteLine("Writing line" + i);
                            System.Console.WriteLine(temp);
                            file.WriteLine(temp);
                        }
                    }
                    //   string[] funny = (string[])arlist.ToArray();
                    //    File.WriteAllLines("WriteLines.txt", funny);
                    xlWorkBook.Close();
                    excel.Quit();
                    MessageBox.Show("Finished writing. File saved in My documents as output_textfile.txt");
                }
            }
        }
        private void button2_Click(object sender, EventArgs e)
        {
            System.IO.Stream myStream = null;
            OpenFileDialog theDialog = new OpenFileDialog();
            theDialog.Title = "Open Text File";
            theDialog.Filter = "TXT files|*.txt";
            theDialog.InitialDirectory = @"C:\";
            if (theDialog.ShowDialog() == DialogResult.OK)
            {
                string sFileName = theDialog.FileName;

                if ((myStream = theDialog.OpenFile()) != null)
                {
                    {
                        string[] lines = System.IO.File.ReadAllLines(sFileName);
                        var excel = new Microsoft.Office.Interop.Excel.Application();
                        if (excel == null)
                        {
                            MessageBox.Show("Excel is not properly installed!!");
                            return;
                        }
                        else
                        {
                            //   MessageBox.Show("Excel is installed!!"); 
                        }

                        excel.Visible = false;
                        excel.DisplayAlerts = false;
                        var xlWorkBook = excel.Workbooks.Add();
                        var xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
                        // Use a tab to indent each line of the file.
                        for (int i = 0; i < lines.Length; i++)
                        {
                            if(names.ContainsKey(lines[i]))
                            {
                                string placeholder;
                                names.TryGetValue(lines[i], out placeholder);
                                xlWorkSheet.Cells[i+1, 2] = placeholder;
                            }
                        }
                        // xlWorkSheet.Cells[1, 1] = para;
                        xlWorkBook.SaveAs("Names.xlsx");
                        xlWorkBook.Close();
                        excel.Quit();
                        MessageBox.Show("Names have been saved in Names.xlsx in my Documents.");



                    }
                }
            }
        }
            private void button1_Click(object sender, EventArgs e)
        {

            System.IO.Stream myStream = null;
            OpenFileDialog theDialog = new OpenFileDialog();
            theDialog.Title = "Open Text File";
            theDialog.Filter = "TXT files|*.txt";
            theDialog.InitialDirectory = @"C:\";
            if (theDialog.ShowDialog() == DialogResult.OK)
            {
                string sFileName = theDialog.FileName;
                
                    if ((myStream = theDialog.OpenFile()) != null)
                    {          
                        {
                            string[] lines = System.IO.File.ReadAllLines(sFileName);
                            var excel = new Microsoft.Office.Interop.Excel.Application();
                            if (excel == null)
                            {
                                MessageBox.Show("Excel is not properly installed!!");
                                return;
                            }
                            else
                            {
                                //   MessageBox.Show("Excel is installed!!"); 
                            }

                            excel.Visible = false;
                            excel.DisplayAlerts = false;
                            var xlWorkBook = excel.Workbooks.Add();
                            var xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
                            // Use a tab to indent each line of the file.
                            for (int i = 0; i < lines.Length; i++)
                            {   
                                if (lines[i].StartsWith("「") && (lines[i].EndsWith("」")))
                                {
                                    xlWorkSheet.Cells[i + 1, 1] = lines[i];
                                System.Console.WriteLine("BOth");
                                }
                                else if (lines[i].StartsWith("「")) 
                                {
                                System.Console.WriteLine("Single");
                                System.Console.WriteLine("lines.Length " + lines.Length);
                                System.Console.WriteLine("i " + i);
                                String concat = lines[i]; 
                                    int j = i + 1;
                                    bool end_detector = false;
                                    int while_ran = 0;
                                    while( end_detector == false)
                                    {
                                    System.Console.WriteLine("j " + j);
                                    concat = concat + lines[j];
                                        if (lines[j].Contains("」"))
                                        {
                                            end_detector = true;
                                        }
                                        j++;
                                        while_ran++;
                                    }
                                    xlWorkSheet.Cells[i + 1, 1] = concat;
                                    i = i + while_ran;
                                }
                                else
                                {
                                   xlWorkSheet.Cells[i + 1, 1] = lines[i];
                                }
                                
                            }
                           // xlWorkSheet.Cells[1, 1] = para;
                            xlWorkBook.SaveAs("text_to_excel.xlsx");
                            xlWorkBook.Close();
                            excel.Quit();
                           MessageBox.Show("Excel file is saved as text_to_excel.xlsx at my documents folder.");



                    }
                    }
                
                
            }
        }

        public void something(string para)
        {
            var excel = new Microsoft.Office.Interop.Excel.Application();
            if (excel == null)
            {
                MessageBox.Show("Excel is not properly installed!!");
                return;
            }
            else
            {
                //   MessageBox.Show("Excel is installed!!"); 
            }
            excel.Visible = false;
            excel.DisplayAlerts = false;
            var xlWorkBook = excel.Workbooks.Add();
            var xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            xlWorkSheet.Cells[1, 1] = para;
            xlWorkBook.SaveAs("hehe.xlsx");
            xlWorkBook.Close();
            excel.Quit();
        }

        private void tableLayoutPanel1_Paint(object sender, PaintEventArgs e)
        {

        }
    }
}
