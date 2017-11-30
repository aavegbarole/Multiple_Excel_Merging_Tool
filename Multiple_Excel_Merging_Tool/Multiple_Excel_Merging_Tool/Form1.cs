using LinqToExcel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using ClosedXML.Excel;

namespace Multiple_Excel_Merging_Tool
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button_browse1_Click(object sender, EventArgs e)
        {
            LoadExcelFile();
        }

   
        private void LoadExcelFile()
        {
            OpenFileDialog fdlg = new OpenFileDialog();
            fdlg.Title = "Select Input File";
            fdlg.InitialDirectory = @"d:\Cost Management";
            fdlg.Filter = "All files (*.*)|*.*|All files (*.*)|*.*";
            fdlg.FilterIndex = 2;
            fdlg.RestoreDirectory = true;
            if (fdlg.ShowDialog() == DialogResult.OK)
            {
                SelectedExcelFiles.Items.Add(fdlg.FileName);
            }
        }

        private List<Row> excelDataRow = new List<Row>();
        private static List<Row> headerContents = new List<Row>();

        private void button_merge_Click(object sender, EventArgs e)
        {
            // 2007 Office System Driver: Data Connectivity Components.

            bool headerCollected = false;

            foreach (var items in SelectedExcelFiles.Items)
            {
                var book = new ExcelQueryFactory(items.ToString());
                var worksheetsNames = book.GetWorksheetNames().ToList();
                bool flag = false;
                int iteration = 0;

                //worksheetsNames.ForEach((item) =>
                //{
                var worksheetAllitems = book.Worksheet(worksheetsNames.FirstOrDefault()).ToList();
               
                worksheetAllitems.ForEach((row) =>
                {
                    row.ForEach((cell) =>
                    {
                        flag = flag || !cell.Value.ToString().Equals("");
                    });

                    if (flag)
                    {
                        if (iteration < 4 && !headerCollected)
                        {
                            headerContents.Add(row);
                            iteration++;
                            return;
                        }
                        excelDataRow.Add(row);
                        iteration++;
                    }
                    else
                    {
                        return;
                    }
                    flag = false;
                });

                headerCollected = true;
                //});
            }

            FinalExcelProcess();


            ToExcelFile("D:\\Hello.xlsx", "Sheet1");

            //DownloadExcel();
        }


        private void FinalExcelProcess()
        {
            headerContents.AddRange(excelDataRow);
        }
        
        public static void ToExcelFile(string filename, string worksheetName = "Sheet1")
        {
            using (var workbook = new XLWorkbook())
            {
                var ws =  workbook.Worksheets.Add(worksheetName);

                //headerContents.ForEach((item) =>
                //{
                //    ws.Cell(1, 1).InsertData(item);
                //});

                for (int i = 0; i < headerContents.Count; i++)
                {
                  //  ws.Cell(i+1, 1).InsertData(headerContents[i]);
            

                    for (int j = 0; j < headerContents[i].Count; j++)
                    {
                        ws.Cell(i+1, j+1).Value = headerContents[i][j];
                    }
                }


                      
      
            
                workbook.SaveAs(filename);
            }
        }

    }
}
