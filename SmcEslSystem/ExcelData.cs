using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace SmcEslSystem
{
    class ExcelData
    {


        public DataTable GetExcelSheetNames(string filePath)
        {
            //Office 2003
            //OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + filePath + ";Extended Properties='Excel 8.0;HDR=YES;IMEX=1'");

            //Office 2007
            try {
                OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filePath + ";Extended Properties='Excel 12.0 Xml;HDR=YES'");

                DataSet ds = new DataSet();
                conn.Open();
                DataTable dt = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                conn.Close();
                return dt;
            }
            catch (Exception ex) {
                MessageBox.Show(ex.ToString());
                return null;
            }
            
        }
        public DataTable GetExcelDataTable(string filePath, string sql)
        {
            //Office 2003
            // OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + filePath + ";Extended Properties='Excel 8.0;HDR=YES;IMEX=1;Readonly=0'");

            //Office 2007
            OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filePath + ";Extended Properties='Excel 12.0 Xml;HDR=YES'");
            OleDbDataAdapter da;
            DataTable dt = new DataTable();
            try
            {
                da = new OleDbDataAdapter(sql, conn);
                da.Fill(dt);
                dt.TableName = "tmp";
                conn.Close();
                
            }
            catch (Exception e)
            {
                MessageBox.Show("錯誤，請關閉Excel");
         //       Console.WriteLine(e.ToString());
            }


            return dt;
        }






        /// <summary>
        /// 將DataGridView匯出到Excel
        /// </summary>
        /// <param name="gridView">DataGridView</param>
        /// <param name="isShowExcle">是否顯示Excel畫面</param>
        /// <returns></returns>
        public bool ExportDataGridview(DataGridView gridView, DataGridView gridView2, DataGridView gridView3, DataGridView gridView4,DataGridView gridView7, bool isShowExcle, string filename)
        {
           // if (gridView.Rows.Count == 0)
           //     return false;
            //建立Excel

            Excel._Application myExcel = null;
            Excel._Workbook myBook = null;
            Excel._Worksheet mySheet = null;
            Excel._Worksheet mySheet2 = null;
            Excel._Worksheet mySheet3 = null;
            Excel._Worksheet mySheet4 = null;


            Excel.Application excel = new Excel.Application();
            excel.Application.Workbooks.Add(true);
           
            excel.Application.Sheets.Add(After: excel.Application.Sheets[excel.Application.Sheets.Count]);
            
            excel.Application.Sheets.Add(After: excel.Application.Sheets[excel.Application.Sheets.Count]);
           
            excel.Application.Sheets.Add(After: excel.Application.Sheets[excel.Application.Sheets.Count]);
            
            mySheet = (Excel._Worksheet)excel.Worksheets["工作表1"];//引用第一張工作表
            mySheet2 = (Excel._Worksheet)excel.Worksheets["工作表2"];//引用第一張工作表
            mySheet3 = (Excel._Worksheet)excel.Worksheets["工作表3"];//引用第一張工作表
            mySheet4 = (Excel._Worksheet)excel.Worksheets["工作表4"];//引用第一張工作表

            //    try
            //  {
            //標題
            int a = 0;
            for (int i = 0; i < gridView.ColumnCount - 3; i++)
            {
                //濾掉選項欄位
                if (i == 0)
                {
                    a = 1;
                }
                else if (i == 1)
                {
                    a = 3;
                }
                mySheet.Cells[1, i + 1] = gridView.Columns[i + a].HeaderText;
                // excel.Cells[1, i + 1] = gridView.Columns[i].HeaderText;
            }
            //數據資料
            for (int i = 0; i < gridView.RowCount - 1; i++)
            {
                for (int j = 0; j < gridView.ColumnCount - 3; j++)
                {
                    //濾掉選項欄位
                    if (j == 0)
                    {
                        a = 1;
                    }
                    else if (j == 1)
                    {
                        a = 3;
                    }
                    /* if (gridView[j, i].ValueType == typeof(string))
                       {
                           excel.Cells[i + 2, j + 1] = gridView[j+a, i].Value.ToString();
                       }
                       else
                       {
                           excel.Cells[i + 2, j + 1] = gridView[j+a, i].Value.ToString();
                       }*/
                    // if (j + a != 2)
                    //  {
                    //  Console.WriteLine("gridView[j + a, i].Value.ToString()" + gridView[j + a, i].Value.ToString());
                        mySheet.Cells[i + 2, j + 1] = gridView[j + a, i].Value.ToString();
                    // }

                }

            }




            //222222222222222222

            a = 0;
                for (int i = 0; i < gridView.ColumnCount - 1; i++)
                {
                    //濾掉選項欄位
                    if (i == 0)
                    {
                        a = 1;
                    }
                if(i<3)
                mySheet2.Cells[1, i + 1] = gridView2.Columns[i + a].HeaderText;
                    // excel.Cells[1, i + 1] = gridView.Columns[i].HeaderText;
                }
               
                // excel.Cells[1, i + 1] = gridView.Columns[i].HeaderText;
                for (int i = 0; i < gridView2.RowCount; i++)
                {

                    for (int j = 0; j < gridView2.ColumnCount-1; j++)
                    {

                        if (j == 0)
                        {
                            a = 1;
                        }
                    mySheet2.Cells[i + 2, j + 1] = gridView2[j + a, i].Value.ToString();
                    /* if (gridView[j + a, i].Value != null) {
                         Console.WriteLine("gridView2[j + a, i].Value.ToString()"+ (j + a )+"bbm"+ "i" + gridView2[j + a, i].Value.ToString());

                         mySheet2.Cells[i + 2, j + 1] = gridView2[j + a, i].Value.ToString();
                     }*/



                }
                }

                //77777777777
            for (int i = 0; i < gridView7.RowCount; i++)
            {

                for (int j = 0; j < gridView7.ColumnCount - 1; j++)
                {

                    if (j == 0)
                    {
                        a = 1;
                    }
                 //   Console.WriteLine("gridView2.RowCount+1+i" + (gridView2.RowCount + 1 + i)+ "gridView2.ColumnCount + 1+j"+(gridView2.ColumnCount + 1 + j)+ gridView7[j + a, i].Value.ToString());
                    mySheet2.Cells[gridView2.RowCount+2+i, j+1] = gridView7[j + a, i].Value.ToString();
                    /* if (gridView[j + a, i].Value != null) {
                         Console.WriteLine("gridView2[j + a, i].Value.ToString()"+ (j + a )+"bbm"+ "i" + gridView2[j + a, i].Value.ToString());

                         mySheet2.Cells[i + 2, j + 1] = gridView2[j + a, i].Value.ToString();
                     }*/



                }
            }

            //33333333333333333

            a = 0;
                for (int i = 0; i < gridView3.ColumnCount - 1; i++)
                {
                    //濾掉選項欄位
                    if (i == 0)
                    {
                        a = 1;
                    }
                mySheet3.Cells[1, i + 1] = gridView3.Columns[i + a].HeaderText;
                    // excel.Cells[1, i + 1] = gridView.Columns[i].HeaderText;
                }
                //數據資料
                for (int i = 0; i < gridView3.RowCount - 1; i++)
                {
                    for (int j = 0; j < gridView3.ColumnCount-1; j++)
                    {

                        if (j == 0)
                        {
                            a = 1;
                        }

                        if(j!=3&&j!=4)
                        mySheet3.Cells[i + 2, j + 1] = gridView3[j + a, i].Value.ToString();

                       
                       

                    }
                }

                //444444444444444444
                a = 0;
                for (int i = 0; i < gridView4.ColumnCount-1; i++)
                {
                    if (i == 0)
                    {
                        a = 1;
                    }

             //   Console.WriteLine("HeaderText4" + gridView4.Columns[i + a].HeaderText);
                mySheet4.Cells[1, i + 1] = gridView4.Columns[i + a].HeaderText;
                    // excel.Cells[1, i + 1] = gridView.Columns[i].HeaderText;
                }
                //數據資料
                for (int i = 0; i < gridView4.RowCount - 1; i++)
                {
                    for (int j = 0; j < gridView4.ColumnCount - 1; j++)
                    {

                        if (j == 0)
                        {
                            a = 1;
                        }


                    if(j!=3&&j!=4)
                        mySheet4.Cells[i + 2, j + 1] = gridView4[j + a, i].Value.ToString();

                        
                        //   Console.WriteLine(i + 2 + "," + j + 1 + "," + gridView4[j, i].Value.ToString());

                    }
                }

                //設定為按照內容自動調整欄寬
                Excel.Range oRng;
                oRng = mySheet.get_Range("A1", "N" + gridView.RowCount);
                oRng.EntireColumn.AutoFit(); // 自動調整欄寬
                Excel.Range oRng3;
                oRng3 = mySheet3.get_Range("A1", "H" + gridView.RowCount);
                oRng3.EntireColumn.AutoFit(); // 自動調整欄寬
                Excel.Range oRng4;
                oRng4 = mySheet4.get_Range("A1", "C" + gridView.RowCount);
                oRng4.EntireColumn.AutoFit(); // 自動調整欄寬


                //設定為置中
                oRng3 = mySheet3.get_Range("A1", "H" + gridView.RowCount);
                oRng3.EntireColumn.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter; //置中

                oRng3 = mySheet3.get_Range("B1", "C" + gridView.RowCount);
                oRng3.EntireColumn.NumberFormatLocal = 0;

                //設定為置中
                oRng = mySheet.get_Range("A1", "N" + gridView.RowCount);
                oRng.EntireColumn.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter; //置中

                oRng = mySheet.get_Range("B1", "C" + gridView.RowCount);
                oRng.EntireColumn.NumberFormatLocal = 0;



                oRng = mySheet.get_Range("D1", "D" + gridView.RowCount); //顏色
                oRng.Font.Color = Color.Red;



                //excel.ActiveWorkbook.SaveCopyAs(filename);
                excel.ActiveWorkbook.SaveCopyAs(filename);
                excel.Visible = isShowExcle;
                // excel.Quit();//離開聯結 
         /*   }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                excel.Save();
                mySheet = null;
                excel.Quit();
                excel = null;
            }*/
            return true;
        }


        public bool AllDataGridviewSave(DataGridView gridView, DataGridView gridView2, DataGridView gridView3, DataGridView gridView4, bool isShowExcle, string filename)
        {
            if (gridView.Rows.Count == 0)
                return false;
            //建立Excel

            Excel.Application excel = new Excel.Application();
            Excel.Workbook excelwb = excel.Workbooks.Open(filename);
            try
            {
                //    excel.Application.Workbooks.Add(true);
                // Excel.Worksheet mySheet = new Excel.Worksheet();
                Excel.Worksheet mySheet = excelwb.Worksheets["工作表1"];//引用第一張工作表
                Excel.Worksheet mySheet2 = excelwb.Worksheets["工作表2"];//引用第一張工作表
                Excel.Worksheet mySheet3 = excelwb.Worksheets["工作表3"];//引用第一張工作表
                Excel.Worksheet mySheet4 = excelwb.Worksheets["工作表4"];//引用第一張工作表


                //標題
                int a = 0;
                for (int i = 0; i < gridView.ColumnCount - 2; i++)
                {
                    //濾掉選項欄位
                    if (i == 0)
                    {
                        a = 1;
                    }
                    else if (i == 2)
                    {
                        a = 2;
                    }
                    mySheet.Cells[1, i + 1] = gridView.Columns[i + a].HeaderText;
                    // excel.Cells[1, i + 1] = gridView.Columns[i].HeaderText;
                }
                //數據資料
                for (int i = 0; i < gridView.RowCount - 1; i++)
                {
                    for (int j = 0; j < gridView.ColumnCount - 2; j++)
                    {
                        //濾掉選項欄位
                        if (j == 0)
                        {
                            a = 1;
                        }
                        else if (j == 2)
                        {
                            a = 2;
                        }
                        /* if (gridView[j, i].ValueType == typeof(string))
                           {
                               excel.Cells[i + 2, j + 1] = gridView[j+a, i].Value.ToString();
                           }
                           else
                           {
                               excel.Cells[i + 2, j + 1] = gridView[j+a, i].Value.ToString();
                           }*/
                        //  if (j + a != 2)
                        // {
                        mySheet.Cells[i + 2, j + 1] = gridView[j + a, i].Value.ToString();
                        //  }

                    }
                }

                //222222222222222222

                a = 0;

                mySheet2.Cells[1, 1] = gridView2.Columns[1].HeaderText;
                // excel.Cells[1, i + 1] = gridView.Columns[i].HeaderText;
                for (int i = 0; i < gridView2.RowCount; i++)
                {
                    if (i == 0)
                    {
                        a = 1;
                    }
                    for (int j = 0; j < gridView2.ColumnCount; j++)
                    {

                        if (j != 0)
                        {
                            int k = i + 2;
                            int s = j + 1;
                            Console.WriteLine("k" + k + "j" + j + gridView2[j, i].Value.ToString());
                            mySheet2.Cells[i + 2, j] = gridView2[j, i].Value.ToString();
                        }

                    }
                }

                //33333333333333333

                a = 0;
                for (int i = 0; i < gridView3.ColumnCount - 1; i++)
                {
                    //濾掉選項欄位
                    if (i == 0)
                    {
                        a = 1;
                    }
                    mySheet3.Cells[1, i + 1] = gridView3.Columns[i + a].HeaderText;
                    // excel.Cells[1, i + 1] = gridView.Columns[i].HeaderText;
                }
                //數據資料
                for (int i = 0; i < gridView3.RowCount - 1; i++)
                {
                    for (int j = 0; j < gridView3.ColumnCount; j++)
                    {

                        if (j != 0 && j != 4)
                        {
                            int k = i + 2;
                            int s = j + 1;
                            //   Console.WriteLine(k + "," + s + "," + gridView3[j, i].Value.ToString());
                            mySheet3.Cells[i + 2, j] = gridView3[j, i].Value.ToString();
                        }

                    }
                }

                //444444444444444444
                a = 0;
                for (int i = 0; i < gridView4.ColumnCount; i++)
                {

                    mySheet4.Cells[1, i + 1] = gridView4.Columns[i + a].HeaderText;
                    // excel.Cells[1, i + 1] = gridView.Columns[i].HeaderText;
                }
                //數據資料
                for (int i = 0; i < gridView4.RowCount - 1; i++)
                {
                    for (int j = 0; j < gridView4.ColumnCount - 1; j++)
                    {

                        mySheet4.Cells[i + 2, j + 1] = gridView4[j + a, i].Value.ToString();
                        //   Console.WriteLine(i + 2 + "," + j + 1 + "," + gridView4[j, i].Value.ToString());

                    }
                }

                //設定為按照內容自動調整欄寬
                Excel.Range oRng;
                oRng = mySheet.get_Range("A1", "N" + gridView.RowCount);
                oRng.EntireColumn.AutoFit(); // 自動調整欄寬
                Excel.Range oRng3;
                oRng3 = mySheet3.get_Range("A1", "H" + gridView.RowCount);
                oRng3.EntireColumn.AutoFit(); // 自動調整欄寬
                Excel.Range oRng4;
                oRng4 = mySheet4.get_Range("A1", "C" + gridView.RowCount);
                oRng4.EntireColumn.AutoFit(); // 自動調整欄寬


                //設定為置中
                oRng3 = mySheet3.get_Range("A1", "H" + gridView.RowCount);
                oRng3.EntireColumn.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter; //置中

                oRng3 = mySheet3.get_Range("B1", "C" + gridView.RowCount);
                oRng3.EntireColumn.NumberFormatLocal = 0;

                //設定為置中
                oRng = mySheet.get_Range("A1", "N" + gridView.RowCount);
                oRng.EntireColumn.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter; //置中

                oRng = mySheet.get_Range("B1", "C" + gridView.RowCount);
                oRng.EntireColumn.NumberFormatLocal = 0;



                oRng = mySheet.get_Range("D1", "D" + gridView.RowCount); //顏色
                oRng.Font.Color = Color.Red;



                excelwb.Save();
                mySheet = null;
                excelwb.Close();
                excelwb = null;
                excel.Quit();
                excel = null;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                excelwb.Save();
               // mySheet = null;
                excelwb.Close();
                excelwb = null;
                excel.Quit();
                excel = null;
            }
            //excel.Visible = isShowExcle;
            // excel.Quit();//離開聯結 
            return true;
        }


        public bool DataGridviewSave(DataGridView gridView, bool isShowExcle, string filename)
        {
         //   Console.WriteLine("工作表1"+ gridView);
            if (gridView.Rows.Count == 0)
                return false;
            //建立Excel

            //Excel._Application myExcel = null;
           // Excel._Workbook myBook = null;
           // Excel._Worksheet mySheet = null;


            Excel.Application excel = new Excel.Application();
            Excel.Workbook excelwb = excel.Workbooks.Open(filename);
            //    excel.Application.Workbooks.Add(true);
            Excel.Worksheet mySheet = new Excel.Worksheet();
             mySheet = excelwb.Worksheets["工作表1"];//引用第一張工作表
            try
            {
                //標題
                int a = 0;
                for (int i = 0; i < gridView.ColumnCount - 2; i++)
                {
                    //濾掉選項欄位
                    if (i == 0)
                    {
                        a = 1;
                    }
                    else if (i == 2)
                    {
                        a = 2;
                    }

                    mySheet.Cells[1, i + 1] = gridView.Columns[i + a].HeaderText;
                    // excel.Cells[1, i + 1] = gridView.Columns[i].HeaderText;
                }
                //數據資料
                for (int i = 0; i < gridView.RowCount - 1; i++)
                {
                    for (int j = 0; j < gridView.ColumnCount - 2; j++)
                    {
                        //濾掉選項欄位
                        if (j == 0)
                        {
                            a = 1;
                        }
                        else if (j == 2)
                        {
                            a = 2;
                        }
                        /* if (gridView[j, i].ValueType == typeof(string))
                           {
                               excel.Cells[i + 2, j + 1] = gridView[j+a, i].Value.ToString();
                           }
                           else
                           {
                               excel.Cells[i + 2, j + 1] = gridView[j+a, i].Value.ToString();
                           }*/
                        //  if (j + a != 2)
                        // {
                        int k = i + 2;
                        int s = j + 1;
                        //        Console.WriteLine(k + "," + s + ","+ gridView[j + a, i].Value.ToString());
                        mySheet.Cells[i + 2, j + 1] = gridView[j + a, i].Value.ToString();
                        //  }

                    }
                }


                //設置禁止彈出保存和覆蓋的詢問提示框
                mySheet.Application.DisplayAlerts = isShowExcle;
                mySheet.Application.AlertBeforeOverwriting = isShowExcle;
                //設定為按照內容自動調整欄寬
                Excel.Range oRng;
                oRng = mySheet.get_Range("A1", "N" + gridView.RowCount);
                oRng.EntireColumn.AutoFit(); // 自動調整欄寬

                //設定為置中
                oRng = mySheet.get_Range("A1", "N" + gridView.RowCount);
                oRng.EntireColumn.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter; //置中

                oRng = mySheet.get_Range("B1", "C" + gridView.RowCount);
                oRng.EntireColumn.NumberFormatLocal = 0;



                oRng = mySheet.get_Range("D1", "D" + gridView.RowCount); //顏色
                oRng.Font.Color = Color.Red;



                //excel.ActiveWorkbook.SaveCopyAs(filename);
                excelwb.Save();
                mySheet = null;
                excelwb.Close();
                excelwb = null;
                excel.Quit();
                excel = null;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                excelwb.Save();
                mySheet = null;
                excelwb.Close();
                excelwb = null;
                excel.Quit();
                excel = null;
            }
            //excel.Visible = isShowExcle;
            // excel.Quit();//離開聯結 
            return true;
        }

        public bool EslStyleCgange(DataGridView gridView, string  style, string styleName, bool isShowExcle, string filename, Excel.Application excel, Excel.Workbook excelwb, Excel.Worksheet mySheet)
        {
            mySheet = excelwb.Worksheets["工作表2"];//引用第一張工作表
           
            for (int i = 1; i < mySheet.Rows.Count; i++)
            {
                
                if (mySheet.Rows[i].Cells[1].Value != null && mySheet.Rows[i].Cells[1].Value.ToString() == styleName)
                {
                    Console.WriteLine("mySheet" + mySheet.Rows[i].Cells[1].Value.ToString() + "styleName" + styleName);
                    mySheet.Rows[i].Cells[2].Value = style;
                    break;
                }
            }
            return true;
        }


        public bool dataGridViewRowCellUpdate(DataGridView gridView,int columns,int row, bool isShowExcle, string filename, Excel.Application excel, Excel.Workbook excelwb, Excel.Worksheet mySheet)
        {
            if (gridView.Rows.Count == 0)
                return false;

            Console.WriteLine("工作表1" + columns+ "row" + row);
            if (gridView.Name == "dataGridView1")
            {
                mySheet = excelwb.Worksheets["工作表1"];//引用第一張工作表
                int c = 0;
                Console.WriteLine("columns" + columns);
                if (columns > 3)
                    c = 2;

                mySheet.Cells[row + 2, columns - c] = gridView.Rows[row].Cells[columns].Value;
            }
            else if(gridView.Name == "dataGridView2")
            {
                mySheet = excelwb.Worksheets["工作表2"];//引用第一張工作表
                mySheet.Cells[row + 2, columns] = gridView.Rows[row].Cells[columns].Value;
            }
            else if (gridView.Name == "dataGridView7")
            {
                mySheet = excelwb.Worksheets["工作表2"];//引用第一張工作表
                mySheet.Cells[row + 2, columns] = gridView.Rows[row].Cells[columns].Value;
            }
            else if (gridView.Name == "dataGridView4")
            {
                Console.WriteLine("dataGridView4");
                mySheet = excelwb.Worksheets["工作表3"];//引用第一張工作表
                mySheet.Cells[row + 2, columns] = gridView.Rows[row].Cells[columns].Value;
                Console.WriteLine("row" + row + "columns"+ columns+":"+ gridView.Rows[row].Cells[columns].Value);
                Console.WriteLine("mySheet.Cells[row + 2, columns]"+ mySheet.Cells[row + 2, columns]);
            }
            else if (gridView.Name == "dataGridView5")
            {
                mySheet = excelwb.Worksheets["工作表4"];//引用第一張工作表
                mySheet.Cells[row + 2, columns] = gridView.Rows[row].Cells[columns].Value;
            }

            return true;
        }



        public bool DataGridview4Update(DataGridView gridView, bool isShowExcle, string filename)
        {
            if (gridView.Rows.Count == 0)
                return false;
            //建立Excel

            //Excel._Application myExcel = null;
            // Excel._Workbook myBook = null;
            // Excel._Worksheet mySheet = null;

      //      Console.WriteLine("工作表3" + gridView);
          //  Console.WriteLine("gridView.ColumnCount" + gridView.ColumnCount);
         //   Console.WriteLine("gridView.RowCount" + gridView.RowCount);
            Excel.Application excel = new Excel.Application();
            Excel.Workbook excelwb = excel.Workbooks.Open(filename);
            //    excel.Application.Workbooks.Add(true);
            Excel.Worksheet mySheet = new Excel.Worksheet();
            mySheet = excelwb.Worksheets["工作表3"];//引用第一張工作表
            //標題
            int a = 0;
            for (int i = 0; i < gridView.ColumnCount - 1; i++)
            {
                //濾掉選項欄位
                if (i == 0)
                {
                    a = 1;
                }
                mySheet.Cells[1, i + 1] = gridView.Columns[i + a].HeaderText;
                // excel.Cells[1, i + 1] = gridView.Columns[i].HeaderText;
            }
            //數據資料
            for (int i = 0; i < gridView.RowCount - 1; i++)
            {
                for (int j = 0; j < gridView.ColumnCount; j++)
                {

                    if (j!= 0&&j != 4)
                    {
                        int k = i + 2;
                        int s = j + 1;
                        mySheet.Cells[i + 2, j] = gridView[j , i].Value.ToString();
                    }

                }
            }


            //設置禁止彈出保存和覆蓋的詢問提示框
            mySheet.Application.DisplayAlerts = isShowExcle;
            mySheet.Application.AlertBeforeOverwriting = isShowExcle;
            //設定為按照內容自動調整欄寬
            Excel.Range oRng;
            oRng = mySheet.get_Range("A1", "H" + gridView.RowCount);
            oRng.EntireColumn.AutoFit(); // 自動調整欄寬

            //設定為置中
            oRng = mySheet.get_Range("A1", "H" + gridView.RowCount);
            oRng.EntireColumn.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter; //置中

            oRng = mySheet.get_Range("B1", "C" + gridView.RowCount);
            oRng.EntireColumn.NumberFormatLocal = 0;



            oRng = mySheet.get_Range("D1", "D" + gridView.RowCount); //顏色
            oRng.Font.Color = Color.Red;



            //excel.ActiveWorkbook.SaveCopyAs(filename);
            excelwb.Save();
            mySheet = null;
            excelwb.Close();
            excelwb = null;
            excel.Quit();
            excel = null;
            //excel.Visible = isShowExcle;
            // excel.Quit();//離開聯結 
            return true;
        }


        public bool DataGridview5Update(DataGridView gridView, bool isShowExcle, string filename, Excel.Application excel, Excel.Workbook excelwb, Excel.Worksheet mySheet)
        {
            if (gridView.Rows.Count == 0)
                return false;
            //建立Excel

            //Excel._Application myExcel = null;
            // Excel._Workbook myBook = null;
            // Excel._Worksheet mySheet = null;

         //   Console.WriteLine("工作表4" + gridView);
          //  Console.WriteLine("gridView.ColumnCount" + gridView.ColumnCount);
          //  Console.WriteLine("gridView.RowCount" + gridView.RowCount);
          //  Excel.Application excel = new Excel.Application();
          //  Excel.Workbook excelwb = excel.Workbooks.Open(filename);
            //    excel.Application.Workbooks.Add(true);
          //  Excel.Worksheet mySheet = new Excel.Worksheet();
            mySheet = excelwb.Worksheets["工作表4"];//引用第一張工作表
            //標題
            int a = 0;
            for (int i = 0; i < gridView.ColumnCount - 1; i++)
            {

                if (i == 0)
                {
                    a = 1;
                }
                mySheet.Cells[1, i + 1] = gridView.Columns[i + a].HeaderText;
                // excel.Cells[1, i + 1] = gridView.Columns[i].HeaderText;
            }
            //數據資料
            for (int i = 0; i < gridView.RowCount - 1; i++)
            {
                for (int j = 0; j < gridView.ColumnCount; j++)
                {

                    if (j != 0&&j != 4)
                    {
                        mySheet.Cells[i + 2, j] = gridView[j, i].Value.ToString();
                    }
                   
                       // Console.WriteLine(i + 2 + "," + j + 1 + "," + gridView[j, i].Value.ToString());

                }
            }


            //設置禁止彈出保存和覆蓋的詢問提示框
        /*    mySheet.Application.DisplayAlerts = isShowExcle;
            mySheet.Application.AlertBeforeOverwriting = isShowExcle;
            //設定為按照內容自動調整欄寬
            Excel.Range oRng;
            oRng = mySheet.get_Range("A1", "C" + gridView.RowCount);
            oRng.EntireColumn.AutoFit(); // 自動調整欄寬

            //設定為置中
            oRng = mySheet.get_Range("A1", "C" + gridView.RowCount);
            oRng.EntireColumn.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter; //置中

            oRng = mySheet.get_Range("B1", "C" + gridView.RowCount);
            oRng.EntireColumn.NumberFormatLocal = 0;



            oRng = mySheet.get_Range("B1", "B" + gridView.RowCount); //顏色
            oRng.Font.Color = Color.Red;*/



            //excel.ActiveWorkbook.SaveCopyAs(filename);
        /*    excelwb.Save();
            mySheet = null;
            excelwb.Close();
            excelwb = null;
            excel.Quit();
            excel = null;*/
            //excel.Visible = isShowExcle;
            // excel.Quit();//離開聯結 
            return true;
        }

        public bool dataGridView2Update(DataGridView dataGrid,string styleName,string fileName,PictureBox pictureBox1, Excel.Application excel, Excel.Workbook excelwb, Excel.Worksheet mySheet,int ESLStyleNumber,int size) {
           // Excel.Application excel = new Excel.Application();
          //  Excel.Workbook excelwb = excel.Workbooks.Open(fileName);
            //    excel.Application.Workbooks.Add(true);
          //  Excel.Worksheet mySheet = new Excel.Worksheet();
            mySheet = excelwb.Worksheets["工作表2"];//引用第一張工作表
            Excel.Range last = mySheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
            int lastUsedRow = last.Row;
            lastUsedRow = lastUsedRow + 1;
            int col = 5;
            DataGridViewRow row = (DataGridViewRow)dataGrid.Rows[0].Clone();
            DataGridViewCheckBoxColumn chk = new DataGridViewCheckBoxColumn();
            dataGrid.Columns.Add(chk);
            int Rowcount = dataGrid.RowCount;
            foreach (Control x in pictureBox1.Controls)
            {

                // mySheet.Rows[lastUsedRow].Add(x.Name, x.Width, x.Height, x.Location.X, x.Location.Y, x.Font);
                //   Console.WriteLine("lastUsedRow" + lastUsedRow);
                /* switch (x.Name)
                 {
                     case "ProName":
                         col = 1;
                         break;
                     case "ProBrand":
                         col = 7;
                         break;
                     case "ProFormat":
                         col = 13;
                         break;
                     case "ProPrice":
                         col = 19;
                         break;
                     case "ProPromotion":
                         col = 25;
                         break;
                     case "ProBarcode":
                         col = 31;
                         break;
                     case "ProESLID":
                         col = 37;
                         break;
                 //}*/
                // Console.WriteLine("col" + col + "lastUsedRow" + lastUsedRow + "x.Tag.ToString()" + x.Tag.ToString());
                mySheet.Cells[lastUsedRow, col] = x.Tag.ToString();
                col++;
                mySheet.Cells[lastUsedRow, col] = x.Name;
                col++;
                mySheet.Cells[lastUsedRow, col] = x.Text;
                col++;
                mySheet.Cells[lastUsedRow, col] = x.Width;
                col++;
                mySheet.Cells[lastUsedRow, col] = x.Height;
                col++;
                mySheet.Cells[lastUsedRow, col] = x.Location.X;
                col++;
                mySheet.Cells[lastUsedRow, col] = x.Location.Y;
                col++;
                mySheet.Cells[lastUsedRow, col] = x.Font.Name;
            //    Console.WriteLine("Name" + x.Name + "width" + x.Width + x.Height + "textBox1.Location" + x.Location + "x.font" + x.Font + " x.ForeColor" + x.ForeColor.A + "," + x.ForeColor.R + "," + x.ForeColor.G + "," + x.ForeColor.B + "x.Font.Style" + x.Font.Style + "x.BackColor" + x.BackColor.A + "," + x.BackColor.R + "," + x.BackColor.G + "," + x.BackColor.B);
                col++;
                mySheet.Cells[lastUsedRow, col] = x.Font.Size;
                col++;
                mySheet.Cells[lastUsedRow, col] = x.Font.Style.ToString();
                col++;
                mySheet.Cells[lastUsedRow, col] = x.ForeColor.A;
                col++;
                mySheet.Cells[lastUsedRow, col] = x.ForeColor.R;
                col++;
                mySheet.Cells[lastUsedRow, col] = x.ForeColor.G;
                col++;
                mySheet.Cells[lastUsedRow, col] = x.ForeColor.B;
                col++;
                mySheet.Cells[lastUsedRow, col] = x.BackColor.A;
                col++;
                mySheet.Cells[lastUsedRow, col] = x.BackColor.R;
                col++;
                mySheet.Cells[lastUsedRow, col] = x.BackColor.G;
                col++;
                mySheet.Cells[lastUsedRow, col] = x.BackColor.B;
                col++;

            }
            mySheet.Cells[lastUsedRow, 3] = ESLStyleNumber;
            mySheet.Cells[lastUsedRow, 2] = "";
            mySheet.Cells[lastUsedRow, 1] = styleName;
            mySheet.Cells[lastUsedRow, 4] = size;
            //mySheet.Cells[lastUsedRow, 2] = "";
            //設置禁止彈出保存和覆蓋的詢問提示框
            mySheet.Application.DisplayAlerts = true;
            mySheet.Application.AlertBeforeOverwriting = true;


            //excel.ActiveWorkbook.SaveCopyAs(filename);
            excelwb.Save();
            mySheet = null;
            excelwb.Close();
            excelwb = null;
            excel.Quit();
            excel = null;
            return true;
        }
        public bool UpdateDataList(bool isShowExcle, string filename,List<Page1> PageList)
        {

            //每天更新紀錄
            string filepath;
            //Excel._Application myExcel = null;
             Excel.Workbook excelwb = null;
             Excel.Worksheet mySheet = null;
            string exeDir = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            //using (StreamWriter sw = File.CreateText(path)) { }
            string today = DateTime.Now.ToString("yyyyMMdd");
           // filename =filename;
            filepath = exeDir +@"\"+ today + filename;
            Excel.Application excel = new Excel.Application();
           // Console.WriteLine("-----L"+ filename);
          //  Console.WriteLine("-----L" + File.Exists(filename));
            if (!File.Exists(filepath))
            {
            //    Console.WriteLine("exist");
                excelwb = excel.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);
                mySheet = new Excel.Worksheet();
                mySheet = excelwb.Worksheets["工作表1"];//引用第一張工作表
                                                     //標題
                int a = 0;
                for (int i = 0; i < PageList.Count; i++)
                {
                  //  Console.WriteLine("PageList[i].product_name"+PageList[i].product_name);
                     a = i + 1;
                    mySheet.Cells[a, 1] = PageList[i].no;
                    mySheet.Cells[a, 2] = PageList[i].barcode;
                    mySheet.Cells[a, 3] = PageList[i].product_name;
                    mySheet.Cells[a, 4] = PageList[i].Brand;
                    mySheet.Cells[a, 5] = PageList[i].specification;
                    mySheet.Cells[a, 6] = PageList[i].price;
                    mySheet.Cells[a, 7] = PageList[i].Special_offer;
                    mySheet.Cells[a, 8] = PageList[i].Web;
                    mySheet.Cells[a, 9] = PageList[i].BleAddress;
                    mySheet.Cells[a, 10] = PageList[i].usingAddress;
                    mySheet.Cells[a, 11] = PageList[i].onsale;
                    mySheet.Cells[a, 12] = PageList[i].UpdateState;
                    mySheet.Cells[a, 13] = PageList[i].UpdateTime;

                }

                //設置禁止彈出保存和覆蓋的詢問提示框
                 mySheet.Application.DisplayAlerts = isShowExcle;
                 mySheet.Application.AlertBeforeOverwriting = isShowExcle;
                 //設定為按照內容自動調整欄寬
                 Excel.Range oRng;
                 oRng = mySheet.get_Range("A1", "N" + PageList.Count);
                 oRng.EntireColumn.AutoFit(); // 自動調整欄寬

                 //設定為置中
                 oRng = mySheet.get_Range("A1", "N" + PageList.Count);
                 oRng.EntireColumn.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter; //置中

                 oRng = mySheet.get_Range("B1", "C" + PageList.Count);
                 oRng.EntireColumn.NumberFormatLocal = 0;



                 oRng = mySheet.get_Range("L1", "L" + PageList.Count); //顏色
                 oRng.Font.Color = Color.Red;

                excelwb.SaveAs(filepath, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                //excel.Save();
               // excelwb.Save();
                mySheet = null;
                excelwb.Close();
                excelwb = null;
                excel.Quit();
                excel = null;
                //excel.Visible = isShowExcle;
                // excel.Quit();//離開聯結 
            }
            else
            {

             //   Console.WriteLine("GGGDDD");
                excelwb = excel.Workbooks.Open(filepath);
             //   Console.WriteLine("工作表3");
                // Excel.Workbook excelwb = excel.Workbooks.Open(filename);
                //    excel.Application.Workbooks.Add(true);
                mySheet = new Excel.Worksheet();
                mySheet = excelwb.Worksheets["工作表1"];//引用第一張工作表
                                                     //標題
                Excel.Range last = mySheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
                int a = 0;
               
                for (int i = 0; i < PageList.Count; i++)
                {
                    a = last.Row + 1+i;
                  //  Console.WriteLine("AA" + a);
                    excel.Cells[a, 1] = PageList[i].no;
                    excel.Cells[a, 2] = PageList[i].barcode;
                    excel.Cells[a, 3] = PageList[i].product_name;
                  //  Console.WriteLine("AA" + PageList[i].product_name);
                    excel.Cells[a, 4] = PageList[i].Brand;
                    excel.Cells[a, 5] = PageList[i].specification;
                    excel.Cells[a, 6] = PageList[i].price;
                    excel.Cells[a, 7] = PageList[i].Special_offer;
                    excel.Cells[a, 8] = PageList[i].Web;
                    excel.Cells[a, 9] = PageList[i].BleAddress;
                    excel.Cells[a, 10] = PageList[i].usingAddress;
                    excel.Cells[a, 11] = PageList[i].onsale;
                  //  Console.WriteLine("AA" + PageList[i].onsale);
                    excel.Cells[a, 12] = PageList[i].UpdateState;
                  //  Console.WriteLine("AA" + PageList[i].UpdateState);
                    excel.Cells[a, 13] = PageList[i].UpdateTime;
                  //  Console.WriteLine("AA" + PageList[i].UpdateTime);

                }

                //設置禁止彈出保存和覆蓋的詢問提示框
                mySheet.Application.DisplayAlerts = isShowExcle;
                mySheet.Application.AlertBeforeOverwriting = isShowExcle;
                //設定為按照內容自動調整欄寬
                Excel.Range oRng;
                oRng = mySheet.get_Range("A1", "M" + PageList.Count);
                oRng.EntireColumn.AutoFit(); // 自動調整欄寬

                //設定為置中
                oRng = mySheet.get_Range("A1", "N" + PageList.Count);
                oRng.EntireColumn.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter; //置中

                oRng = mySheet.get_Range("B1", "C" + PageList.Count);
                oRng.EntireColumn.NumberFormatLocal = 0;



                oRng = mySheet.get_Range("L1", "L" + PageList.Count); //顏色
                oRng.Font.Color = Color.Red;



                excelwb.Save();
                mySheet = null;
                excelwb.Close();
                excelwb = null;
                excel.Quit();
                excel = null;
                //excel.Visible = isShowExcle;
                // excel.Quit();//離開聯結 
            }
            
            return true;
        }



        public bool dataviewdel(DataGridView gridView, List<int> delno,string table ,string filename, Excel.Application excel, Excel.Workbook excelwb,Excel.Worksheet mySheet)
        {
            if (gridView.Rows.Count == 0)
                return false;
            //建立Excel

            //Excel._Application myExcel = null;
            // Excel._Workbook myBook = null;
            // Excel._Worksheet mySheet = null;


            //Excel.Application excel = new Excel.Application();
            // Excel.Workbook excelwb = excel.Workbooks.Open(filename);
            //    excel.Application.Workbooks.Add(true);
            //Excel.Worksheet mySheet = new Excel.Worksheet();
            if (gridView.Name == "dataGridView2" || gridView.Name == "dataGridView7")
            {
                mySheet = excelwb.Worksheets[table];//引用第一張工作表
                Excel.Range last = mySheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
                int count = last.Row + 1;
                for (int i = 2; i < count; i++)
                {
                
                    if (mySheet.Cells[i, 1].Value!=null&&gridView.Rows[delno[0]].Cells[1].Value.ToString() == mySheet.Cells[i, 1].Value.ToString())
                    {
                        Console.WriteLine("(gridView.Ro5555555555555555");
                        mySheet.Rows[i].Delete();
                    }
                }
            }
            else {
                mySheet = excelwb.Worksheets[table];//引用第一張工作表
                for (int i = 0; i < delno.Count; i++)
                {
                    // Console.WriteLine("///////////kk" + mySheet.Rows[delno[i]].Cells[0].Value.ToString());

                    //   Console.WriteLine("excek"+ delno[i]);
                    mySheet.Rows[delno[i]].Delete();

                }
            }
           

        /*    excelwb.Save();
            mySheet = null;
            excelwb.Close();
            excelwb = null;
            excel.Quit();
            excel = null;*/
            return true;

        }

        public bool ESLStyleCover( string StyleName,PictureBox pictureBox1, Excel.Application excel, Excel.Workbook excelwb, Excel.Worksheet mySheet,int size)
        {
            //建立Excel

            //Excel._Application myExcel = null;
            // Excel._Workbook myBook = null;
            // Excel._Worksheet mySheet = null;

            //   Console.WriteLine("工作表4" + gridView);
            //  Console.WriteLine("gridView.ColumnCount" + gridView.ColumnCount);
            //  Console.WriteLine("gridView.RowCount" + gridView.RowCount);
            //  Excel.Application excel = new Excel.Application();
            //  Excel.Workbook excelwb = excel.Workbooks.Open(filename);
            //    excel.Application.Workbooks.Add(true);
            //  Excel.Worksheet mySheet = new Excel.Worksheet();
            mySheet = excelwb.Worksheets["工作表2"];//引用第一張工作表
            int rowNumber = 0;
            string Style = "0";
            string Toah = "";
            Excel.Range last = mySheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
            rowNumber=last.Row + 1;
            int count = last.Row + 1;
            for (int i = 2; i < count; i++)
            {

                if (mySheet.Cells[i, 1].Value != null && StyleName == mySheet.Cells[i, 1].Value.ToString())
                {
                    //rowNumber = i;
                    Style=mySheet.Cells[i, 3].Value.ToString();
                    if (mySheet.Cells[i, 2].Value ==null)
                        Toah = "";
                    else
                        Toah = mySheet.Cells[i, 2].Value.ToString();
                    Console.WriteLine("(gridView.Ro5555555555555555");
                    mySheet.Rows[i].Delete();
                }
            }
            //標題
            /*        int a = 0;
                    for (int i = 1; i < mySheet.Rows.Count; i++)
                    {
                        if (StyleName == mySheet.Rows[i].Cells[1].Value.ToString())
                        {
                            rowNumber = i;
                            for (int p = 1; p < mySheet.Columns.Count-1; p++)
                            {

                               if (p != 3&& p != 2)
                                {
                                    if (mySheet.Rows[i].Cells[p].Value != null && mySheet.Rows[i].Cells[p].Value.ToString() == "")
                                        break;
                                  //  Console.WriteLine(" mySheet.Rows[i].Cells[p].Value" + mySheet.Rows[i].Cells[p].Value.ToString());
                                    mySheet.Rows[i].Cells[p].Value = DBNull.Value;
                                }

                            }
                        }
                    }*/
            int col = 5;
            //----------
            foreach (Control x in pictureBox1.Controls)
            {
                Console.WriteLine("dfffffffffffff");
                // mySheet.Rows[lastUsedRow].Add(x.Name, x.Width, x.Height, x.Location.X, x.Location.Y, x.Font);
                //   Console.WriteLine("lastUsedRow" + lastUsedRow);
                /* switch (x.Name)
                 {
                     case "ProName":
                         col = 1;
                         break;
                     case "ProBrand":
                         col = 7;
                         break;
                     case "ProFormat":
                         col = 13;
                         break;
                     case "ProPrice":
                         col = 19;
                         break;
                     case "ProPromotion":
                         col = 25;
                         break;
                     case "ProBarcode":
                         col = 31;
                         break;
                     case "ProESLID":
                         col = 37;
                         break;
                 //}*/
                // Console.WriteLine("col" + col + "lastUsedRow" + lastUsedRow + "x.Tag.ToString()" + x.Tag.ToString());
                mySheet.Cells[rowNumber, col] = x.Tag.ToString();
                col++;
                mySheet.Cells[rowNumber, col] = x.Name;
                col++;
                mySheet.Cells[rowNumber, col] = x.Text;
                col++;
                mySheet.Cells[rowNumber, col] = x.Width;
                col++;
                mySheet.Cells[rowNumber, col] = x.Height;
                col++;
                mySheet.Cells[rowNumber, col] = x.Location.X;
                col++;
                mySheet.Cells[rowNumber, col] = x.Location.Y;
                col++;
                mySheet.Cells[rowNumber, col] = x.Font.Name;
                //    Console.WriteLine("Name" + x.Name + "width" + x.Width + x.Height + "textBox1.Location" + x.Location + "x.font" + x.Font + " x.ForeColor" + x.ForeColor.A + "," + x.ForeColor.R + "," + x.ForeColor.G + "," + x.ForeColor.B + "x.Font.Style" + x.Font.Style + "x.BackColor" + x.BackColor.A + "," + x.BackColor.R + "," + x.BackColor.G + "," + x.BackColor.B);
                col++;
                mySheet.Cells[rowNumber, col] = x.Font.Size;
                col++;
                mySheet.Cells[rowNumber, col] = x.Font.Style.ToString();
                col++;
                mySheet.Cells[rowNumber, col] = x.ForeColor.A;
                col++;
                mySheet.Cells[rowNumber, col] = x.ForeColor.R;
                col++;
                mySheet.Cells[rowNumber, col] = x.ForeColor.G;
                col++;
                mySheet.Cells[rowNumber, col] = x.ForeColor.B;
                col++;
                mySheet.Cells[rowNumber, col] = x.BackColor.A;
                col++;
                mySheet.Cells[rowNumber, col] = x.BackColor.R;
                col++;
                mySheet.Cells[rowNumber, col] = x.BackColor.G;
                col++;
                mySheet.Cells[rowNumber, col] = x.BackColor.B;
                col++;

            }
            mySheet.Cells[rowNumber, 3] = Style;
            mySheet.Cells[rowNumber, 2] = Toah;
            mySheet.Cells[rowNumber, 1] = StyleName;
            mySheet.Cells[rowNumber, 4] = size;
            //mySheet.Cells[lastUsedRow, 2] = "";
            //設置禁止彈出保存和覆蓋的詢問提示框
            mySheet.Application.DisplayAlerts = true;
            mySheet.Application.AlertBeforeOverwriting = true;


            //excel.ActiveWorkbook.SaveCopyAs(filename);
            excelwb.Save();
            mySheet = null;
            excelwb.Close();
            excelwb = null;
            excel.Quit();
            excel = null;
            return true;



        }

    }
}
