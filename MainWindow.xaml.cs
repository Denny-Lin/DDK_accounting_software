using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Data.OleDb;
using System.IO;
using System.ComponentModel;
using System.Data.Odbc;
using System.Data;
using ExcelDataReader;
using System.Drawing;
using System.Drawing.Printing;
using System.Windows.Xps;
using System.Windows.Xps.Packaging;

namespace knuckle_052521
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public DataSet excel_dataset=new DataSet();//讀取出的excel資料表         

        public MainWindow()
        {
            InitializeComponent();
        }
       
        private void Button_Click(object sender, RoutedEventArgs e)//開EXCEL 顯示EXCEL combobox顯示各表單
        {
            try
            {
                OpenFileDialog ofd = new OpenFileDialog();
                string filder = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                ofd.InitialDirectory = @filder; //設定初始路徑
                ofd.Filter = "Excel檔案(*.xlsx)|*.xlsx|Csv檔案(*.csv)|*.csv|所有檔案(*.*)|*.*"; //設定“另存為檔案型別”或“檔案型別”框中出現的選擇內容
                ofd.FilterIndex = 1; //設定預設顯示檔案型別為xls //Csv檔案(*.csv)|*.csv
                ofd.Title = "開啟檔案"; //獲取或設定檔案對話方塊標題
                ofd.RestoreDirectory = true;////設定對話方塊是否記憶上次開啟的目錄

                if (ofd.ShowDialog() == true)//開檔成功
                {                                                        
                    //txtFilename.Text = ofd.FileName;//顯示檔案完整路徑

                    string filePath = ofd.FileName;

                    System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
                    //open file and returns as Stream
                    using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))//開檔讀檔
                    {                      
                        using (var excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream))//2007
                        {
                            /*
                                //1.1 Reading from a binary Excel file ('97-2003 format; *.xls)
                                IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(stream);

                                //1.2 Reading from a OpenXml Excel file (2007 format; *.xlsx)
                                IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
                            */

                            // reader.IsFirstRowAsColumnNames
                            var conf = new ExcelDataSetConfiguration
                            {
                                ConfigureDataTable = _ => new ExcelDataTableConfiguration
                                {
                                    UseHeaderRow = true
                                }
                            };

                            //DataSet - The result of each spreadsheet will be created in the result.Tables
                            //丟出去給全域變數
                            excel_dataset = excelReader.AsDataSet(conf);

                            // Now you can get data from each sheet by its index or its "name" 
                            var dataTable = excel_dataset.Tables;
                            
                            
                            //excelfilename.Content = dataTable.Rows[1][1];//偵錯 會顯示FF戰隊 顯示特定欄位的資料
                            //var tableCollection = dataSet.Tables;

                            DataGridView1.ItemsSource = dataTable[0].DefaultView;//////////////////重點代碼 工作表第一頁丟給 DataGridView 顯示

                            //excelfilename.Content = dataTable.Rows.Count;//取得行數,由上往下數，有幾層  

                            cboSheet.Items.Clear();//combobox增加各表單選擇鍵

                            for (int i = 0; ; i++) {//有點問題
                                if (excel_dataset.Tables[i] == null)
                                    break;

                                cboSheet.Items.Add(excel_dataset.Tables[i].TableName);//add sheet to combobox
                            } 

                            excelReader.Close();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.Message);
            }
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)//從第3行名字那欄開始找
        {
            //DataTable dt = new DataTable();//假設dt是由"SELECT C1,C2,C3 FROM T1"查詢出來的結果
            try
            {
                var dt = excel_dataset.Tables[0];

                if (cboSheet.SelectedItem.ToString() != null)//如果有2個表單以上，以那個表單做計算
                   dt = excel_dataset.Tables[cboSheet.SelectedItem.ToString()];

                string name = txtFilename.Text;

                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    if (dt.Rows[i][2].ToString() == name)//查詢條件012
                    {
                        txtFilename.Text = name + "買了" + dt.Rows[i][4].ToString() + "隻豬腳,共" + ((double)dt.Rows[i][5] + (double)dt.Rows[i][6]).ToString() + "元";
                        //進行操作
                    }
                }
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.Message);
            }
            /*
                            DataSet 取值
                              DataSet.Table[0].Rows[ i ][ j ]
                              其中 i 代表第 i 行數, j 代表第 j 列數
                            DataSet行數
                              DataSet.Table[0].Rows[ i ].ItemArray[ j ]
                              其中 i 代表第 i 行數, j 代表第 j 列數
                            DataSet列數
                              DataSet.Tables[0].Columns.Count
                              取得表的總列數
                              DataSet總行數　
                            DataSet.Tables[0].Rows.Count
                              取得表的總行數
                            DataSet中取出特定值
                              DataSet.Tables[0].Columns[ i ].ToString()
                              取得表的 i 列名
                            */
            /*
             DataTable的.clear()與.reset()差異
              .clear() 僅清除內容資料。後面重複利用 select 如果欄位不同可能會發生錯誤。
              .reset() 全部重設。後面重複利用 select 如果欄位不同不會發生錯誤。
            */
        }

        private void show_datatable_colum(int colum_num)//找出查到那行的整排資料，DataGridView 顯示
        {
            DataSet ds = new DataSet();

            //var dt = excel_dataset.Tables[0];
            for (int i = 0; i < colum_num; i++)
            {
                ds.Tables[0].Rows[colum_num][i] = excel_dataset.Tables[0].Rows[colum_num][i];
            }

            DataGridView1.ItemsSource = ds.Tables[0].DefaultView;//工作表第一頁丟給 DataGridView 顯示
        }
  
        private void btnPrint_Click(object sender, RoutedEventArgs e)//列印按鈕
        {
            PrintDialog Pdialog = new PrintDialog();
            Pdialog.PageRangeSelection = PageRangeSelection.AllPages;
            Pdialog.UserPageRangeEnabled = true;

            //显示打印框，选择份数和打印机
            if (Pdialog.ShowDialog() == true)
            {               
                //XpsDocument xpsDocument = new XpsDocument("C:\\FixedDocumentSequence.xps", FileAccess.ReadWrite);
                //FixedDocumentSequence fixedDocSeq = xpsDocument.GetFixedDocumentSequence();
                //Pdialog.PrintDocument(fixedDocSeq.DocumentPaginator, "Test print job");

                Pdialog.PrintVisual(DataGridView1, "Print Test");
                //Pdialog.PrintVisual(richText, "测试");
            }

            //直接打印
            // dialog.PrintVisual(richText, "测试");
        }
        
        private void cboSheet_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            //DataGridView1.ItemsSource = excel_dataset.Tables[0].DefaultView;
            DataGridView1.ItemsSource = excel_dataset.Tables[cboSheet.SelectedItem.ToString()].DefaultView;
        }

        private void TextBox_TextChanged(object sender, TextChangedEventArgs e)
        {

        }

        private void txtPath_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

        }
    }
}
