using System;
using System.Collections.Generic;
using System.Linq;
using System.Data;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Windows.Forms;
using ESRI.ArcGIS.Controls;
using ESRI.ArcGIS.esriSystem;
using ESRI.ArcGIS.SystemUI;
using ESRI.ArcGIS.Geometry;
using ESRI.ArcGIS.Carto;
using ESRI.ArcGIS.Display;
using System.Drawing;
using WpfApplication1.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Web.UI.DataVisualization.Charting;
using DevExpress.Xpf.Charts;
using DevExpress.Xpf.DemoBase;
using Model.OperateOracle;

namespace WpfApplication1
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        private static string strWhere ="";//查询条件
        public int pageSize;//页行数
        public int pageIndex;//当前页
        public int pageCount;//总页数
        public long row;     //总行数
        int isNewRow = 0;
        long startIndex;//每页开始的记录行数
        long endIndex;//每页结束的记录行数


        string tableName = null;//查询或插入数据的表名
        DataTable dt;           //存储查询数据的表格
        DataTable newdtb = new DataTable();

      
         AxMapControl mapControl;//地图控件
         AxToolbarControl toolbarControl;//地图工具控件
       
       

        public MainWindow()
        { 
         
            InitializeComponent();
            CreateEngineControls();
            ESRI.ArcGIS.RuntimeManager.Bind(ESRI.ArcGIS.ProductCode.EngineOrDesktop);


            //初始化combox控件选择，绑定表格控件数据
            this.PageRowNum.SelectedIndex = 0;
            this.comboBoxInsert.SelectedIndex = 0;
            pageIndex = 1;
            BindPageGridList(strWhere, tableName);



        }


        /// <summary>
        /// 绑定表格控件数据
        /// </summary>
        /// <param name="strWhere"></param>
        public void BindPageGridList(string strWhere,string TableName)
        {

            this.dataGrid.UnselectAll();
            if (this.PageRowNum.SelectedIndex == 0) { pageSize = 10; }
            else if (this.PageRowNum.SelectedIndex == 1) { pageSize = 20; }
            else { pageSize = 30; };


            this.firstPage.IsEnabled = true;
            this.prePage.IsEnabled = true;
            this.nextPage.IsEnabled = true;
            this.finPage.IsEnabled = true;
            //记录获取开始数
            startIndex = (pageIndex - 1) * pageSize + 1;
            //结束数

            endIndex = pageIndex * pageSize;

            //总行数
            row = OperateOracle.GetRecordCount(TableName,strWhere);

            if (row % pageSize > 0)
            { pageCount = (int)(row / pageSize + 1); }
            else
            { pageCount = (int)(row / pageSize); }

            if (pageIndex == 1)
            {
                this.firstPage.IsEnabled = false;
                this.prePage.IsEnabled = false;
            }

            if (pageIndex == pageCount)
            {
                endIndex = row;
                this.nextPage.IsEnabled = false;
                this.finPage.IsEnabled = false;
            }

            //分页获取数据列表
            dt = OperateOracle.GetListByPage(TableName,strWhere, "", startIndex, endIndex);



            this.dataGrid.ItemsSource = dt;

            // this.dataGrid.DataContext = dt;
            // this.dataTable.DataContext = dt;//?
            this.nowPage.Text = pageIndex.ToString();
            this.totalPage.Content = string.Format(" 共{0}页", pageCount);

        }



        //表格翻页
        private void firstPage_Click(object sender, RoutedEventArgs e)
        {
            pageIndex = 1;
            //绑定分页控件和GridControl数据  
            BindPageGridList(strWhere,tableName);
        }
        private void prePage_Click(object sender, RoutedEventArgs e)
        {
            pageIndex --;
            //绑定分页控件和GridControl数据  
            BindPageGridList(strWhere, tableName);
        }
        private void nextPage_Click(object sender, RoutedEventArgs e)
        {
            pageIndex++;
            //绑定分页控件和GridControl数据  
            BindPageGridList(strWhere, tableName);
        }
        private void finPage_Click(object sender, RoutedEventArgs e)
        {
            pageIndex = pageCount;
            BindPageGridList(strWhere, tableName);
        }
        private void goto_Click(object sender, RoutedEventArgs e)
        {
            pageIndex = int.Parse(this.goPage.Text);
            BindPageGridList(strWhere, tableName);
        }


        //设置查询数据
        private string GetSqlWhere()
        {
            //查询条件  
            string strReturnWhere = string.Empty;


            return strReturnWhere;

        }

        //表格增加，删除，更新，改动
        private void btnRef_Click(object sender, RoutedEventArgs e)
        {

            if (isNewRow > 0)
            {

            }
            this.dataTable.CloseEditor();



            //获取查询条件  
            strWhere = GetSqlWhere();
            BindPageGridList(strWhere, tableName);


        }
        private void btnDel_Click(object sender, RoutedEventArgs e)
        {

            int[] iSelectRow = this.dataGrid.GetSelectedRowHandles();

            if (System.Windows.Forms.MessageBox.Show("你确定要删除选中的记录吗？", "删除提示", MessageBoxButtons.YesNo,
             MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2, 0, false) == System.Windows.Forms.DialogResult.Yes)
            {
                for (int j = 0; j < iSelectRow.Length; j++)
                {
                    Console.WriteLine("Int64.Parse:{0}", long.Parse(this.dataGrid.GetCellDisplayText(iSelectRow[j]-j, dataGrid.Columns[1]))); 
                    
                    OperateOracle.Delete(tableName, long.Parse(this.dataGrid.GetCellDisplayText(iSelectRow[j]-j, dataGrid.Columns[1])));


                 
                    this.dataTable.DeleteRow(iSelectRow[j]-j);

                   // if (iSelectRow[j] == row) { isNewRow -= 1; }
                }
            }
            BindPageGridList(strWhere, tableName);
        }
        private void btnAdd_Click(object sender, RoutedEventArgs e)
        {
            pageIndex = pageCount;
            BindPageGridList(strWhere, tableName);
            if (isNewRow == 0)
            {
                this.dataTable.AddNewRow();
                isNewRow += 1;
                long a;

                long b = long.Parse(this.dataGrid.GetCellDisplayText((int)(endIndex - startIndex), dataGrid.Columns[1]));
                a = b + 1;
                OperateOracle.Insert(tableName, a);
                isNewRow -= 1;
                BindPageGridList(strWhere, tableName);
               
            }
            else
            {
                System.Windows.Forms.MessageBox.Show("请编辑新加行数据！");
            }

        }

        //private void dataTable_InitNewRow(object sender, DevExpress.Xpf.Grid.InitNewRowEventArgs e)
        //{
        //    pageIndex = pageCount;
        //    BindPageGridList(strWhere, tableName);

           
        //    // }
        //    // else
        //    // {
        //    //     System.Windows.Forms.MessageBox.Show("请编辑新加行数据！");
        //    // }

        //}
        private void dataTable_CellValueChanged(object sender, DevExpress.Xpf.Grid.CellValueChangedEventArgs e)
        {
            int b = 0;
            for (int i = 0; i < dataGrid.Columns.Count; i++)
            {
                if (dataGrid.Columns[i] == e.Column) { b = i; Console.WriteLine("b:{0}", b); }
            }

            if (e.RowHandle < row)
            {

                string  a = this.dataGrid.GetCellDisplayText(e.RowHandle, dataGrid.Columns[b]);
                Int64 c = Int64.Parse(this.dataGrid.GetCellDisplayText(e.RowHandle, dataGrid.Columns[1]));
                Console.WriteLine("a:{0}", a);

                OperateOracle.Update(tableName, a, c, b-1);


            }
            else
            {

            }
        }


        //选择不同的表，确定表名
        private void comboBoxInsert_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            switch (this.comboBoxInsert.SelectedIndex)
            {
                case 0: tableName = "REGION"; break;
                case 1: tableName = "HUINFO"; break;
                case 2: tableName = "PEOPLEINFO"; break;
                case 3: tableName = "COLLECTION"; break;
                case 4: tableName = "MEDIA"; break;
                case 5: tableName = "SHIJIPINKUN"; break;
                case 6: tableName = "YIDIBANQIAN"; break;
                case 7: tableName="PERSON";break;
                case 8: tableName="HOUSEHOLDER";break;
                case 9:tableName="PERSONTENSE";break;
                case 10:tableName="HOUSEHOLDERTENSE";break;
                default: tableName = "REGION"; break;
            }
            if (this.comboBoxInsert.SelectedIndex > 6) { this.batchInsert.IsEnabled = false; this.button3.IsEnabled = true; }
            if (this.comboBoxInsert.SelectedIndex < 6) { this.batchInsert.IsEnabled = true; this.button3.IsEnabled = false; }
        }

        //条件查询
        private void dXTabControl1_SelectionChanged(object sender, DevExpress.Xpf.Core.TabControlSelectionChangedEventArgs e)
        {
            if (this.dXTabControl1.SelectedContainer == tiaojianSearch) 
            {
                List<string> listShi =OperateOracle.GetShi();
                this.comboBox1.ItemsSource = listShi;
            }
            
        }
        private void comboBox1_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
           
            string shiStr = (string)this.comboBox1.SelectedValue;
            Console.WriteLine("shiStr:{0}", shiStr);
            List<string> listXian = OperateOracle.GetXian(shiStr);
            this.comboBox2.ItemsSource = listXian;
        }        
        private void comboBox2_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            string shiStr = (string)this.comboBox1.SelectedValue;
         
            string xianStr = (string)this.comboBox2.SelectedValue;
            
            List<string> listZhen = OperateOracle.GetZhen(shiStr,xianStr);
            this.comboBox3.ItemsSource = listZhen;
        }
        private void comboBox3_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            string zhenStr = (string)this.comboBox3.SelectedValue;
            string shiStr = (string)this.comboBox1.SelectedValue;
            string xianStr = (string)this.comboBox2.SelectedValue;
            Console.WriteLine("shiStr:{0}", shiStr);
            Console.WriteLine("xianStr:{0}", xianStr);
           
            Console.WriteLine("zhenStr:{0}", zhenStr);
            List<string> listCun = OperateOracle.GetCun(shiStr, xianStr,zhenStr);
            this.comboBox4.ItemsSource = listCun;

        }
        private void Search1_Click(object sender, RoutedEventArgs e)
        {
           
            string shiStr = (string)this.comboBox1.SelectedValue;
            string xianStr = (string)this.comboBox2.SelectedValue;
            string zhenStr = (string)this.comboBox3.SelectedValue;
            string cunStr = (string)this.comboBox4.SelectedValue;
            long rid=  OperateOracle.GetRegionID(shiStr, xianStr, zhenStr, cunStr);
            strWhere += string.Format(" REGIONID={0}", rid);
            BindPageGridList(strWhere, tableName);
        }
       
        
        //数据批量导入
        private void batchInsert_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Title = "输入文件";
            openFileDialog.Filter = "excel文件|*.xls|excel文件|*.xlsx|所有文件|*.*";
            openFileDialog.FileName = string.Empty;
            openFileDialog.FilterIndex = 1;
            openFileDialog.RestoreDirectory = true;
            openFileDialog.DefaultExt = "xls";

            DialogResult result = openFileDialog.ShowDialog();
            if (result == System.Windows.Forms.DialogResult.Cancel)
            {
                return;
            }
            if (result == System.Windows.Forms.DialogResult.OK)
            {
                string inputFile = openFileDialog.FileName;
               DataSet ds= OperateOracle.ExcelToDataTable(inputFile,true,tableName);
               DataSet ds1 = OperateOracle.removeSame(ds,tableName);             
               OperateOracle.BulkToDB(tableName, ds1);

            }
          
        }

        private void button3_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Title = "输入文件";
            openFileDialog.Filter = "excel文件|*.xls|excel文件|*.xlsx|所有文件|*.*";
            openFileDialog.FileName = string.Empty;
            openFileDialog.FilterIndex = 1;
            openFileDialog.RestoreDirectory = true;
            openFileDialog.DefaultExt = "xls";

            DialogResult result = openFileDialog.ShowDialog();
            if (result == System.Windows.Forms.DialogResult.Cancel)
            {
                return;
            }
            if (result == System.Windows.Forms.DialogResult.OK)
            {
                string inputFile = openFileDialog.FileName;


               DataSet ds = OperateOracle.ExcelToDataTable1(inputFile, true, tableName);
               List<DataSet> ds1ist = OperateOracle.yibiaoruku(ds);
             
               OperateOracle.BulkToDB("PERSON", ds1ist[0]);
               OperateOracle.BulkToDB("PERSONTENSE", ds1ist[1]);
               OperateOracle.BulkToDB("HOUSEHOLDER", ds1ist[2]);
               OperateOracle.BulkToDB("HOUSEHOLDERTENSE", ds1ist[3]);
            }
        }

       //表所有数据导出
        private void button2_Click(object sender, RoutedEventArgs e)
        {


            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Title = "输出文件";
            saveFileDialog.Filter = "excel文件|*.xls|excel文件|*.xlsx|所有文件|*.*";
            saveFileDialog.FileName = string.Empty;
            saveFileDialog.FilterIndex = 1;
            saveFileDialog.RestoreDirectory = true;
            saveFileDialog.DefaultExt = "xls";

            DialogResult result = saveFileDialog.ShowDialog();
            if (result == System.Windows.Forms.DialogResult.Cancel)
            {
                return;
            }
            if (result == System.Windows.Forms.DialogResult.OK)
            {
                string outputFile = saveFileDialog.FileName;
               
                DataSet ds = OperateOracle.DBToDataTable(tableName);
                OperateOracle.DataTableToExcel(ds, outputFile);
            }
        }


        //查询显示表格导出
        private void btnOut_Click(object sender, RoutedEventArgs e)
        {

            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Title = "输出文件";
            saveFileDialog.Filter = "excel文件|*.xls|excel文件|*.xlsx|所有文件|*.*";
            saveFileDialog.FileName = string.Empty;
            saveFileDialog.FilterIndex = 1;
            saveFileDialog.RestoreDirectory = true;
            saveFileDialog.DefaultExt = "xls";

            DialogResult result = saveFileDialog.ShowDialog();
            if (result == System.Windows.Forms.DialogResult.Cancel)
            {
                return;
            }
            if (result == System.Windows.Forms.DialogResult.OK)
            {
                string outputFile = saveFileDialog.FileName;
                DataSet ds = new DataSet();
               DataTable dt1= OperateOracle.GetListByPage(tableName, strWhere, "", 0, row);
                ds.Tables.Add(dt1);
                OperateOracle.DataTableToExcel(ds, outputFile);
            }
        }
         
        //拖动表头移动窗体，之前的window_MouseMove影响表格控件
        private void header_MouseMove(object sender, System.Windows.Input.MouseEventArgs e)
        {
            if (e.LeftButton == MouseButtonState.Pressed)
            {
                         
                Window.DragMove();
            }
        }

      





        void CreateEngineControls()
        {
            mapControl = new AxMapControl();
            mapHost.Child = mapControl;
            toolbarControl = new AxToolbarControl();
            toolbarHost.Child = toolbarControl;
        }
        private void SetControlsProperty()
        {
            //设置控件之间的绑定关系


            ((System.ComponentModel.ISupportInitialize)toolbarControl).BeginInit();

            ((System.ComponentModel.ISupportInitialize)toolbarControl).EndInit();
            toolbarControl.SetBuddyControl(mapControl);

            //添加命令按钮到toolbarControl
            toolbarControl.AddItem("esriControls.ControlsOpenDocCommand");
            toolbarControl.AddItem("esriControls.ControlsAddDataCommand");
            toolbarControl.AddItem("esriControls.ControlsSaveAsDocCommand");
            toolbarControl.AddItem("esriControls.ControlsMapNavigationToolbar");
            toolbarControl.AddItem("esriControls.ControlsMapIdentifyTool");

            //设置空间属性
            toolbarControl.BackColor = System.Drawing.Color.FromArgb(245, 245, 220);


            //挂接事件

        }
        void LoadMap()
        {
            string strMxd = @"G:\sd.mxd";
            if (mapControl.CheckMxFile(strMxd))
                mapControl.LoadMxFile(strMxd);

        }
        private void min_Click(object sender, RoutedEventArgs e)
        {
            this.WindowState = WindowState.Minimized;
        }
        //窗体关闭
        private void close_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
            System.Environment.Exit(System.Environment.ExitCode);
        }
        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            SetControlsProperty();
            //LoadMap(); 

        }
        private void findButton_Click(object sender, RoutedEventArgs e)
        {
            //打开查询的数据表格等
        }
        private void max_Click(object sender, RoutedEventArgs e)
        {
            this.WindowState = WindowState.Maximized;
        }
        private void Button_Click(object sender, RoutedEventArgs e)
        {

        }
        public DataTable loadData()
        {

            Excel.Application xApp = new Excel.ApplicationClass();  //Excel.Application xApp是excel的应用程序
            xApp.Visible = false;                                       //ture or false是打不打开excel进行显示的意思。
            Excel.Workbook xBook = xApp.Workbooks._Open(@"C:/Users/Administrator/Desktop/1.xlsx", System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value);
            Excel.Worksheet xSheet = (Excel.Worksheet)xBook.Sheets[1]; //读取sheet[n],也就是第几张表
            Excel.Range i = (Excel.Range)xSheet.Cells[1, 1];       //读取位置cell[i,j]的值
            //Console.WriteLine(i.Value2);                         //输出值
            int RowCount = xSheet.UsedRange.Rows.Count;
            int ColumnsCount = xSheet.UsedRange.Columns.Count;
            string a = Convert.ToString(i.Value2);
            DataTable mytable = new DataTable();
            mytable.Columns.Add("Id", typeof(int));
            mytable.Columns.Add("AreaName", typeof(string));
            mytable.Columns.Add("总村数", typeof(string));
            mytable.Columns.Add("总贫困村数", typeof(string));
            mytable.Columns.Add("总户数", typeof(string));
            mytable.Columns.Add("总贫困户数", typeof(string));
            mytable.Columns.Add("总人数", typeof(string));
            mytable.Columns.Add("总贫困人数", typeof(string));

            mytable.Columns["Id"].AutoIncrement = true;
            for (int b = 1; b < RowCount; b++)
            {
                mytable.Rows.Add();
            }
            for (int k = 5; k <= RowCount + 1; k++)
            {
                for (int c = 2; c <= ColumnsCount; c++)
                {
                    Microsoft.Office.Interop.Excel.Range thecell = (Microsoft.Office.Interop.Excel.Range)xSheet.Cells[k - 1, c];
                    string s = Convert.ToString(thecell.Value2);
                    //string s = (string)thecell.Value2;
                    if (s != "")
                        mytable.Rows[k - 4][c - 1] = s;
                }

            }


            this.chartControl1.DataSource = mytable.DefaultView;


            PieSeries2D series = this.chartControl1.Diagram.Series[0] as PieSeries2D;

            xSheet = null;
            xBook = null;
            xApp.Quit(); //这一句是非常重要的，否则Excel对象不能从内存中退出             
            xApp = null;
            return mytable;
        }
        private void datatable_Loaded(object sender, RoutedEventArgs e)
        {
            //newdtb= loadData();
            //  this.datatable.ItemsSource = newdtb; 

        }

        private void Change_Click(object sender, RoutedEventArgs e)
        {
            DataChange dataChange = new DataChange();
            dataChange.Show();
        }

      
       


       
    }
}
