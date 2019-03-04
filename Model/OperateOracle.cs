using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Data;
using Oracle.DataAccess.Client;
using Oracle.DataAccess.Types;
using System.IO;

using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

using NPOI.XWPF.UserModel;
using NPOI.OpenXmlFormats.Wordprocessing;


namespace Model.OperateOracle
{
    class OperateOracle
    {


        static List<string> strList = null;

        /// <summary>
        /// 获取查询结果总数  
        /// </summary>
        /// <param name="strWhere"></param>
        /// <returns></returns>
        public static long GetRecordCount(string tableName, string strWhere)
        {
            string strSql = "";
            strSql += string.Format("select count(1) from {0}", tableName);
            if (strWhere.Trim() != "")
            {
                strSql += " where " + strWhere;
            }
            object obj = null;
            //string connString = "Provider=OraOLEDB.Oracle.1;User ID=SYSTEM;Password=123;Data Source=ORCL;";
            string connString = @"Data Source=(DESCRIPTION =(ADDRESS = (PROTOCOL = TCP)(HOST = localhost)(PORT = 1521))(CONNECT_DATA =
      (SERVER = DEDICATED)
      (SERVICE_NAME = orcl)
    )
  ); User Id = SYSTEM; Password =123;";
            OracleConnection conn = new OracleConnection(connString);

            try
            {

                conn.Open();
                OracleCommand cmd = conn.CreateCommand();
                cmd.CommandText = strSql;
                obj = cmd.ExecuteScalar();


            }
            catch (Exception ex)
            {
                MessageBox.Show("查询记录数错误：" + ex.Message.ToString());
            }
            finally
            {
                conn.Close();
            }
            if (obj == null)
            { return 0; }
            else
            { return Convert.ToInt64(obj); }

        }

        /// <summary>
        /// 获取查询数据
        /// </summary>
        /// <param name="tableName">表名</param>
        /// <param name="strWhere">查询条件</param>
        /// <param name="orderby">排序字段</param>
        /// <param name="startIndex">开始行</param>
        /// <param name="endIndex">结束行</param>
        /// <returns>获取所查询的表</returns>
        public static DataTable GetListByPage(string tableName,string strWhere, string orderby, long startIndex, long endIndex)
        {
            string connString = @"Data Source=(DESCRIPTION =(ADDRESS = (PROTOCOL = TCP)(HOST = localhost)(PORT = 1521))(CONNECT_DATA =
      (SERVER = DEDICATED)
      (SERVICE_NAME = orcl)
    )
  ); User Id = SYSTEM; Password =123;";
            OracleConnection conn = new OracleConnection(connString);
            DataTable dt = new DataTable();
             string strSql = "";
             strSql = string.Format("SELECT * FROM {0}", tableName);
            try
            {
                conn.Open();
                OracleCommand cmd1 = conn.CreateCommand();
                cmd1.CommandText = strSql;
                OracleDataReader rd1 = cmd1.ExecuteReader();
                for (int i = 0; i < rd1.FieldCount; i++) { dt.Columns.Add(rd1.GetName(i), rd1.GetFieldType(i)); }


           
            // strSql = "SELECT * FROM ( SELECT shijipinkun.* , ROWNUM   RN FROM  shijipinkun WHERE ROWNUM <= 10 order by col0 asc) WHERE RN >= 1 ";
            strSql = "SELECT * FROM ( ";
            strSql += string.Format("SELECT   {0}.*  FROM  {1} ", tableName, tableName);
            if (!string.IsNullOrEmpty(strWhere.Trim()))
            {
                strSql += " WHERE " + strWhere;
            }
            if (!string.IsNullOrEmpty(orderby.Trim()))
            {
                strSql += " order by " + orderby + " asc";
            }
            else
            {
                strSql += string.Format(" ORDER BY {0} ASC", dt.Columns[0].ColumnName);
            }
            strSql += string.Format(" ) WHERE {0} BETWEEN {1} and {2}",dt.Columns[0].ColumnName, startIndex, endIndex);



            OracleCommand cmd = conn.CreateCommand();
            cmd.CommandText = strSql;
            OracleDataReader rd = cmd.ExecuteReader();
           // for (int i = 0; i < rd.FieldCount; i++) { dt.Columns.Add(rd.GetName(i), rd.GetFieldType(i)); }
                while (rd.Read())//读取数据，如果返回为false的话，就说明到记录集的尾部了  
                {
                    DataRow r = dt.NewRow();

                  
                    for (int i = 0; i < rd.FieldCount; i++)
                    {
                        if(rd.GetFieldType(i)==typeof(string))
                        {
                            if (!rd.IsDBNull(i))
                            {
                                r[dt.Columns[i]] = rd.GetString(i);
                            }
                            else { r[dt.Columns[i]] = DBNull.Value; }
                        }
                        else if (rd.GetFieldType(i) == typeof(DateTime)) 
                        {
                            if (!rd.IsDBNull(i))
                            {
                                r[dt.Columns[i]] = rd.GetGuid(i);
                            }
                            else { r[dt.Columns[i]] = DBNull.Value; }
                        }
                        else  //if (rd.GetFieldType(i) == typeof(Int32) || rd.GetFieldType(i) == typeof(Int64) || rd.GetFieldType(i) == typeof(NumberFormat))
                        {

                            if (!rd.IsDBNull(i))
                            {
                                r[dt.Columns[i]] = rd.GetValue(i);
                            }
                            else { r[dt.Columns[i]] = DBNull.Value; }
                        }
                       
                    }
                    dt.Rows.Add(r);
                }

                // MessageBox.Show("读取表结束");  //弹出显示
                rd.Close();//关闭reader.这是一定要写的  

            }
            catch (Exception ex)
            {
                MessageBox.Show("读取表记录错误：" + ex.Message.ToString());
            }
            finally
            {
                conn.Close();
            }
            return dt;
        }



        //查询数据库表格列名及数据类型
        public static DataTable GetDataTable(string tableName)
        {
            DataTable dt = new DataTable();          
            string strSql = "";
            // strSql = "SELECT * FROM ( SELECT shijipinkun.* , ROWNUM   RN FROM  shijipinkun WHERE ROWNUM <= 10 order by col0 asc) WHERE RN >= 1 ";
           // strSql +="SELECT COLUMN_NAME,data_type FROM all_tab_columns where  TABLE_NAME =";
            strSql += "SELECT * FROM "; 
            strSql += string.Format("{0}", tableName) ;

            string connString = @"Data Source=(DESCRIPTION =
(ADDRESS = (PROTOCOL = TCP)(HOST = localhost)(PORT = 1521))
(CONNECT_DATA =(SERVER = DEDICATED)(SERVICE_NAME = orcl) ));
User Id = SYSTEM; Password =123;";
            OracleConnection conn = new OracleConnection(connString);           

            try
            {
                conn.Open();
                OracleCommand cmd = conn.CreateCommand();
                cmd.CommandText = strSql;
                OracleDataReader rd = cmd.ExecuteReader();
                for (int i = 0; i < rd.FieldCount; i++) { dt.Columns.Add(rd.GetName(i), rd.GetFieldType(i)); }

                              
                rd.Close();//关闭reader.这是一定要写的  

            }
            catch (Exception ex)
            {
                MessageBox.Show("读取表列信息错误：" + ex.Message.ToString());
            }
            finally
            {
                conn.Close();
            }
            return dt;
            
        }


       //数据转换（还没用）
        private static OracleDbType GetOracleDbType(object value)
        {
            OracleDbType dataType = OracleDbType.Object;
            if (value is string[])
            {
                dataType = OracleDbType.Varchar2;
            }
            else if (value is DateTime[])
            {
                dataType = OracleDbType.TimeStamp;
            }
            else if (value is int[] || value is short[])
            {
                dataType = OracleDbType.Int32;
            }
            else if (value is long[])
            {
                dataType = OracleDbType.Int64;
            }
            else if (value is decimal[] || value is double[] || value is float[])
            {
                dataType = OracleDbType.Decimal;
            }
            else if (value is Guid[])
            {
                dataType = OracleDbType.Varchar2;
            }
            else if (value is bool[] || value is Boolean[])
            {
                dataType = OracleDbType.Byte;
            }
            else if (value is byte[])
            {
                dataType = OracleDbType.Blob;
            }
            else if (value is char[])
            {
                dataType = OracleDbType.Char;
            }
            return dataType;
        }



        //添加，删除，更新操作
        public static int Insert(string tableName,long a )
        {
           // string connString = "Provider=OraOLEDB.Oracle.1;User ID=SYSTEM;Password=123;Data Source=orcl;";
           // OleDbConnection conn = new OleDbConnection(connString);
            string connString = @"Data Source=(DESCRIPTION =(ADDRESS = (PROTOCOL = TCP)(HOST = localhost)(PORT = 1521))(CONNECT_DATA =
      (SERVER = DEDICATED)
      (SERVICE_NAME = orcl)
    )
  ); User Id = SYSTEM; Password =123;";
            OracleConnection conn = new OracleConnection(connString);
            try
            {
                conn.Open();//打开指定的连接
                OracleCommand cmd = conn.CreateCommand();
                cmd.CommandText = "select * from " + tableName;
                OracleDataReader rd = cmd.ExecuteReader();
                List<string> nameList=new List<string>();
                for(int i=0;i<rd.FieldCount;i++){nameList.Add(rd.GetName(i));}




                string sqlStr = string.Format("INSERT INTO {0} ( {1} ", tableName, nameList[0]);
                for (int i = 1; i < rd.FieldCount; i++) { sqlStr += string.Format(" ,{0} ", nameList[i]); }
                sqlStr += " ) VALUES ( ";
                sqlStr += string.Format(" {0} ", a);
                for (int i = 1; i < rd.FieldCount; i++) { sqlStr += string.Format(" ,{0} ","null"); }
                sqlStr += ")";



                cmd.CommandText = sqlStr;
                int result = cmd.ExecuteNonQuery();
                if (result > 0)
                {
                    MessageBox.Show("插入记录成功");  //弹出显示
                }




            }
            catch (Exception ex)
            {
                MessageBox.Show("插入记录错误：" + ex.Message.ToString());
            }
            finally
            {

                conn.Close();
            }





            return 1;
        }
        public static int Delete(string tableName, long id)
        {
            //string connString = "Provider=OraOLEDB.Oracle.1;User ID=SYSTEM;Password=123;Data Source=orcl;";
           // OleDbConnection conn = new OleDbConnection(connString);
            string connString = @"Data Source=(DESCRIPTION =(ADDRESS = (PROTOCOL = TCP)(HOST = localhost)(PORT = 1521))(CONNECT_DATA =
      (SERVER = DEDICATED)
      (SERVICE_NAME = orcl)
    )
  ); User Id = SYSTEM; Password =123;";
            OracleConnection conn = new OracleConnection(connString);
            try
            {
                conn.Open();//打开指定的连接


                string sqlStr = string.Format("delete from {0} where id={1}", tableName, id);

                OracleCommand cmd = new OracleCommand(sqlStr, conn);
                int result = cmd.ExecuteNonQuery();
                if (result > 0)
                {
                    MessageBox.Show("删除成功");  //弹出显示
                }
                
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString());
            }
            finally
            {
                conn.Close();
            }
            
            return 1;
        }
        public static int Update(string tableName,string da, long rowid, int colid)
        {
           
            string connString = @"Data Source=(DESCRIPTION =(ADDRESS = (PROTOCOL = TCP)(HOST = localhost)(PORT = 1521))(CONNECT_DATA =
      (SERVER = DEDICATED)
      (SERVICE_NAME = orcl)
    )
  ); User Id = SYSTEM; Password =123;";
            OracleConnection conn = new OracleConnection(connString);
            try
            {
                conn.Open();//打开指定的连接

                OracleCommand cmd = conn.CreateCommand();
                cmd.CommandText = "select * from " + tableName ;
                OracleDataReader rd = cmd.ExecuteReader();
                string columName = rd.GetName(colid);


                string sqlStr = string.Format("update {0} set {1}=\'{2}\' ", tableName, columName, da);
                
                sqlStr += " WHERE id=" + rowid;
                cmd.CommandText = sqlStr;
                
               
                int result = cmd.ExecuteNonQuery();
                if (result > 0)
                {
                    MessageBox.Show("更新记录成功");  //弹出显示
                }




            }
            catch (Exception ex)
            {
                MessageBox.Show("更新记录错误：" + ex.Message.ToString());
            }
            finally
            {
                conn.Close();
            }





            return 1;
        }

        //将错误列表输出文档
        public static int printWord(List<string> str)
        {
            string fileName = string.Format("{0}.doc", DateTime.Now.ToString("yyyyMMddHHmmssfff"));
            int result = 0;

            //创建document对象  
            XWPFDocument doc = new XWPFDocument();



            //创建段落对象  
            XWPFParagraph p1 = doc.CreateParagraph();
            p1.Alignment = ParagraphAlignment.CENTER;//字体居左  

            //创建run对象  
            //本节提到的所有样式都是基于XWPFRun的，  
            //你可以把XWPFRun理解成一小段文字的描述对象，  
            //这也是Word文档的特征，即文本描述性文档。  

            XWPFRun runTitle = p1.CreateRun();
            runTitle.IsBold = true;
            runTitle.SetText("导入数据检查报告");
            runTitle.FontSize = 16;
            runTitle.SetFontFamily("宋体", FontCharRange.None);//设置雅黑字体



            for (int i = 0; i < str.Count; i++)
            {
                XWPFParagraph p2 = doc.CreateParagraph();
                //段落对其方式为居中  
                p2.Alignment = ParagraphAlignment.LEFT;
                XWPFRun r2 = p2.CreateRun();//向该段落中添加文字  
                r2.FontSize = 12;//设置大小  
                r2.FontFamily = "宋体"; //设置字体  
                r2.SetText(string.Format("{0}.",i) + str[i]);

            }




            string docPath = "D:"+ fileName;
            //  if (!Directory.Exists(docPath)) { Directory.CreateDirectory(docPath); }
            FileStream out1 = new FileStream(docPath, FileMode.OpenOrCreate, FileAccess.ReadWrite);
            doc.Write(out1);
            out1.Close();


            result = 1;
            MessageBox.Show("成功");
            return result;
        }
        //数据查重，记录覆盖信息
        public static DataSet removeSame(DataSet ds,string tableName)
        {
            
            DataSet ds1 = new DataSet();


            string connStr = @"Data Source=(DESCRIPTION =(ADDRESS = (PROTOCOL = TCP)(HOST = localhost)(PORT = 1521))(CONNECT_DATA =
      (SERVER = DEDICATED)(SERVICE_NAME = orcl))); User Id = SYSTEM; Password =123;";
            OracleConnection conn = new OracleConnection(connStr);
           

            try
            {
                conn.Open();//打开指定的连接
               // MessageBox.Show(conn.State.ToString());
                int j=0;
                foreach (DataTable dt in ds.Tables)
                {   j++;int m=0;
                    DataTable dt1 = dt.Copy();
                    foreach (DataRow dr in dt.Rows)
                    {m++;
                    string sqlStr = "";
                    sqlStr += string.Format("UPDATE {0} SET ", tableName);
                    sqlStr += string.Format(" {0}=\'{1}\' ", dt1.Columns[0].ColumnName.ToUpper(), dr[dt.Columns[0]]);
                        for (int i = 1; i < dt1.Columns.Count; i++)
                        {
                            sqlStr += string.Format(" ,{0}=\'{1}\' ", dt1.Columns[i].ColumnName.ToUpper(), dr[dt.Columns[i]]);
                        }
                        if (dr["DateTimeIndex"] != "")
                        {
                            sqlStr += string.Format(" WHERE ID={0} and DATETIMEINDEX=\'{1}\' ", dr["ID"], dr["DateTimeIndex"]);
                        }
                        else {
                            sqlStr += string.Format(" WHERE ID={0}  ", dr["ID"]);
                        }
                        Console.WriteLine("id:{0}",dr["ID"]);
                        Console.WriteLine("DateTimeIndex:{0}", dr["DateTimeIndex"]);
                        Console.WriteLine("sqlStr{0}", sqlStr);
                        OracleCommand cmd = new OracleCommand(sqlStr, conn);
                        int result = cmd.ExecuteNonQuery();
                        if (result > 0)
                        {
                           //输出
                            strList.Add(string.Format("第{0}sheet第{1}行覆盖了表中同年数据！",j,m));
                            dt1.Rows.RemoveAt(m-1);
                        }


                    }
                    ds1.Tables.Add(dt1);


                }
                
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString());
            }
            finally
            {

                conn.Close();

            }

            //输出检查结果
           
            printWord(strList);
            return ds1;
        }

       

        //批量导入数据库
        public static DataSet ExcelToDataTable(string filePath, bool isColumnName,string tableName)
        {
            strList = new List<string>();
            DataSet ds = new DataSet();

            FileStream fs = null;
            DataColumn column = null;
            DataRow dataRow = null;
            IWorkbook workbook = null;
            ISheet sheet = null;
            IRow row = null;
            NPOI.SS.UserModel.ICell cell = null;
            int startRow = 0;
            try
            {
                using (fs = File.OpenRead(filePath))
                {
                    // 2007版本  
                    if (filePath.IndexOf(".xlsx") > 0)
                        workbook = new XSSFWorkbook(fs);
                    // 2003版本  
                    else if (filePath.IndexOf(".xls") > 0)
                        workbook = new HSSFWorkbook(fs);

                    if (workbook != null)
                    {

                        for (int k = 0; k < workbook.NumberOfSheets; k++)
                        {
                            DataTable dataTable = GetDataTable(tableName);
                            // string sheetName = workbook.GetNameAt(k).SheetName;

                            //  if (sheetName.Contains("$") && !sheetName.Replace("'", "").EndsWith("$"))//判断sheet是否有效
                            //  {
                            //     continue;
                            //  }

                            sheet = workbook.GetSheetAt(k);//读取第一个sheet，当然也可以循环读取每个sheet  


                            if (sheet != null)
                            {
                                long rowCount = sheet.LastRowNum;//总行数  
                                if (rowCount > 0)
                                {
                                    IRow firstRow = sheet.GetRow(0);//第一行  
                                    long cellCount = firstRow.LastCellNum;//列数  

                                    //构建datatable的列  
                                    if (isColumnName)
                                    {
                                        startRow = 1;//如果第一行是列名，则从第二行开始读取  
                                        for (int i = firstRow.FirstCellNum; i < cellCount; ++i)
                                        {
                                            cell = firstRow.GetCell(i);
                                            if (cell != null)
                                            {
                                                if (cell.StringCellValue != null)
                                                {
                                                    column = new DataColumn(cell.StringCellValue);
                                                  // dataTable.Columns.Add(column);
                                                }
                                            }
                                        }
                                    }
                                    else
                                    {
                                        for (int i = firstRow.FirstCellNum; i < cellCount; ++i)
                                        {
                                            column = new DataColumn("column" + (i + 1));
                                            //dataTable.Columns.Add(column);
                                        }
                                    }
                                   // dataTable.Columns[0].DataType = typeof(long);

                                    //填充行  
                                    for (int i = startRow; i <= rowCount; ++i)
                                    {
                                        row = sheet.GetRow(i);
                                        if (row == null) { strList.Add(string.Format("第{0}行数据为空！", i)); continue; }

                                        dataRow = dataTable.NewRow();
                                        for (int j = row.FirstCellNum; j < cellCount; ++j)
                                        {
                                            cell = row.GetCell(j);
                                            if (cell == null)
                                            {
                                                dataRow[j] = "";
                                                strList.Add(string.Format("第{0}行{1}列数据为空！", i, j));
                                            }
                                            else
                                            {
                                                //CellType(Unknown = -1,Numeric = 0,String = 1,Formula = 2,Blank = 3,Boolean = 4,Error = 5,)  
                                                switch (cell.CellType)
                                                {
                                                    case CellType.Blank:
                                                        dataRow[j] = "";
                                                        break;
                                                    case CellType.Numeric:
                                                        short format = cell.CellStyle.DataFormat;
                                                        //对时间格式（2015.12.5、2015/12/5、2015-12-5等）的处理  
                                                        if (format == 14 || format == 31 || format == 57 || format == 58)
                                                            dataRow[j] = cell.DateCellValue;
                                                        else
                                                        {
                                                            dataRow[j] = cell.NumericCellValue;
                                                            // Console.WriteLine("{0}", dataRow[j]);
                                                        }

                                                        break;
                                                    case CellType.String:
                                                        dataRow[j] = cell.StringCellValue;
                                                        break;
                                                    default: strList.Add(string.Format("第{0}行{1}列数据类型错误！", i, j)); break;
                                                }
                                            }
                                        }
                                        dataTable.Rows.Add(dataRow);
                                    }

                                    ds.Tables.Add(dataTable);

                                }//rowCount <= 0
                                else { MessageBox.Show("行数不能小于0！"); }
                            }//sheet != null
                            else { MessageBox.Show("sheet= null"); }

                        }


                    }//workbook != null
                    else { MessageBox.Show("workbook = null"); }

                }

            }
            catch (Exception ex)
            {
                if (fs != null)
                {
                    fs.Close();
                }
                MessageBox.Show(ex.Message);
                throw ex;
            }
            finally
            {
                if (fs != null)
                {
                    fs.Close();
                }
            }
            return ds;
        }




        public static DataSet ExcelToDataTable1(string filePath, bool isColumnName, string tableName)
        {
            strList = new List<string>();
            DataSet ds = new DataSet();

            FileStream fs = null;
            DataColumn column = null;
            DataRow dataRow = null;
            IWorkbook workbook = null;
            ISheet sheet = null;
            IRow row = null;
            NPOI.SS.UserModel.ICell cell = null;
            int startRow = 0;
            try
            {
                using (fs = File.OpenRead(filePath))
                {
                    // 2007版本  
                    if (filePath.IndexOf(".xlsx") > 0)
                        workbook = new XSSFWorkbook(fs);
                    // 2003版本  
                    else if (filePath.IndexOf(".xls") > 0)
                        workbook = new HSSFWorkbook(fs);

                    if (workbook != null)
                    {

                        for (int k = 0; k < workbook.NumberOfSheets; k++)
                        {
                            DataTable dataTable = new DataTable();
                            // string sheetName = workbook.GetNameAt(k).SheetName;

                            //  if (sheetName.Contains("$") && !sheetName.Replace("'", "").EndsWith("$"))//判断sheet是否有效
                            //  {
                            //     continue;
                            //  }

                            sheet = workbook.GetSheetAt(k);//读取第一个sheet，当然也可以循环读取每个sheet  


                            if (sheet != null)
                            {
                                long rowCount = sheet.LastRowNum;//总行数  
                                if (rowCount > 0)
                                {
                                    IRow firstRow = sheet.GetRow(0);//第一行  
                                    long cellCount = firstRow.LastCellNum;//列数  

                                    //构建datatable的列  
                                    if (isColumnName)
                                    {
                                        startRow = 1;//如果第一行是列名，则从第二行开始读取  
                                        for (int i = firstRow.FirstCellNum; i < cellCount; ++i)
                                        {
                                            cell = firstRow.GetCell(i);
                                            if (cell != null)
                                            {
                                                if (cell.StringCellValue != null)
                                                {
                                                    column = new DataColumn(cell.StringCellValue);
                                                    // dataTable.Columns.Add(column);
                                                }
                                            }
                                        }
                                    }
                                    else
                                    {
                                        for (int i = firstRow.FirstCellNum; i < cellCount; ++i)
                                        {
                                            column = new DataColumn("column" + (i + 1));
                                            //dataTable.Columns.Add(column);
                                        }
                                    }
                                    // dataTable.Columns[0].DataType = typeof(long);
                                    dataTable.Columns.Add(new DataColumn("RegionID", typeof(long)));
                                    dataTable.Columns.Add(new DataColumn("HuzhuName", typeof(string)));
                                    dataTable.Columns.Add(new DataColumn("ChengyuanName", typeof(string)));
                                    dataTable.Columns.Add(new DataColumn("Relationship", typeof(string)));
                                    dataTable.Columns.Add(new DataColumn("Poverty", typeof(string)));
                                    dataTable.Columns.Add(new DataColumn("HuzhuIDCard", typeof(string)));
                                    dataTable.Columns.Add(new DataColumn("ChengyuanIDCard", typeof(string)));
                                    dataTable.Columns.Add(new DataColumn("Phone", typeof(string)));
                                    dataTable.Columns.Add(new DataColumn("National", typeof(string)));
                                    dataTable.Columns.Add(new DataColumn("Culture", typeof(string)));
                                    dataTable.Columns.Add(new DataColumn("School", typeof(string)));
                                    dataTable.Columns.Add(new DataColumn("LaborCapacity", typeof(string)));
                                    dataTable.Columns.Add(new DataColumn("EngageState", typeof(string)));
                                    dataTable.Columns.Add(new DataColumn("EngageTime", typeof(double)));
                                    dataTable.Columns.Add(new DataColumn("LivingState", typeof(string)));
                                    dataTable.Columns.Add(new DataColumn("HealthState", typeof(string)));
                                    dataTable.Columns.Add(new DataColumn("Level", typeof(int)));                                    
                                    dataTable.Columns.Add(new DataColumn("Reason", typeof(int)));
                                    dataTable.Columns.Add(new DataColumn("Income", typeof(double)));
                                    dataTable.Columns.Add(new DataColumn("Relocate", typeof(int)));
                                    dataTable.Columns.Add(new DataColumn("BankName", typeof(string)));
                                    dataTable.Columns.Add(new DataColumn("CreditCard", typeof(string)));
                                    dataTable.Columns.Add(new DataColumn("ProjectID", typeof(string)));
                                    dataTable.Columns.Add(new DataColumn("RecordTime", typeof(DateTime)));
                                    //填充行  
                                    for (int i = startRow; i <= rowCount; ++i)
                                    {
                                        row = sheet.GetRow(i);
                                        if (row == null) { strList.Add(string.Format("第{0}行数据为空！", i)); continue; }

                                        dataRow = dataTable.NewRow();
                                        for (int j = row.FirstCellNum; j < cellCount; ++j)
                                        {
                                            cell = row.GetCell(j);
                                            if (cell == null)
                                            {
                                                dataRow[j] = DBNull.Value;
                                                strList.Add(string.Format("第{0}行{1}列数据为空！", i, j));
                                            }
                                            else
                                            {
                                                //CellType(Unknown = -1,Numeric = 0,String = 1,Formula = 2,Blank = 3,Boolean = 4,Error = 5,)  
                                                switch (cell.CellType)
                                                {
                                                    case CellType.Blank:
                                                        dataRow[j] = DBNull.Value;
                                                        break;
                                                    case CellType.Numeric:
                                                        short format = cell.CellStyle.DataFormat;
                                                        //对时间格式（2015.12.5、2015/12/5、2015-12-5等）的处理  

                                                        if (format == 14 || format == 31 || format == 57 || format == 58)
                                                        { dataRow[j] = cell.DateCellValue; }  
                                                        else
                                                        {
                                                           
                                                                dataRow[j] = cell.NumericCellValue;
                                                            
                                                          
                                                            // Console.WriteLine("{0}", dataRow[j]);
                                                        }
                                                        
                                                        break;
                                                    case CellType.String:                                                        
                                                          dataRow[j] = cell.StringCellValue; break;
                                                    default: strList.Add(string.Format("第{0}行{1}列数据类型错误！", i, j)); break;
                                                }
                                            }
                                        }
                                        dataTable.Rows.Add(dataRow);
                                    }

                                    ds.Tables.Add(dataTable);

                                }//rowCount <= 0
                                else { MessageBox.Show("行数不能小于0！"); }
                            }//sheet != null
                            else { MessageBox.Show("sheet= null"); }

                        }


                    }//workbook != null
                    else { MessageBox.Show("workbook = null"); }

                }

            }
            catch (Exception ex)
            {
                if (fs != null)
                {
                    fs.Close();
                }
                MessageBox.Show(ex.Message);
                throw ex;
            }
            finally
            {
                if (fs != null)
                {
                    fs.Close();
                }
            }
            return ds;
        }



        #region//统计信息表一表入库
        public static DataSet ExcelToDataTable2(string filePath, bool isColumnName, string tableName)
        {
            strList = new List<string>();
            DataSet ds = new DataSet();

            FileStream fs = null;
            DataColumn column = null;
            DataRow dataRow = null;
            IWorkbook workbook = null;
            ISheet sheet = null;
            IRow row = null;
            NPOI.SS.UserModel.ICell cell = null;
            int startRow = 0;
            try
            {
                using (fs = File.OpenRead(filePath))
                {
                    // 2007版本  
                    if (filePath.IndexOf(".xlsx") > 0)
                        workbook = new XSSFWorkbook(fs);
                    // 2003版本  
                    else if (filePath.IndexOf(".xls") > 0)
                        workbook = new HSSFWorkbook(fs);

                    if (workbook != null)
                    {

                        for (int k = 0; k < workbook.NumberOfSheets; k++)
                        {
                            DataTable dataTable = new DataTable();
                            // string sheetName = workbook.GetNameAt(k).SheetName;

                            //  if (sheetName.Contains("$") && !sheetName.Replace("'", "").EndsWith("$"))//判断sheet是否有效
                            //  {
                            //     continue;
                            //  }

                            sheet = workbook.GetSheetAt(k);//读取第一个sheet，当然也可以循环读取每个sheet  


                            if (sheet != null)
                            {
                                long rowCount = sheet.LastRowNum;//总行数  
                                if (rowCount > 0)
                                {
                                    IRow firstRow = sheet.GetRow(0);//第一行  
                                    long cellCount = firstRow.LastCellNum;//列数  

                                    //构建datatable的列  
                                    if (isColumnName)
                                    {
                                        startRow = 1;//如果第一行是列名，则从第二行开始读取  
                                        for (int i = firstRow.FirstCellNum; i < cellCount; ++i)
                                        {
                                            cell = firstRow.GetCell(i);
                                            if (cell != null)
                                            {
                                                if (cell.StringCellValue != null)
                                                {
                                                    column = new DataColumn(cell.StringCellValue);
                                                    // dataTable.Columns.Add(column);
                                                }
                                            }
                                        }
                                    }
                                    else
                                    {
                                        for (int i = firstRow.FirstCellNum; i < cellCount; ++i)
                                        {
                                            column = new DataColumn("column" + (i + 1));
                                            //dataTable.Columns.Add(column);
                                        }
                                    }
                                    // dataTable.Columns[0].DataType = typeof(long);
                                    dataTable.Columns.Add(new DataColumn("NAME", typeof(string)));
                                    dataTable.Columns.Add(new DataColumn("IDCARD", typeof(string)));
                                    dataTable.Columns.Add(new DataColumn("REGIONID", typeof(string)));
                                    dataTable.Columns.Add(new DataColumn("HLATITUDE", typeof(string)));
                                    dataTable.Columns.Add(new DataColumn("HLONGITUDE", typeof(string)));
                                    dataTable.Columns.Add(new DataColumn("GATHERER", typeof(string)));
                                    dataTable.Columns.Add(new DataColumn("COLLECTIONTIME", typeof(string)));
                                    dataTable.Columns.Add(new DataColumn("MARK", typeof(string)));
                                    dataTable.Columns.Add(new DataColumn("HOUSEORDER", typeof(string)));
                                    dataTable.Columns.Add(new DataColumn("FILENAME", typeof(string)));
                                    dataTable.Columns.Add(new DataColumn("AZIMUTH", typeof(string)));
                                    dataTable.Columns.Add(new DataColumn("SATELLITES", typeof(int)));
                                    dataTable.Columns.Add(new DataColumn("LATITUDE", typeof(string)));
                                    dataTable.Columns.Add(new DataColumn("LONGITUDE", typeof(string)));
                                    
                                    //填充行  
                                    for (int i = startRow; i <= rowCount; ++i)
                                    {
                                        row = sheet.GetRow(i);
                                        if (row == null) { strList.Add(string.Format("第{0}行数据为空！", i)); continue; }

                                        dataRow = dataTable.NewRow();
                                        for (int j = row.FirstCellNum; j < cellCount; ++j)
                                        {
                                            cell = row.GetCell(j);
                                            if (cell == null)
                                            {
                                                dataRow[j] = DBNull.Value;
                                                strList.Add(string.Format("第{0}行{1}列数据为空！", i, j));
                                            }
                                            else
                                            {
                                                //CellType(Unknown = -1,Numeric = 0,String = 1,Formula = 2,Blank = 3,Boolean = 4,Error = 5,)  
                                                switch (cell.CellType)
                                                {
                                                    case CellType.Blank:
                                                        dataRow[j] = DBNull.Value;
                                                        break;
                                                    case CellType.Numeric:
                                                        short format = cell.CellStyle.DataFormat;
                                                        //对时间格式（2015.12.5、2015/12/5、2015-12-5等）的处理  

                                                        if (format == 14 || format == 31 || format == 57 || format == 58)
                                                        { dataRow[j] = cell.DateCellValue; }
                                                        else
                                                        {

                                                            dataRow[j] = cell.NumericCellValue;


                                                            // Console.WriteLine("{0}", dataRow[j]);
                                                        }

                                                        break;
                                                    case CellType.String:
                                                        dataRow[j] = cell.StringCellValue; break;
                                                    default: strList.Add(string.Format("第{0}行{1}列数据类型错误！", i, j)); break;
                                                }
                                            }
                                        }
                                        dataTable.Rows.Add(dataRow);
                                    }

                                    ds.Tables.Add(dataTable);

                                }//rowCount <= 0
                                else { MessageBox.Show("行数不能小于0！"); }
                            }//sheet != null
                            else { MessageBox.Show("sheet= null"); }

                        }


                    }//workbook != null
                    else { MessageBox.Show("workbook = null"); }

                }

            }
            catch (Exception ex)
            {
                if (fs != null)
                {
                    fs.Close();
                }
                MessageBox.Show(ex.Message);
                throw ex;
            }
            finally
            {
                if (fs != null)
                {
                    fs.Close();
                }
            }
            return ds;
        }
        public static long getDtMaxID(string tableName)
        {
            string strSql = "";
            strSql += string.Format("select max({0})  from {1}", GetDataTable(tableName).Columns[0].ColumnName, tableName);

            object obj = null;
            //string connString = "Provider=OraOLEDB.Oracle.1;User ID=SYSTEM;Password=123;Data Source=ORCL;";
            string connString = @"Data Source=(DESCRIPTION =(ADDRESS = (PROTOCOL = TCP)(HOST = localhost)(PORT = 1521))(CONNECT_DATA =
      (SERVER = DEDICATED)
      (SERVICE_NAME = orcl)
    )
  ); User Id = SYSTEM; Password =123;";
            OracleConnection conn = new OracleConnection(connString);

            try
            {

                conn.Open();
                OracleCommand cmd = conn.CreateCommand();
                cmd.CommandText = strSql;
                obj = cmd.ExecuteScalar();


            }
            catch (Exception ex)
            {
                MessageBox.Show("查询记录数错误：" + ex.Message.ToString());
            }
            finally
            {
                conn.Close();
            }
            if (obj == null)
            { return 0; }
            else
            { return Convert.ToInt64(obj); }

        }
        public static List<List<long[]>> updateSame1(DataTable dt)
        {
            string connString = @"Data Source=(DESCRIPTION =(ADDRESS = (PROTOCOL = TCP)(HOST = localhost)(PORT = 1521))(CONNECT_DATA =
      (SERVER = DEDICATED)
      (SERVICE_NAME = orcl)
    )
  ); User Id = SYSTEM; Password =123;";



            OracleConnection conn = new OracleConnection(connString);
            List<List<long[]>> ID = new List<List<long[]>>();
            try
            {

                conn.Open();//打开指定的连接                                         

                //  MessageBox.Show(conn.State.ToString());
                OracleCommand cmd = conn.CreateCommand();
                foreach (DataRow dr in dt.Rows)
                {
                    string strSql = "";
                    strSql += string.Format("select {0},{1} from {2} ", "PERSONID", "HOUSEHOLDERID", "PERSON");

                    strSql += string.Format(" where IDCARD=\'{0}\' and REGIONID=\'{1}\' and PERSONNAME=\'{2}\' ", dr["ChengyuanIDCard"], dr["REGIONID"], dr["ChengyuanName"]);
                    //strSql += string.Format(" where IDCARD=\'{0}\' and REGIONID=\'{1}\'  ", dr["ChengyuanIDCard"], dr["REGIONID"]);    

                    cmd.CommandText = strSql;
                    //int result = cmd.ExecuteNonQuery();
                    // if (result > 0)
                    // {

                    List<long[]> b = new List<long[]>();//每个idcard，REGIONID， PERSONNAME对应的所有数据
                    OracleDataReader rd = cmd.ExecuteReader();

                    if (rd.Read())
                    {
                        long a1 = Int64.Parse(rd.GetValue(0).ToString());
                        long a2 = Int64.Parse(rd.GetValue(1).ToString());
                        while (rd.Read())//读取数据，如果返回为false的话，就说明到记录集的尾部了  
                        {


                            long[] a = new long[2] { a1, a2 };
                            b.Add(a);
                            a1 = Int64.Parse(rd.GetValue(0).ToString());
                            a2 = Int64.Parse(rd.GetValue(1).ToString());

                        }
                        b.Add(new long[2] { a1, a2 });
                        rd.Close();
                        ID.Add(b);
                    }
                    // }

                }



                if (ID.Count > 0)
                {

                    DataTable person_dt = GetDataTable("PERSON");
                    DataTable householder_dt = GetDataTable("HOUSEHOLDER");
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        for (int k = 0; k < ID[i].Count; k++)
                        {
                            DataRow person_dr = person_dt.NewRow();
                            person_dr["PERSONID"] = ID[i][k][0];
                            person_dr["HOUSEHOLDERID"] = ID[i][k][1];
                            person_dr["PERSONNAME"] = dt.Rows[i]["ChengyuanName"];
                            person_dr["IDCARD"] = dt.Rows[i]["ChengyuanIDCard"];
                            person_dr["PERSONNATIONAL"] = dt.Rows[i]["National"];
                            person_dr["REGIONID"] = dt.Rows[i]["RegionID"];
                            person_dr["RELATIONSHIP"] = dt.Rows[i]["Relationship"];
                            person_dr["CULTURE"] = dt.Rows[i]["Culture"];
                            person_dr["SCHOOL"] = dt.Rows[i]["School"];
                            person_dt.Rows.Add(person_dr);
                        }
                    }
                    DataRow houseHolder_dr = householder_dt.NewRow();
                    houseHolder_dr["HOUSEHOLDERID"] = ID[0][0][1];
                    houseHolder_dr["PHONE"] = dt.Rows[0]["Phone"];
                    houseHolder_dr["RELOCATE"] = dt.Rows[0]["Relocate"];
                    houseHolder_dr["BANKNAME"] = dt.Rows[0]["BankName"];
                    houseHolder_dr["REGIONID"] = dt.Rows[0]["RegionID"];
                    houseHolder_dr["CREDITCARD"] = dt.Rows[0]["CreditCard"];
                    houseHolder_dr["PROJECTID"] = dt.Rows[0]["ProjectID"];
                    householder_dt.Rows.Add(houseHolder_dr);
                    int j = -1;
                    foreach (DataRow dr in person_dt.Rows)
                    {
                        j++;
                        for (int k = 0; k < ID[j].Count; k++)
                        {
                            string sqlStr1 = "";
                            sqlStr1 += string.Format("UPDATE {0} SET ", "PERSON");
                            sqlStr1 += string.Format(" {0}=\'{1}\' ", person_dt.Columns[1].ColumnName, dr[person_dt.Columns[1]]);
                            for (int i = 2; i < person_dt.Columns.Count; i++)
                            {
                                sqlStr1 += string.Format(" ,{0}=\'{1}\' ", person_dt.Columns[i].ColumnName, dr[person_dt.Columns[i]]);
                            }
                            sqlStr1 += string.Format(" WHERE personid= \'{0}\' ", ID[j][k][0]);
                            cmd.CommandText = sqlStr1;
                            //OracleCommand cmd1 = new OracleCommand(sqlStr1, conn);
                            int result = cmd.ExecuteNonQuery();
                            if (result > 0)
                            {
                                strList.Add(string.Format("数据覆盖了Person表中PersonID:{0}的数据！", ID[j][k][0]));
                                //输出                    
                            }
                        }
                    }
                    j = -1;
                    foreach (DataRow dr in householder_dt.Rows)
                    {
                        j++;
                        string sqlStr2 = "";
                        sqlStr2 += string.Format("UPDATE {0} SET ", "HOUSEHOLDER");
                        sqlStr2 += string.Format(" {0}=\'{1}\' ", householder_dt.Columns[0].ColumnName, dr[householder_dt.Columns[0]]);
                        for (int i = 1; i < householder_dt.Columns.Count; i++)
                        {
                            sqlStr2 += string.Format(" ,{0}=\'{1}\' ", householder_dt.Columns[i].ColumnName, dr[householder_dt.Columns[i]]);
                        }
                        sqlStr2 += string.Format(" where HOUSEHOLDERID= \'{0}\' ", ID[j][0][1]);
                        OracleCommand cmd1 = new OracleCommand(sqlStr2, conn);
                        int result1 = cmd1.ExecuteNonQuery();
                        if (result1 > 0)
                        {
                            strList.Add(string.Format("覆盖了HouseHolder表中HouseHolderID:{0}的数据！", ID[j][0][1]));
                            //输出                    
                        }
                    }



                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("更新数据：" + ex.Message.ToString());
            }
            finally
            {

                conn.Close();

            }


            return ID;
        }       
        public static List<DataSet> yibiaoruku1(DataSet ds)
        {
            List<DataSet> ds_list = new List<DataSet>();
            DataTable statistics_dt = GetDataTable("STATICSINFO");
            DataTable multimedia_dt = GetDataTable("MULTIMEDIAINFO");

            long statisticsID = getDtMaxID("STATICSINFO");
            long multimediaID = getDtMaxID("MULTIMEDIAINFO");
            
            int j = 0; int m = 1; int m1 = 0;
            string name = ds.Tables[0].Rows[0]["NAME"].ToString();
            DataRow houseFirstRow = ds.Tables[0].NewRow();

            DataTable singlehouse = ds.Tables[0].Clone();
            foreach (DataTable dt in ds.Tables)
            {
                j++;

                foreach (DataRow dr in dt.Rows)
                {
                    m1++;
                    if (name == dr["NAME"].ToString())
                    {
                        singlehouse.Rows.Add(dr.ItemArray);
                    }
                    else
                    {
                        houseFirstRow = dr;
                        name = dr["NAME"].ToString();


                        List<List<long[]>> a = updateSame1(singlehouse);
                        if (a.Count > 0)
                        {
                            strList.Add(string.Format("第{0}sheet第{1}到行覆盖了STATICSINFO,MULTIMEDIAINFO表中同年数据,其中:", j, m, m1 - 1));
                            
                         
                            
                        }
                        else
                        {

                            //所有
                            statisticsID++; 
                            for (int i = 0; i < singlehouse.Rows.Count; i++)
                            {
                              
                                      multimediaID++;
                                    DataRow multimedia_dr = multimedia_dt.NewRow();
                                    multimedia_dr["multimediaid"]=multimediaID;
                                    multimedia_dr["statisticsid"]=statisticsID;
                                    multimedia_dr["filename"]=singlehouse.Rows[i]["FILENAME"];
                                    multimedia_dr["azimuth"]=singlehouse.Rows[i]["AZIMUTH"];
                                    multimedia_dr["satellite"]=singlehouse.Rows[i]["SATELLITES"];
                                      multimedia_dr["longitude"]=singlehouse.Rows[i]["LONGITUDE"];
                                      multimedia_dr["latitude "]=singlehouse.Rows[i]["LATITUDE"];
                                    multimedia_dt.Rows.Add(multimedia_dr.ItemArray);

                              

                            }
                             
                            DataRow statistics_dr = statistics_dt .NewRow();
                           statistics_dr["statisticsid"] = statisticsID;
                            statistics_dr["longitude"] = singlehouse.Rows[0]["LONGITUDE"];
                            statistics_dr["latitude"] = singlehouse.Rows[0]["LATITUDE"];
                           statistics_dr["gatherer"] = singlehouse.Rows[0]["GATHERER"];
                           statistics_dr["collectiontime"] = singlehouse.Rows[0]["COLLECTIONTIME"];
                           statistics_dr["mark"] = singlehouse.Rows[0]["MARK"];
                           statistics_dr["houseorder"] = singlehouse.Rows[0]["HOUSEORDER"];
                            statistics_dt.Rows.Add(statistics_dr.ItemArray);

                           


                        }

                        singlehouse.Clear();
                        singlehouse.Rows.Add(houseFirstRow.ItemArray);
                    }

                }

            }
            DataSet statistics_ds = new DataSet(); statistics_ds.Tables.Add(statistics_dt);
            DataSet multimedia_ds = new DataSet(); multimedia_ds.Tables.Add(multimedia_dt);


            ds_list.Add(statistics_ds); ds_list.Add(multimedia_ds); 
            //输出检查结果
            printWord(strList);
            return ds_list;

        }
        #endregion

        #region//贫困人员户主信息一表入库
        public static List<List<long[]>> updateSame(DataTable dt)
        {
            string connString = @"Data Source=(DESCRIPTION =(ADDRESS = (PROTOCOL = TCP)(HOST = localhost)(PORT = 1521))(CONNECT_DATA =
      (SERVER = DEDICATED)
      (SERVICE_NAME = orcl)
    )
  ); User Id = SYSTEM; Password =123;";



            OracleConnection conn = new OracleConnection(connString);
            List<List<long[]>> ID = new List<List<long[]>>();
            try
            {                             

                conn.Open();//打开指定的连接                                         
                  
              //  MessageBox.Show(conn.State.ToString());
                OracleCommand cmd = conn.CreateCommand();
                foreach (DataRow dr in dt.Rows)
                {
                   string strSql = "";
                      strSql += string.Format("select {0},{1} from {2} ", "PERSONID", "HOUSEHOLDERID", "PERSON");
                    
                      strSql += string.Format(" where IDCARD=\'{0}\' and REGIONID=\'{1}\' and PERSONNAME=\'{2}\' ", dr["ChengyuanIDCard"], dr["REGIONID"], dr["ChengyuanName"]);
                      //strSql += string.Format(" where IDCARD=\'{0}\' and REGIONID=\'{1}\'  ", dr["ChengyuanIDCard"], dr["REGIONID"]);    
                    
                        cmd.CommandText = strSql;
                   //int result = cmd.ExecuteNonQuery();
                   // if (result > 0)
                   // {
                 
                        List<long[]> b = new List<long[]>();//每个idcard，REGIONID， PERSONNAME对应的所有数据
                        OracleDataReader rd = cmd.ExecuteReader();

                        if (rd.Read())
                        {
                            long a1 = Int64.Parse(rd.GetValue(0).ToString());
                            long a2 = Int64.Parse(rd.GetValue(1).ToString());
                        while (rd.Read())//读取数据，如果返回为false的话，就说明到记录集的尾部了  
                        {

                            
                            long[] a = new long[2]{ a1, a2 };
                            b.Add(a);
                             a1 = Int64.Parse(rd.GetValue(0).ToString());
                             a2 = Int64.Parse(rd.GetValue(1).ToString());
                           
                       }
                        b.Add(new long[2] { a1, a2 });
                     rd.Close();
                     ID.Add(b);
                    }
                   // }
                                     
               }
                              
               
                
                if (ID.Count>0)
                {

                    DataTable person_dt = GetDataTable("PERSON");
                    DataTable householder_dt = GetDataTable("HOUSEHOLDER");
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        for (int k = 0; k< ID[i].Count; k++) { 
                        DataRow person_dr = person_dt.NewRow();
                        person_dr["PERSONID"] = ID[i][k][0];
                        person_dr["HOUSEHOLDERID"] = ID[i][k][1];
                        person_dr["PERSONNAME"] = dt.Rows[i]["ChengyuanName"];
                        person_dr["IDCARD"] = dt.Rows[i]["ChengyuanIDCard"];
                        person_dr["PERSONNATIONAL"] = dt.Rows[i]["National"];
                        person_dr["REGIONID"] = dt.Rows[i]["RegionID"];
                        person_dr["RELATIONSHIP"] = dt.Rows[i]["Relationship"];
                        person_dr["CULTURE"] = dt.Rows[i]["Culture"];
                        person_dr["SCHOOL"] = dt.Rows[i]["School"];
                        person_dt.Rows.Add(person_dr);
                        }
                    }
                    DataRow houseHolder_dr = householder_dt.NewRow();
                    houseHolder_dr["HOUSEHOLDERID"] = ID[0][0][1];
                    houseHolder_dr["PHONE"] = dt.Rows[0]["Phone"];
                    houseHolder_dr["RELOCATE"] = dt.Rows[0]["Relocate"];
                    houseHolder_dr["BANKNAME"] = dt.Rows[0]["BankName"];
                    houseHolder_dr["REGIONID"] = dt.Rows[0]["RegionID"];
                    houseHolder_dr["CREDITCARD"] = dt.Rows[0]["CreditCard"];
                    houseHolder_dr["PROJECTID"] = dt.Rows[0]["ProjectID"];
                    householder_dt.Rows.Add(houseHolder_dr);
                    int j = -1; 
                    foreach (DataRow dr in person_dt.Rows)
                    {
                        j++;
                        for (int k = 0; k < ID[j].Count; k++)
                        {
                            string sqlStr1 = "";
                            sqlStr1 += string.Format("UPDATE {0} SET ", "PERSON");
                            sqlStr1 += string.Format(" {0}=\'{1}\' ", person_dt.Columns[1].ColumnName, dr[person_dt.Columns[1]]);
                            for (int i = 2; i < person_dt.Columns.Count; i++)
                            {
                                sqlStr1 += string.Format(" ,{0}=\'{1}\' ", person_dt.Columns[i].ColumnName, dr[person_dt.Columns[i]]);
                            }
                            sqlStr1 += string.Format(" WHERE personid= \'{0}\' ", ID[j][k][0]);
                            cmd.CommandText = sqlStr1;
                            //OracleCommand cmd1 = new OracleCommand(sqlStr1, conn);
                            int result = cmd.ExecuteNonQuery();
                            if (result > 0)
                            {
                                strList.Add(string.Format("数据覆盖了Person表中PersonID:{0}的数据！", ID[j][k][0]));
                                //输出                    
                            }
                        }
                    }
                    j=-1;
                    foreach (DataRow dr in householder_dt.Rows)
                    {
                        j++;
                        string sqlStr2 = "";
                        sqlStr2 += string.Format("UPDATE {0} SET ", "HOUSEHOLDER");
                        sqlStr2 += string.Format(" {0}=\'{1}\' ", householder_dt.Columns[0].ColumnName, dr[householder_dt.Columns[0]]);
                        for (int i = 1; i < householder_dt.Columns.Count; i++)
                        {
                            sqlStr2 += string.Format(" ,{0}=\'{1}\' ", householder_dt.Columns[i].ColumnName, dr[householder_dt.Columns[i]]);
                        }
                        sqlStr2 += string.Format(" where HOUSEHOLDERID= \'{0}\' ",ID[j][0][1] );
                        OracleCommand cmd1 = new OracleCommand(sqlStr2, conn);
                        int result1 = cmd1.ExecuteNonQuery();
                        if (result1 > 0)
                        {
                            strList.Add(string.Format("覆盖了HouseHolder表中HouseHolderID:{0}的数据！", ID[j][0][1]));
                            //输出                    
                        }
                   }

                

            }}
            catch (Exception ex)
            {
                MessageBox.Show("更新数据："+ex.Message.ToString());
            }
            finally
            {

                conn.Close();

            }

           
            return ID;
        }       
       
        public static List<DataSet> yibiaoruku(DataSet ds)         
        {
            List<DataSet> ds_list = new List<DataSet>();
            DataTable person_dt = GetDataTable("PERSON");
            DataTable personTense_dt = GetDataTable("PERSONTENSE");
            DataTable householder_dt = GetDataTable("HOUSEHOLDER");
            DataTable householderTense_dt = GetDataTable("HOUSEHOLDERTENSE");
            long houseHolderID = getDtMaxID("HOUSEHOLDER");
            long personID = getDtMaxID("PERSON");
            long PersonTenseID = getDtMaxID("PERSONTENSE");
            long houseHolderTenseID = getDtMaxID("HOUSEHOLDERTENSE");
            int j = 0; int m = 1;int m1 = 0;
            string name = ds.Tables[0].Rows[0]["HuzhuName"].ToString();
            DataRow houseFirstRow = ds.Tables[0].NewRow();

            DataTable singlehouse = ds.Tables[0].Clone();
            foreach (DataTable dt in ds.Tables)
            {
                 j++;    
                
                foreach(DataRow dr in dt.Rows)
                {
                    m1++;
                    if (name == dr["HuzhuName"].ToString())
                    {
                        singlehouse.Rows.Add(dr.ItemArray);
                     }
                    else 
                    {
                        houseFirstRow = dr;
                        name = dr["HuzhuName"].ToString();
                        

                        List<List<long[]>> a = updateSame(singlehouse);
                         if (a.Count>0)
                         {
                             strList.Add(string.Format("第{0}sheet第{1}到行覆盖了Person,HouseHolder表中同年数据,其中:", j,m, m1-1 ));
                             m = m1;
                             //返回覆盖的personID，houseHolderID
                             houseHolderTenseID++;
                             for (int i = 0; i < singlehouse.Rows.Count; i++)
                             {
                                 for (int k = 0; k < a[i].Count; k++)
                                 {
                                     //personID++;
                                     PersonTenseID++;
                                     DataRow personTense_dr = personTense_dt.NewRow();
                                     personTense_dr["PERSONTENSEID"] = PersonTenseID;
                                     personTense_dr["PERSONID"] = a[i][k][0];
                                     personTense_dr["LABORCAPACITY"] = singlehouse.Rows[i]["LaborCapacity"];
                                     personTense_dr["ENGAGESTATE"] = singlehouse.Rows[i]["EngageState"];
                                     personTense_dr["ENGAGETIME"] = singlehouse.Rows[i]["EngageTime"];
                                     personTense_dr["LIVINGSTATE"] = singlehouse.Rows[i]["LivingState"];
                                     personTense_dr["HEALTHSTATE"] = singlehouse.Rows[i]["HealthState"];
                                     personTense_dr["RECORDTIME"] = singlehouse.Rows[i]["RecordTime"];
                                     personTense_dt.Rows.Add(personTense_dr.ItemArray);
                                 }

                             }

                             DataRow houseHolderTense_dr = householderTense_dt.NewRow();
                             houseHolderTense_dr["HOUSEHOLDERTENSEID"] = houseHolderTenseID;
                             houseHolderTense_dr["HOUSEHOLDERID"] = a[0][0][1];
                             houseHolderTense_dr["POVERTY"] = singlehouse.Rows[0]["Poverty"];
                             houseHolderTense_dr["POVERTYLEVEL"] = singlehouse.Rows[0]["Level"];
                             houseHolderTense_dr["REASON"] = singlehouse.Rows[0]["Reason"];
                             houseHolderTense_dr["INCOME"] = singlehouse.Rows[0]["Income"];
                             houseHolderTense_dr["RECORDTIME"] = singlehouse.Rows[0]["RecordTime"];
                             householderTense_dt.Rows.Add(houseHolderTense_dr.ItemArray);
                             
                         //tense
                         }
                         else 
                         {                         

                         //所有
                             houseHolderID++; houseHolderTenseID++;
                             for (int i = 0; i < singlehouse.Rows.Count; i++)
                             {
                                 personID++; PersonTenseID++;
                                 DataRow person_dr = person_dt.NewRow();
                                 person_dr["PERSONID"] = personID;
                                 person_dr["HOUSEHOLDERID"] = houseHolderID;
                                 person_dr["PERSONNAME"] = singlehouse.Rows[i]["ChengyuanName"];
                                 person_dr["IDCARD"] = singlehouse.Rows[i]["ChengyuanIDCard"];
                                 person_dr["PERSONNATIONAL"] = singlehouse.Rows[i]["National"];
                                 person_dr["REGIONID"] = singlehouse.Rows[i]["RegionID"];
                                 person_dr["RELATIONSHIP"] = singlehouse.Rows[i]["Relationship"];
                                 person_dr["CULTURE"] = singlehouse.Rows[i]["Culture"];
                                 person_dr["SCHOOL"] = singlehouse.Rows[i]["School"];
                                 person_dt.Rows.Add(person_dr.ItemArray);



                                 DataRow personTense_dr = personTense_dt.NewRow();
                                 personTense_dr["PERSONTENSEID"] = PersonTenseID;
                                 personTense_dr["PERSONID"] = personID;
                                 personTense_dr["LABORCAPACITY"] = singlehouse.Rows[i]["LaborCapacity"];
                                 personTense_dr["ENGAGESTATE"] = singlehouse.Rows[i]["EngageState"];
                                 personTense_dr["ENGAGETIME"] = singlehouse.Rows[i]["EngageTime"];
                                 personTense_dr["LIVINGSTATE"] = singlehouse.Rows[i]["LivingState"];
                                 personTense_dr["HEALTHSTATE"] = singlehouse.Rows[i]["HealthState"];
                                 personTense_dr["RECORDTIME"] = singlehouse.Rows[i]["RecordTime"];
                                 personTense_dt.Rows.Add(personTense_dr.ItemArray);
                                

                                
                                
                             }
                             DataRow houseHolder_dr = householder_dt.NewRow();
                             houseHolder_dr["HOUSEHOLDERID"] = houseHolderID;
                             houseHolder_dr["REGIONID"] = singlehouse.Rows[0]["RegionID"];
                             houseHolder_dr["PHONE"] = singlehouse.Rows[0]["Phone"];
                             houseHolder_dr["RELOCATE"] = singlehouse.Rows[0]["Relocate"];
                             houseHolder_dr["BANKNAME"] = singlehouse.Rows[0]["BankName"];
                             houseHolder_dr["CREDITCARD"] = singlehouse.Rows[0]["CreditCard"];
                             houseHolder_dr["PROJECTID"] = singlehouse.Rows[0]["ProjectID"];
                             householder_dt.Rows.Add(houseHolder_dr.ItemArray);

                             

                             DataRow houseHolderTense_dr = householderTense_dt.NewRow();
                             houseHolderTense_dr["HOUSEHOLDERTENSEID"] = houseHolderTenseID;
                             houseHolderTense_dr["HOUSEHOLDERID"] = houseHolderID;
                             houseHolderTense_dr["POVERTY"] = singlehouse.Rows[0]["Poverty"];
                             houseHolderTense_dr["POVERTYLEVEL"] = singlehouse.Rows[0]["Level"];
                             houseHolderTense_dr["REASON"] = singlehouse.Rows[0]["Reason"];
                             houseHolderTense_dr["INCOME"] = singlehouse.Rows[0]["Income"];
                             houseHolderTense_dr["RECORDTIME"] = singlehouse.Rows[0]["RecordTime"];
                             householderTense_dt.Rows.Add(houseHolderTense_dr.ItemArray);
                             
                         }

                         singlehouse.Clear();
                         singlehouse.Rows.Add(houseFirstRow.ItemArray);
                    } 

                }

            }
            DataSet person_ds = new DataSet(); person_ds.Tables.Add(person_dt);
            DataSet personTense_ds = new DataSet(); personTense_ds.Tables.Add(personTense_dt);
            DataSet householder_ds= new DataSet(); householder_ds.Tables.Add(householder_dt);
            DataSet householderTense_ds = new DataSet(); householderTense_ds.Tables.Add(householderTense_dt); 
           
            ds_list.Add(person_ds); ds_list.Add(personTense_ds); ds_list.Add(householder_ds); ds_list.Add(householderTense_ds);
            //输出检查结果
            printWord(strList);
            return ds_list;
           
        }
        #endregion


        public static int BulkToDB(string targetTable, DataSet ds)
        {
            int iResult = 0;
            foreach (DataTable dt in ds.Tables)
            {


                if (string.IsNullOrEmpty(targetTable))
                {
                    throw new ArgumentException("必须指定批量插入的表名称", "tableName");
                }


                string connOrcleString = @"Data Source=(DESCRIPTION =(ADDRESS = (PROTOCOL = TCP)(HOST = localhost)(PORT = 1521))(CONNECT_DATA =
      (SERVER = DEDICATED)
      (SERVICE_NAME = orcl)
    )
  ); User Id = SYSTEM; Password =123;";
                OracleConnection conn = new OracleConnection(connOrcleString);
                OracleBulkCopy bulkCopy = new OracleBulkCopy(connOrcleString, OracleBulkCopyOptions.UseInternalTransaction);   //用其它源的数据有效批量加载Oracle表中
                //conn.BeginTransaction();
                //OracleBulkCopy bulkCopy = new OracleBulkCopy(connOrcleString, OracleBulkCopyOptions.Default);


                bulkCopy.BatchSize = dt.Rows.Count;
                bulkCopy.BulkCopyTimeout = 260;
                bulkCopy.DestinationTableName = targetTable;    //服务器上目标表的名称
                bulkCopy.BatchSize = dt.Rows.Count;   //每一批次中的行数


                foreach (DataColumn dc in dt.Columns)
                {
                    bulkCopy.ColumnMappings.Add(dc.ColumnName, dc.ColumnName);
                }
                try
                {
                    conn.Open();
                    if (dt != null && dt.Rows.Count != 0)

                        bulkCopy.WriteToServer(dt);   //将提供的数据源中的所有行复制到目标表中
                }
                catch (Exception ex)
                {
                    throw ex;
                }
                finally
                {
                    conn.Close();
                    if (bulkCopy != null)
                        bulkCopy.Close();

                    iResult = 1;
                }
            }

            return iResult;
        }


        #region //导出数据库
        public static DataSet DBToDataTable(string tableName)
        {

            DataSet ds = new DataSet();
            string connString = @"Data Source=(DESCRIPTION =(ADDRESS = (PROTOCOL = TCP)(HOST = localhost)(PORT = 1521))(CONNECT_DATA =
      (SERVER = DEDICATED)
      (SERVICE_NAME = orcl)
    )
  ); User Id = SYSTEM; Password =123;";

            OracleConnection conn = new OracleConnection(connString);
            DataTable dt = new DataTable();

            string strSql = "select * from " + tableName;
            try
            {
                conn.Open();
                OracleCommand cmd = conn.CreateCommand();
                cmd.CommandText = strSql;
                OracleDataReader rd = cmd.ExecuteReader();

                for (int i = 0; i < rd.FieldCount; i++) { dt.Columns.Add(rd.GetName(i), rd.GetFieldType(i)); }

                while (rd.Read())//读取数据，如果返回为false的话，就说明到记录集的尾部了  
                {
                    DataRow r = dt.NewRow();


                    //  Console.WriteLine(" r[dt.Columns[0]]:{0}", r[dt.Columns[0]]);
                    for (int i = 0; i < rd.FieldCount; i++)
                    {
                        if (rd.GetFieldType(i) == typeof(string))
                        {
                            if (!rd.IsDBNull(i))
                            {
                                r[dt.Columns[i]] = rd.GetString(i);
                            }
                            else { r[dt.Columns[i]] = DBNull.Value; }
                        }
                        else if (rd.GetFieldType(i) == typeof(DateTime))
                        {
                            if (!rd.IsDBNull(i))
                            {
                                r[dt.Columns[i]] = rd.GetGuid(i);
                            }
                            else { r[dt.Columns[i]] = DBNull.Value; }
                        }
                        else //if (rd.GetFieldType(i) == typeof(Int32) || rd.GetFieldType(i) == typeof(Int64) || rd.GetFieldType(i) == typeof(NumberFormat))
                        {

                            if (!rd.IsDBNull(i))
                            {
                                r[dt.Columns[i]] = rd.GetValue(i);
                            }
                            else { r[dt.Columns[i]] = DBNull.Value; }
                        }

                    //    if (!rd.IsDBNull(i))
                    //    {


                    //        r[dt.Columns[i]] = rd.GetValue(i);


                    //    }
                    //    else
                    //    {
                    //        r[dt.Columns[i]] = DBNull.Value;
                    //        //r[dt.Columns[i]] = string.Empty;
                    //    }
                    }
                    dt.Rows.Add(r);
                }

                // MessageBox.Show("读取表结束");  //弹出显示
                rd.Close();//关闭reader.这是一定要写的  
                ds.Tables.Add(dt);

            }
            catch (Exception ex)
            {
                MessageBox.Show("读取表记录错误：" + ex.Message.ToString());
                return null;
            }
            finally
            {
                conn.Close();

            }
            return ds;
        }
        
          public static bool DataTableToExcel(DataSet ds, string str1)
        {
            bool result = false;
            int k = -1;
            foreach(DataTable dt in ds.Tables)
            {
                k++;
                IWorkbook workbook = null;
                FileStream fs = null;
                IRow row = null;
                ISheet sheet = null;
                NPOI.SS.UserModel.ICell cell = null;
                try
                {
                    if (dt != null && dt.Rows.Count > 0)
                    {
                        workbook = new HSSFWorkbook();
                        sheet = workbook.CreateSheet(string.Format("Sheet{0}",k));//创建一个名称为Sheet0的表  
                        int rowCount = dt.Rows.Count;//行数  
                        int columnCount = dt.Columns.Count;//列数  

                        //设置列头  
                        row = sheet.CreateRow(0);//excel第一行设为列头  
                        for (int c = 0; c < columnCount; c++)
                        {
                            cell = row.CreateCell(c);
                            cell.SetCellValue(dt.Columns[c].ColumnName);
                        }

                        //设置每行每列的单元格,  
                        for (int i = 0; i < rowCount; i++)
                        {
                            row = sheet.CreateRow(i + 1);
                            for (int j = 0; j < columnCount; j++)
                            {
                                cell = row.CreateCell(j);//excel第二行开始写入数据  
                                cell.SetCellValue(dt.Rows[i][j].ToString());
                            }
                        }
                        using (fs = File.OpenWrite(str1))
                        {
                            workbook.Write(fs);//向打开的这个xls文件中写入数据  
                            result = true;
                        }
                    }
                   
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                    if (fs != null)
                    {
                        fs.Close();
                    }

                    return false;
                }
               
            }
            return result;
        }

         #endregion


          #region//条件查询读取combox列表

        public static List<string> GetShi()
        {
            List<string> shiStr = new List<string>();

            string connStr = @"Data Source=(DESCRIPTION =(ADDRESS = (PROTOCOL = TCP)(HOST = localhost)(PORT = 1521))(CONNECT_DATA =
      (SERVER = DEDICATED)(SERVICE_NAME = orcl))); User Id = SYSTEM; Password =123;";
            OracleConnection conn = new OracleConnection(connStr);

            try
            {


                conn.Open();//打开指定的连接
                MessageBox.Show(conn.State.ToString());


                //改tableName
                string sqlStr = "select * from " + "region";


                OracleCommand cmd = new OracleCommand(sqlStr, conn);
                OracleDataReader rd = cmd.ExecuteReader();
               // if (rd.Read())
                ///{
                //    MessageBox.Show("读取表成功");  //弹出显示
               // }

                while (rd.Read())//读取数据，如果返回为false的话，就说明到记录集的尾部了  
                {
               string str= rd.GetString(1);
               if (shiStr.Contains(str))
               {
                   continue;
               }else
               {
                   shiStr.Add(str);
                    }

                }

                // MessageBox.Show("读取表结束");  //弹出显示
                rd.Close();//关闭reader.这是一定要写的
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString());
            }
            finally
            {
                conn.Close();
            }







            return shiStr;
        }
        public static List<string> GetXian(string shiName)
        {
            List<string> shiStr = new List<string>();

            string connStr = @"Data Source=(DESCRIPTION =(ADDRESS = (PROTOCOL = TCP)(HOST = localhost)(PORT = 1521))(CONNECT_DATA =
      (SERVER = DEDICATED)(SERVICE_NAME = orcl))); User Id = SYSTEM; Password =123;";
            OracleConnection conn = new OracleConnection(connStr);

            try
            {


                conn.Open();//打开指定的连接
                MessageBox.Show(conn.State.ToString());

                string sqlStr = " select * from " + " REGION T";
                sqlStr += string.Format(" where T.CITY LIKE \'{0}\'", shiName);
               
                OracleCommand cmd = new OracleCommand(sqlStr, conn);
                OracleDataReader rd = cmd.ExecuteReader();
               // if (rd.Read())
              //  {
             //       MessageBox.Show("读取表成功");  //弹出显示
            //    }

                while (rd.Read())//读取数据，如果返回为false的话，就说明到记录集的尾部了  
                {
                    string str = rd.GetString(2);
                    if (shiStr.Contains(str))
                    {
                        continue;
                    }
                    else
                    {
                        shiStr.Add(str);
                    }

                    


                }

                // MessageBox.Show("读取表结束");  //弹出显示
                rd.Close();//关闭reader.这是一定要写的
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString());
            }
            finally
            {
                conn.Close();
            }







            return shiStr;
        }
        public static List<string> GetZhen(string shiName,string xianName)
        {
            List<string> shiStr = new List<string>();

            string connStr = @"Data Source=(DESCRIPTION =(ADDRESS = (PROTOCOL = TCP)(HOST = localhost)(PORT = 1521))(CONNECT_DATA =
      (SERVER = DEDICATED)(SERVICE_NAME = orcl))); User Id = SYSTEM; Password =123;";
            OracleConnection conn = new OracleConnection(connStr);

            try
            {


                conn.Open();//打开指定的连接
                MessageBox.Show(conn.State.ToString());

                string sqlStr = "select * from " + "REGION T ";
                sqlStr += string.Format(" WHERE T.CITY LIKE \'{0}\'", shiName);
                sqlStr += string.Format(" AND   T.COUNTY LIKE \'{0}\'", xianName);
                OracleCommand cmd = new OracleCommand(sqlStr, conn);
                OracleDataReader rd = cmd.ExecuteReader();
              //  if (rd.Read())
               // {
              //      MessageBox.Show("读取表成功");  //弹出显示
               // }

                while (rd.Read())//读取数据，如果返回为false的话，就说明到记录集的尾部了  
                {
                    string str = rd.GetString(3);
                    if (shiStr.Contains(str))
                    {
                        continue;
                    }
                    else
                    {
                        shiStr.Add(str);
                    }

                    

                }

                // MessageBox.Show("读取表结束");  //弹出显示
                rd.Close();//关闭reader.这是一定要写的
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString());
            }
            finally
            {
                conn.Close();
            }







            return shiStr;
        }
        public static List<string> GetCun(string shiName,string xianName,string zhenName)
        {
            List<string> shiStr = new List<string>();

            string connStr = @"Data Source=(DESCRIPTION =(ADDRESS = (PROTOCOL = TCP)(HOST = localhost)(PORT = 1521))(CONNECT_DATA =
      (SERVER = DEDICATED)(SERVICE_NAME = orcl))); User Id = SYSTEM; Password =123;";
            OracleConnection conn = new OracleConnection(connStr);

            try
            {


                conn.Open();//打开指定的连接
                MessageBox.Show(conn.State.ToString());


                //改tableName
                string sqlStr = "select * from " + "region T";

                sqlStr += string.Format(" where T.city LIKE \'{0}\'",  shiName);
                sqlStr += string.Format(" and T.county LIKE \'{0}\'",  xianName);
                sqlStr += string.Format(" and T.town  LIKE   \'{0}\'",  zhenName);
                OracleCommand cmd = new OracleCommand(sqlStr, conn);
                OracleDataReader rd = cmd.ExecuteReader();
             //   if (rd.Read())
              //  {
             //       MessageBox.Show("读取表成功");  //弹出显示
             //   }

                while (rd.Read())//读取数据，如果返回为false的话，就说明到记录集的尾部了  
                {
                    string str = rd.GetString(4);
                    if (shiStr.Contains(str))
                    {
                        continue;
                    }
                    else
                    {
                        shiStr.Add(str);
                    }
                    


                }

                // MessageBox.Show("读取表结束");  //弹出显示
                rd.Close();//关闭reader.这是一定要写的
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString());
            }
            finally
            {
                conn.Close();
            }







            return shiStr;
        }
        //条件查询获取地址id
        public static long GetRegionID(string shiName, string xianName, string zhenName, string cunName)
        {
            long regionid=0;

            string connStr = @"Data Source=(DESCRIPTION =(ADDRESS = (PROTOCOL = TCP)(HOST = localhost)(PORT = 1521))(CONNECT_DATA =
      (SERVER = DEDICATED)(SERVICE_NAME = orcl))); User Id = SYSTEM; Password =123;";
            OracleConnection conn = new OracleConnection(connStr);
           
            try
            {


                conn.Open();//打开指定的连接
                MessageBox.Show(conn.State.ToString());


                //改tableName
                string sqlStr = "select * from " + "region T";

                sqlStr += string.Format(" where T.city LIKE \'{0}\'", shiName);
                sqlStr += string.Format(" and T.county LIKE \'{0}\'", xianName);
                sqlStr += string.Format(" and T.town  LIKE   \'{0}\'", zhenName);
                sqlStr += string.Format(" and T.VILLIAGE  LIKE   \'{0}\'", cunName);
                OracleCommand cmd = new OracleCommand(sqlStr, conn);
                OracleDataReader rd = cmd.ExecuteReader();
               
               
                while (rd.Read())//读取数据，如果返回为false的话，就说明到记录集的尾部了  
                {
                    regionid =Int64.Parse( rd.GetValue(0).ToString());
                    Console.WriteLine(" regionid:{0}", regionid);
                 
                }

                // MessageBox.Show("读取表结束");  //弹出显示
                rd.Close();//关闭reader.这是一定要写的
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString());
            }
            finally
            {
                conn.Close();
            }

             return regionid;
        }
        #endregion
    }
}
