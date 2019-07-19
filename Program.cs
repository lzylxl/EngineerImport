using OfficeOpenXml;
using System;
using System.IO;
using System.Data.SqlClient;
using System.Data;
using System.Collections.Generic;

namespace EngineerImport
{
    class Program
    {
        static void Main(string[] args)
        {
            string filepath = @"C:\Users\Administrator\Documents\WeChat Files\lzy945945\FileStorage\File\2019-07\用户信息.xlsx";
            Program p = new Program();
            p.ImportExcelEngineer(filepath);
            //ImportExcelEngineer_mysql imp = new ImportExcelEngineer_mysql();
            //imp.ImportExcelEngineer(filepath);
        }

        static string connstr = "server=.;database=EngineerDB;uid=sa;pwd=1433547973@qq.com";
        SqlConnection conn = new SqlConnection(connstr);
        
        /// <summary>
        /// 导入Excel到数据库将数据分别存入usertable和abilitytable中
        /// </summary>
        /// <param name="filePath"></param>
        public void ImportExcelEngineer(string filePath)
        {
            try
            {
                FileInfo newFile = new FileInfo(filePath);
                using (ExcelPackage package = new ExcelPackage(newFile))
                {
                    //将Excel中的数据完成的放入engineerSheet中
                    ExcelWorksheet sheet = package.Workbook.Worksheets[0];
                    int rows = sheet.Dimension.End.Row;
                    int cols = sheet.Dimension.End.Column;
                    string[,] engineerSheet = new string[rows, cols];
                    for (int i = 1; i <= rows; i++)
                    {
                        for (int j = 1; j <= cols; j++)
                        {
                            if (sheet.Cells[i, j].Value != null)
                            {
                                engineerSheet[i - 1, j - 1] = sheet.Cells[i, j].Text;
                            }
                            else
                            {
                                engineerSheet[i - 1, j - 1] = "";
                            }
                        }
                    }
                    //找到engineerSheet中列名对应的列号
                    int colNumber = 0;
                    int velocityCol = 0;
                    int EmailCol = 0;
                    int usernameCol = 0;
                    int nameCol = 0;
                    int roleCol = 0;
                    int organizationCol = 0;
                    int passwordCol = 0;
                    foreach (string colname in engineerSheet)
                    {
                        colNumber++;
                        if (colname == "V")
                        {
                            velocityCol = colNumber;
                        }
                        else if(colname== "Email")
                        {
                            EmailCol = colNumber;
                        }
                        else if(colname== "Username")
                        {
                            usernameCol = colNumber;
                        }
                        else if (colname == "Role")
                        {
                            roleCol = colNumber;
                        }
                        else if (colname == "Organization")
                        {
                            organizationCol = colNumber;
                        }
                        else if (colname == "Name")
                        {
                            nameCol = colNumber;
                        }
                        else if (colname == "PassWord")
                        {
                            passwordCol = colNumber;
                        }
                        if (velocityCol != 0 && EmailCol != 0 && usernameCol != 0 && nameCol != 0 && roleCol != 0 && organizationCol != 0 && passwordCol != 0) break;
                    }

                    //将engineerSheet中的数据分割成用户和能力两个表
                    //能力表中有用户表的部分数据，能力表中有V速度和ability能力是用户表中没有的
                    //用户表中的数据是engineerSheet中V列之前的所有数据
                    //能力表中目前只有UserName、Email、V、ability四列数据
                    string[,] userSheet = new string[rows - 1, colNumber - 1];
                    string[,] abilitySheet = new string[rows, 4];

                    //获得数据库中usertable表中各列名的列号
                    conn.Open();
                    string condition = "select* from usertable";
                    SqlDataAdapter data = new SqlDataAdapter(condition, conn);
                    DataSet ds = new DataSet();
                    data.Fill(ds);
                    int colnumber = 0;
                    Dictionary<string, int> userTableCol = new Dictionary<string, int>();
                    foreach (DataColumn column in ds.Tables[0].Columns)
                    {
                        userTableCol.Add(column.ColumnName, colnumber++);
                    }
                    //将engineerSheet中V列之前的所有数据放入userSheet中
                    //此时userSheet中的列的顺序与数据库中usertable表中列的顺序完全一致
                    for (int i = 1; i < rows; i++)
                    {
                        userSheet[i - 1, userTableCol["Username"]] = engineerSheet[i, usernameCol - 1];
                        userSheet[i - 1, userTableCol["Email"]] = engineerSheet[i, EmailCol - 1];
                        userSheet[i - 1, userTableCol["Name"]] = engineerSheet[i, nameCol - 1];
                        userSheet[i - 1, userTableCol["Role"]] = engineerSheet[i, roleCol - 1];
                        userSheet[i - 1, userTableCol["Organization"]] = engineerSheet[i, organizationCol - 1];
                        userSheet[i - 1, userTableCol["PassWord"]] = engineerSheet[i, passwordCol - 1];
                    }

                    //获得数据库中abilitytable表中各列名的列号
                    condition = "select* from abilitytable";
                    data = new SqlDataAdapter(condition, conn);
                    ds = new DataSet();
                    data.Fill(ds);
                    colnumber = 0;
                    Dictionary<string, int> abilityTableCol = new Dictionary<string, int>();
                    foreach (DataColumn column in ds.Tables[0].Columns)
                    {

                        abilityTableCol.Add(column.ColumnName, colnumber++);
                    }
                    //此时abilitySheet中的列的顺序与数据库中abilitytable表中列的顺序完全一致
                    for (int i = 1; i < rows; i++)
                    {
                        abilitySheet[i - 1, abilityTableCol["Username"]] = engineerSheet[i, usernameCol - 1];
                        abilitySheet[i - 1, abilityTableCol["Email"]] = engineerSheet[i, EmailCol - 1];
                        abilitySheet[i - 1, abilityTableCol["V"]] = engineerSheet[i, velocityCol - 1];
                    }
                    //abilitySheet中的ability列需要特殊处理将多列数据合并成一列
                    for (int i = 1; i < rows; i++)
                    {
                        for (int j = velocityCol; j < cols; j++)
                        {
                            if (engineerSheet[i, j] == "1")
                            {
                                if (abilitySheet[i - 1, abilityTableCol["ability"]] == null)
                                {
                                    abilitySheet[i - 1, abilityTableCol["ability"]] += engineerSheet[0, j];
                                }
                                else
                                {
                                    abilitySheet[i - 1, abilityTableCol["ability"]] += "," + engineerSheet[0, j];
                                }
                            }
                        }
                    }
                    //for (int i = 0; i < rows - 1; i++) 
                    //{
                    //    for (int j = 0; j < abilitySheet.GetLength(1); j++)
                    //    {
                    //        Console.Write(abilitySheet[i, j] + "     ");
                    //    }
                    //    Console.WriteLine();
                    //}
                    //for (int i = 0; i < rows - 1; i++) 
                    //{
                    //    for (int j = 0; j < userSheet.GetLength(1); j++)
                    //    {
                    //        Console.Write(userSheet[i, j] + "     ");
                    //    }
                    //    Console.WriteLine();
                    //}


                    //将userSheet和abilitySheet中的数据存入数据库
                    SqlCommand cmd = null;
                    string sql = null;
                    for (int i = 0; i < userSheet.GetLength(0); i++)
                    {
                        //判断用户名是否已经存在，不存在则存入数据，存在则跳过
                        sql = "select count(*) from usertable where userName='" + userSheet[i, userTableCol["Username"]] + "'";
                        cmd = new SqlCommand(sql, conn);
                        bool usernameExist = (int)cmd.ExecuteScalar() > 0;
                        if (usernameExist)
                        {
                            continue;
                        }
                        else
                        {
                            sql = "";
                            for (int j = 1; j < userSheet.GetLength(1) - 1; j++) 
                            {
                                sql += "'" + userSheet[i, j] + "',";
                            }
                            sql = "insert into usertable values('" + userSheet[i, 0] + "'," + sql + "'" + userSheet[i, userSheet.GetLength(1) - 1] + "')";
                            cmd=new SqlCommand(sql, conn);
                            cmd.ExecuteNonQuery();
                        }
                    }
                    //将userSheet和abilitySheet中的数据存入数据库
                    for (int i = 0; i < abilitySheet.GetLength(0); i++)
                    {
                        //判断用户名是否已经存在，不存在则存入数据，存在则跳过
                        sql = "select count(*) from abilitytable where userName='" + abilitySheet[i, abilityTableCol["Username"]] + "'";
                        cmd = new SqlCommand(sql, conn);
                        bool usernameExist = (int)cmd.ExecuteScalar() > 0;
                        if (usernameExist)
                        {
                            continue;
                        }
                        else
                        {
                            sql = "insert into abilitytable values('" + abilitySheet[i, 0] + "','" + abilitySheet[i, 1] + "','" + abilitySheet[i, 2] + "','" + abilitySheet[i, 3] + "')";
                            cmd = new SqlCommand(sql, conn);
                            cmd.ExecuteNonQuery();
                        }
                    }
                    conn.Close();
                }
            }
            catch(Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }
    }
}
