using OfficeOpenXml;
using System;
using System.IO;
using System.Data.SqlClient;

namespace EngineerImport
{
    class Program
    {
        static void Main(string[] args)
        {
            Program p = new Program();
            //p.GetConnection();
            p.ImportExcelEngineer(@"C:\Users\Administrator\Documents\WeChat Files\lzy945945\FileStorage\File\2019-07\用户信息.xlsx");
        }

        static string connstr = "server=.;database=EngineerDB;uid=sa;pwd=1433547973@qq.com";
        SqlConnection conn = new SqlConnection(connstr);
        
        /// <summary>
        /// 导入Excel到数据库
        /// </summary>
        /// <param name="filePath"></param>
        public void ImportExcelEngineer(string filePath)
        {
            try
            {
                FileInfo newFile = new FileInfo(filePath);
                using (ExcelPackage package = new ExcelPackage(newFile))
                {
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
                    int colNumber = 0;
                    //找到速度列的位置
                    int velocityCol = 0;
                    //找到邮箱列的位置
                    int EmailCol = 0;
                    //找到用户名列的位置
                    int usernameCol = 0;
                    //找到姓名列的位置
                    int nameCol = 0;
                    //找到角色列的位置
                    int roleCol = 0;
                    //找到部门列的位置
                    int organizationCol = 0;
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
                        if (velocityCol != 0 && EmailCol != 0 && usernameCol != 0 && nameCol != 0 && roleCol != 0 && organizationCol != 0) break;
                    }
                    string[,] userSheet = new string[rows - 1, 5];
                    string[,] abilitySheet = new string[rows - 1, 4];
                    for (int i = 1; i < rows; i++)
                    {
                        userSheet[i - 1, 0] = engineerSheet[i, usernameCol - 1];
                        userSheet[i - 1, 1] = engineerSheet[i, EmailCol - 1];
                        userSheet[i - 1, 2] = engineerSheet[i, nameCol - 1];
                        userSheet[i - 1, 3] = engineerSheet[i, roleCol - 1];
                        userSheet[i - 1, 4] = engineerSheet[i, organizationCol - 1];
                    }
                    for (int i = 1; i < rows; i++)
                    {
                        abilitySheet[i - 1, 0] = engineerSheet[i, usernameCol - 1];
                        abilitySheet[i - 1, 1] = engineerSheet[i, EmailCol - 1];
                        abilitySheet[i - 1, 2] = engineerSheet[i, velocityCol - 1];
                    }
                    for (int i = 1; i < rows; i++)
                    {
                        for (int j = velocityCol; j < cols; j++)
                        {
                            if (engineerSheet[i, j] == "1")
                            {
                                if (abilitySheet[i-1,3] == null)
                                {
                                    abilitySheet[i - 1, 3] += engineerSheet[0, j];
                                }
                                else
                                {
                                    abilitySheet[i - 1, 3] += "," + engineerSheet[0, j];
                                }
                            }
                        }
                    }
                    //for (int i = 0; i < rows; i++)
                    //{
                    //    for (int j = 0; j < abilitySheet.GetLength(1); j++)
                    //    {
                    //        Console.Write(abilitySheet[i, j] + "     ");
                    //    }
                    //    Console.WriteLine();
                    //}
                    //for (int i = 0; i < rows; i++)
                    //{
                    //    for (int j = 0; j < userSheet.GetLength(1); j++)
                    //    {
                    //        Console.Write(userSheet[i, j] + "     ");
                    //    }
                    //    Console.WriteLine();
                    //}

                    SqlCommand cmd;
                    string sql;
                    conn.Open();
                    for (int i = 0; i < userSheet.GetLength(0); i++)
                    {
                        sql = "select count(*) from usertable where userName='" + userSheet[i, 0] + "'";
                        cmd = new SqlCommand(sql, conn);
                        bool usernameExist = (int)cmd.ExecuteScalar() > 0;
                        if (usernameExist)
                        {
                            continue;
                        }
                        else
                        {
                            sql = "insert into usertable values('" + userSheet[i, 0] + "','" + userSheet[i, 1] + "','" + userSheet[i, 2] + "','" + userSheet[i, 3] + "','" + userSheet[i, 4] + "')";
                            cmd=new SqlCommand(sql, conn);
                            cmd.ExecuteNonQuery();
                        }
                    }
                    for (int i = 0; i < abilitySheet.GetLength(0); i++)
                    {
                        sql = "select count(*) from abilitytable where userName='" + abilitySheet[i, 0] + "'";
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
