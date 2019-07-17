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
            p.ImportExcelEngineer(@"C:\Users\Administrator\Documents\WeChat Files\lzy945945\FileStorage\File\2019-07\用户信息.xlsx");
        }

        //连接数据库
        static string connstr = "server=.database=EngineerDB;uid=root;pwd=1433547973@qq.com";
        SqlConnection conn = new SqlConnection(connstr);
        public void GetConnection()
        {
            try
            {
                conn.Open();
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }

        //导入Excel后的操作
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
                    //将Excel中信息放入engineerSheet二维数组中
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
                    int velocityCol = 0;
                    int EmailCol = 0;
                    int usernameCol = 0;
                    int colNumber = 0;
                    foreach (string colname in engineerSheet)
                    {
                        colNumber++;
                        if (colname == "V")
                        {
                            velocityCol = colNumber;//找到速度列的位置
                        }
                        else if(colname== "Email")
                        {
                            EmailCol = colNumber;//找到邮箱列的位置
                        }
                        else if(colname== "Username")
                        {
                            usernameCol = colNumber;//找到用户名列的位置
                        }
                        if (velocityCol != 0 && EmailCol != 0 && usernameCol != 0) break;

                    }
                    string[,] abilitySheet = new string[rows, 3];
                    for(int i = 1; i < rows; i++)
                    {
                        abilitySheet[i - 1, 0] = engineerSheet[i, usernameCol - 1];
                        abilitySheet[i - 1, 1] = engineerSheet[i, EmailCol - 1];
                    }
                    for(int i = 1; i < rows; i++)
                    {
                        for (int j = velocityCol; j < cols; j++)
                        {
                            if (engineerSheet[i, j] == "1")
                            {
                                if (abilitySheet[i-1,2] == null)
                                {
                                    abilitySheet[i - 1, 2] += engineerSheet[0, j];
                                }
                                else
                                {
                                    abilitySheet[i - 1, 2] += "," + engineerSheet[0, j];
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
                    
                    
                    
                }
            }catch(Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }
    }
}
