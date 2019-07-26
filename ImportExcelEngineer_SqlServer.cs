using OfficeOpenXml;
using System;
using System.IO;
using System.Data.SqlClient;
using System.Data;

namespace EngineerImport
{
    class ImportExcelEngineer_SqlServer
    {
        static string connstr = "server=.;database=EngineerDB;uid=sa;pwd=1433547973@qq.com";
        SqlConnection conn = new SqlConnection(connstr);

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

                    //将Excel表中所有数据包括列名放入engineerSheet中
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

                    //创建一个二维数组userSheet存放用户数据
                    //使userSheet的列名和数据库中usertable对应
                    conn.Open();
                    string condition = "select* from usertable";
                    SqlDataAdapter data = new SqlDataAdapter(condition, conn);
                    DataSet userDS = new DataSet();
                    data.Fill(userDS);
                    int usersheetRows = rows;
                    int usersheetCols = userDS.Tables[0].Columns.Count;
                    string[,] userSheet = new string[usersheetRows, usersheetCols];
                    for (int i = 0; i < usersheetCols; i++)
                    {
                        userSheet[0, i] = userDS.Tables[0].Columns[i].ColumnName;
                    }

                    //创建一个二维数组abilitySheet存放用户的能力数据
                    //使abilitySheet的列名和数据库中abilitytable对应
                    condition = "select* from abilitytable";
                    data = new SqlDataAdapter(condition, conn);
                    DataSet abilityDS = new DataSet();
                    data.Fill(abilityDS);
                    int abilitysheetRows = rows;
                    int abilitysheetCols = abilityDS.Tables[0].Columns.Count;
                    string[,] abilitySheet = new string[abilitysheetRows, abilitysheetCols];
                    for (int i = 0; i < abilitysheetCols; i++)
                    {
                        abilitySheet[0, i] = abilityDS.Tables[0].Columns[i].ColumnName;
                    }

                    //将engineerSheet中的数据分别放入userSheet和abilitySheet中
                    for (int i = 1; i < usersheetRows; i++)
                    {
                        for (int j = 0; j < usersheetCols; j++)
                        {
                            for (int colNum = 0; colNum < engineerSheet.GetLength(1); colNum++)
                            {
                                if (userSheet[0, j] == engineerSheet[0, colNum])
                                {
                                    userSheet[i, j] = engineerSheet[i, colNum];
                                }
                            }
                        }
                    }
                    for (int i = 1; i < abilitysheetRows; i++)
                    {
                        for (int j = 0; j < abilitysheetCols; j++)
                        {
                            for (int colNum = 0; colNum < engineerSheet.GetLength(1); colNum++)
                            {
                                if (abilitySheet[0, j] == engineerSheet[0, colNum])
                                {
                                    abilitySheet[i, j] = engineerSheet[i, colNum];
                                }
                            }
                        }
                    }
                    int colNumber = 0;
                    int velocityCol = 0;
                    foreach (string colname in engineerSheet)
                    {
                        colNumber++;
                        if (colname == "V")
                        {
                            velocityCol = colNumber;
                        }

                        if (velocityCol != 0) break;
                    }
                    for (int i = 1; i < abilitysheetRows; i++)
                    {
                        for (int j = velocityCol; j < cols; j++)
                        {
                            if (engineerSheet[i, j] == "1")
                            {
                                if (abilitySheet[i, abilitysheetCols - 1] == null)
                                {
                                    abilitySheet[i, abilitysheetCols - 1] += engineerSheet[0, j];
                                }
                                else
                                {
                                    abilitySheet[i, abilitysheetCols - 1] += "," + engineerSheet[0, j];
                                }
                            }
                        }
                    }

                    int usernameCol = 0;
                    foreach (DataColumn column in userDS.Tables[0].Columns)
                    {
                        usernameCol++;
                        if (column.ColumnName.ToString().Equals("Username")) break;
                    }
                    SqlCommand cmd = null;
                    string sql = null;
                    for (int i = 1; i < userSheet.GetLength(0); i++)
                    {
                        //判断用户名是否已经存在，不存在则存入数据，存在则跳过
                        sql = "select count(*) from usertable where userName='" + userSheet[i, usernameCol - 1] + "'";
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
                            cmd = new SqlCommand(sql, conn);
                            cmd.ExecuteNonQuery();
                        }
                    }

                    usernameCol = 0;
                    foreach (DataColumn column in abilityDS.Tables[0].Columns)
                    {
                        usernameCol++;
                        if (column.ColumnName.ToString().Equals("Username")) break;
                    }
                    for (int i = 1; i < abilitySheet.GetLength(0); i++)
                    {
                        //判断用户名是否已经存在，不存在则存入数据，存在则跳过
                        sql = "select count(*) from abilitytable where userName='" + abilitySheet[i, usernameCol - 1] + "'";
                        cmd = new SqlCommand(sql, conn);
                        bool usernameExist = (int)cmd.ExecuteScalar() > 0;
                        if (usernameExist)
                        {
                            continue;
                        }
                        else
                        {
                            sql = "insert into abilitytable values('" + abilitySheet[i, 0] + "','" + abilitySheet[i, 1] + "'," + abilitySheet[i, 2] + ",'" + abilitySheet[i, 3] + "')";
                            cmd = new SqlCommand(sql, conn);
                            cmd.ExecuteNonQuery();
                        }

                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }
    }
}
