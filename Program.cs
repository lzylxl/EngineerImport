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
            //ImportExcelEngineer_SqlServer imp1 = new ImportExcelEngineer_SqlServer();
            //imp1.ImportExcelEngineer(filepath);
            ImportExcelEngineer_mysql imp2 = new ImportExcelEngineer_mysql();
            imp2.ImportExcelEngineer(filepath);
        }
    }
}
