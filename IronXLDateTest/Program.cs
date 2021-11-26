using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;

namespace IronXLDateTest
{
    class Program
    {
        static void Main(string[] args)
        {
            var employeeManager = new EmployeeManager();
            var employees = employeeManager.GetEmployees();
            var dataTable = employeeManager.GetDataTableFromCollection(employees);
            var file = employeeManager.CreateFileXls(dataTable);
            System.IO.File.WriteAllBytes($@"{System.IO.Path.GetTempPath()}Filename.xls", file);
        }

    }
}
