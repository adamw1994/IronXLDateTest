using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using IronXL;
using Newtonsoft.Json;

namespace IronXLDateTest
{
    public class EmployeeManager
    {
        public IEnumerable<EmployeeDto> GetEmployees()
        {
            var employees = new List<EmployeeDto>();
            for (int i = 0; i < 300; i++)
            {
                employees.Add(new EmployeeDto
                {
                    Date1 = "2021-06-25",
                    Date2 = "2021-06-25",
                    Date3 = "2021-06-25",
                    Date4 = "2021-06-25"
                });
            }
            return employees;
        }
        public DataTable GetDataTableFromCollection(IEnumerable<EmployeeDto> employees)
        {
            var jsonCollection = JsonConvert.SerializeObject(employees);
            return JsonConvert.DeserializeObject<DataTable>(jsonCollection);
        }

        public byte[] CreateFileXls(DataTable dataTable)
        {
            var wb = new WorkBook(ExcelFileFormat.XLS);
            var sheet = wb.CreateWorkSheet("Sheet1");

            for (int columnIndex = 0; columnIndex < dataTable.Columns.Count; columnIndex++)
            {
                sheet.SetCellValue(0, columnIndex, dataTable.Columns[columnIndex].ColumnName);
            }

            for (int rowIndex = 0; rowIndex < dataTable.Rows.Count; rowIndex++)
            {
                for (int columnIndex = 0; columnIndex < dataTable.Columns.Count; columnIndex++)
                {
                    sheet.SetCellValue(rowIndex + 1, columnIndex, dataTable.Rows[rowIndex][columnIndex]);
                }
            }

            foreach (var columnIndex in sheet.Rows[1].Where(x => x.IsDateTime).Select(x => x.ColumnIndex))
                sheet.GetColumn(columnIndex).FormatString = IronXL.Formatting.BuiltinFormats.ShortDate;
            return wb.ToByteArray();
        }
    }
}
