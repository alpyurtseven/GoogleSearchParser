using System;
using ClosedXML.Excel;

namespace GoogleSearchParser
{
    public class ExcelHelper
    {
        public List<CompanyModel> ReadExcel(string filePath)
        {
            var companies = new List<CompanyModel>();
   
            using (var workbook = new XLWorkbook(filePath))
            {
                var worksheet = workbook.Worksheet(1);
                var range = worksheet.RangeUsed();

                for (int row = 2; row <= range.RowCount(); row++)
                {
                    var company = new CompanyModel
                    {
                        Name = range.Cell(row, 1).GetValue<string>(),
                        Telephone = range.Cell(row, 2).GetValue<string>(),
                        WorkingHours = range.Cell(row, 3).GetValue<string>(),
                        Website = range.Cell(row, 4).GetValue<string>()
                    };

                    companies.Add(company);
                }
            }
 
            return companies;
        }

        public void UpdateExcel(string filePath, List<CompanyModel> companies)
        {
            using (var workbook = new XLWorkbook(filePath))
            {
                var worksheet = workbook.Worksheet(1);

                int lastRow = worksheet.LastRowUsed().RowNumber();
                worksheet.Range(2, 1, lastRow, 4).Clear();

                for (int i = 0; i < companies.Count; i++)
                {
                    var company = companies[i];
                    worksheet.Cell(i + 2, 1).Value = company.Name;
                    worksheet.Cell(i + 2, 2).Value = company.Telephone;
                    worksheet.Cell(i + 2, 3).Value = company.WorkingHours;
                    worksheet.Cell(i + 2, 4).Value = company.Website;
                }

                workbook.Save();
            }
        }
    }

}

