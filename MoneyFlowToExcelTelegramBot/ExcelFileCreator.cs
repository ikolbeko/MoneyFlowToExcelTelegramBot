using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Drawing;

namespace MoneyFlowToExcelTelegramBot;

internal class ExcelFileCreator
{
    private int RowIndex = 2;
    private int currentRowIndex = 0;

    public void Create(List<Transaction> incomeTransactions, List<Transaction> expenseTransactions, long chatId, int dataYear)
    {
        using (ExcelPackage excelPackage = new ExcelPackage())
        {
            ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Report");
            SetHeaderRow(worksheet);

            int expenseRowIndex = PopulateCategoryData(expenseTransactions, worksheet, 2);
            TotalRow(worksheet, "Total Expenses", expenseRowIndex, System.Drawing.Color.FromArgb(255, 204, 204));
            MakeBorderStyle(worksheet, "A1", "O" + currentRowIndex + "");

            RowIndex = expenseRowIndex + 2;
            int incomeRowIndex = PopulateCategoryData(incomeTransactions, worksheet, expenseRowIndex + 2);
            TotalRow(worksheet, "Total Income", incomeRowIndex, System.Drawing.Color.FromArgb(204, 255, 204));
            MakeBorderStyle(worksheet, "A" + RowIndex + "", "O" + currentRowIndex + "");

            worksheet.Cells.AutoFitColumns();

            string filePath = "excel/" + chatId + "-" + dataYear + ".xlsx";
            FileInfo excelFile = new FileInfo(filePath);
            excelPackage.SaveAs(excelFile);
        }
    }

    private void SetHeaderRow(ExcelWorksheet worksheet)
    {
        worksheet.Cells[1, 1].Value = "Item of expenses";
        worksheet.Cells[1, 2].Value = "Average";
        worksheet.Cells[1, 3].Value = "Amount for the year";

        ExcelRange range = worksheet.Cells["A1:O1"];
        range.Style.Fill.PatternType = ExcelFillStyle.Solid;
        range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(153, 204, 255));

        DateTime currentMonth = new DateTime(DateTime.Now.Year, 1, 1);
        for (int i = 0; i < 12; i++)
        {
            worksheet.Cells[1, i + 4].Value = currentMonth.ToString("MMMM");
            currentMonth = currentMonth.AddMonths(1);
        }
    }

    private int PopulateCategoryData(List<Transaction> transactions, ExcelWorksheet worksheet, int rowIndex)
    {
        List<string> uniqueCategories = new List<string>();
        currentRowIndex = rowIndex;

        foreach (Transaction transaction in transactions)
        {
            if (!uniqueCategories.Contains(transaction.Category))
            {
                uniqueCategories.Add(transaction.Category);

                worksheet.Cells[currentRowIndex, 1].Value = transaction.Category;
                worksheet.Cells[currentRowIndex, 2].Formula = "=ROUND(AVERAGE(D" + currentRowIndex + ":O" + currentRowIndex + "), 2)";
                worksheet.Cells[currentRowIndex, 3].Formula = "=SUM(D" + currentRowIndex + ":O" + currentRowIndex + ")";

                int monthColumn = 4;
                DateTime currentMonth = new DateTime(transaction.Date.Year, 1, 1);
                for (int j = 0; j < 12; j++)
                {
                    decimal monthValue = CalculateMonthValue(transactions, transaction.Category, currentMonth);
                    worksheet.Cells[currentRowIndex, monthColumn].Value = monthValue;
                    monthColumn++;
                    currentMonth = currentMonth.AddMonths(1);
                }

                currentRowIndex++;
            }
        }

        return currentRowIndex;
    }

    private void TotalRow(ExcelWorksheet worksheet, string rowLabel, int rowIndex, Color color)
    {
        worksheet.Cells[rowIndex, 1].Value = rowLabel;
        worksheet.Cells[rowIndex, 2].Formula = "=ROUND(AVERAGE(D" + rowIndex + ":O" + rowIndex + "), 2)";
        worksheet.Cells[rowIndex, 3].Formula = "=SUM(C" + RowIndex + ":C" + (rowIndex - 1) + ")";

        ExcelRange range = worksheet.Cells["" + worksheet.Cells[rowIndex, 1].Address + ":" + worksheet.Cells[rowIndex, 15].Address + ""];
        range.Style.Fill.PatternType = ExcelFillStyle.Solid;
        range.Style.Fill.BackgroundColor.SetColor(color);

        int totalRowMonthColumn = 4;
        DateTime currentMonth = new DateTime(DateTime.Now.Year, 1, 1);
        for (int j = 0; j < 12; j++)
        {
            worksheet.Cells[rowIndex, totalRowMonthColumn].Formula = @"=SUM(" + worksheet.Cells[RowIndex, totalRowMonthColumn].Address + ":" +
                "" + worksheet.Cells[currentRowIndex - 1, totalRowMonthColumn].Address + ")";

            totalRowMonthColumn++;
            currentMonth = currentMonth.AddMonths(1);
        }
    }

    private decimal CalculateMonthValue(List<Transaction> transactions, string category, DateTime month)
    {
        decimal monthValue = 0;

        foreach (Transaction transaction in transactions)
        {
            if (transaction.Category == category && transaction.Date.Month == month.Month && transaction.Date.Year == month.Year)
            {
                monthValue += transaction.Amount;
            }
        }

        return monthValue;
    }

    private void MakeBorderStyle(ExcelWorksheet sheet, string firstCell, string secondCell)
    {
        ExcelRange range = sheet.Cells["" + firstCell + ":" + secondCell + ""];
        range.Style.Border.Top.Style = ExcelBorderStyle.Thin;
        range.Style.Border.Left.Style = ExcelBorderStyle.Thin;
        range.Style.Border.Right.Style = ExcelBorderStyle.Thin;
        range.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
    }
}
