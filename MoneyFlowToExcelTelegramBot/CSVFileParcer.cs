using System.Globalization;
using System.Transactions;

namespace MoneyFlowToExcelTelegramBot;

class CSVFileParcer
{
    public int dataYear;
    public void Parce(string destinationFilePath, long chatId)
    {
        List<Transaction> incomeTransactions = ParseCSV(destinationFilePath, true);
        List<Transaction> expenseTransactions = ParseCSV(destinationFilePath, false);

        try
        {
            dataYear = expenseTransactions[0].Date.Year;
        }
        catch
        {
            dataYear = 0;
        }

        // Create the Excel file
        var excel = new ExcelFileCreator();
        excel.Create(incomeTransactions, expenseTransactions, chatId, dataYear);

        Console.WriteLine("Excel file created successfully.");
    }

    private List<Transaction> ParseCSV(string filePath, bool includePositiveAmounts)
    {
        List<Transaction> transactions = new List<Transaction>();

        using (StreamReader reader = new StreamReader(filePath))
        {
            // Skip the header line
            reader.ReadLine();

            // Read each line and parse the transaction
            string line;
            while ((line = reader.ReadLine()) != null)
            {
                string[] fields = line.Split(',');

                DateTime date = DateTime.Parse(fields[0]);
                decimal amount = decimal.Parse(fields[1], NumberStyles.Any, CultureInfo.InvariantCulture);
                string currency = fields[2];
                string account = fields[3];
                string transfer = fields[4];
                string category = fields[5];
                string note = fields[6];
                string tags = fields[7];
                string counterparty = fields[8];
                string place = fields[9];

                Transaction transaction = new Transaction(date, Math.Abs(amount), currency, account, transfer, category, note, tags, counterparty, place);

                if (includePositiveAmounts && amount >= 0)
                {
                    transactions.Add(transaction);
                }
                else if (!includePositiveAmounts && amount < 0)
                {
                    transactions.Add(transaction);
                }
            }
        }

        return transactions;
    }
}