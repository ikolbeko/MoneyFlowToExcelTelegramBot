namespace MoneyFlowToExcelTelegramBot;

internal class Transaction
{
    public DateTime Date { get; }
    public decimal Amount { get; }
    public string Currency { get; }
    public string Account { get; }
    public string Transfer { get; }
    public string Category { get; }
    public string Note { get; }
    public string Tags { get; }
    public string Counterparty { get; }
    public string Place { get; }

    public Transaction(DateTime date, decimal amount, string currency, string account, string transfer, string category, string note, string tags, string counterparty, string place)
    {
        Date = date;
        Amount = amount;
        Currency = currency;
        Account = account;
        Transfer = transfer;
        Category = category;
        Note = note;
        Tags = tags;
        Counterparty = counterparty;
        Place = place;
    }

    public override string ToString()
    {
        return $"{Date}: {Amount} {Currency}, {Account}, {Transfer}, {Category}, {Note}, {Tags}, {Counterparty}, {Place}";
    }
}
