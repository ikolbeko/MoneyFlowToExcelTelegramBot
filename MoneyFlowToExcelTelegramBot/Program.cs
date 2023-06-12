using MoneyFlowToExcelTelegramBot;
using Telegram.Bot;
using Telegram.Bot.Exceptions;
using Telegram.Bot.Polling;
using Telegram.Bot.Types;
using Telegram.Bot.Types.Enums;


var botClient = new TelegramBotClient("6154595963:AAHJOo8lAThxUPpiHiiYHH1mk4NIUfFCHOI");

Directory.CreateDirectory("csv/");
Directory.CreateDirectory("excel/");

using CancellationTokenSource cts = new();

// StartReceiving does not block the caller thread. Receiving is done on the ThreadPool.
ReceiverOptions receiverOptions = new()
{
    AllowedUpdates = Array.Empty<UpdateType>() // receive all update types except ChatMember related updates
};

botClient.StartReceiving(
    updateHandler: HandleUpdateAsync,
    pollingErrorHandler: HandlePollingErrorAsync,
    receiverOptions: receiverOptions,
    cancellationToken: cts.Token
);

var me = await botClient.GetMeAsync();

Console.WriteLine($"Start listening for @{me.Username}");
Console.ReadLine();

// Send cancellation request to stop bot
cts.Cancel();

async Task HandleUpdateAsync(ITelegramBotClient botClient, Update update, CancellationToken cancellationToken)
{
    // Only process Message updates: https://core.telegram.org/bots/api#message
    if (update.Message is not { } message)
        return;

    if (message.Document is not { } document)
        return;

    var chatId = message.Chat.Id;
    var parcer = new CSVFileParcer();
    var fileId = document.FileId;
    var fileInfo = await botClient.GetFileAsync(fileId);
    var filePath = fileInfo.FilePath;

    Console.WriteLine($"Received a '{filePath}' message in chat {chatId}.");


    // Download
    string destinationFilePath = "csv/" + chatId + ".csv";

    using (Stream fileStream = System.IO.File.Create(destinationFilePath))
    {
        await botClient.DownloadFileAsync(
            filePath: filePath,
            destination: fileStream,
            cancellationToken: cancellationToken);
        Console.WriteLine("csv file: " + chatId + ".csv downloaded!");
    }

    // Parcing & Create Excel
    parcer.Parce(destinationFilePath, chatId);

    // Upload
    using (Stream stream = System.IO.File.OpenRead("excel/" + chatId + "-" + parcer.dataYear + ".xlsx"))
    {
        Message messageDocument = await botClient.SendDocumentAsync(
        chatId: chatId,
        document: InputFile.FromStream(stream: stream, fileName: parcer.dataYear + ".xlsx"));

        Console.WriteLine("excel file: " + chatId + " - " + parcer.dataYear + ".xlsx uploaded!");
    }
}

Task HandlePollingErrorAsync(ITelegramBotClient botClient, Exception exception, CancellationToken cancellationToken)
{
    var ErrorMessage = exception switch
    {
        ApiRequestException apiRequestException
            => $"Telegram API Error:\n[{apiRequestException.ErrorCode}]\n{apiRequestException.Message}",
        _ => exception.ToString()
    };

    Console.WriteLine(ErrorMessage);
    return Task.CompletedTask;
}
