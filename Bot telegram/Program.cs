using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Telegram.Bot;
using Telegram.Bot.Polling;
using Telegram.Bot.Types;
using Telegram.Bot.Types.Enums;
using static System.Runtime.InteropServices.JavaScript.JSType;
using Telegram.Bot.Types.ReplyMarkups;

// Конфигурация бота
var botToken = "8389569202:AAE0khuQlcHh1TmuSduZSzgZ5KyCfU-7nZ8";
var botClient = new TelegramBotClient(botToken);
var sender = new TelegramBotSender(botToken);

// Настройки получения обновлений
var receiverOptions = new ReceiverOptions {
    AllowedUpdates = Array.Empty<UpdateType>(), // Получаем все типы обновлений
    DropPendingUpdates = true, // Игнорируем накопленные сообщения при старте
};

// Статистика бота
var me = await botClient.GetMe();
Console.WriteLine($"Бот @{me.Username} запущен и готов к работе!");

// Запускаем обработку сообщений
using var cts = new CancellationTokenSource();
botClient.StartReceiving(
    updateHandler: HandleUpdateAsync,
    errorHandler: HandlePollingErrorAsync,
    receiverOptions: receiverOptions,
    cancellationToken: cts.Token
);

// Бесконечное ожидание (Ctrl+C для остановки)
await Task.Delay(-1, cts.Token);

// Обработчик входящих сообщений
async Task HandleUpdateAsync(ITelegramBotClient botClient, Update update, CancellationToken ct) {
    try {
        // Обрабатываем только текстовые сообщения
        if (update.Message is not { Text: { } messageText } message)
            return;
        var chatId = message.Chat.Id;
        var username = message.From?.Username ?? "Пользователь";
        Console.WriteLine($"[{DateTime.Now}] {username}: {messageText}");
        // Ответные действия
        switch (messageText) {
            case "/start":
            await SendStartMenu(chatId, ct);
            break;

            case "/help":
            await SendHelpMessage(chatId, ct);
            break;

            case "/CurrentTask":
            ShowCurrentTask(chatId);
            break;

            case "/reload":
            SendReload(chatId, ct);
            break;

            case "/BurnSLA":
            ShowBurnSLA(chatId);
            break;


            default:
            await EchoMessage(chatId, messageText, ct);
            break;
        }
        
    } catch (Exception ex) {
        Console.WriteLine($"Ошибка обработки сообщения: {ex.Message}");
    }
}

// Обработчик ошибок
Task HandlePollingErrorAsync(ITelegramBotClient botClient, Exception error, CancellationToken ct) {
    Console.WriteLine($"Ошибка: {error.Message}");
    return Task.CompletedTask;
}

// Методы отправки сообщений
async Task SendStartMenu(long chatId, CancellationToken ct) {
    var menuText = "Добро пожаловать! Доступные команды:\n" +
                  "/help - Справка\n" +
                  "/menu - Основное меню\n" + 
                  "/CurrentTask - Заявки\n" +
                  "/start\n" + "/BurnSLA \n";

    await botClient.SendMessage(
        chatId: chatId,
        text: menuText,
        cancellationToken: ct);
}
async Task SendHelpMessage(long chatId, CancellationToken ct) {
    await botClient.SendMessage(
        chatId: chatId,
        text: "Это бот-помощник. Отправьте любое сообщение для эхо-ответа.",
        cancellationToken: ct);
}
async Task EchoMessage(long chatId, string text, CancellationToken ct) {
    await botClient.SendMessage(
        chatId: chatId,
        text: $"Вы написали?: {text}",
        cancellationToken: ct);
}
async Task SendCurrentTaskMessage(long chatId, CancellationToken ct) {
    await botClient.SendMessage(
        chatId: chatId,
        text: $"Вот актуальный список заявок",
        cancellationToken: ct);
}
async Task SendReload(long chatId, CancellationToken ct) {
    await botClient.SendMessage(chatId: chatId, text: "reset", replyMarkup: new ReplyKeyboardRemove { Selective = false});
}


////Часть логики преобразования файлов
void ShowCurrentTask(long chatId) {
    var reader = new ExcelColumnIndexReader();
    string filePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "files", "export.xlsx");
    // Читаем столбцы с индексами 1, 3 и 5 (A, C, E)
    var data = reader.ReadColumnsByIndex(filePath, 1, 2, 3, 4, 5, 6, 7);
    string WriteOut;
    string SLA = null;
    // Выводим результаты
    for (int i = 0; i < data.Count; i++) {
        string excelDateStr = data[i][4];
        // Преобразуем строку в double
        if (double.TryParse(excelDateStr, NumberStyles.Any, CultureInfo.InvariantCulture, out double excelDate)) {
            // Конвертируем число Excel в DateTime
            DateTime date = DateTime.FromOADate(excelDate);
            SLA = $"SLA: {date.ToString("dd.MM.yyyy HH:mm:ss")}";
        } else {
            Console.WriteLine("Некорректный формат числа");
        }
        WriteOut = $"Ссылка: https://tabertrade.worktkur.ru/Task/View/{data[i][0]}\n" +
            $"Номер заявки: {data[i][0]}\n" +
            $"Код объекта: {data[i][1]}\n" +
            $"Город: {data[i][2]} \n" +
            $"Адрес: {data[i][3]}\n" +
            $"{SLA}";

        sender.SendTextToChatAsync(chatId, WriteOut);
        //Thread.Sleep(5000);
    }
}
void ShowBurnSLA (long chatId) {
    var reader = new ExcelColumnIndexReader();
    string filePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "files", "export.xlsx");
    // Читаем столбцы с индексами 1, 3 и 5 (A, C, E)
    var data = reader.ReadColumnsByIndex(filePath, 1, 2, 3, 4, 5, 6, 7);
    string WriteOut;
    // Выводим результаты
    for (int i = 0; i < data.Count; i++) {
        string excelDateStr = data[i][4];
        DateTime datesla = new DateTime();
        if (double.TryParse(excelDateStr, NumberStyles.Any, CultureInfo.InvariantCulture, out double excelDate)) {
            // Конвертируем число Excel в DateTime
            datesla = DateTime.FromOADate(excelDate);
        } else {
            Console.WriteLine("Некорректный формат числа");
        }
        TimeSpan SLATimer = datesla - DateTime.Now;
        TimeSpan SLADeadline = new TimeSpan(2, 00, 00); // Критерий проверки горящего СЛА
        if (SLATimer < SLADeadline) {
            //Console.WriteLine($"Task {data[i][0]} - SLA Timer {SLATimer.ToString("hh\\:mm\\:ss")}");
            WriteOut = $"Ссылка: https://tabertrade.worktkur.ru/Task/View/{data[i][0]}\n" +
                       $"Адрес: {data[i][3]}\n" +
                       $"SLA Timer {SLATimer.ToString("hh\\:mm\\:ss")}";
            sender.SendTextToChatAsync(chatId, WriteOut);
            } else {
            }
    }
    }

public class ExcelColumnIndexReader {
    public List<List<string>> ReadColumnsByIndex(string filePath, params int[] columnIndexes) {
        var result = new List<List<string>>();
        using (SpreadsheetDocument doc = SpreadsheetDocument.Open(filePath, false)) {
            WorkbookPart workbookPart = doc.WorkbookPart;
            WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();
            SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();
            foreach (Row row in sheetData.Elements<Row>().Skip(1)) {
                var rowData = new List<string>();
                var cells = row.Elements<Cell>().ToList();
                foreach (int colIndex in columnIndexes) {
                    // Excel использует индексацию с 1, а список с 0
                    if (colIndex <= cells.Count) {
                        Cell cell = cells[colIndex - 1]; // преобразуем в 0-based индекс
                        string cellValue = GetCellValue(cell, workbookPart);
                        rowData.Add(cellValue);
                    } else {
                        rowData.Add(string.Empty); // если столбец отсутствует
                    }
                }
                result.Add(rowData);
            }
        }
        return result;
    }
    private string GetCellValue(Cell cell, WorkbookPart workbookPart) {
        if (cell.DataType?.Value == CellValues.SharedString) {
            SharedStringTablePart stringTable = workbookPart.SharedStringTablePart;
            if (stringTable != null) {
                return stringTable.SharedStringTable.ElementAt(int.Parse(cell.InnerText)).InnerText;
            }
        }
        return cell.InnerText;
    }
}
public class TelegramBotSender {
    private readonly ITelegramBotClient _botClient;

    // Конструктор (инициализация бота)
    public TelegramBotSender(string botToken) {
        _botClient = new TelegramBotClient(botToken);
    }

    // Функция отправки текста в чат
    public async Task SendTextToChatAsync(long chatId, string textToSend) {
        try {
            await _botClient.SendMessage(
                chatId: chatId,
                text: textToSend);

            Console.WriteLine($"Сообщение отправлено в чат {chatId}");
        } catch (Exception ex) {
            Console.WriteLine($"Ошибка при отправке: {ex.Message}");
        }
    }
}