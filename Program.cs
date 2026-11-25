using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;

namespace Lab5
{
    internal class Program
    {
        private const int PageSize = 20;
        private const string DefaultOutput = "Freelance.xlsx";

        private FreelanceDatabase? _database;
        private bool _changed;

        private static string SourcePath =>
            Path.Combine(AppContext.BaseDirectory,
                "LR5-var6.xls");

        private static void Main()
        {
            Console.OutputEncoding = Encoding.UTF8;
            Encoding.RegisterProvider(
                CodePagesEncodingProvider.Instance);
            try
            {
                new Program().Run();
            }
            catch (EndOfStreamException)
            {
                Console.WriteLine("Ввод завершён.");
            }
        }

        private void Run()
        {
            Console.WriteLine(
                "Лабораторная 5. Вариант 6. Фриланс");
            TryExecute(() => Open(SourcePath));
            while (true)
            {
                Console.WriteLine("""

                    1 - Открыть книгу Excel
                    2 - Просмотр таблицы
                    3 - Удаление по коду
                    4 - Добавление записи
                    5 - Запросы LINQ
                    6 - Сохранить книгу XLSX
                    0 - Выход
                    """);
                if (_changed)
                {
                    Console.WriteLine(
                        "Есть несохранённые изменения.");
                }
                int command = ConsoleInput.ReadRange(
                    "Команда: ", 0, 6);
                if (command != 0)
                {
                    TryExecute(() => Execute(command));
                }
                else if (CanDiscardChanges())
                {
                    return;
                }
            }
        }

        private void Execute(int command)
        {
            if (command == 1)
            {
                if (CanDiscardChanges())
                {
                    Open(ReadPath("Путь к книге", SourcePath));
                }
                return;
            }

            FreelanceDatabase database = _database
                ?? throw new InvalidOperationException(
                    "Сначала откройте книгу Excel.");
            switch (command)
            {
                case 2:
                    ShowTable(database);
                    break;
                case 3:
                    RemoveRow(database);
                    break;
                case 4:
                    AddRow(database);
                    break;
                case 5:
                    RunQuery(database);
                    break;
                case 6:
                    Save(database);
                    break;
            }
        }

        private void Open(string path)
        {
            FreelanceDatabase database =
                FreelanceDatabase.Load(path);
            _database = database;
            _changed = false;
            Console.WriteLine($"""
                Открыта книга {Path.GetFileName(path)}
                Заказы: {database.Orders.Count}
                Услуги: {database.Services.Count}
                Исполнители: {database.Contractors.Count}
                """);
        }

        private void Save(FreelanceDatabase database)
        {
            string path = ReadPath("Путь для сохранения",
                DefaultOutput);
            if (File.Exists(path)
                && ConsoleInput.ReadRange(
                    "Заменить файл? (1 - да, 0 - нет): ",
                    0, 1) == 0)
            {
                return;
            }
            database.Save(path);
            _changed = false;
            Console.WriteLine(
                $"Сохранено: {Path.GetFullPath(path)}");
        }

        private bool CanDiscardChanges() =>
            !_changed || ConsoleInput.ReadRange(
                "Продолжить без сохранения? (1 - да, 0 - нет): ",
                0, 1) == 1;

        private static void ShowTable(FreelanceDatabase database)
        {
            switch (ReadTable())
            {
                case 1:
                    PrintPages(database.Orders, "Заказы");
                    break;
                case 2:
                    PrintPages(database.Services,
                        "Код | Название");
                    break;
                case 3:
                    PrintPages(database.Contractors,
                        "Код | Возраст | Гражданство");
                    break;
            }
        }

        private static void PrintPages<T>(
            IReadOnlyList<T> rows, string title,
            string emptyMessage = "Таблица пуста.")
        {
            if (rows.Count == 0)
            {
                Console.WriteLine(emptyMessage);
                return;
            }
            int page = 1;
            int pageCount =
                (rows.Count + PageSize - 1) / PageSize;
            while (page != 0)
            {
                Console.WriteLine(title);
                foreach (T row in rows
                    .Skip((page - 1) * PageSize).Take(PageSize))
                {
                    Console.WriteLine(row);
                }
                Console.WriteLine(
                    $"Страница {page} из {pageCount}");
                if (pageCount == 1)
                {
                    return;
                }
                page = ConsoleInput.ReadRange(
                    "Номер страницы (0 - назад): ",
                    0, pageCount);
            }
        }

        private void RemoveRow(FreelanceDatabase database)
        {
            int table = ReadTable();
            if (table == 0)
            {
                return;
            }
            int id = ReadId("Код для удаления: ");
            bool removed = table switch
            {
                1 => database.RemoveOrder(id),
                2 => database.RemoveService(id),
                _ => database.RemoveContractor(id)
            };
            _changed = _changed || removed;
            Console.WriteLine(removed
                ? "Запись удалена." : "Код не найден.");
        }

        private void AddRow(FreelanceDatabase database)
        {
            int table = ReadTable();
            if (table == 0)
            {
                return;
            }
            int id = ReadId("Новый код: ");
            switch (table)
            {
                case 1:
                    int serviceId = ReadId("Код услуги: ");
                    int contractorId = ReadId(
                        "Код исполнителя: ");
                    decimal price = ConsoleInput.ReadPrice(
                        "Стоимость, руб.: ");
                    database.AddOrder(
                        id, serviceId, contractorId, price);
                    break;
                case 2:
                    string name = ConsoleInput.ReadText(
                        "Название: ");
                    database.AddService(id, name);
                    break;
                case 3:
                    int age = ConsoleInput.ReadRange(
                        "Возраст: ", 1,
                        FreelanceDatabase.MaxAge);
                    string country = ConsoleInput.ReadText(
                        "Гражданство: ");
                    database.AddContractor(id, age, country);
                    break;
            }
            _changed = true;
            Console.WriteLine("Запись добавлена.");
        }

        private static void RunQuery(FreelanceDatabase database)
        {
            Console.WriteLine("""
                1 - Услуги по фрагменту названия
                2 - Общая стоимость по услуге
                3 - Выручка по услугам для страны
                4 - Лучший Python-программист из Китая (31+)
                0 - Назад
                """);
            int query = ConsoleInput.ReadRange("Запрос: ", 0, 4);
            switch (query)
            {
                case 1:
                    string part = ConsoleInput.ReadText(
                        "Фрагмент названия: ");
                    PrintResults(database.FindServices(part));
                    break;
                case 2:
                    string name = ConsoleInput.ReadText(
                        "Полное название услуги: ");
                    decimal total =
                        database.SumOrdersByService(name);
                    Console.WriteLine($"Сумма: {total:0.##} ₽");
                    break;
                case 3:
                    string country = ConsoleInput.ReadText(
                        "Гражданство: ");
                    var results = database
                        .SumServicesByCitizenship(country)
                        .Select(service => $"{service.Id} | "
                            + $"{service.Name} | "
                            + $"{service.Total:0.##} ₽");
                    PrintResults(results);
                    break;
                case 4:
                    int? id = database.FindTopContractor(
                        "Китай", "Python-программист", 30);
                    Console.WriteLine(id.HasValue
                        ? $"Код исполнителя: {id}"
                        : "Подходящих заказов нет.");
                    break;
            }
        }

        private static void PrintResults<T>(
            IEnumerable<T> results)
        {
            PrintPages(results.ToList(), "Результат запроса",
                "Совпадений нет.");
        }

        private static int ReadTable()
        {
            Console.WriteLine("""
                1 - Заказы
                2 - Услуги
                3 - Исполнители
                0 - Назад
                """);
            return ConsoleInput.ReadRange("Таблица: ", 0, 3);
        }

        private static int ReadId(string message) =>
            ConsoleInput.ReadRange(message, 1, int.MaxValue);

        private static string ReadPath(
            string message, string defaultPath)
        {
            string fileName = Path.GetFileName(defaultPath);
            string path = ConsoleInput.ReadLine(
                $"{message} (Enter - {fileName}): ").Trim();
            return path.Length == 0
                ? defaultPath : path.Trim('"');
        }

        private static void TryExecute(Action action)
        {
            try
            {
                action();
            }
            catch (Exception error) when (
                error is IOException and not EndOfStreamException
                    or UnauthorizedAccessException
                    or InvalidDataException
                    or InvalidOperationException
                    or ArgumentException
                    or XmlException)
            {
                Console.WriteLine($"Ошибка! {error.Message}");
            }
        }
    }
}
