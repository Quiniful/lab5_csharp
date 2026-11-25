using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml.Linq;
using Sylvan.Data.Excel;

namespace Lab5
{
    internal static class FreelanceFile
    {
        private const string Orders = "Заказы";
        private const string Services = "Услуги";
        private const string Contractors = "Исполнители";

        private static readonly (string Name, Type Type)[]
            _orderColumns =
        [
            ("Код", typeof(int)),
            ("Код услуги", typeof(int)),
            ("Код исполнителя", typeof(int)),
            ("Стоимость", typeof(decimal))
        ];

        private static readonly (string Name, Type Type)[]
            _serviceColumns =
        [
            ("Код", typeof(int)), ("Название", typeof(string))
        ];

        private static readonly (string Name, Type Type)[]
            _contractorColumns =
        [
            ("Код", typeof(int)), ("Возраст", typeof(int)),
            ("Гражданство", typeof(string))
        ];

        public static FreelanceDatabase Load(string path)
        {
            string extension = Path.GetExtension(path)
                .ToLowerInvariant();
            if (extension is not (".xls" or ".xlsx"))
            {
                throw new InvalidDataException(
                    "Выберите файл .xls или .xlsx.");
            }
            try
            {
                using var reader = ExcelDataReader.Create(path,
                    new ExcelDataReaderOptions
                    {
                        Schema = ExcelSchema.NoHeaders,
                        ReadHiddenRows = true,
                        ReadHiddenWorksheets = true
                    });
                List<Order> orders = ReadTable(reader, Orders,
                    _orderColumns, cells => new Order(
                        ReadInt(cells[0]), ReadInt(cells[1]),
                        ReadInt(cells[2]), ReadPrice(cells[3])));
                List<Service> services = ReadTable(reader,
                    Services, _serviceColumns,
                    cells => new Service(ReadInt(cells[0]),
                        ReadText(cells[1])));
                List<Contractor> contractors = ReadTable(reader,
                    Contractors, _contractorColumns,
                    cells => new Contractor(ReadInt(cells[0]),
                        ReadInt(cells[1],
                            FreelanceDatabase.MaxAge),
                        ReadText(cells[2])));
                return new FreelanceDatabase(
                    orders, services, contractors);
            }
            catch (Exception error) when (
                error is IOException
                    and not (FileNotFoundException
                        or DirectoryNotFoundException)
                or FormatException
                or ArgumentOutOfRangeException
                or IndexOutOfRangeException
                or NotSupportedException)
            {
                throw new InvalidDataException(
                    "Не удалось прочитать книгу Excel.", error);
            }
        }

        public static void Save(
            string path, FreelanceDatabase database)
        {
            if (!string.Equals(Path.GetExtension(path), ".xlsx",
                StringComparison.OrdinalIgnoreCase))
            {
                throw new ArgumentException(
                    "Укажите имя файла с расширением .xlsx.");
            }
            string destination = Path.GetFullPath(path);
            string temporary = Path.Combine(
                Path.GetDirectoryName(destination)!,
                $".{Guid.NewGuid():N}.tmp");
            FileStream stream = File.Open(
                temporary, FileMode.CreateNew);
            try
            {
                using (stream)
                {
                    WriteWorkbook(stream, database);
                    FixDefaultFill(stream);
                }
                File.Move(temporary, destination, true);
            }
            finally
            {
                File.Delete(temporary);
            }
        }

        private static void WriteWorkbook(Stream stream,
            FreelanceDatabase database)
        {
            using var writer = ExcelDataWriter.Create(
                stream, ExcelWorkbookType.ExcelXml,
                new() { OwnsStream = false });
            WriteTable(writer, Orders, _orderColumns,
                database.Orders, order =>
                [
                    order.Id, order.ServiceId,
                    order.ContractorId, order.Price
                ]);
            WriteTable(writer, Services, _serviceColumns,
                database.Services, service =>
                [service.Id, service.Name]);
            WriteTable(writer, Contractors, _contractorColumns,
                database.Contractors, person =>
                [person.Id, person.Age, person.Citizenship]);
        }

        private static void WriteTable<T>(ExcelDataWriter writer,
            string name, (string Name, Type Type)[] columns,
            IEnumerable<T> rows, Func<T, object[]> toCells)
        {
            using var table = new DataTable();
            foreach ((string header, Type type) in columns)
            {
                table.Columns.Add(header, type);
            }
            foreach (T row in rows)
            {
                table.Rows.Add(toCells(row));
            }
            using var reader = table.CreateDataReader();
            writer.Write(reader, name);
        }

        private static void FixDefaultFill(Stream stream)
        {
            // Sylvan 0.5.8 создаёт пустой элемент fill.
            stream.Position = 0;
            using var archive = new ZipArchive(
                stream, ZipArchiveMode.Update, leaveOpen: true);
            ZipArchiveEntry entry = archive
                .GetEntry("xl/styles.xml")
                ?? throw new InvalidDataException(
                    "В книге нет таблицы стилей.");
            using Stream styles = entry.Open();
            XElement root = XElement.Load(styles);
            XNamespace ns = root.Name.Namespace;
            XElement? fill = root.Element(ns + "fills")
                ?.Element(ns + "fill");
            if (fill != null && !fill.HasElements)
            {
                fill.Add(new XElement(ns + "patternFill",
                    new XAttribute("patternType", "none")));
                styles.SetLength(0);
                root.Save(styles);
            }
        }

        private static List<T> ReadTable<T>(
            ExcelDataReader reader, string name,
            (string Name, Type Type)[] columns,
            Func<object[], T> create)
        {
            int matches = reader.WorksheetNames
                .Count(sheet => sheet == name);
            if (matches != 1 || !reader.TryOpenWorksheet(name))
            {
                throw new InvalidDataException(
                    $"В книге должен быть один лист «{name}».");
            }
            if (!reader.Read() || reader.RowNumber != 1
                || reader.RowFieldCount == 0)
            {
                throw new InvalidDataException(
                    $"Лист «{name}»: нет строки заголовка.");
            }

            string?[] headers = Enumerable
                .Range(0, reader.RowFieldCount)
                .Select(index => reader.GetExcelValue(index))
                .Select(value => (value as string)?.Trim())
                .ToArray();
            int[] positions = columns
                .Select(column =>
                    FindColumn(headers, name, column.Name))
                .ToArray();
            List<T> result = new List<T>();
            while (reader.Read())
            {
                object[] cells = positions.Select(index =>
                    index < reader.RowFieldCount
                        ? reader.GetExcelValue(index)
                        : DBNull.Value).ToArray();
                int missing =
                    cells.Count(cell => cell is DBNull);
                if (missing == cells.Length)
                {
                    continue;
                }
                try
                {
                    if (missing > 0)
                    {
                        throw new ArgumentException(
                            "Заполнены не все ячейки.");
                    }
                    result.Add(create(cells));
                }
                catch (ArgumentException error)
                {
                    throw new InvalidDataException(
                        $"«{name}», строка {reader.RowNumber}. "
                        + error.Message, error);
                }
            }
            return result;
        }

        private static int FindColumn(string?[] headers,
            string sheet, string name)
        {
            int position = Array.IndexOf(headers, name);
            if (position < 0
                || position != Array.LastIndexOf(headers, name))
            {
                throw new InvalidDataException(
                    $"Проверьте «{name}» на листе «{sheet}».");
            }
            return position;
        }

        private static int ReadInt(
            object value, int max = int.MaxValue) =>
            value is double number && number >= 1
            && number <= max && number == Math.Truncate(number)
                ? (int)number
                : throw new ArgumentException(
                    $"Ожидалось целое число от 1 до {max}.");

        private static decimal ReadPrice(object value)
        {
            if (value is double number && number >= 0
                && number <= (double)FreelanceDatabase.MaxPrice)
            {
                decimal price = (decimal)number;
                if (FreelanceDatabase.IsValidPrice(price))
                {
                    return price;
                }
            }
            throw new ArgumentException(
                "Ожидалась стоимость в рублях и копейках.");
        }

        private static string ReadText(object value) =>
            value is string text
                ? FreelanceDatabase.ValidateText(text)
                : throw new ArgumentException(
                    "Ожидалось текстовое поле.");
    }
}
