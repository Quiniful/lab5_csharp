using System;
using System.Globalization;
using System.IO;

namespace Lab5
{
    internal static class ConsoleInput
    {
        public static string ReadLine(string message)
        {
            Console.Write(message);
            return Console.ReadLine()
                ?? throw new EndOfStreamException();
        }

        public static string ReadText(string message)
        {
            string input = ReadLine(message).Trim();
            while (input.Length == 0)
            {
                Console.WriteLine("Ошибка! Строка пустая.");
                input = ReadLine(message).Trim();
            }
            return input;
        }

        public static int ReadInt(string message)
        {
            int result;
            while (!int.TryParse(ReadLine(message), out result))
            {
                Console.WriteLine("Введите целое число Int32.");
            }
            return result;
        }

        public static int ReadRange(
            string message, int min, int max)
        {
            int result = ReadInt(message);
            while (result < min || result > max)
            {
                Console.WriteLine(
                    $"Введите число от {min} до {max}");
                result = ReadInt(message);
            }
            return result;
        }

        public static decimal ReadPrice(string message)
        {
            while (true)
            {
                string input = ReadLine(message).Trim()
                    .Replace(',', '.');
                const NumberStyles Style =
                    NumberStyles.AllowDecimalPoint
                    | NumberStyles.AllowLeadingSign;
                int point = input.IndexOf('.');
                bool validScale = point < 0
                    || input.AsSpan(point + 1).TrimEnd('0')
                        .Length <= 2;
                if (validScale && decimal.TryParse(input, Style,
                    CultureInfo.InvariantCulture,
                    out decimal result)
                    && FreelanceDatabase.IsValidPrice(result))
                {
                    return result;
                }
                Console.WriteLine(
                    "Введите стоимость в рублях и копейках.");
            }
        }
    }
}
