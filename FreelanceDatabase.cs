using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml;

namespace Lab5
{
    internal class FreelanceDatabase
    {
        public const int MaxAge = 150;

        // Excel хранит до 15 значащих цифр.
        public const decimal MaxPrice = 9_999_999_999_999.99m;
        private const int MaxRowCount = 1_048_575;
        private readonly List<Order> _orders;
        private readonly List<Service> _services;
        private readonly List<Contractor> _contractors;

        public FreelanceDatabase(List<Order> orders,
            List<Service> services, List<Contractor> contractors)
        {
            UniqueKeys(orders.Select(order => order.Id),
                "Заказы");
            HashSet<int> serviceIds = UniqueKeys(
                services.Select(service => service.Id),
                "Услуги");
            HashSet<int> contractorIds = UniqueKeys(
                contractors.Select(person => person.Id),
                "Исполнители");
            Order? invalid = orders.FirstOrDefault(order =>
                !serviceIds.Contains(order.ServiceId)
                || !contractorIds.Contains(order.ContractorId));
            if (invalid != null)
            {
                throw new InvalidDataException(
                    $"Заказ {invalid.Id}: нарушены связи.");
            }
            _orders = orders;
            _services = services;
            _contractors = contractors;
        }

        public IReadOnlyList<Order> Orders => _orders;
        public IReadOnlyList<Service> Services => _services;
        public IReadOnlyList<Contractor> Contractors
            => _contractors;

        public static FreelanceDatabase Load(string path) =>
            FreelanceFile.Load(path);

        public void Save(string path) =>
            FreelanceFile.Save(path, this);

        public void AddOrder(int id, int serviceId,
            int contractorId, decimal price)
        {
            CheckNewKey(id, _orders.Select(order => order.Id),
                _orders.Count);
            if (!_services.Any(service =>
                service.Id == serviceId))
            {
                throw new ArgumentException(
                    "Услуга не найдена.");
            }
            if (!_contractors.Any(person =>
                person.Id == contractorId))
            {
                throw new ArgumentException(
                    "Исполнитель не найден.");
            }
            if (!IsValidPrice(price))
            {
                throw new ArgumentException(
                    "Недопустимая стоимость заказа.");
            }
            _orders.Add(new Order(id, serviceId, contractorId,
                price));
        }

        public void AddService(int id, string name)
        {
            CheckNewKey(id,
                _services.Select(service => service.Id),
                _services.Count);
            _services.Add(new Service(id, ValidateText(name)));
        }

        public void AddContractor(int id, int age,
            string citizenship)
        {
            CheckNewKey(id,
                _contractors.Select(person => person.Id),
                _contractors.Count);
            if (age < 1 || age > MaxAge)
            {
                throw new ArgumentException(
                    "Недопустимый возраст.");
            }
            _contractors.Add(new Contractor(
                id, age, ValidateText(citizenship)));
        }

        public bool RemoveOrder(int id) =>
            Remove(_orders, order => order.Id == id);

        public bool RemoveService(int id)
        {
            if (_orders.Any(order => order.ServiceId == id))
            {
                throw new ArgumentException(
                    "На эту услугу ссылаются заказы.");
            }
            return Remove(_services,
                service => service.Id == id);
        }

        public bool RemoveContractor(int id)
        {
            if (_orders.Any(order => order.ContractorId == id))
            {
                throw new ArgumentException(
                    "На этого исполнителя ссылаются заказы.");
            }
            return Remove(_contractors,
                person => person.Id == id);
        }

        public IEnumerable<Service> FindServices(string text) =>
            from service in _services
            where service.Name.Contains(text,
                StringComparison.OrdinalIgnoreCase)
            orderby service.Id
            select service;

        public decimal SumOrdersByService(string name)
        {
            var prices =
                from order in _orders
                join service in _services
                    on order.ServiceId equals service.Id
                where TextEquals(service.Name, name)
                select order.Price;
            return prices.Sum();
        }

        public IEnumerable<(int Id, string Name, decimal Total)>
            SumServicesByCitizenship(string citizenship) =>
            from entry in OrdersWithDetails()
            where TextEquals(
                entry.Person.Citizenship, citizenship)
            group entry.Price by entry.Service into prices
            let total = prices.Sum()
            orderby total descending, prices.Key.Id
            select (prices.Key.Id, prices.Key.Name, total);

        public int? FindTopContractor(string citizenship,
            string serviceName, int minimumAge)
        {
            var ids =
                from entry in OrdersWithDetails()
                where entry.Person.Age > minimumAge
                    && TextEquals(
                        entry.Person.Citizenship, citizenship)
                    && TextEquals(
                        entry.Service.Name, serviceName)
                group entry.Price by entry.Person.Id into prices
                orderby prices.Sum() descending, prices.Key
                select (int?)prices.Key;
            return ids.FirstOrDefault();
        }

        private IEnumerable<(decimal Price, Service Service,
            Contractor Person)> OrdersWithDetails() =>
            from order in _orders
            join service in _services
                on order.ServiceId equals service.Id
            join person in _contractors
                on order.ContractorId equals person.Id
            select (order.Price, service, person);

        public static bool IsValidPrice(decimal price) =>
            price >= 0 && price <= MaxPrice
            && decimal.Round(price, 2) == price;

        public static string ValidateText(string value)
        {
            value = value.Trim();
            if (value.Length == 0 || value.Length > 32767
                || value.Any(char.IsControl))
            {
                throw new ArgumentException(
                    "Нужен непустой текст до 32767 знаков.");
            }
            try
            {
                return XmlConvert.VerifyXmlChars(value);
            }
            catch (XmlException)
            {
                throw new ArgumentException(
                    "Текст содержит недопустимый символ.");
            }
        }

        private static HashSet<int> UniqueKeys(
            IEnumerable<int> keys, string table)
        {
            var unique = new HashSet<int>();
            foreach (int key in keys)
            {
                if (!unique.Add(key))
                {
                    throw new InvalidDataException(
                        $"«{table}»: повтор кода {key}.");
                }
            }
            return unique;
        }

        private static void CheckNewKey(
            int id, IEnumerable<int> keys, int count)
        {
            if (id <= 0)
            {
                throw new ArgumentException(
                    "Код должен быть положительным.");
            }
            if (keys.Contains(id))
            {
                throw new ArgumentException(
                    "Код уже существует.");
            }
            if (count == MaxRowCount)
            {
                throw new ArgumentException(
                    "Лист Excel заполнен.");
            }
        }

        private static bool Remove<T>(
            List<T> rows, Func<T, bool> matches)
            where T : class =>
            rows.FirstOrDefault(matches) is T row
            && rows.Remove(row);

        private static bool TextEquals(string left, string right)
            => string.Equals(left, right,
                StringComparison.OrdinalIgnoreCase);
    }
}
