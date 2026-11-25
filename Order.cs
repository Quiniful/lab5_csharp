namespace Lab5
{
    internal class Order(
        int id, int serviceId, int contractorId, decimal price)
    {
        public int Id { get; } = id;
        public int ServiceId { get; } = serviceId;
        public int ContractorId { get; } = contractorId;
        public decimal Price { get; } = price;

        public override string ToString() =>
            $"Заказ {Id} | услуга {ServiceId} | "
            + $"исполнитель {ContractorId} | {Price:0.##} ₽";
    }
}
