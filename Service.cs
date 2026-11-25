namespace Lab5
{
    internal class Service(int id, string name)
    {
        public int Id { get; } = id;
        public string Name { get; } = name;

        public override string ToString() => $"{Id} | {Name}";
    }
}
