namespace Lab5
{
    internal class Contractor(
        int id, int age, string citizenship)
    {
        public int Id { get; } = id;
        public int Age { get; } = age;
        public string Citizenship { get; } = citizenship;

        public override string ToString() =>
            $"{Id} | {Age} | {Citizenship}";
    }
}
