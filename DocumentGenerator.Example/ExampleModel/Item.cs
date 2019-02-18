namespace DocumentGenerator.Example.ExampleModel
{
    public class Item
    {
        public string Property1 { get; }
        public string Property2 { get; }

        public Item(string property1, string property2)
        {
            Property1 = property1;
            Property2 = property2;
        }
    }
}
