namespace DocumentGenerator.Tests
{
    public class FakeModel
    {
        public string Property1 { get; }
        public string Property2 { get; }
        public FakeSubModel SubModel { get; }
        public FakeItem[] Items { get; }

        public FakeModel(string property1, string property2, FakeSubModel subModel, FakeItem[] items)
        {
            Property1 = property1;
            Property2 = property2;
            SubModel = subModel;
            Items = items;
        }

        public class FakeSubModel
        {
            public string Property1 { get; }
            public string Property2 { get; }

            public FakeSubModel(string property1, string property2)
            {
                Property1 = property1;
                Property2 = property2;
            }
        }

        public class FakeItem
        {
            public string Property1 { get;  }
            public string Property2 { get; }

            public FakeItem(string property1, string property2)
            {
                Property1 = property1;
                Property2 = property2;
            }

        }
    }
}
