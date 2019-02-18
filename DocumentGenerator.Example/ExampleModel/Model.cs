using System;
using System.Numerics;

namespace DocumentGenerator.Example.ExampleModel
{
    public class Model
    {
        public string Property1 { get; }
        public string Property2 { get; }
        public int Property3 { get; }
        public decimal Property4 { get; }
        public DateTime Property5 { get; }
        public DateTimeOffset Property6 { get; }
        public TimeSpan Property7 { get; }
        public Vector2 Property8 { get; }
        public Guid Property9 { get; }
        public SubModel SubModel { get; }
        public Item[] Items { get; }

        public Model(string property1,
            string property2,
            int property3,
            decimal property4,
            DateTime property5,
            DateTimeOffset property6,
            TimeSpan property7,
            Vector2 property8,
            Guid property9,
            SubModel subModel,
            Item[] items)
        {
            Property1 = property1;
            Property2 = property2;
            Property3 = property3;
            Property4 = property4;
            Property5 = property5;
            Property6 = property6;
            Property7 = property7;
            Property8 = property8;
            Property9 = property9;
            SubModel = subModel;
            Items = items;
        }
    }
}
