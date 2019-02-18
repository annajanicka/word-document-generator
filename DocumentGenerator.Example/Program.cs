using DocumentGenerator.Example.ExampleModel;
using System;
using System.IO;
using System.Numerics;
using System.Reflection;

namespace DocumentGenerator.Example
{
    class Program
    {
        static void Main(string[] args)
        {
            var assembly = Assembly.GetExecutingAssembly();
            using (var stream = assembly.GetManifestResourceStream("DocumentGenerator.Example.form.docx"))
            using (MemoryStream memoryStream = new MemoryStream())
            {
                stream.CopyTo(memoryStream);
                var result = new WordDocumentGenerator().GetDocument(GetModel(), memoryStream.ToArray());
                {
                    File.WriteAllBytes("output.docx", result);
                }
            }
        }

        private static Model GetModel()
        {
            return new Model(
                Guid.NewGuid().ToString(),
                Guid.NewGuid().ToString(),
                123123,
                2.2M,
                DateTime.Now,
                DateTimeOffset.Now,
                TimeSpan.FromDays(123232),
                Vector2.UnitY,
                Guid.NewGuid(),
                new SubModel(Guid.NewGuid().ToString(), Guid.NewGuid().ToString()),
                new[]
                {
                    new Item(Guid.NewGuid().ToString(), Guid.NewGuid().ToString()),
                    new Item(Guid.NewGuid().ToString(), Guid.NewGuid().ToString()),
                    new Item(Guid.NewGuid().ToString(), Guid.NewGuid().ToString())
                });
        }
    }
}
