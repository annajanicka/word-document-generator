using System.IO;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using NUnit.Framework;
using Shouldly;

namespace DocumentGenerator.Tests
{
    [TestFixture]
    public class WhenDocumentDoesNotContainModelProperties
    {
        private WordDocumentGenerator _formatter;
        [SetUp]
        public void Setup()
        {
            _formatter = new WordDocumentGenerator();
        }

        [Test]
        public void WhenDocumentDoesNotContainModelPropertiesReturnDocument()
        {
            var model = new FakeModel(
                "Value of property 1",
                "Value of property 2",
                new FakeModel.FakeSubModel("Value of property 1 of sub model", "Value of property 2 of sub model"),
                new []
                {
                    new FakeModel.FakeItem("Value of property 1 of fake item 1", "Value of property 2 of fake item 1"),
                    new FakeModel.FakeItem("Value of property 1 of fake item 1", "Value of property 2 of fake item 1")
                });
            var result = _formatter.GetDocument(model, GetDocument());

            using (var memoryStream = new MemoryStream(result))
            using (var document = WordprocessingDocument.Open(memoryStream, true))
            using (var streamReader = new StreamReader(document.MainDocumentPart.GetStream()))
            {
                var content = streamReader.ReadToEnd();
                content.ShouldContain(model.Property1);
                content.ShouldNotContain(model.Property2);
                content.ShouldContain(model.SubModel.Property1);
                content.ShouldNotContain(model.SubModel.Property2);

                foreach (var item in model.Items)
                {
                    content.ShouldContain(item.Property1);
                    content.ShouldNotContain(item.Property2);
                }

                var regex = new Regex("{model(.*?)}");
                var placeholdersToReplace = regex.Matches(content);
                placeholdersToReplace.Count.ShouldBe(0);
                Regex.Matches(content, "related-content").Count.ShouldBe(model.Items.Length);
            }
        }

        private byte[] GetDocument()
        {
            using (MemoryStream memoryStream = new MemoryStream())
            using (var wordDocument = WordprocessingDocument.Create(memoryStream, WordprocessingDocumentType.Document))
            {
                MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();
                mainPart.Document = new Document(
                    new Body(
                        new Paragraph(new Run(new Text("{model.Property1}"))),
                        new Paragraph(new Run(new Text("{model.SubModel.Property1}"))),
                        new Table(
                            new TableRow(
                                new TableCell(new Paragraph(new Run(new Text("{model.Items[i].Property1}"))))),
                            new TableRow(
                                new TableCell(new Paragraph(new Run(new Text("{i-related} related-content"))))))
                    ));

                wordDocument.Close();
                return memoryStream.ToArray();
            }
        }
    }
}
