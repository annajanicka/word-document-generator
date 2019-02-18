using System;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using NUnit.Framework;

namespace DocumentGenerator.Tests
{
    [TestFixture]
    public class WhenDocumentContainsUnknowMappings
    {
        private WordDocumentGenerator _formatter;

        [SetUp]
        public void Setup()
        {
            _formatter = new WordDocumentGenerator();
        }

        [Test]
        public void WhenDocumentContainsMappingUnknownForModelThrowException()
        {
            var model = new FakeModel(
                "Value of property 1",
                "Value of property 2",
                new FakeModel.FakeSubModel("Value of property 1 of sub model", "Value of property 2 of sub model"),
                new []
                {
                    new FakeModel.FakeItem("Value of property 1 of fake item 1", "Value of property 2 of fake item 1")
                });

            var exception = Assert.Throws<NullReferenceException>(() => _formatter.GetDocument(model, GetDocument()));
            Assert.That(exception.Message, Is.EqualTo("property is null for 'NonExistingProperty'. Document template contains unknown model expression."));
        }

        private byte[] GetDocument()
        {
            using (MemoryStream memoryStream = new MemoryStream())
            using (var wordDocument = WordprocessingDocument.Create(memoryStream, WordprocessingDocumentType.Document))
            {
                MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();
                mainPart.Document = new Document(
                    new Body(
                        new Paragraph(new Run(new Text("{model.NonExistingProperty}")))
                    ));

                wordDocument.Close();
                return memoryStream.ToArray();
            }
        }
    }
}
