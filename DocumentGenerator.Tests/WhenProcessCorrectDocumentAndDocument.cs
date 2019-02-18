using System;
using System.IO;
using System.Numerics;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Web;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using NUnit.Framework;
using Shouldly;

namespace DocumentGenerator.Tests
{
    [TestFixture]
    public class WhenProcessCorrectDocumentAndDocument
    {
        WordDocumentGenerator _formatter;
        FakeModelWithDiversePropertyTypes _model;

        [SetUp]
        public void Setup()
        {
            _formatter = new WordDocumentGenerator();
            _model = GetModel();
        }

        [Test]
        public void OpeningWordpressingDocumentAndReplacingAllModelIndicatorsReturnsCorrectDocument()
        {
            var assembly = Assembly.GetExecutingAssembly();
            using (var stream = assembly.GetManifestResourceStream("DocumentGenerator.Tests.form.docx"))
            using (MemoryStream mem = new MemoryStream())
            {
                stream.CopyTo(mem);
                var result = _formatter.GetDocument(_model, mem.ToArray());
                using (var memoryStream = new MemoryStream(result))
                using (var document = WordprocessingDocument.Open(memoryStream, true))
                using (var streamReader = new StreamReader(document.MainDocumentPart.GetStream()))
                {
                    var content = streamReader.ReadToEnd();
                    content.ShouldContain(_model.Property1);
                    content.ShouldContain(_model.Property2);
                    content.ShouldContain(_model.Property3.ToString());
                    content.ShouldContain(_model.Property4.ToString("F"));
                    content.ShouldContain(_model.Property5.ToString("G"));
                    content.ShouldContain(_model.Property6.ToString("D"));
                    content.ShouldContain(_model.Property7.ToString("hh"));
                    content.ShouldContain(HttpUtility.HtmlEncode(_model.Property8.ToString()));
                    content.ShouldContain(_model.Property9.ToString("N"));
                    content.ShouldContain(_model.SubModel.Property1);
                    content.ShouldContain(_model.SubModel.Property2);

                    foreach (var item in _model.Items)
                    {
                        content.ShouldContain(item.Property1);
                        content.ShouldContain(item.Property2);
                    }

                    var regex = new Regex("{model(.*?)}");
                    var placeholdersToReplace = regex.Matches(content);
                    placeholdersToReplace.Count.ShouldBe(0);
                }
            }
        }

        [Test]
        public void CreatingWordpressingDocumentAndReplacingAllModelIndicatorsReturnsCorrectDocument()
        {
            var result = _formatter.GetDocument(_model, GetDocument());
            using (var memoryStream = new MemoryStream(result))
            using (var document = WordprocessingDocument.Open(memoryStream, true))
            using (var streamReader = new StreamReader(document.MainDocumentPart.GetStream()))
            {
                var content = streamReader.ReadToEnd();
                content.ShouldContain(_model.Property1);
                content.ShouldContain(_model.Property2);
                content.ShouldContain(_model.Property3.ToString());
                content.ShouldContain(_model.Property4.ToString("F"));
                content.ShouldContain(_model.Property5.ToString("G"));
                content.ShouldContain(_model.Property6.ToString("D"));
                content.ShouldContain(_model.Property7.ToString("hh"));
                content.ShouldContain(HttpUtility.HtmlEncode(_model.Property8.ToString()));
                content.ShouldContain(_model.Property9.ToString("N"));
                content.ShouldContain(_model.SubModel.Property1);
                content.ShouldContain(_model.SubModel.Property2);

                foreach (var item in _model.Items)
                {
                    content.ShouldContain(item.Property1);
                    content.ShouldContain(item.Property2);
                }

                var regex = new Regex("{model(.*?)}");
                var placeholdersToReplace = regex.Matches(content);
                placeholdersToReplace.Count.ShouldBe(0);
                Regex.Matches(content, "related-content").Count.ShouldBe(_model.Items.Length);
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
                        new Paragraph(new Run(new Text("{model.Property2}"))),
                        new Paragraph(new Run(new Text("{model.Property3}"))),
                        new Paragraph(new Run(new Text("{model.Property4:F}"))),
                        new Paragraph(new Run(new Text("{model.Property5:G}"))),
                        new Paragraph(new Run(new Text("{model.Property6:D}"))),
                        new Paragraph(new Run(new Text("{model.Property7:hh}"))),
                        new Paragraph(new Run(new Text("{model.Property8}"))),
                        new Paragraph(new Run(new Text("{model.Property9:N}"))),
                        new Paragraph(new Run(new Text("{model.SubModel.Property1}"))),
                        new Paragraph(new Run(new Text("{model.SubModel.Property2}"))),
                        new Table(
                            new TableRow(
                                new TableCell(new Paragraph(new Run(new Text("{model.Items[i].Property1}")))),
                                new TableCell(new Paragraph(new Run(new Text("{model.Items[i].Property2}"))))),
                            new TableRow(
                                new TableCell(new Paragraph(new Run(new Text("{i-related} related-content"))))))
                    ));

                wordDocument.Close();
                return memoryStream.ToArray();
            }
        }

        private FakeModelWithDiversePropertyTypes GetModel()
        {
            return new FakeModelWithDiversePropertyTypes(
                Guid.NewGuid().ToString(),
                Guid.NewGuid().ToString(),
                123123,
                2.2M,
                DateTime.Now, 
                DateTimeOffset.Now, 
                TimeSpan.FromDays(123232),
                Vector2.UnitY, 
                Guid.NewGuid(),                
                new FakeModelWithDiversePropertyTypes.FakeSubModel(Guid.NewGuid().ToString(), Guid.NewGuid().ToString()),
                new[]
                {
                    new FakeModelWithDiversePropertyTypes.FakeItem(Guid.NewGuid().ToString(), Guid.NewGuid().ToString()),
                    new FakeModelWithDiversePropertyTypes.FakeItem(Guid.NewGuid().ToString(), Guid.NewGuid().ToString()),
                    new FakeModelWithDiversePropertyTypes.FakeItem(Guid.NewGuid().ToString(), Guid.NewGuid().ToString()),
                    new FakeModelWithDiversePropertyTypes.FakeItem(Guid.NewGuid().ToString(), Guid.NewGuid().ToString()),
                    new FakeModelWithDiversePropertyTypes.FakeItem(Guid.NewGuid().ToString(), Guid.NewGuid().ToString()),
                    new FakeModelWithDiversePropertyTypes.FakeItem(Guid.NewGuid().ToString(), Guid.NewGuid().ToString()),
                    new FakeModelWithDiversePropertyTypes.FakeItem(Guid.NewGuid().ToString(), Guid.NewGuid().ToString()),
                    new FakeModelWithDiversePropertyTypes.FakeItem(Guid.NewGuid().ToString(), Guid.NewGuid().ToString()),
                    new FakeModelWithDiversePropertyTypes.FakeItem(Guid.NewGuid().ToString(), Guid.NewGuid().ToString()),
                    new FakeModelWithDiversePropertyTypes.FakeItem(Guid.NewGuid().ToString(), Guid.NewGuid().ToString())
                });
        }
    }
}
