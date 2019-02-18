using System;
using System.IO;
using System.Numerics;
using System.Text.RegularExpressions;
using System.Web;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using NUnit.Framework;
using Shouldly;

namespace DocumentGenerator.Tests
{
    public class WhenUseCustomModelPrefix
    {
        private WordDocumentGenerator _formatter;
        [SetUp]
        public void Setup()
        {
            _formatter = new WordDocumentGenerator();
        }

        [Test]
        public void CreatingWordpressingDocumentAndReplacingAllModelIndicatorsReturnsCorrectDocument()
        {
            var model = GetModel();
            var result = _formatter.GetDocument(model, GetDocument(), new ProcessingOptions("customModelPrefix"));
            using (var memoryStream = new MemoryStream(result))
            using (var document = WordprocessingDocument.Open(memoryStream, true))
            using (var streamReader = new StreamReader(document.MainDocumentPart.GetStream()))
            {
                var content = streamReader.ReadToEnd();
                content.ShouldContain(model.Property1);
                content.ShouldContain(model.Property2);
                content.ShouldContain(model.Property3.ToString());
                content.ShouldContain(model.Property4.ToString("F"));
                content.ShouldContain(model.Property5.ToString("G"));
                content.ShouldContain(model.Property6.ToString("D"));
                content.ShouldContain(model.Property7.ToString("hh"));
                content.ShouldContain(HttpUtility.HtmlEncode(model.Property8.ToString()));
                content.ShouldContain(model.Property9.ToString("N"));
                content.ShouldContain(model.SubModel.Property1);
                content.ShouldContain(model.SubModel.Property2);

                foreach (var item in model.Items)
                {
                    content.ShouldContain(item.Property1);
                    content.ShouldContain(item.Property2);
                }

                var regex = new Regex("{customModelPrefix(.*?)}");
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
                        new Paragraph(new Run(new Text("{customModelPrefix.Property1}"))),
                        new Paragraph(new Run(new Text("{customModelPrefix.Property2}"))),
                        new Paragraph(new Run(new Text("{customModelPrefix.Property3}"))),
                        new Paragraph(new Run(new Text("{customModelPrefix.Property4:F}"))),
                        new Paragraph(new Run(new Text("{customModelPrefix.Property5:G}"))),
                        new Paragraph(new Run(new Text("{customModelPrefix.Property6:D}"))),
                        new Paragraph(new Run(new Text("{customModelPrefix.Property7:hh}"))),
                        new Paragraph(new Run(new Text("{customModelPrefix.Property8}"))),
                        new Paragraph(new Run(new Text("{customModelPrefix.Property9:N}"))),
                        new Paragraph(new Run(new Text("{customModelPrefix.SubModel.Property1}"))),
                        new Paragraph(new Run(new Text("{customModelPrefix.SubModel.Property2}"))),
                        new Table(
                            new TableRow(
                                new TableCell(new Paragraph(new Run(new Text("{customModelPrefix.Items[i].Property1}")))),
                                new TableCell(new Paragraph(new Run(new Text("{customModelPrefix.Items[i].Property2}"))))),
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
