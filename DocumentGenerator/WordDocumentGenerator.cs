using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXmlPowerTools;

namespace DocumentGenerator
{
    public class WordDocumentGenerator : IWordDocumentGenerator
    {
        private const string RELATED_ROWS_INDICATOR = "{i-related}";
        private const string RUN_ELEMENT_NAME = "r";
        private const string TABLE_ROW_ELEMENT_NAME = "tr";
        private readonly IModelParser _modelParser;

        public WordDocumentGenerator() : this(new ModelParser())
        {
        }

        public WordDocumentGenerator(IModelParser modelParser)
        {
            _modelParser = modelParser;
        }

        public byte[] GetDocument(object model, byte[] fileBytes)
        {
            return GetDocument(model, fileBytes, CultureInfo.CurrentCulture, new ProcessingOptions("model"));
        }

        public byte[] GetDocument(object model, byte[] fileBytes, IProcessionOptions options)
        {
            return GetDocument(model, fileBytes, CultureInfo.CurrentCulture, options);
        }

        public byte[] GetDocument(object model, byte[] fileBytes, CultureInfo culture, IProcessionOptions options)
        {
            if (model == null)
            {
                throw new ArgumentException($"{nameof(model)} is required", nameof(model));
            }

            if (fileBytes == null)
            {
                throw new ArgumentException($"{nameof(fileBytes)} is required", nameof(fileBytes));
            }

            using (var stream = new MemoryStream())
            {
                stream.Write(fileBytes, 0 , fileBytes.Length);
                using (var document = WordprocessingDocument.Open(stream, true))
                {
                    SimplifyMarkup(document);
                    PopulateTemplateRows(model, document, options);
                    UpdateContent(model, document, culture, options);
                    document.Close();
                    return stream.ToArray();
                }                
            }
        }

        private void UpdateContent(object model, WordprocessingDocument document, CultureInfo culture, IProcessionOptions options)
        {
            var modelRegex = new Regex($"(?<={options.ModelPrefix}.).+?(?=}})");

            var elementsToUpdate = document.MainDocumentPart.Document.Descendants()
                .Where(d => modelRegex.IsMatch(d.InnerText) && d.LocalName == RUN_ELEMENT_NAME)
                .ToList();
            foreach (var element in elementsToUpdate)
            {
                var expressionValue = modelRegex.Match(element.InnerText).Value;
                var newValue = element.InnerText
                    .Replace(expressionValue, GetReplacement(model, expressionValue, culture))
                    .Replace($"{{{options.ModelPrefix}.", string.Empty)
                    .Replace("}", string.Empty);
                element.RemoveAllChildren<Text>();
                element.AppendChild(new Text(newValue));
            }
        }

        private void PopulateTemplateRows(object model, WordprocessingDocument document, IProcessionOptions options)
        {
            var templateRows = document.MainDocumentPart.Document.Descendants()
                .Where(d => d.InnerText.Contains("[i]") && d.LocalName == TABLE_ROW_ELEMENT_NAME).ToList();
            var regex = new Regex($@"(?<={options.ModelPrefix}.).+?(?=\[)");
            foreach (var row in templateRows)
            {
                var propertyName = regex.Match(row.InnerText).Value;
                var elementsLength = (_modelParser.GetPropertyValue(model, propertyName) as IEnumerable<object>)?.Count();

                var relatedRowsForTemplateRow = row.Parent.ChildElements
                    .Where(c => c.InnerXml.Contains(RELATED_ROWS_INDICATOR) && c.LocalName == TABLE_ROW_ELEMENT_NAME)
                    .ToList();

                var lastRow = row;
                for (var i = 0; i < elementsLength; i++)
                {
                    var newRow = row.CloneNode(true);
                    newRow.InnerXml = newRow.InnerXml.Replace("[i]", $"[{i}]").Replace("{i}", (i + 1).ToString());
                    lastRow.InsertAfterSelf(newRow);
                    lastRow = newRow;
                    foreach (var relatedRow in relatedRowsForTemplateRow)
                    {
                        var node = relatedRow.CloneNode(true);
                        node.InnerXml = node.InnerXml.Replace(RELATED_ROWS_INDICATOR, string.Empty);
                        lastRow.InsertAfterSelf(node);
                        lastRow = node;
                    }
                }

                RemoveElements(relatedRowsForTemplateRow);
            }

            RemoveElements(templateRows);
        }

        private void RemoveElements(IEnumerable<OpenXmlElement> elements)
        {
            foreach (var element in elements)
            {
                element.Remove();
            }
        }

        private void SimplifyMarkup(WordprocessingDocument originalDocument)
        {
            SimplifyMarkupSettings settings = new SimplifyMarkupSettings
            {
                RemoveComments = true,
                RemoveContentControls = true,
                RemoveEndAndFootNotes = true,
                RemoveFieldCodes = false,
                RemoveLastRenderedPageBreak = true,
                RemovePermissions = true,
                RemoveProof = true,
                RemoveRsidInfo = true,
                RemoveSmartTags = true,
                RemoveSoftHyphens = true,
                ReplaceTabsWithSpaces = true,
                RemoveBookmarks = true,
                RemoveGoBackBookmark = true,
                RemoveHyperlinks = true,
               
            };
            MarkupSimplifier.SimplifyMarkup(originalDocument, settings);
        }

        private string GetReplacement(object model, string expressionValue, CultureInfo culture)
        {
            var splitByFormat = expressionValue.Split(":", 2);
            expressionValue = splitByFormat[0];
            var format = splitByFormat.Length > 1 ? splitByFormat[1] : string.Empty;
            var propertyValue = _modelParser.GetPropertyValue(model, expressionValue);
            if (propertyValue == null)
            {
                return string.Empty;
            }

            if (!string.IsNullOrEmpty(format) && propertyValue is IFormattable)
            {
                propertyValue = string.Format(culture, $"{{0:{format}}}", propertyValue);
            }

            return Convert.ToString(propertyValue);
        }
    }
}
