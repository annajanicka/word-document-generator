using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;


[assembly: InternalsVisibleTo("DocumentGenerator.Tests")]
namespace DocumentGenerator
{
    internal class ModelParser : IModelParser
    {
        public object GetPropertyValue(object model, string propertyName)
        {
            if (propertyName.Contains("."))
            {
                var temp = propertyName.Split(new[] { '.' }, 2);
                var parent = GetPropertyValue(model, temp[0]);
                return GetPropertyValue(parent, temp[1]);
            }

            int? index = null;

            var regex = new Regex(@"\[.*\]");
            if (regex.IsMatch(propertyName))
            {
                var match = regex.Match(propertyName);
                index = int.Parse(match.Value.Replace("[", string.Empty).Replace("]", string.Empty));
                propertyName = propertyName.Replace(match.Value, string.Empty);
            }

            var property = model.GetType().GetProperty(propertyName);
            if (property == null)
            {
                throw new NullReferenceException($"{nameof(property)} is null for '{propertyName}'. Document template contains unknown model expression.");
            }

            if (property.GetValue(model, null) is IEnumerable<object> collection)
            {
                return index.HasValue ? collection?.ElementAt(index.Value) : collection;
            }

            return property.GetValue(model, null);
        }
    }
}
