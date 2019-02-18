namespace DocumentGenerator
{
    public interface IModelParser
    {
        object GetPropertyValue(object model, string propertyName);
    }
}
