namespace DocumentGenerator
{
    public interface IWordDocumentGenerator
    {
        byte[] GetDocument(object model, byte[] fileBytes);
    }
}
