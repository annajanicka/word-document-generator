namespace DocumentGenerator
{
    public class ProcessingOptions : IProcessionOptions
    {
        public string ModelPrefix { get; }

        public ProcessingOptions(string modelPrefix)
        {
            ModelPrefix = modelPrefix;
        }
    }
}
