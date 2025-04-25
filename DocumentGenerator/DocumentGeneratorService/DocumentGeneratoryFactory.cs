using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocumentGeneratorService
{
    public static class DocumentGeneratoryFactory
    {
       // public DocumentGeneratoryFactory() { }

        public static IGenerateDocument GetDocumentGeneratorByType(string filename)
        {
            if (string.IsNullOrWhiteSpace(filename))
            {
                throw new ArgumentException("Filename cannot be null or empty.", nameof(filename));
            }

            string extension = Path.GetExtension(filename)?.ToLower();

            switch (extension)
            {
                case ".nfp":
                    return new NintexDocumentationGenerator();
                case ".xsn":
                    return new InfoPathGenerateDocument();
                default:
                    throw new NotSupportedException($"File extension '{extension}' is not supported.");
            }
        }
    }
}
