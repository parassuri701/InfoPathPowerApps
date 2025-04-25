using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocumentGeneratorService
{
    public interface IGenerateDocument
    {
        public void GenerateDocumentation(string infoPathFilePath, string outputDocPath);
        public string Output { get; set; }
    }
}
