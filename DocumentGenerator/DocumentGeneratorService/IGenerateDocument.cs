using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocumentGeneratorService
{
    // This interface defines the contract for generating documentation based on either infoPath or Nintex forms.
    public interface IGenerateDocument
    {
        public void GenerateDocumentation(string infoPathFilePath, string outputDocPath);
        public string Output { get; set; }
    }
}
