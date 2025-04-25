using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Data;
using System.Diagnostics;
using System.IO.Compression;
using System.Reflection;
using System.Reflection.Metadata;
using System.Reflection.PortableExecutable;
using System.Text;
using System.Xml.Linq;
using Color = DocumentFormat.OpenXml.Wordprocessing.Color;
using Document = DocumentFormat.OpenXml.Wordprocessing.Document;
using Text = DocumentFormat.OpenXml.Wordprocessing.Text;

namespace DocumentGeneratorService
{
    public class InfoPathGenerateDocument : IGenerateDocument
    {
        public string Output { get; set; }
        // Main class properties to store form information
        private List<Control> Controls { get; set; } = new List<Control>();
        private List<DataSource> DataSources { get; set; } = new List<DataSource>();
        private List<Rule> Rules { get; set; } = new List<Rule>();
        private List<Script> Scripts { get; set; } = new List<Script>();
        private List<Workflow> Workflows { get; set; } = new List<Workflow>();
        private Dictionary<string, List<string>> Dependencies { get; set; } = new Dictionary<string, List<string>>();
        private List<Dependency> DependencyList { get; set; } = new List<Dependency>();
        private List<ImageFile> ImageFiles { get; set; } = new List<ImageFile>();

        // Method to generate documentation
        public void GenerateDocumentation(string infoPathFilePath, string outputDocPath)
        {
            // First extract the XSN contents (XSN is a CAB file)         
            // string tempFolder = Path.Combine(AppContext.BaseDirectory, "stagingfiles_" + Guid.NewGuid().ToString("N"));
            string tempFolder = Path.Combine(Path.GetDirectoryName(outputDocPath), Path.GetFileNameWithoutExtension(outputDocPath));
            Directory.CreateDirectory(tempFolder);

            try
            {
                // Extract .xsn file which is basically a ZIP/CAB file
                ExtractXsnContents(infoPathFilePath, tempFolder);

                // Parse form definition
                ParseFormDefinition(tempFolder, outputDocPath);

                // Generate Word document
                CreateWordDocument(outputDocPath);
            }
            finally
            {

            }
        }

        private void ExtractXsnContents(string xsnFilePath, string targetFilePath)
        {
            Console.WriteLine("Extracting InfoPath form contents...");

            try
            {
                // Define the expand.exe command
                Process expandProcess = new Process
                {
                    StartInfo = new ProcessStartInfo
                    {
                        /* Path to expand.exe - usually it is resided in C:\Windows\System32\expand.exe
                         or C:\Windows\SysWOW64\expand.exe  */
                        FileName = "expand.exe",
                        Arguments = $"\"{xsnFilePath}\" -F:* \"{targetFilePath}\"", // Command arguments
                        RedirectStandardOutput = true,
                        RedirectStandardError = true,
                        UseShellExecute = false,
                        CreateNoWindow = true
                    }
                };

                // Start the process
                expandProcess.Start();

                // Read and display the output
                string output = expandProcess.StandardOutput.ReadToEnd();
                string error = expandProcess.StandardError.ReadToEnd();
                this.Output += "Extracting " + xsnFilePath + " ";
                this.Output += output;
                expandProcess.WaitForExit();

                // Check for success or errors
                if (expandProcess.ExitCode == 0)
                {
                    Console.WriteLine("Extraction completed successfully:");
                    Console.WriteLine(output);
                }
                else
                {
                    Console.WriteLine("An error occurred during extraction:");
                    Console.WriteLine(error);
                }

                Console.WriteLine("Form extracted successfully.");
            }
            catch (Exception ex)
            {
                throw new Exception("Failed to extract InfoPath form: " + ex.Message, ex);
            }
        }

        private void ParseFormDefinition(string extractedFolder, string outputPath)
        {
            Console.WriteLine("Parsing form definition...");

            // Look for key files
            string manifestFile = Path.Combine(extractedFolder, "manifest.xsf");
            if (!File.Exists(manifestFile))
            {
                throw new FileNotFoundException("Manifest file (manifest.xsf) not found in the InfoPath form package.");
            }

            // Parse the manifest file
            XDocument manifest = XDocument.Load(manifestFile);
            XNamespace xsfNs = "http://schemas.microsoft.com/office/infopath/2003/solutionDefinition";

            // Get basic form information
            var formName = manifest.Descendants(xsfNs + "title").FirstOrDefault()?.Value ?? "Unnamed Form";
            Console.WriteLine($"Form name: {formName}");

            // Parse data sources
            ParseDataSources(manifest, xsfNs);

            // Parse form files to find controls

            //TODO: Add functionality to read controls if necessary from other view files (view2 etc)
            string defaultViewFile = Path.Combine(extractedFolder, "view1.xsl");
            if (File.Exists(defaultViewFile))
            {
                ParseViewFile(defaultViewFile);
            }

            // Look for script files
            foreach (var scriptFile in Directory.GetFiles(extractedFolder, "*.js"))
            {
                ParseScriptFile(scriptFile);
            }

            foreach (var dllFile in Directory.GetFiles(extractedFolder, "*.dll"))
            {
                ParseDllFile(dllFile);
            }
            // Parse rules and validations from manifest
            ParseRulesAndValidations(manifest, xsfNs);
            ParseSchemaForValidations(extractedFolder, xsfNs);
            // Try to find workflow information
            ParseWorkflows(manifest, xsfNs);

            ParseDependencies(manifest, xsfNs, extractedFolder);
            ParseImages(manifest, xsfNs);

            //CopyFilesIntoReferenceFolder(outputPath, extractedFolder);
            Console.WriteLine("Parsing completed.");
        }

        private void CopyFilesIntoReferenceFolder(string outputDirPath, string extractedFilesPath)
        {

            string outputFilesPath = outputDirPath.Split(".")[0];
            if (ImageFiles.Count > 0 || DependencyList.Count > 0)
            {
                if (!Directory.Exists(outputFilesPath))
                {
                    Directory.CreateDirectory(outputFilesPath);
                }

                //foreach (var image in ImageFiles)
                //{

                //    CopyFile(outputFilesPath, extractedFilesPath, image.FileName);
                //}

                //foreach (var dependencyFile in DependencyList)
                //{
                //    CopyFile(outputFilesPath, extractedFilesPath, dependencyFile.Name);
                //}
            }
        }
        private void CopyFile(string outputFilesPath, string extractedFilesPath, string fileName)
        {
            string imageFilePath = Path.Combine(extractedFilesPath, fileName);
            string outputImagePath = Path.Combine(outputFilesPath, fileName);
            if (File.Exists(imageFilePath))
            {
                File.Copy(imageFilePath, outputImagePath, true);
            }
        }


        private void ParseSchemaForValidations(string extractedFolder, XNamespace xsfNs)
        {
            Console.WriteLine($"Beginning to parse schema for validation rules with {Rules.Count} validations .");

            List<string> schemaFiles = Directory.GetFiles(extractedFolder).Where(x => x.EndsWith(".xsd")).ToList();
            foreach (var schemaFile in schemaFiles)
            {
                if (!File.Exists(schemaFile))
                {
                    throw new FileNotFoundException("Schema file (schema.xsd) not found in the InfoPath form package.");
                }
                XDocument schema = XDocument.Load(schemaFile);
                var descendentNodes = schema.DescendantNodes().Distinct().ToList();

                foreach (var node in descendentNodes)
                {
                    if (node.NodeType.ToString() == "Element")
                    {
                        var elementNode = (XElement)node;
                        var type = elementNode.Attribute("type")?.Value ?? string.Empty;
                        if (type.Contains("required", StringComparison.OrdinalIgnoreCase))
                        {
                            Rules.Add(new Rule
                            {
                                Name = elementNode.Attribute("name")?.Value ?? string.Empty,
                                Type = "Required",
                                Expression = type,
                                Message = type,
                                Field = elementNode.Attribute("name")?.Value ?? string.Empty
                            });

                        }
                    }

                }
            }
            Console.WriteLine($"Total of {Rules.Count} validations found.");

        }
        private void ParseDependencies(XDocument manifest, XNamespace xsfNs, string extractedFolder)
        {
            Console.WriteLine("Parsing dependencies from xsf file...");
            // Parse for explicit DLL references via <xsf:script>
            var scriptElements = manifest.Descendants(xsfNs + "script")
                .Where(s => s.Attribute("name") != null &&
                            s.Attribute("name").Value.EndsWith(".dll", StringComparison.OrdinalIgnoreCase));

            foreach (var script in scriptElements)
            {
                DependencyList.Add(new Dependency
                {
                    Name = script.Attribute("name")?.Value,
                    Project = script.Attribute("project")?.Value ?? "Unknown",
                    Language = script.Attribute("language")?.Value ?? "Unknown",
                    ScriptLanguageVersion = script.Attribute("scriptLanguageVersion")?.Value ?? "Unknown"
                });
            }

            // Parse for implicit DLL references in <xsf:file> with rootAssembly
            var fileElements = manifest.Descendants(xsfNs + "file")
                .Where(f => f.Attribute("name") != null && f.Elements(xsfNs + "fileProperties")
                    .Any(p => p.Elements(xsfNs + "property")
                        .Any(prop => prop.Attribute("name")?.Value == "fileType" &&
                                     prop.Attribute("value")?.Value == "rootAssembly")));

            foreach (var file in fileElements)
            {
                DependencyList.Add(new Dependency
                {
                    Name = file.Attribute("name")?.Value,
                    Project = "Implicit DLL",
                    Language = "Managed Code",
                    ScriptLanguageVersion = "N/A"
                });
            }

            Console.WriteLine("Scanning for JavaScript and VBScript files...");

            if (string.IsNullOrWhiteSpace(extractedFolder) || !Directory.Exists(extractedFolder))
            {
                Console.WriteLine("Invalid extracted folder path for scanning scripts.");
                return;
            }

            var jsFiles = Directory.GetFiles(extractedFolder, "*.js");
            var vbsFiles = Directory.GetFiles(extractedFolder, "*.vbs");

            foreach (var js in jsFiles)
            {
                DependencyList.Add(new Dependency
                {
                    Name = Path.GetFileName(js),
                    Project = "Client Script",
                    Language = "JavaScript",
                    ScriptLanguageVersion = "N/A"
                });
            }

            foreach (var vbs in vbsFiles)
            {
                DependencyList.Add(new Dependency
                {
                    Name = Path.GetFileName(vbs),
                    Project = "Client Script",
                    Language = "VBScript",
                    ScriptLanguageVersion = "N/A"
                });
            }

            Console.WriteLine($"Total dependencies found: {DependencyList.Count}");
        }

        private void ParseDataSources(XDocument manifest, XNamespace xsfNs)
        {
            Console.WriteLine("Parsing data sources...");
            XNamespace xsfNss = "http://schemas.microsoft.com/office/infopath/2006/solutionDefinition/extensions";

            var dataSources = manifest.Descendants(xsfNs + "dataObject");

            foreach (var dataObject in dataSources)
            {
                var name = dataObject.Attribute("name")?.Value ?? "Unnamed Data Source";
                var type = "";
                var desc = "";
                // Identify Adapter Type
                var adapter = dataObject.Descendants().Skip(1).FirstOrDefault();
                var adapterType = "";
                if (adapter != null)
                {
                    adapterType = adapter.Name.LocalName;
                    Console.WriteLine($"  - Adapter Type: {adapter.Name.LocalName}");

                    if (adapterType == "SharePointListAdapter")
                    { type = "SharePoint List/Library"; desc = "Get/update SharePoint data"; }
                    else if (adapterType == "adoAdapter")
                    { type = "SQL Database"; desc = "Query SQL Server"; }
                    else if (adapterType == "httpAdapter")
                    { type = "REST API (HTTP GET/POST)"; desc = "Retrieve JSON/XML data from REST API"; }
                    else if (adapterType == "fileAdapter")
                    { type = "XML File Connection"; desc = "Load XML from a file or URL"; }
                    else if (adapterType == "webServiceAdapter")
                    { type = "Web Service (SOAP)"; desc = "Query/submit data via SOAP"; }
                    else if (adapterType == "emailAdapter")
                    { type = "Email Submission"; desc = "Submit form data via email"; }
                    else if (adapterType == "udcAdapter")
                    { type = "Universal Data Connection (UDC)"; desc = "External UDC connection (SQL, SharePoint)"; }
                    else if (name == "Main")
                    { type = "Main"; desc = "Primary form data storage"; }
                    else
                    { type = "Unknown"; desc = "Unknown data source"; }

                    // Get key attributes of each adapter
                    foreach (var attr in adapter.Attributes())
                    {
                        Console.WriteLine($"    - {attr.Name}: {attr.Value}");
                    }

                    DataSources.Add(new DataSource
                    {
                        Name = name,
                        Type = adapterType,
                        Description = desc
                    });
                }

                Console.WriteLine(new string('-', 50)); // Separator
            }

            dataSources = manifest.Descendants(xsfNss + "dataConnections").Descendants();

            foreach (var dataSource in dataSources)
            {
                var adapterType = "";
                var sourceName = "";
                var sourceDesc = "";
                if (dataSource != null)
                {
                    adapterType = dataSource.Name.LocalName;

                    if (adapterType == "emailAdapterExtension")
                    {
                        sourceName = "Email Submission Extension";
                        sourceDesc = "EmailAttachmentType (XML)";
                    }
                    else if (adapterType == "sharepointListAdapterExtended")
                    {
                        sourceName = dataSource.Attributes("name").FirstOrDefault()?.Value ?? "Unknown name";
                        sourceDesc = "SharePoint Universal Data Connection(.udcx)";
                    }
                    else
                    {
                        sourceName = dataSource.Attributes("name").FirstOrDefault()?.Value ?? "Unknown name";
                        sourceDesc = "Unknown data source";
                    }

                    DataSources.Add(new DataSource
                    {
                        Name = sourceName,
                        Type = adapterType,
                        Description = sourceDesc
                    });
                }
            }

            Console.WriteLine($"Found {DataSources.Count} data sources.");
        }

        private void ParseViewFile(string viewFilePath)
        {
            Console.WriteLine($"Parsing view file: {Path.GetFileName(viewFilePath)}");

            try
            {
                XDocument viewDoc = XDocument.Load(viewFilePath);

                // The XSL file contains the form UI definition
                // Look for controls - typically in xsl:template elements
                var templates = viewDoc.Descendants().Where(e => e.Name.LocalName == "template");

                foreach (var template in templates)
                {
                    ParseControlsFromTemplate(template);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Warning: Could not parse view file {viewFilePath}: {ex.Message}");
            }
        }

        private void ParseControlsFromTemplate(XElement template)
        {
            // Look for common InfoPath control patterns
            // This is a simplified approach - real implementation would be more comprehensive

            // Look for spans with xd:binding attributes (usually bound controls)
            var boundControls = template.Descendants()
                .Where(e => e.Attributes().Any(a => a.Name.LocalName == "binding" || a.Name.LocalName == "xctname"));

            foreach (var control in boundControls)
            {
                string controlType = DetermineControlType(control);
                string controlName = control.Attribute(XName.Get("xctname", "http://schemas.microsoft.com/office/infopath/2003"))?.Value
                    ?? control.Attribute(XName.Get("binding", "http://schemas.microsoft.com/office/infopath/2003"))?.Value
                    ?? "Unnamed Control";
                string controlOptions = string.Empty;

                if (controlType.Equals("Drop-down List", StringComparison.OrdinalIgnoreCase))
                {
                    var descendentOptions = control.DescendantNodes();
                    foreach (var node in descendentOptions)
                    {
                        if (node.NodeType.ToString() == "Element")
                        {
                            var ElementNode = (XElement)node;
                            if (ElementNode.Name.LocalName == "option")
                            {
                                var option = ElementNode.Attribute("value")?.Value ?? string.Empty;
                                if (!string.IsNullOrEmpty(option))
                                {
                                    if (controlOptions.Length > 0)
                                    {
                                        controlOptions += $", {option}";
                                    }
                                    else
                                    {
                                        controlOptions += option;
                                    }
                                }

                            }
                        }
                    }

                }
                else if (controlName.Equals("CheckBox", StringComparison.OrdinalIgnoreCase))
                {
                    string title = control.Attribute("title")?.Value ?? " TODO: Still Pending ";
                    controlOptions = title;
                }
                else if (controlName.Equals("OptionButton", StringComparison.OrdinalIgnoreCase))
                {
                    if (control.NextNode?.NodeType == System.Xml.XmlNodeType.Text)
                    {
                        var nextNode = (XText)control.NextNode;
                        controlOptions = nextNode.Value;
                    }
                    else
                    {
                        controlOptions = "TODO: Still Pending";
                    }
                }
                else if (controlName.Equals("Section", StringComparison.OrdinalIgnoreCase))
                {
                    string controlClass = control.Attribute("class")?.Value ?? string.Empty;
                    controlOptions = controlClass;
                }
                else if (controlType == "Unknown Control" && !controlName.Contains("DTPicker"))
                {
                    controlOptions = "TODO: Still Pending";
                }


                Controls.Add(new Control
                {
                    Name = controlName,
                    Type = controlType,
                    Binding = control.Attribute(XName.Get("binding", "http://schemas.microsoft.com/office/infopath/2003"))?.Value ?? "Not bound",
                    Details = controlOptions
                });

            }
        }

        private string DetermineControlType(XElement control)
        {
            // Attempt to determine control type based on attributes or parent elements
            // This is a simplified approach

            if (control.Name.LocalName == "input")
            {
                var type = control.Attribute("type")?.Value;
                return !string.IsNullOrEmpty(type) ? type : "Text Box";
            }
            else if (control.Name.LocalName == "select")
            {
                return "Drop-down List";
            }
            else if (control.Name.LocalName == "button")
            {
                return "Button";
            }
            else if (control.Name.LocalName == "span" && control.Attributes().Any(a => a.Name.LocalName == "xctname"))
            {
                var xctname = control.Attribute(XName.Get("xctname", "http://schemas.microsoft.com/office/infopath/2003"))?.Value ?? "";

                if (xctname.Contains("CheckBox"))
                {
                    return "Check Box";
                }
                else if (xctname.Contains("DatePicker"))
                {
                    return "Date Picker";
                }
                else if (xctname.Contains("ListBox"))
                {
                    return "List Box";
                }
                else
                {
                    return xctname;
                }
            }
            else if ((control.Name.LocalName == "tbody" || control.Name.LocalName == "div")
                && control.Attributes().Any(a => a.Name.LocalName == "xctname"))
            {
                var xctname = control.Attribute(XName.Get("xctname", "http://schemas.microsoft.com/office/infopath/2003"))?.Value ?? "";
                return xctname;
            }

            return "Unknown Control";
        }
        private void ParseDllFile(string dllFilePath)
        {
            Console.WriteLine($"Parsing Dll file: {Path.GetFileName(dllFilePath)}");
            try
            {
                Assembly assembly = Assembly.LoadFrom(dllFilePath);


                using (var stream = new FileStream(dllFilePath, FileMode.Open, FileAccess.Read))
                {
                    using (var peReader = new PEReader(stream))
                    {
                        if (!peReader.HasMetadata)
                        {
                            Console.WriteLine("The file does not contain metadata.");
                            return;
                        }

                        var metadataReader = peReader.GetMetadataReader();
                        foreach (var methodDefinitionHandle in metadataReader.MethodDefinitions)
                        {
                            var methodDefinition = metadataReader.GetMethodDefinition(methodDefinitionHandle);
                            var methodName = metadataReader.GetString(methodDefinition.Name);
                            Console.WriteLine($"Method: {methodName}");
                            // Retrieve parameter list
                            var parameterHandles = methodDefinition.GetParameters();
                            StringBuilder paramList = new StringBuilder();
                            foreach (var parameterHandle in parameterHandles)
                            {
                                var parameter = metadataReader.GetParameter(parameterHandle);
                                var parameterName = metadataReader.GetString(parameter.Name);

                                // Optional: Get parameter type information (if available)
                                paramList.Append(parameterName + ", ");
                            }

                            // Remove trailing ", " from parameter list
                            if (paramList.Length > 2)
                            {
                                paramList.Length -= 2;
                            }

                            Scripts.Add(new Script
                            {
                                Name = methodName,
                                Parameters = paramList.ToString(),
                                FilePath = Path.GetFileName(dllFilePath)
                            });
                        }
                    }
                    stream.Close();
                    stream.Dispose();
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine($"Warning: Could not parse Dll file {dllFilePath}: {ex.Message}");
            }
        }
        private void ParseScriptFile(string scriptFilePath)
        {
            Console.WriteLine($"Parsing script file: {Path.GetFileName(scriptFilePath)}");

            try
            {
                string scriptContent = File.ReadAllText(scriptFilePath);

                // Look for function definitions
                var functionMatches = System.Text.RegularExpressions.Regex.Matches(
                    scriptContent,
                    @"function\s+([A-Za-z0-9_:]+)\s*\(([^)]*)\)"
                );

                foreach (System.Text.RegularExpressions.Match match in functionMatches)
                {
                    string functionName = match.Groups[1].Value;
                    string parameters = match.Groups[2].Value;

                    Scripts.Add(new Script
                    {
                        Name = functionName,
                        Parameters = parameters,
                        FilePath = Path.GetFileName(scriptFilePath)
                    });
                }

                Console.WriteLine($"Found {functionMatches.Count} script functions.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Warning: Could not parse script file {scriptFilePath}: {ex.Message}");
            }
        }

        private void ParseRulesAndValidations(XDocument manifest, XNamespace xsfNs)
        {
            Console.WriteLine("Parsing rules and validations...");

            var descendentNodes = manifest.DescendantNodes().Distinct().ToList();

            foreach (var node in descendentNodes)
            {
                if (node.NodeType.ToString() == "Element")
                {
                    var elementNode = (XElement)node;

                    if (elementNode.Name.LocalName == "errorCondition")
                    {
                        string validationName = elementNode.Attribute("name")?.Value ?? "Unnamed Validation";
                        string expression = elementNode.Attribute("expression")?.Value ?? "";
                        string message = "No message node found";
                        string field = elementNode.Attribute("expressionContext")?.Value ?? "";

                        var messageNode = elementNode.Descendants().Where(e => e.Name.LocalName == "errorMessage")?.FirstOrDefault();
                        message = messageNode.Attribute("shortMessage")?.ToString() ?? "No shortMessage present";
                        if (String.IsNullOrEmpty(field) || field == ".")
                        {
                            var matchValue = elementNode?.Attribute("match")?.Value ?? "";
                            field = matchValue?.Split('/')?.LastOrDefault() ?? "No field specified";
                        }

                        Rules.Add(new Rule
                        {
                            Name = validationName,
                            Type = "Validation",
                            Expression = expression,
                            Message = message,
                            Field = field
                        });
                    }
                }
                else if (node.NodeType.ToString() == "Comment" || node.NodeType.ToString() == "Text")
                {
                    //Do nothing for now
                }
                else if (node.NodeType.ToString() == "ProcessingInstruction")
                {
                    Console.WriteLine($"{node.NodeType.ToString()} {node.ToString()}");
                }
                else
                {
                    Console.WriteLine($"Unknown Type of node {node.NodeType.ToString()}");
                }
            }

            Console.WriteLine($"Found {Rules.Count} rules and validations.");
        }

        private void ParseWorkflows(XDocument manifest, XNamespace xsfNs)
        {
            Console.WriteLine("Looking for workflow information...");

            // Look for workflow elements - this varies by InfoPath version and SharePoint integration
            var workflowElements = manifest.Descendants().Where(e =>
                e.Name.LocalName.Contains("workflow") ||
                e.Name.LocalName.Contains("Workflow"));

            foreach (var workflowElement in workflowElements)
            {
                string workflowName = workflowElement.Attribute("name")?.Value ??
                                      workflowElement.Attribute("displayName")?.Value ??
                                      "Unnamed Workflow";

                Workflows.Add(new Workflow
                {
                    Name = workflowName,
                    Description = GetWorkflowDescription(workflowElement)
                });
            }

            // If we didn't find any specific workflow elements, check for SharePoint Submit options
            if (Workflows.Count == 0)
            {
                var submitOptions = manifest.Descendants(xsfNs + "submit");
                if (submitOptions.Any())
                {
                    Workflows.Add(new Workflow
                    {
                        Name = "Form Submit Process",
                        Description = "Standard form submission process"
                    });
                }
            }

            Console.WriteLine($"Found {Workflows.Count} workflows.");
        }
        private void ParseImages(XDocument manifest, XNamespace xsfNs)
        {
            Console.WriteLine("Parsing images...");

            var imageElements = manifest.Descendants().Where(e =>
                e.Name.LocalName.Equals("file", StringComparison.InvariantCultureIgnoreCase));

            foreach (var imageElement in imageElements)
            {
                var fileName = imageElement.Attribute("name")?.Value ?? "Unnamed Image";
                if (!fileName.Contains(".xsd")
                && !fileName.Contains(".xsl")
                && !fileName.Contains(".xml")
                && !fileName.Contains(".js")
                && !fileName.Contains(".dll")
                && !fileName.Contains(".vbs"))
                {
                    ImageFiles.Add(new ImageFile
                    {
                        FileName = fileName,
                    });
                }
            }

            Console.WriteLine($"Found {ImageFiles.Count} images.");
        }

        private string GetWorkflowDescription(XElement workflowElement)
        {
            // Try to extract a description from the workflow element
            var description = workflowElement.Elements().FirstOrDefault(e =>
                e.Name.LocalName.Contains("description") ||
                e.Name.LocalName.Contains("Description"));

            if (description != null)
            {
                return description.Value;
            }

            return "No description available";
        }

        private void CreateWordDocument(string outputPath)
        {
            Console.WriteLine("Generating Word documentation...");
            FileStream fileStream = new FileStream(outputPath, FileMode.Create, FileAccess.ReadWrite, FileShare.None);
            using (WordprocessingDocument wordDocument = WordprocessingDocument.Create(fileStream, WordprocessingDocumentType.Document))
            {
                // Add a main document part
                MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();

                // Create the document structure
                mainPart.Document = new Document();
                Body body = new Body();
                mainPart.Document.Append(body);

                // Add styles
                AddStyles(wordDocument);

                // Add title and document info
                AddDocumentTitle(body, "InfoPath Form Documentation");

                // Add table of contents
                AddTableOfContents(body);

                // Add sections
                int sectionNumber = 1;

                // Controls Section
                AddSection(body, $"{sectionNumber++}. Controls", "Heading1");
                AddControlsTable(body);
                body.Append(new Paragraph(new Run(new Break()))); // Add a line break

                // Rules Section
                AddSection(body, $"{sectionNumber++}. Validation Rules", "Heading1");
                AddRulesTable(body);
                body.Append(new Paragraph(new Run(new Break()))); // Add a line break

                // Data Sources Section
                AddSection(body, $"{sectionNumber++}. Data Sources", "Heading1");
                AddDataSourcesTable(body);
                body.Append(new Paragraph(new Run(new Break()))); // Add a line break

                // Scripts Section
                AddSection(body, $"{sectionNumber++}. Scripts and Business Logic", "Heading1");
                AddScriptsTable(body);
                body.Append(new Paragraph(new Run(new Break()))); // Add a line break

                // Workflows Section
                AddSection(body, $"{sectionNumber++}. Workflows", "Heading1");
                AddWorkflowsTable(body);
                body.Append(new Paragraph(new Run(new Break()))); // Add a line break

                // Dependencies Section
                AddSection(body, $"{sectionNumber++}. Dependencies", "Heading1");
                AddDependenciesTable(body);
                body.Append(new Paragraph(new Run(new Break()))); // Add a line break

                // Images Section
                AddSection(body, $"{sectionNumber++}. Images", "Heading1");
                AddImagesTable(body);

                // Add page numbers
                AddPageNumbers(wordDocument);

                // Update the table of contents
                UpdateTableOfContents(wordDocument);

                Console.WriteLine("Word document generation completed.");
            }
        }

        private void AddStyles(WordprocessingDocument wordDocument)
        {
            // Add a StylesPart to the document
            StyleDefinitionsPart stylesPart = wordDocument.MainDocumentPart.AddNewPart<StyleDefinitionsPart>();

            // Create Styles element
            Styles styles = new Styles();

            // Create Title style
            Style titleStyle = new Style() { Type = StyleValues.Paragraph, StyleId = "Title" };
            StyleName titleStyleName = new StyleName() { Val = "Title" };
            titleStyle.Append(titleStyleName);
            PrimaryStyle primaryStyle1 = new PrimaryStyle();
            titleStyle.Append(primaryStyle1);

            StyleRunProperties titleRunProperties = new StyleRunProperties();
            Bold titleBold = new Bold();
            Color titleColor = new Color() { Val = "2F5496" };
            FontSize titleFontSize = new FontSize() { Val = "32" };
            titleRunProperties.Append(titleBold);
            titleRunProperties.Append(titleColor);
            titleRunProperties.Append(titleFontSize);
            titleStyle.Append(titleRunProperties);

            // Create Heading1 style
            Style heading1Style = new Style() { Type = StyleValues.Paragraph, StyleId = "Heading1" };
            StyleName heading1StyleName = new StyleName() { Val = "Heading 1" };
            heading1Style.Append(heading1StyleName);
            PrimaryStyle primaryStyle2 = new PrimaryStyle();
            heading1Style.Append(primaryStyle2);

            StyleRunProperties heading1RunProperties = new StyleRunProperties();
            Bold heading1Bold = new Bold();
            Color heading1Color = new Color() { Val = "2F5496" };
            FontSize heading1FontSize = new FontSize() { Val = "28" };
            heading1RunProperties.Append(heading1Bold);
            heading1RunProperties.Append(heading1Color);
            heading1RunProperties.Append(heading1FontSize);
            heading1Style.Append(heading1RunProperties);

            // Create Heading2 style
            Style heading2Style = new Style() { Type = StyleValues.Paragraph, StyleId = "Heading2" };
            StyleName heading2StyleName = new StyleName() { Val = "Heading 2" };
            heading2Style.Append(heading2StyleName);
            PrimaryStyle primaryStyle3 = new PrimaryStyle();
            heading2Style.Append(primaryStyle3);

            StyleRunProperties heading2RunProperties = new StyleRunProperties();
            Bold heading2Bold = new Bold();
            Color heading2Color = new Color() { Val = "2F5496" };
            FontSize heading2FontSize = new FontSize() { Val = "24" };
            heading2RunProperties.Append(heading2Bold);
            heading2RunProperties.Append(heading2Color);
            heading2RunProperties.Append(heading2FontSize);
            heading2Style.Append(heading2RunProperties);

            // Create Normal style
            Style normalStyle = new Style() { Type = StyleValues.Paragraph, StyleId = "Normal", Default = true };
            StyleName normalStyleName = new StyleName() { Val = "Normal" };
            normalStyle.Append(normalStyleName);
            PrimaryStyle primaryStyle4 = new PrimaryStyle();
            normalStyle.Append(primaryStyle4);

            // Add styles to the Styles element
            styles.Append(titleStyle);
            styles.Append(heading1Style);
            styles.Append(heading2Style);
            styles.Append(normalStyle);

            // Add the Styles element to the StylesPart
            stylesPart.Styles = styles;
        }

        private void AddDocumentTitle(Body body, string title)
        {
            Paragraph titleParagraph = new Paragraph();
            ParagraphProperties titleProperties = new ParagraphProperties();
            ParagraphStyleId titleStyleId = new ParagraphStyleId() { Val = "Title" };
            titleProperties.Append(titleStyleId);
            titleParagraph.Append(titleProperties);

            Run titleRun = new Run();
            RunProperties titleRunProperties = new RunProperties();
            titleRun.Append(titleRunProperties);
            Text titleText = new Text(title);
            titleRun.Append(titleText);
            titleParagraph.Append(titleRun);

            body.Append(titleParagraph);

            // Add a page break
            Paragraph pageBreak = new Paragraph();
            Run pageBreakRun = new Run();
            Break break1 = new Break() { Type = BreakValues.Page };
            //pageBreakRun.Append(break1);
            // pageBreak.Append(pageBreakRun);
            body.Append(pageBreak);
        }

        private void AddTableOfContents(Body body)
        {
            AddSection(body, "Table of Contents", "Heading1");

            // Add a TOC field
            Paragraph tocParagraph = new Paragraph();
            Run tocRun = new Run();

            // TOC field
            //tocRun.Append(new SimpleField() { Instruction = "TOC \\o \"1-3\" \\h \\z \\u" });
            var sdtBlock = new SdtBlock();
            sdtBlock.InnerXml = GetTOC("", 16);

            tocParagraph.Append(tocRun);
            body.Append(sdtBlock);

            // Add a page break
            Paragraph pageBreak = new Paragraph();
            Run pageBreakRun = new Run();
            Break break1 = new Break() { Type = BreakValues.Page };
            pageBreakRun.Append(break1);
            pageBreak.Append(pageBreakRun);
            body.Append(pageBreak);
        }

        private void AddSection(Body body, string sectionTitle, string headingStyle)
        {
            Paragraph sectionParagraph = new Paragraph();
            ParagraphProperties sectionProperties = new ParagraphProperties();
            ParagraphStyleId sectionStyleId = new ParagraphStyleId() { Val = headingStyle };
            sectionProperties.Append(sectionStyleId);
            sectionParagraph.Append(sectionProperties);

            Run sectionRun = new Run();
            Text sectionText = new Text(sectionTitle);
            sectionRun.Append(sectionText);
            sectionParagraph.Append(sectionRun);

            body.Append(sectionParagraph);
        }

        private void AddText(Body body, string text, string style)
        {
            Paragraph paragraph = new Paragraph();
            ParagraphProperties properties = new ParagraphProperties();
            ParagraphStyleId styleId = new ParagraphStyleId() { Val = style };
            properties.Append(styleId);
            paragraph.Append(properties);

            Run run = new Run();
            Text textElement = new Text(text);
            run.Append(textElement);
            paragraph.Append(run);

            body.Append(paragraph);
        }

        private void AddControlsTable(Body body)
        {
            if (Controls.Count == 0)
            {
                AddText(body, "No controls found in the form.", "Normal");
                return;
            }

            Table table = new Table();

            // Set table properties
            TableProperties tableProperties = new TableProperties();
            TableBorders tableBorders = new TableBorders();
            TopBorder topBorder = new TopBorder() { Val = BorderValues.Single, Size = 4U };
            BottomBorder bottomBorder = new BottomBorder() { Val = BorderValues.Single, Size = 4U };
            LeftBorder leftBorder = new LeftBorder() { Val = BorderValues.Single, Size = 4U };
            RightBorder rightBorder = new RightBorder() { Val = BorderValues.Single, Size = 4U };
            InsideHorizontalBorder insideHBorder = new InsideHorizontalBorder() { Val = BorderValues.Single, Size = 4U };
            InsideVerticalBorder insideVBorder = new InsideVerticalBorder() { Val = BorderValues.Single, Size = 4U };

            tableBorders.Append(topBorder);
            tableBorders.Append(bottomBorder);
            tableBorders.Append(leftBorder);
            tableBorders.Append(rightBorder);
            tableBorders.Append(insideHBorder);
            tableBorders.Append(insideVBorder);

            tableProperties.Append(tableBorders);
            table.Append(tableProperties);

            // Add header row
            TableRow headerRow = new TableRow();

            // Header cells
            headerRow.Append(CreateTableCell("Control Name", true));
            headerRow.Append(CreateTableCell("Type", true));
            headerRow.Append(CreateTableCell("Binding", true));
            headerRow.Append(CreateTableCell("Details", true));

            table.Append(headerRow);

            // Add data rows
            foreach (var control in Controls)
            {
                TableRow dataRow = new TableRow();

                dataRow.Append(CreateTableCell(control.Name));
                dataRow.Append(CreateTableCell(control.Type));
                dataRow.Append(CreateTableCell(control.Binding));
                dataRow.Append(CreateTableCell(control.Details));

                table.Append(dataRow);
            }

            body.Append(table);
        }

        private void AddRulesTable(Body body)
        {
            if (Rules.Count == 0)
            {
                AddText(body, "No validation rules found in the form.", "Normal");
                return;
            }

            Table table = new Table();

            // Set table properties
            TableProperties tableProperties = new TableProperties();
            TableBorders tableBorders = new TableBorders();
            TopBorder topBorder = new TopBorder() { Val = BorderValues.Single, Size = 4U };
            BottomBorder bottomBorder = new BottomBorder() { Val = BorderValues.Single, Size = 4U };
            LeftBorder leftBorder = new LeftBorder() { Val = BorderValues.Single, Size = 4U };
            RightBorder rightBorder = new RightBorder() { Val = BorderValues.Single, Size = 4U };
            InsideHorizontalBorder insideHBorder = new InsideHorizontalBorder() { Val = BorderValues.Single, Size = 4U };
            InsideVerticalBorder insideVBorder = new InsideVerticalBorder() { Val = BorderValues.Single, Size = 4U };

            tableBorders.Append(topBorder);
            tableBorders.Append(bottomBorder);
            tableBorders.Append(leftBorder);
            tableBorders.Append(rightBorder);
            tableBorders.Append(insideHBorder);
            tableBorders.Append(insideVBorder);

            tableProperties.Append(tableBorders);
            table.Append(tableProperties);

            // Add header row
            TableRow headerRow = new TableRow();

            // Header cells
            headerRow.Append(CreateTableCell("Field", true));
            headerRow.Append(CreateTableCell("Type", true));
            headerRow.Append(CreateTableCell("Details", true));

            table.Append(headerRow);

            // Add data rows
            foreach (var rule in Rules)
            {
                TableRow dataRow = new TableRow();

                dataRow.Append(CreateTableCell(rule.Field));
                dataRow.Append(CreateTableCell(rule.Type));

                string details = rule.Type == "Validation"
                    ? $"Expression: {rule.Expression}\nMessage: {rule.Message}"
                    : $"Event: {rule.Event}\nFunction: {rule.FunctionName}";
                if (rule.Type.Equals("Required", StringComparison.OrdinalIgnoreCase))
                {
                    details = $"Type: {rule.Expression}";
                }
                dataRow.Append(CreateTableCell(details));

                table.Append(dataRow);
            }

            body.Append(table);
        }

        private void AddDataSourcesTable(Body body)
        {
            if (DataSources.Count == 0)
            {
                AddText(body, "No data sources found in the form.", "Normal");
                return;
            }

            Table table = new Table();

            // Set table properties
            TableProperties tableProperties = new TableProperties();
            TableBorders tableBorders = new TableBorders();
            TopBorder topBorder = new TopBorder() { Val = BorderValues.Single, Size = 4U };
            BottomBorder bottomBorder = new BottomBorder() { Val = BorderValues.Single, Size = 4U };
            LeftBorder leftBorder = new LeftBorder() { Val = BorderValues.Single, Size = 4U };
            RightBorder rightBorder = new RightBorder() { Val = BorderValues.Single, Size = 4U };
            InsideHorizontalBorder insideHBorder = new InsideHorizontalBorder() { Val = BorderValues.Single, Size = 4U };
            InsideVerticalBorder insideVBorder = new InsideVerticalBorder() { Val = BorderValues.Single, Size = 4U };

            tableBorders.Append(topBorder);
            tableBorders.Append(bottomBorder);
            tableBorders.Append(leftBorder);
            tableBorders.Append(rightBorder);
            tableBorders.Append(insideHBorder);
            tableBorders.Append(insideVBorder);

            tableProperties.Append(tableBorders);
            table.Append(tableProperties);

            // Add header row
            TableRow headerRow = new TableRow();

            // Header cells
            headerRow.Append(CreateTableCell("Data Source Name", true));
            headerRow.Append(CreateTableCell("Type", true));
            headerRow.Append(CreateTableCell("Description", true));

            table.Append(headerRow);

            // Add data rows
            foreach (var dataSource in DataSources)
            {
                TableRow dataRow = new TableRow();

                dataRow.Append(CreateTableCell(dataSource.Name));
                dataRow.Append(CreateTableCell(dataSource.Type));
                dataRow.Append(CreateTableCell(dataSource.Description));

                table.Append(dataRow);
            }

            body.Append(table);
        }

        private void AddScriptsTable(Body body)
        {
            if (Scripts.Count == 0)
            {
                AddText(body, "No scripts or business logic found in the form.", "Normal");
                return;
            }

            Table table = new Table();

            // Set table properties
            TableProperties tableProperties = new TableProperties();
            TableBorders tableBorders = new TableBorders();
            TopBorder topBorder = new TopBorder() { Val = BorderValues.Single, Size = 4U };
            BottomBorder bottomBorder = new BottomBorder() { Val = BorderValues.Single, Size = 4U };
            LeftBorder leftBorder = new LeftBorder() { Val = BorderValues.Single, Size = 4U };
            RightBorder rightBorder = new RightBorder() { Val = BorderValues.Single, Size = 4U };
            InsideHorizontalBorder insideHBorder = new InsideHorizontalBorder() { Val = BorderValues.Single, Size = 4U };
            InsideVerticalBorder insideVBorder = new InsideVerticalBorder() { Val = BorderValues.Single, Size = 4U };

            tableBorders.Append(topBorder);
            tableBorders.Append(bottomBorder);
            tableBorders.Append(leftBorder);
            tableBorders.Append(rightBorder);
            tableBorders.Append(insideHBorder);
            tableBorders.Append(insideVBorder);

            tableProperties.Append(tableBorders);
            table.Append(tableProperties);

            // Add header row
            TableRow headerRow = new TableRow();

            // Header cells
            headerRow.Append(CreateTableCell("Function Name", true));
            headerRow.Append(CreateTableCell("Parameters", true));
            headerRow.Append(CreateTableCell("Source File", true));

            table.Append(headerRow);

            // Add data rows
            foreach (var script in Scripts)
            {
                TableRow dataRow = new TableRow();

                dataRow.Append(CreateTableCell(script.Name));
                dataRow.Append(CreateTableCell(script.Parameters));
                dataRow.Append(CreateTableCell(script.FilePath));

                table.Append(dataRow);
            }

            body.Append(table);
        }

        private void AddWorkflowsTable(Body body)
        {
            if (Workflows.Count == 0)
            {
                AddText(body, "No workflows found in the form.", "Normal");
                return;
            }

            Table table = new Table();

            // Set table properties
            TableProperties tableProperties = new TableProperties();
            TableBorders tableBorders = new TableBorders();
            TopBorder topBorder = new TopBorder() { Val = BorderValues.Single, Size = 4U };
            BottomBorder bottomBorder = new BottomBorder() { Val = BorderValues.Single, Size = 4U };
            LeftBorder leftBorder = new LeftBorder() { Val = BorderValues.Single, Size = 4U };
            RightBorder rightBorder = new RightBorder() { Val = BorderValues.Single, Size = 4U };
            InsideHorizontalBorder insideHBorder = new InsideHorizontalBorder() { Val = BorderValues.Single, Size = 4U };
            InsideVerticalBorder insideVBorder = new InsideVerticalBorder() { Val = BorderValues.Single, Size = 4U };

            tableBorders.Append(topBorder);
            tableBorders.Append(bottomBorder);
            tableBorders.Append(leftBorder);
            tableBorders.Append(rightBorder);
            tableBorders.Append(insideHBorder);
            tableBorders.Append(insideVBorder);

            tableProperties.Append(tableBorders);
            table.Append(tableProperties);

            // Add header row
            TableRow headerRow = new TableRow();

            // Header cells
            headerRow.Append(CreateTableCell("Workflow Name", true));
            headerRow.Append(CreateTableCell("Description", true));

            table.Append(headerRow);

            // Add data rows
            foreach (var workflow in Workflows)
            {
                TableRow dataRow = new TableRow();

                dataRow.Append(CreateTableCell(workflow.Name));
                dataRow.Append(CreateTableCell(workflow.Description));

                table.Append(dataRow);
            }

            body.Append(table);
        }

        private void AddDependenciesTable(Body body)
        {
            if (DependencyList.Count == 0)
            {
                AddText(body, "No external dependencies detected.", "Normal");
                return;
            }

            Table table = new Table();

            // Set table borders (using the same style as other tables)
            TableProperties tableProperties = new TableProperties();
            TableBorders tableBorders = new TableBorders(
                new TopBorder() { Val = BorderValues.Single, Size = 4U },
                new BottomBorder() { Val = BorderValues.Single, Size = 4U },
                new LeftBorder() { Val = BorderValues.Single, Size = 4U },
                new RightBorder() { Val = BorderValues.Single, Size = 4U },
                new InsideHorizontalBorder() { Val = BorderValues.Single, Size = 4U },
                new InsideVerticalBorder() { Val = BorderValues.Single, Size = 4U }
            );
            tableProperties.Append(tableBorders);
            table.Append(tableProperties);

            // Header row
            TableRow headerRow = new TableRow();
            headerRow.Append(CreateTableCell("DLL Name", true));
            headerRow.Append(CreateTableCell("Project", true));
            headerRow.Append(CreateTableCell("Language", true));
            headerRow.Append(CreateTableCell("Script Language Version", true));
            table.Append(headerRow);

            // Data rows for each dependency
            foreach (var dep in DependencyList)
            {
                TableRow row = new TableRow();
                row.Append(CreateTableCell(dep.Name));
                row.Append(CreateTableCell(dep.Project));
                row.Append(CreateTableCell(dep.Language));
                row.Append(CreateTableCell(dep.ScriptLanguageVersion));
                table.Append(row);
            }

            body.Append(table);
        }

        private void AddImagesTable(Body body)
        {
            if (ImageFiles.Count == 0)
            {
                AddText(body, "No images detected.", "Normal");
                return;
            }

            Table table = new Table();

            // Set table borders (using the same style as other tables)
            TableProperties tableProperties = new TableProperties();
            TableBorders tableBorders = new TableBorders(
                new TopBorder() { Val = BorderValues.Single, Size = 4U },
                new BottomBorder() { Val = BorderValues.Single, Size = 4U },
                new LeftBorder() { Val = BorderValues.Single, Size = 4U },
                new RightBorder() { Val = BorderValues.Single, Size = 4U },
                new InsideHorizontalBorder() { Val = BorderValues.Single, Size = 4U },
                new InsideVerticalBorder() { Val = BorderValues.Single, Size = 4U }
            );
            tableProperties.Append(tableBorders);
            table.Append(tableProperties);

            // Header row
            TableRow headerRow = new TableRow();
            headerRow.Append(CreateTableCell("FileName", true));
            table.Append(headerRow);

            // Data rows for each dependency
            foreach (var img in ImageFiles)
            {
                TableRow row = new TableRow();
                row.Append(CreateTableCell(img.FileName));
                table.Append(row);
            }

            body.Append(table);
        }
        private TableCell CreateTableCell(string text, bool isHeader = false)
        {
            TableCell cell = new TableCell();

            // Create paragraph
            Paragraph paragraph = new Paragraph();
            Run run = new Run();

            // Add text
            Text cellText = new Text(text);
            run.Append(cellText);

            // Make header bold
            if (isHeader)
            {
                RunProperties runProperties = new RunProperties();
                Bold bold = new Bold();
                runProperties.Append(bold);
                run.PrependChild(runProperties);
            }

            paragraph.Append(run);
            cell.Append(paragraph);

            return cell;
        }

        private void AddPageNumbers(WordprocessingDocument wordDocument)
        {
            // Get the main document part
            MainDocumentPart mainPart = wordDocument.MainDocumentPart;

            // Create a new footer part
            FooterPart footerPart = mainPart.AddNewPart<FooterPart>();

            // Create footer
            Footer footer = new Footer();

            // Create paragraph for the footer
            Paragraph footerParagraph = new Paragraph();
            ParagraphProperties paragraphProperties = new ParagraphProperties();
            Justification justification = new Justification() { Val = JustificationValues.Center };
            paragraphProperties.Append(justification);
            footerParagraph.Append(paragraphProperties);

            // Create run for the page number
            Run run = new Run();
            Text text = new Text("Page ");
            run.Append(text);
            footerParagraph.Append(run);

            // Add the page number field
            Run pageFieldRun = new Run();
            pageFieldRun.Append(new SimpleField() { Instruction = "PAGE" });
            footerParagraph.Append(pageFieldRun);

            // Add " of " text
            Run ofRun = new Run();
            Text ofText = new Text(" of ");
            ofRun.Append(ofText);
            footerParagraph.Append(ofRun);

            // Add the num pages field
            Run numPagesFieldRun = new Run();
            numPagesFieldRun.Append(new SimpleField() { Instruction = "NUMPAGES" });
            footerParagraph.Append(numPagesFieldRun);

            // Add the paragraph to the footer
            footer.Append(footerParagraph);

            // Add the footer to the footer part
            footerPart.Footer = footer;

            // Create a new section properties element
            SectionProperties sectionProperties = new SectionProperties();

            // Create a new footer reference
            FooterReference footerReference = new FooterReference()
            {
                Type = HeaderFooterValues.Default,
                Id = mainPart.GetIdOfPart(footerPart)
            };

            // Add the footer reference to the section properties
            sectionProperties.Append(footerReference);

            // Add the section properties to the document
            Body body = mainPart.Document.Body;

            // If there are existing section properties, replace them; otherwise add new ones
            SectionProperties existingSectionProperties = body.Elements<SectionProperties>().FirstOrDefault();
            if (existingSectionProperties != null)
            {
                // Add the footer reference to the existing section properties
                existingSectionProperties.Append(footerReference);
            }
            else
            {
                // Add new section properties
                body.Append(sectionProperties);
            }
        }

        private void UpdateTableOfContents(WordprocessingDocument wordDocument)
        {
            MainDocumentPart mainPart = wordDocument.MainDocumentPart;
            Body body = mainPart.Document.Body;

            // Find the TOC section
            var tocParagraphs = body.Elements<Paragraph>()
                .Where(p => p.InnerText.Contains("Table of Contents"))
                .ToList();

            if (tocParagraphs.Count > 0)
            {
                var tocParagraph = tocParagraphs[0];
                var nextParagraph = tocParagraph.NextSibling<Paragraph>();

                if (nextParagraph != null)
                {
                    //Update Fields on open
                    var settingsPart = mainPart.AddNewPart<DocumentSettingsPart>();
                    settingsPart.Settings = new Settings { BordersDoNotSurroundFooter = new BordersDoNotSurroundFooter() { Val = true } };
                    settingsPart.Settings.Append(new UpdateFieldsOnOpen() { Val = true });
                }
            }
        }

        private static string GetTOC(string title, int titleFontSize)
        {
            return $@"<w:sdt>
                     <w:sdtPr>
                        <w:id w:val=""-493258456"" />
                        <w:docPartObj>
                           <w:docPartGallery w:val=""Table of Contents"" />
                           <w:docPartUnique />
                        </w:docPartObj>
                     </w:sdtPr>
                     <w:sdtEndPr>
                        <w:rPr>
                           <w:rFonts w:asciiTheme=""minorHAnsi"" w:eastAsiaTheme=""minorHAnsi"" w:hAnsiTheme=""minorHAnsi"" w:cstheme=""minorBidi"" />
                           <w:b />
                           <w:bCs />
                           <w:noProof />
                           <w:color w:val=""auto"" />
                           <w:sz w:val=""22"" />
                           <w:szCs w:val=""22"" />
                        </w:rPr>
                     </w:sdtEndPr>
                     <w:sdtContent>
                        <w:p w:rsidR=""00095C65"" w:rsidRDefault=""00095C65"">
                           <w:pPr>
                              <w:pStyle w:val=""TOCHeading"" />
                              <w:jc w:val=""center"" /> 
                           </w:pPr>
                           <w:r>
                                <w:rPr>
                                  <w:b /> 
                                  <w:color w:val=""2E74B5"" w:themeColor=""accent1"" w:themeShade=""BF"" /> 
                                  <w:sz w:val=""{titleFontSize * 2}"" /> 
                                  <w:szCs w:val=""{titleFontSize * 2}"" /> 
                              </w:rPr>
                              <w:t>{title}</w:t>
                           </w:r>
                        </w:p>
                        <w:p w:rsidR=""00095C65"" w:rsidRDefault=""00095C65"">
                           <w:r>
                              <w:rPr>
                                 <w:b />
                                 <w:bCs />
                                 <w:noProof />
                              </w:rPr>
                              <w:fldChar w:fldCharType=""begin"" />
                           </w:r>
                           <w:r>
                              <w:rPr>
                                 <w:b />
                                 <w:bCs />
                                 <w:noProof />
                              </w:rPr>
                              <w:instrText xml:space=""preserve""> TOC \o ""1-3"" \h \z \u </w:instrText>
                           </w:r>
                           <w:r>
                              <w:rPr>
                                 <w:b />
                                 <w:bCs />
                                 <w:noProof />
                              </w:rPr>
                              <w:fldChar w:fldCharType=""separate"" />
                           </w:r>
                           <w:r>
                              <w:rPr>
                                 <w:noProof />
                              </w:rPr>
                              <w:t>No table of contents entries found.</w:t>
                           </w:r>
                           <w:r>
                              <w:rPr>
                                 <w:b />
                                 <w:bCs />
                                 <w:noProof />
                              </w:rPr>
                              <w:fldChar w:fldCharType=""end"" />
                           </w:r>
                        </w:p>
                     </w:sdtContent>
                  </w:sdt>";
        }
    }

    // Data models for storing InfoPath form information
    public class Control
    {
        public string Details { get; set; }
        public string Name { get; set; }
        public string Type { get; set; }
        public string Binding { get; set; }
        public string UniqueId { get; set; }
        public string IsVisible { get; set; }
    }

    public class DataSource
    {
        public string Name { get; set; }
        public string Type { get; set; }
        public string Description { get; set; }
    }

    public class Rule
    {
        public string Name { get; set; }
        public string Type { get; set; }
        public string Expression { get; set; }
        public string Message { get; set; }
        public string Event { get; set; }
        public string FunctionName { get; set; }
        public string Field { get; set; }
    }

    public class Script
    {
        public string Name { get; set; }
        public string Parameters { get; set; }
        public string FilePath { get; set; }
    }

    public class Workflow
    {
        public string Name { get; set; }
        public string Description { get; set; }
    }

    public class Dependency
    {
        public string Name { get; set; }
        public string Project { get; set; }
        public string Language { get; set; }
        public string ScriptLanguageVersion { get; set; }
    }

    public class ImageFile
    {
        public string FileName { get; set; }
    }

    class UserFormVariable
    {
        public string Name { get; set; }
        public string Type { get; set; }
    }
}