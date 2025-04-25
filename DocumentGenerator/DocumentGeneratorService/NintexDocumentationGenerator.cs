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
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Color = DocumentFormat.OpenXml.Wordprocessing.Color;
using Document = DocumentFormat.OpenXml.Wordprocessing.Document;
using Text = DocumentFormat.OpenXml.Wordprocessing.Text;

namespace DocumentGeneratorService
{
    public class NintexDocumentationGenerator : IGenerateDocument
    {
        public string Output { get; set; }
        // Main class properties to store form information
        private List<Control> Controls { get; set; } = new List<Control>();
        private List<Layouts> FormLayouts { get; set; } = new List<Layouts>();
        private List<NintexFormRule> Rules { get; set; } = new List<NintexFormRule>();
        private List<Script> Scripts { get; set; } = new List<Script>();
        private List<Workflow> Workflows { get; set; } = new List<Workflow>();
        private Dictionary<string, List<string>> Dependencies { get; set; } = new Dictionary<string, List<string>>();
        private List<Dependency> DependencyList { get; set; } = new List<Dependency>();
        private List<ImageFile> ImageFiles { get; set; } = new List<ImageFile>();

        private List<UserFormVariable> formVariables = new List<UserFormVariable>();
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

        private void LoadRules(string directoryPath)
        {
            try
            {
                // Find all files ending with "rules.xml" in the specified directory and subdirectories
                var ruleFiles = Directory.GetFiles(directoryPath, "*Rules.xml", SearchOption.AllDirectories);

                if (ruleFiles.Length == 0)
                {
                    Console.WriteLine("No files ending with 'Rules.xml' found in the specified directory.");
                    return;
                }

                // Process each file
                foreach (var filePath in ruleFiles)
                {
                    Console.WriteLine($"Processing: {filePath}");

                    // Load the XML from the file
                    XDocument xmlDoc = XDocument.Load(filePath);

                    // Define namespaces
                    XNamespace defaultNs = "http://schemas.datacontract.org/2004/07/Nintex.Forms";
                    XNamespace xsiNs = "http://www.w3.org/2001/XMLSchema-instance";

                    // Extract FormControlProperties, including nested elements
                    var rules = xmlDoc.Descendants(defaultNs + "Rule")
                                         .Select(rule => new NintexFormRule
                                         {
                                             Expression = rule.Element(defaultNs + "Expression")?.Value,
                                             ExpressionValue = rule.Element(defaultNs + "ExpressionValue")?.Value,
                                             RuleType = rule.Element(defaultNs + "RuleType")?.Value,
                                             ControlName = Controls.FirstOrDefault(c => c.UniqueId == rule.Element(defaultNs + "ControlIds")?.Value)?.Name ?? "Unknown Control",
                                             IsDisabled = rule.Element(defaultNs + "Disable")?.Value == "true" ? "Yes" : "No"
                                         })
                                         .ToList();

                    //var controlInfo = Controls.FirstOrDefault(c => c.UniqueId == controlIdsValue);
                    //var controlName = controlInfo?.Name ?? "Unknown Control";

                    this.Rules.AddRange(rules);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        private void LoadLayouts(string directoryPath)
        {
            try
            {
                // Find all files ending with "rules.xml" in the specified directory and subdirectories
                var layoutFiles = Directory.GetFiles(directoryPath, "*Layouts.xml", SearchOption.AllDirectories);

                if (layoutFiles.Length == 0)
                {
                    Console.WriteLine("No files ending with 'Layouts.xml' found in the specified directory.");
                    return;
                }

                // Process each file
                foreach (var filePath in layoutFiles)
                {
                    Console.WriteLine($"Processing: {filePath}");

                    // Load the XML from the file
                    XDocument xmlDoc = XDocument.Load(filePath);

                    // Define namespaces
                    XNamespace defaultNs = "http://schemas.datacontract.org/2004/07/Nintex.Forms";
                    XNamespace xsiNs = "http://www.w3.org/2001/XMLSchema-instance";

                    // Extract FormControlProperties, including nested elements
                    var layouts = xmlDoc.Descendants(defaultNs + "FormLayout")
                                         .Select(layout => new Layouts
                                         {
                                             Name = layout.Element(defaultNs + "DeviceName")?.Value,
                                             Title = layout.Element(defaultNs + "Title")?.Value
                                         })
                                         .ToList();

                    this.FormLayouts.AddRange(layouts);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        private void LoadUserVariables(string directoryPath)
        {
            //string xmlFilePath = "UserVariables.xml";
            var xmlFilePath = Directory.GetFiles(directoryPath, "*variables.xml", SearchOption.AllDirectories);
           
            foreach (var file in xmlFilePath)
            {
                XDocument xmlDoc = XDocument.Load(file);

                // Get the namespace from XML
                XNamespace ns = "http://schemas.datacontract.org/2004/07/Nintex.Forms";


                foreach (XElement variable in xmlDoc.Root.Elements(ns + "UserFormVariable"))
                {
                    formVariables.Add(new UserFormVariable
                    {
                        Name = variable.Element(ns + "Name")?.Value,
                        Type = variable.Element(ns + "Type")?.Value
                    });
                }
            }
        }
        private void LoadControls(string directoryPath)
        {
           // string directoryPath = @"C:\path\to\search";

            try
            {
                // Find all files ending with "controls.xml" in the specified directory and subdirectories
                var controlFiles = Directory.GetFiles(directoryPath, "*controls.xml", SearchOption.AllDirectories);

                if (controlFiles.Length == 0)
                {
                    Console.WriteLine("No files ending with 'controls.xml' found in the specified directory.");
                    return;
                }

                this.Controls = new List<Control>();

                // Process each file
                foreach (var filePath in controlFiles)
                {
                    Console.WriteLine($"Processing: {filePath}");

                    // Load the XML from the file
                    XDocument xmlDoc = XDocument.Load(filePath);

                    // Define namespaces
                    XNamespace defaultNs = "http://schemas.datacontract.org/2004/07/Nintex.Forms.FormControls";
                    XNamespace xsiNs = "http://www.w3.org/2001/XMLSchema-instance";

                    // Extract FormControlProperties, including nested elements
                    var controls = xmlDoc.Descendants(defaultNs + "FormControlProperties")
                                         .Select(control => new Control
                                         {
                                             Name = control.Element(defaultNs + "DisplayName")?.Value,
                                             Type = control.Attribute(xsiNs + "type")?.Value,
                                             UniqueId = control.Element(defaultNs + "UniqueId")?.Value,
                                             IsVisible = control.Element(defaultNs + "IsVisible")?.Value
                                         })
                                         .ToList();


                    Controls.AddRange(controls);
                }

                // Print the total number of controls and their details
                //Console.WriteLine($"Total Controls Found: {Controls.Count}");
                //foreach (var control in allControls)
               // {
                 //   Console.WriteLine($"Details: {control.Details}, Name: {control.Name}, Type: {control.Type}, Binding: {control.Binding}");
               // }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
        private void ExtractXsnContents(string xsnFilePath, string targetFilePath)
        {
            Console.WriteLine("Extracting InfoPath form contents...");

            try
            {
                
                ZipFile.ExtractToDirectory(xsnFilePath, targetFilePath);
                string[] files = Directory.GetFiles(targetFilePath, "*", SearchOption.AllDirectories);

                this.Output = xsnFilePath + "\n\n";
                foreach (string file in files)
                {
                    this.Output += file + "\n";
                }

                this.Output += files.Length + " files extracted\n\n";
            }
            catch (Exception ex)
            {
                throw new Exception("Failed to extract InfoPath form: " + ex.Message, ex);
            }
        }

        private void ParseFormDefinition(string extractedFolder, string outputPath)
        {
            Console.WriteLine("Parsing form definition...");

            ////Look for key files

            string manifestFile = Path.Combine(extractedFolder, "manifest.xml");
            if (!File.Exists(manifestFile))
            {
                throw new FileNotFoundException("Manifest file (manifest.xsf) not found in the InfoPath form package.");
            }

            // Parse the manifest file
            XDocument manifest = XDocument.Load(manifestFile);
            XNamespace xsfNs = "http://schemas.microsoft.com/office/infopath/2003/solutionDefinition";
            
            // Parse controls
            LoadControls(extractedFolder);

            // Parse User Variables
            LoadUserVariables(extractedFolder);

            // Parse Layouts
            LoadLayouts(extractedFolder);

            // Parse Rules
            LoadRules(extractedFolder);

            // Look for script files
            foreach (var scriptFile in Directory.GetFiles(extractedFolder, "*.js"))
            {
                ParseScriptFile(scriptFile);
            }
            foreach (var scriptFile in Directory.GetFiles(extractedFolder, "*scripts.xml",SearchOption.AllDirectories))
            {
                ParseScriptFile(scriptFile);
            }
            foreach (var dllFile in Directory.GetFiles(extractedFolder, "*.dll"))
            {
                ParseDllFile(dllFile);
            }
                        
            ParseImages(manifest, xsfNs);

            //CopyFilesIntoReferenceFolder(outputPath, extractedFolder);
            Console.WriteLine("Parsing completed.");
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
                XDocument xmlDoc = XDocument.Load(scriptFilePath);
                string scriptContent = xmlDoc.Root.Element("Script").Value;

                MatchCollection matches = Regex.Matches(scriptContent, @"function\s+(\w+)\s*\(([^)]*)\)\s*{([^}]*)}", RegexOptions.Singleline);

                foreach (Match match in matches)
                {
                    Scripts.Add(new Script
                    {
                        Name = match.Groups[1].Value,
                        Parameters = match.Groups[2].Value.Replace(" ", ""), // Remove extra spaces
                        FilePath = match.Groups[3].Value
                    });
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Warning: Could not parse script file {scriptFilePath}: {ex.Message}");
            }
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
                AddDocumentTitle(body, "Nintex Form Documentation");

                // Add table of contents
                AddTableOfContents(body);

                // Add sections
                int sectionNumber = 1;

                // Controls Section
                AddSection(body, $"{sectionNumber++}. Controls", "Heading1");
                AddControlsTable(body);
                body.Append(new Paragraph(new Run(new Break()))); // Add a line break

                // Rules Section
                AddSection(body, $"{sectionNumber++}. Rules", "Heading1");
                AddRulesTable(body);
                body.Append(new Paragraph(new Run(new Break()))); // Add a line break

                // Data Sources Section
                AddSection(body, $"{sectionNumber++}. Layouts", "Heading1");
                AddFormLayouts(body);
                body.Append(new Paragraph(new Run(new Break()))); // Add a line break

                // Scripts Section
                AddSection(body, $"{sectionNumber++}. Scripts and Business Logic", "Heading1");
                AddScriptsTable(body);
                body.Append(new Paragraph(new Run(new Break()))); // Add a line break

                // Dependencies Section
                AddSection(body, $"{sectionNumber++}. User Variables", "Heading1");
                AddUserVariables(body);
                body.Append(new Paragraph(new Run(new Break()))); // Add a line break

                // Dependencies Section
                AddSection(body, $"{sectionNumber++}. Assets", "Heading1");
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
            headerRow.Append(CreateTableCell("IsVisible", true));

            table.Append(headerRow);

            // Add data rows
            foreach (var control in Controls)
            {
                TableRow dataRow = new TableRow();

                dataRow.Append(CreateTableCell(control.Name));
                dataRow.Append(CreateTableCell(control.Type));
                dataRow.Append(CreateTableCell(control.IsVisible));

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
            headerRow.Append(CreateTableCell("Control Name", true));
            headerRow.Append(CreateTableCell("Rule Type", true));
            headerRow.Append(CreateTableCell("IsDisabled", true));
            headerRow.Append(CreateTableCell("Expression", true));
            headerRow.Append(CreateTableCell("Expression Value", true));

            table.Append(headerRow);

            // Add data rows
            foreach (var rule in Rules)
            {
                TableRow dataRow = new TableRow();
                dataRow.Append(CreateTableCell(rule.ControlName));
                dataRow.Append(CreateTableCell(rule.RuleType));
                dataRow.Append(CreateTableCell(rule.IsDisabled));
                dataRow.Append(CreateTableCell(rule.Expression));
                dataRow.Append(CreateTableCell(rule.ExpressionValue));               

                table.Append(dataRow);
            }

            body.Append(table);
        }

        private void AddFormLayouts(Body body)
        {
            if (FormLayouts.Count == 0)
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
            headerRow.Append(CreateTableCell("Device Name", true));
            headerRow.Append(CreateTableCell("Title", true));
           // headerRow.Append(CreateTableCell("Description", true));

            table.Append(headerRow);

            // Add data rows
            foreach (var formLayout in FormLayouts)
            {
                TableRow dataRow = new TableRow();

                dataRow.Append(CreateTableCell(formLayout.Name));
                dataRow.Append(CreateTableCell(formLayout.Title));

                table.Append(dataRow);
            }

            body.Append(table);
        }

        private void AddUserVariables(Body body)
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
            headerRow.Append(CreateTableCell("Name", true));
            headerRow.Append(CreateTableCell("Type", true));
            // headerRow.Append(CreateTableCell("Source File", true));

            table.Append(headerRow);

            // Add data rows
            foreach (var script in formVariables)
            {
                TableRow dataRow = new TableRow();

                dataRow.Append(CreateTableCell(script.Name));
                dataRow.Append(CreateTableCell(script.Type));
                //  dataRow.Append(CreateTableCell(script.FilePath));

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
           // headerRow.Append(CreateTableCell("Source File", true));

            table.Append(headerRow);

            // Add data rows
            foreach (var script in Scripts)
            {
                TableRow dataRow = new TableRow();

                dataRow.Append(CreateTableCell(script.Name));
                dataRow.Append(CreateTableCell(script.Parameters));
              //  dataRow.Append(CreateTableCell(script.FilePath));

                table.Append(dataRow);
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

    public class Layouts
    {
        public string Name { get; set; }
        public string Title { get; set; }
    }

    public class NintexFormRule
    {        
        public string Expression { get; set; }        
        public string ExpressionValue { get; set; }
        public string RuleType { get; set; }
        public string ControlName { get; set; }
        public string IsDisabled { get; set; }
    }    
}