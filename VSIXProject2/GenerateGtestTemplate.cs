using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
//using Community.VisualStudio.Toolkit;

using Task = System.Threading.Tasks.Task;
using System.Windows.Forms;
using System.Text;
using System.Linq;
using Microsoft.VisualStudio.Settings;
using Microsoft.VisualStudio.Shell.Settings;
using System.Text.RegularExpressions;

namespace VSIXProject2
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class GenerateGtestTemplate
    {
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 0x0100;
        public const int CommandIdMock = 0x0105;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("4604f7c9-dbe1-4db8-a82c-148761b330b2");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly AsyncPackage package;
        WritableSettingsStore userSettingsStore;
        string text;
        string[] eachline;

        /// <summary>
        /// Initializes a new instance of the <see cref="GenerateGtestTemplate"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        /// <param name="commandService">Command service to add command to, not null.</param>
        private GenerateGtestTemplate(AsyncPackage package, OleMenuCommandService commandService)
        {
            this.package = package ?? throw new ArgumentNullException(nameof(package));
            commandService = commandService ?? throw new ArgumentNullException(nameof(commandService));

            var menuCommandID = new CommandID(CommandSet, CommandId);
            var menuItem = new MenuCommand(this.Execute, menuCommandID);
            commandService.AddCommand(menuItem);
            var menuCommandIDMock = new CommandID(CommandSet, CommandIdMock);
            var menuItemMock = new MenuCommand(this.ExecuteForMock, menuCommandIDMock);
            commandService.AddCommand(menuItemMock);
        }


        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        public static GenerateGtestTemplate Instance
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets the service provider from the owner package.
        /// </summary>
        private IServiceProvider ServiceProvider
        {
            get
            {
                return this.package;
            }
        }

        /// <summary>
        /// Initializes the singleton instance of the command.
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        public static async Task InitializeAsync(AsyncPackage package)
        {
            // Switch to the main thread - the call to AddCommand in GenerateGtestTemplate's constructor requires
            // the UI thread.
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

            OleMenuCommandService commandService = await package.GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;
            Instance = new GenerateGtestTemplate(package, commandService);
        }

        private void ExecuteForMock(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            var dte = ServiceProvider.GetService(typeof(DTE)) as DTE2;
            if (dte == null)
            {
                ShowMessage("Unknown error occured while loading");
                return;
            }
            var currentlyOpenTabfilePath = dte.ActiveDocument.FullName;
            if (string.IsNullOrEmpty(currentlyOpenTabfilePath))
                return;

            string ext = Path.GetExtension(currentlyOpenTabfilePath);
            if (ext != ".h")
            {
                ShowMessage("Select the header file for which you need to create a mock");
                return;
            }

            var selection = (TextSelection)dte.ActiveDocument.Selection;
            var activePoint = selection.ActivePoint;
            string entireLine = activePoint.CreateEditPoint().GetLines(activePoint.Line, activePoint.Line + 1);
            if (entireLine == "")
            {
                ShowMessage("Select the class name for which you want to create a mock");
                return;
            }
            bool isInterface = false;
            if (entireLine.Contains("__interface"))
            {
                isInterface = true;
            }
            string className = "";
            string IclassName = GetCurrentClassName(GetHeaderFileName(currentlyOpenTabfilePath));
            char secondLetter = IclassName.ElementAt(1);
            if (isInterface && IclassName.StartsWith("I") && Char.IsUpper(secondLetter))
            {
                className = IclassName.Remove(0, 1);
            }
            else
            {
                className = IclassName;
            }
            string text = File.ReadAllText(currentlyOpenTabfilePath);
            string[] lines = text.Split('\n');
            int i = 0;
            while(!(lines[i].Contains(entireLine)))
            {
                i++;               
            }
            string l = lines[i+1].Replace("\r\n", "").Replace("\r", "").Replace("\n", "");
            i += 1;
            if (l.Length == 1 && l == "{")
            {
                i = i + 1;
            }
            StringBuilder generatedMock = new StringBuilder();
            generatedMock.Append("class" + " "+className + "Mock" + ": " + "public" + " " +IclassName+"\n");
            generatedMock.Append("{\n");
            generatedMock.Append("public:\n");
            for (int j = i; j < lines.Length;j++)
            {
                string combinedLine = "";
                if (lines[j].Contains("};"))
                {
                    generatedMock.Append("};").Append("\n\n");
                    break;

                }
                if (isInterface || lines[j].Contains("virtual"))
                {
                    lines[j] = lines[j].Replace("virtual", "");
                    if (lines[j].Contains("(") && !(lines[j].Contains(")")))
                    {
                        while (!(lines[j].Contains(")")))
                        {
                            combinedLine += lines[j];
                            j++;
                        }
                        combinedLine += lines[j];
                        GetMockMethod(ref generatedMock, combinedLine);
                    }
                    else if (lines[j].Contains(")"))
                    {
                        GetMockMethod(ref generatedMock, lines[j]);
                    }
                }
               
            }
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.InitialDirectory = GetCurrentDirectory(currentlyOpenTabfilePath);
            openFileDialog.Filter = "Source files|*.cpp|Header files|*.h";
            string fileName = "";
            if(openFileDialog.ShowDialog() == DialogResult.OK)
            {
                fileName = openFileDialog.FileName;
            }
            if (fileName != "")
            {
                text = File.ReadAllText(fileName);
                if (text.Length == 0)
                {
                    string headers = GetHeaders(currentlyOpenTabfilePath);
                    generatedMock = generatedMock.Insert(0, headers);
                    WriteToFile(fileName, generatedMock.ToString());
                }
                else
                {
                    dte.ExecuteCommand("File.OpenFile", fileName);
                    dte.ExecuteCommand("File.Close", string.Empty);
                    int position = GetPositionToAppend(text);
                    if (position < text.Length)
                    {
                        generatedMock.Append(text.Substring(position));
                        generatedMock.Append("\n");
                        AppendToFile(fileName, generatedMock.ToString(), position);
                    }
                    else
                    {
                        string generatedMockString = generatedMock.ToString();
                        int pos = getGeneratedTemplateToAppend(fileName, ref generatedMockString);
                        AppendToFile(fileName, generatedMockString, pos);
                    }
                }
                dte.ExecuteCommand("File.OpenFile", fileName);
                dte.ExecuteCommand("Edit.FormatDocument", string.Empty);
            }
            else
            {
                ShowMessage("Enter a path in which to create the mock file!");
            }
        }

        private static void GetMockMethod(ref StringBuilder generatedMock, string line)
        {
            string ReturnTypeAndMethodName = line.Substring(0, line.IndexOf('('));
            ReturnTypeAndMethodName = ReturnTypeAndMethodName.Trim();
            int lengthOfArgs = (line.IndexOf(')') - line.IndexOf('(')) + 1;
            string Args = line.Substring(line.IndexOf('('), lengthOfArgs);
            string[] ReturnTypeAndMethodNameSplit = ReturnTypeAndMethodName.Split(' ');
            generatedMock.Append("MOCK_METHOD" + "(" + ReturnTypeAndMethodNameSplit[0] + "," + ReturnTypeAndMethodNameSplit[ReturnTypeAndMethodNameSplit.Length - 1] + "," + Args + "," + "(override)" + ")" + ";");
            generatedMock.Append("\n");
        }

        private int GetPositionToAppend(string text)
        {
            int position = 0;
            string[] lines = text.Split('\n');
            for(int i=0; i<lines.Length;i++)
            {
                if (lines[i].Contains("TYPED_TEST_SUITE"))
                {
                    position -= ((lines[i - 1].Length)+1);
                    break;
                }
                else
                {
                    position += lines[i].Length + 1;
                }
            }
            return position;
        }

        /// <summary>
        /// This function is the callback used to execute the command when the menu item is clicked.
        /// See the constructor to see how the menu item is associated with this function using
        /// OleMenuCommandService service and MenuCommand class.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event args.</param>
        private void Execute(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            //IVsObjectManager2 vsObjectManager2= ServiceProvider.GetService(typeof(SVsObjectManager)) as IVsObjectManager2;
            //IVsLibrary2 lib;
            //vsObjectManager2.FindLibrary()


            var dte = ServiceProvider.GetService(typeof(DTE)) as DTE2;
            if (dte == null)
            {
                ShowMessage("Unknown error occured while loading");
                return;
            }
            var currentlyOpenTabfilePath = dte.ActiveDocument.FullName;
            if (string.IsNullOrEmpty(currentlyOpenTabfilePath))
                return;

            string ext = Path.GetExtension(currentlyOpenTabfilePath);
            if (ext != ".cpp")
            {
                ShowMessage("Sorry ! Template are created only for cpp files.");
                return;
            }

            var selection = (TextSelection)dte.ActiveDocument.Selection;
            var activePoint = selection.ActivePoint;
            string entireLine = activePoint.CreateEditPoint().GetLines(activePoint.Line, activePoint.Line + 1);
            if (entireLine == "")
            {
                ShowMessage("Select the method name for which you want to create a test");
                return;
            }
            string[] splitMethodNameAndArgs = entireLine.Split('(');
            string[] methodNameWithAccessSpecifiers = splitMethodNameAndArgs[0].Split(' ');
            string methodName = methodNameWithAccessSpecifiers[methodNameWithAccessSpecifiers.Length - 1];
            if (methodName.Contains("::"))
            {
                methodName = methodName.Substring(methodName.IndexOf("::") + 2);
            }
            int index1 = entireLine.IndexOf('(') + 1;
            int index2 = entireLine.IndexOf(')');
            if (index2 < index1)
            {
                ShowMessage("Select a valid method name for which you want to create a test");
                return;
            }
            string args = getArgs(entireLine.Substring(index1, index2 - index1));

            string absoluteFilePath = "";
            SettingsManager settingsManager = new ShellSettingsManager(ServiceProvider);
            userSettingsStore = settingsManager.GetWritableSettingsStore(SettingsScope.UserSettings);

            bool isFirstTimeForTheFile = false;
            CompareInfo Compare = CultureInfo.InvariantCulture.CompareInfo;
            userSettingsStore.CreateCollection("Gtest Template\\");
            bool hasPath = userSettingsStore.CollectionExists("Gtest Template\\" + currentlyOpenTabfilePath);
            if (!hasPath)
            {
                userSettingsStore.CreateCollection("Gtest Template\\" + currentlyOpenTabfilePath);
                isFirstTimeForTheFile = true;
            }

            if (isFirstTimeForTheFile)
            {
                absoluteFilePath = "";
                FolderBrowserDialog fbd = new FolderBrowserDialog();
                fbd.SelectedPath = GetCurrentDirectory(currentlyOpenTabfilePath);
                if (fbd.ShowDialog() == DialogResult.OK)
                {
                    absoluteFilePath = fbd.SelectedPath + "\\";
                }
                else
                {
                    absoluteFilePath = GetCurrentDirectory(currentlyOpenTabfilePath) + "\\";
                }
                userSettingsStore.SetBoolean("Gtest Template", "useUserDefinedPath", true);
                userSettingsStore.SetString("Gtest Template", currentlyOpenTabfilePath, absoluteFilePath);

            }
            else
            {
                if (userSettingsStore.GetBoolean("Gtest Template", "useUserDefinedPath"))
                {
                    absoluteFilePath = userSettingsStore.GetString("Gtest Template", currentlyOpenTabfilePath);
                }
            }
            if (absoluteFilePath == "")
            {
                ShowMessage("Enter a path in which to create the test file!");
            }

            absoluteFilePath = absoluteFilePath + GetSourceFileName(currentlyOpenTabfilePath);
            string fileName = absoluteFilePath.Replace(".cpp", "Test.cpp");

            string generatedTest =
               "TYPED_TEST" + "(" + GetTypeName(fileName, isFirstTimeForTheFile && File.Exists(fileName), currentlyOpenTabfilePath) + "," + "Should" + methodName + ")" +
               "{" + "\n"
               + args + GetAssertString() +
               "}";
            if (isFirstTimeForTheFile && !File.Exists(fileName))
            {
                getTemplateForTestToWrite(absoluteFilePath, ref generatedTest);
                WriteToFile(fileName, generatedTest);
            }
            else
            {
                generatedTest = "\n" + generatedTest;
                generatedTest += "\n";
                int position = getGeneratedTemplateToAppend(fileName, ref generatedTest);
                dte.ExecuteCommand("File.OpenFile", fileName);
                dte.ExecuteCommand("File.Close", string.Empty);
                AppendToFile(fileName, generatedTest, position);
            }

            dte.ExecuteCommand("File.OpenFile", fileName);
            dte.ExecuteCommand("Edit.FormatDocument", string.Empty);
            //dte.ExecuteCommand("File.SaveAll", string.Empty);
            //dte.ExecuteCommand("File.Close", string.Empty);
            //dte.ExecuteCommand("File.OpenFile", fileName);
        }

        private int getGeneratedTemplateToAppend(string fileName, ref string generatedTest)
        {
            int count = 0;
            text = File.ReadAllText(fileName);
            eachline = text.Split('\n');
            foreach (string line in eachline)
            {
                string lineTrimmed = line.Trim();
                if (lineTrimmed.StartsWith("namespace"))
                {
                    count++;
                    generatedTest += "}" + "\n";
                }
                if (line.Contains("class") || line.Contains("TYPED_TEST") || line.Contains("TEST_F") || line.Contains("TEST"))
                {
                    break;
                }
            }
            int val = 0;
            int position = text.Length - 1;
            while (val < count)
            {
                if (text.ElementAt(position) == '}')
                {
                    val++;
                }
                position--;
            }

            return position;
        }

        private void getTemplateForTestToWrite(string absoluteFilePath, ref string generatedTest)
        {
            generatedTest = GetHeaders(absoluteFilePath) + "\n" + AddNamespace() + "\n" + DefineTypes(absoluteFilePath) + "\n" + GetClassInitialization(absoluteFilePath) + "{" + GetConstructorAndDestructor(absoluteFilePath) + "};" + "\n" + "\n"+ generatedTest + "\n" + "}" + "\n" +"}"+"\n";
        }

        private string GetTypeName(string fileName, bool shouldGetTypeFromFile, string currentlyOpenTabfilePath)
        {
            try
            {
                if (shouldGetTypeFromFile)
                {
                    text = File.ReadAllText(fileName);
                    eachline = text.Split('\n');
                    string typeLine = "";
                    foreach (string line in eachline)
                    {
                        if (line.Contains("TYPED_TEST_SUITE("))
                        {
                            typeLine = line;
                        }
                    }
                    int index1 = typeLine.IndexOf("TYPED_TEST_SUITE(") + "TYPED_TEST_SUITE(".Length;
                    int length = typeLine.IndexOf(',') - index1;
                    string TypeName = typeLine.Substring(index1, length);
                    userSettingsStore.CreateCollection("Gtest template\\TypeName");
                    userSettingsStore.SetString("Gtest template", fileName, TypeName);
                    return TypeName;
                }
                else if (File.Exists(fileName))
                {
                    return userSettingsStore.GetString("Gtest template", fileName);
                }
                else
                {
                    return GetTestClassName(GetHeaderFileName(currentlyOpenTabfilePath));
                }
            }
            catch(Exception e)
            {
                return GetTestClassName(GetHeaderFileName(currentlyOpenTabfilePath));
            }
        }

        private void ShowMessage(string text)
        {
            string message = string.Format(CultureInfo.CurrentCulture, text, this.GetType().FullName);
            string title = "GenerateGtestTemplate";

            // Show a message box to prove we were here
            VsShellUtilities.ShowMessageBox(
                this.package,
                message,
                title,
                OLEMSGICON.OLEMSGICON_INFO,
                OLEMSGBUTTON.OLEMSGBUTTON_OK,
                OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
        }

        private void WriteToFile(string fileName, string generatedTemplate)
        {
            
            using (Stream stream = new FileStream(fileName, FileMode.OpenOrCreate))
            {
                stream.Seek(0, SeekOrigin.Begin);
                stream.Write(Encoding.ASCII.GetBytes(generatedTemplate), 0, generatedTemplate.Length);
            }
        }

        private string DefineTypes(string currentlyOpenTabfilePath)
        {
            string stringToReturn = "";
            string className = GetTestClassName(GetHeaderFileName(currentlyOpenTabfilePath)) ;
            stringToReturn += "using"+className+"Types = ::testing::Types<bool>;" +"\n"
            + "TYPED_TEST_SUITE("+className+","+className+"Types);"+"\n";
            return stringToReturn;
        }

        private void AppendToFile(string fileName, string generatedTemplate, int position)
        {
            
            
            using (Stream stream = new FileStream(fileName, FileMode.OpenOrCreate, FileAccess.ReadWrite))
            {
                stream.Seek(position, SeekOrigin.Begin);
                stream.Write(Encoding.ASCII.GetBytes(generatedTemplate), 0, generatedTemplate.Length);
            }              
           
            //dte.ExecuteCommand("File.OpenFile", fileName);
            //dte.ExecuteCommand("Edit.FormatDocument", string.Empty);
        }

        private string getArgs(string v)
        {
            string[] varList = v.Split(',');
            string stringToReturn = "";
            if (v == "")
                return stringToReturn;
            foreach (string arg in varList)
            {
                //string trimmedArg = arg.Trim();
                string[] typeAndVariableName = arg.Trim().Split(' ');
                string initializer = "";

                if (typeAndVariableName[0] == "int" || typeAndVariableName[0] == "DWORD")
                {
                    initializer = "0";

                }
                else if (typeAndVariableName[0] == "string")
                {
                    initializer = "\"\"";

                }
                else if (typeAndVariableName[0] == "wstring")
                {
                    initializer = "L\"\"";

                }

                if (initializer.Length > 0)
                    stringToReturn = stringToReturn + "\t" + arg + "=" + initializer + ';' + '\n';

                else
                    stringToReturn = stringToReturn + "\t" + arg + ';' + '\n';

            }
            return stringToReturn;


        }

        private string GetAssertString()
        {
            string stringToReturn = "";
            stringToReturn += "\t" + "EXPECT_TRUE" + "(2 == 2)" + ";" + "\n";
            stringToReturn += "\t" + "EXPECT_FALSE" + "(2 == 1)" + ";" + "\n";
            stringToReturn += "\t" + "EXPECT_EQUALS" + "(2, 2)" + ";" + "\n";

            return stringToReturn;
        }

        private string GetClassInitialization(string fileName)
        {
            string testClassName = GetTestClassName(GetHeaderFileName(fileName));
            string baseClass = "public testing::Test";
            return "class " + testClassName + ":" + baseClass + "\n";

        }

        private string GetHeaderFileName(string currentlyOpenTabfilePath)
        {
            currentlyOpenTabfilePath = currentlyOpenTabfilePath.Replace(".cpp", ".h");
            string[] fileName = currentlyOpenTabfilePath.Split('\\');
            return fileName[fileName.Length - 1];
        }

        private string GetTestClassName(string className)
        {
            return className.Replace(".h", "") + "Test";
        }
        private string GetCurrentClassName(string className)
        {
            return className.Replace(".h", "");
        }

        private string GetConstructorAndDestructor(string currentlyOpenTabfilePath)
        {

            string stringToReturn = "";
            var testClassName = GetTestClassName(GetHeaderFileName(currentlyOpenTabfilePath));
            stringToReturn = "\n" + "public:" + "\n";
            stringToReturn += testClassName + "()" + "{" + "\n" + "}" + "\n";
            stringToReturn += "~" + testClassName + "()" + "{" + "\n" + "}" + "\n";
            return stringToReturn;
        }

        private string GetHeaders(string currentlyOpenTabfilePath)
        {
            var className = GetHeaderFileName(currentlyOpenTabfilePath);
            string stringToReturn = "";
            stringToReturn += "#include " + "\"pch.h\"" + "\n";
            stringToReturn += "#include " + "\"iostream\"" + "\n";            
            stringToReturn += "#include " + "\"gtest/gtest.h\"" + "\n";
            stringToReturn += "#include " + "\"gmock/gmock.h\"" + "\n";
            stringToReturn += "#include " + "\"gmock/gmock-generated-function-mockers.h\"" + "\n";
            stringToReturn += "#include " + "\"" + className + "\"" + "\n"+"\n";
            return stringToReturn;
        }

        private string GetSourceFileName(string currentlyOpenTabfilePath)
        {
            return GetHeaderFileName(currentlyOpenTabfilePath).Replace(".h", ".cpp");
        }

        private string GetCurrentDirectory(string currentlyOpenTabfilePath)
        {
            return Path.GetDirectoryName(currentlyOpenTabfilePath);
        }

        private string AddNamespace()
        {
            string stringToReturn = "namespace unittest {";
            stringToReturn += "namespace UnitTesters {" +
             "using testing::_;\n" +
            "using testing::A;\n" +
            "using testing::An;" +
            "using testing::AnyNumber;\n" +
            "using testing::Const;\n" +
            "using testing::DoDefault;\n" +
            "using testing::Eq;\n" +
            "using testing::Lt;\n" +
            "using testing::MockFunction;\n" +
            "using testing::Ref;\n" +
            "using testing::Return;\n" +
            "using testing::ReturnRef;\n" +
            "using testing::TypedEq;\n" +
            "\n" +
            "template < typename T >\n" +

        "class TemplatedCopyable\n" +
        "{\n" +
            "public:\n" +
            "TemplatedCopyable() { }\n" +
            "\n" +
            "template<typename U>\n" +
            "TemplatedCopyable(const U& other) {} \n" +
        "};\n";
            return stringToReturn;
        }

    }
}