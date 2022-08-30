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

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("4604f7c9-dbe1-4db8-a82c-148761b330b2");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly AsyncPackage package;

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
            int index1 = entireLine.IndexOf('(') + 1;
            int index2 = entireLine.IndexOf(')');
            if (index2 < index1)
            {
                ShowMessage("Select a valid method name for which you want to create a test");
                return;
            }
            string args = getArgs(entireLine.Substring(index1, index2 - index1));
            string generatedTest =
                "TYPED_TEST" + "(" + GetTestClassName(GetHeaderFileName(currentlyOpenTabfilePath)) + "," + "Should" + methodName + ")" +
                "{" + "\n"
                + args + GetAssertString() +
                "}";

            string absoluteFilePath = "";
            SettingsManager settingsManager = new ShellSettingsManager(ServiceProvider);
            WritableSettingsStore userSettingsStore = settingsManager.GetWritableSettingsStore(SettingsScope.UserSettings);

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
                absoluteFilePath = GetCurrentDirectory(currentlyOpenTabfilePath);
            }

            absoluteFilePath = absoluteFilePath + GetSourceFileName(currentlyOpenTabfilePath);
            string fileName = absoluteFilePath.Replace(".cpp", "Test.cpp");

            if (isFirstTimeForTheFile)
                WriteToFile(absoluteFilePath, fileName, generatedTest);
            else
                AppendToFile(absoluteFilePath, fileName, generatedTest,dte);

            dte.ExecuteCommand("File.OpenFile", fileName);
            dte.ExecuteCommand("Edit.FormatDocument", string.Empty);
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

        private void WriteToFile(string absoluteFilePath, string fileName, string generatedTest)
        {
            generatedTest = GetHeaders(absoluteFilePath) + "\n" + AddNamespace() + "\n" +"\n"+ DefineTypes(absoluteFilePath) + "\n"+ "\n"+ GetClassInitialization(absoluteFilePath) + "{" + GetConstructorAndDestructor(absoluteFilePath) + "};" + "\n" + generatedTest + "\n" + "}" + "\n";
            using (Stream stream = new FileStream(fileName, FileMode.OpenOrCreate))
            {
                stream.Seek(0, SeekOrigin.Begin);
                stream.Write(Encoding.ASCII.GetBytes(generatedTest), 0, generatedTest.Length);
            }
        }

        private string DefineTypes(string currentlyOpenTabfilePath)
        {
            string stringToReturn = "";
            string className = GetTestClassName(GetHeaderFileName(currentlyOpenTabfilePath)) ;
            stringToReturn += "using"+className+"Types = ::testing::Types<bool>;" +"\n"
            + "TYPED_TEST_SUITE("+className+className+"Types);";
            return stringToReturn;
        }

        private void AppendToFile(string absoluteFilePath, string fileName, string generatedTest, DTE2 dte)
        {
            generatedTest = "\n" + generatedTest;
            generatedTest += "\n"+"}";
            string text = File.ReadAllText(fileName);
            //var lineCount = File.ReadLines(fileName).Count();
            using (Stream stream = new FileStream(fileName, FileMode.OpenOrCreate, FileAccess.ReadWrite))
            {
                int len = text.Length;
                stream.Seek(text.Length-2, SeekOrigin.Begin);
                stream.Write(Encoding.ASCII.GetBytes(generatedTest), 0, generatedTest.Length);
            }
            //dte.ExecuteCommand("File.OpenFile", fileName);
            //dte.ExecuteCommand("Edit.FormatDocument", string.Empty);
        }

        private string getArgs(string v)
        {
            string[] varList = v.Split(',');
            string stringToReturn = "";

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
                    stringToReturn = stringToReturn + "\t" + typeAndVariableName[0] + " " + typeAndVariableName[1] + "=" + initializer + ';' + '\n';

                else
                    stringToReturn = stringToReturn + "\t" + typeAndVariableName[0] + " " + typeAndVariableName[1] + ';' + '\n';

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
            stringToReturn += "#include " + "\"gmock/gmock.h\"" + "\n" + "\n";
            stringToReturn += "#include " + "\"" + className + "\"" + "\n";
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
            return "namespace unittest {";
        }


    }
}