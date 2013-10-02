// -----------------------------------------------------------------------
// <copyright file="Engine.cs" company="B1C Canada Inc.">
//   Copyright (c) B1C Canada Inc. All rights reserved.
// </copyright>
// <author>Frank Alunni</author>
// <summary>
//    Product        : B1C SAP Core Utilities
//    Author Email   : frank.alunni@B1C.com
//    Created Date   : 03/17/2009
//    Prerequisites  : - SAP Business One v.2007a
//                     - SQL Server 2005
//                     - Micorsoft .NET Framework 3.5
// </summary>
//-----------------------------------------------------------------------
namespace B1C.Utility.Script
{
    #region Using Directive

    using System;
    using System.CodeDom.Compiler;
    using System.Reflection;
    using System.Text;

    using Microsoft.CSharp;
    using Microsoft.VisualBasic;

    #endregion

    /// <summary>
    /// Instance of the Engine class.
    /// </summary>
    /// <author>Frank Alunni</author>
    public class Engine : MarshalByRefObject 
    {
        #region Fields
        /// <summary>
        /// Teh assembly generated
        /// </summary>
        private Assembly assembly;

        /// <summary>
        /// Dom Compiler
        /// </summary>
        private CodeDomProvider compiler;

        /// <summary>
        /// The evaluator object
        /// </summary>
        private object evaluator;

        /// <summary>
        /// The evaluator type
        /// </summary>
        private Type evaluatorType;

        /// <summary>
        /// The compilation paramter list
        /// </summary>
        private CompilerParameters parameters;

        /// <summary>
        /// The compilation results
        /// </summary>
        private CompilerResults results;

        /// <summary>
        /// Template function
        /// </summary>
        private string templateFunction;

        /// <summary>
        /// Template module
        /// </summary>
        private string templateModule;

        /// <summary>
        /// Template variable
        /// </summary>
        private string templateVariable;

        /// <summary>
        /// The variables
        /// </summary>
        private string variables;

        #endregion Fields

        #region Constructor

        /// <summary>
        /// Initializes a new instance of the <see cref="Engine"/> class.
        /// </summary>
        public Engine() : this(LanguageUsed.VBasic)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="Engine"/> class.
        /// </summary>
        /// <param name="language">
        /// The language.
        /// </param>
        public Engine(LanguageUsed language)
        {
            this.Language = language;

            // Reset Variables
            this.Code = string.Empty;
            this.variables = string.Empty;
            this.CodeSnip = string.Empty;
            this.SetTemplates();
        }

        #endregion Constructor

        #region Enums

        /// <summary>
        /// Possible Languages
        /// </summary>
        public enum LanguageUsed
        {
            /// <summary>
            /// Language vb.NET option
            /// </summary>
            VBasic,

            /// <summary>
            /// Language C# option
            /// </summary>
            CSharp
        }

        #endregion

        #region Properties

        /// <summary>
        /// Gets or sets the name of the assembly file.
        /// </summary>
        /// <value>The name of the assembly file.</value>
        public string AssemblyFileName
        {
            get;
            set;
        }

        /// <summary>
        /// Gets or sets a value indicating whether [run optimized].
        /// </summary>
        /// <value><c>true</c> if [run optimized]; otherwise, <c>false</c>.</value>
        public bool RunOptimized
        {
            get;
            set;
        }

        /// <summary>
        /// Gets or sets the code.
        /// </summary>
        /// <value>The generated code.</value>
        public string Code
        {
            get;

            set;
        }

        /// <summary>
        /// Gets or sets CodeSnip.
        /// </summary>
        /// <value>
        /// The code snip.
        /// </value>
        public string CodeSnip
        {
            get;

            set;
        }

        /// <summary>
        /// Gets or sets Language.
        /// </summary>
        /// <value>
        /// The language.
        /// </value>
        public LanguageUsed Language
        {
            get;

            set;
        }

        /// <summary>
        /// Gets or sets Source.
        /// </summary>
        /// <value>
        /// The source.
        /// </value>
        public string Source
        {
            get;

            set;
        }

        /// <summary>
        /// Gets or sets the messages.
        /// </summary>
        /// <value>The messages.</value>
        public string[] Messages
        {
            get;

            set;
        }

        #endregion Properties

        #region Public Methods

        /// <summary>
        /// Adds the code snip.
        /// </summary>
        /// <param name="codeToAdd">The code to add.</param>
        public void AddCodeSnip(string codeToAdd)
        {
            var sb = new StringBuilder();
            if (this.CodeSnip != string.Empty)
            {
                sb.AppendLine(this.CodeSnip);
            }

            sb.AppendLine("        " + codeToAdd);

            this.CodeSnip = sb.ToString();
        }

        /// <summary>
        /// Adds the function.
        /// </summary>
        /// <param name="variableName">Name of the variable.</param>
        /// <param name="variableType">Type of the variable.</param>
        /// <param name="codeToAdd">The code to add.</param>
        public void AddFunction(string variableName, string variableType, string codeToAdd)
        {
            if (codeToAdd.StartsWith("Result"))
            {
                this.variables += string.Format(this.templateFunction, variableName, variableType, codeToAdd);
            }
            else
            {
                this.variables += string.Format(this.templateFunction, variableName, variableType, "Result = " + codeToAdd);
            }

            // System.Diagnostics.Debug.Write(this.variables);
        }

        /// <summary>
        /// Adds the variable.
        /// </summary>
        /// <param name="variableName">Name of the variable.</param>
        /// <param name="variableType">Type of the variable.</param>
        public void AddVariable(string variableName, string variableType)
        {
            this.variables += string.Format(this.templateVariable, variableName, variableType);

            // System.Diagnostics.Debug.Write(this.variables);
        }

        /// <summary>
        /// Clean the formula folder. this instance.
        /// </summary>
        /// <returns>True if the folder is successfully cleaned</returns>
        public bool Clean()
        {
            Assembly runningAssembly = Assembly.GetExecutingAssembly();
            if (runningAssembly.Location != null)
            {
                string dllLocation = runningAssembly.Location.Replace(runningAssembly.ManifestModule.Name, "Formulas\\");
                if (System.IO.Directory.Exists(dllLocation))
                {
                    System.IO.Directory.Delete(dllLocation, true);
                }
            }

            return true;
        }

        /// <summary>
        /// Compiles this instance.
        /// </summary>
        /// <returns>True if the code compiles successfully</returns>
        public bool Compile()
        {
            /**********************************************************
             *  MODULE 
             * 0 ==> Variables
             * 1 ==> Code
             * 2 ==> Code Snips
             *********************************************************/
            this.Source = string.Format(this.templateModule, this.variables.Trim(), this.Code.Trim(), this.CodeSnip.Trim());

            // System.Diagnostics.Debug.Write(this.Source);
            this.parameters = new CompilerParameters 
            { 
               GenerateInMemory = false,
               GenerateExecutable = false,
               IncludeDebugInformation = false
            };

            if (this.RunOptimized)
            {
                // Save dll to a specific folder
                Assembly runningAssembly = Assembly.GetExecutingAssembly();
                string dllLocation = string.Empty;
                if (runningAssembly.Location != null)
                {
                    dllLocation = runningAssembly.Location.Replace(runningAssembly.ManifestModule.Name, "Formulas\\");
                }

                if (!string.IsNullOrEmpty(this.AssemblyFileName))
                {
                    string dllFilePath = string.Format("{0}{1}", dllLocation, this.AssemblyFileName);

                    // Check it the assemply exists if so load it
                    if (System.IO.File.Exists(dllFilePath))
                    {
                        this.assembly = Assembly.LoadFile(dllFilePath);
                        this.evaluatorType = this.assembly.GetType("UserScript.RunScript");
                        this.evaluator = Activator.CreateInstance(this.evaluatorType);
                        return true;
                    }

                    if (!System.IO.Directory.Exists(dllLocation))
                    {
                        System.IO.Directory.CreateDirectory(dllLocation);
                    }

                    this.parameters.OutputAssembly = dllFilePath;
                    this.parameters.TempFiles = new TempFileCollection(dllLocation, true);
                }
            }

            switch (this.Language)
            {
                case LanguageUsed.CSharp:
                    this.compiler = new CSharpCodeProvider();
                    break;
                default: // VB.NET
                    this.compiler = new VBCodeProvider();
                    break;
            }

            this.results = this.compiler.CompileAssemblyFromSource(this.parameters, this.Source);

            // Check for compile errors / warnings
            if (this.results.Errors.HasErrors || this.results.Errors.HasWarnings)
            {
                this.Messages = new string[this.results.Errors.Count];
                for (int i = 0; i < this.results.Errors.Count; i++)
                {
                    this.Messages[i] = this.results.Errors[i].ToString();
                }

                return false;
            }

            if (this.RunOptimized)
            {
                // Keep source file name
                if (!string.IsNullOrEmpty(this.AssemblyFileName))
                {
                    string sourceFileOriginalName = string.Format("{0}.0.{1}", this.results.TempFiles.BasePath, this.compiler.FileExtension);
                    string newFileName = sourceFileOriginalName.Substring(0, sourceFileOriginalName.LastIndexOf(@"\")) + @"\" + this.AssemblyFileName.Replace(".dll", ".vb");

                    System.IO.File.Copy(sourceFileOriginalName, newFileName);
                    System.IO.File.Delete(sourceFileOriginalName);
                    System.IO.File.Delete(this.results.TempFiles.BasePath + ".err");
                    System.IO.File.Delete(this.results.TempFiles.BasePath + ".cmdline");
                    System.IO.File.Delete(this.results.TempFiles.BasePath + ".out");
                    System.IO.File.Delete(this.results.TempFiles.BasePath + ".tmp");
                }
            }

            this.Messages = null;
            this.assembly = this.results.CompiledAssembly;
            this.evaluatorType = this.assembly.GetType("UserScript.RunScript");
            this.evaluator = Activator.CreateInstance(this.evaluatorType);
            this.compiler.Dispose();
            GC.Collect();
            return true;
        }

        /// <summary>
        /// Evaluates this instance.
        /// </summary>
        /// <returns>The result of the formula</returns>
        public double Evaluate()
        {
            object o = this.evaluatorType.InvokeMember("Eval", BindingFlags.InvokeMethod, null, this.evaluator, new object[] { });
            string s = o.ToString();
            return double.Parse(s);
        }

        /// <summary>
        /// Evaluates the property.
        /// </summary>
        /// <param name="propertyName">Name of the property.</param>
        /// <returns>The result of the property</returns>
        public double EvaluateProperty(string propertyName)
        {
            try
            {
                object o = this.evaluatorType.InvokeMember(propertyName, BindingFlags.GetProperty, null, this.evaluator, new object[] { });

                string s = o.ToString();
                return double.Parse(s);
            }
            catch
            {
                return -1;                
            }
        }

        #region Variable Methods

        /// <summary>
        /// Gets the variable.
        /// </summary>
        /// <param name="variableName">Name of the variable.</param>
        /// <param name="retValue">The ret value.</param>
        public void GetVariable(string variableName, out double retValue)
        {
            object o = this.evaluatorType.InvokeMember(
                    "Get" + variableName, BindingFlags.InvokeMethod, null, this.evaluator, new object[] { });
            string s = o.ToString();
            retValue = double.Parse(s);
        }

        /// <summary>
        /// Gets the variable.
        /// </summary>
        /// <param name="variableName">Name of the variable.</param>
        /// <param name="retValue">The ret value.</param>
        public void GetVariable(string variableName, out int retValue)
        {
            object o = this.evaluatorType.InvokeMember(
                    "Get" + variableName, BindingFlags.InvokeMethod, null, this.evaluator, new object[] { });
            string s = o.ToString();
            retValue = int.Parse(s);
        }

        /// <summary>
        /// Gets the variable.
        /// </summary>
        /// <param name="variableName">Name of the variable.</param>
        /// <param name="retValue">The ret value.</param>
        public void GetVariable(string variableName, out string retValue)
        {
            object o = this.evaluatorType.InvokeMember(
                    string.Format("Get{0}", variableName), BindingFlags.InvokeMethod, null, this.evaluator, new object[] { });
            string s = o.ToString();
            retValue = s;
        }

        /// <summary>
        /// Sets the variable.
        /// </summary>
        /// <param name="variableName">Name of the variable.</param>
        /// <param name="value">The value.</param>
        public void SetVariable(string variableName, double value)
        {
            this.evaluatorType.InvokeMember("Set" + variableName, BindingFlags.InvokeMethod, null, this.evaluator, new object[] { value });
        }

        /// <summary>
        /// Sets the variable.
        /// </summary>
        /// <param name="variableName">The variable name.</param>
        /// <param name="value">The value.</param>
        public void SetVariable(string variableName, int value)
        {
            this.evaluatorType.InvokeMember("Set" + variableName, BindingFlags.InvokeMethod, null, this.evaluator, new object[] { value });
        }

        /// <summary>
        /// Sets the variable.
        /// </summary>
        /// <param name="variableName">The variable name.</param>
        /// <param name="value">The value.</param>
        public void SetVariable(string variableName, string value)
        {
           this.evaluatorType.InvokeMember("Set" + variableName, BindingFlags.InvokeMethod, null, this.evaluator, new object[] { value });
        }

        #endregion Variable Methods

        #endregion Public Methods

        #region Private Methods

        /// <summary>
        /// Sets the templates.
        /// </summary>
        private void SetTemplates()
        {
            // Build Variable Template
            var sb = new StringBuilder();
            switch (this.Language)
            {
                case LanguageUsed.CSharp:

                    /**********************************************************
                     *  VARIABLE TEMPLATE 
                     * 0 ==> Variable Name
                     * 1 ==> Variable Type
                     *********************************************************/
                    sb.AppendLine("     {1} {0} = 0;                        ");
                    sb.AppendLine("     public void Set{0}({1} x)           ");
                    sb.AppendLine("     {                                   ");
                    sb.AppendLine("         {0} = x;                        ");
                    sb.AppendLine("     }                                   ");
                    sb.AppendLine("     public {1} Get{0}()                 ");
                    sb.AppendLine("     {                                   ");
                    sb.AppendLine("         return {0};                     ");
                    sb.AppendLine("     }                                   ");
                    sb.AppendLine(string.Empty);
                    this.templateVariable = sb.ToString();

                    /**********************************************************
                     *  FUNCTION TEMPLATE 
                     * 0 ==> Variable Name
                     * 1 ==> Variable Type
                     * 2 ==> Function Code
                     *********************************************************/
                    sb = new StringBuilder();
                    sb.AppendLine("     {1} {0} = 0;                        ");
                    sb.AppendLine("     public {1} Get{0}()                 ");
                    sb.AppendLine("     {                                   ");
                    sb.AppendLine("         double Return = double.NaN;     ");
                    sb.AppendLine("         {2}                             ");
                    sb.AppendLine("         {0} = Return;                   ");
                    sb.AppendLine("         return {0};                     ");
                    sb.AppendLine("     }                                   ");
                    sb.AppendLine(string.Empty);
                    this.templateFunction = sb.ToString();

                    /**********************************************************
                     *  MODULE TEMPLATE 
                     * 0 ==> Variables
                     * 1 ==> Code
                     * 2 ==> Code Snips
                     *********************************************************/
                    sb = new StringBuilder();
                    sb.AppendLine("using System;                            ");
                    sb.AppendLine("namespace UserScript                     ");
                    sb.AppendLine("{                                        ");
                    sb.AppendLine(" public class RunScript                  ");

                    sb.AppendLine(" {                                       ");
                    sb.AppendLine("{0}                                      ");
                    sb.AppendLine("{2}                                      ");
                    sb.AppendLine("                                         ");
                    sb.AppendLine("     public double Eval()                ");
                    sb.AppendLine("     {                                   ");
                    sb.AppendLine("         double Result = {0}.NaN;        ");
                    sb.AppendLine("         {1}                             ");
                    sb.AppendLine("         return Result;                  ");
                    sb.AppendLine("     }                                   ");
                    sb.AppendLine(" }                                       ");
                    sb.AppendLine("}                                        ");
                    this.templateModule = sb.ToString();
                    break;

                default: // Visual Basic

                    /**********************************************************
                     *  VARIABLE TEMPLATE 
                     * 0 ==> Variable Name
                     * 1 ==> Variable Type
                     *********************************************************/
                    sb.AppendLine(string.Empty);
                    sb.AppendLine("        Dim {0} As {1}                      ");
                    sb.AppendLine("        Public Sub Set{0}(AVal As {1})      ");
                    sb.AppendLine("            {0} = AVal                      ");
                    sb.AppendLine("        End Sub                             ");
                    sb.AppendLine("        Public Function Get{0} As {1}       ");
                    sb.AppendLine("            Return {0}                      ");
                    sb.AppendLine("        End Function                        ");
                    this.templateVariable = sb.ToString();

                    /**********************************************************
                     *  FUNCTION TEMPLATE 
                     * 0 ==> Variable Name
                     * 1 ==> Variable Type
                     * 2 ==> Function Code
                     *********************************************************/
                    sb = new StringBuilder();

                    sb.AppendLine(string.Empty);
                    sb.AppendLine("        Public ReadOnly Property {0} As {1} ");
                    sb.AppendLine("            Get                             ");
                    sb.AppendLine("                Dim Result As {1}           ");
                    sb.AppendLine("                Result = {1}.NaN            ");
                    sb.AppendLine("                {2}                         ");
                    sb.AppendLine("                Return Result               ");
                    sb.AppendLine("            End Get                         ");
                    sb.AppendLine("        End Property                        ");
                    this.templateFunction = sb.ToString();

                    /**********************************************************
                     *  MODULE TEMPLATE 
                     * 0 ==> Variables
                     * 1 ==> Code
                     * 2 ==> Code Snips
                     *********************************************************/
                    sb = new StringBuilder();
                    sb.AppendLine("Imports System                           ");
                    sb.AppendLine("Namespace UserScript                     ");
                    sb.AppendLine("    Public Class RunScript               ");
                    sb.AppendLine("        {0}                              ");
                    sb.AppendLine("                                         ");
                    sb.AppendLine("        {2}                              ");
                    sb.AppendLine("                                         ");
                    sb.AppendLine("         Public Function Eval() As Double");
                    sb.AppendLine("             Dim Result As Double        ");
                    sb.AppendLine("             {1}                         ");
                    sb.AppendLine("             Return Result               ");
                    sb.AppendLine("         End Function                    ");
                    sb.AppendLine("    End Class                            ");
                    sb.AppendLine("End Namespace                            ");
                    this.templateModule = sb.ToString();
                    break;
            }
        }

        #endregion Private Methods
    }
}