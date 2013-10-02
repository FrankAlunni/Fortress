//-----------------------------------------------------------------------
// <copyright file="Program.cs" company="Fortress Technology Inc.">
//     Copyright (c) Fortress Technology Inc. All rights reserved.
// </copyright>
// <author>Frank Alunni</author>
// <email>frank.alunni@b1computing.com</email>
//-----------------------------------------------------------------------

namespace B1C.SAP.Addons.Configurator
{
    #region Using Directive(s)

    using System;
    using System.Windows.Forms;

    #endregion Using Directive(s)

    internal static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        private static void Main()
        {
            new AddOn("B1C", "PC") { Name = Application.ProductName };

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false); 
            Application.Run();
        }
    }
}
