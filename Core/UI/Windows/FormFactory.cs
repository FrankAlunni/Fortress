// --------------------------------------------------------------------------------------------------------------------
// <copyright file="FormFactory.cs" company="B1C Canada Inc.">
//   B1C Canada Inc.
// </copyright>
// <summary>
//   The Form Factory Class
// </summary>
// --------------------------------------------------------------------------------------------------------------------

namespace B1C.SAP.UI.Windows
{
    using System;
    using System.Linq.Expressions;
    using Model.Interfaces;

    /// <summary>
    /// The Form Factory Class
    /// </summary>
    public class FormFactory
    {
        /// <summary>
        /// Private Instance of the Add On
        /// </summary>
        private readonly AddOn _addOn;

        /// <summary>
        /// Initializes a new instance of the <see cref="FormFactory"/> class.
        /// </summary>
        /// <param name="addon">The addon.</param>
        public FormFactory(AddOn addon)
        {
            this._addOn = addon;
        }

        /// <summary>
        /// Creates the form.
        /// </summary>
        /// <typeparam name="TFormType">The type of the orm type.</typeparam>
        /// <param name="menu">Index of the menu.</param>
        /// <returns>The Form created</returns>
        public TFormType Create<TFormType>(SAPbouiCOM.MenuItem menu) where TFormType : class
        {
            menu.Activate();
            return this.Create<TFormType>();
        }

        /// <summary>
        /// Creates the form.
        /// </summary>
        /// <typeparam name="TFormType">The type of the form type.</typeparam>
        /// <returns>The Form created</returns>
        public TFormType Create<TFormType>() where TFormType : class
        {
            if (this._addOn.Forms.ContainsKey(this._addOn.Application.Forms.ActiveForm.UniqueID))
            {
                return (TFormType) this._addOn.Forms[this._addOn.Application.Forms.ActiveForm.UniqueID];
            }

            var p = ConstructorInstantiator<TFormType>.Create(this._addOn);

            if (p is IForm)
            {
                ((IForm)p).AddOn.AddForm((IForm)p);
            }

            return p;
        }

        /// <summary>
        /// Creates the form.
        /// </summary>
        /// <typeparam name="TFormType">The type of the orm type.</typeparam>
        /// <param name="form">The SAP form.</param>
        /// <returns>The Form created</returns>
        public TFormType Create<TFormType>(SAPbouiCOM.Form form) where TFormType : class
        {
            if (this._addOn.Forms.ContainsKey(form.UniqueID))
            {
                return (TFormType)this._addOn.Forms[form.UniqueID];
            }

            var p = ConstructorInstantiator<TFormType>.Create(this._addOn, (SAPbouiCOM.FormClass)form);

            if (p is IForm)
            {
                ((IForm)p).AddOn.AddForm((IForm)p);
            }

            return p;
        }

        public TFormType Create<TFormType>(string menuId) where TFormType : class
        {
            this._addOn.Application.Menus.Item(menuId).Activate();
            return this.Create<TFormType>();
        }

        /// <summary>
        /// Creates the form.
        /// </summary>
        /// <typeparam name="TFormType">The type of the form type.</typeparam>
        /// <param name="fileName">Name of the file.</param>
        /// <param name="type">The form type.</param>
        /// <returns>The Form created</returns>
        public TFormType Create<TFormType>(string fileName, string type) where TFormType : class
        {
            var p = ConstructorInstantiator<TFormType>.Create(this._addOn, this._addOn.Location + fileName, type);

            if (p is IForm)
            {
                ((IForm)p).AddOn.AddForm((IForm)p);
            }

            return p;
        }

        /// <summary>
        /// The Constructor instaitiator
        /// </summary>
        /// <typeparam name="TFormType">The Form Type</typeparam>
        private static class ConstructorInstantiator<TFormType>
        {
            /// <summary>
            /// The XML constructor function
            /// </summary>
            private static Func<AddOn, string, string, TFormType> _xmlConstructor;

            /// <summary>
            /// The SAP Active Form constuctor function
            /// </summary>
            private static Func<AddOn, TFormType> _sapActiveFormConstructor;

            /// <summary>
            /// The SAP Form constuctor function
            /// </summary>
            private static Func<AddOn, SAPbouiCOM.FormClass, TFormType> _sapFormConstructor;

            /// <summary>
            /// Creates the specified add on.
            /// </summary>
            /// <param name="addOn">The add on.</param>
            /// <param name="xmlPath">The XML path.</param>
            /// <param name="formType">The form type.</param>
            /// <returns>Create an instance of type T class</returns>
            public static TFormType Create(AddOn addOn, string xmlPath, string formType)
            {
                if (_xmlConstructor == null)
                {
                    _xmlConstructor = XmlFormConstructor();
                }

                return _xmlConstructor(addOn, xmlPath, formType);
            }

            /// <summary>
            /// Creates the specified add on.
            /// </summary>
            /// <param name="addOn">The add on.</param>
            /// <param name="form">The SAP form.</param>
            /// <returns>Create an instance of type T class</returns>
            public static TFormType Create(AddOn addOn, SAPbouiCOM.FormClass form)
            {
                if (_sapFormConstructor == null)
                {
                    _sapFormConstructor = SapFormConstructor();
                }

                return _sapFormConstructor(addOn, form);
            }

            /// <summary>
            /// Creates the specified add on.
            /// </summary>
            /// <param name="addOn">The add on.</param>
            /// <returns>Create an instance of type T class</returns>
            public static TFormType Create(AddOn addOn)
            {
                if (_sapActiveFormConstructor == null)
                {
                    _sapActiveFormConstructor = SapActiveFormConstructor();
                }

                return _sapActiveFormConstructor(addOn);
            }

            /// <summary>
            /// The XML Form Constructor class
            /// </summary>
            /// <returns>
            /// The Constructor os a form build using XML
            /// </returns>
            /// <exception cref="InvalidOperationException">
            /// </exception>
            private static Func<AddOn, string, string, TFormType> XmlFormConstructor()
            {
                var paramAddOn = Expression.Parameter(typeof(AddOn), "addOn");
                var paramXmlPath = Expression.Parameter(typeof(string), "xmlPath");
                var paramType = Expression.Parameter(typeof(string), "formType");

                var ci = typeof(TFormType).GetConstructor(new[] { typeof(AddOn), typeof(string), typeof(string) });
                if (ci == null)
                {
                    throw new InvalidOperationException("Constructor protected ctor(B1C.SAP.UI.AddOn, string, string) not found in the form you are trying to create.");
                }

                var body = Expression.New(ci, paramAddOn, paramXmlPath, paramType);
                return Expression.Lambda<Func<AddOn, string, string, TFormType>>(body, paramAddOn, paramXmlPath, paramType).Compile();
            }

            /// <summary>
            /// The Sap form constructor.
            /// </summary>
            /// <returns>
            /// The Constructor os a form build using SAP Form
            /// </returns>
            /// <exception cref="InvalidOperationException">
            /// </exception>
            private static Func<AddOn, SAPbouiCOM.FormClass, TFormType> SapFormConstructor()
            {
                var paramAddOn = Expression.Parameter(typeof(AddOn), "addOn");
                var paramSapForm = Expression.Parameter(typeof(SAPbouiCOM.FormClass), "form");

                var ci = typeof(TFormType).GetConstructor(new[] { typeof(AddOn), typeof(SAPbouiCOM.FormClass) });
                if (ci == null)
                {
                    throw new InvalidOperationException("Constructor protected ctor(B1C.SAP.UI.AddOn, SAPbouiCOM.Form) not found in the form you are trying to create.");
                }

                var body = Expression.New(ci, paramAddOn, paramSapForm);
                return Expression.Lambda<Func<AddOn, SAPbouiCOM.FormClass, TFormType>>(body, paramAddOn, paramSapForm).Compile();
            }

            /// <summary>
            /// Saps the active form constructor.
            /// </summary>
            /// <returns>
            /// The Constructor os a form build using SAP Form
            /// </returns>
            /// <exception cref="InvalidOperationException">
            /// </exception>
            private static Func<AddOn, TFormType> SapActiveFormConstructor()
            {
                var paramAddOn = Expression.Parameter(typeof(AddOn), "addOn");

                var ci = typeof(TFormType).GetConstructor(new[] { typeof(AddOn) });
                if (ci == null)
                {
                    throw new InvalidOperationException("Constructor protected ctor() not found in the form you are trying to create.");
                }

                var body = Expression.New(ci, paramAddOn);
                return Expression.Lambda<Func<AddOn, TFormType>>(body, paramAddOn).Compile();
            }
        } 
    }
}
