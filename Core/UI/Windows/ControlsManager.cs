// --------------------------------------------------------------------------------------------------------------------
// <copyright file="ControlsManager.cs" company="B1C Canada Inc.">
//   B1C Canada Inc.
// </copyright>
// <summary>
//   Defines the ControlManager type.
// </summary>
// --------------------------------------------------------------------------------------------------------------------

using B1C.SAP.UI.Adapters;

namespace B1C.SAP.UI.Windows
{
    /// <summary>
    /// The COntrols Manager class
    /// </summary>
    public class ControlManager
    {
        /// <summary>
        /// The form instance
        /// </summary>
        private readonly Form form;

        /// <summary>
        /// Initializes a new instance of the <see cref="ControlManager"/> class.
        /// </summary>
        /// <param name="currentForm">The current form.</param>
        public ControlManager(Form currentForm)
        {
            this.form = currentForm;
        }

        /// <summary>
        /// Matrixes the specified unique id.
        /// </summary>
        /// <param name="uniqueId">The unique id.</param>
        /// <returns>The Matrix Requested</returns>
        public Matrix Matrix(string uniqueId)
        {
            return Adapters.Matrix.Instance(this.form.SapForm, uniqueId);
        }

        /// <summary>
        /// Edits the text.
        /// </summary>
        /// <param name="uniqueId">The unique id.</param>
        /// <returns>The Edit text requested</returns>
        public EditText EditText(string uniqueId)
        {
            return Adapters.EditText.Instance(this.form.SapForm, uniqueId);
        }

        /// <summary>
        /// Buttons the specified unique id.
        /// </summary>
        /// <param name="uniqueId">The unique id.</param>
        /// <returns>The button requested</returns>
        public Button Button(string uniqueId)
        {
            return Adapters.Button.Instance(this.form.SapForm, uniqueId);
        }

        /// <summary>
        /// Statics the text.
        /// </summary>
        /// <param name="uniqueId">The unique id.</param>
        /// <returns>The Static text requested</returns>
        public StaticText StaticText(string uniqueId)
        {
            return Adapters.StaticText.Instance(this.form.SapForm, uniqueId);
        }

        /// <summary>
        /// Comboes the box.
        /// </summary>
        /// <param name="uniqueId">The unique id.</param>
        /// <returns>The Combo box requested</returns>
        public ComboBox ComboBox(string uniqueId)
        {
            return Adapters.ComboBox.Instance(this.form.SapForm, uniqueId);
        }

        /// <summary>
        /// Items the specified unique id.
        /// </summary>
        /// <param name="uniqueId">The unique id.</param>
        /// <returns>The Item Control Adapter</returns>
        public Item Item(string uniqueId)
        {
            return new Item(this.form.SapForm, uniqueId);
        }
    }
}
