//-----------------------------------------------------------------------
// <copyright file="CopyConfigurator.cs" company="Fortress Technology Inc.">
//     Copyright (c) Fortress Technology Inc. All rights reserved.
// </copyright>
// <author>Frank Alunni</author>
// <email>frank.alunni@b1computing.com</email>
//-----------------------------------------------------------------------

namespace B1C.SAP.Addons.Configurator.Forms
{
    #region Usings
    using UI.Model.EventArguments;
    using System;
    using DI.BusinessAdapters;
    using DI.Business;
    using Fortress.DI.BusinessAdapters.Sales;
    using B1C.SAP.DI.Helpers;
    #endregion Usings

    public class CopyConfigurator : Window.FortressForm 
    {
        #region Constuctors
        /// <summary>
        /// Initializes a new instance of the <see cref="CopyConfigurator"/> class.
        /// </summary>
        /// <param name="addOn">The add on.</param>
        /// <param name="xmlPath">The XML path.</param>
        /// <param name="formType">Type of the form.</param>
        public CopyConfigurator(UI.AddOn addOn, string xmlPath, string formType) : base(addOn, xmlPath, formType) { }
        #endregion

        #region Methods
        /// <summary>
        /// Initializes this instance.
        /// </summary>
        public override void Initialize()
        {
            this.ButtonClick += this.OnButtonClick;
            this.AfterArgumentsSet += this.OnAfterArgumentsSet;
            this.ChooseFromList += this.OnChooseFromList;
            base.Initialize();
        }

        /// <summary>
        /// Performs application-defined tasks associated with freeing, releasing, or resetting unmanaged resources.
        /// </summary>
        public override void Dispose()
        {
            this.ButtonClick -= this.OnButtonClick;
            this.AfterArgumentsSet -= this.OnAfterArgumentsSet;
            this.ChooseFromList -= this.OnChooseFromList;
            base.Dispose();
        }

        private void OnChooseFromList(object sender, ChooseFromListEventArgs e)
        {
            var sel = e.SelectedObjects;
            var edit = this.ControlManager.EditText("soT");

            edit.Value = sel.GetValue("DocNum", 0).ToString();

            string query = "SELECT VisOrder + 1 as VisOrder, T1.DocEntry, LineNum, ItemCode FROM RDR1 T0 JOIN ORDR T1 ON T0.DocEntry = T1.DocEntry ";
            if (this.IsConveyor())
            {
                query += " JOIN [@XX_CONVBLT] T2 ON T1.DocEntry = T2.U_XX_OrderNo AND T0.LineNum = T2.U_XX_OrdrLnNo ";
            }
            else if (this.IsMetalDetector())
            {
                query += "JOIN [@XX_METDET] T2 ON T1.DocEntry = T2.U_XX_OrderNo AND T0.LineNum = T2.U_XX_OrdrLnNo ";
            }
            query += string.Format(" WHERE T1.DocNum = {0} AND T0.ItemCode like '{1}%' ", edit.Value, this.ItemCode.Substring(0, 1));

            // Reload matrix
            var matrix = this.ControlManager.Matrix("mtx_0");
            var dt = this.DataSource.UserTable("ordLnDT");
            dt.ExecuteQuery(query);

            matrix.BaseObject.LoadFromDataSource();

        }


        /// <summary>
        /// Called when [after arguments set].
        /// </summary>
        /// <param name="sender">The sender.</param>
        private void OnAfterArgumentsSet(object sender)
        {
            // AddChooseFromList();
            // var edit = this.ControlManager.EditText("soT");
            // edit.BaseObject.ChooseFromListUID = "UDCFL1";
            // edit.BaseObject.ChooseFromListAlias = "DocNum";            


            //string query = "SELECT CAST(T0.DocEntry as varchar(max)) + '-' + CAST(T0.LineNum as varchar(max)),";
            //query += "RIGHT('000000' + CAST(T3.DocNum as varchar(max)), 6) + '-' +";
            //query += "RIGHT('00' + CAST((T0.VisOrder + 1) as varchar(max)), 2) + ' [' + ";
            //query += "T0.ItemCode + ']' as Descr ";
            //query += "FROM RDR1 T0 WITH (NOLOCK) INNER JOIN ORDR T3 WITH (NOLOCK) on T0.DocEntry = T3.DOcEntry  ";
            //if (this.IsConveyor())
            //{
            //    query += "JOIN [@XX_CONVBLT] T1 ON T0.DocEntry = T1.U_XX_OrderNo AND T0.LineNum = T1.U_XX_OrdrLnNo ";
            //}
            //else if (this.IsMetalDetector())
            //{
            //    query += "JOIN [@XX_METDET] T1 ON T0.DocEntry = T1.U_XX_OrderNo AND T0.LineNum = T1.U_XX_OrdrLnNo ";
            //}
            //query += "WHERE T0.ItemCode like '" + this.ItemCode.Substring(0,1) +  "%' ";
            //query += "ORDER BY T0.DocEntry, T0.VisOrder";
            
            //var combo = this.ControlManager.ComboBox("soItemC");

            
            //using(var rs = new RecordsetAdapter(this.AddOn.Company, query))
            //{
            //    while (!rs.EoF)
            //    {
            //        combo.BaseObject.ValidValues.Add(rs.FieldValue(0).ToString(), rs.FieldValue(1).ToString());
            //        rs.MoveNext();
            //    }
            //}


        }


        private void AddChooseFromList()
        {
            // Choose From List Collection object to store the collection
            SAPbouiCOM.ChooseFromListCollection oCFLs;

            // Getting the Form Choose From Lists 
            oCFLs = this.SapForm.ChooseFromLists;

            // Choose From List object to store the ChooseFromList
            SAPbouiCOM.ChooseFromList oCFL;

            // ChooseFromListCreationParams to create the parameters for the CFL
            SAPbouiCOM.ChooseFromListCreationParams oCFLCreationParams;
            oCFLCreationParams = (SAPbouiCOM.ChooseFromListCreationParams)this.AddOn.Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams);
            oCFLCreationParams.MultiSelection = false;
            oCFLCreationParams.ObjectType = "17";
            oCFLCreationParams.UniqueID = "UDCFL1";
            oCFL = oCFLs.Add(oCFLCreationParams);

            // Adding Conditions to CFL
            //oCons = oCFL.GetConditions();
            //oCon = oCons.Add();
            //oCon.Alias = "U_HasPSC";
            //oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
            //oCon.CondVal = "Y";
            //oCFL.SetConditions(oCons);
        }


        /// <summary>
        /// Called when [button click].
        /// </summary>
        /// <param name="sender">The sender.</param>
        /// <param name="e">The <see cref="B1C.SAP.UI.Model.EventArguments.ItemEventArgs"/> instance containing the event data.</param>
        private void OnButtonClick(object sender, ItemEventArgs e)
        {
            switch (e.ItemId)
            {
                case "btnCopy":
                    try
                    {

                        var matrix = this.ControlManager.Matrix("mtx_0");

                        if (matrix.SelectedRow() < 0)
                        {
                            this.AddOn.StatusBar.ErrorMessage = "Please select the line to copy the configuration from.";
                            return;
                        }

                        string fromOrder = matrix[matrix.SelectedRow(), "V_2"];
                        string fromLine = matrix[matrix.SelectedRow(), "V_1"];

                        var order = new FortressSalesOrderAdapter(this.AddOn.Company, int.Parse(fromOrder));

                        order.DuplicateConfiguration(int.Parse(fromLine), int.Parse(this.OrderId), int.Parse(this.OrderLine));

                        this.AddOn.StatusBar.Message = "Configuration has being copied successfully.";
                        this.ControlManager.Button("2").Focus();
                    }
                    catch (Exception ex)
                    {
                        this.AddOn.HandleException(ex);
                    }

                    break;
            }
        }

        /// <summary>
        /// Determines whether this instance is a conveyor.
        /// </summary>
        /// <returns>
        ///   <c>true</c> if this instance is conveyor; otherwise, <c>false</c>.
        /// </returns>
        private bool IsConveyor()
        {
            string prefix;
            ConfigurationHelper.GlobalConfiguration.Load(this.AddOn.Company, "ConvPref", out prefix);

            if (!this.ItemCode.StartsWith(prefix))
            {
                return false;
            }

            return true;
        }

        /// <summary>
        /// Determines whether this instance is a conveyor.
        /// </summary>
        /// <returns>
        ///   <c>true</c> if this instance is conveyor; otherwise, <c>false</c>.
        /// </returns>
        private bool IsMetalDetector()
        {
            string prefix;
            ConfigurationHelper.GlobalConfiguration.Load(this.AddOn.Company, "DetPref", out prefix);

            if (!this.ItemCode.StartsWith(prefix))
            {
                return false;
            }

            return true;
        }
        #endregion Methods

    }
}
