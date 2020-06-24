using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using BomDisassembly.Models;
using SAPbouiCOM.Framework;

namespace BomDisassembly.Forms
{
    [FormAttribute("BomDisassembly.Forms.DissassemblyBomList", "Forms/DissassemblyBomList.b1f")]
    class DissassemblyBomList : UserFormBase
    {
        public DissassemblyBomList()
        {
        }

        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.Grid0 = ((SAPbouiCOM.Grid)(this.GetItem("Item_0").Specific));
            this.Grid0.DoubleClickAfter += new SAPbouiCOM._IGridEvents_DoubleClickAfterEventHandler(this.Grid0_DoubleClickAfter);
            this.Grid0.ClickAfter += new SAPbouiCOM._IGridEvents_ClickAfterEventHandler(this.Grid0_ClickAfter);
            this.OnCustomInitialize();

        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
        }

        private SAPbouiCOM.Grid Grid0;

        private void OnCustomInitialize()
        {
            Refresh();
            BomDissasembly.RefreshListBom += Refresh;
        }

        private void Grid0_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            if (pVal.Row == -1)
            {
                return;
            }
            Grid0.Rows.SelectedRows.Clear();
            Grid0.Rows.SelectedRows.Add(pVal.Row);
        }

        public void Refresh()
        {
            Grid0.DataTable.ExecuteQuery(DiManager.QueryHanaTransalte("select ROW_NUMBER() OVER (ORDER BY Code) as [#], U_ItemCode as [Parent Item], U_ItemName as [Description],  convert(varchar,convert(datetime, U_PostingDate, 126), 1) as [Posting Date], U_ItemTotalCost as [Total Cost], U_UOM as [Unit Of Measure],U_WhsCode as [Warehouse Code], U_WhsName as [Warehouse Name], U_PlannedQuantity as [Planned Quantity], U_Comment as [Comment], Code from [@RSM_BOMD_HEADER] order by Code desc"));
        }

        private void Grid0_DoubleClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            if (pVal.Row == -1)
            {
                return;
            }

            string code = Grid0.DataTable.GetValue("Code", pVal.Row).ToString();
            DisassemblyBom bomForm = new DisassemblyBom(code);
            bomForm.Show();
        }
    }
}
