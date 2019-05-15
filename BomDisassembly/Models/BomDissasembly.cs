using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using SAPbobsCOM;
using SAPbouiCOM;
using Application = SAPbouiCOM.Framework.Application;

namespace BomDisassembly.Models
{
    class BomDissasembly
    {
        public string Code { get; set; }
        public int BplId { get; set; }
        public string ItemCode { get; set; }
        public string GoodsIssueDocEntry { get; set; }
        public string GoodsIssueDocNum { get; set; }
        public string GoodsReceiptDocNum { get; set; }
        public string GoodsReceiptDocEntry { get; set; }
        public string ItemName { get; set; }
        public string Uom { get; set; }
        public string WhsCode { get; set; }
        public string WhsName { get; set; }
        public string Comment { get; set; }
        public double PlannedQuantity { get; set; }
        public double ItemCost { get; set; }
        public double ItemTotalCost { get; set; }
        public DateTime PostingDate { get; set; }
        public string AccountCode { get; set; }
        public double DiffBetweenCosts { get; set; }


        public IEnumerable<BomDissasemblyRow> BomDissasemblyRows { get; set; }

        public void AddHeader()
        {
            Thread.CurrentThread.CurrentCulture = CultureInfo.InvariantCulture;

            if (!DiManager.Company.InTransaction)
            {
                DiManager.Company.StartTransaction();
            }

            UserTable rsmBomHeader = DiManager.Company.UserTables.Item("RSM_BOMD_HEADER");

            rsmBomHeader.UserFields.Fields.Item("U_ItemCode").Value = ItemCode;
            rsmBomHeader.UserFields.Fields.Item("U_GoodsIssueDocEntry").Value = GoodsIssueDocEntry;
            rsmBomHeader.UserFields.Fields.Item("U_GoodsIssueDocNum").Value = GoodsIssueDocNum;
            rsmBomHeader.UserFields.Fields.Item("U_GoodsReceiptDocNum").Value = GoodsReceiptDocNum ?? "";
            rsmBomHeader.UserFields.Fields.Item("U_GoodsReceiptDocEntry").Value = GoodsReceiptDocEntry ?? "";
            rsmBomHeader.UserFields.Fields.Item("U_ItemName").Value = ItemName;
            rsmBomHeader.UserFields.Fields.Item("U_Uom").Value = Uom;
            rsmBomHeader.UserFields.Fields.Item("U_WhsCode").Value = WhsCode;
            rsmBomHeader.UserFields.Fields.Item("U_WhsName").Value = WhsName;
            rsmBomHeader.UserFields.Fields.Item("U_ItemCost").Value = ItemCost.ToString(CultureInfo.InvariantCulture);
            rsmBomHeader.UserFields.Fields.Item("U_PlannedQuantity").Value = PlannedQuantity.ToString(CultureInfo.InvariantCulture);
            rsmBomHeader.UserFields.Fields.Item("U_ItemTotalCost").Value = ItemTotalCost.ToString(CultureInfo.InvariantCulture);
            rsmBomHeader.UserFields.Fields.Item("U_WhsCode").Value = WhsCode;
            rsmBomHeader.UserFields.Fields.Item("U_PostingDate").Value = PostingDate.ToString("s");
            rsmBomHeader.UserFields.Fields.Item("U_Comment").Value = Comment;
            rsmBomHeader.UserFields.Fields.Item("U_AccountCode").Value = AccountCode;

            int res = rsmBomHeader.Add();

            if (res != 0)
            {
                Application.SBO_Application.SetStatusBarMessage(DiManager.Company.GetLastErrorDescription(),
                    BoMessageTime.bmt_Short, true);

                if (DiManager.Company.InTransaction)
                {
                    DiManager.Company.EndTransaction(BoWfTransOpt.wf_RollBack);
                }
                return;
            }
            Recordset recSet = (Recordset)DiManager.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
            recSet.DoQuery(DiManager.QueryHanaTransalte("select top(1) Code from [@RSM_BOMD_HEADER] order by Code desc"));
            Code = recSet.Fields.Item("Code").Value.ToString();
            if (DiManager.Company.InTransaction)
            {
                DiManager.Company.EndTransaction(BoWfTransOpt.wf_Commit);
            }
        }

        public void AddRows()
        {
            if (!DiManager.Company.InTransaction)
            {
                DiManager.Company.StartTransaction();
            }
            UserTable rsmBomRow = DiManager.Company.UserTables.Item("RSM_BOMD_ROW");
            foreach (BomDissasemblyRow bomDissasemblyRow in BomDissasemblyRows)
            {
                rsmBomRow.UserFields.Fields.Item("U_ComponentCode").Value = bomDissasemblyRow.ComponentCode;
                rsmBomRow.UserFields.Fields.Item("U_ComponentName").Value = bomDissasemblyRow.ComponentName;
                rsmBomRow.UserFields.Fields.Item("U_Quantity").Value = bomDissasemblyRow.Quantity.ToString(CultureInfo.InvariantCulture);
                rsmBomRow.UserFields.Fields.Item("U_TotalQuantity").Value = bomDissasemblyRow.TotalQuantity.ToString(CultureInfo.InvariantCulture);
                rsmBomRow.UserFields.Fields.Item("U_ComponentCost").Value = bomDissasemblyRow.ComponentCost.ToString(CultureInfo.InvariantCulture);
                rsmBomRow.UserFields.Fields.Item("U_ComponentTotalCost").Value = bomDissasemblyRow.ComponentTotalCost.ToString(CultureInfo.InvariantCulture);
                rsmBomRow.UserFields.Fields.Item("U_WhsCode").Value = bomDissasemblyRow.WhsCode;
                rsmBomRow.UserFields.Fields.Item("U_WhsName").Value = bomDissasemblyRow.WhsName ?? "";
                rsmBomRow.UserFields.Fields.Item("U_Comment").Value = Comment;
                rsmBomRow.UserFields.Fields.Item("U_BplId").Value = bomDissasemblyRow.BplId.ToString(CultureInfo.InvariantCulture);
                rsmBomRow.UserFields.Fields.Item("U_ParentCode").Value = Code;
                int resRow = rsmBomRow.Add();
                if (resRow != 0)
                {
                    Application.SBO_Application.SetStatusBarMessage(DiManager.Company.GetLastErrorDescription(),
                        BoMessageTime.bmt_Short, true);

                    if (DiManager.Company.InTransaction)
                    {
                        DiManager.Company.EndTransaction(BoWfTransOpt.wf_RollBack);
                    }
                    return;
                }


                Recordset recSet2 = (Recordset)DiManager.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
                recSet2.DoQuery(DiManager.QueryHanaTransalte($"select top(1) Code from [@RSM_BOMD_ROW] order by Code desc"));
                if (!recSet2.EoF)
                {
                    bomDissasemblyRow.Code = recSet2.Fields.Item("Code").Value.ToString();
                }
                if (DiManager.Company.InTransaction)
                {
                    DiManager.Company.EndTransaction(BoWfTransOpt.wf_Commit);
                }
            }
        } 

    }

    class BomDissasemblyRow
    {
        public string ComponentCode { get; set; }
        public string Code { get; set; }
        public string ParentCode { get; set; }
        public string ComponentName { get; set; }
        public double Quantity { get; set; }
        public double TotalQuantity { get; set; }
        public double ComponentCost { get; set; }
        public double ComponentTotalCost { get; set; }
        public string WhsCode { get; set; }
        public string WhsName { get; set; }
        public int BplId { get; set; }


    }
}
