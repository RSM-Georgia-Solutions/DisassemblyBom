using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SAPbobsCOM;
using SAPbouiCOM;
using Application = SAPbouiCOM.Framework.Application;

namespace BomDisassembly.Initialize
{
    class CreateFields : IRunnable
    {
        public void Run(DiManager diManager)
        {
            if (!DiManager.Company.InTransaction)
            {
                DiManager.Company.StartTransaction();
            }

            if (
                diManager.AddField("RSM_BOMD_HEADER", "AccountCode", "Account Code", BoFieldTypes.db_Alpha, 250, false) &&
                diManager.AddField("RSM_BOMD_HEADER", "ItemCode", "Item Code", BoFieldTypes.db_Alpha, 250, false) &&
                diManager.AddField("RSM_BOMD_HEADER", "GoodsIssueDocEntry", "Goods Issue DocEntry", BoFieldTypes.db_Alpha, 250, false) &&
                diManager.AddField("RSM_BOMD_HEADER", "GoodsIssueDocNum", "Goods Issue DocNum", BoFieldTypes.db_Alpha, 250, false) &&
                diManager.AddField("RSM_BOMD_HEADER", "GoodsReceiptDocNum", "Goods Receipt DocNum", BoFieldTypes.db_Alpha, 250, false) &&
                diManager.AddField("RSM_BOMD_HEADER", "GoodsReceiptDocEntry", "Goods Receipt DocEntry", BoFieldTypes.db_Alpha, 250, false) &&
                diManager.AddField("RSM_BOMD_HEADER", "ItemName", "Item Name", BoFieldTypes.db_Alpha, 250, false) &&
                diManager.AddField("RSM_BOMD_HEADER", "UOM", "Unit Of Measure", BoFieldTypes.db_Alpha, 250, false) &&
                diManager.AddField("RSM_BOMD_HEADER", "WhsCode", "WareHouse Code", BoFieldTypes.db_Alpha, 250, false) &&
                diManager.AddField("RSM_BOMD_HEADER", "WhsName", "WareHouse Name", BoFieldTypes.db_Alpha, 250, false) &&
                diManager.AddField("RSM_BOMD_HEADER", "PostingDate", "Posting Date", BoFieldTypes.db_Alpha, 250, false) &&
                diManager.AddField("RSM_BOMD_HEADER", "PlannedQuantity", "PlannedQuantity", BoFieldTypes.db_Alpha, 250, false) &&
                diManager.AddField("RSM_BOMD_HEADER", "ItemCost", "Item Cost", BoFieldTypes.db_Alpha, 250, false) &&
                diManager.AddField("RSM_BOMD_HEADER", "ItemTotalCost", "ItemTotalCost", BoFieldTypes.db_Alpha, 250, false) &&
                diManager.AddField("RSM_BOMD_HEADER", "Comment", "Comment", BoFieldTypes.db_Alpha, 250, false) &&

                diManager.AddField("RSM_BOMD_ROW", "BplId", "Branch Code", BoFieldTypes.db_Alpha, 250, false) &&
                diManager.AddField("RSM_BOMD_ROW", "ComponentCode", "Item Code", BoFieldTypes.db_Alpha, 250, false) &&
                diManager.AddField("RSM_BOMD_ROW", "ParentCode", "Parent Code", BoFieldTypes.db_Alpha, 250, false) &&
                diManager.AddField("RSM_BOMD_ROW", "ComponentName", "Item Name", BoFieldTypes.db_Alpha, 250, false) &&
                diManager.AddField("RSM_BOMD_ROW", "Quantity", "Quantity", BoFieldTypes.db_Alpha, 250, false) &&
                diManager.AddField("RSM_BOMD_ROW", "TotalQuantity", "TotalQuantity", BoFieldTypes.db_Alpha, 250, false) &&
                diManager.AddField("RSM_BOMD_ROW", "ComponentCost", "Item Cost", BoFieldTypes.db_Alpha, 250, false) &&
                diManager.AddField("RSM_BOMD_ROW", "ComponentTotalCost", "Item Total Cost", BoFieldTypes.db_Alpha, 250, false) &&
                diManager.AddField("RSM_BOMD_ROW", "WhsCode", "WareHouse Code", BoFieldTypes.db_Alpha, 250, false) &&
                diManager.AddField("RSM_BOMD_ROW", "WhsName", "WareHouse Name", BoFieldTypes.db_Alpha, 250, false) &&
                diManager.AddField("RSM_BOMD_ROW", "Comment", "Comment", BoFieldTypes.db_Alpha, 250, false)
                )
            {
                if (DiManager.Company.InTransaction)
                {
                    Application.SBO_Application.SetStatusBarMessage("ველები წარმატებით შეიქმნა",
                        BoMessageTime.bmt_Short, false);
                    DiManager.Company.EndTransaction(BoWfTransOpt.wf_Commit);
                }
            }
            else
            {
                if (DiManager.Company.InTransaction)
                {
                    Application.SBO_Application.SetStatusBarMessage("პრობლემა მოხდა ველების შეიქმნისას",
                        BoMessageTime.bmt_Short, true);
                    DiManager.Company.EndTransaction(BoWfTransOpt.wf_RollBack);
                }
            }

        }
    }
}
