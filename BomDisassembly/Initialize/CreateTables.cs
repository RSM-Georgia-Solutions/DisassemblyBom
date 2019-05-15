using SAPbobsCOM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SAPbouiCOM;
using Application = SAPbouiCOM.Framework.Application;

namespace BomDisassembly.Initialize
{
    class CreateTables : IRunnable
    {
        public void Run(DiManager diManager)
        {
            if (!DiManager.Company.InTransaction)
            {
                DiManager.Company.StartTransaction();
            }

            if (diManager.CreateTable("RSM_BOMD_HEADER", BoUTBTableType.bott_NoObjectAutoIncrement) &&
                diManager.CreateTable("RSM_BOMD_ROW", BoUTBTableType.bott_NoObjectAutoIncrement))
            {
                if (!DiManager.Company.InTransaction)
                {
                    Application.SBO_Application.SetStatusBarMessage("ცხირლები წარმატებით შეიქმნა",
                        BoMessageTime.bmt_Short, false);
                    DiManager.Company.EndTransaction(BoWfTransOpt.wf_Commit);
                }
            }
        }
    }
}
