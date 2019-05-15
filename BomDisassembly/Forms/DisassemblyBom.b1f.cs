using System;
using System.CodeDom;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading;
using System.Xml.Linq;
using BomDisassembly.Models;
using SAPbobsCOM;
using SAPbouiCOM;
using SAPbouiCOM.Framework;
using Application = SAPbouiCOM.Framework.Application;

namespace BomDisassembly.Forms
{
    [FormAttribute("BomDisassembly.Forms.DisassemblyBom", "Forms/DisassemblyBom.b1f")]
    class DisassemblyBom : UserFormBase
    {
        private readonly BomDissasembly _bomDissasembly;
        private string _gridQuery;
        public DisassemblyBom()
        {
            _bomDissasembly = new BomDissasembly();
        }

        private bool isFromCode = false;

        public DisassemblyBom(string code)
        {
            _bomDissasembly = new BomDissasembly();
            FillModelFromCode(code);
            isFromCode = true;
        }



        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.StaticText0 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_0").Specific));
            this.EditText0 = ((SAPbouiCOM.EditText)(this.GetItem("Item_1").Specific));
            this.EditText0.ChooseFromListAfter += new SAPbouiCOM._IEditTextEvents_ChooseFromListAfterEventHandler(this.EditText0_ChooseFromListAfter);
            this.StaticText1 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_2").Specific));
            this.EditText1 = ((SAPbouiCOM.EditText)(this.GetItem("Item_3").Specific));
            this.EditText2 = ((SAPbouiCOM.EditText)(this.GetItem("Item_4").Specific));
            this.StaticText2 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_5").Specific));
            this.StaticText3 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_6").Specific));
            this.EditText3 = ((SAPbouiCOM.EditText)(this.GetItem("Item_7").Specific));
            this.EditText3.ValidateAfter += new SAPbouiCOM._IEditTextEvents_ValidateAfterEventHandler(this.EditText3_ValidateAfter);
            this.StaticText4 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_8").Specific));
            this.EditText4 = ((SAPbouiCOM.EditText)(this.GetItem("Item_9").Specific));
            this.StaticText5 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_10").Specific));
            this.EditText5 = ((SAPbouiCOM.EditText)(this.GetItem("Item_11").Specific));
            this.EditText5.ChooseFromListAfter += new SAPbouiCOM._IEditTextEvents_ChooseFromListAfterEventHandler(this.EditText5_ChooseFromListAfter);
            this.StaticText6 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_12").Specific));
            this.EditText6 = ((SAPbouiCOM.EditText)(this.GetItem("Item_13").Specific));
            this.EditText7 = ((SAPbouiCOM.EditText)(this.GetItem("Item_14").Specific));
            this.StaticText7 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_15").Specific));
            this.StaticText8 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_16").Specific));
            this.EditText8 = ((SAPbouiCOM.EditText)(this.GetItem("Item_17").Specific));
            this.Grid0 = ((SAPbouiCOM.Grid)(this.GetItem("Item_18").Specific));
            this.Grid0.ValidateAfter += new SAPbouiCOM._IGridEvents_ValidateAfterEventHandler(this.Grid0_ValidateAfter);
            this.Grid0.ClickAfter += new SAPbouiCOM._IGridEvents_ClickAfterEventHandler(this.Grid0_ClickAfter);
            this.Button0 = ((SAPbouiCOM.Button)(this.GetItem("Item_19").Specific));
            this.Button0.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.Button0_PressedAfter);
            this.StaticText9 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_20").Specific));
            this.EditText9 = ((SAPbouiCOM.EditText)(this.GetItem("Item_21").Specific));
            this.StaticText10 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_22").Specific));
            this.EditText10 = ((SAPbouiCOM.EditText)(this.GetItem("Item_23").Specific));
            this.LinkedButton0 = ((SAPbouiCOM.LinkedButton)(this.GetItem("Item_24").Specific));
            this.LinkedButton1 = ((SAPbouiCOM.LinkedButton)(this.GetItem("Item_25").Specific));
            this.LinkedButton2 = ((SAPbouiCOM.LinkedButton)(this.GetItem("Item_26").Specific));
            this.Button1 = ((SAPbouiCOM.Button)(this.GetItem("0").Specific));
            this.Button1.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.Button1_PressedAfter);
            this.Grid1 = ((SAPbouiCOM.Grid)(this.GetItem("Item_29").Specific));
            this.Grid2 = ((SAPbouiCOM.Grid)(this.GetItem("Item_30").Specific));
            this.EditText11 = ((SAPbouiCOM.EditText)(this.GetItem("GoodsIssue").Specific));
            this.EditText12 = ((SAPbouiCOM.EditText)(this.GetItem("GoodsR").Specific));
            this.EditText13 = ((SAPbouiCOM.EditText)(this.GetItem("Item_27").Specific));
            this.StaticText11 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_28").Specific));
            this.Button2 = ((SAPbouiCOM.Button)(this.GetItem("Item_31").Specific));
            this.Button2.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.Button2_PressedAfter);
            this.EditText14 = ((SAPbouiCOM.EditText)(this.GetItem("Item_32").Specific));
            this.EditText15 = ((SAPbouiCOM.EditText)(this.GetItem("Item_33").Specific));
            this.StaticText12 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_34").Specific));
            this.EditText16 = ((SAPbouiCOM.EditText)(this.GetItem("Item_35").Specific));
            this.StaticText13 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_36").Specific));
            this.OnCustomInitialize();

        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
            this.VisibleAfter += new SAPbouiCOM.Framework.FormBase.VisibleAfterHandler(this.Form_VisibleAfter);
            this.CloseAfter += new CloseAfterHandler(this.Form_CloseAfter);

        }


        private string GenereteGridQuery(double quantity = 1)
        {
            return $"select  OITW.ItemCode, ItemName ,Quantity, Quantity*{quantity} as [TotalQuantity], OITW.AvgPrice as [Cost], " +
                   $" OITW.AvgPrice * Quantity*{quantity} as [TotalCost], WareHouse  from ITT1 JOIN OITM on OITM.ItemCode = ITT1.Code  JOIN OITW on OITW.ItemCode = ITT1.Code AND OITW.WhsCode = ITT1.Warehouse";
        }

        private SAPbouiCOM.StaticText StaticText0;

        private void OnCustomInitialize()
        {
            SAPbouiCOM.Framework.Application.SBO_Application.ItemEvent += new _IApplicationEvents_ItemEventEventHandler(this.SBO_Application_ItemEvent_ChooseFromList);
            string query = $"{GenereteGridQuery()} Where 1 = 2";
            Grid0.DataTable.ExecuteQuery(DiManager.QueryHanaTransalte(query));

            GetItem("GoodsIssue").Left = 1000;
            GetItem("GoodsR").Left = 1000;
            GetItem("Item_21").Enabled = false;
            GetItem("Item_23").Enabled = false;
            GetItem("Item_32").Left = 1000;
        }

        private void SBO_Application_ItemEvent_ChooseFromList(string FormUID, ref ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            if (pVal.FormTypeEx != "BomDisassembly.Forms.DisassemblyBom")
            {
                return;
            }

            ChooseFromList(FormUID, pVal, "Item_1", "Item_3", "ItemCode", "ItemName");
            ChooseFromList(FormUID, pVal, "Item_11", "Item_13", "WhsCode", "WhsName");
            ChooseFromList(FormUID, pVal, "Item_33", "", "AccCode", "WhsName");

            if (pVal.ColUID == "ItemCode")
            {
                ChooseFromList(FormUID, pVal, "ItemCode", "ItemName", "CFL_ItemCmp", "", true, "Item_18");

            }
            if (pVal.ColUID == "WareHouse")
            {
                ChooseFromList(FormUID, pVal, "WareHouse", "WareHouse", "WhsCode", "WhsName", true, "Item_18");
            }

            if ((FormUID == "CFL1") & (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD))
            {
                System.Windows.Forms.Application.Exit();
            }
        }


        private void ChooseFromList(string FormUID, ItemEvent pVal, string itemUId, string itemUIdDesc, string dataSourceId = "", string dataSourceDescId = "", bool isMatrix = false, string matrixUid = "")
        {
            if (pVal.EventType == BoEventTypes.et_CHOOSE_FROM_LIST)
            {
                string val = null;
                string val2 = null;
                try
                {
                    IChooseFromListEvent oCFLEvento = null;
                    oCFLEvento = ((IChooseFromListEvent)(pVal));
                    string sCFL_ID = null;
                    sCFL_ID = oCFLEvento.ChooseFromListUID;
                    Form oForm = null;
                    oForm = SAPbouiCOM.Framework.Application.SBO_Application.Forms.Item(FormUID);
                    SAPbouiCOM.ChooseFromList oCFL = null;
                    oCFL = oForm.ChooseFromLists.Item(sCFL_ID);

                    if (oCFLEvento.BeforeAction == false)
                    {
                        DataTable oDataTable = null;
                        oDataTable = oCFLEvento.SelectedObjects;

                        try
                        {
                            val = Convert.ToString(oDataTable.GetValue(0, 0));
                            val2 = Convert.ToString(oDataTable.GetValue(1, 0));
                        }
                        catch (Exception ex)
                        {

                        }
                        if (pVal.ItemUID == itemUId || pVal.ItemUID == matrixUid)
                        {
                            if (isMatrix)
                            {
                                Grid0.DataTable.SetValue(itemUId, pVal.Row, val);
                                Grid0.DataTable.SetValue(itemUIdDesc, pVal.Row, val2);
                            }
                            else if (pVal.ItemUID == itemUId)
                            {
                                var xz = Application.SBO_Application.Forms.GetForm("BomDisassembly.Forms.DisassemblyBom", 1);

                                xz.DataSources.UserDataSources.Item(dataSourceId).Value = val;
                                if (!string.IsNullOrWhiteSpace(dataSourceDescId))
                                {
                                    xz.DataSources.UserDataSources.Item(dataSourceDescId).Value = val2;
                                }

                            }
                        }
                    }
                }
                catch (Exception e)
                {
                }
            }
        }

        private void Form_VisibleAfter(SAPbouiCOM.SBOItemEventArg pVal)
        {
            if (isFromCode)
            {
                FillFormFromModel();
            }

            var activeForm = SAPbouiCOM.Framework.Application.SBO_Application.Forms.ActiveForm;
            if (activeForm.Title == "Dissembly Bill Of Material")
            {
                RefreshGrid();

                if (Grid0.Item.Enabled == true)
                {
                    DiManager.AddRightClickMenue("Dissembly Bill Of Material", "AddRow", "Add Row");
                    DiManager.AddRightClickMenue("Dissembly Bill Of Material", "RmvRow", "Remove Row");
                }


                ChooseFromListCollection oCfLs = activeForm.ChooseFromLists; ChooseFromListCreationParams oCflCreationParams = (ChooseFromListCreationParams)Application.SBO_Application.CreateObject(BoCreatableObjectType
                    .cot_ChooseFromListCreationParams);
                oCflCreationParams.ObjectType = "66";
                oCflCreationParams.UniqueID = "CFL_ItemCode";

                SAPbouiCOM.ChooseFromList oCfl = oCfLs.Item("CFL_ItemCode");

                Conditions oCons = oCfl.GetConditions();
                Condition oCon = oCons.Add();

                oCon.Alias = "TreeType";
                oCon.Operation = BoConditionOperation.co_EQUAL;
                oCon.CondVal = "P";
                oCfl.SetConditions(oCons);



                SAPbouiCOM.EditTextColumn gridColumnWhs = (EditTextColumn)Grid0.Columns.Item("WareHouse");
                gridColumnWhs.ChooseFromListUID = "CFL_WhsCmp";

                SAPbouiCOM.EditTextColumn gridColumnitemCode = (EditTextColumn)Grid0.Columns.Item("ItemCode");
                gridColumnitemCode.ChooseFromListUID = "CFL_ItemCmp";


                activeForm.Items.Item("Item_3").Enabled = false;
                activeForm.Items.Item("Item_4").Enabled = false;
                activeForm.Items.Item("Item_13").Enabled = false;
                activeForm.Items.Item("Item_14").Enabled = false;
                activeForm.Items.Item("Item_17").Enabled = false;


            }
        }

        private SAPbouiCOM.EditText EditText0;
        private SAPbouiCOM.StaticText StaticText1;
        private SAPbouiCOM.EditText EditText1;
        private SAPbouiCOM.EditText EditText2;
        private SAPbouiCOM.StaticText StaticText2;
        private SAPbouiCOM.StaticText StaticText3;
        private SAPbouiCOM.EditText EditText3;
        private SAPbouiCOM.StaticText StaticText4;
        private SAPbouiCOM.EditText EditText4;
        private SAPbouiCOM.StaticText StaticText5;
        private SAPbouiCOM.EditText EditText5;
        private SAPbouiCOM.StaticText StaticText6;
        private SAPbouiCOM.EditText EditText6;
        private SAPbouiCOM.EditText EditText7;
        private SAPbouiCOM.StaticText StaticText7;
        private SAPbouiCOM.StaticText StaticText8;
        private SAPbouiCOM.EditText EditText8;
        private SAPbouiCOM.Grid Grid0;
        private SAPbouiCOM.Button Button0;
        private SAPbouiCOM.StaticText StaticText9;
        private SAPbouiCOM.EditText EditText9;
        private SAPbouiCOM.StaticText StaticText10;
        private SAPbouiCOM.EditText EditText10;
        private SAPbouiCOM.LinkedButton LinkedButton0;
        private SAPbouiCOM.LinkedButton LinkedButton1;

        private void Button0_PressedAfter(object sboObject, SBOItemEventArg pVal)
        {

        }

        private LinkedButton LinkedButton2;

        private void EditText0_ChooseFromListAfter(object sboObject, SBOItemEventArg pVal)
        {
            CalculateHeaderVariables();
            RefreshGrid();
        }

        private void CalculateGridVariables()
        {

        }

        private void RefreshGrid()
        {
            Thread.CurrentThread.CurrentCulture = CultureInfo.InvariantCulture;
            if (EditText0.Value != "")
            {
                double quantity = double.Parse(EditText3.Value, CultureInfo.InvariantCulture);
                Grid0.DataTable.ExecuteQuery(DiManager.QueryHanaTransalte($"{GenereteGridQuery(quantity)} Where Father = N'{EditText0.Value}'"));

                CalculateDifference();

                Grid0.AutoResizeColumns();
            }
        }

        private void CalculateDifference()
        {
            Thread.CurrentThread.CurrentCulture = CultureInfo.InvariantCulture;
            double totalCostHeader = double.Parse(EditText8.Value, CultureInfo.InvariantCulture);
            double totalCostRow = 0;

            for (int i = 0; i < Grid0.DataTable.Rows.Count; i++)
            {
                double totalCost = double.Parse(Grid0.DataTable.GetValue("TotalCost", i).ToString(), CultureInfo.InvariantCulture);
                totalCostRow += totalCost;
            }

            double diff = Math.Round(totalCostHeader - totalCostRow, 2);
            string diffString = diff.ToString(CultureInfo.InvariantCulture);

            var bomDissForm = SAPbouiCOM.Framework.Application.SBO_Application.Forms.GetForm("BomDisassembly.Forms.DisassemblyBom", 1);

            bomDissForm.DataSources.UserDataSources.Item("Diff").Value = diffString;
            //EditText16.Value = diffString;
            _bomDissasembly.DiffBetweenCosts = diff;
        }

        private void CalculateHeaderVariables()
        {
            Thread.CurrentThread.CurrentCulture = CultureInfo.InvariantCulture;
            if (EditText0.Value == "" || EditText5.Value == "")
            {
                return;
            }

            string itemCode = EditText0.Value;
            string whsCode = EditText5.Value;
            string accountCode = EditText15.Value;
            string itemUom = string.Empty;
            string bplId = string.Empty;
            double quantity = double.Parse(EditText3.Value, CultureInfo.InvariantCulture);
            // ReSharper disable once CompareOfFloatsByEqualityOperator
            if (quantity == 0)
            {
                quantity = 1;
            }
            double cost = 0;

            Recordset recSet2 = (Recordset)DiManager.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
            recSet2.DoQuery(DiManager.QueryHanaTransalte($"select OWHS.BPLid from OWHS where WhsCode = N'{whsCode}'"));
            if (!recSet2.EoF)
            {
                bplId = recSet2.Fields.Item("BPLid").Value.ToString();
            }

            SAPbobsCOM.Items item = (SAPbobsCOM.Items)DiManager.Company.GetBusinessObject(BoObjectTypes.oItems);
            item.GetByKey(itemCode);
            itemUom = item.SalesUnit;

            Recordset recSet = (Recordset)DiManager.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
            recSet.DoQuery(DiManager.QueryHanaTransalte($"select AvgPrice from OITW where ItemCode = N'{itemCode}' And WhsCode = N'{whsCode}'"));

            if (!recSet.EoF)
            {
                var x = recSet.Fields.Item("AvgPrice").Value.ToString();
                cost = double.Parse(recSet.Fields.Item("AvgPrice").Value.ToString(), CultureInfo.InvariantCulture);
            }

            var bomDissForm = SAPbouiCOM.Framework.Application.SBO_Application.Forms.GetForm("BomDisassembly.Forms.DisassemblyBom", 1);

            bomDissForm.DataSources.UserDataSources.Item("Uom").Value = itemUom;
            bomDissForm.DataSources.UserDataSources.Item("AccCode").Value = accountCode;
            bomDissForm.DataSources.UserDataSources.Item("BplId").Value = bplId;
            bomDissForm.DataSources.UserDataSources.Item("Cost").Value = cost.ToString(CultureInfo.InvariantCulture);
            bomDissForm.DataSources.UserDataSources.Item("Quantity").Value = quantity.ToString(CultureInfo.InvariantCulture);
            bomDissForm.DataSources.UserDataSources.Item("TotalCost").Value = (quantity * cost).ToString(CultureInfo.InvariantCulture);

        }



        private Button Button1;

        private void EditText5_ChooseFromListAfter(object sboObject, SBOItemEventArg pVal)
        {
            CalculateHeaderVariables();
            RefreshGrid();
        }



        /// <summary>
        /// Multiply Cost to Quantity
        /// </summary>
        private void CostToQuantity()
        {
            var bomDissForm = SAPbouiCOM.Framework.Application.SBO_Application.Forms.GetForm("BomDisassembly.Forms.DisassemblyBom", 1);
            double cost = double.Parse(bomDissForm.DataSources.UserDataSources.Item("Cost").Value, CultureInfo.InvariantCulture);
            double quantity = double.Parse(EditText3.Value, CultureInfo.InvariantCulture);
            bomDissForm.DataSources.UserDataSources.Item("TotalCost").Value = (quantity * cost).ToString(CultureInfo.InvariantCulture);
        }

        private void EditText3_ValidateAfter(object sboObject, SBOItemEventArg pVal)
        {
            CostToQuantity();
            RefreshGrid();
        }

        private Grid Grid1;
        private Grid Grid2;

        private bool ValidateForm()
        {
            var bomDissForm = SAPbouiCOM.Framework.Application.SBO_Application.Forms.GetForm("BomDisassembly.Forms.DisassemblyBom", 1);

            if (EditText0.Value == "")
            {
                Application.SBO_Application.SetStatusBarMessage("ItemCode Is Required",
                    BoMessageTime.bmt_Short, true);
                return false;
            }
            if (EditText15.Value == "")
            {
                Application.SBO_Application.SetStatusBarMessage("AccountCode Is Required",
                    BoMessageTime.bmt_Short, true);
                return false;
            }
            if (EditText3.Value == "0.0")
            {
                Application.SBO_Application.SetStatusBarMessage("Quantity Must Be Greater Then Zero",
                    BoMessageTime.bmt_Short, true);
                return false;
            }
            if (EditText4.Value == "")
            {
                Application.SBO_Application.SetStatusBarMessage("Posting Date Is Required",
                    BoMessageTime.bmt_Short, true);
                return false;
            }
            if (EditText5.Value == "")
            {
                Application.SBO_Application.SetStatusBarMessage("Warehosue Is Required",
                    BoMessageTime.bmt_Short, true);
                return false;
            }

            return true;
        }

        private void FillModelFromCode(string code)
        {
            Thread.CurrentThread.CurrentCulture = CultureInfo.InvariantCulture;
            Recordset recSet = (Recordset)DiManager.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
            Recordset Whs = (Recordset)DiManager.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
            recSet.DoQuery(DiManager.QueryHanaTransalte($@"SELECT [@RSM_BOMD_HEADER].*, [@RSM_BOMD_ROW].U_ComponentCode,               [@RSM_BOMD_ROW].U_ComponentName,
                [@RSM_BOMD_ROW].Code as [CodeRow],
                [@RSM_BOMD_ROW].U_ComponentCost,
                [@RSM_BOMD_ROW].U_ComponentTotalCost,
                [@RSM_BOMD_ROW].U_Quantity, [@RSM_BOMD_ROW].U_TotalQuantity, [@RSM_BOMD_ROW].U_WhsCode as [WhsCodeCmp],
                [@RSM_BOMD_ROW].U_WhsName as [WhsNameCmp], [@RSM_BOMD_ROW].U_BplId as [Branch]
                FROM [@RSM_BOMD_HEADER]
                Left join [@RSM_BOMD_ROW] on [@RSM_BOMD_HEADER].Code =
                [@RSM_BOMD_ROW].U_ParentCode WHERE [@RSM_BOMD_HEADER].Code = { code}"));


            _bomDissasembly.WhsCode = recSet.Fields.Item("U_WhsCode").Value.ToString();
            Whs.DoQuery(DiManager.QueryHanaTransalte($"SELECT BPLid FROM OWHS WHERE WhsCode = N'{ _bomDissasembly.WhsCode}'"));
            int bplId = int.Parse(Whs.Fields.Item("BPLid").Value.ToString());

            var bomDissForm = Application.SBO_Application.Forms.GetForm("BomDisassembly.Forms.DisassemblyBom", 1);
            bomDissForm.DataSources.UserDataSources.Item("BplId").Value = Whs.Fields.Item("BPLid").Value.ToString();

            _bomDissasembly.BplId = bplId;
            _bomDissasembly.Code = recSet.Fields.Item("Code").Value.ToString();
            _bomDissasembly.Comment = recSet.Fields.Item("U_Comment").Value.ToString();
            _bomDissasembly.ItemTotalCost = double.Parse(recSet.Fields.Item("U_ItemTotalCost").Value.ToString(), CultureInfo.InvariantCulture);
            _bomDissasembly.PlannedQuantity = double.Parse(recSet.Fields.Item("U_PlannedQuantity").Value.ToString(), CultureInfo.InvariantCulture);
            _bomDissasembly.GoodsIssueDocEntry = recSet.Fields.Item("U_GoodsIssueDocEntry").Value.ToString();
            _bomDissasembly.GoodsIssueDocNum = recSet.Fields.Item("U_GoodsIssueDocNum").Value.ToString();
            _bomDissasembly.GoodsReceiptDocEntry = recSet.Fields.Item("U_GoodsReceiptDocEntry").Value.ToString();
            _bomDissasembly.GoodsReceiptDocNum = recSet.Fields.Item("U_GoodsReceiptDocNum").Value.ToString();
            _bomDissasembly.ItemCode = recSet.Fields.Item("U_ItemCode").Value.ToString();
            _bomDissasembly.ItemName = recSet.Fields.Item("U_ItemName").Value.ToString();
            _bomDissasembly.Uom = recSet.Fields.Item("U_UOM").Value.ToString();
            _bomDissasembly.WhsName = recSet.Fields.Item("U_WhsName").Value.ToString();
            _bomDissasembly.ItemCost = double.Parse(recSet.Fields.Item("U_ItemCost").Value.ToString());
            _bomDissasembly.AccountCode = recSet.Fields.Item("U_AccountCode").Value.ToString();
            _bomDissasembly.PostingDate = DateTime.Parse(recSet.Fields.Item("U_PostingDate").Value.ToString());

            List<BomDissasemblyRow> rows = new List<BomDissasemblyRow>();

            while (!recSet.EoF)
            {
                var componentCode = recSet.Fields.Item("U_ComponentCode").Value.ToString(); ;
                if (string.IsNullOrWhiteSpace(componentCode))
                {
                    recSet.MoveNext();
                    continue;
                }
                BomDissasemblyRow row = new BomDissasemblyRow();
                row.ParentCode = _bomDissasembly.Code;
                row.ComponentCode = recSet.Fields.Item("CodeRow").Value.ToString();
                row.ComponentCode = recSet.Fields.Item("U_ComponentCode").Value.ToString();
                row.ComponentName = recSet.Fields.Item("U_ComponentName").Value.ToString();
                row.ComponentCost = double.Parse(recSet.Fields.Item("U_ComponentCost").Value.ToString());
                row.ComponentTotalCost = double.Parse(recSet.Fields.Item("U_ComponentTotalCost").Value.ToString());
                row.Quantity = double.Parse(recSet.Fields.Item("U_Quantity").Value.ToString());
                row.TotalQuantity = double.Parse(recSet.Fields.Item("U_TotalQuantity").Value.ToString());
                row.WhsCode = recSet.Fields.Item("WhsCodeCmp").Value.ToString();
                row.WhsName = recSet.Fields.Item("WhsNameCmp").Value.ToString();
                row.BplId = int.Parse(recSet.Fields.Item("Branch").Value.ToString());
                rows.Add(row);
                recSet.MoveNext();
            }

            if (rows.Count > 0)
            {
                _bomDissasembly.BomDissasemblyRows = rows;
            }
            //else
            //{
            //    RefreshGrid();
            //    List<BomDissasemblyRow> rows2 = new List<BomDissasemblyRow>();
            //    for (int i = 0; i < Grid0.DataTable.Rows.Count; i++)
            //    {
            //        string warehouse = Grid0.DataTable.GetValue("WareHouse", i).ToString();
            //        recSet.DoQuery(DiManager.QueryHanaTransalte($"SELECT BPLid FROM OWHS WHERE WhsCode = N'{warehouse}'"));
            //        int rowbplId = int.Parse(recSet.Fields.Item("BPLid").Value.ToString());
            //        BomDissasemblyRow row =
            //            new BomDissasemblyRow
            //            {
            //                ComponentCode = Grid0.DataTable.GetValue("ItemCode", i).ToString(),
            //                ComponentName = Grid0.DataTable.GetValue("ItemName", i).ToString(),
            //                Quantity = double.Parse(Grid0.DataTable.GetValue("Quantity", i).ToString(),
            //                    CultureInfo.InvariantCulture),
            //                TotalQuantity = double.Parse(Grid0.DataTable.GetValue("TotalQuantity", i).ToString(),
            //                    CultureInfo.InvariantCulture),
            //                ComponentCost = double.Parse(Grid0.DataTable.GetValue("Cost", i).ToString(),
            //                    CultureInfo.InvariantCulture),
            //                ComponentTotalCost = double.Parse(Grid0.DataTable.GetValue("TotalCost", i).ToString(),
            //                    CultureInfo.InvariantCulture),
            //                WhsCode = warehouse,
            //                BplId = rowbplId
            //            };
            //        rows2.Add(row);
            //    }

            //    _bomDissasembly.BomDissasemblyRows = rows2;

            //    int index = 0;
            //    SAPbouiCOM.Framework.Application.SBO_Application.Forms.ActiveForm.Freeze(true);
            //    Grid0.DataTable.Rows.Add(_bomDissasembly.BomDissasemblyRows.Count() - 1);
            //    foreach (BomDissasemblyRow bomDissasemblyBomDissasemblyRow in _bomDissasembly.BomDissasemblyRows)
            //    {
            //        Grid0.DataTable.SetValue("ItemCode", index, bomDissasemblyBomDissasemblyRow.ComponentCode);
            //        Grid0.DataTable.SetValue("ItemName", index, bomDissasemblyBomDissasemblyRow.ComponentName);
            //        Grid0.DataTable.SetValue("Quantity", index, bomDissasemblyBomDissasemblyRow.Quantity);
            //        Grid0.DataTable.SetValue("TotalQuantity", index, bomDissasemblyBomDissasemblyRow.TotalQuantity);
            //        Grid0.DataTable.SetValue("Cost", index, bomDissasemblyBomDissasemblyRow.ComponentCost);
            //        Grid0.DataTable.SetValue("WareHouse", index, bomDissasemblyBomDissasemblyRow.WhsCode);
            //        Grid0.DataTable.SetValue("TotalCost", index, bomDissasemblyBomDissasemblyRow.ComponentTotalCost);
            //        index++;
            //    }

            //    SAPbouiCOM.Framework.Application.SBO_Application.Forms.ActiveForm.Freeze(false);
            //}


            try
            {
                EditText14.Active = true;
                Grid0.Item.Enabled = false;
            }
            catch (Exception e)
            {
            }
        }


        private void FillFormFromModel()
        {
            Thread.CurrentThread.CurrentCulture = CultureInfo.InvariantCulture;
            EditText14.Active = true;
            EditText14.Item.Left = 1000;
            EditText4.Item.Enabled = false;
            EditText0.Item.Enabled = false;
            EditText1.Item.Enabled = false;
            EditText2.Item.Enabled = false;
            EditText3.Item.Enabled = false;
            EditText5.Item.Enabled = false;
            EditText6.Item.Enabled = false;
            EditText7.Item.Enabled = false;
            EditText8.Item.Enabled = false;
            EditText15.Item.Enabled = false;
            EditText4.Value = _bomDissasembly.PostingDate.ToString("yyyyMMdd");
            EditText0.Value = _bomDissasembly.ItemCode;
            EditText1.Value = _bomDissasembly.ItemName;
            EditText2.Value = _bomDissasembly.Uom;
            EditText3.Value = _bomDissasembly.PlannedQuantity.ToString(CultureInfo.InvariantCulture);
            EditText5.Value = _bomDissasembly.WhsCode;
            EditText6.Value = _bomDissasembly.WhsName;
            EditText7.Value = _bomDissasembly.ItemCost.ToString(CultureInfo.InvariantCulture);
            EditText8.Value = _bomDissasembly.ItemTotalCost.ToString(CultureInfo.InvariantCulture);
            EditText13.Value = _bomDissasembly.Comment;
            EditText11.Value = _bomDissasembly.GoodsIssueDocEntry;
            EditText9.Value = _bomDissasembly.GoodsIssueDocNum;
            EditText12.Value = _bomDissasembly.GoodsReceiptDocEntry;
            EditText10.Value = _bomDissasembly.GoodsReceiptDocNum;
            EditText15.Value = _bomDissasembly.AccountCode;
            EditText14.Active = true;

            if (!string.IsNullOrWhiteSpace(_bomDissasembly.GoodsReceiptDocEntry))
            {
                Grid0.Item.Enabled = false;
            }

            int i = 0;
            SAPbouiCOM.Framework.Application.SBO_Application.Forms.ActiveForm.Freeze(true);
            try
            {
                Grid0.DataTable.Rows.Add(_bomDissasembly.BomDissasemblyRows.Count() - 1);
                foreach (BomDissasemblyRow bomDissasemblyBomDissasemblyRow in _bomDissasembly.BomDissasemblyRows)
                {
                    Grid0.DataTable.SetValue("ItemCode", i, bomDissasemblyBomDissasemblyRow.ComponentCode);
                    Grid0.DataTable.SetValue("ItemName", i, bomDissasemblyBomDissasemblyRow.ComponentName);
                    Grid0.DataTable.SetValue("Quantity", i, bomDissasemblyBomDissasemblyRow.Quantity);
                    Grid0.DataTable.SetValue("TotalQuantity", i, bomDissasemblyBomDissasemblyRow.TotalQuantity);
                    Grid0.DataTable.SetValue("Cost", i, bomDissasemblyBomDissasemblyRow.ComponentCost);
                    Grid0.DataTable.SetValue("WareHouse", i, bomDissasemblyBomDissasemblyRow.WhsCode);
                    Grid0.DataTable.SetValue("TotalCost", i, bomDissasemblyBomDissasemblyRow.ComponentTotalCost);
                    i++;
                }
            }
            catch (Exception e)
            {
 
            }

            SAPbouiCOM.Framework.Application.SBO_Application.Forms.ActiveForm.Freeze(false);

        }

        private void FillModelFromForm()
        {
            Thread.CurrentThread.CurrentCulture = CultureInfo.InvariantCulture;
            var bomDissForm = SAPbouiCOM.Framework.Application.SBO_Application.Forms.GetForm("BomDisassembly.Forms.DisassemblyBom", 1);

            var bplId = int.Parse(bomDissForm.DataSources.UserDataSources.Item("BplId").Value, CultureInfo.InvariantCulture);

            Recordset recSet = (Recordset)DiManager.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
            Recordset recSet2 = (Recordset)DiManager.Company.GetBusinessObject(BoObjectTypes.BoRecordset);


            _bomDissasembly.PostingDate = DateTime.ParseExact(EditText4.Value, "yyyyMMdd", CultureInfo.InvariantCulture);
            _bomDissasembly.BplId = bplId;
            _bomDissasembly.ItemCode = EditText0.Value;
            _bomDissasembly.ItemName = EditText1.Value;
            _bomDissasembly.Uom = EditText2.Value;
            _bomDissasembly.PlannedQuantity = double.Parse(EditText3.Value, CultureInfo.InvariantCulture);
            _bomDissasembly.WhsCode = EditText5.Value;
            _bomDissasembly.WhsName = EditText6.Value;
            _bomDissasembly.ItemCost = double.Parse(EditText7.Value, CultureInfo.InvariantCulture);
            _bomDissasembly.ItemTotalCost = double.Parse(EditText8.Value, CultureInfo.InvariantCulture);
            _bomDissasembly.Comment = EditText13.Value;
            _bomDissasembly.AccountCode = EditText15.Value;


            List<BomDissasemblyRow> rows = new List<BomDissasemblyRow>();
            for (int i = 0; i < Grid0.DataTable.Rows.Count; i++)
            {
                string warehouse = Grid0.DataTable.GetValue("WareHouse", i).ToString();
                recSet.DoQuery(DiManager.QueryHanaTransalte($"SELECT BPLid FROM OWHS WHERE WhsCode = N'{warehouse}'"));
                int rowbplId = int.Parse(recSet.Fields.Item("BPLid").Value.ToString());
                BomDissasemblyRow row =
                    new BomDissasemblyRow
                    {
                        ComponentCode = Grid0.DataTable.GetValue("ItemCode", i).ToString(),
                        ComponentName = Grid0.DataTable.GetValue("ItemName", i).ToString(),
                        Quantity = double.Parse(Grid0.DataTable.GetValue("Quantity", i).ToString(),
                            CultureInfo.InvariantCulture),
                        TotalQuantity = double.Parse(Grid0.DataTable.GetValue("TotalQuantity", i).ToString(),
                            CultureInfo.InvariantCulture),
                        ComponentCost = double.Parse(Grid0.DataTable.GetValue("Cost", i).ToString(),
                            CultureInfo.InvariantCulture),
                        ComponentTotalCost = double.Parse(Grid0.DataTable.GetValue("TotalCost", i).ToString(),
                            CultureInfo.InvariantCulture),
                        WhsCode = warehouse,
                        BplId = rowbplId
                    };
                rows.Add(row);
            }

            _bomDissasembly.BomDissasemblyRows = rows;

        }

        private void Button1_PressedAfter(object sboObject, SBOItemEventArg pVal)
        {
            if (!ValidateForm())
            {
                return;
            }

            if (!string.IsNullOrWhiteSpace(_bomDissasembly.Code))
            {
                Application.SBO_Application.SetStatusBarMessage("Goods Issue Already Added",
                    BoMessageTime.bmt_Short, true);
                return;
            }

            FillModelFromForm();

            SAPbobsCOM.Documents goodsIssue =
                (Documents)DiManager.Company.GetBusinessObject(BoObjectTypes.oInventoryGenExit);
            goodsIssue.DocDate = _bomDissasembly.PostingDate;
            goodsIssue.BPL_IDAssignedToInvoice = _bomDissasembly.BplId;

            goodsIssue.Lines.SetCurrentLine(0);
            goodsIssue.Lines.ItemCode = _bomDissasembly.ItemCode;
            goodsIssue.Lines.Quantity = _bomDissasembly.PlannedQuantity;
            goodsIssue.Lines.WarehouseCode = _bomDissasembly.WhsCode;
            goodsIssue.Lines.AccountCode = _bomDissasembly.AccountCode;
            var res = goodsIssue.Add();
            if (res != 0)
            {
                Application.SBO_Application.SetStatusBarMessage(DiManager.Company.GetLastErrorDescription(),
                    BoMessageTime.bmt_Short, true);
            }
            else
            {
                _bomDissasembly.GoodsIssueDocEntry = DiManager.Company.GetNewObjectKey();
                string code = string.Empty;
                DiManager.Company.GetNewObjectCode(out code);
                _bomDissasembly.GoodsIssueDocNum = code;
                EditText11.Value = _bomDissasembly.GoodsIssueDocEntry;
                EditText9.Value = _bomDissasembly.GoodsIssueDocNum;
                SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetSystemMessage("Success", BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Success);
                goodsIssue.GetByKey(int.Parse(_bomDissasembly.GoodsIssueDocEntry));
                var costs = getCost(goodsIssue);
                _bomDissasembly.ItemTotalCost = double.Parse(costs[0]) * _bomDissasembly.PlannedQuantity;

                _bomDissasembly.AddHeader();
            }
            GetItem("Item_32").Click();
            GetItem("Item_33").Enabled = false;
            GetItem("Item_9").Enabled = false;
            GetItem("Item_11").Enabled = false;
            GetItem("Item_7").Enabled = false;



        }


        public List<string> getCost(Documents objectXml)
        {
            var deliveryXmlString = objectXml.GetAsXML();
            XElement DelXml = XElement.Parse(deliveryXmlString);
            List<string> costs = new List<string>();
            try
            {
                foreach (var node in DelXml.Element("BO").Element("IGE1").Elements("row").Elements().Where(x => x.Name == "StockPrice"))
                {
                    var cost = node.Value;
                    costs.Add(cost);
                }
            }
            catch (Exception e)
            {

            }
            return costs;
        }

        private static void CreateChoseFromList(Form form, string objectType, string cflId, string editTextItem, string cflAlias, string dataSourceId, string conditionField = "", BoConditionOperation equality = BoConditionOperation.co_EQUAL, string conditionValue = "", string cflParamsId = "")
        {
            Form activeForm = form;
            ChooseFromListCollection oCfLs = activeForm.ChooseFromLists;
            ChooseFromListCreationParams oCflCreationParams = (ChooseFromListCreationParams)Application.SBO_Application.CreateObject(BoCreatableObjectType
                .cot_ChooseFromListCreationParams);
            oCflCreationParams.ObjectType = objectType;
            oCflCreationParams.UniqueID = cflId;
            SAPbouiCOM.ChooseFromList oCfl = oCfLs.Add(oCflCreationParams);
            Conditions oCons = oCfl.GetConditions();
            Condition oCon = oCons.Add();

            if (conditionField != "")
            {
                oCon.Alias = conditionField;
                oCon.Operation = equality;
                oCon.CondVal = conditionValue;
                oCfl.SetConditions(oCons);
            }
            oCflCreationParams.UniqueID = cflParamsId;
            oCfl = oCfLs.Add(oCflCreationParams);
            activeForm.DataSources.UserDataSources.Add(dataSourceId, BoDataType.dt_SHORT_TEXT, 254);
            EditText oEdit = ((EditText)(activeForm.Items.Item(editTextItem).Specific));
            oEdit.DataBind.SetBound(true, "", dataSourceId);
            oEdit.ChooseFromListUID = cflId;
            oEdit.ChooseFromListAlias = cflAlias;
        }

        private EditText EditText11;
        private EditText EditText12;
        private EditText EditText13;
        private StaticText StaticText11;
        private Button Button2;

        private void Button2_PressedAfter(object sboObject, SBOItemEventArg pVal)
        {

            CalculateDifference();

            if (string.IsNullOrWhiteSpace(_bomDissasembly.Code))
            {
                Application.SBO_Application.SetStatusBarMessage("Goods Issue Should Be Already Posted",
                    BoMessageTime.bmt_Short, true);
                return;
            }
            if (_bomDissasembly.DiffBetweenCosts != 0)
            {
                Application.SBO_Application.SetStatusBarMessage("Difference Should be Zero",
                    BoMessageTime.bmt_Short, true);
                return;
            }

            if (!string.IsNullOrWhiteSpace(_bomDissasembly.GoodsReceiptDocEntry))
            {
                Application.SBO_Application.SetStatusBarMessage("Goods Receipt Already Posted",
                    BoMessageTime.bmt_Short, true);
                return;
            }

            FillModelFromForm();

            SAPbobsCOM.Documents goodsReceipt =
                (Documents)DiManager.Company.GetBusinessObject(BoObjectTypes.oInventoryGenEntry);
            goodsReceipt.DocDate = _bomDissasembly.PostingDate;
            goodsReceipt.BPL_IDAssignedToInvoice = _bomDissasembly.BomDissasemblyRows.First().BplId;

            foreach (BomDissasemblyRow bomDissasemblyRow in _bomDissasembly.BomDissasemblyRows)
            {
                goodsReceipt.Lines.ItemCode = bomDissasemblyRow.ComponentCode;
                goodsReceipt.Lines.Quantity = bomDissasemblyRow.TotalQuantity;
                goodsReceipt.Lines.WarehouseCode = bomDissasemblyRow.WhsCode;
                goodsReceipt.Lines.AccountCode = _bomDissasembly.AccountCode;
                goodsReceipt.Lines.Price = bomDissasemblyRow.ComponentCost;
                goodsReceipt.Lines.Add();
            }

            var res = goodsReceipt.Add();
            if (res != 0)
            {
                Application.SBO_Application.SetStatusBarMessage(DiManager.Company.GetLastErrorDescription(),
                    BoMessageTime.bmt_Short, true);
            }
            else
            {
                _bomDissasembly.GoodsReceiptDocEntry = DiManager.Company.GetNewObjectKey();
                string code = string.Empty;
                DiManager.Company.GetNewObjectCode(out code);
                _bomDissasembly.GoodsReceiptDocNum = code;
                EditText12.Value = _bomDissasembly.GoodsReceiptDocEntry;
                EditText10.Value = _bomDissasembly.GoodsReceiptDocNum;
                SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetSystemMessage("Success", BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Success);
                _bomDissasembly.AddRows();
                EditText14.Active = true;
                Grid0.Item.Enabled = false;
            }


        }

        private EditText EditText14;
        private EditText EditText15;
        private StaticText StaticText12;
        private EditText EditText16;
        private StaticText StaticText13;


        private void Grid0_ClickAfter(object sboObject, SBOItemEventArg pVal)
        {
            SAPbouiCOM.Framework.Application.SBO_Application.Forms.ActiveForm.EnableMenu("RmvRow", true);
            SAPbouiCOM.Framework.Application.SBO_Application.Forms.ActiveForm.EnableMenu("AddRow", true);
            if (pVal.Row == -1)
            {
                return;
            }
            Grid0.Rows.SelectedRows.Clear();
            Grid0.Rows.SelectedRows.Add(pVal.Row);

        }

        private void Form_CloseAfter(SBOItemEventArg pVal)
        {
            SAPbouiCOM.Framework.Application.SBO_Application.Forms.ActiveForm.EnableMenu("RmvRow", false);
            SAPbouiCOM.Framework.Application.SBO_Application.Forms.ActiveForm.EnableMenu("AddRow", false);
        }

        private void Grid0_ValidateAfter(object sboObject, SBOItemEventArg pVal)
        {
            CalculateGridVariables(pVal);
        }

        private void CalculateGridVariables(SBOItemEventArg pVal)
        {
            Thread.CurrentThread.CurrentCulture = CultureInfo.InvariantCulture;

            if (pVal.ColUID == "Quantity")
            {
                double quantity = double.Parse(Grid0.DataTable.GetValue("Quantity", pVal.Row).ToString());
                double totalQuantity = Math.Round(double.Parse(EditText3.Value) * quantity, 3);
                Grid0.DataTable.SetValue("TotalQuantity", pVal.Row, totalQuantity);
            }
            else if (pVal.ColUID == "TotalQuantity")
            {
                double totalQuantity = double.Parse(Grid0.DataTable.GetValue("TotalQuantity", pVal.Row).ToString());
                // double totalCost = double.Parse(Grid0.DataTable.GetValue("TotalCost", pVal.Row).ToString());
                double cost = double.Parse(Grid0.DataTable.GetValue("Cost", pVal.Row).ToString());
                double quantity = Math.Round(totalQuantity / double.Parse(EditText3.Value), 3);
                double totoalCost = Math.Round(totalQuantity * cost, 2);
                Grid0.DataTable.SetValue("Quantity", pVal.Row, quantity);
                Grid0.DataTable.SetValue("TotalCost", pVal.Row, totoalCost);
            }
            else if (pVal.ColUID == "Cost")
            {
                double cost = double.Parse(Grid0.DataTable.GetValue("Cost", pVal.Row).ToString());
                double totalQuantity = double.Parse(Grid0.DataTable.GetValue("TotalQuantity", pVal.Row).ToString());
                double totalCost = Math.Round(cost * totalQuantity, 2);
                Grid0.DataTable.SetValue("TotalCost", pVal.Row, totalCost);
            }
            else if (pVal.ColUID == "TotalCost")
            {
                double totalCost = double.Parse(Grid0.DataTable.GetValue("TotalCost", pVal.Row).ToString());
                double totalQuantity = double.Parse(Grid0.DataTable.GetValue("TotalQuantity", pVal.Row).ToString());
                double cost = Math.Round(totalCost / totalQuantity, 6);
                Grid0.DataTable.SetValue("Cost", pVal.Row, cost);
                CalculateDifference();
            }
        }
    }
}
