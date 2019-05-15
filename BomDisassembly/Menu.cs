using System;
using System.Collections.Generic;
using System.Text;
using BomDisassembly.Forms;
using SAPbouiCOM.Framework;
using BoOrderType = SAPbouiCOM.BoOrderType;
using Grid = SAPbouiCOM.Grid;
using ItemEvent = SAPbouiCOM.ItemEvent;
using SBOItemEventArg = SAPbouiCOM.SBOItemEventArg;

namespace BomDisassembly
{
    class Menu
    {
        public void AddMenuItems()
        {
            SAPbouiCOM.Menus oMenus = null;
            SAPbouiCOM.MenuItem oMenuItem = null;

            oMenus = Application.SBO_Application.Menus;

            SAPbouiCOM.MenuCreationParams oCreationPackage = null;
            oCreationPackage = ((SAPbouiCOM.MenuCreationParams)(Application.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)));
            oMenuItem = Application.SBO_Application.Menus.Item("43520"); // moudles' 

            oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_POPUP;
            oCreationPackage.UniqueID = "BomDisassembly";
            oCreationPackage.String = "BomDisassembly";
            oCreationPackage.Enabled = true;
            oCreationPackage.Position = -1;



            oMenus = oMenuItem.SubMenus;

            try
            {
                //  If the manu already exists this code will fail
                oMenus.AddEx(oCreationPackage);
            }
            catch (Exception e)
            {

            }

            try
            {
                // Get the menu collection of the newly added pop-up item
                oMenuItem = Application.SBO_Application.Menus.Item("BomDisassembly");
                oMenus = oMenuItem.SubMenus;

                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                oCreationPackage.UniqueID = "BomDisassembly.Forms.DissassemblyBomList";
                oCreationPackage.String = "Disassembly Bom List";
                oMenus.AddEx(oCreationPackage);

                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                oCreationPackage.UniqueID = "BomDisassembly.Forms.DisassemblyBom";
                oCreationPackage.String = "Disassembly Bom";
                oMenus.AddEx(oCreationPackage);
                // Create s sub menu
                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                oCreationPackage.UniqueID = "BomDisassembly.Form1";
                oCreationPackage.String = "Form1";
                oMenus.AddEx(oCreationPackage);
            }
            catch (Exception er)
            { //  Menu already exists
                Application.SBO_Application.SetStatusBarMessage("Menu Already Exists", SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
        }


        public void SBO_Application_MenuEvent(ref SAPbouiCOM.MenuEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            try
            {
                SAPbouiCOM.Form activeFormEvent = SAPbouiCOM.Framework.Application.SBO_Application.Forms.ActiveForm;

                if (activeFormEvent.Title == "Dissembly Bill Of Material" && pVal.MenuUID == "AddRow" && pVal.BeforeAction == false)
                {
                    if (((Grid)activeFormEvent.Items.Item("Item_18").Specific).Item.Enabled)
                    {
                        ((Grid)activeFormEvent.Items.Item("Item_18").Specific).DataTable.Rows.Add();
                    }
                }
                else if (activeFormEvent.Title == "Dissembly Bill Of Material" && pVal.MenuUID == "RmvRow" && pVal.BeforeAction == false)
                {
                    try
                    {
                        int rwoIndex = ((Grid)activeFormEvent.Items.Item("Item_18").Specific).Rows.SelectedRows.Item(0, BoOrderType.ot_RowOrder);
                        if (((Grid)activeFormEvent.Items.Item("Item_18").Specific).Item.Enabled)
                        {
                            ((Grid)activeFormEvent.Items.Item("Item_18").Specific).DataTable.Rows.Remove(rwoIndex);
                        }
                    }
                    catch (Exception e)
                    {
                        //row not selected
                    }
                }

                else if (pVal.BeforeAction && pVal.MenuUID == "BomDisassembly.Form1")
                {
                    Form1 activeForm = new Form1();
                    activeForm.Show();
                }
                else if (pVal.BeforeAction && pVal.MenuUID == "BomDisassembly.Forms.DisassemblyBom")
                {
                    DisassemblyBom activeForm = new DisassemblyBom();
                    activeForm.Show();
                }
                else if (pVal.BeforeAction && pVal.MenuUID == "BomDisassembly.Forms.DissassemblyBomList")
                {
                    DissassemblyBomList activeForm = new DissassemblyBomList();
                    activeForm.Show();
                }


            }
            catch (Exception ex)
            {
                Application.SBO_Application.MessageBox(ex.ToString(), 1, "Ok", "", "");
            }
        }

    }
}
