using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
namespace RouteAddOn.Class
{
    class RouteMaster
    {
        SAPbouiCOM.Form oForm;
        SAPbouiCOM.Matrix oMat;
        SAPbobsCOM.Recordset oRec;
        string Query;
        bool IsMatrixUpdated;
        #region Item Event
        public void ItemEvent(ref ItemEvent pVal, ref bool bubbleEvent, int formType)
        {
            try
            {
                oForm = (SAPbouiCOM.Form)ST.SBO_Application.Forms.Item("frmRte");

                if (IsMatrixUpdated)
                {
                    return;
                }
                GC.Collect();
                switch (pVal.EventType)
                {

                    #region et_CHOOSE_FROM_LIST Event
                    case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:
                        try
                        {
                            if (pVal.BeforeAction)
                            {
                                if (pVal.ColUID == "mcLc")
                                {
                                    CFL("CFL_STLDC", 0);
                                }
                                else if (pVal.ColUID == "mcAll")
                                {
                                    CFL("CFL_STALW", 0);
                                }
                                else if (pVal.ColUID == "mcCar")
                                {
                                    CFL("CFL_STCAT", 0);
                                }
                            }
                            else
                            {
                                SAPbouiCOM.IChooseFromListEvent oCFLEvento = ((SAPbouiCOM.IChooseFromListEvent)(pVal));
                                SAPbouiCOM.DataTable oDataTable = oCFLEvento.SelectedObjects;
                                oMat = (SAPbouiCOM.Matrix)oForm.Items.Item(pVal.ItemUID).Specific;
                                if (oCFLEvento.SelectedObjects != null && pVal.ActionSuccess)
                                {
                                    if (pVal.ColUID == "mcAll")
                                    {
                                        try
                                        {
                                            IsMatrixUpdated = true;
                                            bubbleEvent = false;
                                            if (oDataTable.Rows.Count > 0)
                                            {
                                                for(int i = 0; i < oDataTable.Rows.Count; i++)
                                                {
                                                    ((SAPbouiCOM.EditText)oMat.Columns.Item("mcAll").Cells.Item(i+1).Specific).Value = oDataTable.GetValue("Code", 0).ToString();
                                                    AddMatrixRow();
                                                }
                                            }

                                          
                                            //((SAPbouiCOM.EditText)oForm.Items.Item("mcAll").Specific).Value = oDataTable.GetValue("Code", 0).ToString();
                                        }
                                        catch { }finally
                                        {
                                            IsMatrixUpdated = false;
                                            bubbleEvent = false;
                                        }

                                    }
                                    else if (pVal.ColUID == "mcLc")
                                    {
                                        if (pVal.ColUID == "mcLc")
                                        {
                                            try
                                            {

                                                ((SAPbouiCOM.EditText)oMat.Columns.Item("mcLc").Cells.Item(pVal.Row).Specific).Value = oDataTable.GetValue("Code", 0).ToString();

                                                //((SAPbouiCOM.EditText)oForm.Items.Item("mcAll").Specific).Value = oDataTable.GetValue("Code", 0).ToString();
                                            }
                                            catch { }

                                        }
                                    }
                                    else if (pVal.ColUID == "mcCar")
                                    {
                                        if (pVal.ColUID == "mcCar")
                                        {
                                            try
                                            {


                                                ((SAPbouiCOM.EditText)oMat.Columns.Item("mcCar").Cells.Item(pVal.Row).Specific).Value = oDataTable.GetValue("Code", 0).ToString();

                                                //((SAPbouiCOM.EditText)oForm.Items.Item("mcAll").Specific).Value = oDataTable.GetValue("Code", 0).ToString();
                                            }
                                            catch { }

                                        }
                                    }
                                }

                            }
                        }
                        catch (Exception e)
                        {
                            ST.SetText("et_CHOOSE_FROM_LIST : " + e.Message, BoStatusBarMessageType.smt_Success);
                        }
                        break;

                    #endregion

                    #region Click Event
                    case BoEventTypes.et_CLICK:

                        try
                        {
                            if (pVal.BeforeAction)
                            {
                                if (pVal.FormMode == 1 || pVal.FormMode == 2)
                                {
                                    if (new[] { "tRCode", "tRName", "cType" }.Contains(pVal.ItemUID))
                                    {
                                        ST.SBO_Application.SetStatusBarMessage("You Can't Edit This Field.", BoMessageTime.bmt_Short, false);
                                        bubbleEvent = false;
                                    }

                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            ST.SetText("Click Event: " + ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                        }
                        break;

                    #endregion

                    #region Item Pressed Event
                    case BoEventTypes.et_ITEM_PRESSED:
                        if (pVal.BeforeAction)

                        {
                            if (pVal.ItemUID == "1" && (oForm.Mode == BoFormMode.fm_ADD_MODE || oForm.Mode == BoFormMode.fm_UPDATE_MODE))
                            {
                                if (Validation() == false)
                                {
                                    bubbleEvent = false;
                                }
                                ////ST.SetText("Item Pressed BeforeAction", BoStatusBarMessageType.smt_Success);
                            }

                        }
                        else
                        {
                            if (pVal.ItemUID == "1")
                            {
                                if (oForm.Mode == BoFormMode.fm_ADD_MODE)
                                {
                                    SetCode();
                                }
                                else if (oForm.Mode == BoFormMode.fm_OK_MODE)
                                {
                                    EditableAfterAdd();
                                }
                            }
                            //ST.SetText("Item Pressed AfterAction", BoStatusBarMessageType.smt_Success);
                        }




                        break;
                    #endregion

                    #region et_MATRIX_LINK_PRESSED Event
                    case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED:
                        try
                        {
                            if (!pVal.BeforeAction)
                            {
                                if (pVal.ItemUID == "mat")
                                {
                                    if (pVal.ColUID == "mcAll")
                                    {
                                        string AllowanceCode = ((SAPbouiCOM.EditText)oMat.Columns.Item("mcAll").Cells.Item(pVal.Row).Specific).Value.ToString();
                                        oRec = (SAPbobsCOM.Recordset)ST.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                        Query = "SELECT T0.\"Code\" FROM \"@STALW\" T0 WHERE T0.\"Code\" = '" + AllowanceCode + "'";
                                        oRec.DoQuery(Query);
                                        if (oRec.RecordCount > 0)
                                        {
                                            if (ST.ExistingForm("frmAlw"))
                                            {
                                                ST.LoadFromXML("frmAlw");
                                                oForm = ST.SBO_Application.Forms.ActiveForm;
                                                oForm.Freeze(true);
                                                oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                                                oForm.Items.Item("tCode").Enabled = true;
                                                ((SAPbouiCOM.EditText)oForm.Items.Item("tCode").Specific).Value = oRec.Fields.Item("Code").Value.ToString();
                                                oForm.Items.Item("1").Click();
                                                oForm.Items.Item("tCode").Enabled = false;
                                                oForm.Freeze(false);
                                            }
                                        }
                                    }else if (pVal.ColUID == "mcCar")
                                    {
                                        string CargoTypeCode = ((SAPbouiCOM.EditText)oMat.Columns.Item("mcCar").Cells.Item(pVal.Row).Specific).Value.ToString();
                                        oRec = (SAPbobsCOM.Recordset)ST.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                        Query = "SELECT T0.\"Code\" FROM \"@STCAT\" T0 WHERE T0.\"Code\" = '" + CargoTypeCode + "'";
                                        oRec.DoQuery(Query);
                                        if (oRec.RecordCount > 0)
                                        {
                                            if (ST.ExistingForm("frmCt"))
                                            {
                                                ST.LoadFromXML("frmCt");
                                                oForm = ST.SBO_Application.Forms.ActiveForm;
                                                oForm.Freeze(true);
                                                oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                                                oForm.Items.Item("tCode").Enabled = true;
                                                ((SAPbouiCOM.EditText)oForm.Items.Item("tCode").Specific).Value = oRec.Fields.Item("Code").Value.ToString();
                                                oForm.Items.Item("1").Click();
                                                oForm.Items.Item("tCode").Enabled = false;
                                                oForm.Freeze(false);
                                            }
                                        }
                                    }
                                    else if (pVal.ColUID == "mcLc")
                                    {
                                        string LoadCapacity = ((SAPbouiCOM.EditText)oMat.Columns.Item("mcLc").Cells.Item(pVal.Row).Specific).Value.ToString();
                                        oRec = (SAPbobsCOM.Recordset)ST.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                        Query = "SELECT T0.\"Code\" FROM \"@STLDC\" T0 WHERE T0.\"Code\" = '" + LoadCapacity + "'";
                                        oRec.DoQuery(Query);
                                        if (oRec.RecordCount > 0)
                                        {
                                            if (ST.ExistingForm("frmLc"))
                                            {
                                                ST.LoadFromXML("frmLc");
                                                oForm = ST.SBO_Application.Forms.ActiveForm;
                                                oForm.Freeze(true);
                                                oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                                                oForm.Items.Item("tCode").Enabled = true;
                                                ((SAPbouiCOM.EditText)oForm.Items.Item("tCode").Specific).Value = oRec.Fields.Item("Code").Value.ToString();
                                                oForm.Items.Item("1").Click();
                                                oForm.Items.Item("tCode").Enabled = false;
                                                oForm.Freeze(false);
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            ST.SetText("Matrix Linked Pressed Event: " + ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                        }
                        break;

                        #endregion
                }
            }
            catch (Exception e)
            {
                ST.SetText("ItemEvent : " + e.Message, BoStatusBarMessageType.smt_Error);
            }
        }
        #endregion

        #region Menu Click Event
        public void MenuEvent(ref MenuEvent pVal, string uniqueID, string type)
        {
            try
            {
                oForm = (SAPbouiCOM.Form)ST.SBO_Application.Forms.ActiveForm;
                if (type == "Add")
                {
                    SetCode();
                }
                else if (type == "Find")
                {
                    DisabledAllFields();
                    oForm.Items.Item("tRCode").Enabled = true;
                }
                else if (type == "AddR")
                {
                    AddMatrixRow();
                }

                else if (type == "DelR")
                {
                    DeleteRow1();
                }
            }
            catch (Exception e)
            {

                ST.SetText("MenuEvent : " + e.Message, BoStatusBarMessageType.smt_Error);
            }
        }
        #endregion

        #region Right Click Event
        public void RightClickEvent(ref ContextMenuInfo eventInfo, ref bool bubbleEvent, string FormUID)
        {
            try
            {
                if (eventInfo.FormUID == FormUID && eventInfo.BeforeAction == true)
                {
                    ST.DELLINE = eventInfo.Row;
                    oForm = ST.SBO_Application.Forms.Item(FormUID);
                    oForm.EnableMenu("1292", false);                //Add Row
                    oForm.EnableMenu("1293", false);                //Delete Row
                    oForm.EnableMenu("1294", false);                //Duplicate Row
                    oForm.EnableMenu("1299", false);                //Close Row
                    oForm.EnableMenu("1287", false);                //Close Row
                    oForm.EnableMenu("771", false);                 //Cut

                    oForm.EnableMenu("772", true);                  //Copy
                    oForm.EnableMenu("773", true);                 //Paste Row

                    oForm.EnableMenu("774", false);                 //Delete Row
                    oForm.EnableMenu("784", false);                 //Copy Table
                    oForm.EnableMenu("1287", true);                //Duplicate
                    oForm.EnableMenu("1283", false);                //Remove                

                    if (eventInfo.ItemUID == "mat")
                    {
                        oForm.EnableMenu("1292", true);             //Add Row
                        oForm.EnableMenu("1293", true);             //Delete Row
                        oForm.EnableMenu("784", false);             //Copy Table
                        oForm.EnableMenu("1287", false);            //Copy Table

                        //if (eventInfo.ColUID == "cPCODE")
                        //{
                        //    oForm.EnableMenu("5377", false);         //list of CFL
                        //}
                    }

                    if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE) oForm.EnableMenu("773", true);//Paste Row                
                }
            }
            catch (Exception ex)
            {
                ST.SBO_Application.SetStatusBarMessage("Right Click Event: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short);
            }
        }
        #endregion

        #region Other Methods
        private bool Validation()
        {
            try
            {
                oMat = ((SAPbouiCOM.Matrix)(oForm.Items.Item("mat").Specific));
                if (String.IsNullOrEmpty(((SAPbouiCOM.EditText)oForm.Items.Item("tRCode").Specific).Value))
                {
                    oForm.Items.Item("tRCode").Click(BoCellClickType.ct_Regular);
                    ST.SetText("Route Code is missing", BoStatusBarMessageType.smt_Error);
                    return false;
                }
                else if (String.IsNullOrEmpty(((SAPbouiCOM.EditText)oForm.Items.Item("tRName").Specific).Value))
                {
                    oForm.Items.Item("tRName").Click(BoCellClickType.ct_Regular);
                    ST.SetText("Route Name is missing", BoStatusBarMessageType.smt_Error);
                    return false;
                }
                else if (String.IsNullOrEmpty(((SAPbouiCOM.EditText)oForm.Items.Item("tLoLc").Specific).Value))
                {
                    oForm.Items.Item("tLoLc").Click(BoCellClickType.ct_Regular);
                    ST.SetText("Loading Location is missing", BoStatusBarMessageType.smt_Error);
                    return false;
                }
                else if (String.IsNullOrEmpty(((SAPbouiCOM.EditText)oForm.Items.Item("tDes").Specific).Value))
                {
                    oForm.Items.Item("tDes").Click(BoCellClickType.ct_Regular);
                    ST.SetText("Destination is missing", BoStatusBarMessageType.smt_Error);
                    return false;
                }
                else if (String.IsNullOrEmpty(((SAPbouiCOM.EditText)oForm.Items.Item("tDis").Specific).Value))
                {
                    oForm.Items.Item("tDis").Click(BoCellClickType.ct_Regular);
                    ST.SetText("Distance is missing", BoStatusBarMessageType.smt_Error);
                    return false;
                }
                else if (Convert.ToInt32(((SAPbouiCOM.EditText)oForm.Items.Item("tDis").Specific).Value) < 0)
                {
                    oForm.Items.Item("tDis").Click(BoCellClickType.ct_Regular);
                    ST.SetText("Distance cannot be negative", BoStatusBarMessageType.smt_Error);
                    return false;
                }
                else if (String.IsNullOrEmpty(((SAPbouiCOM.EditText)oForm.Items.Item("tEth").Specific).Value))
                {
                    oForm.Items.Item("tEth").Click(BoCellClickType.ct_Regular);
                    ST.SetText("Estimated Trip Hours is missing", BoStatusBarMessageType.smt_Error);
                    return false;
                }
                else if (Convert.ToInt32(((SAPbouiCOM.EditText)oForm.Items.Item("tEth").Specific).Value) < 0)
                {
                    oForm.Items.Item("tEth").Click(BoCellClickType.ct_Regular);
                    ST.SetText("Estimated Trip Hours cannot be negative", BoStatusBarMessageType.smt_Error);
                    return false;
                }
                else if (String.IsNullOrEmpty(((SAPbouiCOM.EditText)oForm.Items.Item("tTpm").Specific).Value))
                {
                    oForm.Items.Item("tTpm").Click(BoCellClickType.ct_Regular);
                    ST.SetText("Turnaround Per Month is missing", BoStatusBarMessageType.smt_Error);
                    return false;
                }
                else if (Convert.ToInt32(((SAPbouiCOM.EditText)oForm.Items.Item("tTpm").Specific).Value) < 0)
                {
                    ST.SetText("Turnaround Per Month cannot be negative", BoStatusBarMessageType.smt_Error);
                    oForm.Items.Item("tTpm").Click(BoCellClickType.ct_Regular);
                    return false;
                }
                else if (String.IsNullOrEmpty(((SAPbouiCOM.ComboBox)oForm.Items.Item("cType").Specific).Value))
                {
                    oForm.Items.Item("cType").Click(BoCellClickType.ct_Regular);
                    ST.SetText("Type of Route is missing", BoStatusBarMessageType.smt_Error);
                    return false;
                }
                else
                 if (oMat.RowCount > 0)
                {
                    for (int i = 1; i <= oMat.RowCount; i++)
                    {
                        if (string.IsNullOrEmpty(((SAPbouiCOM.EditText)oMat.Columns.Item("mcAll").Cells.Item(i).Specific).Value.ToString()))
                        {
                            
                            oMat.Columns.Item("mcAll").Cells.Item(i).Click(BoCellClickType.ct_Regular);
                            ST.SetText("Allowance is missing at row number " + i, BoStatusBarMessageType.smt_Error);
                            return false;
                        }
                        else if (string.IsNullOrEmpty(((SAPbouiCOM.EditText)oMat.Columns.Item("mtAmt").Cells.Item(i).Specific).Value.ToString()))
                        {
                            oMat.Columns.Item("mtAmt").Cells.Item(i).Click(BoCellClickType.ct_Regular);
                            ST.SetText("Amount is missing at row number " + i, BoStatusBarMessageType.smt_Error);
                            return false;

                        }
                        else if (Convert.ToDecimal(((SAPbouiCOM.EditText)oMat.Columns.Item("mtAmt").Cells.Item(i).Specific).Value.ToString()) < 0)
                        {
                            oMat.Columns.Item("mtAmt").Cells.Item(i).Click(BoCellClickType.ct_Regular);
                            ST.SetText("Amount cannot be negative", BoStatusBarMessageType.smt_Error);
                            return false;
                        }
                        else if (string.IsNullOrEmpty(((SAPbouiCOM.EditText)oMat.Columns.Item("mcLc").Cells.Item(i).Specific).Value.ToString()))
                        {
                            oMat.Columns.Item("mcLc").Cells.Item(i).Click(BoCellClickType.ct_Regular);
                            ST.SetText("Load Capacity is missing at row number " + i, BoStatusBarMessageType.smt_Error);
                            return false;

                        }
                        else if (string.IsNullOrEmpty(((SAPbouiCOM.EditText)oMat.Columns.Item("mcCar").Cells.Item(i).Specific).Value.ToString()))
                        {
                            oMat.Columns.Item("mcCar").Cells.Item(i).Click(BoCellClickType.ct_Regular);
                            ST.SetText("Cargo Type is missing at row number " + i, BoStatusBarMessageType.smt_Error);
                            return false;

                        }
                    }
                }
                else
                {
                    ST.SetText("Validated", BoStatusBarMessageType.smt_Success);
                }

                return true;
            }
            catch (Exception e)
            {
                ST.SetText("Validation : " + e.Message, BoStatusBarMessageType.smt_Error);
            }
            return true;
        }
        private void SetCode()
        {
            try
            {

                int nextNumber = oForm.BusinessObject.GetNextSerialNumber("DocEntry", "STRTE");
                //string nextRCode = String.Format("R{0:D5}", nextNumber);
                oForm.Freeze(true);
                oForm.Mode = BoFormMode.fm_ADD_MODE;
                oForm.State = BoFormStateEnum.fs_Maximized;



                ((SAPbouiCOM.ComboBox)oForm.Items.Item("cType").Specific).Select(0, BoSearchKey.psk_Index);
                ((SAPbouiCOM.ComboBox)oForm.Items.Item("cType").Specific).ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly;

                oForm.Items.Item("tCode").Enabled = false;
                //oForm.Items.Item("tRCode").Enabled = false;

                ((SAPbouiCOM.EditText)oForm.Items.Item("tCode").Specific).Value = nextNumber.ToString();
                //((SAPbouiCOM.EditText)oForm.Items.Item("tRCode").Specific).Value = nextRCode;

                oMat = ((SAPbouiCOM.Matrix)oForm.Items.Item("mat").Specific);
                oMat.AutoResizeColumns();
                AddMatrixRow();


                //oMat.Columns.Item("mcLc").Editable = true;
                //oMat.Columns.Item("mcAll").Editable = true;
                //oMat.Columns.Item("mcCar").Editable = true;
                //oMat.Columns.Item("mtAmt").Editable = true;

                oForm.Items.Item("tRCode").Click(BoCellClickType.ct_Regular);
                oForm.Freeze(false);
            }
            catch (Exception e)
            {
                ST.SetText("SetCode : " + e.Message, BoStatusBarMessageType.smt_Error);
            }
        }
        private void DisabledAllFields()
        {
            try
            {
                oMat = ((SAPbouiCOM.Matrix)(oForm.Items.Item("mat").Specific));
                oForm.Items.Item("tRName").Enabled = false;
                oForm.Items.Item("tLoLc").Enabled = false;
                oForm.Items.Item("tDes").Enabled = false;
                oForm.Items.Item("tDis").Enabled = false;
                oForm.Items.Item("tEth").Enabled = false;
                oForm.Items.Item("tTpm").Enabled = false;
                oForm.Items.Item("cType").Enabled = false;

                oForm.Items.Item("tRCode").Enabled = false;
                oForm.Items.Item("tCode").Enabled = false;


                oMat.Item.Enabled = false;
            }
            catch (Exception e)
            {
                ST.SetText("DisabledAllFields : " + e.Message, BoStatusBarMessageType.smt_Error);
            }
        }

        private void EditableAfterAdd()
        {
            try
            {

                //Editable After Add - fields
                oForm.Items.Item("tLoLc").Enabled = true;
                oForm.Items.Item("tDes").Enabled = true;
                oForm.Items.Item("tDis").Enabled = true;
                oForm.Items.Item("tEth").Enabled = true;
                oForm.Items.Item("tTpm").Enabled = true;

                //Not Editable After Add - fields
                oForm.Items.Item("tRName").Enabled = false;
                oForm.Items.Item("cType").Enabled = false;
                oForm.Items.Item("tRCode").Enabled = false;
                oForm.Items.Item("tCode").Enabled = false;


                //Fields of row table are not mentioned here
                //Update it once you start working on it
            }
            catch (Exception e)
            {
                ST.SetText("DisabledAllFields : " + e.Message, BoStatusBarMessageType.smt_Error);
            }

        }


        private bool CFL(string CFLID, Int32 Row)
        {
            try
            {
                SAPbouiCOM.Condition oCond = default(SAPbouiCOM.Condition);
                SAPbouiCOM.Conditions oConds = default(SAPbouiCOM.Conditions);
                SAPbouiCOM.ChooseFromList oCFL = default(SAPbouiCOM.ChooseFromList);
                oCFL = oForm.ChooseFromLists.Item(CFLID);
                oConds = (SAPbouiCOM.Conditions)ST.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_Conditions);

                if (CFLID == "CFL_STALW")
                {
                    oRec = (SAPbobsCOM.Recordset)ST.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    Query = "SELECT T0.\"Code\",T0.\"U_ACODE\",T0.\"U_ANAME\" FROM \"@STALW\" T0 ";
                    oRec.DoQuery(Query);
                    if (oRec.RecordCount > 0)
                    {
                        for (int i = 1; i <= oRec.RecordCount; i++)
                        {
                            oCond = oConds.Add();
                            oCond.Alias = "Code";
                            oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                            oCond.CondVal = oRec.Fields.Item("Code").Value.ToString();
                            if (i != oRec.RecordCount)
                                oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR;
                            oRec.MoveNext();
                        }
                    }
                    else
                    {
                        oCond = oConds.Add();
                        oCond.Alias = "";
                        oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                        oCond.CondVal = "";
                    }
                    oCFL.SetConditions(oConds);
                }
                else if (CFLID == "CFL_STLDC")
                {
                    oRec = (SAPbobsCOM.Recordset)ST.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    Query = "SELECT T0.\"Code\",T0.\"U_LCODE\",T0.\"U_LNAME\"  FROM \"@STLDC\" T0 ";
                    oRec.DoQuery(Query);
                    if (oRec.RecordCount > 0)
                    {
                        for (int i = 1; i <= oRec.RecordCount; i++)
                        {
                            oCond = oConds.Add();
                            oCond.Alias = "";
                            oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                            oCond.CondVal = oRec.Fields.Item(0).Value.ToString();
                            if (i != oRec.RecordCount)
                                oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR;
                            oRec.MoveNext();
                        }
                    }
                    else
                    {
                        oCond = oConds.Add();
                        oCond.Alias = "";
                        oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                        oCond.CondVal = "";
                    }
                    oCFL.SetConditions(oConds);
                }
                else if (CFLID == "CFL_STCAT")
                {
                    oRec = (SAPbobsCOM.Recordset)ST.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    Query = "SELECT T0.\"Code\",T0.\"U_CTCODE\",T0.\"U_CTNAME\"  FROM \"@STCAT\" T0 ";
                    oRec.DoQuery(Query);
                    if (oRec.RecordCount > 0)
                    {
                        for (int i = 1; i <= oRec.RecordCount; i++)
                        {
                            oCond = oConds.Add();
                            oCond.Alias = "";
                            oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                            oCond.CondVal = oRec.Fields.Item(0).Value.ToString();
                            if (i != oRec.RecordCount)
                                oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR;
                            oRec.MoveNext();
                        }
                    }
                    else
                    {
                        oCond = oConds.Add();
                        oCond.Alias = "";
                        oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                        oCond.CondVal = "";
                    }
                    oCFL.SetConditions(oConds);
                }

                return true;
            }
            catch (Exception ex)
            {
                ST.SBO_Application.SetStatusBarMessage("CFL Before Action: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short);
            }
            return true;
        }
        private void DeleteRow1()
        {
            try
            {
                oMat = ((SAPbouiCOM.Matrix)(oForm.Items.Item("mat").Specific));

                if (oMat.RowCount == 1)
                {
                    oMat.ClearRowData(1);
                }
                else
                {
                    oMat.DeleteRow(ST.DELLINE);
                }
                for (int j = 1; j <= oMat.RowCount; j++)
                {
                    ((SAPbouiCOM.EditText)oMat.Columns.Item("#").Cells.Item(j).Specific).Value = j.ToString();
                }
                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                {
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                }
            }
            catch (Exception ex)
            {
                ST.SBO_Application.SetStatusBarMessage("Delete Row: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short);
            }
        }

        private void AddMatrixRow()
        {
            try
            {
                oMat = ((SAPbouiCOM.Matrix)(oForm.Items.Item("mat").Specific));
                if (oMat.RowCount == 0)
                {
                    oMat.AddRow();
                    ((SAPbouiCOM.EditText)oMat.Columns.Item("#").Cells.Item(1).Specific).Value = "1";
                }
                else
                {
                    if (((SAPbouiCOM.EditText)oMat.Columns.Item("mcAll").Cells.Item(oMat.RowCount).Specific).Value != "")
                    {
                        oMat.AddRow();
                        int I = oMat.RowCount;
                        oMat.ClearRowData(I);
                        ((SAPbouiCOM.EditText)oMat.Columns.Item("#").Cells.Item(I).Specific).Value = I.ToString();
                    }
                }
                 oMat.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                ST.SBO_Application.SetStatusBarMessage("Add Matrix Row: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short);
            }
        }

        private void DeleteMatrixBlankRow()
        {
            try
            {
                oMat = ((SAPbouiCOM.Matrix)(oForm.Items.Item("mat").Specific));
                for (int i = oMat.RowCount; i >= 1; i--)
                {

                    if (string.IsNullOrEmpty(((SAPbouiCOM.EditText)oMat.Columns.Item("cPCODE").Cells.Item(i).Specific).Value))
                    {
                        oMat.DeleteRow(i);
                    }

                }
            }
            catch (Exception ex)
            {
                ST.SBO_Application.SetStatusBarMessage("Delete Matrix Blank Row: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short);
            }
        }
        #endregion
    }
}
