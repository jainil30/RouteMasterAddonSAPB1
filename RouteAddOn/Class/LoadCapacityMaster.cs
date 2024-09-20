using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RouteAddOn.Class
{
    class LoadCapacityMaster
    {
        SAPbouiCOM.Form oForm;
        #region Item Event
        public void ItemEvent(ref ItemEvent pVal, ref bool BubbleEvent, int formType)
        {
            try
            {
                oForm = (SAPbouiCOM.Form)ST.SBO_Application.Forms.Item("frmLc");
                switch (pVal.EventType)
                {
                    #region ITEM PRESSED EVENT
                    case BoEventTypes.et_ITEM_PRESSED:
                        try
                        {
                            if (pVal.BeforeAction)
                            {
                                if (pVal.ItemUID == "1")
                                {
                                    if (oForm.Mode == BoFormMode.fm_ADD_MODE || oForm.Mode == BoFormMode.fm_UPDATE_MODE)
                                    {
                                        if (!Validation())
                                        {
                                            BubbleEvent = false;
                                        }
                                    }

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
                                }
                            }
                        }
                        catch (Exception e)
                        {
                            ST.SetText("Item Pressed Event: " + e.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                        }
                        break;

                    #endregion

                    #region et_CLICK Event
                    case SAPbouiCOM.BoEventTypes.et_CLICK:
                        try
                        {
                            if (pVal.BeforeAction)
                            {
                                if (pVal.FormMode != 3)
                                {
                                    if (new[] { "tLCode", "tLName" }.Contains(pVal.ItemUID))
                                    {
                                        ST.SBO_Application.SetStatusBarMessage("You Can't Edit This Field.", BoMessageTime.bmt_Short, false);
                                        BubbleEvent = false;
                                    }
                                }
                            }
                        }
                        catch (Exception e)
                        {
                            ST.SetText("Click Event: " + e.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                        }
                        break;

                    #endregion

                    #region et_KEY_DOWN Event
                    case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:
                        try
                        {
                            if (pVal.BeforeAction)
                            {
                                if (pVal.FormMode != 3)
                                {
                                    if (new[] { "tLCode", "tLName" }.Contains(pVal.ItemUID))
                                    {
                                        ST.SBO_Application.SetStatusBarMessage("You Can't Edit This Field.", BoMessageTime.bmt_Short, false);
                                        BubbleEvent = false;
                                    }
                                }
                            }

                        }
                        catch (Exception ex)
                        {
                            ST.SetText("Key Down Event: " + ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
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

        #region Menu Event
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
                    oForm.Items.Item("tLCode").Enabled = true;
                    oForm.Items.Item("tLCode").Click(BoCellClickType.ct_Regular);
                }
                else
                {
                    DisabledAllFields();
                }
            }
            catch (Exception e)
            {
                ST.SetText("MenuEvent : " + e.Message, BoStatusBarMessageType.smt_Error);
                throw;
            }
        }
        #endregion

        #region Other Methods
        private void SetCode()
        {
            try
            {
                oForm.Freeze(true);
                oForm.Mode = BoFormMode.fm_ADD_MODE;
                //oForm.State = BoFormStateEnum.fs_Maximized;
                string str = oForm.BusinessObject.GetNextSerialNumber("Code", "STLDC").ToString();
                oForm.Items.Item("tCode").Enabled = false;
                ((SAPbouiCOM.EditText)oForm.Items.Item("tCode").Specific).Value = str;
                oForm.Items.Item("tLCode").Click(BoCellClickType.ct_Regular);
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
                oForm.Items.Item("tCode").Enabled = false;
                oForm.Items.Item("tLCode").Enabled = false;
                oForm.Items.Item("tLName").Enabled = false;
            }
            catch (Exception e)
            {
                ST.SetText("DisabledAllFields : " + e.Message, BoStatusBarMessageType.smt_Error);
            }
        }

        private bool Validation()
        {
            try
            {
                if (String.IsNullOrEmpty(((SAPbouiCOM.EditText)oForm.Items.Item("tLCode").Specific).Value))
                {
                    ST.SetText("Load Capacity Code is missing", BoStatusBarMessageType.smt_Error);
                    return false;
                }
                else if (String.IsNullOrEmpty(((SAPbouiCOM.EditText)oForm.Items.Item("tLName").Specific).Value))
                {
                    ST.SetText("Load Capacity Name is missing", BoStatusBarMessageType.smt_Error);
                    return false;
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
        #endregion
    }
}
