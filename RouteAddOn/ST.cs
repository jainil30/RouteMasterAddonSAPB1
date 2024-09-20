using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RouteAddOn
{
    class ST
    {

        #region Variable
        public static SAPbouiCOM.Application SBO_Application;
        public static SAPbobsCOM.Company oCompany;
        private static SAPbouiCOM.Form oForm = default(SAPbouiCOM.Form);
        private static SAPbobsCOM.Recordset oRecordSet = default(SAPbobsCOM.Recordset);
        private SAPbouiCOM.MenuItem oMenuItem;
        private SAPbouiCOM.Menus oMenus;
        private SAPbouiCOM.MenuCreationParams oCreationPackage = null;
        public static string Qry = "";
        private SAPbouiCOM.EventFilters oFilters;
        private SAPbouiCOM.EventFilter oFilter;
        public static bool FORMCLOSE = false;
        public static int DELLINE = 0;
        public static string l = null;
        string[] FindColoums = null;
        string[] ChildTable = null;
        private string ErrMsg = null;
        private int RetCode = 0;
        public static string Query_Item = null;
        string[,] str;
        Class.RouteMaster RM = new Class.RouteMaster();
        Class.AllowanceMaster AM = new Class.AllowanceMaster();
        Class.LoadCapacityMaster LC = new Class.LoadCapacityMaster();
        Class.CargoTypeMaster CT = new Class.CargoTypeMaster();

        #endregion

        #region Constructor
        public ST()
        {
            try
            {
                SBO_Application = SAPbouiCOM.Framework.Application.SBO_Application;
                oCompany = ((SAPbobsCOM.Company)SBO_Application.Company.GetDICompany());

                SetFilters();
                SBO_Application.ItemEvent += new SAPbouiCOM._IApplicationEvents_ItemEventEventHandler(SBO_Application_ItemEvent);
                SBO_Application.MenuEvent += new SAPbouiCOM._IApplicationEvents_MenuEventEventHandler(SBO_Application_MenuEvent);
                SBO_Application.AppEvent += new SAPbouiCOM._IApplicationEvents_AppEventEventHandler(SBO_Application_AppEvent);
                SBO_Application.RightClickEvent += new SAPbouiCOM._IApplicationEvents_RightClickEventEventHandler(SBO_Application_RightClickEvent);

                CreateDBStructure();
                CreateAddonMenu();
                SBO_Application.StatusBar.SetText("Add-on Route Master is Connected..!!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
            }
            catch (Exception ex)
            {
                SBO_Application.StatusBar.SetText(ex.Message.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short);
            }
        }
        #endregion

        #region APP Event
        private void SBO_Application_AppEvent(SAPbouiCOM.BoAppEventTypes EventType)
        {
            try
            {
                switch (EventType)
                {
                    case SAPbouiCOM.BoAppEventTypes.aet_ShutDown:
                    case SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition:
                        {
                            System.Windows.Forms.Application.Exit();
                        }
                        break;
                    case SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged:
                        {
                            SBO_Application.SetStatusBarMessage("Company Change", SAPbouiCOM.BoMessageTime.bmt_Short, false);
                            System.Windows.Forms.Application.Exit();
                        }
                        break;
                }
            }
            catch (Exception e)
            {
                SetText("SBO_Application_AppEvent : " + e.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }
        #endregion

        #region ItemEvent
        private void SBO_Application_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                GC.Collect();
                switch (FormUID)
                {
                    case "frmRte":
                        RM.ItemEvent(ref pVal, ref BubbleEvent, pVal.FormType);
                        break;

                    case "frmAlw":
                        AM.ItemEvent(ref pVal, ref BubbleEvent, pVal.FormType);
                        break;

                    case "frmLc":
                        LC.ItemEvent(ref pVal, ref BubbleEvent, pVal.FormType);
                        break;

                    case "frmCt":
                        CT.ItemEvent(ref pVal, ref BubbleEvent, pVal.FormType);
                        break;
                }

                if (pVal.BeforeAction)
                {
                    if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD)
                    {
                        if (FORMCLOSE == true)
                        {
                            BubbleEvent = false;
                            FORMCLOSE = false;
                            int Msg = SBO_Application.MessageBox("Do you want to save the changes?", 1, "Yes", "No", "Cancel");
                            if (Msg == 1 || Msg == 2)
                            {
                                SBO_Application.Forms.Item(FormUID).Items.Item("2").Click();
                            }
                        }
                    }
                }
                else
                {
                    ST.SetText("Testing", SAPbouiCOM.BoStatusBarMessageType.smt_None);
                }
            }
            catch (Exception e){
                SetText("SBO_Application_ItemEvent : " + e.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }
        #endregion

        #region MenuEvents
        public void SBO_Application_MenuEvent(ref SAPbouiCOM.MenuEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                GC.Collect();
                if (pVal.BeforeAction)
                {
                    #region Add Mode
                    if (pVal.MenuUID == "1282")
                    {
                        oForm = SBO_Application.Forms.ActiveForm;
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            BubbleEvent = false;
                        }
                    }
                    #endregion

                    #region Delete Row
                    if (pVal.MenuUID == "1293")
                    {
                        oForm = SBO_Application.Forms.ActiveForm;
                        switch (oForm.UniqueID)
                        {
                            case "frmRte":
                                RM.MenuEvent(ref pVal, oForm.UniqueID, "DelR");
                                BubbleEvent = false;
                                break;
                        }


                    }
                    #endregion
                }

                if (!pVal.BeforeAction)
                {
                    #region Add Mode
                    if (pVal.MenuUID == "1282")
                    {
                        oForm = SBO_Application.Forms.ActiveForm;
                        switch (oForm.UniqueID)
                        {
                            case "frmRte":
                                RM.MenuEvent(ref pVal, oForm.UniqueID, "Add");
                                break;
                            case "frmAlw":
                                AM.MenuEvent(ref pVal, oForm.UniqueID, "Add");
                                break;
                            case "frmLc":
                                LC.MenuEvent(ref pVal, oForm.UniqueID, "Add");
                                break;
                            case "frmCt":
                                CT.MenuEvent(ref pVal, oForm.UniqueID, "Add");
                                break;
                        }
                    }
                    #endregion

                    #region Duplicate Mode
                    if (pVal.MenuUID == "1287")
                    {
                        oForm = SBO_Application.Forms.ActiveForm;
                        switch (oForm.UniqueID)
                        {
                            case "frmRte":
                                RM.MenuEvent(ref pVal, oForm.UniqueID, "Dupl");
                                break;

                            
                        }
                    }
                    #endregion

                    #region Find Mode
                    if (pVal.MenuUID == "1281")
                    {
                        oForm = SBO_Application.Forms.ActiveForm;
                        switch (oForm.UniqueID)
                        {
                            case "frmRte":
                                RM.MenuEvent(ref pVal, oForm.UniqueID, "Find");
                                break;
                            case "frmAlw":
                                AM.MenuEvent(ref pVal, oForm.UniqueID, "Find");
                                break;
                            case "frmLc":
                                LC.MenuEvent(ref pVal, oForm.UniqueID, "Find");
                                break;
                            case "frmCt":
                                CT.MenuEvent(ref pVal, oForm.UniqueID, "Find");
                                break;
                        }

                    }
                    #endregion                

                    #region Navigation
                    if (pVal.MenuUID == "1288" || pVal.MenuUID == "1289" || pVal.MenuUID == "1290" || pVal.MenuUID == "1291")
                    {
                        oForm = SBO_Application.Forms.ActiveForm;
                        switch (oForm.UniqueID)
                        {
                            case "frmRte":
                                RM.MenuEvent(ref pVal, oForm.UniqueID, "Nav");
                                break;
                            case "frmAlw":
                                AM.MenuEvent(ref pVal, oForm.UniqueID, "Nav");
                                break;
                            case "frmLc":
                                LC.MenuEvent(ref pVal, oForm.UniqueID, "Nav");
                                break;
                            case "frmCt":
                                CT.MenuEvent(ref pVal, oForm.UniqueID, "Nav");
                                break;
                        }
                    }
                    #endregion

                    #region Add Row
                    if (pVal.MenuUID == "1292")
                    {
                        oForm = SBO_Application.Forms.ActiveForm;
                        switch (oForm.UniqueID)
                        {
                            case "frmRte":
                                RM.MenuEvent(ref pVal, oForm.UniqueID, "AddR");
                                break;
                        }
                    }
                    #endregion

                    #region Form Open    

                    #region Route Master
                    else if (pVal.MenuUID == "mRoute")
                    {
                        if (ExistingForm("frmRte"))
                        {
                            LoadFromXML("frmRte");
                            oForm.DataBrowser.BrowseBy = "tCode";
                            //oForm.EnableMenu("1288", true);
                            oForm.EnableMenu("1289", true);
                            oForm.EnableMenu("1290", true);
                            oForm.EnableMenu("1291", true);
                            oForm.EnableMenu("1287", true);
                            oForm.EnableMenu("6913", false);
                            oForm.EnableMenu("1304", true);
                            RM.MenuEvent(ref pVal, oForm.UniqueID, "Add");
                        }
                    }
                    else if (pVal.MenuUID == "mAlw")
                    {
                        if (ExistingForm("frmAlw"))
                        {
                            LoadFromXML("frmAlw");
                            oForm.DataBrowser.BrowseBy = "tCode";
                            //oForm.EnableMenu("1288", true);
                            oForm.EnableMenu("1281", false);
                            oForm.EnableMenu("1289", true);
                            oForm.EnableMenu("1290", true);
                            oForm.EnableMenu("1291", true);
                            oForm.EnableMenu("1287", true);
                            oForm.EnableMenu("6913", false);
                            oForm.EnableMenu("1304", true);
                            AM.MenuEvent(ref pVal, oForm.UniqueID, "Add");
                        }
                    }
                    else if (pVal.MenuUID == "mLc")
                    {
                        if (ExistingForm("frmLc"))
                        {
                            LoadFromXML("frmLc");
                            oForm.DataBrowser.BrowseBy = "tCode";
                            //oForm.EnableMenu("1288", true);
                            oForm.EnableMenu("1281", false);
                            oForm.EnableMenu("1289", true);
                            oForm.EnableMenu("1290", true);
                            oForm.EnableMenu("1291", true);
                            oForm.EnableMenu("1287", true);
                            oForm.EnableMenu("6913", false);
                            oForm.EnableMenu("1304", true);
                            LC.MenuEvent(ref pVal, oForm.UniqueID, "Add");
                        }
                    }
                    else if (pVal.MenuUID == "mCt")
                    {
                        if (ExistingForm("frmCt"))
                        {
                            LoadFromXML("frmCt");
                            oForm.DataBrowser.BrowseBy = "tCode";
                            //oForm.EnableMenu("1288", true);
                            oForm.EnableMenu("1281", false);
                            oForm.EnableMenu("1289", true);
                            oForm.EnableMenu("1290", true);
                            oForm.EnableMenu("1291", true);
                            oForm.EnableMenu("1287", true);
                            oForm.EnableMenu("6913", false);
                            oForm.EnableMenu("1304", true);
                            CT.MenuEvent(ref pVal, oForm.UniqueID, "Add");
                        }
                    }
                    #endregion


                    #endregion


                }
            }
            catch (Exception ex)
            {
                SBO_Application.SetStatusBarMessage("ST Menu Event: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short);
            }
        }
        #endregion

        #region Right Click Event
        private void SBO_Application_RightClickEvent(ref SAPbouiCOM.ContextMenuInfo eventInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                if (eventInfo.BeforeAction)
                {
                    switch (eventInfo.FormUID)
                {
                    case "frmRte":
                        RM.RightClickEvent(ref eventInfo, ref BubbleEvent, eventInfo.FormUID);
                        break;
                }
                }
            }
            catch { }
            
                
            
        }
        #endregion       

        #region Create Menu
        private void CreateAddonMenu()
        {
            try
            {
                oMenus = null;
                oMenuItem = null;
                oCreationPackage = null;

                oMenus = SBO_Application.Menus;
                oMenuItem = SBO_Application.Menus.Item("43520");

                CreateMenu_Item(0, "", SAPbouiCOM.BoMenuType.mt_POPUP, "Add-on", "mRoutePop", oMenus);
                oMenuItem = SBO_Application.Menus.Item("mRoutePop");
                oMenus = oMenuItem.SubMenus;
                CreateMenu_Item(-1, "", SAPbouiCOM.BoMenuType.mt_STRING, "Route Master", "mRoute", oMenus);

                CreateMenu_Item(-1, "", SAPbouiCOM.BoMenuType.mt_STRING, "Allowance Master", "mAlw", oMenus);
                CreateMenu_Item(-1, "", SAPbouiCOM.BoMenuType.mt_STRING, "Load Capacity Master", "mLc", oMenus);
                CreateMenu_Item(-1, "", SAPbouiCOM.BoMenuType.mt_STRING, "Cargo Type Master", "mCt", oMenus);


            }
            catch (Exception ex)
            {
                SBO_Application.SetStatusBarMessage("Create Addon Menu: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short);
            }
        }

        private void CreateMenu_Item(int Position, string ImageName, SAPbouiCOM.BoMenuType Type, string MenuLabel, string MenuId, SAPbouiCOM.Menus oMenus)
        {
            try
            {
                oCreationPackage = ((SAPbouiCOM.MenuCreationParams)(SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)));
                oCreationPackage.Type = Type;
                oCreationPackage.String = MenuLabel;
                oCreationPackage.UniqueID = MenuId;
                oCreationPackage.Enabled = true;
                oCreationPackage.Position = Position;
                if (SBO_Application.ClientType == SAPbouiCOM.BoClientType.ct_Desktop)
                {
                    string sPath = System.Windows.Forms.Application.StartupPath.ToString();
                    oCreationPackage.Image = sPath + "\\" + ImageName;
                }
                if (!oMenus.Exists(MenuId))
                {
                    try
                    {
                        oMenus = oMenuItem.SubMenus;
                        oMenus.AddEx(oCreationPackage);
                    }
                    catch { }
                }
            }
            catch (Exception ex)
            {
                SBO_Application.SetStatusBarMessage("Create Menu Package: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short);
            }
        }

        #endregion

        #region Create DataBase

        private bool CreateDBStructure()
        {
            bool DBCREATE = false;
            try
            {
                oRecordSet = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                oRecordSet.DoQuery("Select USER_CODE From OUSR Where USER_CODE = '" + oCompany.UserName + "'");

                if (oRecordSet.Fields.Item("U_DB").Value.ToString() == "Y")
                {
                    DBCREATE = true;
                }
            }
            catch
            {
                DBCREATE = true;
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
                oRecordSet = null;
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            if (DBCREATE == true)
            {
                try
                {
                    SBO_Application.SetStatusBarMessage("Database structure creation in progress..", SAPbouiCOM.BoMessageTime.bmt_Medium, false);
                    CreateStructure();
                    oRecordSet = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    oRecordSet.DoQuery("Update OUSR Set U_DB = 'N' Where USER_CODE = '" + oCompany.UserName + "'");
                }
                catch (Exception ex)
                {
                    SBO_Application.SetStatusBarMessage("Create DB Structure: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short);
                }
                finally
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
                    oRecordSet = null;
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                }
            }
            return true;
        }

        private bool CreateStructure()
        {
            try
            {
                #region Route Master
                //STRTE Table
                CreateTable("STRTE", "Route Master", SAPbobsCOM.BoUTBTableType.bott_MasterData);


                CreateColumn("@STRTE", "RCODE", "Route Code", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50);
                CreateColumn("@STRTE", "RNAME", "Route Name", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100);
                CreateColumn("@STRTE", "LOADLOC", "Loading Location", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100);
                CreateColumn("@STRTE", "DEST", "Destination", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100);
                CreateColumn("@STRTE", "DIST", "Distance", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None);
                CreateColumn("@STRTE", "ESTHOURS", "Estimated Trip Hours", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None,0,null,"0");
                CreateColumn("@STRTE", "TURNARND", "Turn around Per Month", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, 0, null, "0");


                str = new string[,] { { "Local", "Local" }, { "Transit", "Transit" } };
                CreateColumn("@STRTE", "TROUTE", "Type of Route", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None,25,str);


                //STRTE1 Columns
                CreateTable("STRTE1", "Route Master Row", SAPbobsCOM.BoUTBTableType.bott_MasterDataLines);

                //str = new string[,] { { "Escort Allowance", "Escort Allowance" }, { "Border Charges", "Border Charges" }, { "Road Permit", "Road Permit" } };
                CreateColumn("@STRTE1", "ALLOW", "Allowance", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100); //, str, "Escort Allowance");

                //str = new string[,] { { "Half 7-18MT", "Half 7-18MT"}, { "Full 18MT", "Full 18MT" } };
                CreateColumn("@STRTE1", "LCAP", "Load Capacity", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100); //, str, "Half 7-18MT");
                CreateColumn("@STRTE1", "AMOUNT", "Amount", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Price);


                //str = new string[,] { { "10 FT", "10 FT" }, { "20 FT", "20 FT" } , { "40 FT", "40 FT" } };
                CreateColumn("@STRTE1", "CTYPE", "Cargo Type", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50); //, str, "10 FT");


                //Creating Object for STRTE
                Array.Resize(ref FindColoums, 1);
                FindColoums[0] = "U_RCODE";

                Array.Resize(ref ChildTable, 1);
                ChildTable[0] = "STRTE1";


                CreateObject("STRTE", "Route Master Object", "STRTE", ChildTable,FindColoums,SAPbobsCOM.BoUDOObjType.boud_MasterData,"N");


                //Allowance Table
                CreateTable("STALW", "Allowance Table", SAPbobsCOM.BoUTBTableType.bott_MasterData);
                Array.Resize(ref FindColoums, 1);
                FindColoums[0] = "Code";


                CreateColumn("@STALW", "ACODE", "Allowance Code", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50);
                CreateColumn("@STALW", "ANAME", "Allowance Name", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100);

                CreateObject("STALW", "Allowance Object", "STALW", null, FindColoums, SAPbobsCOM.BoUDOObjType.boud_MasterData, "N");

                //Load Capacity
                CreateTable("STLDC", "Load Capacity Table", SAPbobsCOM.BoUTBTableType.bott_MasterData);


                CreateColumn("@STLDC", "LCCODE", "Load Capacity Code", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50);
                CreateColumn("@STLDC", "LCNAME", "Load Capacity Name", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100);

                CreateObject("STLDC", "Load Capacity Object", "STLDC", null, FindColoums, SAPbobsCOM.BoUDOObjType.boud_MasterData, "N");


                //Cargo Type
                CreateTable("STCAT", "Cargo Type Table", SAPbobsCOM.BoUTBTableType.bott_MasterData);


                CreateColumn("@STCAT", "CTCODE", "Cargo Type Code", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50);
                CreateColumn("@STCAT", "CTNAME", "Cargo Type Name", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100);

                CreateObject("STCAT", "Cargo Type Object", "STCAT", null, FindColoums, SAPbobsCOM.BoUDOObjType.boud_MasterData, "N");

                #endregion

                return true;
            }
            catch(Exception e)
            {
                SetText("CreateStructure : " + e.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                return false;
            }
        }

        private bool CreateTable(string TableName, string TableDesc, SAPbobsCOM.BoUTBTableType TableType)
        {
            SAPbobsCOM.UserTablesMD oUserTablesMD = default(SAPbobsCOM.UserTablesMD);
            try
            {
                oUserTablesMD = (SAPbobsCOM.UserTablesMD)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables);
                if (oUserTablesMD.GetByKey(TableName) == true)
                {
                    return true;
                }
                oUserTablesMD.TableName = TableName;
                oUserTablesMD.TableDescription = TableDesc;
                oUserTablesMD.TableType = TableType;
                RetCode = oUserTablesMD.Add();
                if (RetCode != 0)
                {
                    oCompany.GetLastError(out RetCode, out ErrMsg);
                    SBO_Application.StatusBar.SetText("Table Creation Failed: " + ErrMsg + "", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                }
                else
                {
                    SBO_Application.StatusBar.SetText("Table Created: " + TableDesc + "", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                }
                return true;
            }
            catch
            {
                SBO_Application.StatusBar.SetText("Table Created: " + TableDesc + "", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                return false;
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserTablesMD);
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        private bool CreateColumn(string TableName, string FieldName, string FieldDesc, SAPbobsCOM.BoFieldTypes FieldType, SAPbobsCOM.BoFldSubTypes FieldSubType, int FieldSize = 0, string[,] ValidValues = null, string DefaultVal = null)
        {
            SAPbobsCOM.UserFieldsMD oUserFieldsMD = default(SAPbobsCOM.UserFieldsMD);
            try
            {
                oUserFieldsMD = ((SAPbobsCOM.UserFieldsMD)(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)));
                if (FieldExist(TableName, FieldName) == false)
                {
                    oUserFieldsMD.TableName = TableName;
                    oUserFieldsMD.Name = FieldName;
                    oUserFieldsMD.Description = FieldDesc;
                    oUserFieldsMD.Type = FieldType;
                    oUserFieldsMD.SubType = FieldSubType;
                    oUserFieldsMD.EditSize = FieldSize;

                    if (ValidValues == null)
                    {
                    }
                    else
                    {
                        for (int k = 0; k < ValidValues.Length / 2; k++)
                        {
                            oUserFieldsMD.ValidValues.Value = ValidValues[k, 0];
                            oUserFieldsMD.ValidValues.Description = ValidValues[k, 1];
                            oUserFieldsMD.ValidValues.Add();
                        }
                    }

                    if (!string.IsNullOrEmpty(DefaultVal))
                    {
                        if (DefaultVal == "OUSR")
                        {
                            oUserFieldsMD.LinkedSystemObject = SAPbobsCOM.UDFLinkedSystemObjectTypesEnum.ulUsers;
                        }
                        else
                        {
                            oUserFieldsMD.DefaultValue = DefaultVal;
                        }
                    }
                    RetCode = oUserFieldsMD.Add();
                    if (RetCode != 0)
                    {
                        if (RetCode == -2035 || RetCode == -1120)
                        {
                            return false;
                        }
                        else
                        {
                            oCompany.GetLastError(out RetCode, out ErrMsg);
                            SBO_Application.SetStatusBarMessage("Field Creation Error in : " + TableName + " As : " + FieldDesc + " " + ErrMsg, SAPbouiCOM.BoMessageTime.bmt_Short, false);
                        }
                    }
                    else
                    {
                        SBO_Application.SetStatusBarMessage("Field Created in: " + TableName + " As : " + FieldDesc, SAPbouiCOM.BoMessageTime.bmt_Short, false);
                    }
                }
                return true;
            }
            catch (Exception ex)
            {
                SBO_Application.SetStatusBarMessage("Field Creation Error in: " + TableName + " As : " + FieldDesc + " " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, false);
                return false;
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserFieldsMD);
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        private bool CreateObject(string CodeID, string Name, string TableName, string[] ChildTableName, string[] FindColoums, SAPbobsCOM.BoUDOObjType ObjectType, string ManageSeries) //used for registration of user defined table
        {
            SAPbobsCOM.UserObjectsMD oUserObjectMD = default(SAPbobsCOM.UserObjectsMD);
            try
            {
                oUserObjectMD = ((SAPbobsCOM.UserObjectsMD)(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD)));
                if (oUserObjectMD.GetByKey(CodeID) == true)
                {
                    return true;
                }
                oUserObjectMD.Code = CodeID;
                oUserObjectMD.Name = Name;
                oUserObjectMD.TableName = TableName;

                oUserObjectMD.ObjectType = ObjectType;
                oUserObjectMD.CanClose = SAPbobsCOM.BoYesNoEnum.tNO;
                oUserObjectMD.CanDelete = SAPbobsCOM.BoYesNoEnum.tNO;
                oUserObjectMD.CanCancel = SAPbobsCOM.BoYesNoEnum.tNO;
                oUserObjectMD.CanFind = SAPbobsCOM.BoYesNoEnum.tYES;
                oUserObjectMD.CanLog = SAPbobsCOM.BoYesNoEnum.tYES;

                if (ManageSeries == "Y")
                {
                    oUserObjectMD.ManageSeries = SAPbobsCOM.BoYesNoEnum.tYES;
                }
                else
                {
                    oUserObjectMD.ManageSeries = SAPbobsCOM.BoYesNoEnum.tNO;
                }

                if (ChildTableName != null)
                {
                    for (int i = 0; i <= ChildTableName.Length - 1; i++)
                    {
                        if (ChildTableName[i].Trim() != string.Empty)
                        {
                            oUserObjectMD.ChildTables.TableName = ChildTableName[i];
                            oUserObjectMD.ChildTables.Add();
                        }
                    }
                }
                if (FindColoums != null)
                {
                    for (int i = 0; i <= FindColoums.Length - 1; i++)
                    {
                        if (FindColoums[i].Trim() != string.Empty)
                        {
                            oUserObjectMD.FindColumns.ColumnAlias = FindColoums[i];
                            oUserObjectMD.FindColumns.Add();
                        }
                    }
                }
                // check for errors in the process
                RetCode = oUserObjectMD.Add();

                if (RetCode != 0)
                {
                    if (RetCode != -1)
                    {
                        ST.oCompany.GetLastError(out RetCode, out ErrMsg);
                        SBO_Application.StatusBar.SetText("Object Creation Failed: " + ErrMsg + "", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    }
                }
                else
                {
                    SBO_Application.StatusBar.SetText("Object Registered: " + Name + "", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                }
                return true;
            }
            catch (Exception ex)
            {
                SBO_Application.SetStatusBarMessage("Object Register Error: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, false);
                return false;
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserObjectMD);
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        private bool FieldExist(string TableName, string ColumnName)
        {
            oRecordSet = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            try
            {
                oRecordSet.DoQuery("SELECT COUNT(\"TableID\") FROM CUFD WHERE \"TableID\" = '" + TableName + "' AND \"AliasID\" = '" + ColumnName + "'");
                if ((Convert.ToInt32(oRecordSet.Fields.Item(0).Value) == 0))
                {
                    return false;
                }
                else
                {
                    return true;
                }
            }
            catch (Exception ex)
            {
                SBO_Application.SetStatusBarMessage("Field exist: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, false);
                return false;
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
            }
        }

        #endregion

        #region Other Method

        private void SetFilters()
        {
            oFilters = new SAPbouiCOM.EventFilters();
            oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_LOAD);
            oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_MENU_CLICK);
            oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_CLICK);
            oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_CLOSE);
            oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED);
            oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_KEY_DOWN);
            oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_RIGHT_CLICK);
            oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST);
            oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_COMBO_SELECT);
            oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED);
            SBO_Application.SetFilter(oFilters);
        }

        public static bool ExistingForm(string FormName)
        {
            try
            {
                for (int i = SBO_Application.Forms.Count - 1; i >= 0; i--)
                {
                    if (SBO_Application.Forms.Item(i).UniqueID == FormName)
                    {
                        SBO_Application.Forms.Item(i).Select();
                        return false;
                    }
                }
                return true;
            }
            catch (Exception ex)
            {
                SBO_Application.SetStatusBarMessage("ExistingForm: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short);
                return false;
            }
        }

        private string[] DateSupport()
        {
            string[] DateSupport = new string[] {
                                 "dd/MM/yy",
                                 "dd-MM-yy",
                                 "dd.MM.yy",
                                "dd/MM/yyyy",
                                 "dd-MM-yyyy",
                                 "dd.MM.yyyy",
                                 "MM/dd/yy",
                                 "MM-dd-yy",
                                 "MM.dd.yy",
                                 "MM/dd/yyyy",
                                 "MM-dd-yyyy",
                                 "MM.dd.yyyy",
                                 "ddMMMyyyy",
                                 "ddMMyyyy",
                                 "yyyyMMdd",
                                 "yyyy-MM-dd"
                                                };
            return DateSupport;
        }

        public static void LoadFromXML(string FileName)
        {
            try
            {
                System.Xml.XmlDocument oXmlDoc = null;
                oXmlDoc = new System.Xml.XmlDocument();

                // load the content of the XML File
                string sPath = null;
                FileName = FileName + ".xml";
                sPath = System.Windows.Forms.Application.StartupPath.ToString();
                oXmlDoc.Load(sPath + @"\" + FileName);

                // load the form to the SBO application in one batch
                string tmpStr;
                tmpStr = oXmlDoc.InnerXml;
                SBO_Application.LoadBatchActions(ref tmpStr);
                sPath = SBO_Application.GetLastBatchResults();
                oForm = SBO_Application.Forms.ActiveForm;
            }
            catch (Exception ex)
            {
                SBO_Application.SetStatusBarMessage("Load From XML: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short);
            }
        }

        public static void SetText(string Text, SAPbouiCOM.BoStatusBarMessageType Flag)
        {
            SBO_Application.StatusBar.SetText(Text, SAPbouiCOM.BoMessageTime.bmt_Short, Flag);
        }


        #endregion
    }
}
