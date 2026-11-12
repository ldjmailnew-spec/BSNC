//****************************************************************************
//  File       :   ZCHLIST.cls
//  Module     :   공통
//  Desc       :   선택리스트
//  FormType   :   ZCHLIST
//  Create Date   : 2013.05.03
//  Creator       : Kim Baek Bum
//  MEMO          : 2022.12.15.BJH ZQM707B_ZPI518 관련 SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK 수정, Return_Data() 수정, LoadDataTable 수정
//****************************************************************************

using System;
using System.Windows.Forms;

class ZCHLIST
{
    private static string oFather;
    private static SAPbouiCOM.Form oForm;
    private static SAPbouiCOM.Grid oGrid;

    //*******************************************************************************************
    public void SBO_Application_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
    {
        SAPbouiCOM.IChooseFromListEvent oCFL = null;
        string oAcctCode;

        BubbleEvent = true;

        if (oForm.UniqueID != FormUID)
        {
            oForm = ZSBO.SBO_Application.Forms.Item(FormUID);
            //SetoForm();
        }

        try
        {
            switch (pVal.EventType)
            {
                //==============================================================================
                case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
                    if (pVal.BeforeAction & pVal.ItemUID == "BtnSel")
                    {
                        if (oGrid.Rows.SelectedRows.Count == 0)
                        {
                            BubbleEvent = false;
                            return;
                        }
                        Return_Data();
                    }
                    else if (pVal.BeforeAction == false & pVal.ItemUID == "Grd1" & pVal.ColUID != "RowsHeader" & pVal.Row != -1)
                    {
                        if (oGrid.Rows.IsSelected(pVal.Row) == true)
                            oGrid.Rows.SelectedRows.Remove(pVal.Row);
                        else if (oGrid.Rows.IsSelected(pVal.Row) == false)
                            oGrid.Rows.SelectedRows.Add(pVal.Row);
                    }
                    break;

                //==============================================================================
                case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK:
                    if (pVal.BeforeAction)
                    {
                        if (oGrid.Rows.SelectedRows.Count == 0)
                        {
                            if (oFather == "ZQM707B_ZPI518")
                            {
                                //공정최종검사결과 더블클릭 정렬 2022.12.15.BJH                           
                            }
                            else
                            {
                                BubbleEvent = false;
                            }
                            return;
                        }
                        Return_Data();
                    }
                    else if (pVal.BeforeAction == false & pVal.ItemUID == "Grd1" & pVal.ColUID != "RowsHeader" & pVal.Row != -1)
                    {
                        if (oGrid.Rows.IsSelected(pVal.Row) == true)
                            oGrid.Rows.SelectedRows.Remove(pVal.Row);
                        else if (oGrid.Rows.IsSelected(pVal.Row) == false)
                            oGrid.Rows.SelectedRows.Add(pVal.Row);
                    }
                    break;

                //==============================================================================
                case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED:

                    if (pVal.BeforeAction & pVal.ItemUID == "Grd1" & pVal.ColUID == "의뢰번호" & pVal.Row >= 0)
                    {
                        //ZCMP303 oZCMP303 = new ZCMP303(); oZCMP303.LoadForm("ZCMP303", oGrid.DataTable.Columns.Item(pVal.ColUID).Cells.Item(oGrid.GetDataTableRowIndex(pVal.Row)).Value.ToString());

                        BubbleEvent = false;
                    }
                    break;

            }
        }
        catch (Exception Error)
        {
            ZSBO.SBO_Application.SetStatusBarMessage(Error.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        }
    }

    //*******************************************************************************************
    private void Return_Data()
    {

/*
        switch (oFather)
        {
            //case "SAMPLE": 
            //SAMPLE oSAMPLE = new SAMPLE();                                // SAMPLE
            //oSAMPLE.SetChooseList(oForm.DataSources.DataTables.Item("DT1"), oGrid.Rows.SelectedRows); 
            //break;






        }
*/
        oForm.Close();
    }

    //*******************************************************************************************
    public void LoadChooseList(string oFUid, string oQuery, SAPbouiCOM.BoMatrixSelect oSelMode)
    {
        SAPbouiCOM.FormCreationParams oCreationParams = null;
        System.Xml.XmlDocument oXmlDoc = null;
        string sPath = null;

        oXmlDoc = new System.Xml.XmlDocument();

        // load the content of the XML File
        //sPath = System.IO.Directory.GetParent(Application.StartupPath).ToString();
        //sPath = System.IO.Directory.GetParent(sPath).ToString();
        sPath = Application.StartupPath;

        oXmlDoc.Load(sPath + @"\Screen\ZCHLIST.srf");
        oXmlDoc.SelectSingleNode("Application/forms/action/form/@uid").Value = "";
        oXmlDoc.SelectSingleNode("Application/forms/action/form/@top").Value
            = Convert.ToString(Convert.ToInt16(oXmlDoc.SelectSingleNode("Application/forms/action/form/@top").Value) + (ZSBO.SBO_Application.Forms.Count + 5));
        oXmlDoc.SelectSingleNode("Application/forms/action/form/@left").Value
            = Convert.ToString(Convert.ToInt16(oXmlDoc.SelectSingleNode("Application/forms/action/form/@left").Value) + (ZSBO.SBO_Application.Forms.Count + 5));

        // add a new form
        oCreationParams = ((SAPbouiCOM.FormCreationParams)(ZSBO.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams)));
        oCreationParams.XmlData = oXmlDoc.InnerXml.ToString();

        oForm = ZSBO.SBO_Application.Forms.AddEx(oCreationParams);
        oGrid = (SAPbouiCOM.Grid)oForm.Items.Item("Grd1").Specific;

        oFather = oFUid;
        //==========================================================================================
        if (LoadDataTable(oQuery, oSelMode, oFather) == false)
        {
            ZSBO.SBO_Application.SetStatusBarMessage("선택리스트를 실행하는 동안에 에러가 발생했습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
            oForm.Close();
        }
        //==========================================================================================

        oForm.Visible = true;

        oGrid.AutoResizeColumns();


    }

    //*******************************************************************************************
    private Boolean LoadDataTable(string oQuery, SAPbouiCOM.BoMatrixSelect oSelMode, string oFather)
    {
        SAPbouiCOM.DataTable oDTable;
        SAPbouiCOM.BoFieldsType oType;
        int i;

        try
        {
            oForm.Freeze(true);

            oDTable = oForm.DataSources.DataTables.Item("DT1");
            oDTable.Clear();

            oDTable.ExecuteQuery(oQuery);
            oGrid.DataTable = oForm.DataSources.DataTables.Item("DT1");
            oGrid.SelectionMode = oSelMode;

            for (i = 0; i < oGrid.Columns.Count; i++)
            {
                oType = oGrid.DataTable.Columns.Item(i).Type;
                if (oType == SAPbouiCOM.BoFieldsType.ft_Float | oType == SAPbouiCOM.BoFieldsType.ft_Quantity
                    | oType == SAPbouiCOM.BoFieldsType.ft_Price | oType == SAPbouiCOM.BoFieldsType.ft_Sum)
                {
                    oGrid.Columns.Item(i).RightJustified = true;
                }

            }
            if (oFather == "ZQM707B_ZPI518") //공정최종검사결과 정렬 2022.12.15.BJH
            {
                // 컬럼 Sort          
                for (i = 0; i <= oGrid.Columns.Count - 1; i++)
                {
                    oGrid.Columns.Item(i).TitleObject.Sortable = true;
                }
            }
            if (oFather == "ZQM705B")
            {
                ((SAPbouiCOM.EditTextColumn)oGrid.Columns.Item("의뢰번호")).LinkedObjectType = "ZCMP303";
            }

            
            oForm.Freeze(false);
            oForm.Update();
            oGrid.AutoResizeColumns();
            return true;
        }
        catch (Exception Error)
        {
            oForm.Freeze(false);
            return false;
        }
    }


}