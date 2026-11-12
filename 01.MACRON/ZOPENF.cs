//****************************************************************************
//  File        :   ZOPENF.cls
//  Module      :   공통
//  Desc        :   파일경로참조
//  FormType    :   ZOPENF
//  Create Date :   2011.07.18
//  Creator     :   Kim Baek Bum
//  Memo        :   2018.10.22.SYM.윈도우탐색기창뒤로숨는부분수정OpenFileExcell수정(이정남대리수정내역참조)
//****************************************************************************

using System;
using System.Windows.Forms;
class ZOPENF
{
    private static SAPbouiCOM.Form oForm = null;
    private SAPbouiCOM.IEditText oEditText;
    private static string oReturnClass;

    //*******************************************************************************************
    public void SBO_Application_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
    {
        SAPbouiCOM.IChooseFromListEvent oCFL = null;

        BubbleEvent = true;

        if (oForm.UniqueID != FormUID)
        {
            oForm = ZSBO.SBO_Application.Forms.Item(FormUID);
        }

        try
        {
            switch (pVal.EventType)
            {
                //==============================================================================
                case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
                    if (pVal.BeforeAction & pVal.ItemUID == "BtnSel")
                    {
                        System.Threading.Thread th = new System.Threading.Thread(OpenFileExcell);
                        th.SetApartmentState(System.Threading.ApartmentState.STA);
                        th.Start();
                        //th.Join();
                    }
                    else if (pVal.BeforeAction & pVal.ItemUID == "BtnOK")
                    {
                        ReturnData();
                        oForm.Close();
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
    public void SBO_Application_MenuEvent(ref SAPbouiCOM.MenuEvent pVal, out bool BubbleEvent)
    {

        BubbleEvent = true;

        try
        {
            if (pVal.BeforeAction == true && pVal.MenuUID == "ZOPENF")
            {
                LoadForm("ZOPENF", "");
            }
        }
        catch (Exception Error)
        {
            ZSBO.SBO_Application.SetStatusBarMessage("SBO_Application_MenuEvent Error : " + Error.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        }
    }

    //*******************************************************************************************
    private void OpenFileExcell()
    {
        OpenFileDialog pFileDlg = new OpenFileDialog();
        String strFullPathFile = "";

        //pFileDlg.Filter = "Text Files(*.txt)|*.txt|All Files(*.*)|*.*";
        pFileDlg.Filter = "Excel Files(*.xlsx)|*.xlsx|All Files(*.*)|*.*";
        //pFileDlg.Title = "파일을 선택하여 주세요.";

        ZSBO.SBO_Application.SendKeys("%{TAB}");

        // 경로 가져오기
        if (pFileDlg.ShowDialog(WindowWrapper.GetForegroundWindowWrapper()) == DialogResult.OK)
        {
            strFullPathFile = pFileDlg.FileName;
            pFileDlg.Dispose();
        }
        else
        {
            return;
        }

        oForm.DataSources.UserDataSources.Item("Path").Value = strFullPathFile;

        //////////OpenFileDialog pFileDlg = new OpenFileDialog();
        //////////String strFullPathFile = "";

        ////////////pFileDlg.Filter = "Text Files(*.txt)|*.txt|All Files(*.*)|*.*";
        //////////pFileDlg.Filter = "Excel Files(*.xlsx)|*.xlsx|All Files(*.*)|*.*";
        ////////////pFileDlg.Title = "파일을 선택하여 주세요.";
        
        //////////// 경로 가져오기
        //////////if (pFileDlg.ShowDialog() == DialogResult.OK)
        //////////    strFullPathFile = pFileDlg.FileName;

        //////////oForm.DataSources.UserDataSources.Item("Path").Value = strFullPathFile;
    }

    //*******************************************************************************************
    private void ReturnData()
    {
        string oPath, oSheetName;
        int oHLine;
        oPath = oForm.DataSources.UserDataSources.Item("Path").Value;
        oSheetName = oForm.DataSources.UserDataSources.Item("SheetNm").Value;
        oHLine = Convert.ToInt16(oForm.DataSources.UserDataSources.Item("HLine").Value);
/*
        switch (oReturnClass)
        {
            //case "SAMPLE":
            //    SAMPLE oSAMPLE = new SAMPLE();                              // SAMPLE
            //    oSAMPLE.SetFileTarget(oPath, oSheetName, oHLine); 
            //    break;

            //////////case "ZCMP703":
            //////////    ZCMP703 oZCMP703 = new ZCMP703();                              // SAMPLE
            //////////    oZCMP703.SetFileTarget(oPath, oSheetName, oHLine);
            //////////    break;

            //////////case "ZCMP701":
            //////////    ZCMP701 oZCMP701 = new ZCMP701();                              // SAMPLE
            //////////    oZCMP701.SetFileTarget(oPath, oSheetName, oHLine);
            //////////    break;

        }
 */ 
    }

    //*******************************************************************************************
    public void LoadForm(string FileName, string oString)
    {
        SAPbouiCOM.FormCreationParams oCreationParams = null;
        System.Xml.XmlDocument oXmlDoc = null;

        oXmlDoc = new System.Xml.XmlDocument();

        // load the content of the XML File
        oXmlDoc.Load(Application.StartupPath + @"\Screen\" + FileName + ".srf");
        oXmlDoc.SelectSingleNode("Application/forms/action/form/@uid").Value = "";
        oXmlDoc.SelectSingleNode("Application/forms/action/form/@top").Value
            = Convert.ToString(Convert.ToInt16(oXmlDoc.SelectSingleNode("Application/forms/action/form/@top").Value) + (ZSBO.SBO_Application.Forms.Count + 5));
        oXmlDoc.SelectSingleNode("Application/forms/action/form/@left").Value
            = Convert.ToString(Convert.ToInt16(oXmlDoc.SelectSingleNode("Application/forms/action/form/@left").Value) + (ZSBO.SBO_Application.Forms.Count + 5));

        // add a new form
        oCreationParams = ((SAPbouiCOM.FormCreationParams)(ZSBO.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams)));
        oCreationParams.XmlData = oXmlDoc.InnerXml.ToString();

        oForm = ZSBO.SBO_Application.Forms.AddEx(oCreationParams);
        //oForm.Title = ZSBO.ZAPI.FGetFormTitle(oForm.Title, FileName, oForm.TypeEx, oForm.BusinessObject.Type);
        //oForm.DataBrowser.BrowseBy = "DocEntry";
        CreateItems();
        oReturnClass = oString;
        oForm.Visible = true;

    }

    //*******************************************************************************************
    private void CreateItems()
    {

        oForm.DataSources.UserDataSources.Add("Path", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 255);
        oForm.DataSources.UserDataSources.Add("SheetNm", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
        oForm.DataSources.UserDataSources.Add("HLine", SAPbouiCOM.BoDataType.dt_SHORT_NUMBER, 10);

        //바인딩=====================================================================
        oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("Path").Specific;
        oEditText.DataBind.SetBound(true, "", "Path");
        oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("SheetNm").Specific;
        oEditText.DataBind.SetBound(true, "", "SheetNm");
        oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("HLine").Specific;
        oEditText.DataBind.SetBound(true, "", "HLine");

        oForm.DataSources.UserDataSources.Item("SheetNm").Value = "Sheet1";
        oForm.DataSources.UserDataSources.Item("HLine").Value = "1";

        oForm.SupportedModes = 1;
        
    }

}

