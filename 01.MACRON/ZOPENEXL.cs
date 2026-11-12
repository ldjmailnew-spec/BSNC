//************************************************************************************************************
//  File            : ZOPENEXL.cls
//  Module          : 공통
//  Desc            : 엑셀업로딩
//  FormType        : ZOPENEXL
//  Create Date     : 2011.08.29
//  Modified Date   : 2011.11.17
//  Creator         : Seo Young Mi
//  Memo            : 2015.02.05.SYM. 추가(ZVAT100A)
//                  : 2015.06.03.SYM. 추가(ZFD701,ZFD704,ZFD713)
//                  : 2018.10.22.SYM.윈도우탐색기창뒤로숨는부분수정OpenFileExcell수정(이정남대리수정내역참조)
//  Copyright (c) BSNC
//************************************************************************************************************
using System;
using System.Windows.Forms;


class ZOPENEXL
{
    //************************************************************************************************************
    private static SAPbouiCOM.Form oForm;
    private static string oFather;


    //************************************************************************************************************
    // ItemEventHander
    //************************************************************************************************************
    public void SBO_Application_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
    {
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
                        if (System.IO.File.Exists(oForm.DataSources.UserDataSources.Item("Path").Value) == false)
                        {
                            ZSBO.SBO_Application.StatusBar.SetText("파일이 존재하지 않습니다. 경로 또는 파일명을 다시 확인하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                            BubbleEvent = false;
                        }
                        else
                        {
                            ReturnData();
                            oForm.Close();
                        }
                    }
                    break;
            }
        }
        //-------------------------------------------------------------------------------------------------------
        catch (Exception Error)
        {
            ZSBO.SBO_Application.SetStatusBarMessage(Error.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        }
    }


    //************************************************************************************************************
    // MenuEventHander
    //************************************************************************************************************
    public void SBO_Application_MenuEvent(ref SAPbouiCOM.MenuEvent pVal, out bool BubbleEvent)
    {
        BubbleEvent = true;

        try
        {
            if (pVal.BeforeAction == true && pVal.MenuUID == "ZOPENEXL")    {   LoadForm("ZOPENEXL", "");   }
        }
        //-------------------------------------------------------------------------------------------------------
        catch (Exception Error)
        {
            ZSBO.SBO_Application.SetStatusBarMessage("SBO_Application_MenuEvent Error : " + Error.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        }
    }


    //************************************************************************************************************
    // 
    //************************************************************************************************************
    private void OpenFileExcell()
    {
        OpenFileDialog pFileDlg = new OpenFileDialog();
        String strFullPathFile = "";
        pFileDlg.Filter = "Excel Files|*.xl*|All Files(*.*)|*.*";
        pFileDlg.Multiselect = false;
        pFileDlg.RestoreDirectory = true;
        pFileDlg.Title = "파일선택";

        try
        {
            if (pFileDlg.ShowDialog(WindowWrapper.GetForegroundWindowWrapper()) == DialogResult.OK)
            {
                strFullPathFile = pFileDlg.FileName;
                oForm.DataSources.UserDataSources.Item("Path").Value = strFullPathFile;
            }
        }
        catch (System.Exception Error)
        {
            ZSBO.SBO_Application.SetStatusBarMessage("OpenFileExcell Error : " + Error.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        }

        //try
        //{
        //    OpenFileDialog pFileDlg = new OpenFileDialog();
        //    String strFullPathFile = "";
        //    pFileDlg.Filter = "Excel Files|*.xl*|All Files(*.*)|*.*";
        //    pFileDlg.Multiselect = false;
        //    pFileDlg.RestoreDirectory = true;
        //    pFileDlg.Title = "파일선택";

        //    // 경로 가져오기
        //    if (pFileDlg.ShowDialog() == DialogResult.OK) { strFullPathFile = pFileDlg.FileName; }
        //    try
        //    {
        //        oForm.DataSources.UserDataSources.Item("Path").Value = strFullPathFile;
        //    }
        //    catch
        //    {
        //    }
        //}
        //catch
        //{
        //}
    }

    
    //************************************************************************************************************
    // 
    //************************************************************************************************************
    private void ReturnData()
    {
        string oPath = oForm.DataSources.UserDataSources.Item("Path").Value;
        string oSheetName = oForm.DataSources.UserDataSources.Item("SheetNm").Value;
        int oHLine = Convert.ToInt16(oForm.DataSources.UserDataSources.Item("HLine").Value);
/*
        switch (oFather)
        {
            //case "SAMPLE":
            //    SAMPLE oSAMPLE = new SAMPLE();                          // SAMPLE
            //    oSAMPLE.SetFileTarget(oPath, oSheetName, oHLine);
            //    break;





        }
 */ 
    }


    //************************************************************************************************************
    // 
    //************************************************************************************************************
    public void LoadForm(string FileName, string oString)
    {
        try
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
            oFather = oString;
            CreateItems();
            oForm.SupportedModes = 1;
            oForm.Visible = true;
        }
        //-------------------------------------------------------------------------------------------------------
        catch (Exception Error)
        {
            ZSBO.SBO_Application.SetStatusBarMessage("LoadForm_Error:" + Error.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        }
    }


    //************************************************************************************************************
    // 
    //************************************************************************************************************
    private void CreateItems()
    {
        oForm.DataSources.UserDataSources.Add("Path", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 255);
        oForm.DataSources.UserDataSources.Add("SheetNm", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
        oForm.DataSources.UserDataSources.Add("HLine", SAPbouiCOM.BoDataType.dt_SHORT_NUMBER, 10);
        ((SAPbouiCOM.EditText)oForm.Items.Item("Path").Specific).DataBind.SetBound(true, "", "Path");
        ((SAPbouiCOM.EditText)oForm.Items.Item("SheetNm").Specific).DataBind.SetBound(true, "", "SheetNm");
        ((SAPbouiCOM.EditText)oForm.Items.Item("HLine").Specific).DataBind.SetBound(true, "", "HLine");
        oForm.DataSources.UserDataSources.Item("SheetNm").Value = "Sheet1";
        oForm.DataSources.UserDataSources.Item("HLine").Value = "1";
    }
}
