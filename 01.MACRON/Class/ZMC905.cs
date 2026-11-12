//************************************************************************************************************
//  File             : ZMC905.cs
//  Module           : 마스터 스케쥴 
//  Desc             : 마스터 스케쥴
//  FormType         : ZMC905
//  Create Date      : 2026.05.14
//  Creator          : KIM SUN MI
//  memo             : 
//  Copyright (c) BSNC
//************************************************************************************************************

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

class ZMC905
{
    private static SAPbouiCOM.Form oForm;
    private static SAPbouiCOM.Grid oGRD1;
    private SAPbouiCOM.IEditText oEditText;

    //************************************************************************************************************
    // ItemEventHander    
    //************************************************************************************************************
    public void SBO_Application_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
    {
        SAPbouiCOM.IChooseFromListEvent oCFL = null;
        string oStr, oCUID;
        BubbleEvent = true;

        if (oForm.UniqueID != FormUID)
        {
            oForm = ZSBO.SBO_Application.Forms.Item(FormUID);
            oGRD1 = (SAPbouiCOM.Grid)oForm.Items.Item("GRD1").Specific;
        }

        try
        {
            switch (pVal.EventType)
            {
                //=====================================================================================
                case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
                //=====================================================================================
                    if (pVal.Before_Action & pVal.ItemUID == "BTN_SRH")
                    {
                        if (ChkInData() == true)
                        {
                            if (ZSBO.SBO_Application.MessageBox("조회하시겠습니까?", 1, "예", "아니오", "") == 1)
                            {
                                LoadData("S");
                            }
                            else
                            {
                                BubbleEvent = false;
                            }
                        }
                        else
                        {
                            BubbleEvent = false;
                        }

                    }
                    //if (pVal.Before_Action & pVal.ItemUID == "1")
                    //{
                    //}
                    break;

                //=====================================================================================
                case SAPbouiCOM.BoEventTypes.et_CLICK:
                //=====================================================================================
                    if (pVal.BeforeAction & pVal.ItemUID == "GRD1" & pVal.ColUID != "RowsHeader" & pVal.Row >= 0)
                    {
                        if (oGRD1.Rows.IsSelected(pVal.Row) == false)
                        {
                            oGRD1.Rows.SelectedRows.Clear();
                            oGRD1.Rows.SelectedRows.Add(pVal.Row);      
                        }
                    }
                    break;

                //=====================================================================================
                case SAPbouiCOM.BoEventTypes.et_VALIDATE:
                //=====================================================================================
                    if (pVal.Before_Action == false & pVal.ItemChanged & oForm.Items.Item(pVal.ItemUID).Type == SAPbouiCOM.BoFormItemTypes.it_EDIT)
                    {
                        oStr = ((SAPbouiCOM.EditText)oForm.Items.Item(pVal.ItemUID).Specific).String.Trim();
                        FlushToItemValue(pVal.ItemUID, pVal.ColUID, pVal.Row, oStr);
                    }
                    break;

                //=====================================================================================
                case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE:
                //=====================================================================================
                    if (pVal.Before_Action == false & oGRD1.Rows.Count > 0)
                    {
                        oGRD1.AutoResizeColumns();
                    }
                    break;
            }

        }
        catch (Exception Error)
        {
            ZSBO.SBO_Application.SetStatusBarMessage("SBO_Application_ItemEvent Error :" + Error.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
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
            if (pVal.BeforeAction & pVal.MenuUID == "ZMC905") { LoadForm("ZMC905", ""); }

        }
        catch (Exception Error)
        {
            ZSBO.SBO_Application.SetStatusBarMessage("SBO_Application_MenuEvent Error : " + Error.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        }
        finally
        {

        }

    }
    
    //************************************************************************************************************
    // FlushToItemValue        
    //************************************************************************************************************
    private void FlushToItemValue(string sUID, string sCol, int sRow, string sValue)
    {
        try
        {


        }
        catch (Exception Error)
        {
            ZSBO.SBO_Application.StatusBar.SetText("FlushToItemValue_Error : " + Error.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
        }
        finally
        {
            //if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
        }

    }


    //************************************************************************************************************
    // ChkInData     
    //************************************************************************************************************
    private bool ChkInData()
    {
        try
        {
            // 프로젝트번호 체크
            if (oForm.DataSources.UserDataSources.Item("PRJCODE").Value == "")
            {
                ZSBO.SBO_Application.SetStatusBarMessage("프로젝트번호가 누락되었습니다", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                return false;
            }
           

            return true;

        }
        catch (Exception Error)
        {
            ZSBO.SBO_Application.StatusBar.SetText("ChkInData_Error : " + Error.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            return false;
        }
        finally
        {
            oForm.Freeze(false);
        }

    }

    //************************************************************************************************************
    // LoadData        
    //************************************************************************************************************
    public bool LoadData(string sType)
    {
        SAPbouiCOM.EditTextColumn oCol;
        string sQry = "";
        string oPRJCODE;
        string[] oArray;

        try
        {
            oForm.Freeze(true);

            oPRJCODE = oForm.DataSources.UserDataSources.Item("PRJCODE").Value.Trim();

            sQry = "EXEC dbo.[ZTMSP_ZMC905] '" + oPRJCODE + "'";
            oForm.DataSources.DataTables.Item("GRD1").ExecuteQuery(sQry);
            oGRD1.DataTable = oForm.DataSources.DataTables.Item("GRD1");

            // 숫자 오른쪽 정렬 처리
            oGRD1.Columns.Item("수주금액").RightJustified = true;
            oGRD1.Columns.Item("매출금액").RightJustified = true;
            oGRD1.Columns.Item("수주잔액").RightJustified = true;
            oGRD1.Columns.Item("수주금액 (원화)").RightJustified = true;
            oGRD1.Columns.Item("매출금액 (원화)").RightJustified = true;
            oGRD1.Columns.Item("수주잔액 (원화)").RightJustified = true;


            //oForm.Freeze(false);

            ZSBO.SBO_Application.StatusBar.SetText("LoadData_Success", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
            return true;

        }
        catch (Exception Error)
        {
            //oForm.Freeze(false);
            ZSBO.SBO_Application.StatusBar.SetText("LoadData_Error : " + Error.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            return false;
        }
        finally
        {
            oForm.Freeze(false);
        }
    }

    //************************************************************************************************************
    // LoadForm         
    //************************************************************************************************************
    public void LoadForm(string FileName, string oString)
    {

        try
        {
            SAPbouiCOM.FormCreationParams oCreationParams = null;
            System.Xml.XmlDocument oXmlDoc = null;
            string sPath = null;

            oXmlDoc = new System.Xml.XmlDocument();
            sPath = Application.StartupPath;

            oXmlDoc.Load(sPath + @"\Screen\" + FileName + ".srf");
            //일치하는 첫번째노드 선택
            //oXmlDoc.SelectSingleNode("Application/forms/action/form/@uid").Value = "";
            //oXmlDoc.SelectSingleNode("Application/forms/action/form/@top").Value = Convert.ToString(Convert.ToInt32(oXmlDoc.SelectSingleNode("Application/forms/action/form/@top").Value) + (ZSBO.SBO_Application.Forms.Count + 5));
            //oXmlDoc.SelectSingleNode("Application/forms/action/form/@left").Value = Convert.ToString(Convert.ToInt32(oXmlDoc.SelectSingleNode("Application/forms/action/form/@left").Value) + (ZSBO.SBO_Application.Forms.Count + 5));

            oXmlDoc.SelectSingleNode("Application/forms/action/form/@uid").Value = "";
            oXmlDoc.SelectSingleNode("Application/forms/action/form/@top").Value
                = Convert.ToString(Convert.ToInt16(oXmlDoc.SelectSingleNode("Application/forms/action/form/@top").Value) + (ZSBO.SBO_Application.Forms.Count + 5));
            oXmlDoc.SelectSingleNode("Application/forms/action/form/@left").Value
                = Convert.ToString(Convert.ToInt16(oXmlDoc.SelectSingleNode("Application/forms/action/form/@left").Value) + (ZSBO.SBO_Application.Forms.Count + 5));


            // add a new form
            oCreationParams = ((SAPbouiCOM.FormCreationParams)(ZSBO.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams)));
            oCreationParams.XmlData = oXmlDoc.InnerXml.ToString();

            oForm = ZSBO.SBO_Application.Forms.AddEx(oCreationParams);
            oForm.Title = ZSBO.ZAPI.FGetFormTitle(oForm.Title, FileName, oForm.TypeEx, oForm.BusinessObject.Type);
            oGRD1 = (SAPbouiCOM.Grid)oForm.Items.Item("GRD1").Specific;
            oForm.DataSources.DataTables.Add("GRD1");

            CreateItems();
            SetFormBase();
            LoadData("B");

            oForm.Visible = true;           // 기본으로 false


        }
        catch (Exception Error)
        {
            ZSBO.SBO_Application.SetStatusBarMessage("LoadForm_Error:" + Error.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        }

    }

    //************************************************************************************************************
    // CreateItems : 콤보리스트/체크박스/매트릭스/폴더/(총계:합계 구하는 컬럼도 여기서한다.)       
    //************************************************************************************************************
    private void CreateItems()
    {

        try
        {
            // Userdatasauce 생성
            oForm.DataSources.UserDataSources.Add("PRJCODE", SAPbouiCOM.BoDataType.dt_LONG_TEXT, 100);

            // Binding (스크핀페인터랑 소스랑 연결하기)
            ((SAPbouiCOM.EditText)oForm.Items.Item("PRJCODE").Specific).DataBind.SetBound(true, "", "PRJCODE");

            // 프로젝트번호 "" 호출
            //oForm.DataSources.UserDataSources.Item("PRJCODE").ValueEx = "";

        }
        catch (Exception Error)
        {
            ZSBO.SBO_Application.SetStatusBarMessage("CreateItems_Error : " + Error.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        }

    }

    //************************************************************************************************************
    // SetFormBase : 기본값셋팅     
    //************************************************************************************************************
    private void SetFormBase()
    {

    }

    //************************************************************************************************************
    // SetToolBarEnabled        
    //************************************************************************************************************
    private void SetToolBarEnabled()
    {
        oForm.EnableMenu("1281", false); // 찾기
        oForm.EnableMenu("1282", false); // 추가
        oForm.EnableMenu("1283", false); // 제거
        oForm.EnableMenu("1284", false); // 취소
        oForm.EnableMenu("1285", false); // 복원
        oForm.EnableMenu("1286", false); // 닫기
        oForm.EnableMenu("1287", false); // 복제

    }

















}
