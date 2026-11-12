//************************************************************************************************************
//  File             : ZMC935.cs
//  Module           : 생산관리_구성품 재고확보 현황조회
//  Desc             : 생산관리_구성품 재고확보 현황조회
//  FormType         : ZMC935
//  Create Date      : 2026.09.18
//  Creator          : LEE DAE JIN
//  memo             : 
//  Copyright (c) BSNC
//************************************************************************************************************

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

class ZMC935
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
            if (pVal.BeforeAction & pVal.MenuUID == "ZMC935") { LoadForm("ZMC935", ""); }

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
        // 사용안함, 배치 번호를 입력하지 않고 전체를 조회할수도 있음

        try
        {
            // 배치번호 지정 체크
            if (oForm.DataSources.UserDataSources.Item("PRDSTATUS").Value == "")
            {
                ZSBO.SBO_Application.SetStatusBarMessage("생산오더 상태 지정이 누락되었습니다", SAPbouiCOM.BoMessageTime.bmt_Short, true);
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
        string oPRDSTATUS = "";

        string[] oArray;

        try
        {
            oForm.Freeze(true);


            SAPbouiCOM.ComboBox oComboStatus = (SAPbouiCOM.ComboBox)oForm.Items.Item("PRDSTATUS").Specific;
            if (oComboStatus.Selected != null)
            {
                //oWHSCODE = oForm.DataSources.UserDataSources.Item("WHSCODE").Value.Trim();
                oPRDSTATUS = oComboStatus.Selected.Description.Trim();
                //oWHSCODE = oComboStatus.Selected.Value.ToString();
            }

            sQry = "EXEC dbo.[ZTMSP_ZMC935] '" + sType + "','" + oPRDSTATUS + "'";
            oForm.DataSources.DataTables.Item("GRD1").ExecuteQuery(sQry);
            oGRD1.DataTable = oForm.DataSources.DataTables.Item("GRD1");

            //링크
            ((SAPbouiCOM.EditTextColumn)oGRD1.Columns.Item("구성품코드")).LinkedObjectType = "4";    // 품목마스터 링크설정

            //숫자 정렬
            oGRD1.Columns.Item("총필요수량").RightJustified = true;
            oGRD1.Columns.Item("WH10현재고").RightJustified = true;
            oGRD1.Columns.Item("WH20현재고").RightJustified = true;
            oGRD1.Columns.Item("총가용차이수량").RightJustified = true;

            /* 링크 참고사항 
  
                        LinkedObjectType	대상 화면 (오브젝트)	관련 테이블
                        "2"	    비즈니스 파트너 마스터 (BP Master Data)	OCRD
                        "4"	    품목 마스터 (Item Master Data)	OITM
                        "63"	프로젝트 (Project Setup)	OPRJ
                        "64"	창고 (Warehouse Setup)	OWHS
                        "157"	코스트 센터 (Cost Center)	OOCR
                        "171"	사원 마스터 (Employee Master Data)	OHEM
                        2. 영업 및 A/R (Sales - A/R)		
		
                        LinkedObjectType	대상 화면 (오브젝트)	관련 테이블
                        "13"	A/R 송장 (A/R Invoice)	OINV
                        "14"	A/R 대변 메모 (A/R Credit Memo)	ORIN
                        "15"	납품 (Delivery)	ODLN
                        "16"	반품 (Returns)	ORDN
                        "17"	판매 오더 (Sales Order)	ORDR
                        "23"	판매 견적 (Sales Quotation)	OQUT
                        3. 구매 및 A/P (Purchasing - A/P)		
		
                        LinkedObjectType	대상 화면 (오브젝트)	관련 테이블
                        "18"	A/P 송장 (A/P Invoice)	OPCH
                        "19"	A/P 대변 메모 (A/P Credit Memo)	ORPC
                        "20"	A/P 입고 (Goods Receipt PO)	OPDN
                        "21"	A/P 반품 (Goods Return)	ORPD
                        "22"	구매 오더 (Purchase Order)	OPOR
                        "54"	구매 견적 (Purchase Quotation)	OPQT
                        4. 재고 및 생산 (Inventory & Production)		
		
                        LinkedObjectType	대상 화면 (오브젝트)	관련 테이블
                        "59"	입고 (Goods Receipt)	OIGN
                        "60"	출고 (Goods Issue)	OIGE
                        "67"	재고 이전 (Inventory Transfer)	OWTR
                        "202"	생산 오더 (Production Order)	OWOR
                        5. 재무 및 은행 업무 (Financials & Banking)		
		
                        LinkedObjectType	대상 화면 (오브젝트)	관련 테이블
                        "24"	수금 (Incoming Payments)	ORCT
                        "30"	분개 (Journal Entry)	OJDT
                        "46"	지급 (Outgoing Payments)	OVPM

            */

            /*
                        //oForm.Freeze(false);
            */

            if (sType == "S")  // 초기 빈 타이틀 조회시 실행 안함
            {
                ZSBO.SBO_Application.StatusBar.SetText("LoadData_Success", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
            }
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
        SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)ZSBO.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        SAPbouiCOM.ComboBox oCombo = null;
        string[] oArray;
        string oQry;

        try
        {


            // Userdatasauce 생성
            oForm.DataSources.UserDataSources.Add("PRDSTATUS", SAPbouiCOM.BoDataType.dt_LONG_TEXT, 50);

            // Binding (스크핀페인터랑 소스랑 연결하기)
            ((SAPbouiCOM.ComboBox)oForm.Items.Item("PRDSTATUS").Specific).DataBind.SetBound(true, "", "PRDSTATUS");

            oCombo = (SAPbouiCOM.ComboBox)oForm.Items.Item("PRDSTATUS").Specific;
            oCombo.ExpandType = SAPbouiCOM.BoExpandType.et_ValueDescription; //.et_DescriptionOnly;

            ((SAPbouiCOM.ComboBox)oForm.Items.Item("PRDSTATUS").Specific).ValidValues.Add("", "");
            oCombo.ValidValues.Add("계획","P");
            oCombo.ValidValues.Add("릴리스","R");
            oCombo.ValidValues.Add("마감","L");
            oCombo.ValidValues.Add("취소됨","C");

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
