//************************************************************************************************************
//  File             : ZMC201.cs
//  Module           : 프로젝트 공정관리
//  Desc             : 프로젝트 공정관리 
//  FormType         : ZMC201
//  Create Date      : 2026.10.15
//  Creator          : LEE DAE JIN
//  memo             : 
//  Copyright (c) BSNC
//************************************************************************************************************

using System;
using System.Windows.Forms;
using System.Threading;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Collections.Generic;

class ZMC201
{
    private static SAPbouiCOM.Form oForm;
    private static SAPbouiCOM.Grid oGRD1;
    private static SAPbouiCOM.ComboBox oCombo;
    private SAPbouiCOM.IEditText oEditText;


    //ZMC201A
    private static SAPbouiCOM.Form aForm;
    private static SAPbouiCOM.Grid aGrd1;

    private bool aModal = false;


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


        SAPbouiCOM.BoEventTypes EventEnum = 0;
        // For FORM_CLOSE check variable
        EventEnum = pVal.EventType;


        ///////////////////////////////////////////////////////////////////////////////////////
        /// [참조 Form] : Modal Form Open 확인                                               ///
        //////////////////////////////////////////////////////////////////////////////////////


        if (aModal)
        {
            // 현재벤트가 팝업창(ZMC201A)의 닫기 이벤트가 아닐 경우
            if (pVal.FormTypeEx == "ZMC201A")
            {
                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_CLOSE && !pVal.BeforeAction)
                    aModal = false; // 팝업창이 닫힐 때 모달 해제
            }
            else if (aForm != null)
            {
                try
                {
                    aForm.Select(); // 팝업창을 강제로 앞으로 가져옴
                }
                catch { } //BubbleEvent = false;

                BubbleEvent = false; // 기존 이벤트 전파 차단
                return;
            }
        }



        try
        {


            // 참조 Form 이벤트
            if (pVal.FormTypeEx == "ZMC201A")
            {
                ZMCLIST_ItemEventA(ref pVal, out BubbleEvent);
                return;
            }


            switch (pVal.EventType)
            {
                //=====================================================================================
                case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
                    //=====================================================================================
                    if (pVal.Before_Action & pVal.ItemUID == "BTN_SRH")
                    {
                        // if (ChkInData() == true)
                        // {
                        if (ZSBO.SBO_Application.MessageBox("조회하시겠습니까?", 1, "예", "아니오", "") == 1)
                        {
                            LoadData("S");
                        }
                        else
                        {
                            BubbleEvent = false;
                        }
                        // }
                        // else
                        // {
                        //     BubbleEvent = false;
                        // }

                    }
                    else if (pVal.BeforeAction & pVal.ItemUID == "BTN_SREF")
                    { //if (pVal.BeforeAction == false & (pVal.FormMode != (int)SAPbouiCOM.BoFormMode.fm_FIND_MODE) & pVal.ItemUID == "BTN_SREF")




                        //종복 팝업창 뜨지 않도록 
                        foreach (SAPbouiCOM.Form f in ZSBO.SBO_Application.Forms)
                        {
                            if (f.TypeEx == "ZMC201A")
                            {
                                f.Select();
                                return;
                            }
                        }


                        //핀메오더참조 팝업창 
                        SAPbouiCOM.Form mainForm = ZSBO.SBO_Application.Forms.Item(FormUID);

                        OpenPopupForm1(mainForm, "ZMC201A", "");


                        //for (int i = 0; i < ZSBO.SBO_Application.Forms.Count; i++)
                        //{
                        //    if (ZSBO.SBO_Application.Forms.Item(i).TypeEx == "ZMC201A")
                        //    {
                        //        ZSBO.SBO_Application.Forms.Item(i).Select();
                        //BubbleEvent = false;
                        //return;
                        //    }
                        // }

                    }

                    break;

                //=====================================================================================
                case SAPbouiCOM.BoEventTypes.et_CLICK:
                    //=====================================================================================


                    if (pVal.BeforeAction & pVal.ItemUID == "GRD1")
                    {
                        BubbleEvent = true;
                        // 사용자가 키를 입력하려고 하거나 값을 변경하려 할 때 이벤트를 취소
                        if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_KEY_DOWN || pVal.ColUID != "선택")
                        {
                            BubbleEvent = false; // 이벤트 전파 중단 (입력 방지)
                        }

                    }


                    //그리드 선택라인 지정 
                    if (pVal.BeforeAction & pVal.ItemUID == "GRD1" & pVal.ColUID != "RowsHeader" & pVal.Row >= 0)
                    {
                        if (oGRD1.Rows.IsSelected(pVal.Row) == false)
                        {
                            oGRD1.Rows.SelectedRows.Clear();
                            oGRD1.Rows.SelectedRows.Add(pVal.Row);
                        }
                    }
                    else if (!pVal.BeforeAction & pVal.ItemUID == "GRD1" && pVal.ColUID == "선택")
                    {

                        int selectedRow = pVal.Row;
                        string currentValue = Convert.ToString(oGRD1.DataTable.GetValue("선택", selectedRow));

                        try
                        {

                            oForm.Freeze(true);
                            // 클릭한 행이 'Y'로 체크되었을 경우
                            if (currentValue == "Y")
                            {
                                // 다른 모든 행의 Check 컬럼을 'N'으로 변경
                                for (int i = 0; i < oGRD1.DataTable.Rows.Count; i++)
                                {
                                    if (i != selectedRow)
                                    {
                                        oGRD1.DataTable.SetValue("선택", i, "N");
                                    }
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            ZSBO.SBO_Application.SetStatusBarMessage("단일 선택 처리 오류: ", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                        }
                        finally
                        {

                            oForm.Freeze(false);
                        }


                    }

                    //상단 선택 옵션 설정 
                    if (!pVal.BeforeAction & pVal.ItemUID == "RNDCONDT")
                    {
                        GetSelectedComboBoxValue();
                    }

                    //하단 판매오더 참조 설정
                    if (!pVal.BeforeAction & pVal.ItemUID == "BTN_SREF")
                    {

                        SAPbouiCOM.Form mainForm = ZSBO.SBO_Application.Forms.Item(FormUID);

                        OpenPopupForm1(mainForm, "ZMC201A", "");
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


                //==============================================================================
                case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE:
                    //==============================================================================
                    if (pVal.BeforeAction)
                    {
                        switch (ZSBO.SBO_Application.MessageBox(" 종료하시겠습니까?", 1, "예", "취소", ""))
                        {
                            case 1: oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE; break;
                            case 2: BubbleEvent = false; break;
                        }
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
            if (pVal.BeforeAction & pVal.MenuUID == "ZMC201") { LoadForm("ZMC201", ""); }

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
            //if (oForm.DataSources.UserDataSources.Item("SHIPNUM").Value == "")
            //{
            //    ZSBO.SBO_Application.SetStatusBarMessage("호선번호H(*) 지정이 누락되었습니다", SAPbouiCOM.BoMessageTime.bmt_Short, true);
            //    return false;
            //}


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
        //string oSHIPNUM = "";

        string[] oArray;

        try
        {
            oForm.Freeze(true);

            //oSHIPNUM = oForm.DataSources.UserDataSources.Item("SHIPNUM").Value.Trim();

            sQry = "EXEC dbo.[ZTMSP_ZMC201] '" + sType + "'";
            oForm.DataSources.DataTables.Item("GRD1").ExecuteQuery(sQry);
            oGRD1.DataTable = oForm.DataSources.DataTables.Item("GRD1");

            //선택 체크박스 활성화 
            ((SAPbouiCOM.EditTextColumn)oGRD1.Columns.Item("선택")).Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox;

            /* 이벤트에서 선택칼럼을 제외한 모든 칼럼의 클릭이벤트 취소 시킴, 화변 배경이 흰색유지 
             * EDIT 비활성 시키면 회으로 변경되어 바탕색을 흰색으로 유지시킴 
            //그리드 칼럼 비활성화
            ((SAPbouiCOM.EditTextColumn)oGRD1.Columns.Item("순번")).Editable = false;
            ((SAPbouiCOM.EditTextColumn)oGRD1.Columns.Item("담당자")).Editable = false;
            ((SAPbouiCOM.EditTextColumn)oGRD1.Columns.Item("문서번호")).Editable = false;
            ((SAPbouiCOM.EditTextColumn)oGRD1.Columns.Item("전표번호")).Editable = false;
            ((SAPbouiCOM.EditTextColumn)oGRD1.Columns.Item("프로젝트번호")).Editable = false;
            ((SAPbouiCOM.EditTextColumn)oGRD1.Columns.Item("조선소명")).Editable = false;
            ((SAPbouiCOM.EditTextColumn)oGRD1.Columns.Item("호선번호")).Editable = false;
            ((SAPbouiCOM.EditTextColumn)oGRD1.Columns.Item("대표호선번호")).Editable = false;
            ((SAPbouiCOM.EditTextColumn)oGRD1.Columns.Item("거래처코드")).Editable = false;
            ((SAPbouiCOM.EditTextColumn)oGRD1.Columns.Item("거래처명")).Editable = false;
            ((SAPbouiCOM.EditTextColumn)oGRD1.Columns.Item("선박유형")).Editable = false;
            ((SAPbouiCOM.EditTextColumn)oGRD1.Columns.Item("선급")).Editable = false;
            ((SAPbouiCOM.EditTextColumn)oGRD1.Columns.Item("주문품목")).Editable = false;
            ((SAPbouiCOM.EditTextColumn)oGRD1.Columns.Item("영업접수")).Editable = false;
            ((SAPbouiCOM.EditTextColumn)oGRD1.Columns.Item("YARD PLAN")).Editable = false;
            ((SAPbouiCOM.EditTextColumn)oGRD1.Columns.Item("YARD 송부")).Editable = false;
            ((SAPbouiCOM.EditTextColumn)oGRD1.Columns.Item("YARD 승인")).Editable = false;
            ((SAPbouiCOM.EditTextColumn)oGRD1.Columns.Item("CLASS 송부")).Editable = false;
            ((SAPbouiCOM.EditTextColumn)oGRD1.Columns.Item("CLASS 승인")).Editable = false;
            ((SAPbouiCOM.EditTextColumn)oGRD1.Columns.Item("납기일(최초)")).Editable = false;
            ((SAPbouiCOM.EditTextColumn)oGRD1.Columns.Item("납기일(변경)")).Editable = false;
            ((SAPbouiCOM.EditTextColumn)oGRD1.Columns.Item("계획(납기90 OR 180일 전)")).Editable = false;
            ((SAPbouiCOM.EditTextColumn)oGRD1.Columns.Item("승인일자")).Editable = false;
            ((SAPbouiCOM.EditTextColumn)oGRD1.Columns.Item("PROJECT 상태")).Editable = false;
            ((SAPbouiCOM.EditTextColumn)oGRD1.Columns.Item("판매오더번호")).Editable = false;
            ((SAPbouiCOM.EditTextColumn)oGRD1.Columns.Item("진행상황1")).Editable = false;
            ((SAPbouiCOM.EditTextColumn)oGRD1.Columns.Item("진행상황2")).Editable = false;
            ((SAPbouiCOM.EditTextColumn)oGRD1.Columns.Item("진행상황3")).Editable = false;
            ((SAPbouiCOM.EditTextColumn)oGRD1.Columns.Item("문서상태")).Editable = false;
            */

            //링크
            ((SAPbouiCOM.EditTextColumn)oGRD1.Columns.Item("전표번호")).LinkedObjectType = "112";    // 품목마스터 링크설정
            ((SAPbouiCOM.EditTextColumn)oGRD1.Columns.Item("판매오더번호")).LinkedObjectType = "17";    // 품목마스터 링크설정


            //숫자 정렬
            oGRD1.Columns.Item("순번").RightJustified = true;

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
            LoadData("S");

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
        SAPbouiCOM.ComboBox oCombo1 = null;
        SAPbouiCOM.ComboBox oCombo2 = null;
        string[] oArray;
        string oQry;

        try
        {

            // 오더기간 설정 
            oForm.DataSources.UserDataSources.Add("DOCSDATE", SAPbouiCOM.BoDataType.dt_DATE, 10);
            oForm.DataSources.UserDataSources.Add("DOCEDATE", SAPbouiCOM.BoDataType.dt_DATE, 10);

            // Binding (스크핀페인터랑 소스랑 연결하기)
            ((SAPbouiCOM.EditText)oForm.Items.Item("DOCSDATE").Specific).DataBind.SetBound(true, "", "DOCSDATE");
            ((SAPbouiCOM.EditText)oForm.Items.Item("DOCEDATE").Specific).DataBind.SetBound(true, "", "DOCEDATE");

            //초기화 
            oForm.DataSources.UserDataSources.Item("DOCSDATE").ValueEx = "";
            oForm.DataSources.UserDataSources.Item("DOCEDATE").ValueEx = "";


            // 담당자 설정 
            oCombo1 = (SAPbouiCOM.ComboBox)oForm.Items.Item("DESIGNER").Specific;
            oCombo1.ExpandType = SAPbouiCOM.BoExpandType.et_ValueDescription; //.et_DescriptionOnly;

            oQry = "SELECT FLDVALUE AS 'CODE', DESCR AS 'NAME' FROM UFD1 WHERE TABLEID = 'ORDR' AND FieldID = '56' AND FLDVALUE <> 'NONE'  ORDER BY IndexID ";
            oRecordSet.DoQuery(oQry);

            ((SAPbouiCOM.ComboBox)oForm.Items.Item("DESIGNER").Specific).ValidValues.Add("", "");
            while (oRecordSet.EoF == false)
            {
                oCombo1.ValidValues.Add(oRecordSet.Fields.Item(1).Value.ToString(), oRecordSet.Fields.Item(0).Value.ToString());
                oRecordSet.MoveNext();
            }

            // 조선소 설정 
            oCombo1 = (SAPbouiCOM.ComboBox)oForm.Items.Item("MCSHYID").Specific;
            oCombo1.ExpandType = SAPbouiCOM.BoExpandType.et_ValueDescription; //.et_DescriptionOnly;

            oQry = "SELECT CODE,NAME FROM [@ZMD154] ORDER BY DOCENTRY ";
            oRecordSet.DoQuery(oQry);

            ((SAPbouiCOM.ComboBox)oForm.Items.Item("MCSHYID").Specific).ValidValues.Add("", "");
            while (oRecordSet.EoF == false)
            {
                oCombo1.ValidValues.Add(oRecordSet.Fields.Item(1).Value.ToString(), oRecordSet.Fields.Item(0).Value.ToString());
                oRecordSet.MoveNext();
            }

            //호선번호 
            oForm.DataSources.UserDataSources.Add("MCSHMBER", SAPbouiCOM.BoDataType.dt_LONG_TEXT, 100);
            //대표호선 
            oForm.DataSources.UserDataSources.Add("MCSHBANM", SAPbouiCOM.BoDataType.dt_LONG_TEXT, 100);
            //제품군 
            oCombo2 = (SAPbouiCOM.ComboBox)oForm.Items.Item("MCORITMID").Specific;
            oCombo2.ExpandType = SAPbouiCOM.BoExpandType.et_ValueDescription; //.et_DescriptionOnly;

            oQry = "SELECT CODE,NAME FROM [@ZMD152] ";
            oRecordSet.DoQuery(oQry);

            ((SAPbouiCOM.ComboBox)oForm.Items.Item("MCORITMID").Specific).ValidValues.Add("", "");
            while (oRecordSet.EoF == false)
            {
                oCombo2.ValidValues.Add(oRecordSet.Fields.Item(1).Value.ToString(), oRecordSet.Fields.Item(0).Value.ToString());
                oRecordSet.MoveNext();
            }

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

        //상단 조건 설정 
        // 1. 콤보박스 컨트롤 객체 가져오기
        oCombo = (SAPbouiCOM.ComboBox)oForm.Items.Item("RNDCONDT").Specific;

        oCombo.ValidValues.Add("전체", "001");
        oCombo.ValidValues.Add("오더기간", "002");
        oCombo.ValidValues.Add("설계담당", "003");


        //디폴드토 001일 가져올때 
        //oCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);


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

    //************************************************************************************************************
    // Grid ItemEvent
    //************************************************************************************************************

    public void SBO_HandleGridItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
    {
        BubbleEvent = true;

        // CheckBox 클릭(CellDataChange) 이벤트 감지
        if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK //.et_CELL_DATA_CHANGE 
            && pVal.ItemUID == "GRD1"
            && pVal.ColUID == "선택"
            && !pVal.BeforeAction)
        {
            SAPbouiCOM.Form oForm = ZSBO.SBO_Application.Forms.Item(FormUID);
            SAPbouiCOM.Grid oGrid = (SAPbouiCOM.Grid)oForm.Items.Item("GRD1").Specific;
            SAPbouiCOM.DataTable oDataTable = oGrid.DataTable;

            try
            {
                oForm.Freeze(true);

                int selectedRow = pVal.Row;
                string currentValue = Convert.ToString(oDataTable.GetValue("선택", selectedRow));

                // 클릭한 행이 'Y'로 체크되었을 경우
                if (currentValue == "Y")
                {
                    // 다른 모든 행의 Check 컬럼을 'N'으로 변경
                    for (int i = 0; i < oDataTable.Rows.Count; i++)
                    {
                        if (i != selectedRow)
                        {
                            oDataTable.SetValue("선택", i, "N");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                ZSBO.SBO_Application.SetStatusBarMessage("단일 선택 처리 오류: ", SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }
    }


    //************************************************************************************************************
    // 상단조건 콤보박스 이벤트
    //************************************************************************************************************
    private void GetSelectedComboBoxValue() //SAPbouiCOM.Form oForm)
    {
        try
        {
            oCombo = (SAPbouiCOM.ComboBox)oForm.Items.Item("RNDCONDT").Specific;

            // 선택된 항목이 있는 경우에만 처리
            if (oCombo.Selected != null)
            {
                // A. 선택된 코드 (Value) 가져오기 (예: "001", "002", "003")
                //string selectedCode = oCombo.Selected.Value;

                // B. 선택된 명칭 (Description) 가져오기 (예: "전체", "담당자", "오더기간")
                string selectedDesc = oCombo.Selected.Description;

                // C. 조건문 분기 처리 예시
                if (selectedDesc == "001")
                {

                    // 전체 조회 로직 수행
                    oForm.Items.Item("DOCSDATE").Enabled = false;
                    oForm.Items.Item("DOCEDATE").Enabled = false;


                }
                else if (selectedDesc == "002")
                {
                    // 담당자 조건 조회 로직 수행
                    oForm.Items.Item("DOCSDATE").Enabled = false;
                    oForm.Items.Item("DOCEDATE").Enabled = false;
                }
                else if (selectedDesc == "003")
                {
                    // 오더기간 조건 조회 로직 수행
                    oForm.Items.Item("DOCSDATE").Enabled = true;
                    oForm.Items.Item("DOCEDATE").Enabled = true;
                }
            }
            else
            {
                ZSBO.SBO_Application.StatusBar.SetText("선택된 콤보박스 값이 없습니다.",
                    SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
            }
        }
        catch (Exception ex)
        {
            ZSBO.SBO_Application.StatusBar.SetText("Get Selected Value Error: " + ex.Message,
                SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
        }
    }



    //************************************************************************************************************
    // 하단 판매오더참조 버튼 클릭 이벤트 
    //************************************************************************************************************
    private void OpenPopupForm1(SAPbouiCOM.Form parentForm, string FileName, string oString)
    {

        //SAPbouiCOM.Form aForm = null; // oForm 변수 선언 추가

        try
        {


            string formUID = FileName; // 판매오더 참조 팝업 

            // 1. 이미 팝업 창이 열려 있는지 확인 (중복 생성 방지)
            try
            {
                SAPbouiCOM.Form existingForm = ZSBO.SBO_Application.Forms.Item(formUID);
                if (existingForm != null)
                {
                    existingForm.Select(); // 이미 열려 있다면 앞으로 가져오기
                    return;
                }
            }
            catch
            {
                // 창이 존재하지 않으면 예외가 발생하므로 새로 생성 진행
            }

            SAPbouiCOM.FormCreationParams oCreationParams = null;
            System.Xml.XmlDocument oXmlDoc = null;
            string sPath = null;

            oXmlDoc = new System.Xml.XmlDocument();
            sPath = Application.StartupPath;

            oXmlDoc.Load(sPath + @"\Screen\" + FileName + ".srf");
            oXmlDoc.SelectSingleNode("Application/forms/action/form/@uid").Value = "";
            oXmlDoc.SelectSingleNode("Application/forms/action/form/@top").Value
                = Convert.ToString(Convert.ToInt16(oXmlDoc.SelectSingleNode("Application/forms/action/form/@top").Value) + (ZSBO.SBO_Application.Forms.Count + 5));
            oXmlDoc.SelectSingleNode("Application/forms/action/form/@left").Value
                = Convert.ToString(Convert.ToInt16(oXmlDoc.SelectSingleNode("Application/forms/action/form/@left").Value) + (ZSBO.SBO_Application.Forms.Count + 5));

            // FormCreationParams 생성
            oCreationParams = ((SAPbouiCOM.FormCreationParams)(ZSBO.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams)));
            oCreationParams.XmlData = oXmlDoc.InnerXml.ToString();


            aForm = ZSBO.SBO_Application.Forms.AddEx(oCreationParams);


            aGrd1 = (SAPbouiCOM.Grid)aForm.Items.Item("GRD1").Specific;
            aForm.DataSources.DataTables.Add("GRD1");

            //popup 조건지정 활성화 
            CreateItems_aForm(aForm);



            if (parentForm != null)
            {
                // BoFormModality.fm_Modal 설정 시 부모 폼 제어가 차단됩니다.
                oCreationParams.Modality = SAPbouiCOM.BoFormModality.fm_Modal;
            }

            aForm.Visible = true;

            aModal = true;

        }
        catch (Exception ex)
        {
            ZSBO.SBO_Application.StatusBar.SetText("팝업 창 생성 오류: {ex.Message}",
                SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
        }
    }




    //************************************************************************************************************
    // CreateItems : 콤보리스트/체크박스/매트릭스/폴더/(총계:합계 구하는 컬럼도 여기서한다.)       
    //************************************************************************************************************
    private void CreateItems_aForm(SAPbouiCOM.Form xForm)
    {

        try
        {

            //상단 조건 설정 
            // 1. 콤보박스 컨트롤 객체 가져오기
            SAPbouiCOM.ComboBox oCombos = (SAPbouiCOM.ComboBox)xForm.Items.Item("CMBORDER").Specific;

            oCombos.ValidValues.Add("판매오더초안", "A");
            oCombos.ValidValues.Add("판매오더", "B");
            oCombos.Select("판매오더초안", SAPbouiCOM.BoSearchKey.psk_ByValue);

            // 오더기간 설정 
            xForm.DataSources.UserDataSources.Add("SDATE", SAPbouiCOM.BoDataType.dt_DATE, 10);
            xForm.DataSources.UserDataSources.Add("EDATE", SAPbouiCOM.BoDataType.dt_DATE, 10);

            // Binding (스크핀페인터랑 소스랑 연결하기)
            ((SAPbouiCOM.EditText)xForm.Items.Item("SDATE").Specific).DataBind.SetBound(true, "", "SDATE");
            ((SAPbouiCOM.EditText)xForm.Items.Item("EDATE").Specific).DataBind.SetBound(true, "", "EDATE");

            //초기화 
            xForm.DataSources.UserDataSources.Item("SDATE").ValueEx = "20250101";
            xForm.DataSources.UserDataSources.Item("EDATE").ValueEx = "20261031";


            xForm.Items.Item("CMBORDER").Enabled = true;
            xForm.Items.Item("SDATE").Enabled = true;
            xForm.Items.Item("EDATE").Enabled = true;

            //그리드 타이틀 표시 
            ZMSLIST_DATALOAD("B");


        }
        catch (Exception Error)
        {
            ZSBO.SBO_Application.SetStatusBarMessage("CreateItems_Error : " + Error.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        }

    }


    //************************************************************************************************************
    // [참조 Form] : ZMSLIST_ItemEventA
    //************************************************************************************************************
    private void ZMCLIST_ItemEventA(ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
    {
        string oStr;
        SAPbouiCOM.IChooseFromListEvent oCFL = null;

        BubbleEvent = true;

        try
        {
            switch (pVal.EventType)
            {
                //========================================================================================================
                case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
                    //========================================================================================================

                    // 조회버튼 클릭 
                    if (pVal.Before_Action & pVal.ItemUID == "BTN_SRH")
                    {
                        if (ZMSLIST_ChkInData() == true)
                        {
                            if (ZSBO.SBO_Application.MessageBox("조회하시겠습니까?", 1, "예", "아니오", "") == 1)
                            {
                                ZMSLIST_DATALOAD("S");
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
                    // 공정관리 등록 
                    if (pVal.Before_Action & pVal.ItemUID == "BTN_ORD")
                    {
                        if (ZMSLIST_ChkInData() == true)
                        {
                            if (ZSBO.SBO_Application.MessageBox("선택하신 판매오더를 공정관리에 등록하시겠습니까?", 1, "예", "아니오", "") == 1)
                            {
                                ZMSLIST_ORDERADD();
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

                    break;

                //========================================================================================================
                case SAPbouiCOM.BoEventTypes.et_VALIDATE:
                    //========================================================================================================
                    //if (pVal.BeforeAction == false & pVal.ItemChanged & aForm.Items.Item(pVal.ItemUID).Type == SAPbouiCOM.BoFormItemTypes.it_EDIT)
                    //{
                    //    oStr = ((SAPbouiCOM.EditText)aForm.Items.Item(pVal.ItemUID).Specific).String.Trim();
                    //    ZMSLIST_FlushToItemValueA(pVal.ItemUID, pVal.ColUID, pVal.Row, oStr);
                    //}
                    break;

                //========================================================================================================
                case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:
                    //========================================================================================================
                    if (aForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE) return;

                    break;

                //========================================================================================================
                case SAPbouiCOM.BoEventTypes.et_CLICK:
                    //========================================================================================================

                    //그리드 선택라인 지정 
                    if (!pVal.BeforeAction & pVal.ItemUID == "GRD1" & pVal.Row >= 0)  //& pVal.ColUID != "RowsHeader" 
                    {
                        if (aGrd1.Rows.IsSelected(pVal.Row) == false)
                        {
                            aGrd1.Rows.SelectedRows.Clear();
                            aGrd1.Rows.SelectedRows.Add(pVal.Row);
                        }
                    }

                    //셀 클릭방지(입력방지) 
                    if (pVal.BeforeAction & pVal.ItemUID == "GRD1")
                    {
                        BubbleEvent = true;
                        // 사용자가 키를 입력하려고 하거나 값을 변경하려 할 때 이벤트를 취소
                        if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_KEY_DOWN || pVal.ColUID != "선택")
                        {
                            BubbleEvent = false; // 이벤트 전파 중단 (입력 방지)
                        }

                    }

                    break;

                //========================================================================================================
                case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK:
                    //========================================================================================================

                    break;

                //========================================================================================================
                case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE:
                    //========================================================================================================
                    if (pVal.BeforeAction == false)
                    {
                        try
                        {
                            if (aForm.Items.Count > 0)
                            {
                                aGrd1.AutoResizeColumns();
                            }
                        }
                        catch (Exception ex)
                        {

                        }
                    }
                    break;
            }
        }
        catch (Exception ex)
        {
            /*
                        if (pVal.BeforeAction == true) BubbleEvent = false;
                        if (aForm != null)
                        {
                            aForm.Close(); aForm = null;
                        }
            */
            ZSBO.SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        }
    }

    //************************************************************************************************************
    // [참조 Form  판매오더초안+판매오더] : ZMSLIST_DATALOAD(aForm)
    //************************************************************************************************************
    private void ZMSLIST_DATALOAD(string sType)
    {
        SAPbouiCOM.EditTextColumn oCol;
        string sQry = "";
        string sOpt = "";
        string sDate = "";
        string eDate = "";
        try
        {
            aForm.Freeze(true);

            //판매오더 선택 
            SAPbouiCOM.ComboBox oComboStatus = (SAPbouiCOM.ComboBox)aForm.Items.Item("CMBORDER").Specific;
            if (oComboStatus.Selected != null)
            {
                sOpt = oComboStatus.Selected.Description.Trim();
            }

            //조회구분
            if (sType == "S")  //전기일자 적용  B: 타이틀만 적용 
            {
                sDate = aForm.DataSources.UserDataSources.Item("SDATE").Value.ToString();
                eDate = aForm.DataSources.UserDataSources.Item("EDATE").Value.ToString();
            }


            sQry = "EXEC dbo.[ZTMSP_ZMC201A] '" + sType + "','" + sOpt + "','" + sDate + "','" + eDate + "'";
            aForm.DataSources.DataTables.Item("GRD1").ExecuteQuery(sQry);
            aGrd1.DataTable = aForm.DataSources.DataTables.Item("GRD1");

            //선택 체크박스 활성화 
            ((SAPbouiCOM.EditTextColumn)aGrd1.Columns.Item("선택")).Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox;



        }
        catch (Exception Error)
        {
            //oForm.Freeze(false);
            ZSBO.SBO_Application.StatusBar.SetText("LoadData_Error : " + Error.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
        }
        finally
        {
            aForm.Freeze(false);
        }


    }

    //************************************************************************************************************
    // [참조 Form  판매오더초안+판매오더] : 데이터 조회전 전기일자 지정여부 체크 
    //************************************************************************************************************
    private bool ZMSLIST_ChkInData()
    {

        try
        {
            // 전기일자 지정 체크 
            if (aForm.DataSources.UserDataSources.Item("SDATE").Value.ToString() == "" || aForm.DataSources.UserDataSources.Item("EDATE").Value.ToString() == "")
            {
                ZSBO.SBO_Application.SetStatusBarMessage("전기일자 기간 지정이 누락되었습니다", SAPbouiCOM.BoMessageTime.bmt_Short, true);
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
    // [참조 Form  판매오더초안+판매오더] : 선택한 판매오더 공정관리에 등록 
    //************************************************************************************************************
    private void ZMSLIST_ORDERADD()
    {

        try
        {

            // 1. 전표번호를 담을 리스트 생성
            List<string> docNumList = new List<string>();

            // 2. 그리드의 DataTable 데이터 순회
            SAPbouiCOM.DataTable aDataTable = aGrd1.DataTable; // 메인 그리드 기준 (팝업이라면 aGrd1 사용)

            for (int i = 0; i < aDataTable.Rows.Count; i++)
            {
                // "선택" 컬럼의 값이 "Y"인지 확인
                string isSelected = Convert.ToString(aDataTable.GetValue("선택", i)).Trim();

                if (isSelected == "Y")
                {
                    // "전표번호" 컬럼의 값 가져오기
                    string docNum = Convert.ToString(aDataTable.GetValue("전표번호", i)).Trim();

                    if (!string.IsNullOrEmpty(docNum))
                    {
                        docNumList.Add(docNum);
                    }
                }
            }

            // 3. 선택된 항목이 없는 경우 처리
            if (docNumList.Count == 0)
            {
                ZSBO.SBO_Application.StatusBar.SetText("선택된 항목이 없습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                return;
            }


            // 4. 배열 또는 콤마로 구분된 문자열(CSV)로 변환
            string[] docNumArray = docNumList.ToArray(); // 배열로 변환 필요시
            string docNumString = string.Join(",", docNumList.ToArray());// 프로시저 전달용 (예: "1001,1002,1003")


            // 5. 데이터베이스 프로시저 호출 (Recordset 또는 Command 이용)
            string sOpt = "";
            SAPbouiCOM.ComboBox oComboStatus = (SAPbouiCOM.ComboBox)aForm.Items.Item("CMBORDER").Specific;
            if (oComboStatus.Selected != null)
            {
                sOpt = oComboStatus.Selected.Description.Trim();
            }

            ZMSLIST_ORDERWRITE(docNumString, sOpt);
        }
        catch (Exception ex)
        {
            ZSBO.SBO_Application.StatusBar.SetText("공정관리에 등록 진행시 오류가 발생하였습니다. ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
        }


    }


    //************************************************************************************************************
    // [참조 Form  판매오더초안+판매오더] : 선택한 판매오더 공정관리에 등록 > 프로시져 전달 등록 
    //************************************************************************************************************
    private void ZMSLIST_ORDERWRITE(string docNums, string sOpt)
    {
        SAPbobsCOM.Recordset oRecordSet = null;
        try
        {
            SAPbobsCOM.Recordset aRecordSet = (SAPbobsCOM.Recordset)ZSBO.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            // 전달받은 전표번호 문자열을 파라미터로 프로시저 실행(전표번호 + 구분) 
            string sQry = "EXEC dbo.[ZTMSP_ZMC201A] '{docNums}'" + "','" + sOpt + "'";
            aRecordSet.DoQuery(sQry);

            ZSBO.SBO_Application.StatusBar.SetText("선택한 판매오더를 공정관리에 등록하였습니다. ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
        }
        catch (Exception ex)
        {
            ZSBO.SBO_Application.StatusBar.SetText("선택한 판매오더를 DB에 등록시 오류가 발생하였습니다. ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
        }
        finally
        {
            // 메모리 해제
            if (oRecordSet != null)
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
            }
        }
    }




}
