//************************************************************************************************************
//  File             : ZMC202.cs
//  Module           : 프로젝트 공정관리
//  Desc             : 프로젝트 공정관리 
//  FormType         : ZMC202
//  Create Date      : 2026.10.08
//  Creator          : LEE DAE JIN
//  memo             : 매크론 설계팀 개발 요청(판매오더초안 +판매오더 수주 단계별 진행관리 체크를 위해 개발)
//                     수주관리 정보 + 관련 설계도면관리(web) 개발 연계  
//  Copyright (c) BSNC
//************************************************************************************************************

using System;
using System.Windows.Forms;
using System.Threading;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Collections.Generic;

class ZMC202
{
    private static SAPbouiCOM.Form oForm;
    private static SAPbouiCOM.Grid oGRD1;
    private static SAPbouiCOM.ComboBox oCombo;
    private SAPbouiCOM.IEditText oEditText;


    //ZMC202A
    private static SAPbouiCOM.Form aForm;
    private static SAPbouiCOM.Matrix aMAT1;
    private static SAPbouiCOM.Grid aGrd1;
    private static SAPbouiCOM.DBDataSource aDBDS;


    //ZMC202A 팝업 리사이즈용 
    private class RectI { public int L, T, W, H; }
    private Dictionary<string, RectI> aOrig = new Dictionary<string, RectI>();
    private int aOrigCW = 0;            // 오픈 시점 폼 ClientWidth
    private int aOrigCH = 0;            // 오픈 시점 폼 ClientHeight
    private bool aLayoutInit = false;


    //ZMC202B
    private static SAPbouiCOM.Form bForm;
    private static SAPbouiCOM.DBDataSource bDBDS;
    private static string bCode = "";                                                   // ZMC202B 에서 편집 중인 [@ZMC202] Code
    private static Dictionary<string, string> bOrig = new Dictionary<string, string>(); // 팝업 오픈 시점 값 (변경 여부 비교용)

    //클래스 필드
    private bool bSaving = false;      // 저장/취소 중 이벤트 차단용
    private string mainFormUID = "";   // 메인 폼 UID (팝업 오픈 시 저장해 둘 것)

    //모달
    private bool aModal = false;
    private bool bModal = false;

    private string oFormUID = "";
    //private bool bReloadPending = false;   // 공정단계관리 팝업 종료 후 재조회 요청 플래그

    private int _prevRow = -1;

    //************************************************************************************************************
    // ItemEventHander    
    //************************************************************************************************************
    public void SBO_Application_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
    {
        SAPbouiCOM.IChooseFromListEvent oCFL = null;
        string oStr, oCUID;
        BubbleEvent = true;
        //oFormUID = FormUID;

        // [수정] 메인폼(ZMC202) 이벤트일 때만 oForm/oGRD1 갱신
        // (기존: 팝업(ZMC202A) 이벤트에서도 oForm 이 팝업으로 덮어써져, 팝업을 닫은 뒤 oForm.UniqueID 에서 오류 발생 가능)
        if (pVal.FormTypeEx != "ZMC202A" && (oForm == null || oForm.UniqueID != FormUID))
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
        if (aModal)   //판매오더 참조 
        {
            // 현재벤트가 팝업창(ZMC202A)의 닫기 이벤트가 아닐 경우
            if (pVal.FormTypeEx == "ZMC202A")
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


        if (bModal)   // 공정단계관리 
        {
            // 현재벤트가 팝업창(ZMC202A)의 닫기 이벤트가 아닐 경우
            if (pVal.FormTypeEx == "ZMC202B")
            {
                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_CLOSE && !pVal.BeforeAction)
                    bModal = false; // 팝업창이 닫힐 때 모달 해제
            }
            else if (bForm != null)
            {
                try
                {
                    bForm.Select(); // 팝업창을 강제로 앞으로 가져옴
                }
                catch { } //BubbleEvent = false;

                BubbleEvent = false; // 기존 이벤트 전파 차단
                return;
            }
        }


        try
        {


            // 참조 Form 이벤트
            if (pVal.FormTypeEx == "ZMC202A")
            {
                ZMCLIST_ItemEventA(ref pVal, out BubbleEvent);

                // 팝업(ZMC202A)이 닫힌 후 메인 조회 결과 재조회
                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD && pVal.BeforeAction)
                {
                    aModal = false;      // 모달 해제 (안전장치)

                    aLayoutInit = false;     // 팝업 리사이즈 조정 취소 
                    aOrig.Clear();           // 팝업 리사이즈 조정 값 취소 

                    LoadData("S");       // 메인 폼 재조회
                }

                return;
            }

            else if (pVal.FormTypeEx == "ZMC202B")
            {
                ZMCLIST_ItemEventB(ref pVal, out BubbleEvent);

                // 팝업(ZMC202B)이 닫힌 후 메인 조회 결과 재조회
                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD && pVal.BeforeAction)
                {
                    bModal = false;      // 모달 해제 (안전장치)
                    //LoadData("S");       // 메인 폼 재조회
                }

                return;

            }


            switch (pVal.EventType)
            {
                //=====================================================================================
                case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
                    //=====================================================================================
                    if (pVal.Before_Action & pVal.ItemUID == "BTN_SRH")
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
                    else if (pVal.BeforeAction & pVal.ItemUID == "BTN_SREF")  //판매오더 참조 
                    { 

                        //종복 팝업창 뜨지 않도록 
                        foreach (SAPbouiCOM.Form f in ZSBO.SBO_Application.Forms)
                        {
                            if (f.TypeEx == "ZMC202A")
                            {
                                f.Select();
                                return;
                            }
                        }


                        //판매오더참조 팝업창 
                        SAPbouiCOM.Form mainForm = ZSBO.SBO_Application.Forms.Item(FormUID);

                        OpenPopupForm1(mainForm, "ZMC202A", "");


                    }
                    else if (pVal.BeforeAction & pVal.ItemUID == "BTN_WREF")  //공정단계관리 
                    {

                        //최종 finally에서 데이터 로드시 사용예정 
                        mainFormUID = FormUID;

                        //종복 팝업창 뜨지 않도록 
                        foreach (SAPbouiCOM.Form z in ZSBO.SBO_Application.Forms)
                        {
                            if (z.TypeEx == "ZMC202B")
                            {
                                z.Select();
                                return;
                            }
                        }


                        //공정단계관리 팝업창 : 메인 그리드에서 [선택] 체크한 라인 기준 
                        int chkCnt;
                        int selRow = GetCheckedRow(out chkCnt);

                        if (selRow < 0)
                        {
                            ZSBO.SBO_Application.StatusBar.SetText("공정단계를 관리할 라인을 [선택] 체크한 후 클릭하세요.",
                                SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                            BubbleEvent = false;
                            return;
                        }

                        if (chkCnt > 1)
                        {
                            ZSBO.SBO_Application.StatusBar.SetText("공정단계관리는 한 라인만 선택할 수 있습니다.",
                                SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                            BubbleEvent = false;
                            return;
                        }


                        //판매오더참조 팝업창 
                        SAPbouiCOM.Form mainForm = ZSBO.SBO_Application.Forms.Item(FormUID);

                        OpenPopupFormB(mainForm,"ZMC202B",selRow);

                    }
                    else if (pVal.BeforeAction & pVal.ItemUID == "BTN_CROM")  //크롬브라우저 파일관리 불러오기  
                    {

                        string url = "http://211.174.63.68:8080/main.do";

                        try
                        {
                            // 1. 크롬 실행파일을 직접 호출 (PATH에 등록된 경우)
                            System.Diagnostics.Process.Start("chrome.exe", url);
                        }
                        catch
                        {
                            try
                            {
                                // 2. 64비트 기본 경로 시도
                                string chromePath = @"C:\Program Files\Google\Chrome\Application\chrome.exe";
                                System.Diagnostics.Process.Start(chromePath, url);
                            }
                            catch
                            {
                                try
                                {
                                    // 3. 32비트 기본 경로 시도
                                    string chromePathX86 = @"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe";
                                    System.Diagnostics.Process.Start(chromePathX86, url);
                                }
                                catch (Exception ex)
                                {
                                    // 4. 모든 방법 실패 시 기본 웹 브라우저로 실행
                                    System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(url) { UseShellExecute = true });
                                }
                            }
                        }

                    }
                    break;

                //=====================================================================================
                case SAPbouiCOM.BoEventTypes.et_CLICK:
                    //=====================================================================================

                    if (pVal.ItemUID == "GRD1" & pVal.Row >= 0)   //그리드에서 발생한 클릭이벤트 + 헤더클릭이 아닌경우 
                    {

                        oGRD1.Rows.SelectedRows.Clear();  //그리드 클릭한 라인의 백그라운드 색을 지운다. 밑에 하단에서 선택한 라인의 백그라운드 색을 다시 입힌다. 

                        if (pVal.BeforeAction & pVal.ItemUID == "GRD1")
                        {
                            BubbleEvent = true;
                            // 사용자가 키를 입력하려고 하거나 값을 변경하려 할 때 이벤트를 취소
                            if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_KEY_DOWN || pVal.ColUID != "선택")
                            {
                                BubbleEvent = false; // 이벤트 전파 중단 (입력 방지)
                            }


                            oForm.Freeze(true);

                            if (_prevRow >= -1) // (_prevRow >= 0)
                            {
                                //oGRD1.CommonSetting.SetRowBackColor(_prevRow + 1, -1);   // 기본색 복원
                                for (int i = 1; i <= oGRD1.DataTable.Rows.Count; i++)
                                {
                                    oGRD1.CommonSetting.SetRowBackColor(i, -1);   // 기본색 복원
                                }


                            }

                            if (pVal.Row != -1)  //헤더를 클릭한경우 제외 
                            {
                                oGRD1.CommonSetting.SetRowBackColor(pVal.Row + 1, SelectedRowColor);
                            }
                            _prevRow = pVal.Row;

                            oForm.Freeze(false);

                        }


                        //그리드 선택라인 지정 
                        if (pVal.BeforeAction & pVal.ItemUID == "GRD1" & pVal.ColUID != "선택" & pVal.Row >= 0)
                        {
                            if (oGRD1.Rows.IsSelected(pVal.Row) == false)
                            {
                                oGRD1.Rows.SelectedRows.Clear();
                                oGRD1.Rows.SelectedRows.Add(pVal.Row);
                            }
                        }
                        else if (!pVal.BeforeAction & pVal.ItemUID == "GRD1" && pVal.ColUID == "선택" & pVal.Row >= 0)
                        {

                            try
                            {
                                oForm.Freeze(true);

                                // 1) 화면(Grid) 행 -> DataTable 행 인덱스로 변환 (정렬 후에도 정확)
                                int selectedDtRow = oGRD1.GetDataTableRowIndex(pVal.Row);
                                if (selectedDtRow < 0) return;   // 유효하지 않은 행

                                // 2) 방금 클릭한 행의 체크 값 (DataTable 인덱스 사용)
                                string currentValue = Convert.ToString(oGRD1.DataTable.GetValue("선택", selectedDtRow));

                                // 3) 'Y'로 체크된 경우, 나머지 모든 행을 'N'으로 (DataTable 기준 순회)
                                if (currentValue == "Y")
                                {
                                    int rowCount = oGRD1.DataTable.Rows.Count;
                                    for (int i = 0; i < rowCount; i++)
                                    {
                                        if (i == selectedDtRow) continue;

                                        if (Convert.ToString(oGRD1.DataTable.GetValue("선택", i)) != "N")
                                        {
                                            oGRD1.DataTable.SetValue("선택", i, "N");
                                        }
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                ZSBO.SBO_Application.SetStatusBarMessage("단일 선택 처리 오류: " + ex.Message,
                                    SAPbouiCOM.BoMessageTime.bmt_Short, true);
                            }
                            finally
                            {
                                oForm.Freeze(false);   // 예외가 나도 반드시 해제
                            }
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

                        OpenPopupForm1(mainForm, "ZMC202A", "");
                    }

                    break;

                //=====================================================================================
                case SAPbouiCOM.BoEventTypes.et_GRID_SORT:
                //=====================================================================================
                    // 그리드 헤더 더블클릭시 소트하면 라인번호를 다시 지정함.
                    if (!pVal.Before_Action)
                    {
                        if (pVal.ItemUID == "GRD1" & pVal.Row == -1) // 대상 그리드의 Item UID 확인, 헤더 클릭시 ROW는 -1 임
                        {

                            // 정렬 후 라인 번호 다시 매기기
                            RenumberGridRows(oGRD1);  //그리드와 클릭한 헤더의 칼럼명을 전달 한다. 
                            
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

                //==============================================================================
                case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED:   //링크를 클릭한경우 (판매오더인데 초안을 클릭한경우때문에 체크)
                //==============================================================================
                    string currentValue2 = Convert.ToString(oGRD1.DataTable.GetValue("진행상태", pVal.Row));
                    if ((currentValue2 == "판매오더" || currentValue2 == "판매오더(초안승인)") && pVal.ColUID == "전표번호" )  //판매오더초안(ODRF)이 아닌 판매오더(ORDR)인경우 체크
                    {
                        ZSBO.SBO_Application.SetStatusBarMessage("판매오더 번호를 클릭하세요: ", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                        BubbleEvent = false; break;
                    } 

                    break;
            }

        }
        catch (Exception Error)
        {
            ZSBO.SBO_Application.SetStatusBarMessage("SBO_Application_ItemEvent Error :" + Error.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        }

    }


    // RGB 값을 SAP가 요구하는 BGR 정수로 변환
    private static int ToSapColor(int r, int g, int b)
    {
        return r + (g * 256) + (b * 65536);
    }

    // 확인한 RGB 값을 넣으세요 (아래는 임시 추정값)
    private static readonly int SelectedRowColor = ToSapColor(255, 204, 102);

    //************************************************************************************************************
    // MenuEventHander       
    //************************************************************************************************************
    public void SBO_Application_MenuEvent(ref SAPbouiCOM.MenuEvent pVal, out bool BubbleEvent)
    {
        BubbleEvent = true;

        try
        {
            if (pVal.BeforeAction & pVal.MenuUID == "ZMC202") { LoadForm("ZMC202", ""); }

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
        int i = 0;
        //string oSHIPNUM = "";
        string oRNDCONDT = "";
        string oDOCSDATE = "";
        string oDOCEDATE = "";
        string oDESIGNER = "";
        string oMCPRJCDE = "";
        string oMCSHYID = "";
        string oMCSHMBER = "";
        string oMCSHBANM = "";
        string oMCORITMD = "";

        _prevRow = -1;  //조회시 선택 라인값을 초기화 한다. 

        string[] oArray;

        try
        {


            oForm.Freeze(true);

            //DESIGNER 담담자 콤보
            //MCPRJCDE 프로젝트 텍스트 
            //MCSHYID 조선소 콤보
            //MCSHMBER 호선번호 텍스트
            //MCSHBANM 대표호선 텍스트
            //MCORITMD 품목구분 콤보  


            SAPbouiCOM.ComboBox oComboDesigner = (SAPbouiCOM.ComboBox)oForm.Items.Item("RNDCONDT").Specific; //선택
            if (oComboDesigner.Selected != null)   //처음 로드시 null 체크 
            {
                oRNDCONDT = oComboDesigner.Selected.Description;
            }

            oDOCSDATE = oForm.DataSources.UserDataSources.Item("DOCSDATE").ValueEx; //전기 시작일 
            oDOCEDATE = oForm.DataSources.UserDataSources.Item("DOCEDATE").ValueEx; //전기 종료일 

            SAPbouiCOM.ComboBox oComboDesigner2 = (SAPbouiCOM.ComboBox)oForm.Items.Item("DESIGNER").Specific; //담당자 선택 
            if (oComboDesigner2.Selected != null)   //처음 로드시 null 체크 
            {
                oDESIGNER = ((SAPbouiCOM.ComboBox)oForm.Items.Item("DESIGNER").Specific).Value;//설계담당자 이름을 적용 
                //oDESIGNER = oComboDesigner2.Selected.Description;
            }

            oMCPRJCDE = ((SAPbouiCOM.EditText)oForm.Items.Item("MCPRJCDE").Specific).Value.Trim(); //프로젝트 

            SAPbouiCOM.ComboBox oComboDesigner3 = (SAPbouiCOM.ComboBox)oForm.Items.Item("MCSHYID").Specific; //조선소 선택
            if (oComboDesigner3.Selected != null)   //처음 로드시 null 체크 
            {
                oMCSHYID = oComboDesigner3.Selected.Description;
            }

            oMCSHMBER = ((SAPbouiCOM.EditText)oForm.Items.Item("MCSHMBER").Specific).Value.Trim(); //호선번호  
            oMCSHBANM = ((SAPbouiCOM.EditText)oForm.Items.Item("MCSHBANM").Specific).Value.Trim(); //대표호선번호  

            SAPbouiCOM.ComboBox oComboDesigner4 = (SAPbouiCOM.ComboBox)oForm.Items.Item("MCORITMD").Specific; //품목구분 선택
            if (oComboDesigner4.Selected != null)   //처음 로드시 null 체크 
            {
                oMCORITMD = oComboDesigner4.Selected.Description;
            }


            sQry = "EXEC dbo.[ZTMSP_ZMC202] '" + sType + "','" + oRNDCONDT + "','" + oDOCSDATE + "','" + oDOCEDATE + "','" + oDESIGNER + "','" + oMCPRJCDE + "','" + oMCSHYID + "','" + oMCSHMBER + "','" + oMCSHBANM + "','" + oMCORITMD + "'";
            oForm.DataSources.DataTables.Item("GRD1").ExecuteQuery(sQry);
            oGRD1.DataTable = oForm.DataSources.DataTables.Item("GRD1");
            oGRD1.RowHeaders.Width = 27;//그리드 좌측 순번 칼럼폭조정   

            //선택 체크박스 활성화 
            ((SAPbouiCOM.EditTextColumn)oGRD1.Columns.Item("선택")).Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox;

            /* 이벤트에서 선택칼럼을 제외한 모든 칼럼의 클릭이벤트 취소 시킴, 화변 배경이 흰색유지 
             * EDIT 비활성 시키면 회으로 변경되어 바탕색을 흰색으로 유지시킴 
             */
            //그리드 칼럼 비활성화
            //((SAPbouiCOM.EditTextColumn)oGRD1.Columns.Item("순번")).Editable = false;
            ((SAPbouiCOM.EditTextColumn)oGRD1.Columns.Item("담당자")).TitleObject.Sortable = true;
            ((SAPbouiCOM.EditTextColumn)oGRD1.Columns.Item("문서번호")).TitleObject.Sortable = true;
            ((SAPbouiCOM.EditTextColumn)oGRD1.Columns.Item("전표번호")).TitleObject.Sortable = true;
            ((SAPbouiCOM.EditTextColumn)oGRD1.Columns.Item("프로젝트번호")).TitleObject.Sortable = true;
            ((SAPbouiCOM.EditTextColumn)oGRD1.Columns.Item("조선소명")).TitleObject.Sortable = true;
            ((SAPbouiCOM.EditTextColumn)oGRD1.Columns.Item("호선번호")).TitleObject.Sortable = true;
            ((SAPbouiCOM.EditTextColumn)oGRD1.Columns.Item("대표호선번호")).TitleObject.Sortable = true;
            ((SAPbouiCOM.EditTextColumn)oGRD1.Columns.Item("거래처코드")).TitleObject.Sortable = true;
            ((SAPbouiCOM.EditTextColumn)oGRD1.Columns.Item("거래처명")).TitleObject.Sortable = true;
            ((SAPbouiCOM.EditTextColumn)oGRD1.Columns.Item("선박유형")).TitleObject.Sortable = true;
            ((SAPbouiCOM.EditTextColumn)oGRD1.Columns.Item("선급")).TitleObject.Sortable = true;
            ((SAPbouiCOM.EditTextColumn)oGRD1.Columns.Item("주문품목")).TitleObject.Sortable = true;
            ((SAPbouiCOM.EditTextColumn)oGRD1.Columns.Item("영업접수")).TitleObject.Sortable = true;
            ((SAPbouiCOM.EditTextColumn)oGRD1.Columns.Item("YARD PLAN")).TitleObject.Sortable = true;
            ((SAPbouiCOM.EditTextColumn)oGRD1.Columns.Item("YARD 송부")).TitleObject.Sortable = true;
            ((SAPbouiCOM.EditTextColumn)oGRD1.Columns.Item("YARD 승인")).TitleObject.Sortable = true;
            ((SAPbouiCOM.EditTextColumn)oGRD1.Columns.Item("CLASS 송부")).TitleObject.Sortable = true;
            ((SAPbouiCOM.EditTextColumn)oGRD1.Columns.Item("CLASS 승인")).TitleObject.Sortable = true;
            ((SAPbouiCOM.EditTextColumn)oGRD1.Columns.Item("납기일(최초)")).TitleObject.Sortable = true;
            ((SAPbouiCOM.EditTextColumn)oGRD1.Columns.Item("납기일(변경)")).TitleObject.Sortable = true;
            ((SAPbouiCOM.EditTextColumn)oGRD1.Columns.Item("계획(납기90 OR 180일 전)")).TitleObject.Sortable = true;
            ((SAPbouiCOM.EditTextColumn)oGRD1.Columns.Item("승인일자")).TitleObject.Sortable = true;
            ((SAPbouiCOM.EditTextColumn)oGRD1.Columns.Item("PROJECT 상태")).TitleObject.Sortable = true;
            ((SAPbouiCOM.EditTextColumn)oGRD1.Columns.Item("판매오더번호")).TitleObject.Sortable = true;
            //((SAPbouiCOM.EditTextColumn)oGRD1.Columns.Item("진행상황1")).Editable = false;
            //((SAPbouiCOM.EditTextColumn)oGRD1.Columns.Item("진행상황2")).Editable = false;
            //((SAPbouiCOM.EditTextColumn)oGRD1.Columns.Item("진행상황3")).Editable = false;
            ((SAPbouiCOM.EditTextColumn)oGRD1.Columns.Item("문서상태")).TitleObject.Sortable = true;
            

            //링크
            ((SAPbouiCOM.EditTextColumn)oGRD1.Columns.Item("전표번호")).LinkedObjectType = "112";    // 품목마스터 링크설정
            ((SAPbouiCOM.EditTextColumn)oGRD1.Columns.Item("판매오더번호")).LinkedObjectType = "17";    // 품목마스터 링크설정


            //그리드 칼럼 고정 
            oGRD1.CommonSetting.FixedColumnsCount = 7;
            oGRD1.AutoResizeColumns();


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

            //그리드 라인번호 지정
            RenumberGridRows(oGRD1);

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
    // 결과 조회후 그리드 라인번호 재지정          
    //************************************************************************************************************
    public void RenumberGridRows(SAPbouiCOM.Grid oGRD1)
    {

        oForm.Freeze(true);
 

        try
        {
  

            for (int i = 0; i < oGRD1.Rows.Count; i++)
            {
                oGRD1.RowHeaders.SetText(i, (i + 1).ToString()); //라인번호 지정 
                oGRD1.CommonSetting.SetRowBackColor(i+1, -1);   // 기본색 복원
                oGRD1.DataTable.SetValue("선택", i, "N");  //선택 칼럼에 체크가 있으면 초기화 
                
            }

            _prevRow = -1;

            oForm.Freeze(false);

        }
        catch (Exception Error) 
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
            _prevRow = -1;

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


            // add a new form
            oCreationParams = ((SAPbouiCOM.FormCreationParams)(ZSBO.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams)));
            oCreationParams.XmlData = oXmlDoc.InnerXml.ToString();

            oForm = ZSBO.SBO_Application.Forms.AddEx(oCreationParams);
            oForm.Title = ZSBO.ZAPI.FGetFormTitle(oForm.Title, FileName, oForm.TypeEx, oForm.BusinessObject.Type);
            oGRD1 = (SAPbouiCOM.Grid)oForm.Items.Item("GRD1").Specific;
            oForm.DataSources.DataTables.Add("GRD1");

            CreateItems();
            SetFormBase();

            if (oCombo.Selected == null)  //첫화면 로드시 선택콤보 없음
            {
                LoadData("B");  // 타이틀면 조회하도록 조정
            }
            else
            {
                LoadData("S");  // 콤보가 선택된경우 해당 선택값 기준으로 조회 
            }

            oForm.Visible = true;           // 기본으로 false

            // 첫번째 조건에 커서 지정 
            oForm.ActiveItem = "RNDCONDT";


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

            //프로젝트번호
            oForm.DataSources.UserDataSources.Add("MCPRJCDE", SAPbouiCOM.BoDataType.dt_LONG_TEXT, 100);

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
            oCombo2 = (SAPbouiCOM.ComboBox)oForm.Items.Item("MCORITMD").Specific;
            oCombo2.ExpandType = SAPbouiCOM.BoExpandType.et_ValueDescription; //.et_DescriptionOnly;

            oQry = "SELECT CODE,NAME FROM [@ZMD152] ";
            oRecordSet.DoQuery(oQry);

            ((SAPbouiCOM.ComboBox)oForm.Items.Item("MCORITMD").Specific).ValidValues.Add("", "");
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

        oCombo.ValidValues.Add("전체관리", "A");
       // oCombo.ValidValues.Add("오더기간", "B");
       // oCombo.ValidValues.Add("설계담당", "C");
        oCombo.ValidValues.Add("공정완료", "D");


        //디폴드토 001일 가져올때 
        //oCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);


    }

    //************************************************************************************************************
    // SetToolBarEnabled        
    //************************************************************************************************************
    private void SetToolBarEnabled_oform()
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

            oForm.Freeze(true);

            oCombo = (SAPbouiCOM.ComboBox)oForm.Items.Item("RNDCONDT").Specific;

            // 선택된 항목이 있는 경우에만 처리
            if (oCombo.Selected != null)
            {
                // A. 선택된 코드 (Value) 가져오기 (예: "001", "002", "003")
                //string selectedCode = oCombo.Selected.Value;

                // B. 선택된 명칭 (Description) 가져오기 (예: "전체", "담당자", "오더기간")
                string selectedDesc = oCombo.Selected.Description;

                // C. 조건문 분기 처리 예시
                if (selectedDesc == "A")
                {

                    // 전체 조회 로직 수행

                    oForm.DataSources.UserDataSources.Item("DOCSDATE").ValueEx = "";
                    oForm.DataSources.UserDataSources.Item("DOCEDATE").ValueEx = "";

                    oForm.Items.Item("DOCSDATE").Enabled = true;
                    oForm.Items.Item("DOCEDATE").Enabled = true;

                    oForm.Items.Item("DESIGNER").Enabled = true; //설계담당자 
                    oForm.Items.Item("MCPRJCDE").Enabled = true; //프로젝트번호
                    oForm.Items.Item("MCSHYID").Enabled = true; //조선소 
                    oForm.Items.Item("MCSHMBER").Enabled = true; //호선번호 
                    oForm.Items.Item("MCSHBANM").Enabled = true; //대표호선 
                    oForm.Items.Item("MCORITMD").Enabled = true; //품목구분 


                }
/*
                else if (selectedDesc == "B")
                {
                    // 오더기간 조건 조회 로직 수행
                    oForm.Items.Item("DOCSDATE").Enabled = true;
                    oForm.Items.Item("DOCEDATE").Enabled = true;

                    oForm.Items.Item("DESIGNER").Enabled = true; //설계담당자 
                    oForm.Items.Item("MCPRJCDE").Enabled = true; //프로젝트번호
                    oForm.Items.Item("MCSHYID").Enabled = true; //조선소 
                    oForm.Items.Item("MCSHMBER").Enabled = true; //호선번호 
                    oForm.Items.Item("MCSHBANM").Enabled = true; //대표호선 
                    oForm.Items.Item("MCORITMD").Enabled = true; //품목구분 

                }
                else if (selectedDesc == "C")
                {
                    // 설계담당자 조회 
                    oForm.Items.Item("DOCSDATE").Enabled = false;
                    oForm.Items.Item("DOCEDATE").Enabled = false;

                    oForm.Items.Item("DESIGNER").Enabled = true; //설계담당자 
                    oForm.Items.Item("MCPRJCDE").Enabled = true; //프로젝트번호
                    oForm.Items.Item("MCSHYID").Enabled = true; //조선소 
                    oForm.Items.Item("MCSHMBER").Enabled = true; //호선번호 
                    oForm.Items.Item("MCSHBANM").Enabled = true; //대표호선 
                    oForm.Items.Item("MCORITMD").Enabled = true; //품목구분 

                }
 */ 
                else if (selectedDesc == "D")
                {
                    // 공정관리 완료 문서 조회 
                    // 공정관리 완료 문서는 현재 관리 하는 문서가 아닌, 완료후 관리제외 문서이기에 전체를 
                    // 조회하는 것으로 하고, 추후 데이터가 싸이면 조건을 붙여 관리한는 것으로 함.
                    oForm.DataSources.UserDataSources.Item("DOCSDATE").ValueEx = "";
                    oForm.DataSources.UserDataSources.Item("DOCEDATE").ValueEx = "";
                    oForm.Items.Item("DOCSDATE").Enabled = true;
                    oForm.Items.Item("DOCEDATE").Enabled = true;

                    oForm.Items.Item("DESIGNER").Enabled = true; //설계담당자 
                    oForm.Items.Item("MCPRJCDE").Enabled = true; //프로젝트번호
                    oForm.Items.Item("MCSHYID").Enabled = true; //조선소 
                    oForm.Items.Item("MCSHMBER").Enabled = true; //호선번호 
                    oForm.Items.Item("MCSHBANM").Enabled = true; //대표호선 
                    oForm.Items.Item("MCORITMD").Enabled = true; //품목구분 




                }

                //초기화 
                oForm.DataSources.UserDataSources.Item("DOCSDATE").ValueEx = "";
                oForm.DataSources.UserDataSources.Item("DOCEDATE").ValueEx = "";

                SAPbouiCOM.ComboBox oComboDesigner = (SAPbouiCOM.ComboBox)oForm.Items.Item("DESIGNER").Specific;
                oComboDesigner.Select(0, SAPbouiCOM.BoSearchKey.psk_Index); // 0번째(첫 번째) 항목 선택

                SAPbouiCOM.ComboBox oComboDesigner1 = (SAPbouiCOM.ComboBox)oForm.Items.Item("MCSHYID").Specific;
                oComboDesigner1.Select(0, SAPbouiCOM.BoSearchKey.psk_Index); // 0번째(첫 번째) 항목 선택

                ((SAPbouiCOM.EditText)oForm.Items.Item("MCPRJCDE").Specific).Value = "";
                ((SAPbouiCOM.EditText)oForm.Items.Item("MCSHMBER").Specific).Value = "";
                ((SAPbouiCOM.EditText)oForm.Items.Item("MCSHBANM").Specific).Value = "";

                SAPbouiCOM.ComboBox oComboDesigner3 = (SAPbouiCOM.ComboBox)oForm.Items.Item("MCORITMD").Specific;
                oComboDesigner3.Select(0, SAPbouiCOM.BoSearchKey.psk_Index); // 0번째(첫 번째) 항목 선택

                oForm.ActiveItem = "RNDCONDT";//선택 곰보에 커서를 지정한다. 


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
        finally
        {
            oForm.Freeze(false);
        }
    }



    //************************************************************************************************************
    // 하단 판매오더참조 버튼 클릭 이벤트 
    //************************************************************************************************************
    private void OpenPopupForm1(SAPbouiCOM.Form parentForm, string FileName, string oString)
    {

        try
        {


            string formUID = FileName; // 판매오더 참조 팝업 

            // 1. 이미 팝업 창이 열려 있는지 확인 (중복 생성 방지)
            //--------------------------------------------------------
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

            aForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

            //popup 조건지정 활성화 
            CreateItems_aForm(aForm);


            if (parentForm != null)
            {
                // BoFormModality.fm_Modal 설정 시 부모 폼 제어가 차단됩니다.
                oCreationParams.Modality = SAPbouiCOM.BoFormModality.fm_Modal;
            }

            aForm.Visible = true;

            InitPopupLayout(aForm);     // 팝업 사이즈 조정 추가

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
            xForm.DataSources.UserDataSources.Item("SDATE").ValueEx = DateTime.Today.ToString("yyyyMMdd"); //초기값으로 오늘 날짜를 지정 
            xForm.DataSources.UserDataSources.Item("EDATE").ValueEx = DateTime.Today.ToString("yyyyMMdd"); //초기값으로 오늘 날짜를 지정  


            xForm.Items.Item("CMBORDER").Enabled = true;
            xForm.Items.Item("SDATE").Enabled = true;
            xForm.Items.Item("EDATE").Enabled = true;

            //그리드 타이틀 표시 
            ZMSLIST_DATALOAD("S");


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
            // [수정] 팝업 폼 객체(aForm/aGrd1/aMAT1/aDBDS)를 이벤트가 발생한 폼 기준으로 갱신
            if (pVal.EventType != SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD &&
                pVal.EventType != SAPbouiCOM.BoEventTypes.et_FORM_CLOSE)
            {
                try { RefreshPopupRefs(pVal.FormUID); }
                catch { }   // 폼 생성 초기 이벤트 등에서는 무시
            }

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
                    //선택한 판매오더 리스트 매트릭스에 전달 
                    else if (pVal.Before_Action & pVal.ItemUID == "BTN_SLT")
                    {
                        if (ZSBO.SBO_Application.MessageBox("선택하신 리스트를 저장 단계로 이관하겠습니까?", 1, "예", "아니오", "") == 1)
                        {
                            ZMSLIST_DATALOAD("N");
                        }
                        else
                        {
                            BubbleEvent = false;
                        }


                    }
                    // 공정관리 등록 : 매트릭스 값을 [@ZMC202] 테이블에 저장
                    else if (pVal.BeforeAction & pVal.ItemUID == "BTN_ORD")
                    {
                        BubbleEvent = false;   // 일반 버튼이므로 기본 처리 불필요 (저장은 코드에서 수행)

                        if (ZMCLIST_ChkInData() == false)
                        {
                            ZSBO.SBO_Application.StatusBar.SetText("등록할 내역이 없습니다. 판매오더 선택 후 [선택적용]을 먼저 실행하세요.",
                                SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                            return;
                        }

                        if (ZSBO.SBO_Application.MessageBox("선택하신 판매오더를 공정관리에 등록하시겠습니까?", 1, "예", "아니오", "") != 1) return;

                        ZMCLIST_DataSave();
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


                    //매트릭스 메뉴 이벤트 적용 
                    if (pVal.BeforeAction & pVal.ItemUID == "MAT1")
                    {

                        //매트릭스 메뉴 추가 
                        SetToolBarEnabled();


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
                            ApplyPopupLayout(aForm);     // ★ 최대화/복원 시 크기 조정

                            if (aForm.Items.Count > 0)
                            {
                                aGrd1.AutoResizeColumns();
                            }
                        }
                        catch (Exception ex)
                        {
                            ZSBO.SBO_Application.StatusBar.SetText("리사이즈 오류 : " + ex.Message,
               SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);

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
        SAPbobsCOM.Recordset aRecordSet = (SAPbobsCOM.Recordset)ZSBO.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        SAPbouiCOM.EditTextColumn oCol;
        int i;
        string sQry = "";
        string sOpt = "";
        string sDate = "";
        string eDate = "";
        try
        {

            aForm.Freeze(true);

            //조회구분
            if (sType == "S")  //전기일자 적용  B: 타이틀만 적용 
            {

                //판매오더 선택 
                SAPbouiCOM.ComboBox oComboStatus = (SAPbouiCOM.ComboBox)aForm.Items.Item("CMBORDER").Specific;
                if (oComboStatus.Selected != null)
                {
                    sOpt = oComboStatus.Selected.Description.Trim();
                }

                sDate = aForm.DataSources.UserDataSources.Item("SDATE").Value.ToString();
                eDate = aForm.DataSources.UserDataSources.Item("EDATE").Value.ToString();



                sQry = "EXEC dbo.[ZTMSP_ZMC202A] '" + sType + "','" + sOpt + "','" + sDate + "','" + eDate + "'";
                aForm.DataSources.DataTables.Item("GRD1").ExecuteQuery(sQry);
                aGrd1.DataTable = aForm.DataSources.DataTables.Item("GRD1");

                //선택 체크박스 활성화 
                ((SAPbouiCOM.EditTextColumn)aGrd1.Columns.Item("선택")).Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox;

                if (sOpt == "A")  //판매오더 초안 
                {
                    ((SAPbouiCOM.EditTextColumn)aGrd1.Columns.Item("전표번호")).LinkedObjectType = "112";    // 판매오더초안 링크설정
                }
                else if (sOpt == "B" || sOpt == "C")  //판매오더(B) 처음부터 판매오더를 선택하여 관리하는 문서, 판매오더(C) 판매오더초안을 승인하여 판매오더로 넘어온 문서 
                {
                    ((SAPbouiCOM.EditTextColumn)aGrd1.Columns.Item("전표번호")).LinkedObjectType = "17";    // 판매오더초안 링크설정
                }

                ((SAPbouiCOM.EditTextColumn)aGrd1.Columns.Item("고객")).LinkedObjectType = "2";    // BP마스터 링크설정


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



            }
            else if (sType == "N")   //선택한 판매오더 리스트 매트릭스로 이관 
            {

                //판애오더초안,판매오더 구분 선택 
                SAPbouiCOM.ComboBox oComboStatus = (SAPbouiCOM.ComboBox)aForm.Items.Item("CMBORDER").Specific;
                if (oComboStatus.Selected != null)
                {
                    sOpt = oComboStatus.Selected.Description.Trim();
                }

                aMAT1 = (SAPbouiCOM.Matrix)aForm.Items.Item("MAT1").Specific;   // [수정] oForm -> aForm
                aDBDS = aForm.DataSources.DBDataSources.Item("@ZMC202");

                int aRow = 0;
                string chkY = "N";
                for (aRow = 0; aRow < aGrd1.Rows.Count; aRow++)
                {
                    chkY = aGrd1.DataTable.Columns.Item(0).Cells.Item(aRow).Value.ToString();
                    if (chkY == "Y") break;
                }

                if (chkY == "N")
                {
                    ZSBO.SBO_Application.MessageBox("선택한 내역이 없습니다!", 1, "확인", "", "");
                    aForm.Freeze(false);
                    return;
                }


                //선택 품번에 대한 Matrix 반영
                aMAT1.Clear();
                aDBDS.Clear();      // [수정] 이전 적용분 제거 (DBDS 가 남아 있으면 SetValue 인덱스가 어긋남)

                int oRow = 0;
                for (int n = 0; n < aGrd1.Rows.Count; n++)
                {
                    string selY = aGrd1.DataTable.Columns.Item(0).Cells.Item(n).Value.ToString();
                    if (selY == "Y")
                    {
                        aDBDS.InsertRecord(aDBDS.Size);     // 맨 뒤에 신규 레코드 추가
                        int idx = aDBDS.Size - 1;            // [수정] 방금 추가된 레코드의 인덱스
                        aDBDS.Offset = idx;

                        aDBDS.SetValue("U_PDOCNUM", idx, aGrd1.DataTable.Columns.Item("문서번호").Cells.Item(n).Value.ToString());
                        aDBDS.SetValue("U_PDOCENTY", idx, aGrd1.DataTable.Columns.Item("전표번호").Cells.Item(n).Value.ToString());
                        aDBDS.SetValue("U_CARDCODE", idx, aGrd1.DataTable.Columns.Item("고객").Cells.Item(n).Value.ToString());
                        aDBDS.SetValue("U_CARDNAME", idx, aGrd1.DataTable.Columns.Item("고객명").Cells.Item(n).Value.ToString());
                        aDBDS.SetValue("U_DESIGNER", idx, aGrd1.DataTable.Columns.Item("담당자명").Cells.Item(n).Value.ToString());
                        aDBDS.SetValue("U_PRJCODE", idx, aGrd1.DataTable.Columns.Item("프로젝트번호").Cells.Item(n).Value.ToString());
                        aDBDS.SetValue("U_SHYIDNM", idx, aGrd1.DataTable.Columns.Item("조선소명").Cells.Item(n).Value.ToString());
                        aDBDS.SetValue("U_SHMBER", idx, aGrd1.DataTable.Columns.Item("호선번호").Cells.Item(n).Value.ToString());
                        aDBDS.SetValue("U_BASHNM", idx, aGrd1.DataTable.Columns.Item("대표호선번호").Cells.Item(n).Value.ToString());
                        aDBDS.SetValue("U_DOCGBN", idx, sOpt.ToString());
                        oRow++;
                    }
                }

                if (oRow > 0)
                {
                    //그리드 선택 라인 해제 
                    for (int n = 0; n < aGrd1.Rows.Count; n++)
                    {
                        aGrd1.DataTable.Columns.Item("선택").Cells.Item(n).Value = "0";
                    }

                    aMAT1.LoadFromDataSource();

                }
                else
                {
                    // 체크박스
                    aGrd1.Columns.Item(0).Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox;
                    ZSBO.SBO_Application.MessageBox("선택한 판매오더가 없습니다.", 1, "확인", "", "");
                    aForm.Freeze(false);
                }



            }

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
            aForm.Freeze(false);
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
            string sQry = "EXEC dbo.[ZTMSP_ZMC202A] '{docNums}'" + "','" + sOpt + "'";
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


    //************************************************************************************************************
    // 데이터 체크: 저장전 데이터 체크 
    //************************************************************************************************************
    private bool ZMCLIST_ChkInData()
    {
        aMAT1 = (SAPbouiCOM.Matrix)aForm.Items.Item("MAT1").Specific;   // [수정] oForm -> aForm
        if (aMAT1.RowCount > 0)
        {
            return true;
        }
        else
        {
            return false;
        }

    }


    //************************************************************************************************************
    // Matrix Data Save 
    //************************************************************************************************************
    private const string UDO_CODE = "ZMC202";     // ★ OUDO.Code 와 일치해야 함 : SELECT * FROM OUDO WHERE TableName = 'ZMC202'
    private const bool CHECK_DUPLICATE = true;    // 동일 전표번호 + 구분(초안/오더) 중복등록 방지

    // 팝업 폼 객체 핸들 갱신
    private void RefreshPopupRefs(string formUID)
    {
        aForm = ZSBO.SBO_Application.Forms.Item(formUID);
        aGrd1 = (SAPbouiCOM.Grid)aForm.Items.Item("GRD1").Specific;
        aMAT1 = (SAPbouiCOM.Matrix)aForm.Items.Item("MAT1").Specific;
        aDBDS = aForm.DataSources.DBDataSources.Item("@ZMC202");
    }

    // [@ZMC202] = Master Data 형 UDO 테이블 (Code/Name 구조, LineId/자식 테이블 없음)
    //  → 매트릭스 1행 = [@ZMC202] 1건 (GeneralData 헤더 1건)
    private bool ZMCLIST_DataSave()
    {
        SAPbobsCOM.Recordset rs = null;
        SAPbobsCOM.GeneralService svc = null;
        bool startedTrans = false;
        int saved = 0, dup = 0;

        try
        {

            aForm.Freeze(true);

            // 1) 화면(매트릭스) 값을 DBDS 로 반영
            aMAT1.FlushToDataSource();

            if (aDBDS.Size == 0)
            {
                ZSBO.SBO_Application.StatusBar.SetText("등록할 내역이 없습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                aForm.Freeze(false);
                return false;
            }

            rs = (SAPbobsCOM.Recordset)ZSBO.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            // Code 채번
            rs.DoQuery("SELECT ISNULL(MAX(TRY_CONVERT(int, Code)), 0) + 1 FROM [@ZMC202]");
            int nextCode = Convert.ToInt32(rs.Fields.Item(0).Value);

            // ▼ UDO 코드 조회 (기존 GetGeneralService(UDO_CODE) 줄 대체)
            rs.DoQuery("SELECT TOP 1 Code FROM OUDO WHERE TableName IN ('ZMC202', '@ZMC202')");
            if (rs.RecordCount == 0)
                throw new Exception("[@ZMC202] 테이블에 연결된 UDO가 등록되어 있지 않습니다. (OUDO 확인)");
            string udoCode = rs.Fields.Item(0).Value.ToString().Trim();

            svc = ZSBO.oCompany.GetCompanyService().GetGeneralService(udoCode);

            // 3) 트랜잭션 시작 (여러 건 중 하나라도 실패하면 전체 롤백)
            if (!ZSBO.oCompany.InTransaction)
            {
                ZSBO.oCompany.StartTransaction();
                startedTrans = true;
            }

            for (int i = 0; i < aDBDS.Size; i++)
            {
                string sEntry = aDBDS.GetValue("U_PDOCENTY", i).Trim();
                if (sEntry == "") continue;                       // 빈 행 제외

                int pDocEntry;
                int pDocNum;
                if (!int.TryParse(sEntry, out pDocEntry))
                    throw new Exception((i + 1).ToString() + "행의 전표번호가 올바르지 않습니다. (" + sEntry + ")");
                int.TryParse(aDBDS.GetValue("U_PDOCNUM", i).Trim(), out pDocNum);

                string sGbn = aDBDS.GetValue("U_DOCGBN", i).Trim();

                // 중복 체크 (이미 등록된 건은 제외)
                if (CHECK_DUPLICATE)
                {
                    rs.DoQuery("SELECT COUNT(*) FROM [@ZMC202] WHERE U_PDOCENTY = " + pDocEntry.ToString()
                             + " AND ISNULL(U_DOCGBN,'') = '" + sGbn.Replace("'", "''") + "' AND CANCELED = 'N'");
                    if (Convert.ToInt32(rs.Fields.Item(0).Value) > 0) { dup++; continue; }
                }

                // 4) UDO 등록
                SAPbobsCOM.GeneralData gd = (SAPbobsCOM.GeneralData)svc.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData);

                string sCode = (nextCode++).ToString();
                gd.SetProperty("Code", sCode);
                gd.SetProperty("Name", sCode);
                gd.SetProperty("U_PDOCNUM", pDocNum);        // int 필드는 int 로 전달
                gd.SetProperty("U_PDOCENTY", pDocEntry);
                gd.SetProperty("U_CARDCODE", aDBDS.GetValue("U_CARDCODE", i).Trim());
                gd.SetProperty("U_CARDNAME", aDBDS.GetValue("U_CARDNAME", i).Trim());
                gd.SetProperty("U_DESIGNER", aDBDS.GetValue("U_DESIGNER", i).Trim());
                gd.SetProperty("U_PRJCODE", aDBDS.GetValue("U_PRJCODE", i).Trim());
                gd.SetProperty("U_SHYIDNM", aDBDS.GetValue("U_SHYIDNM", i).Trim());
                gd.SetProperty("U_SHMBER", aDBDS.GetValue("U_SHMBER", i).Trim());
                gd.SetProperty("U_BASHNM", aDBDS.GetValue("U_BASHNM", i).Trim());
                gd.SetProperty("U_DOCGBN", sGbn);

                svc.Add(gd);
                saved++;
            }

            // 5) 커밋
            if (startedTrans) { ZSBO.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit); startedTrans = false; }

            string msg = "공정관리에 " + saved.ToString() + "건 등록하였습니다.";
            if (dup > 0) msg += " (이미 등록된 " + dup.ToString() + "건 제외)";

            ZSBO.SBO_Application.StatusBar.SetText(msg, SAPbouiCOM.BoMessageTime.bmt_Short,
                saved > 0 ? SAPbouiCOM.BoStatusBarMessageType.smt_Success : SAPbouiCOM.BoStatusBarMessageType.smt_Warning);

            // 저장 완료분은 화면에서 제거 (재등록 방지)
            if (saved > 0)
            {
                aMAT1.Clear();
                aDBDS.Clear();

                ZMSLIST_DATALOAD("S"); //저장 문서 제외후 재조회 
            }

            return saved > 0;
        }
        catch (Exception ex)
        {
            if (startedTrans && ZSBO.oCompany.InTransaction)
            {
                try { ZSBO.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack); }
                catch { }
            }

            // DI API 상세 오류 (UDO 코드 오류, 필수값 누락, Code 중복 등)
            int errCode; string errMsg;
            ZSBO.oCompany.GetLastError(out errCode, out errMsg);
            string detail = (errCode != 0) ? ("[" + errCode.ToString() + "] " + errMsg) : ex.Message;

            ZSBO.SBO_Application.StatusBar.SetText("공정관리 등록 오류 : " + detail,
                SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            aForm.Freeze(false);
            return false;
        }
        finally
        {
            if (rs != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(rs);

            aForm.Freeze(false);
        }
    }



    //************************************************************************************************************
    // SetToolBarEnabled
    //************************************************************************************************************
    private void SetToolBarEnabled()
    {
        aForm.Freeze(true);

        aForm.EnableMenu("1283", false);    // 제거
        aForm.EnableMenu("1284", false);    // 취소
        aForm.EnableMenu("1285", false);    // 복원
        aForm.EnableMenu("1286", false);    // 닫기
        aForm.EnableMenu("1287", false);    // 복제

        aForm.EnableMenu("1293", true);    // 라인삭제

 
        aForm.Freeze(false);
    }


    //************************************************************************************************************
    // [참조 Form] : ZMSLIST_ItemEventB   프로젝트공정관리 공정단계관리 이벤트 관리 
    //************************************************************************************************************
    private void ZMCLIST_ItemEventB(ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
    {
        string oStr;
        SAPbouiCOM.IChooseFromListEvent oCFL = null;



        BubbleEvent = true;

        try
        {
            // [수정] 팝업 폼 객체(aForm/aGrd1/aMAT1/aDBDS)를 이벤트가 발생한 폼 기준으로 갱신
            if (pVal.EventType != SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD &&
                pVal.EventType != SAPbouiCOM.BoEventTypes.et_FORM_CLOSE)
            {
                try { RefreshPopupRefs(pVal.FormUID); }
                catch { }   // 폼 생성 초기 이벤트 등에서는 무시
            }

            switch (pVal.EventType)
            {
                //========================================================================================================
                case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
                    //========================================================================================================


                    // 등록 : Enabled=true 입력박스 값을 UDO([@ZMC202]) 로 Update
                    if (pVal.ItemUID == "BTN_UPT")
                    {
                        if (pVal.BeforeAction)
                        {
                            if (ZSBO.SBO_Application.MessageBox("입력하신 공정단계 내용을 등록(수정)하시겠습니까?", 1, "예", "아니오", "") != 1)
                                BubbleEvent = false;
                        }
                        else
                        {
                            // After 에서 처리 : 저장 후 팝업을 닫아도 이벤트 처리 중 폼이 사라지는 문제 없음
                            ZMCB_DataSave();
                        }
                    }

                    break;


                //========================================================================================================
                case SAPbouiCOM.BoEventTypes.et_VALIDATE:

                    break;

                //========================================================================================================
                case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:
                    //========================================================================================================
                    if (aForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE) return;

                    break;

                //========================================================================================================
                case SAPbouiCOM.BoEventTypes.et_CLICK:
                    //========================================================================================================

                    break;

                //========================================================================================================
                case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK:
                    //========================================================================================================

                    break;

                //========================================================================================================
                case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE:
                    //========================================================================================================

                    break;
                case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD:

                    //LoadData("S");

                    break;
            }
        }
        catch (Exception ex)
        {

            ZSBO.SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        }
    }



    //************************************************************************************************************
    // 메인 그리드에서 [선택]=Y 인 첫 라인 (없으면 -1), count = 체크된 라인 수
    //************************************************************************************************************
    private int GetCheckedRow(out int count)
    {
        int first = -1;
        count = 0;

        SAPbouiCOM.DataTable dt = oGRD1.DataTable;
        for (int i = 0; i < dt.Rows.Count; i++)
        {
            if (Convert.ToString(dt.GetValue("선택", i)).Trim() == "Y")
            {
                if (first < 0) first = i;
                count++;
            }
        }
        return first;
    }


    //************************************************************************************************************
    // 메인 [공정단계관리] 버튼 : ZMC202B 팝업 오픈 + 선택 라인 수주정보 표시
    //************************************************************************************************************
    private void OpenPopupFormB(SAPbouiCOM.Form parentForm,string FileName,int selRow)
    {
        try
        {
            // 1) 선택 라인의 UDO Code 확인 (공정관리에 등록되지 않은 라인이면 팝업을 열지 않음)
            string sCode = ResolveUdoCode(selRow);
            if (sCode == "") return;

            // 2) 화면 로드 (OpenPopupForm1 과 동일 방식)
            System.Xml.XmlDocument oXmlDoc = new System.Xml.XmlDocument();
            string sPath = Application.StartupPath;

            oXmlDoc.Load(sPath + @"\Screen\"+FileName+".srf");
            oXmlDoc.SelectSingleNode("Application/forms/action/form/@uid").Value = "";
            oXmlDoc.SelectSingleNode("Application/forms/action/form/@top").Value
                = Convert.ToString(Convert.ToInt16(oXmlDoc.SelectSingleNode("Application/forms/action/form/@top").Value) + (ZSBO.SBO_Application.Forms.Count + 5));
            oXmlDoc.SelectSingleNode("Application/forms/action/form/@left").Value
                = Convert.ToString(Convert.ToInt16(oXmlDoc.SelectSingleNode("Application/forms/action/form/@left").Value) + (ZSBO.SBO_Application.Forms.Count + 5));

            SAPbouiCOM.FormCreationParams oCreationParams
                = ((SAPbouiCOM.FormCreationParams)(ZSBO.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams)));
            oCreationParams.XmlData = oXmlDoc.InnerXml.ToString();

            bForm = ZSBO.SBO_Application.Forms.AddEx(oCreationParams);

            // 3) 아이템 바인딩 + 선택 라인 값 표시
            bCode = sCode;
            bOrig.Clear();

            CreateItems_bForm(bForm);
            LoadSelectedToB(selRow);

            //bForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;

            if (parentForm != null)
            {
                // BoFormModality.fm_Modal 설정 시 부모 폼 제어가 차단됩니다.
                oCreationParams.Modality = SAPbouiCOM.BoFormModality.fm_Modal;
            }


            bForm.Visible = true;
            bModal = true;
        }
        catch (Exception ex)
        {
            ZSBO.SBO_Application.StatusBar.SetText("공정단계관리 팝업 생성 오류 : " + ex.Message,
                SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
        }
    }


    //************************************************************************************************************
    // ZMC202B 아이템 → UserDataSource 바인딩
    //   (Enabled=false 박스도 UserDataSource 로 값을 넣을 수 있음)
    //************************************************************************************************************
    private void CreateItems_bForm(SAPbouiCOM.Form xForm)
    {
        foreach (BFld f in B_FIELDS)
        {
            string uds = f.Item;  //"UD_" + f.Item;

            if (f.Kind == 'D')
                xForm.DataSources.UserDataSources.Add(uds, SAPbouiCOM.BoDataType.dt_DATE, 10);
            else if (f.Kind == 'M')
                xForm.DataSources.UserDataSources.Add(uds, SAPbouiCOM.BoDataType.dt_LONG_TEXT, 2000);
            else if (f.Kind == 'C')
                xForm.DataSources.UserDataSources.Add(uds, SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
            else
                xForm.DataSources.UserDataSources.Add(uds, SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 254);

            SAPbouiCOM.Item it = xForm.Items.Item(f.Item);
                
            if (f.Kind == 'C')
            {
                SAPbouiCOM.CheckBox oCheckBox = (SAPbouiCOM.CheckBox)it.Specific;
                oCheckBox.ValOn = "Y";
                oCheckBox.ValOff = "N";
                oCheckBox.DataBind.SetBound(true, "", uds);

                it.TextStyle = 1; //글자를 볼드 처리하여 눈에 뛰게 한다. 
            }
            else
            {
                ((SAPbouiCOM.EditText)it.Specific).DataBind.SetBound(true, "", uds);
            }

            it.AffectsFormMode = false;      // 값 변경 시 폼 모드(Update) 전환 방지 → 닫을 때 저장 확인창 방지
        }
    }



    //************************************************************************************************************
    // [ZMC202B] 화면 아이템 <-> 메인 그리드 컬럼 <-> UDO(UDF) 매핑
    //   Item     : ZMC202B.srf 의 아이템 UID
    //   GridCols : 메인 GRD1 컬럼명 후보 (앞에서부터 존재하는 첫 컬럼 사용)  → 팝업 표시값
    //   Udf      : [@ZMC202] UDF 필드명 (U_ 포함)                            → 등록 시 Update 대상
    //   Kind     : T=문자, D=일자(yyyyMMdd), M=긴문자(진행상황)
    //   실제로 Update 되는 항목은 화면에서 Enabled=true 인 박스만 (ZMCB_DataSave 에서 런타임 체크)
    //************************************************************************************************************
    private class BFld
    {
        public string Item;
        public string[] GridCols;
        public string Udf;
        public char Kind;

        public BFld(string item, string[] gridCols, string udf, char kind)
        {
            Item = item; GridCols = gridCols; Udf = udf; Kind = kind;
        }
    }

    private static readonly BFld[] B_FIELDS = new BFld[]
    {
        // ---- 조회 전용 (srf 에서 enabled=0) ----
        new BFld("U_PDOCNUM",   new string[] { "문서번호" },                 "U_PDOCNUM",  'T'),
        new BFld("U_PDOCENTY",  new string[] { "전표번호" },                 "U_PDOCENTY", 'T'),
        new BFld("U_CARDCODE",  new string[] { "거래처코드", "고객" },       "U_CARDCODE", 'T'),
        new BFld("U_CARDNAME",  new string[] { "거래처명", "고객명" },       "U_CARDNAME", 'T'),
        new BFld("U_PRJCODE",   new string[] { "프로젝트번호" },             "U_PRJCODE",  'T'),
        new BFld("U_SHYIDNM",   new string[] { "조선소명", "조선소" },       "U_SHYIDNM",  'T'),
        new BFld("U_SHMBER",    new string[] { "호선번호" },                 "U_SHMBER",   'T'),
        new BFld("U_BASHNM",    new string[] { "대표호선번호" },             "U_BASHNM",   'T'),
        new BFld("U_SHIPTYPE",  new string[] { "선박유형" },                 "U_SHIPTYPE", 'T'),
        new BFld("U_DESIGNER",  new string[] { "담당자", "담당자명" },       "U_DESIGNER", 'T'),
        new BFld("U_ORDITEM",   new string[] { "주문품목" },                 "U_ORDITEM",  'T'),
        new BFld("U_SALEDATE",  new string[] { "영업접수", "영업접수일" },   "U_SALEDATE", 'D'),
        new BFld("U_YARDPLN",   new string[] { "YARD PLAN" },                "U_YARDPLN",  'D'),
        new BFld("U_DLVFRST",   new string[] { "납기일(최초)" },             "U_DLVFRST",  'D'),
        new BFld("U_DESAPPDT",  new string[] { "승인일자" },                 "U_DESAPPDT", 'D'),

        // ---- 입력 가능 (srf 에서 enabled=1) ----
        new BFld("U_YARDSND", new string[] { "YARD 송부" },                 "U_YARDSND",    'D'),
        new BFld("U_YARDAGR", new string[] { "YARD 승인" },                 "U_YARDAGR",    'D'),
        new BFld("U_CLASSSND", new string[] { "CLASS 송부" },               "U_CLASSSND",   'M'),
        new BFld("U_CLASSAGR", new string[] { "CLASS 승인" },               "U_CLASSAGR",   'M'),
        new BFld("U_CDLVDATE", new string[] { "납기일(변경)" },             "U_CDLVDATE",   'D'),
        new BFld("U_PLANSTT1", new string[] { "진행상황1" },                "U_PLANSTT1",   'M'),
        new BFld("U_PLANSTT2", new string[] { "진행상황2" },                "U_PLANSTT2",   'M'),
        new BFld("U_PLANSTT3", new string[] { "진행상황3" },                "U_PLANSTT3",   'M'),
        new BFld("U_DOCSTATS", new string[] { "공정관리" },                 "U_DOCSTATS",   'C'),
        new BFld("CANCELED",   new string[] { "문서취소" },                 "CANCELED",     'C'),

    };

    // true : 팝업 오픈 시점 대비 값이 바뀐 박스만 Update / false : Enabled 박스 전체 Update
    private const bool UPDATE_CHANGED_ONLY = true;



    //************************************************************************************************************
    // 메인 그리드 선택 라인 → ZMC202B 박스 표시
    //************************************************************************************************************
    private void LoadSelectedToB(int selRow)
    {
        SAPbouiCOM.DataTable dt = oGRD1.DataTable;

        foreach (BFld f in B_FIELDS)
        {
            string v = GridVal(dt, f.GridCols, selRow);
            if (f.Kind == 'D') v = NormDate(v);
            if (f.Kind == 'C')  //공정관료 완료 상태의 문서를 선택한 경우 Y 상태값을 체크 박스에 전달 한다. 
            {
                if (v == "공정완료")
                {
                    v = "Y";
                }
                else
                {
                    v = "N";
                }
            }
            bForm.DataSources.UserDataSources.Item(f.Item).ValueEx = v;
            bOrig[f.Item] = v;
        }
    }


    //************************************************************************************************************
    // 일자 문자열 → yyyyMMdd (변환 불가 시 "")
    //************************************************************************************************************
    private static string NormDate(string s)
    {
        if (string.IsNullOrEmpty(s)) return "";

        DateTime d;
        if (s.Length == 8 && DateTime.TryParseExact(s, "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture,
                                                    System.Globalization.DateTimeStyles.None, out d))
            return s;

        if (DateTime.TryParse(s.Replace('.', '-').Replace('/', '-'), out d))
            return d.Year < 1900 ? "" : d.ToString("yyyyMMdd");

        return "";
    }



    //************************************************************************************************************
    // DataTable 값 읽기 : 컬럼명 후보 중 존재하는 첫 컬럼 값 (없으면 "")
    //************************************************************************************************************
    private static string GridVal(SAPbouiCOM.DataTable dt, string[] cols, int row)
    {
        foreach (string c in cols)
        {
            try
            {
                object v = dt.GetValue(c, row);
                if (v == null) return "";
                if (v is DateTime)
                {
                    DateTime d = (DateTime)v;
                    return d.Year < 1900 ? "" : d.ToString("yyyyMMdd");
                }
                return Convert.ToString(v).Trim();
            }
            catch
            {
            }
        }
        return "";
    }



    //************************************************************************************************************
    // 선택 라인 → [@ZMC202] Code 찾기 (못 찾으면 "" + 상태바 안내)
    //   1) 그리드에 Code 컬럼이 있으면 그대로 사용
    //   2) 없으면 전표번호(U_PDOCENTY) 로 조회, 2건 이상이면 문서구분(U_DOCGBN) 으로 좁힘
    //************************************************************************************************************
    private string ResolveUdoCode(int selRow)
    {
        SAPbobsCOM.Recordset rs = null;

        try
        {
            SAPbouiCOM.DataTable dt = oGRD1.DataTable;

            string code = GridVal(dt, new string[] { "Code", "코드", "관리번호" }, selRow);
            if (code != "") return code;

            string sEntry = GridVal(dt, new string[] { "전표번호" }, selRow);
            int pDocEntry;
            if (!int.TryParse(sEntry, out pDocEntry))
            {
                ZSBO.SBO_Application.StatusBar.SetText("선택한 라인의 전표번호를 확인할 수 없습니다. (" + sEntry + ")",
                    SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                return "";
            }

            rs = (SAPbobsCOM.Recordset)ZSBO.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            string q = "SELECT Code FROM [@ZMC202] WHERE CANCELED = 'N' AND U_PDOCENTY = " + pDocEntry.ToString();
            rs.DoQuery(q);

            if (rs.RecordCount == 1) return rs.Fields.Item(0).Value.ToString().Trim();

            if (rs.RecordCount == 0)
            {
                ZSBO.SBO_Application.StatusBar.SetText("공정관리([@ZMC202])에 등록되지 않은 전표입니다. (전표번호 " + sEntry + ")",
                    SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                return "";
            }

            // 동일 전표번호가 2건 이상 (초안/오더 구분 필요)
            string sGbn = GridVal(dt, new string[] { "문서구분", "문서", "DOCGBN" }, selRow);
            if (sGbn != "")
            {
                rs.DoQuery(q + " AND ISNULL(U_DOCGBN,'') = '" + sGbn.Replace("'", "''") + "'");
                if (rs.RecordCount == 1) return rs.Fields.Item(0).Value.ToString().Trim();
            }

            ZSBO.SBO_Application.StatusBar.SetText("동일 전표번호가 2건 이상 등록되어 있습니다. 메인 조회(ZTMSP_ZMC202) 결과에 Code 컬럼을 추가해 주세요.",
                SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
            return "";
        }
        catch (Exception ex)
        {
            ZSBO.SBO_Application.StatusBar.SetText("ResolveUdoCode_Error : " + ex.Message,
                SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            return "";
        }
        finally
        {
            if (rs != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(rs);
        }
    }



    //************************************************************************************************************
    // [등록] : ZMC202B 에서 Enabled=true 인 박스 값을 UDO([@ZMC202]) 에 Update
    //************************************************************************************************************
    private bool ZMCB_DataSave()
    {
        SAPbobsCOM.Recordset rs = null;
        SAPbobsCOM.GeneralService svc = null;
        bool inDI = false;
        bool closed = false;
        bool doRefresh = false;

        SAPbobsCOM.GeneralDataParams gp = null;
        SAPbobsCOM.GeneralData gd = null;

        try
        {
            bSaving = true;
            bForm.Freeze(true);

            if (bCode == "")
                throw new Exception("수정 대상 공정관리 Code 를 알 수 없습니다. 팝업을 닫고 다시 여십시오.");

            // 1) Enabled=true 박스 중 Update 대상 수집
            List<BFld> targets = new List<BFld>();
            List<string> vals = new List<string>();

            foreach (BFld f in B_FIELDS)
            {
                if (f.Udf == null) continue;
                if (!bForm.Items.Item(f.Item).Enabled) continue;      // enabled = true 인 박스만

                string v = bForm.DataSources.UserDataSources.Item(f.Item).ValueEx.Trim(); //"UD_" +

                string o;
                if (!bOrig.TryGetValue(f.Item, out o)) o = "";
                if (UPDATE_CHANGED_ONLY && v == o) continue;           // 변경 없는 박스 제외

                targets.Add(f);
                if (f.Kind == 'D' && v == "")   //날짜등록된 상태에서 지우면서 update 하는 경우 null 처리 한다. 
                {
                    vals.Add(null);
                }
                else
                {
                    vals.Add(v);
                }
            }

            if (targets.Count == 0)
            {
                ZSBO.SBO_Application.StatusBar.SetText("변경된 내용이 없습니다.",
                    SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                return false;
            }

            rs = (SAPbobsCOM.Recordset)ZSBO.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            // 2) UDF 존재/타입 확인 (CUFD.AliasID 는 U_ 제외 이름)
            Dictionary<string, string> udfType = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            rs.DoQuery("SELECT AliasID, TypeID FROM CUFD WHERE TableID = '@ZMC202'");
            while (rs.EoF == false)
            {
                udfType["U_" + rs.Fields.Item(0).Value.ToString().Trim()] = rs.Fields.Item(1).Value.ToString().Trim();
                rs.MoveNext();
            }

            string missing = "";
            foreach (BFld f in targets)
            {
                if (f.Udf != "CANCELED")
                {
                    if (!udfType.ContainsKey(f.Udf)) missing += (missing == "" ? "" : ", ") + f.Udf;
                }
            }
            if (missing != "")
                throw new Exception("[@ZMC202] 에 없는 UDF 입니다 : " + missing + "  (B_FIELDS 의 Udf 이름 확인)");

            // 3) UDO 코드 조회 (ZMCLIST_DataSave 와 동일)
            rs.DoQuery("SELECT TOP 1 Code FROM OUDO WHERE TableName IN ('ZMC202', '@ZMC202')");
            if (rs.RecordCount == 0)
                throw new Exception("[@ZMC202] 테이블에 연결된 UDO가 등록되어 있지 않습니다. (OUDO 확인)");
            string udoCode = rs.Fields.Item(0).Value.ToString().Trim();

            // 4) UDO 조회 → 값 설정 → Update
            inDI = true;
            svc = ZSBO.oCompany.GetCompanyService().GetGeneralService(udoCode);

            gp = (SAPbobsCOM.GeneralDataParams)svc.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams);
            gp.SetProperty("Code", bCode);

            gd = svc.GetByParams(gp);

            string chk_canceled = "N";

            for (int i = 0; i < targets.Count; i++)
            {
                inDI = false;     // 값 변환 오류는 DI 오류가 아님
                if (targets[i].Udf != "CANCELED")
                {
                    SetUdoValue(gd, targets[i].Udf, udfType[targets[i].Udf], vals[i]);
                }
                else
                {
                    if (vals[i] == "Y")
                    {
                        chk_canceled = "Y";
                    }

                }
            }

            inDI = true;

            if (chk_canceled == "N")  //문서를 취소 하는 경우는 udo 필드를 사용하지 않으므로 
            {
                svc.Update(gd);
            }
            else
            {
                svc.Cancel(gp);   //취소 기능 사용하려면 오브젝트 등록마법사 진행사 취소 기능을 체크해 주고 sap 재시작해야 함.
            }


            inDI = false;

            // UDO COM 객체를 폼 닫기 전에 먼저 해제
            if (gd != null) { System.Runtime.InteropServices.Marshal.ReleaseComObject(gd); gd = null; }
            if (gp != null) { System.Runtime.InteropServices.Marshal.ReleaseComObject(gp); gp = null; }
            if (svc != null) { System.Runtime.InteropServices.Marshal.ReleaseComObject(svc); svc = null; }


            // 5) 팝업 닫고 메인 재조회
            bForm.Freeze(false);
            bForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
            try { bForm.Close(); } catch { }   // 이미 닫힌 경우 방어
            closed = true;
            bModal = false;

            doRefresh = true; //등록작업했다는 구분값, finally 에서 사용 

            //LoadData("S");

            ZSBO.SBO_Application.StatusBar.SetText(chk_canceled == "Y"?"공정단계 문서를 취소하였습니다.":"공정단계 정보를 등록(수정)하였습니다. (" + targets.Count.ToString() + "개 항목)",
                SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);

            return true;
        }
        catch (Exception ex)
        {
            string detail = ex.Message;

            if (inDI)
            {
                int errCode; string errMsg;
                ZSBO.oCompany.GetLastError(out errCode, out errMsg);
                if (errCode != 0) detail = "[" + errCode.ToString() + "] " + errMsg;
            }

            ZSBO.SBO_Application.StatusBar.SetText("공정단계 등록 오류 : " + detail,
                SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            return false;
        }
        finally
        {

            if (rs != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(rs);
            if (gd != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(gd);
            if (gp != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(gp);
            if (rs != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(rs);

            if (!closed)
            {
                try { bForm.Freeze(false); }
                catch { }
            }



            bSaving = false;

            // 팝업이 완전히 닫힌 뒤, 메인 폼을 UID로 명시해서 재조회
            if (doRefresh)
            {
                try
                {
                    SAPbouiCOM.Form mainForm = ZSBO.SBO_Application.Forms.Item(mainFormUID);
                    mainForm.Select();

                    LoadData("S");   //메인폼 데이터 재 로드 
                }
                catch (Exception ex)
                {
                    ZSBO.SBO_Application.StatusBar.SetText("재조회 오류 : " + ex.Message,
                        SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                }
            }

        }
    }

    //************************************************************************************************************
    // UDF 타입(CUFD.TypeID)에 맞게 변환하여 GeneralData 에 설정
    //************************************************************************************************************
    private void SetUdoValue(SAPbobsCOM.GeneralData gd, string udf, string typeId, string v)
    {
        switch (typeId)
        {
            case "D":   // 일자

                if (v == null)           // 일자 삭제 → null 로 갱신
                {
                    gd.SetProperty(udf, "");
                    break;
                }

                if (v == "") return;     // 빈 일자는 기존 값 유지
                DateTime d;
                if (!DateTime.TryParseExact(v, "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture,
                                            System.Globalization.DateTimeStyles.None, out d))
                    throw new Exception(udf + " : 일자 형식 오류 (" + v + ")");
                gd.SetProperty(udf, d);
                break;

            case "N":   // 정수
                int n;
                if (v == "") n = 0;
                else if (!int.TryParse(v, out n)) throw new Exception(udf + " : 숫자 형식 오류 (" + v + ")");
                gd.SetProperty(udf, n);
                break;

            case "B":   // 소수(금액/수량/비율 등)
                double dbl;
                if (v == "") dbl = 0;
                else if (!double.TryParse(v, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out dbl))
                    throw new Exception(udf + " : 숫자 형식 오류 (" + v + ")");
                gd.SetProperty(udf, dbl);
                break;

            default:    // A(문자) / M(메모)

                gd.SetProperty(udf, v == null ? "" : v);
                break;
        }
    }



    //************************************************************************************************************
    // [ZMC202A] 오픈 시점의 폼/아이템 크기 저장
    //************************************************************************************************************
    private void InitPopupLayout(SAPbouiCOM.Form f)
    {
        aOrig.Clear();

        for (int i = 0; i < f.Items.Count; i++)
        {
            SAPbouiCOM.Item it = f.Items.Item(i);
            RectI r = new RectI();
            r.L = it.Left; r.T = it.Top; r.W = it.Width; r.H = it.Height;
            aOrig[it.UniqueID] = r;
        }

        aOrigCW = f.ClientWidth;
        aOrigCH = f.ClientHeight;
        aLayoutInit = true;
    }


    //************************************************************************************************************
    // [ZMC202A] 폼 크기 변경(최대화/복원/드래그) 시 그리드·매트릭스 크기 조정
    //   - 가로 : 증가분만큼 GRD1, MAT1 폭 확대
    //   - 세로 : 증가분을 GRD1 : MAT1 원래 높이 비율로 배분
    //   - 그리드 아래 아이템은 그리드 증가분만큼, 매트릭스 아래 아이템은 전체 증가분만큼 아래로 이동
    //************************************************************************************************************
    private void ApplyPopupLayout(SAPbouiCOM.Form f)
    {
        if (!aLayoutInit || f == null) return;
        if (!aOrig.ContainsKey("GRD1") || !aOrig.ContainsKey("MAT1")) return;

        // 원래 크기보다 작아지면 원래 크기로 고정 (음수 방지)
        int dW = Math.Max(0, f.ClientWidth - aOrigCW);
        int dH = Math.Max(0, f.ClientHeight - aOrigCH);

        RectI g = aOrig["GRD1"];
        RectI m = aOrig["MAT1"];

        int gridExtra = (g.H + m.H) > 0 ? (int)Math.Round(dH * (double)g.H / (g.H + m.H)) : 0;
        int matExtra = dH - gridExtra;

        int gridBottom = g.T + g.H;
        int matBottom = m.T + m.H;

        try
        {
            f.Freeze(true);

            foreach (KeyValuePair<string, RectI> kv in aOrig)
            {
                SAPbouiCOM.Item it = f.Items.Item(kv.Key);
                RectI o = kv.Value;

                if (kv.Key == "GRD1")
                {
                    it.Left = o.L; it.Top = o.T;
                    it.Width = o.W + dW;
                    it.Height = o.H + gridExtra;
                }
                else if (kv.Key == "MAT1")
                {
                    it.Left = o.L;
                    it.Top = o.T + gridExtra;           // 그리드가 커진 만큼 아래로
                    it.Width = o.W + dW;
                    it.Height = o.H + matExtra;
                }
                else
                {
                    // 그 외 아이템(버튼/라벨 등): 항상 원본 기준으로 위치를 다시 지정
                    int shift = 0;
                    if (o.T >= matBottom - 1) shift = dH;          // 매트릭스 아래
                    else if (o.T >= gridBottom - 1) shift = gridExtra;   // 그리드 아래, 매트릭스 위

                    it.Left = o.L;
                    it.Top = o.T + shift;     // shift 가 0 이어도 반드시 대입 (복원 시 원위치)
                    it.Width = o.W;
                    it.Height = o.H;
                }
            }

            // 컬럼 폭 재조정
            ((SAPbouiCOM.Grid)f.Items.Item("GRD1").Specific).AutoResizeColumns();
            ((SAPbouiCOM.Matrix)f.Items.Item("MAT1").Specific).AutoResizeColumns();
        }
        finally
        {
            f.Freeze(false);
        }
    }

}
