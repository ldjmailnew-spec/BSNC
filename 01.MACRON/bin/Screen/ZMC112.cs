//***********************************************************************************************************************************************
//  File             : ZMC112.cs
//  Module           : 품목대체등록(배치시리얼)
//  Desc             : 재고관리 > 품목대체
//  Creator          : LEE KEE CHAN
//  Create Date      : 2026.07.24
//  Memo             : 2026.09.07.LKC. 품명 재배치(너비 조정)로 인한 수정. 추가 후 CFL 버튼 비활성화 방지 조치
//                                     배치품목은 TAG번호가 입력되지 않게 체크. ZMC112A 타이틀 품번대체(배치,시리얼) 선택으로 변경
//  Procedure        : [ZTMSP_ZMC112A], [ZTMSP_ZMC112_BIN]
//  Copyright (c) BSNC
//***********************************************************************************************************************************************
using System;
using System.Windows.Forms;
using System.Threading;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Collections.Generic;

class ZMC112    
{
    //Main Form(ZMC112)
    private static SAPbouiCOM.Form oForm;
    private static SAPbouiCOM.Matrix oMAT1;
    private static SAPbouiCOM.Matrix oMAT2;
    private static SAPbouiCOM.DBDataSource oDBDSH;
    private static SAPbouiCOM.DBDataSource oDBDSL;
    private static SAPbouiCOM.DBDataSource oDBDSM;

    private static string m_ActiveMatrixUID = "MAT1"; // 행 추가를 위해 가장 최근 활성 상태인 Matrix UID 저장
    private static bool m_BatchUpdated = false; // CFL 선택 시 일괄 변경 중복 실행 방지 플래그

    //Modal Form(ZMC112A)
    private static SAPbouiCOM.Form aForm;
    private static SAPbouiCOM.DBDataSource ZDBDSH;
    private static SAPbouiCOM.DBDataSource ZDBDSL;
    private static SAPbouiCOM.DBDataSource ZDBDSM;
    private static SAPbouiCOM.Matrix ZMAT1;
    private static SAPbouiCOM.Matrix ZMAT2;
    private static SAPbouiCOM.Grid oGrd1;

    private bool aModal = false;

    //===========================================================================================================
    [DllImport("shell32.dll")]
    private static extern int ShellExecute(int hwnd, String lpOperation, String lpFile, String lpParameters, String lpDirectory, int nShowCmd);

    //************************************************************************************************************
    // ItemEventHander
    //************************************************************************************************************
    public void SBO_Application_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
    {
        SAPbouiCOM.Matrix oMatrix;
        string oStr = "", oCUID = "";
        BubbleEvent = true;
        SAPbouiCOM.IChooseFromListEvent oCFL = null;

        if (oForm.UniqueID != FormUID)
        {
            oForm = ZSBO.SBO_Application.Forms.Item(FormUID);
            oMAT1 = (SAPbouiCOM.Matrix)oForm.Items.Item("MAT1").Specific;
            oMAT2 = (SAPbouiCOM.Matrix)oForm.Items.Item("MAT2").Specific;
            oDBDSH = oForm.DataSources.DBDataSources.Item("@ZMC112H");
            oDBDSL = oForm.DataSources.DBDataSources.Item("@ZMC112L");
            oDBDSM = oForm.DataSources.DBDataSources.Item("@ZMC112M");
        }

        ///////////////////////////////////////////////////////////////////////////////////////
        /// [참조 Form] : Modal Form 처리                                           ///
        //////////////////////////////////////////////////////////////////////////////////////
        // Don't let the user to move to other forms
        SAPbouiCOM.BoEventTypes EventEnum = 0;
        // For FORM_CLOSE check variable
        EventEnum = pVal.EventType;

        // If the modal form UID is not equal....
        if (aModal & pVal.FormTypeEx != "ZMC112A")
        {
            aForm.Select(); //  Select the modal form
            //aForm.Freeze(false);
            BubbleEvent = false;
            return;
        }

        if (ZSBO.SBO_Application.Forms.Item(FormUID).TypeEx == "ZMC112A" & (EventEnum == SAPbouiCOM.BoEventTypes.et_FORM_CLOSE) & aModal)
            aModal = false;
        //////////////////////////////////////////////////////////////////////////////////////

        try
        {   /////////////////////////////////////////////////////////////////////////////////
            /// 참조 Form 이벤트                                                   ///
            /////////////////////////////////////////////////////////////////////////////////
            if (pVal.FormTypeEx == "ZMC112A")
            {
                ZMSLIST_ItemEventA(ref pVal, out BubbleEvent);

                return;
            }
            //활성 상태 Matrix 저장
            if (pVal.ItemUID != null && pVal.ItemUID.StartsWith("MAT") &&
               (pVal.EventType == SAPbouiCOM.BoEventTypes.et_GOT_FOCUS || pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK))
            {
                m_ActiveMatrixUID = pVal.ItemUID;
            }

            ////////////////////////////////////////////////////////////////////////////////
            switch (pVal.EventType)
            {
                //==============================================================================
                case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
                //==============================================================================
                    if (pVal.Before_Action & pVal.ItemUID == "1" & (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE | oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE))
                    {
                        if (ZSBO.SBO_Application.MessageBox("해당 내역을 저장하시겠습니까?", 1, "예", "아니오", "") == 2) { BubbleEvent = false; return; }
                        if (ChkInData() == false)
                        {
                            AddRow(oMAT1, oDBDSL, "U_ITEMCODE", 0);
                            //AddRow(oMAT2, oDBDSM, "U_ITEMCODE", 0);
                            BubbleEvent = false; return;
                        }
                        else
                        {
                            if (MakeDI() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }
                        }
                    }
                    else if (pVal.Before_Action == false & pVal.ItemUID == "1")
                    {
                        if (!pVal.Action_Success & (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE | oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE))
                        {
                            AddRow(oMAT1, oDBDSL, "U_ITEMCODE", 0);
                            //AddRow(oMAT2, oDBDSM, "U_ITEMCODE", 0);
                            SetFocusByFormMode();
                        }
                        else if (pVal.Action_Success & oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            //AddRow(oMAT1, oDBDSL, "U_ITEMCODE", 0);
                            SetFormBase();
                            SetItemEnabled();
                            SetToolBarEnabled();
                            SetFocusByFormMode();
                            // 틀고정
                            oMAT1.CommonSetting.FixedColumnsCount = 2;
                            oMAT2.CommonSetting.FixedColumnsCount = 2;
                        }
                        else if (pVal.Action_Success & oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                            AddRow(oMAT1, oDBDSL, "U_ITEMCODE", 0);
                            //AddRow(oMAT2, oDBDSM, "U_ITEMCODE", 0);
                            SetItemEnabled();
                            SetToolBarEnabled();

                            SetFocusByFormMode();
                            oMAT1.AutoResizeColumns();
                            oMAT2.AutoResizeColumns();
                            // 틀고정
                            oMAT1.CommonSetting.FixedColumnsCount = 2;
                            oMAT2.CommonSetting.FixedColumnsCount = 2;
                        }
                        else if (pVal.BeforeAction == false & oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                        {
                            SetItemEnabled();
                            SetToolBarEnabled();
                            SetFocusByFormMode();
                            // 틀고정
                            oMAT1.CommonSetting.FixedColumnsCount = 2;
                            oMAT2.CommonSetting.FixedColumnsCount = 2;
                        }
                        else if (pVal.BeforeAction == false & oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                            SetItemEnabled();
                            SetToolBarEnabled();
                            SetFocusByFormMode();
                            // 틀고정
                            oMAT1.CommonSetting.FixedColumnsCount = 2;
                            oMAT2.CommonSetting.FixedColumnsCount = 2;
                        }
                    }
                    ////Modal Form(요청 참조)
                    //else if (pVal.BeforeAction == false & (pVal.FormMode != (int)SAPbouiCOM.BoFormMode.fm_FIND_MODE) & pVal.ItemUID == "BTN_REF")
                    //{
                    //    if (ZMSLIST_ChkInDataA())
                    //    {
                    //        for (int i = 0; i < ZSBO.SBO_Application.Forms.Count; i++)
                    //        {
                    //            if (ZSBO.SBO_Application.Forms.Item(i).TypeEx == "ZMC112A")
                    //            {
                    //                ZSBO.SBO_Application.Forms.Item(i).Select();
                    //                //BubbleEvent = false;
                    //                //return;
                    //            }
                    //        }

                    //        if (ZSBO.SBO_Application.MessageBox("현재 내역이 모두 삭제됩니다. 계속하시겠습니까?", 1, "예", "취소", "") == 2) { BubbleEvent = false; return; }
                    //        ZMSLIST_LoadFormA(ref oDBDSH, ref oDBDSL, ref oMAT1);
                    //    }
                    //}
                    //Modal Form(배치/시리얼 참조)
                    else if (pVal.BeforeAction == false & (pVal.FormMode != (int)SAPbouiCOM.BoFormMode.fm_FIND_MODE) & pVal.ItemUID == "BTN_BAT")
                    {
                        if (ZMSLIST_ChkInDataA())
                        {
                            for (int i = 0; i < ZSBO.SBO_Application.Forms.Count; i++)
                            {
                                if (ZSBO.SBO_Application.Forms.Item(i).TypeEx == "ZMC112A")
                                {
                                    ZSBO.SBO_Application.Forms.Item(i).Select();
                                    //BubbleEvent = false;
                                    //return;
                                }
                            }

                            if (ZSBO.SBO_Application.MessageBox("현재 내역이 모두 삭제됩니다. 계속하시겠습니까?", 1, "예", "취소", "") == 2) { BubbleEvent = false; return; }
                            ZMSLIST_LoadFormA(ref oDBDSH, ref oDBDSL, ref oMAT1);
                        }
                    }
                    break;

                //==============================================================================
                case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:
                //==============================================================================
                    if (pVal.Before_Action & pVal.Modifiers == SAPbouiCOM.BoModifiersEnum.mt_None & pVal.ItemUID.Substring(0, 3) == "MAT" & pVal.Row > 0)
                    {
                        oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item(pVal.ItemUID).Specific;
                        if (pVal.CharPressed == 38)         // 방향키-좌(37),위(38),우(39),아래(40)
                        {
                            if (pVal.Row == 1 & oMatrix.RowCount == 1) { }
                            else if (pVal.Row == 1 & oMatrix.RowCount > 1) { oMatrix.Columns.Item(pVal.ColUID).Cells.Item(oMatrix.RowCount).Click(SAPbouiCOM.BoCellClickType.ct_Regular, 1); }
                            else { oMatrix.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row - 1).Click(SAPbouiCOM.BoCellClickType.ct_Regular, 1); }
                            BubbleEvent = false;
                        }
                        else if (pVal.CharPressed == 40)     // 방향키-좌(37),위(38),우(39),아래(40)
                        {
                            if (pVal.Row == 1 & oMatrix.RowCount == 1) { }
                            else if (pVal.Row == oMatrix.RowCount) { oMatrix.Columns.Item(pVal.ColUID).Cells.Item(1).Click(SAPbouiCOM.BoCellClickType.ct_Regular, 1); }
                            else { oMatrix.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row + 1).Click(SAPbouiCOM.BoCellClickType.ct_Regular, 1); }
                            BubbleEvent = false;
                        }
                    }
                    break;

                //==============================================================================
                case SAPbouiCOM.BoEventTypes.et_CLICK:
                //==============================================================================
                    //MAT1
                    if (pVal.BeforeAction == false & pVal.ItemUID == "MAT1" && pVal.Row > 0) // && pVal.ColUID == "LineId")
                    {
                        if (oMAT1.IsRowSelected(pVal.Row) == false)
                        {
                            oMAT1.SelectRow(pVal.Row, true, false);
                        }
                        else if (oMAT1.IsRowSelected(pVal.Row) == true)
                        {
                            oMAT1.SelectRow(pVal.Row, false, false);
                            //BubbleEvent = false;
                        }
                    }
                    //MAT2
                    if (pVal.BeforeAction == false & pVal.ItemUID == "MAT2" && pVal.Row > 0)
                    {
                        if (oMAT2.IsRowSelected(pVal.Row) == false)
                        {
                            oMAT2.SelectRow(pVal.Row, true, false);
                        }
                        else if (oMAT2.IsRowSelected(pVal.Row) == true)
                        {
                            oMAT2.SelectRow(pVal.Row, false, false);
                            //BubbleEvent = false;
                        }
                    }                    
                    break;

                //==============================================================================
                case SAPbouiCOM.BoEventTypes.et_VALIDATE:
                //==============================================================================
                    if (pVal.BeforeAction == false & pVal.ItemChanged & oForm.Items.Item(pVal.ItemUID).Type == SAPbouiCOM.BoFormItemTypes.it_MATRIX)
                    {
                        oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item(pVal.ItemUID).Specific;
                        oStr = ((SAPbouiCOM.EditText)oMatrix.GetCellSpecific(pVal.ColUID, pVal.Row)).String.Trim();
                        FlushToItemValue(pVal.ItemUID, pVal.ColUID, pVal.Row, oStr);
                    }
                    if (pVal.BeforeAction == false & pVal.ItemChanged & oForm.Items.Item(pVal.ItemUID).Type == SAPbouiCOM.BoFormItemTypes.it_EDIT)
                    {
                        oStr = ((SAPbouiCOM.EditText)oForm.Items.Item(pVal.ItemUID).Specific).String.Trim();
                        FlushToItemValue(pVal.ItemUID, pVal.ColUID, pVal.Row, oStr);
                    }
                    break;

                //========================================================================================================
                case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:
                //========================================================================================================
                    if (pVal.Before_Action == false & pVal.ItemChanged & oForm.Items.Item(pVal.ItemUID).Type == SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX)
                    {
                        //if (pVal.ItemUID == "WHSCODE")  2026.02.19.LKC. 헤더 출고창고 칼럼 관련 내용 주석 처리(라인으로 이동) 
                        //{
                        //    if (oMAT1.VisualRowCount <= 1) return;
                        //    if (ZSBO.SBO_Application.MessageBox("라인에 창고를 일괄 변경하시겠습니까?", 1, "예", "아니오", "") == 2) { BubbleEvent = false; return; }
                        //    oStr = "";
                        //    FlushToItemValue(pVal.ItemUID, pVal.ColUID, pVal.Row, oStr);
                        //}
                        //else
                        //{
                        oStr = "";
                        FlushToItemValue(pVal.ItemUID, pVal.ColUID, pVal.Row, oStr);
                        //}
                    }
                    //if (pVal.Before_Action == false & pVal.ItemChanged & oForm.Items.Item(pVal.ItemUID).Type == SAPbouiCOM.BoFormItemTypes.it_MATRIX)
                    //{
                    //    oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item(pVal.ItemUID).Specific;
                    //    oStr = ((SAPbouiCOM.ComboBox)oMatrix.GetCellSpecific(pVal.ColUID, pVal.Row)).Value.ToString().Trim();
                    //    FlushToItemValue(pVal.ItemUID, pVal.ColUID, pVal.Row, oStr);
                    //}
                    break;

                //==============================================================================
                case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:
                //==============================================================================
                    if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE) return;
                    oCFL = ((SAPbouiCOM.IChooseFromListEvent)(pVal));

                    if (pVal.Before_Action == false &&
                        (oCFL.ChooseFromListUID == "HITEMCD"  || oCFL.ChooseFromListUID == "HITEMNM"  ||
                         oCFL.ChooseFromListUID == "HTITEMCD" || oCFL.ChooseFromListUID == "HTITEMNM" ||
                         oCFL.ChooseFromListUID == "ITEMCODE" || oCFL.ChooseFromListUID == "ITEMNAME" ||
                         oCFL.ChooseFromListUID == "TITEMCD"  || oCFL.ChooseFromListUID == "TITEMNM"  ||
                         oCFL.ChooseFromListUID == "WHSCODE"  || oCFL.ChooseFromListUID == "TWHSCODE" ))
                    {
                        if (oCFL.SelectedObjects == null) return;
                        GetSelectCode(oCFL.SelectedObjects, pVal.Row, oCFL.ChooseFromListUID);                        
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                    }

                    if (oCFL.ChooseFromListUID == "BINCODE" || oCFL.ChooseFromListUID == "TBINCODE")
                    {
                        if (oCFL.ChooseFromListUID == "BINCODE")
                        {
                            oMAT1.FlushToDataSource();
                            SetChooseFromList("BINCODE", oDBDSL.GetValue("U_WHSCODE", pVal.Row - 1).Trim(), pVal.Row - 1);   // CFL Conditions 설정(입고 기준)
                        }
                        else if (oCFL.ChooseFromListUID == "TBINCODE")
                        {
                            oMAT2.FlushToDataSource();
                            SetChooseFromList("TBINCODE", oDBDSM.GetValue("U_WHSCODE", pVal.Row - 1).Trim(), pVal.Row - 1);   // CFL Conditions 설정(창고 기준)
                        }

                        if (oCFL.SelectedObjects == null) return;
                        GetSelectCode(oCFL.SelectedObjects, pVal.Row, oCFL.ChooseFromListUID); //CFL 값선택 처리
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                    }
                    break;

                //==============================================================================
                case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE:
                //==============================================================================
                    if (pVal.BeforeAction & (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE | oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE))
                    {
                        switch (ZSBO.SBO_Application.MessageBox("저장되지 않은 데이터가 있습니다. 종료하시겠습니까?", 1, "예", "취소", ""))
                        {
                            case 1: oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE; break;
                            case 2: BubbleEvent = false; break;
                        }
                    }
                    break;

                ////================================================================================================================
                //case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED: 미사용
                ////================================================================================================================
                //    if (pVal.BeforeAction & pVal.ColUID == "BSENTRY")
                //    {
                //        ZMC101 aZMC101 = new ZMC101();
                //        oStr = oDBDSH.GetValue("U_BSENTRY", 0);

                //        aZMC101.LoadForm("ZMC101", oStr);
                //        aZMC101 = null;
                //        BubbleEvent = false;
                //    }
                //    break;

                //========================================================================================================
                case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE:
                //========================================================================================================
                    if (pVal.BeforeAction == false)
                    {
                        try
                        {
                            if (oForm.Items.Count > 0)
                            {
                                oForm.Freeze(true);

                                SAPbouiCOM.Item oItemMAT1 = oForm.Items.Item("MAT1");
                                SAPbouiCOM.Item oItemMAT2 = oForm.Items.Item("MAT2");

                                int nFormW = oForm.ClientWidth;   // 현재 폼 Client 너비
                                int nFormH = oForm.ClientHeight;  // 현재 폼 Client 높이

                                int nGap       = 8;   // MAT1과 MAT2 사이 간격
                                int nLeftPad   = 6;   // 왼쪽 여백
                                int nRightPad  = 6;   // 오른쪽 여백
                                //int nMatTop    = 81;  // MAT 시작 top 위치 (고정)
                                int nMatTop = 96;     // MAT 시작 top 위치 (고정) 2026.09.07.LKC. 품번명 재배치로 인한 수정
                                //int nBottomH   = 112; // 하단 영역 높이 (top=430~542 = 112)
                                int nBottomH = 127; // 하단 영역 높이
                                int nBottomGap = 12;  // MAT 하단 ~ 하단 아이템 사이 간격

                                // MAT 너비 계산: 폼 너비에서 좌우 여백, 간격을 빼고 2등분
                                int nMatW = (nFormW - nLeftPad - nRightPad - nGap) / 2;
                                if (nMatW < 200) nMatW = 200; // 최소 너비 보장

                                // MAT 높이 계산: 폼 높이에서 상단, 하단 영역, 간격을 빼기
                                int nMatH = nFormH - nMatTop - nBottomGap - nBottomH;
                                if (nMatH < 100) nMatH = 100; // 최소 높이 보장

                                // MAT1 배치 (왼쪽)
                                oItemMAT1.Left   = nLeftPad;
                                oItemMAT1.Width  = nMatW;
                                oItemMAT1.Top    = nMatTop;
                                oItemMAT1.Height = nMatH;

                                // MAT2 배치 (오른쪽)
                                oItemMAT2.Left   = nLeftPad + nMatW + nGap;
                                oItemMAT2.Width  = nMatW;
                                oItemMAT2.Top    = nMatTop;
                                oItemMAT2.Height = nMatH;

                                // 하단 아이템 top 재배치 (MAT 아래로 이동)
                                int nBottomTop = nMatTop + nMatH + nBottomGap; // 하단 영역 시작 top
                                int nOrigBottomTop = 430; // SRF 기준 원래 하단 시작 top
                                int nDiffTop = nBottomTop - nOrigBottomTop;    // 이동량

                                // 하단 아이템들 절대 위치 재배치 (SRF 원본 top 기준 + 오프셋)
                                // { UID, 원본 SRF top }
                                string[,] bottomItems = {
                                    { "25",      "430" }, { "REMARK",  "430" },
                                    { "28",      "430" }, { "DINUM",   "430" }, { "30",      "430" },
                                    { "31",      "447" }, { "DINUM2",  "447" }, { "36",      "447" },
                                    { "32",      "464" }, { "DICNUM",  "464" }, { "34",      "464" },
                                    { "33",      "481" }, { "DICNUM2", "481" }, { "35",      "481" },
                                    { "BTN_REF", "488" },
                                    { "BTN_BAT", "516" }, { "1",       "516" }, { "2",       "516" }
                                };
                                for (int bi = 0; bi < bottomItems.GetLength(0); bi++)
                                {
                                    try
                                    {
                                        SAPbouiCOM.Item oBottomItem = oForm.Items.Item(bottomItems[bi, 0]);
                                        oBottomItem.Top = int.Parse(bottomItems[bi, 1]) + nDiffTop;
                                    }
                                    catch { }
                                }

                                oMAT1.AutoResizeColumns();
                                oMAT2.AutoResizeColumns();
                                //oMAT1.Columns.Item("ITEMNAME").Width = 200;
                                //oMAT2.Columns.Item("ITEMNAME").Width = 200;

                                oForm.Freeze(false);
                            }
                        }
                        catch (Exception ex)
                        {
                            oForm.Freeze(false);
                        }
                    }
                    break;
            }
        }
        //-------------------------------------------------------------------------------------------------------
        catch (Exception Error)
        {
            oForm.Freeze(false);
            ZSBO.SBO_Application.SetStatusBarMessage(Error.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        }
    }

    //************************************************************************************************************
    // [Get_SelectCode] : 선택리스트 이후 데이터 바로 가져오기
    //************************************************************************************************************
    private void GetSelectCode(SAPbouiCOM.DataTable oSelTable, int oRow, string oItemUID)
    {
        SAPbobsCOM.Recordset oRecordset = ((SAPbobsCOM.Recordset)ZSBO.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset));
        double oPrc, oAmt, oVat, oQty;
        int i, oSRow, oSize;
        string oQry, sVatGrp, oBinCode, oItemCode, oWhsCode;
        var oBinCodesList = new List<string>();
        string[] oBinCodes;

        try
        {
            switch (oItemUID)
            {
                case "HITEMCD": //헤더-품번
                case "HITEMNM":
                    oDBDSH.SetValue("U_ITEMCODE", 0, oSelTable.GetValue("ItemCode", 0).ToString().Trim());
                    oDBDSH.SetValue("U_ITEMNAME", 0, oSelTable.GetValue("ItemName", 0).ToString().Trim());
                    break;

                case "HTITEMCD": //헤더-대체품번
                case "HTITEMNM":
                    oDBDSH.SetValue("U_TITEMCD", 0, oSelTable.GetValue("ItemCode", 0).ToString().Trim());
                    oDBDSH.SetValue("U_TITEMNM", 0, oSelTable.GetValue("ItemName", 0).ToString().Trim());

                    // CFL 선택 후 MAT2(oDBDSM) U_ITEMCODE 일괄 변경
                    m_BatchUpdated = true; // FlushToItemValue에서 중복 실행 방지
                    if (ZSBO.SBO_Application.MessageBox("대체품번을 초기화 합니다.", 1, "예", "", "") == 1)
                    {
                        oDBDSM.Clear();
                        oMAT2.Clear();
                        AddRow(oMAT2, oDBDSM, "U_ITEMCODE", 0);
                        oDBDSM.Offset = oDBDSM.Size - 1;
                        oMAT2.FlushToDataSource();
                        oDBDSM.SetValue("U_ITEMCODE", oDBDSM.Size - 1, oDBDSH.GetValue("U_TITEMCD", 0).ToString().Trim());
                        oDBDSM.SetValue("U_ITEMNAME", oDBDSM.Size - 1, oDBDSH.GetValue("U_TITEMNM", 0).ToString().Trim());
                        oMAT2.LoadFromDataSource();
                    }
                    break;

                case "ITEMCODE": //oMAT1
                case "ITEMNAME":
                    oSRow = oRow;
                    oMAT1.FlushToDataSource();
                    for (i = 0; i <= oSelTable.Rows.Count - 1; i++)
                    {
                        if (i != 0) oDBDSL.InsertRecord(oRow - 1);
                        oDBDSL.Offset = oRow - 1;
                        oDBDSL.SetValue("U_ITEMCODE", oRow - 1, oSelTable.GetValue("ItemCode", i).ToString().Trim());
                        oDBDSL.SetValue("U_ITEMNAME", oRow - 1, oSelTable.GetValue("ItemName", i).ToString().Trim());
                        oDBDSL.SetValue("U_WHSCODE",  oRow - 1, "");
                        oDBDSL.SetValue("U_BINCODE",  oRow - 1, "");
                        oDBDSL.SetValue("U_ONHAND",   oRow - 1, "");
                        oDBDSL.SetValue("U_WONHAND",  oRow - 1, "");
                        oRow++;
                    }
                    if (oDBDSL.GetValue("U_ITEMCODE", oDBDSL.Size - 1).ToString() != "")
                        oDBDSL.InsertRecord(oDBDSL.Size);
                    for (i = 0; i < oDBDSL.Size; i++)
                    {
                        oDBDSL.Offset = i;
                        oDBDSL.SetValue("LineId", i, (i + 1).ToString());
                    }
                    oMAT1.LoadFromDataSource();
                    oMAT1.AutoResizeColumns();
                    oMAT1.Columns.Item(oItemUID).Cells.Item(oSRow).Click(SAPbouiCOM.BoCellClickType.ct_Regular, 1);
                    break;

                case "TITEMCD": //oMAT2
                case "TITEMNM":
                    oSRow = oRow;
                    oMAT2.FlushToDataSource();
                    for (i = 0; i <= oSelTable.Rows.Count - 1; i++)
                    {
                        if (i != 0) oDBDSM.InsertRecord(oRow - 1);
                        oDBDSM.Offset = oRow - 1;
                        oDBDSM.SetValue("U_ITEMCODE", oRow - 1, oSelTable.GetValue("ItemCode", i).ToString().Trim());
                        oDBDSM.SetValue("U_ITEMNAME", oRow - 1, oSelTable.GetValue("ItemName", i).ToString().Trim());
                        //oDBDSM.SetValue("U_ITEMNM01", oRow - 1, oSelTable.GetValue("FrgnName", i).ToString().Trim());
                        oDBDSM.SetValue("U_WHSCODE", oRow - 1, ""); //대체품번변경에 따른 창고/BIN/ 초기화
                        oDBDSM.SetValue("U_BINCODE", oRow - 1, "");
                        //oDBDSM.SetValue("U_ONHAND", oRow - 1, "");
                        oRow++;
                    }
                    if (oDBDSM.GetValue("U_ITEMCODE", oDBDSM.Size - 1).ToString() != "")
                        oDBDSM.InsertRecord(oDBDSM.Size);
                    for (i = 0; i < oDBDSM.Size; i++)
                    {
                        oDBDSM.Offset = i;
                        oDBDSM.SetValue("LineId", i, (i + 1).ToString());
                    }
                    oMAT2.LoadFromDataSource();
                    oMAT2.AutoResizeColumns();
                    if (oItemUID == "TITEMCD")
                        oMAT2.Columns.Item("ITEMCODE").Cells.Item(oSRow).Click(SAPbouiCOM.BoCellClickType.ct_Regular, 1);
                    else
                        oMAT2.Columns.Item("ITEMNAME").Cells.Item(oSRow).Click(SAPbouiCOM.BoCellClickType.ct_Regular, 1);
                    break;

                case "WHSCODE": //oMAT1
                    oMAT1.GetLineData(oRow);
                    oDBDSL.Offset = oRow - 1;
                    oDBDSL.SetValue("U_WHSCODE", oRow - 1, oSelTable.GetValue("WhsCode", 0).ToString().Trim());

                    oItemCode = oDBDSL.GetValue("U_ITEMCODE", oRow - 1).ToString().Trim(); //2026.03.11.LKC. 창고 변경시 BIN 로케이션 및 재고 설정
                    oWhsCode  = oDBDSL.GetValue("U_WHSCODE",  oRow - 1).ToString().Trim();
                    //oConds = (SAPbouiCOM.Conditions)ZSBO.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_Conditions);

                    oRecordset.DoQuery("EXEC [DBO].[ZTMSP_ZMC112_BIN] @TYPE     = 'TO' " //출고BIN : 최근 입고기준 TOP 1 Bin Location 2026.03.13.LKC.
                                                                + " , @ITEMCODE = '" + oItemCode + "', @WHSCODE  = '" + oWhsCode + "' ");
                    while (oRecordset.EoF == false)
                    {
                        //같은 값은 제외함
                        oBinCode = oRecordset.Fields.Item(0).Value.ToString();
                        if (!oBinCodesList.Contains(oBinCode)) oBinCodesList.Add(oBinCode); //중복 제외
                        oRecordset.MoveNext();
                    }

                    oBinCodes = oBinCodesList.ToArray();                    
                    if (oBinCodesList.Count == 1)
                    {
                        oBinCode = oBinCodes.GetValue(0).ToString().Trim();
                        //출고BIN
                        oDBDSL.SetValue("U_BINCODE", oRow - 1, oBinCode);
                        //재고계산
                        oDBDSL.SetValue("U_ONHAND",  oRow - 1, GetOnHand("B", oItemCode, oWhsCode, oBinCode).ToString()); //BIN재고
                        oDBDSL.SetValue("U_WONHAND", oRow - 1, GetOnHand("W", oItemCode, oWhsCode, oBinCode).ToString()); //창고재고 2026.04.24.LKC.
                    }
                    else
                    {
                        oDBDSL.SetValue("U_BINCODE", oRow - 1, "");
                        oDBDSL.SetValue("U_ONHAND",  oRow - 1, 0.ToString()); //BIN재고
                        oDBDSL.SetValue("U_WONHAND", oRow - 1, 0.ToString()); //창고재고 2026.04.24.LKC.
                    }
                    oMAT1.SetLineData(oRow);
                    break;

                case "TWHSCODE": // 창고변경시 BIN 초기화되게 수정, oMAT2
                    oMAT2.GetLineData(oRow);
                    oDBDSM.Offset = oRow - 1;
                    oDBDSM.SetValue("U_WHSCODE", oRow - 1, oSelTable.GetValue("WhsCode", 0).ToString().Trim());
                    oDBDSM.SetValue("U_BINCODE", oRow - 1, "");
                    oMAT2.SetLineData(oRow);
                    break;

                case "BINCODE": //출고BIN, oMAT1
                    oMAT1.GetLineData(oRow);
                    oDBDSL.Offset = oRow - 1;
                    if (oSelTable != null)
                    {
                        oBinCode = CheckIsNull(oSelTable.GetValue("BinCode", 0));
                        oDBDSL.SetValue("U_BINCODE", oRow - 1, oBinCode);

                        //BIN로케이션 재고 표시
                        oItemCode = oDBDSL.GetValue("U_ITEMCODE", oRow - 1).ToString();
                        oWhsCode  = oDBDSL.GetValue("U_WHSCODE",  oRow - 1).ToString();
                        oDBDSL.SetValue("U_ONHAND",  oRow - 1, GetOnHand("B", oItemCode, oWhsCode, oBinCode).ToString()); //BIN재고
                        oDBDSL.SetValue("U_WONHAND", oRow - 1, GetOnHand("W", oItemCode, oWhsCode, oBinCode).ToString()); //창고재고 2026.04.24.LKC.
                    }
                    else
                    {
                        oDBDSL.SetValue("U_ONHAND",  oRow - 1, 0.ToString()); //BIN재고
                        oDBDSL.SetValue("U_WONHAND", oRow - 1, 0.ToString()); //창고재고 2026.04.24.LKC.
                    }
                    oMAT1.SetLineData(oRow);
                    break;

                case "TBINCODE": //입고BIN, oMAT2
                    oMAT2.GetLineData(oRow);
                    oDBDSM.Offset = oRow - 1;
                    if (oSelTable != null)
                    {
                        oBinCode = CheckIsNull(oSelTable.GetValue("BinCode", 0));
                        oDBDSM.SetValue("U_BINCODE", oRow - 1, oBinCode);
                    }
                    oMAT2.SetLineData(oRow);
                    break;
            }

        }
        catch (Exception Error)
        {
            oForm.Freeze(false);
            ZSBO.SBO_Application.SetStatusBarMessage("GetSelectCode Error : " + Error.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        }
    }

    //************************************************************************************************************
    // MenuEventHander
    //************************************************************************************************************
    public void SBO_Application_MenuEvent(ref SAPbouiCOM.MenuEvent pVal, out bool BubbleEvent)
    {
        BubbleEvent = true;
        SAPbouiCOM.Matrix oMatrix = null;
        SAPbouiCOM.DBDataSource oDBDS = null;
        string oFldName = "", oColUID = "";
        string activeMatUID = "";
        int oRow = 0;

        try
        {
            ///////////////////////////////////////////////////////////////////////////////////////
            /// [참조 Form] : Modal Form Open 확인                                               ///
            //////////////////////////////////////////////////////////////////////////////////////
            if (aModal)
            {
                BubbleEvent = false;
            }
            //////////////////////////////////////////////////////////////////////////////////////

            if (pVal.BeforeAction & pVal.MenuUID == "ZMC112") { LoadForm("ZMC112", ""); return; }

            if (pVal.BeforeAction == true)
            {
                switch (pVal.MenuUID)
                {
                    case "1284": //취소
                        if (ZSBO.SBO_Application.MessageBox("취소하시겠습니까?", 1, "예", "아니오", "") == 2) { BubbleEvent = false; return; }
                        if (ChkCInData() == false) { BubbleEvent = false; return; }
                        if (CancelDI() == false)   { BubbleEvent = false; return; }
                        break;
                    ////////////case "1287": // 복제 
                    ////////////    if (ZSBO.SBO_Application.MessageBox("복제하시겠습니까?", 1, "예", "아니오", "") == 2) { BubbleEvent = false; return; }
                    ////////////    break;
                    case "1293":
                        //라인삭제
                        activeMatUID = m_ActiveMatrixUID;
                        if (activeMatUID == "MAT1")
                        {
                            oRow = ZSBO.ZBSFC.FGetIndexRow(oMAT1);
                            if (oRow > 0)
                            {
                                oRow = ZSBO.ZBSFC.FGetIndexRow(oMAT1);
                                DeleteRow(oMAT1, oDBDSL, "U_ITEMCODE");
                                //복제 비활성화
                                //oForm.EnableMenu("1287", false);
                                AddRow(oMAT1, oDBDSL, "U_ITEMCODE", 0);
                                if (oRow == 0 | oRow > oMAT1.VisualRowCount) { oRow = oMAT1.VisualRowCount; }
                                oMAT1.Columns.Item("ITEMCODE").Cells.Item(oRow).Click(SAPbouiCOM.BoCellClickType.ct_Regular, 1);
                                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                            }
                            else
                            {
                                ZSBO.SBO_Application.MessageBox("삭제할 행을 선택해 주세요!", 1, "확인", "", "");
                            }
                        }
                        else if (activeMatUID == "MAT2")
                        {
                            oRow = ZSBO.ZBSFC.FGetIndexRow(oMAT2);
                            if (oRow > 0)
                            {
                                oRow = ZSBO.ZBSFC.FGetIndexRow(oMAT2);
                                DeleteRow(oMAT2, oDBDSM, "U_ITEMCODE");
                                //복제 비활성화
                                //oForm.EnableMenu("1287", false);
                                //AddRow(oMAT2, oDBDSM, "U_ITEMCODE", 0);
                                if (oRow == 0 | oRow > oMAT2.VisualRowCount) { oRow = oMAT2.VisualRowCount; }
                                oMAT2.Columns.Item("ITEMCODE").Cells.Item(oRow).Click(SAPbouiCOM.BoCellClickType.ct_Regular, 1);
                                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                            }
                            else
                                ZSBO.SBO_Application.MessageBox("삭제할 행을 선택해 주세요!", 1, "확인", "", "");
                        }
                        BubbleEvent = false;
                        break;

                    case "1294":
                        activeMatUID = m_ActiveMatrixUID;
                        if (activeMatUID == "MAT1")
                        {
                            ZSBO.SBO_Application.StatusBar.SetText("대체품번만 행복제 가능합니다.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                            BubbleEvent = false;
                            return;
                        }
                        oRow = ZSBO.ZBSFC.FGetIndexRow(oMAT2);
                        if (oRow == 0)
                        {
                            ZSBO.SBO_Application.StatusBar.SetText("선택된 행이 없어 복제 할 수 없습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                            BubbleEvent = false;
                            return;
                        }
                        if (oRow > 0)
                        {
                            oForm.Freeze(true);
                            oMAT2.FlushToDataSource();
                            // 선택 행 아래에 새 레코드 삽입 (oRow: 0-based 기준으로 선택행 다음 위치)
                            oDBDSM.InsertRecord(oRow);
                            oDBDSM.Offset = oRow;
                            // 선택 행(oRow - 1) 데이터를 새 행(oRow)으로 복사
                            for (int oCOL_CNT = 1; oCOL_CNT < oMAT2.Columns.Count; oCOL_CNT++) // LineId 제외하기 위해 1부터 시작
                            {
                                string colUID = oMAT2.Columns.Item(oCOL_CNT).UniqueID.ToString();
                                oDBDSM.SetValue("U_" + colUID, oRow, oDBDSM.GetValue("U_" + colUID, oRow - 1).Trim());
                            }
                            // LineId 재정렬
                            for (int i = 0; i < oDBDSM.Size; i++)
                            {
                                oDBDSM.Offset = i;
                                oDBDSM.SetValue("LineId", i, (i + 1).ToString());
                            }
                            oMAT2.LoadFromDataSource();
                            oMAT2.Columns.Item("ITEMCODE").Cells.Item(oRow).Click(SAPbouiCOM.BoCellClickType.ct_Regular, 1);
                            //if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                            oForm.Freeze(false);
                        }
                        BubbleEvent = false;
                        break;
                }
            }
            else if (pVal.BeforeAction == false)
            {
                switch (pVal.MenuUID)
                {
                    case "1281":    // 찾기
                        //Find Mode
                        SetItemEnabled();
                        SetToolBarEnabled();
                        SetFocusByFormMode();
                        // 틀고정
                        oForm.Freeze(true);
                        oMAT1.CommonSetting.FixedColumnsCount = 2;
                        oMAT2.CommonSetting.FixedColumnsCount = 2;
                        oForm.Freeze(false);
                        break;
                    case "1284":    // 취소 완료 후  2026.09.04.LKC
                        SetItemEnabled();
                        SetToolBarEnabled();
                        break;
                    case "1282":    // 추가
                        //Add Mode
                        SetFormBase();
                        SetItemEnabled();
                        SetToolBarEnabled();
                        SetFocusByFormMode();
                        // 틀고정
                        oForm.Freeze(true);
                        oMAT1.CommonSetting.FixedColumnsCount = 2;
                        oMAT2.CommonSetting.FixedColumnsCount = 2;
                        oForm.Freeze(false);
                        break;
                    //case "1287":
                    //    //복제 처리
                    //    bool oEnableMenu = false;
                    //    oForm.Freeze(true);
                    //    //문서번호
                    //    oDBDSH.SetValue("DocEntry",  0, ZSBO.ZBSFC.FGetDocNum("ZMC112"));
                    //    //문서 기본값
                    //    oDBDSH.SetValue("U_DOCSTAT", 0, "O");
                    //    //Canceled 기본값
                    //    oDBDSH.SetValue("CANCELED",  0, "N");
                    //    oDBDSH.SetValue("U_DINUM",   0, "");
                    //    oDBDSH.SetValue("U_DINUM2",  0, "");
                    //    oDBDSH.SetValue("U_DICNUM",  0, "");
                    //    oDBDSH.SetValue("U_DICNUM2", 0, "");
                    //    oMAT1.FlushToDataSource();
                    //    for (int i = 0; i < oMAT1.VisualRowCount; i++)
                    //    {
                    //        oDBDSL.SetValue("DocEntry", i, "");
                    //    }
                    //    oMAT1.LoadFromDataSource();
                    //    oForm.Items.Item("DOCENTRY").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    //    if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) oEnableMenu = true;
                    //    oForm.EnableMenu("1287", oEnableMenu);
                    //    oForm.Freeze(false);
                    //    //문서 닫힘, CANCELED상태에서 복제시 화면 처리
                    //    SetItemEnabled();
                    //    SetToolBarEnabled();
                    //    SetFocusByFormMode();
                    //    break;
                    case "1288":
                    case "1289":
                    case "1290":
                    case "1291":                
                        SetItemEnabled();
                        SetToolBarEnabled();

                        oMAT1.AutoResizeColumns();
                        oMAT2.AutoResizeColumns();
                        break;
                    case "1292": // 행 추가
                        activeMatUID = m_ActiveMatrixUID;
                        if (activeMatUID == "MAT1")
                        {
                            oRow = ZSBO.ZBSFC.FGetIndexRow(oMAT1);
                            AddRow(oMAT1, oDBDSL, "U_ITEMCODE", oRow);
                            if (oRow == 0) { oRow = oMAT1.VisualRowCount; }
                            //행추가 후 위치
                            oMAT1.Columns.Item("ITEMCODE").Cells.Item(oRow).Click(SAPbouiCOM.BoCellClickType.ct_Regular, 1);
                        }
                        else if (activeMatUID == "MAT2")
                        {
                            oRow = ZSBO.ZBSFC.FGetIndexRow(oMAT2);
                            AddRow(oMAT2, oDBDSM, "U_ITEMCODE", oRow);
                            if (oRow == 0) { oRow = oMAT2.VisualRowCount; }
                            //행추가 후 위치
                            oMAT2.Columns.Item("ITEMCODE").Cells.Item(oRow).Click(SAPbouiCOM.BoCellClickType.ct_Regular, 1);
                        }
                        break;
                }
            }
        }
        //-------------------------------------------------------------------------------------------------------
        catch (Exception Error)
        {
            oForm.Freeze(false);
            ZSBO.SBO_Application.SetStatusBarMessage(Error.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        }
    }

    //************************************************************************************************************
    // FlushToItemValue
    //************************************************************************************************************
    private void FlushToItemValue(string sUID, string sCol, int sRow, string sValue)
    {
        oForm.Freeze(true);
        SAPbobsCOM.Recordset oRecordset = (SAPbobsCOM.Recordset)ZSBO.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        oMAT1 = (SAPbouiCOM.Matrix)oForm.Items.Item("MAT1").Specific;
        oDBDSH = oForm.DataSources.DBDataSources.Item("@ZMC112H");
        oDBDSL = oForm.DataSources.DBDataSources.Item("@ZMC112L");
        oDBDSM = oForm.DataSources.DBDataSources.Item("@ZMC112M");
        //int i = 0;

        var oBinCodesList = new List<string>();
        string[] oBinCodes;
        string oBinCode, oItemCode, oWhsCode;

        try
        {
            switch (sUID)
            {
                case "ITEMCODE":
                    if (sValue == "")
                    {
                        oDBDSH.SetValue("U_ITEMCODE", 0, "");
                        oDBDSH.SetValue("U_ITEMNAME", 0, "");
                    }
                    else
                    {
                        oRecordset.DoQuery("SELECT ItemCode, ItemName FROM OITM WHERE ItemCode = '" + sValue + "'");
                        oDBDSH.SetValue("U_ITEMCODE", 0, oRecordset.Fields.Item(0).Value.ToString());
                        oDBDSH.SetValue("U_ITEMNAME", 0, oRecordset.Fields.Item(1).Value.ToString());
                    }
                    oForm.Items.Item("ITEMCODE").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    break;

                case "ITEMNAME":
                    if (sValue == "")
                    {
                        oDBDSH.SetValue("U_ITEMCODE", 0, "");
                        oDBDSH.SetValue("U_ITEMNAME", 0, "");
                    }
                    else
                    {
                        oRecordset.DoQuery("SELECT ItemCode, ItemName FROM OITM WHERE ItemCode = '" + sValue + "'");
                        oDBDSH.SetValue("U_ITEMCODE", 0, oRecordset.Fields.Item(0).Value.ToString());
                        oDBDSH.SetValue("U_ITEMNAME", 0, oRecordset.Fields.Item(1).Value.ToString());
                    }
                    oForm.Items.Item("ITEMNAME").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    break;

                case "TITEMCD":
                    if (sValue == "")
                    {
                        oDBDSH.SetValue("U_TITEMCD", 0, "");
                        oDBDSH.SetValue("U_TITEMNM", 0, "");
                    }
                    else
                    {
                        oRecordset.DoQuery("SELECT ItemCode, ItemName FROM OITM WHERE ItemCode = '" + sValue + "'");
                        oDBDSH.SetValue("U_TITEMCD", 0, oRecordset.Fields.Item(0).Value.ToString());
                        oDBDSH.SetValue("U_TITEMNM", 0, oRecordset.Fields.Item(1).Value.ToString());
                    }

                    if (!m_BatchUpdated) // CFL에서 이미 처리된 경우 스킵
                    {
                        if (ZSBO.SBO_Application.MessageBox("대체품번을 초기화 합니다.", 1, "예", "", "") == 1)
                        {
                            oDBDSM.Clear();
                            oMAT2.Clear();
                            AddRow(oMAT2, oDBDSM, "U_ITEMCODE", 0);
                            oDBDSM.Offset = oDBDSM.Size - 1;
                            oMAT2.FlushToDataSource();
                            oDBDSM.SetValue("U_ITEMCODE", oDBDSM.Size - 1, oDBDSH.GetValue("U_TITEMCD", 0).ToString().Trim());
                            oDBDSM.SetValue("U_ITEMNAME", oDBDSM.Size - 1, oDBDSH.GetValue("U_TITEMNM", 0).ToString().Trim());
                            oMAT2.LoadFromDataSource();
                        }
                    }
                    m_BatchUpdated = false;

                    oForm.Items.Item("TITEMCD").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    break;

                case "TITEMNM":
                    if (sValue == "")
                    {
                        oDBDSH.SetValue("U_TITEMCD", 0, "");
                        oDBDSH.SetValue("U_TITEMNM", 0, "");
                    }
                    else
                    {
                        oRecordset.DoQuery("SELECT ItemCode, ItemName FROM OITM WHERE ItemCode = '" + sValue + "'");
                        oDBDSH.SetValue("U_TITEMCD", 0, oRecordset.Fields.Item(0).Value.ToString());
                        oDBDSH.SetValue("U_TITEMNM", 0, oRecordset.Fields.Item(1).Value.ToString());
                    }

                    if (!m_BatchUpdated) // CFL에서 이미 처리된 경우 스킵
                    {
                        if (ZSBO.SBO_Application.MessageBox("대체품번을 초기화 합니다.", 1, "예", "", "") == 1)
                        {
                            oDBDSM.Clear();
                            oMAT2.Clear();
                            AddRow(oMAT2, oDBDSM, "U_ITEMCODE", 0);
                            oDBDSM.Offset = oDBDSM.Size - 1;
                            oMAT2.FlushToDataSource();
                            oDBDSM.SetValue("U_ITEMCODE", oDBDSM.Size - 1, oDBDSH.GetValue("U_TITEMCD", 0).ToString().Trim());
                            oDBDSM.SetValue("U_ITEMNAME", oDBDSM.Size - 1, oDBDSH.GetValue("U_TITEMNM", 0).ToString().Trim());
                            oMAT2.LoadFromDataSource();
                        }
                    }
                    m_BatchUpdated = false;

                    oForm.Items.Item("TITEMNM").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    break;
            }

            switch (sUID + sCol)
            {
                case "MAT1ITEMCODE":
                    oMAT1.FlushToDataSource();
                    oDBDSL.Offset = sRow - 1;
                    if (sValue == "")
                    {
                        oDBDSL.SetValue("U_ITEMCODE", sRow - 1, "");
                        oDBDSL.SetValue("U_ITEMNAME", sRow - 1, "");
                    }
                    else
                    {
                        oRecordset.DoQuery("SELECT ItemCode, ItemName FROM OITM WHERE ItemCode = '" + sValue + "'");
                        oDBDSL.SetValue("U_ITEMCODE", sRow - 1, oRecordset.Fields.Item(0).Value.ToString());
                        oDBDSL.SetValue("U_ITEMNAME", sRow - 1, oRecordset.Fields.Item(1).Value.ToString());
                        //oDBDSL.SetValue("U_ITEMNM01", sRow - 1, oRecordset.Fields.Item(2).Value.ToString());
                        oMAT1.SetLineData(sRow);
                    }
                    oMAT1.SetLineData(sRow);
                    if (sRow == oMAT1.VisualRowCount & oDBDSL.GetValue("U_ITEMCODE", sRow - 1).Trim() != "")
                    {
                        AddRow(oMAT1, oDBDSL, "U_ITEMCODE", sRow + 1);
                        oMAT1.Columns.Item(sCol).Cells.Item(sRow + 1).Click(SAPbouiCOM.BoCellClickType.ct_Regular, 1);
                    }
                    break;

                case "MAT1ITEMNAME":
                    oMAT1.FlushToDataSource();
                    oDBDSL.Offset = sRow - 1;
                    if (sValue == "")
                    {
                        oDBDSL.SetValue("U_ITEMCODE", sRow - 1, "");
                        oDBDSL.SetValue("U_ITEMNAME", sRow - 1, "");
                    }
                    else
                    {
                        oRecordset.DoQuery("SELECT ItemCode, ItemName FROM OITM WHERE ItemName = '" + sValue + "'");
                        oDBDSL.SetValue("U_ITEMCODE", sRow - 1, oRecordset.Fields.Item(0).Value.ToString());
                        oDBDSL.SetValue("U_ITEMNAME", sRow - 1, oRecordset.Fields.Item(1).Value.ToString());
                        //oDBDSL.SetValue("U_ITEMNM01", sRow - 1, oRecordset.Fields.Item(2).Value.ToString());
                        oMAT1.SetLineData(sRow);
                    }
                    oMAT1.SetLineData(sRow);
                    if (sRow == oMAT1.VisualRowCount & oDBDSL.GetValue("U_ITEMNAME", sRow - 1).Trim() != "")
                    {
                        AddRow(oMAT1, oDBDSL, "U_ITEMNAME", sRow + 1);
                        oMAT1.Columns.Item(sCol).Cells.Item(sRow + 1).Click(SAPbouiCOM.BoCellClickType.ct_Regular, 1);
                    }
                    break;

                case "MAT2ITEMCODE":
                    oMAT2.FlushToDataSource();
                    oDBDSM.Offset = sRow - 1;
                    if (sValue == "")
                    {
                        oDBDSM.SetValue("U_ITEMCODE", sRow - 1, "");
                        oDBDSM.SetValue("U_ITEMNAME", sRow - 1, "");
                    }
                    else
                    {
                        oRecordset.DoQuery("SELECT ItemCode, ItemName FROM OITM WHERE ItemCode = '" + sValue + "'");
                        oDBDSM.SetValue("U_ITEMCODE", sRow - 1, oRecordset.Fields.Item(0).Value.ToString());
                        oDBDSM.SetValue("U_ITEMNAME", sRow - 1, oRecordset.Fields.Item(1).Value.ToString());
                        //oDBDSM.SetValue("U_TITEMNM1", sRow - 1, oRecordset.Fields.Item(2).Value.ToString());
                        oMAT2.SetLineData(sRow);
                    }
                    oMAT2.SetLineData(sRow);
                    break;

                case "MAT2TITEMNAME":
                    oMAT2.FlushToDataSource();
                    oDBDSM.Offset = sRow - 1;
                    if (sValue == "")
                    {
                        oDBDSM.SetValue("U_ITEMCODE", sRow - 1, "");
                        oDBDSM.SetValue("U_ITEMNAME", sRow - 1, "");
                    }
                    else
                    {
                        oRecordset.DoQuery("SELECT ItemCode, ItemName FROM OITM WHERE ItemName = '" + sValue + "'");
                        oDBDSM.SetValue("U_ITEMCODE", sRow - 1, oRecordset.Fields.Item(0).Value.ToString());
                        oDBDSM.SetValue("U_ITEMNAME", sRow - 1, oRecordset.Fields.Item(1).Value.ToString());
                        //oDBDSM.SetValue("U_TITEMNM1", sRow - 1, oRecordset.Fields.Item(2).Value.ToString());
                        oMAT2.SetLineData(sRow);
                    }
                    oMAT2.SetLineData(sRow);
                    break;

                case "MAT1WHSCODE":  //출고창고((MAT1)
                    oMAT1.FlushToDataSource();
                    oDBDSL.Offset = sRow - 1;
                    oRecordset.DoQuery("SELECT WhsCode FROM OWHS WHERE WhsCode = '" + sValue + "' AND Inactive = 'N' ");
                    if (sValue == "")
                    {
                        oDBDSL.SetValue("U_WHSCODE", sRow - 1, "");
                        oDBDSL.SetValue("U_BINCODE", sRow - 1, ""); //창고변경시 BinCode는 초기화
                        oDBDSL.SetValue("U_ONHAND", sRow - 1, 0.ToString()); //창고변경시 BIN재고수량 초기화
                        oDBDSL.SetValue("U_WONHAND", sRow - 1, 0.ToString()); //창고변경시 창고재고수량 초기화 2026.04.24.LKC.
                    }
                    else
                    {
                        oDBDSL.SetValue("U_WHSCODE", sRow - 1, oRecordset.Fields.Item(0).Value.ToString());
                        oItemCode = oDBDSL.GetValue("U_ITEMCODE", sRow - 1).ToString().Trim(); //2026.03.11.LKC. 창고 변경시 BIN 로케이션 및 재고 설정
                        oWhsCode = oDBDSL.GetValue("U_WHSCODE", sRow - 1).ToString().Trim();
                        //oConds = (SAPbouiCOM.Conditions)ZSBO.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_Conditions);

                        oRecordset.DoQuery("EXEC [DBO].[ZTMSP_ZMC112_BIN] @TYPE     = 'TO' " //2026.03.13.LKC. 파라미터 값 변경 'O' -> 'TO'
                                                                    + " , @ITEMCODE = '" + oItemCode + "', @WHSCODE  = '" + oWhsCode + "' ");
                        while (oRecordset.EoF == false)
                        {
                            //중복 값은 제외함
                            oBinCode = oRecordset.Fields.Item(0).Value.ToString();
                            if (!oBinCodesList.Contains(oBinCode)) oBinCodesList.Add(oBinCode); //중복 제외
                            oRecordset.MoveNext();
                        }

                        oBinCodes = oBinCodesList.ToArray();
                        if (oBinCodesList.Count == 1)
                        {
                            oBinCode = oBinCodes.GetValue(0).ToString().Trim();
                            //출고BIN
                            oDBDSL.SetValue("U_BINCODE", sRow - 1, oBinCode);
                            //재고계산
                            oDBDSL.SetValue("U_ONHAND", sRow - 1, GetOnHand("B", oItemCode, oWhsCode, oBinCode).ToString()); //BIN재고
                            oDBDSL.SetValue("U_WONHAND", sRow - 1, GetOnHand("W", oItemCode, oWhsCode, oBinCode).ToString()); //창고재고 2026.04.24.LKC.
                        }
                        else
                        {
                            oDBDSL.SetValue("U_BINCODE", sRow - 1, "");
                            oDBDSL.SetValue("U_ONHAND", sRow - 1, 0.ToString()); //BIN재고
                            oDBDSL.SetValue("U_WONHAND", sRow - 1, 0.ToString()); //창고재고 2026.04.24.LKC.
                        }
                    }
                    oMAT1.SetLineData(sRow);

                    //SetChooseFormList("WHSCODE", oRecordSet.Fields.Item(0).Value.ToString(), 0);
                    break;

                case "MAT2TWHSCODE": //입고창고(MAT2)
                    oMAT2.FlushToDataSource();
                    oDBDSM.Offset = sRow - 1;
                    oRecordset.DoQuery("SELECT WhsCode FROM OWHS WHERE WhsCode = '" + sValue + "' AND Inactive = 'N' ");
                    if (sValue == "")
                    {
                        oDBDSM.SetValue("U_TWHSCODE", sRow - 1, "");
                        oDBDSM.SetValue("U_TBINCODE", sRow - 1, ""); //창고변경시 BinCode는 초기화
                    }
                    else
                    {
                        oItemCode = oDBDSM.GetValue("U_TITEMCD", sRow - 1).ToString().Trim();
                        oWhsCode = oDBDSM.GetValue("U_TWHSCODE", sRow - 1).ToString().Trim();

                        oRecordset.DoQuery("EXEC [DBO].[ZTMSP_ZMC112_BIN] @TYPE     = 'TI' " //입고창고 변경시 입고BIN 재설정. 
                                                                    + " , @ITEMCODE = '" + oItemCode + "', @WHSCODE  = '" + oWhsCode + "' ");
                        while (oRecordset.EoF == false)
                        {
                            //중복 값은 제외함
                            oBinCode = oRecordset.Fields.Item(0).Value.ToString();
                            if (!oBinCodesList.Contains(oBinCode)) oBinCodesList.Add(oBinCode); //중복 제외
                            oRecordset.MoveNext();
                        }

                        oBinCodes = oBinCodesList.ToArray();
                        if (oBinCodesList.Count == 1)
                        {
                            oBinCode = oBinCodes.GetValue(0).ToString().Trim();
                            //입고BIN
                            oDBDSM.SetValue("U_BINCODE", sRow - 1, oBinCode);
                        }
                        else
                        {
                            oDBDSM.SetValue("U_BINCODE", sRow - 1, "");
                        }
                    }
                    oMAT2.SetLineData(sRow);
                    break;

                case "MAT1BINCODE": //출고BIN(MAT1)
                    oMAT1.FlushToDataSource();
                    oDBDSL.Offset = sRow - 1;
                    if (sValue == "")
                    {
                        oDBDSL.SetValue("U_BINCODE", sRow - 1, "");
                        //창고변경에 따른 BIN 재고수량 초기화
                        oDBDSL.SetValue("U_ONHAND", sRow - 1, 0.ToString()); //출고BIN 변경시 BIN재고수량 초기화
                        oDBDSL.SetValue("U_WONHAND", sRow - 1, 0.ToString()); //출고BIN 변경시 창고재고수량 초기화. 2026.04.24.LKC.
                    }
                    oMAT1.SetLineData(sRow);
                    break;

                case "MAT2BINCODE": //입고BIN(MAT2)
                    oMAT2.FlushToDataSource();
                    oDBDSM.Offset = sRow - 1;
                    if (sValue == "")
                    {
                        oDBDSM.SetValue("U_BINCODE", sRow - 1, "");
                    }
                    oMAT2.SetLineData(sRow);
                    break;

                case "MAT1SERIALNO": //일련번호
                    if (sValue != "")
                    {
                        oMAT1.FlushToDataSource();
                        oDBDSL.Offset = sRow - 1;
                        oRecordset.DoQuery("SELECT ManBtchNum, ManSerNum FROM OITM WHERE ITEMCODE = '" + oDBDSL.GetValue("U_ITEMCODE", sRow - 1).Trim() + "'");
                        if (oRecordset.Fields.Item("ManSerNum").Value.ToString() != "Y")
                        {
                            oDBDSL.SetValue("U_SERIALNO", sRow - 1, "");
                            ZSBO.SBO_Application.StatusBar.SetText("시리얼번호 입력대상 품번이 아닙니다.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                        }
                        oMAT1.SetLineData(sRow);
                    }
                    break;

                case "MAT2SERIALNO": //일련번호
                    if (sValue != "")
                    {
                        oMAT2.FlushToDataSource();
                        oDBDSM.Offset = sRow - 1;
                        oRecordset.DoQuery("SELECT ManBtchNum, ManSerNum FROM OITM WHERE ITEMCODE = '" + oDBDSM.GetValue("U_ITEMCODE", sRow - 1).Trim() + "'");
                        if (oRecordset.Fields.Item("ManSerNum").Value.ToString() != "Y")
                        {
                            oDBDSM.SetValue("U_SERIALNO", sRow - 1, "");
                            oDBDSM.SetValue("U_LREMARK", sRow - 1, "시리얼 오류");
                            ZSBO.SBO_Application.StatusBar.SetText("시리얼번호 입력대상 품번이 아닙니다.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                        }
                        else
                        { 
                            oDBDSM.SetValue("U_LREMARK", sRow - 1, ""); 
                        }
                        oMAT2.SetLineData(sRow);
                    }
                    break;

                case "MAT1BATCHNO": //배치번호
                    if (sValue != "")
                    {
                        oMAT1.FlushToDataSource();
                        oDBDSL.Offset = sRow - 1;
                        oRecordset.DoQuery("SELECT ManBtchNum, ManSerNum FROM OITM WHERE ITEMCODE = '" + oDBDSL.GetValue("U_ITEMCODE", sRow - 1).Trim() + "'");
                        if (oRecordset.Fields.Item("ManBtchNum").Value.ToString() != "Y")
                        {
                            oDBDSL.SetValue("U_BATCHNO", sRow - 1, "");
                            ZSBO.SBO_Application.StatusBar.SetText("배치번호 입력대상 품번이 아닙니다.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                        }
                        oMAT1.SetLineData(sRow);
                    }
                    break;

                case "MAT2BATCHNO": //배치번호
                    if (sValue != "")
                    {
                        oMAT2.FlushToDataSource();
                        oDBDSM.Offset = sRow - 1;
                        oRecordset.DoQuery("SELECT ManBtchNum, ManSerNum FROM OITM WHERE ITEMCODE = '" + oDBDSM.GetValue("U_ITEMCODE", sRow - 1).Trim() + "'");
                        if (oRecordset.Fields.Item("ManBtchNum").Value.ToString() != "Y")
                        {
                            oDBDSM.SetValue("U_BATCHNO", sRow - 1, "");
                            oDBDSM.SetValue("U_LREMARK", sRow - 1, "배치번호 오류");
                            ZSBO.SBO_Application.StatusBar.SetText("배치번호 입력대상 품번이 아닙니다.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                        }
                        else
                        {
                            oDBDSM.SetValue("U_LREMARK", sRow - 1, "");
                        }
                        oMAT2.SetLineData(sRow);
                    }
                    break;
                case "MAT1TAGNO": //MAT1 TAG번호
                    if (sValue != "")
                    {
                        oMAT1.FlushToDataSource();
                        oDBDSL.Offset = sRow - 1;
                        oRecordset.DoQuery("SELECT ManBtchNum, ManSerNum FROM OITM WHERE ITEMCODE = '" + oDBDSL.GetValue("U_ITEMCODE", sRow - 1).Trim() + "'");
                        if (oRecordset.Fields.Item("ManBtchNum").Value.ToString() == "Y")
                        {
                            oDBDSL.SetValue("U_TAGNO", sRow - 1, "");
                            //oDBDSM.SetValue("U_LREMARK", sRow - 1, "TAG번호 입력 오류");
                            ZSBO.SBO_Application.StatusBar.SetText("배치번호는 TAG번호를 입력할 수 없습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                        }
                        //else
                        //{
                        //    oDBDSL.SetValue("U_LREMARK", sRow - 1, "");
                        //}
                        oMAT1.SetLineData(sRow);
                    }
                    break;
                case "MAT2TAGNO": //MAT2 TAG번호
                    if (sValue != "")
                    {
                        oMAT2.FlushToDataSource();
                        oDBDSM.Offset = sRow - 1;
                        oRecordset.DoQuery("SELECT ManBtchNum, ManSerNum FROM OITM WHERE ITEMCODE = '" + oDBDSM.GetValue("U_ITEMCODE", sRow - 1).Trim() + "'");
                        if (oRecordset.Fields.Item("ManBtchNum").Value.ToString() == "Y")
                        {
                            oDBDSM.SetValue("U_TAGNO", sRow - 1, "");
                            //oDBDSM.SetValue("U_LREMARK", sRow - 1, "TAG번호 입력 오류");
                            ZSBO.SBO_Application.StatusBar.SetText("배치번호는 TAG번호를 입력할 수 없습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                        }
                        //else
                        //{
                        //    oDBDSM.SetValue("U_LREMARK", sRow - 1, "");
                        //}
                        oMAT2.SetLineData(sRow);
                    }
                    break;
            }
            //oForm.Freeze(false);
        }
        //-------------------------------------------------------------------------------------------------------
        catch (Exception ex)
        {
            
            ZSBO.SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
        }
        finally
        {
            oForm.Freeze(false);
            if (oRecordset != null) { oRecordset = null; }
        }
    }

    //************************************************************************************************************
    // FormDataEventHander
    //************************************************************************************************************
    public void SBO_Application_FormDataEvent(ref SAPbouiCOM.BusinessObjectInfo oInfo, out bool BubbleEvent)
    {
        BubbleEvent = true;

        try
        {
            if (oInfo.BeforeAction) { return; }
            //---------------------------------------------------------------------------------------------------
            switch (oInfo.EventType)
            {
                case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:
                    SetItemEnabled();
                    SetToolBarEnabled();
                    if (oForm.Items.Item("MAT1").Enabled == true) { AddRow(oMAT1, oDBDSL, "U_ITEMCODE", 0); }
                    if (oForm.Items.Item("MAT2").Enabled == true) { AddRow(oMAT2, oDBDSM, "U_ITEMCODE", 0); }
                    oMAT1.Columns.Item("ITEMCODE").Cells.Item(1).Click(SAPbouiCOM.BoCellClickType.ct_Regular, 1);
                    break;
            }
        }
        //-------------------------------------------------------------------------------------------------------
        catch (Exception ex)
        {
            oForm.Freeze(false);
            ZSBO.SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        }
    }

    //************************************************************************************************************
    // 데이터 체크
    //************************************************************************************************************
    private bool ChkInData()
    {
        SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)ZSBO.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        double bQty, wQty, qQty, oMat1Qty = 0, oMat2Qty = 0;


        oForm.Freeze(true);
        try
        {
            DelRow(oMAT1, oDBDSL, "U_ITEMCODE");           // 행삭제
            DelRow(oMAT2, oDBDSM, "U_ITEMCODE");           // 행삭제
            //if (oDBDSH.GetValue("U_REQTYPE", 0).ToString().Trim() == "") 2026.02.19.LKC. 필요없는 칼럼 주석처리
            //{
            //    oForm.Items.Item("REQTYPE").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
            //    ZSBO.SBO_Application.StatusBar.SetText("요청구분이 누락되었습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            //    return false;
            //}
            if (oDBDSH.GetValue("U_IOTYPE", 0).ToString().Trim() == "")
            {
                oForm.Items.Item("IOTYPE").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                ZSBO.SBO_Application.StatusBar.SetText("입출고유형이 누락되었습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                return false;
            }
            //if (oDBDSH.GetValue("U_ITEMCODE", 0).ToString().Trim() == "")
            //{
            //    oForm.Items.Item("ITEMCODE").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
            //    ZSBO.SBO_Application.StatusBar.SetText("품번이 누락되었습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            //    return false;
            //}
            //if (oDBDSH.GetValue("U_TITEMCD", 0).ToString().Trim() == "")
            //{
            //    oForm.Items.Item("TITEMCD").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
            //    ZSBO.SBO_Application.StatusBar.SetText("대체품번이 누락되었습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            //    return false;
            //}
            if (oDBDSH.GetValue("U_DOCDATE", 0).ToString().Trim() == "")
            {
                oForm.Items.Item("DOCDATE").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                ZSBO.SBO_Application.StatusBar.SetText("전기일이 누락되었습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                return false;
            }

            if (oMAT1.VisualRowCount == 0)
            {
                ZSBO.SBO_Application.StatusBar.SetText("품번 내역이 누락되었습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                return false;
            }
            if (oMAT2.VisualRowCount == 0)
            {
                ZSBO.SBO_Application.StatusBar.SetText("대체품번 내역이 누락되었습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                return false;
            }

            // 라인1
            oMAT1.FlushToDataSource(); //출고체크
            string oItemCode = oDBDSH.GetValue("U_ITEMCODE", 0).ToString().Trim();
            string tItemCode = oDBDSH.GetValue("U_TITEMCD",  0).ToString().Trim();
            for (int i = 0; i <= oMAT1.VisualRowCount - 1; i++)
            {
                if (oDBDSL.GetValue("U_ITEMCODE", i).ToString() == "")
                {
                    ZSBO.SBO_Application.StatusBar.SetText("품번이 누락되었습니다. [" + (i + 1).ToString() + "라인]", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    return false;
                }
                if (CheckIsNull(oItemCode) != "" && oItemCode != oDBDSL.GetValue("U_ITEMCODE", i).ToString())
                {
                    ZSBO.SBO_Application.StatusBar.SetText("헤더의 품번과 일치하지 않습니다. [" + (i + 1).ToString() + "라인]", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    return false;
                }
                if (oDBDSL.GetValue("U_ITEMNAME", i).ToString() == "")
                {
                    ZSBO.SBO_Application.StatusBar.SetText("품명이 누락되었습니다. [" + (i + 1).ToString() + "라인]", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    return false;
                }
                if (oDBDSL.GetValue("U_QUANTITY", i).ToString() == "" || Double.Parse(oDBDSL.GetValue("U_QUANTITY", i).ToString().Trim()) <= 0)
                {
                    ZSBO.SBO_Application.StatusBar.SetText("수량이 누락되었습니다. [" + (i + 1).ToString() + "라인]", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    return false;
                }

                //재고체크                
                bQty = double.Parse(GetOnHand("B", oDBDSL.GetValue("U_ITEMCODE", i), oDBDSL.GetValue("U_WHSCODE", i), oDBDSL.GetValue("U_BINCODE", i)).ToString()); //BIN재고수량
                //wQty = double.Parse(GetOnHand("W", oDBDSL.GetValue("U_ITEMCODE", i), oDBDSL.GetValue("U_WHSCODE", i), oDBDSL.GetValue("U_BINCODE", i)).ToString()); //창고재고수량
                qQty = double.Parse(oDBDSL.GetValue("U_QUANTITY", i).Trim());           //출고품번수량

                //if ( bQty - qQty  < 0 || wQty - qQty  < 0) //BIN재고만 체크하기로 함
                if (bQty - qQty < 0) 
                {
                    ZSBO.SBO_Application.StatusBar.SetText(oDBDSL.GetValue("U_ITEMCODE", i).ToString() + " 품번의 재고수량이 부족합니다. [" + (i + 1).ToString() + "라인]", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    return false;
                }
                oMat1Qty = oMat1Qty + qQty;
            }

            //라인2
            oMAT2.FlushToDataSource(); //입고 체크
            for (int i = 0; i <= oMAT2.VisualRowCount - 1; i++)
            {
                if (oDBDSM.GetValue("U_ITEMCODE", i).ToString() == "")
                {
                    ZSBO.SBO_Application.StatusBar.SetText("대체품번이 누락되었습니다. [" + (i + 1).ToString() + "라인]", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    return false;
                }
                if (CheckIsNull(tItemCode) != "" && tItemCode != oDBDSM.GetValue("U_ITEMCODE", i).ToString())
                {
                    ZSBO.SBO_Application.StatusBar.SetText("헤더의 대체품번과 일치하지 않습니다. [" + (i + 1).ToString() + "라인]", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    return false;
                }
                if (oDBDSM.GetValue("U_ITEMNAME", i).ToString() == "")
                {
                    ZSBO.SBO_Application.StatusBar.SetText("대체품명이 누락되었습니다. [" + (i + 1).ToString() + "라인]", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    return false;
                }
                if (oDBDSM.GetValue("U_QUANTITY", i).ToString() == "" || Double.Parse(oDBDSM.GetValue("U_QUANTITY", i).ToString().Trim()) <= 0)
                {
                    ZSBO.SBO_Application.StatusBar.SetText("대체품번의 수량이 누락되었습니다. [" + (i + 1).ToString() + "라인]", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    return false;
                }

                qQty = double.Parse(oDBDSM.GetValue("U_QUANTITY", i).Trim());           //입고품번수량
                oMat2Qty = oMat2Qty + qQty;
            }

            if (oMat1Qty != oMat2Qty)
            {
                ZSBO.SBO_Application.StatusBar.SetText("출고수량 합계와 입고수량 합계가 일치하지 않습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                return false;
            }

            if (oDBDSH.GetValue("U_ITEMCODE", 0).ToString().Trim() != "" && oDBDSH.GetValue("U_TITEMCD", 0).ToString().Trim() != "")
            {                
                if (oDBDSH.GetValue("U_ITEMCODE", 0).ToString().Trim() == oDBDSH.GetValue("U_TITEMCD", 0).ToString().Trim())
                {
                    oForm.Items.Item("ITEMCODE").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    ZSBO.SBO_Application.StatusBar.SetText("품번과 대체품번이 같습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    return false;
                }
            }

            return true;
        }
        //-------------------------------------------------------------------------------------------------------
        catch (Exception Error)
        {
            ZSBO.SBO_Application.StatusBar.SetText(Error.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            return false;
        }
        finally
        {
            oForm.Freeze(false);
        }
    }

    //************************************************************************************************************
    // 취소 데이터 체크(재고 체크)
    //************************************************************************************************************
    private bool ChkCInData()
    {
        SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)ZSBO.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        oForm.Freeze(true);
        try
        {
            // 라인
            double bQty = 0, wQty = 0, qQty = 0;
            //출고,입고DI 문서번호 유무 확인
            if (CheckIsNull(oDBDSH.GetValue("U_DINUM", 0)) == "" || CheckIsNull(oDBDSH.GetValue("U_DINUM2", 0)) == "")
            {
                ZSBO.SBO_Application.StatusBar.SetText("입출고DI 발생내역이 없습니다. ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                return false;
            }
            //출고,입고취소DI 문서번호 유무 확인
            if (CheckIsNull(oDBDSH.GetValue("U_DICNUM", 0)) != "" || CheckIsNull(oDBDSH.GetValue("U_DICNUM2", 0)) != "")
            {
                ZSBO.SBO_Application.StatusBar.SetText("입출고취소DI 발생내역이 있습니다. ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                return false;
            }
            oMAT2.FlushToDataSource(); //취소일 경우 입고 품목의 재고 체크함
            for (int i = 0; i <= oMAT2.VisualRowCount - 1; i++)
            {
                bQty = bQty + double.Parse(GetOnHand("B", oDBDSM.GetValue("U_ITEMCODE", i), oDBDSM.GetValue("U_WHSCODE", i), oDBDSM.GetValue("U_BINCODE", i)).ToString()); //BIN재고수량
                qQty = qQty + double.Parse(oDBDSM.GetValue("U_QUANTITY", i).Trim());           //출고품번수량
            }
            if (bQty - qQty < 0)
            {
                ZSBO.SBO_Application.StatusBar.SetText(oDBDSM.GetValue("U_ITEMCODE", 0).ToString() + " 품번의 재고수량이 부족합니다.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                return false;
            }
            return true;
        }
        //-------------------------------------------------------------------------------------------------------
        catch (Exception Error)
        {
            ZSBO.SBO_Application.StatusBar.SetText(Error.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            return false;
        }
        finally
        {
            oForm.Freeze(false);
        }
    }

    //************************************************************************************************************
    // 행추가 (sChkFld = 실제 필드 이름. ex) U_GCode)
    //************************************************************************************************************
    private void AddRow(SAPbouiCOM.Matrix sMatrix, SAPbouiCOM.DBDataSource sDBDS, string sChkFld, int sRow)
    {
        oForm.Freeze(true);
        int i = 0, j = 0;
        bool oBool = false;
        string oKeyFld = "DocEntry";
        //---------------------------------------------------------------------------------------------------
        sMatrix.FlushToDataSource();
        if (sRow <= 0)
        {
            if (sDBDS.Size == 0) { oBool = true; }
            if (sDBDS.Size != 0)
            {
                sDBDS.Offset = sDBDS.Size - 1;
                if (sDBDS.GetValue(sChkFld, sDBDS.Size - 1).Trim() != "") { oBool = true; }
            }
            if (oBool == true) { sDBDS.InsertRecord(sDBDS.Size); }
        }
        else
        {
            oBool = true;
            sDBDS.InsertRecord(sRow - 1);
            for (j = 0; j < sDBDS.Fields.Count; j++)
            {
                sDBDS.Offset = sRow - 1;
                sDBDS.SetValue(sDBDS.Fields.Item(j).Name, sRow - 1, ZSBO.ZBSFC.FGetDfltValue(sDBDS.Fields.Item(j).Type, ""));
            }
        }
        //-----------------------------------------------------------------------------------------------
        if (oBool == true | sMatrix.VisualRowCount == 0)
        {
            for (i = 0; i < sDBDS.Size; i++)
            {
                sDBDS.Offset = i;
                if (oForm.Mode != SAPbouiCOM.BoFormMode.fm_ADD_MODE) { sDBDS.SetValue(oKeyFld, sDBDS.Size - 1, oDBDSH.GetValue(oKeyFld, 0).ToString()); }
                sDBDS.SetValue("LineId", i, (i + 1).ToString());
                //if (sDBDS.TableName.ToString().Trim() == "@ZMC112M" && CheckIsNull(oDBDSH.GetValue("U_TITEMCD", 0)) != "")
                //{
                //    sDBDS.SetValue("U_ITEMCODE", i, oDBDSH.GetValue("U_TITEMCD", 0).ToString());
                //    sDBDS.SetValue("U_ITEMNAME", i, oDBDSH.GetValue("U_TITEMNM", 0).ToString());
                //}
            }
            sMatrix.LoadFromDataSource();
        }
        //---------------------------------------------------------------------------------------------------
        oForm.Freeze(false);
    }

    //************************************************************************************************************
    // 행삭제 (기존 행삭제는 포커스값을 가져오지 못해서 새로 생성) 2022.01.26 BJH
    //************************************************************************************************************
    private void DeleteRow(SAPbouiCOM.Matrix sMatrix, SAPbouiCOM.DBDataSource sDBDS, string sChkFld)
    {
        int oRow;
        string oKeyFld = "DocEntry";
        oForm.Freeze(true);

        sMatrix.FlushToDataSource();
        oRow = ZSBO.ZBSFC.FGetIndexRow(sMatrix);

        if (oRow < 0) { return; }

        sDBDS.Offset = oRow - 1;
        sMatrix.DeleteRow(oRow);
        sDBDS.RemoveRecord(oRow - 1);

        for (int i = 0; i < sDBDS.Size; i++)
        {
            sDBDS.Offset = i;
            if (oForm.Mode != SAPbouiCOM.BoFormMode.fm_ADD_MODE) { sDBDS.SetValue(oKeyFld, sDBDS.Size - 1, sDBDS.GetValue(oKeyFld, 0).ToString()); }
            sDBDS.SetValue("LineId", i, (i + 1).ToString());
        }
        sMatrix.LoadFromDataSource();

        oForm.Freeze(false);
    }

    //************************************************************************************************************
    // 행삭제 (sChkFld = 실제 필드 이름. ex) U_GCode)
    //************************************************************************************************************
    private void DelRow(SAPbouiCOM.Matrix sMatrix, SAPbouiCOM.DBDataSource sDBDS, string sChkFld)
    {
        oForm.Freeze(true);
        int i = 0, j = 0;
        bool oBool = false;
        string oKeyFld = "DocEntry";
        //---------------------------------------------------------------------------------------------------
        sMatrix.FlushToDataSource();
        if (sDBDS.Size != 0)
        {
            sDBDS.Offset = sDBDS.Size - 1;
            if (sDBDS.GetValue(sChkFld, sDBDS.Size - 1).Trim() == "") { oBool = true; }
        }
        if (oBool == true) { sDBDS.RemoveRecord(sDBDS.Size - 1); }
        for (i = 0; i < sDBDS.Size; i++)
        {
            sDBDS.Offset = i;
            if (oForm.Mode != SAPbouiCOM.BoFormMode.fm_ADD_MODE) { sDBDS.SetValue(oKeyFld, sDBDS.Size - 1, oDBDSH.GetValue(oKeyFld, 0).ToString()); }
            sDBDS.SetValue("LineId", i, (i + 1).ToString());
        }
        sMatrix.LoadFromDataSource();
        //---------------------------------------------------------------------------------------------------
        oForm.Freeze(false);
    }

    //************************************************************************************************************
    // Form : SetFormBase()
    //************************************************************************************************************
    private void SetFormBase()
    {
        SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)ZSBO.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        SAPbouiCOM.ComboBox oCombo = null;
        string oDATE, sqlString, oDocDate;

        try
        {   
            oForm.Freeze(true);

            //문서번호
            oDBDSH.SetValue("DocEntry", 0, ZSBO.ZBSFC.FGetDocNum("ZMC112"));

            // 기본값 변경
            sqlString = "SELECT U_GRP1, U_GRP2 FROM [@ZPGMD] WHERE CODE = 'ZMC112' AND U_TABLE = 'BVALUE' AND U_VALUE ='" + ZSBO.oCompany.UserName + "' AND U_VALID = 'Y'";
            oRecordSet.DoQuery(sqlString);

            if (oRecordSet.RecordCount > 0)
            {
                // 입출고 유형
                oCombo = (SAPbouiCOM.ComboBox)oForm.Items.Item("IOTYPE").Specific;
                oCombo.Select(oRecordSet.Fields.Item(0).Value, SAPbouiCOM.BoSearchKey.psk_ByValue);

                //// 창고
                //oCombo = (SAPbouiCOM.ComboBox)oForm.Items.Item("WHSCODE").Specific;
                //oCombo.Select(oRecordSet.Fields.Item(1).Value, SAPbouiCOM.BoSearchKey.psk_ByValue);
            }

            //전기일 기본값
            sqlString = "SELECT CONVERT(NVARCHAR,GETDATE(),112)";
            oRecordSet.DoQuery(sqlString);
            oDATE = oRecordSet.Fields.Item(0).Value.ToString();
            oDBDSH.SetValue("U_DOCDATE", 0, oDATE);

            //문서 기본값
            oDBDSH.SetValue("U_DOCSTAT", 0, "O");
            //추가 상태로 초기화 시 이전 문서의 종료(DI처리/취소) 정보가 남아있으면
            //SetItemEnabled()가 문서를 닫힌 상태로 판단해 MAT1/MAT2 및 CFL('...')이 비활성화되므로 초기화함
            oDBDSH.SetValue("CANCELED",  0, "N");   // 2026.09.04.LKC. 주석 해제 - 취소문서 조회 후 추가 시 CFL 버튼 비활성화 방지
            oDBDSH.SetValue("U_DINUM",   0, "");    // 2026.09.04.LKC. 주석 해제
            oDBDSH.SetValue("U_DINUM2",  0, "");    // 2026.09.04.LKC. 주석 해제
            oDBDSH.SetValue("U_DICNUM",  0, "");    // 2026.09.04.LKC. 주석 해제
            oDBDSH.SetValue("U_DICNUM2", 0, "");    // 2026.09.04.LKC. 주석 해제

            // 라인 DataSource Clear
            oDBDSL.Clear();
            oMAT1.Clear();
            oDBDSM.Clear();
            oMAT2.Clear();
            if (oForm.Mode != SAPbouiCOM.BoFormMode.fm_FIND_MODE)
            {
                AddRow(oMAT1, oDBDSL, "U_ITEMCODE", 0);
                AddRow(oMAT2, oDBDSM, "U_ITEMCODE", 0);
            }
        }
        catch (Exception ex)
        {
            ZSBO.SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        }
        finally
        {
            oForm.Freeze(false);
            if (oRecordSet != null) oRecordSet = null;
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
            oXmlDoc.SelectSingleNode("Application/forms/action/form/@uid").Value = "";
            oXmlDoc.SelectSingleNode("Application/forms/action/form/@top").Value = Convert.ToString(Convert.ToInt16(oXmlDoc.SelectSingleNode("Application/forms/action/form/@top").Value) + (ZSBO.SBO_Application.Forms.Count + 5));
            oXmlDoc.SelectSingleNode("Application/forms/action/form/@left").Value = Convert.ToString(Convert.ToInt16(oXmlDoc.SelectSingleNode("Application/forms/action/form/@left").Value) + (ZSBO.SBO_Application.Forms.Count + 5));

            // add a new form
            oCreationParams = ((SAPbouiCOM.FormCreationParams)(ZSBO.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams)));
            oCreationParams.XmlData = oXmlDoc.InnerXml.ToString();

            oForm = ZSBO.SBO_Application.Forms.AddEx(oCreationParams);
            oMAT1 = (SAPbouiCOM.Matrix)oForm.Items.Item("MAT1").Specific;
            oMAT2 = (SAPbouiCOM.Matrix)oForm.Items.Item("MAT2").Specific;
            oDBDSH = oForm.DataSources.DBDataSources.Item("@ZMC112H");
            oDBDSL = oForm.DataSources.DBDataSources.Item("@ZMC112L");
            oDBDSM = oForm.DataSources.DBDataSources.Item("@ZMC112M");

            oForm.Title = ZSBO.ZAPI.FGetFormTitle(oForm.Title, FileName, oForm.TypeEx, oForm.BusinessObject.Type);
            oForm.DataBrowser.BrowseBy = "DOCENTRY";

            oForm.Visible = true;
            oForm.Freeze(true);

            CreateItems();
            SetChooseFromList("", "", 0);

            if (oString != "")
            {
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                oForm.Items.Item("DOCENTRY").Enabled = true;
                ((SAPbouiCOM.EditText)oForm.Items.Item("DOCENTRY").Specific).Value = oString;
                oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
            }
            else
            {
                SetFormBase();
                SetItemEnabled();
                SetToolBarEnabled();
                SetFocusByFormMode();
            }

            // 틀고정
            oMAT1.CommonSetting.FixedColumnsCount = 2;
            oMAT2.CommonSetting.FixedColumnsCount = 2;

            oForm.Freeze(false);
        }
        //-------------------------------------------------------------------------------------------------------
        catch (Exception Error)
        {
            oForm.Freeze(false);
            ZSBO.SBO_Application.SetStatusBarMessage(Error.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        }
    }

    //************************************************************************************************************
    // CreateItems
    //************************************************************************************************************
    private void CreateItems()
    {
        SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)ZSBO.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        try
        {
            // 입출고유형
            oRecordSet.DoQuery("SELECT U_VALUE, U_VALNAME FROM [@ZPGMD] WHERE CODE = 'ZMC112' AND U_TABLE = 'IOTYPE' AND U_VALID = 'Y' ORDER BY LineId");   // 입출고유형 환경변수 세팅변경 (2025.07.01.KDI)
            while (oRecordSet.EoF == false)
            {
                ((SAPbouiCOM.ComboBox)oForm.Items.Item("IOTYPE").Specific).ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim());
                oRecordSet.MoveNext();
            }
            
            // 첫번째 행 지정
            if (((SAPbouiCOM.ComboBox)oForm.Items.Item("IOTYPE").Specific).ValidValues.Count > 0)
                ((SAPbouiCOM.ComboBox)oForm.Items.Item("IOTYPE").Specific).Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

            // 문서상태
            oRecordSet.DoQuery("SELECT U_VALUE, U_VALNAME FROM [@ZPGMD] WHERE U_TABLE = 'DOCSTAT' AND Code = 'ZMC112'");
            while (oRecordSet.EoF == false)
            {
                ((SAPbouiCOM.ComboBox)oForm.Items.Item("DOCSTAT").Specific).ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim());
                oRecordSet.MoveNext();
            }
        }
        //-------------------------------------------------------------------------------------------------------
        catch (Exception Error)
        {
            ZSBO.SBO_Application.SetStatusBarMessage(Error.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        }
        finally
        {
            if (oRecordSet != null) oRecordSet = null;
        }
    }

    //************************************************************************************************************
    // SetItemEnabled     
    //************************************************************************************************************
    private void SetItemEnabled()
    {
        //2026.04.23.LKC. SetAutoManagedAttribute 사용하여 화면 제어시 CFL 정보가 삭제되는 현상이 발생되어 변경함
        //SAPbouiCOM.BoAutoManagedAttr oEditable = SAPbouiCOM.BoAutoManagedAttr.ama_Editable;
        //SAPbouiCOM.BoAutoManagedAttr oVisible = SAPbouiCOM.BoAutoManagedAttr.ama_Visible;
        //SAPbouiCOM.BoModeVisualBehavior oTrue = SAPbouiCOM.BoModeVisualBehavior.mvb_True;
        //SAPbouiCOM.BoModeVisualBehavior oFalse = SAPbouiCOM.BoModeVisualBehavior.mvb_False;

        string[] oArray;
        oForm.Freeze(true);
        try
        {
            string docStat = "";
            string canceled = "";
            string diNum = "";

            try { docStat  = oDBDSH.GetValue("U_DOCSTAT", 0).ToString().Trim(); } catch { }
            try { canceled = oDBDSH.GetValue("CANCELED", 0).ToString().Trim();  } catch { }
            try { diNum    = oDBDSH.GetValue("U_DINUM", 0).ToString().Trim();   } catch { }

            bool isDocOpen  = (docStat  == "O");
            bool isCanceled = (canceled == "Y");
            bool hasDiNum   = (diNum    != "" );
            
            // 문서가 닫힘(C)이거나, DI처리가 완료되었거나, 취소(Y)된 경우 닫힌 문서로 간주
            bool isDocClosed = (docStat == "C" || isCanceled || hasDiNum);

            bool isFind = (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE);
            bool isAdd  = (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE );

            oArray = new string[] { "DOCENTRY", "CANCELED", "DOCSTAT", "DINUM", "DINUM2", "DICNUM", "DICNUM2" };
            foreach (string oStr in oArray)
            {
                try { oForm.Items.Item(oStr).Enabled = isFind; } catch { }
            }

            // 2026.09.04.LKC. MAT1/MAT2는 Enabled=false 시 CFL('...' 버튼) 바인딩이 파괴되므로 항상 활성 유지
            // 문서 보호는 SetToolBarEnabled()에서 저장/행추가/행삭제 메뉴 비활성화로 처리
            oArray = new string[] { "MAT1", "MAT2" };
            foreach (string oStr in oArray)
            {
                try 
                { 
                    oForm.Items.Item(oStr).Enabled = !isFind;
                    if (!isDocOpen || isCanceled || hasDiNum)
                    {
                        oForm.Items.Item(oStr).Enabled = false;
                    }
                }
                catch { }
            }

            //oArray = new string[] { "MAT1", "MAT2" };
            //foreach (string oStr in oArray)
            //{
            //    oForm.Items.Item(oStr).SetAutoManagedAttribute(oEditable, (int)SAPbouiCOM.BoAutoFormMode.afm_Add, oTrue);
            //    oForm.Items.Item(oStr).SetAutoManagedAttribute(oEditable, (int)SAPbouiCOM.BoAutoFormMode.afm_Find, oFalse);
            //    oForm.Items.Item(oStr).SetAutoManagedAttribute(oEditable, (int)SAPbouiCOM.BoAutoFormMode.afm_Ok, oTrue);
            //    oForm.Items.Item(oStr).SetAutoManagedAttribute(oEditable, (int)SAPbouiCOM.BoAutoFormMode.afm_View, oTrue);
            //}

            // 2026.09.04.LKC. Matrix 컬럼의 CFL('...' 버튼) 바인딩 재설정
            // SAP B1 UI API에서 모드 전환/저장/취소 시 Matrix 컬럼의 ChooseFromListUID 바인딩이 해제될 수 있으므로 명시적으로 재설정
            if (!isFind)
            {
                ResetMatrixCFL();
            }

            // BTN_BAT는 닫힌 문서에서 비활성화
            try 
            { 
                if (isFind)
                    oForm.Items.Item("BTN_BAT").Enabled = false;
                else
                    oForm.Items.Item("BTN_BAT").Enabled = !isDocClosed;
            } 
            catch { }

            oArray = new string[] { "IOTYPE", "DOCDATE", "ITEMCODE", "ITEMNAME", "TITEMCD", "TITEMNM" };
            foreach (string oStr in oArray)
            {
                try
                {
                    if (isFind)
                        oForm.Items.Item(oStr).Enabled = true;
                    else
                        oForm.Items.Item(oStr).Enabled = !isDocClosed;
                }
                catch { }
            }

            //////// 2026.09.04.LKC. 헤더 품번/대체품번 CFL('...' 버튼) 항목도 Matrix와 동일하게
            //////// Enabled=false 시 ChooseFromListUID 바인딩이 파괴되므로 항상 활성 유지 후 명시적으로 재설정
            //////// 문서 보호는 SetToolBarEnabled()/MakeDI()의 재처리 방지 로직으로 처리
            //////oArray = new string[] { "ITEMCODE", "ITEMNAME", "TITEMCD", "TITEMNM" };
            //////foreach (string oStr in oArray)
            //////{
            //////    //try { oForm.Items.Item(oStr).Enabled = isFind; }
            //////    try { oForm.Items.Item(oStr).Enabled = isFind; }
            //////    catch { }
            //////}
            if (!isFind)
            {
                ResetHeaderCFL();
            }
        }
        catch (Exception Error)
        {
            ZSBO.SBO_Application.SetStatusBarMessage("SetItemEnabled Error : " + Error.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        }
        finally
        {
            oForm.Freeze(false);
        }
    }

    //************************************************************************************************************
    // SetToolBarEnabled
    //************************************************************************************************************
    private void SetToolBarEnabled()
    {
        oForm.Freeze(true);

        oForm.EnableMenu("1283", false);    // 제거
        oForm.EnableMenu("1284", false);    // 취소
        oForm.EnableMenu("1285", false);    // 복원
        oForm.EnableMenu("1286", false);    // 닫기
        oForm.EnableMenu("1287", false);    // 복제

        // 취소됨
        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE & oDBDSH.GetValue("CANCELED", 0) == "N")
        {
            oForm.EnableMenu("1284", true);    // 취소
        }

        //행삭제, 행추가
        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
        {
            oForm.EnableMenu("1292", true);    // 라인추가
            oForm.EnableMenu("1293", true);    // 라인삭제
            oForm.EnableMenu("1294", true);    // 행복제
        }

        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
        {
            oForm.EnableMenu("1292", true);    // 라인추가
            oForm.EnableMenu("1293", true);    // 라인삭제
            //oForm.EnableMenu("1287", true);    // 복제
            oForm.EnableMenu("1294", true);   // 행복제
        }

        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
        {
            if (oDBDSH.GetValue("CANCELED", 0) == "Y")
            {
                oForm.EnableMenu("1284", false);    // 취소
                oForm.EnableMenu("1292", false);    // 라인추가
                oForm.EnableMenu("1293", false);    // 라인삭제
                oForm.EnableMenu("1294", false);    // 행복제
                //oMAT1.Item.Enabled = false;
                //oMAT2.Item.Enabled = false;
                oForm.Items.Item("DOCENTRY").Enabled = false;
            }
        }

        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
        {
            if (oDBDSH.GetValue("U_DOCSTAT", 0) != "O")
            {
                oForm.EnableMenu("1292", false);    // 라인추가
                oForm.EnableMenu("1293", false);    // 라인삭제
                oForm.EnableMenu("1294", false);    // 행복제
                //oMAT1.Item.Enabled = false;
                //oMAT2.Item.Enabled = false;
                oForm.Items.Item("DOCENTRY").Enabled = false;
            }
        }

        oForm.Freeze(false);
    }

    //************************************************************************************************************
    // [Main Form - SetFocusByFormMode] : Form Mode 에 따른 Focus 위치 설정
    //************************************************************************************************************
    private void SetFocusByFormMode()
    {
        //2024.11.06  화면제어관련 수정
        if (CheckIsNull(oDBDSH.GetValue("U_DOCSTAT", 0)) != "O" || CheckIsNull(oDBDSH.GetValue("CANCELED", 0)) != "Y")
        {
            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
            {
                //라인 첫번째 행과 칼럼에 위치
                oMAT1.Columns.Item("ITEMCODE").Cells.Item(1).Click(SAPbouiCOM.BoCellClickType.ct_Regular, 1);
            }
        }
    }

    //************************************************************************************************************
    // 선택리스트 셋팅
    //************************************************************************************************************
    private void SetChooseFromList(string sUID, string sValue, int sRow)
    {
        SAPbobsCOM.Recordset oRecordset = (SAPbobsCOM.Recordset)ZSBO.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        SAPbouiCOM.Conditions oConds;
        string oCondVal1 = "", oCondVal2 = "", oCondVal3 = "";
        
        var oBinCodesList = new List<string>();
        string[] oBinCodes;
        string binCode, oItemCode, oWhsCode;

        if (sUID == "ITEMCODE" || sUID == "ITEMNAME" || CheckIsNull(sUID) == "")
        {
            //품목 조건 설정
            oConds = (SAPbouiCOM.Conditions)ZSBO.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_Conditions);
            //상품
            oConds.Add();
            oConds.Item(oConds.Count - 1).BracketOpenNum = 1;
            oConds.Item(oConds.Count - 1).Alias = "ItmsGrpCod";
            oConds.Item(oConds.Count - 1).Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
            oConds.Item(oConds.Count - 1).CondVal = "100";
            oConds.Item(oConds.Count - 1).BracketCloseNum = 1;

            //And
            oConds.Item(oConds.Count - 1).Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR;

            //원자재
            oConds.Add();
            oConds.Item(oConds.Count - 1).BracketOpenNum = 1;
            oConds.Item(oConds.Count - 1).Alias = "ItmsGrpCod";
            oConds.Item(oConds.Count - 1).Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
            oConds.Item(oConds.Count - 1).CondVal = "103";
            oConds.Item(oConds.Count - 1).BracketCloseNum = 1;

            oForm.ChooseFromLists.Item("HITEMCD").SetConditions(oConds);
            oForm.ChooseFromLists.Item("HITEMNM").SetConditions(oConds);
            oForm.ChooseFromLists.Item("HTITEMCD").SetConditions(oConds);
            oForm.ChooseFromLists.Item("HTITEMNM").SetConditions(oConds);
            oForm.ChooseFromLists.Item("ITEMCODE").SetConditions(oConds);
            oForm.ChooseFromLists.Item("ITEMNAME").SetConditions(oConds);
            oForm.ChooseFromLists.Item("TITEMCD").SetConditions(oConds);
            oForm.ChooseFromLists.Item("TITEMNM").SetConditions(oConds);
        }

        //if (sUID == "TITEMCD" || sUID == "TITEMNM" || CheckIsNull(sUID) == "")
        //{
        //    //품목 조건 설정
        //    oConds = (SAPbouiCOM.Conditions)ZSBO.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_Conditions);
        //    //상품
        //    oConds.Add();
        //    oConds.Item(oConds.Count - 1).BracketOpenNum = 1;
        //    oConds.Item(oConds.Count - 1).Alias = "ItmsGrpCod";
        //    oConds.Item(oConds.Count - 1).Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
        //    oConds.Item(oConds.Count - 1).CondVal = "100";
        //    oConds.Item(oConds.Count - 1).BracketCloseNum = 1;

        //    //And
        //    oConds.Item(oConds.Count - 1).Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR;

        //    //원자재
        //    oConds.Add();
        //    oConds.Item(oConds.Count - 1).BracketOpenNum = 1;
        //    oConds.Item(oConds.Count - 1).Alias = "ItmsGrpCod";
        //    oConds.Item(oConds.Count - 1).Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
        //    oConds.Item(oConds.Count - 1).CondVal = "103";
        //    oConds.Item(oConds.Count - 1).BracketCloseNum = 1;

        //    oForm.ChooseFromLists.Item("TITEMCD").SetConditions(oConds);
        //    oForm.ChooseFromLists.Item("TITEMNM").SetConditions(oConds);
        //}

        //BIN 로케이션 조건 설정
        if (sUID == "BINCODE") 
        {
            //if (CheckIsNull(sValue) == "") oCondVal2 = ""; else oCondVal2 = sValue;
            //oConds = (SAPbouiCOM.Conditions)ZSBO.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_Conditions);
            //oConds.Add();
            //oConds.Item(oConds.Count - 1).Alias = "WhsCode";
            //oConds.Item(oConds.Count - 1).Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
            //oConds.Item(oConds.Count - 1).CondVal = oCondVal2;

            //oForm.ChooseFromLists.Item("BINCODE").SetConditions(oConds);
            ////oForm.ChooseFromLists.Item("BINNAME").SetConditions(oConds);

            oItemCode = oDBDSL.GetValue("U_ITEMCODE", sRow).ToString().Trim(); //2026.03.11.LKC. 품목/창고 감안하여 반영함
            oWhsCode = sValue; // oDBDSL.GetValue("U_WHSCODE", sRow).ToString().Trim();
            oConds = (SAPbouiCOM.Conditions)ZSBO.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_Conditions);

            oRecordset.DoQuery("EXEC [DBO].[ZTMSP_ZMC112_BIN] @TYPE     = 'WO' " 
                                                        + " , @ITEMCODE = '" + oItemCode + "', @WHSCODE  = '" + oWhsCode + "' ");
            while (oRecordset.EoF == false)
            {
                //같은 값은 제외함
                binCode = oRecordset.Fields.Item(0).Value.ToString();
                if (!oBinCodesList.Contains(binCode)) oBinCodesList.Add(binCode); //중복 제외
                oRecordset.MoveNext();
            }

            if (oBinCodesList.Count > 0)
            {
                oBinCodes = oBinCodesList.ToArray();
                //oConds.Item(oConds.Count - 1).Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;

                for (int i = 0; i < oBinCodes.Length; i++)
                {
                    oConds.Add();
                    if (i == 0) oConds.Item(oConds.Count - 1).BracketOpenNum = 1;
                    oConds.Item(oConds.Count - 1).Alias = "BinCode";
                    oConds.Item(oConds.Count - 1).Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                    oConds.Item(oConds.Count - 1).CondVal = oBinCodes[i];

                    if (i != oBinCodes.Length - 1)
                        oConds.Item(oConds.Count - 1).Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR;

                    if (i == oBinCodes.Length - 1)
                        oConds.Item(oConds.Count - 1).BracketCloseNum = 1;
                }
            }
            else
            {
                oConds.Add();
                oConds.Item(oConds.Count - 1).Alias = "BinCode";
                oConds.Item(oConds.Count - 1).Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                oConds.Item(oConds.Count - 1).CondVal = "";
            }

            oForm.ChooseFromLists.Item("BINCODE").SetConditions(oConds);
        }

        if (sUID == "TBINCODE") //입고BIN은 창고 기준으로 표시하게 수정 2026.03.13.LKC
        {
            if (CheckIsNull(sValue) == "") oCondVal3 = ""; else oCondVal3 = sValue;

            oConds = (SAPbouiCOM.Conditions)ZSBO.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_Conditions);
            oConds.Add();
            oConds.Item(oConds.Count - 1).Alias = "WhsCode";
            oConds.Item(oConds.Count - 1).Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
            oConds.Item(oConds.Count - 1).CondVal = oCondVal3;

            oForm.ChooseFromLists.Item("TBINCODE").SetConditions(oConds);
        }
    }

    //************************************************************************************************************
    // [Main Form] : ResetMatrixCFL - Matrix 컬럼의 CFL('...' 버튼) 바인딩 재설정 2026.09.04.LKC
    // SAP B1 UI API에서 폼 모드 전환/저장/취소/Navigation 시 Matrix 컬럼의 ChooseFromListUID 바인딩이
    // 내부적으로 해제되어 '...' 버튼이 사라지는 현상을 방지하기 위해 명시적으로 재설정함
    //************************************************************************************************************
    private void ResetMatrixCFL()
    {
        // MAT1 컬럼 CFL 바인딩 (ITEMCODE는 LinkedButton(type=116)이므로 제외)
        SetColumnCFL(oMAT1, "ITEMNAME", "ITEMNAME", "ItemName");
        SetColumnCFL(oMAT1, "WHSCODE",  "WHSCODE",  "WhsCode");
        SetColumnCFL(oMAT1, "BINCODE",  "BINCODE",  "BinCode");

        // MAT2 컬럼 CFL 바인딩 (ITEMCODE는 LinkedButton(type=116)이므로 제외)
        SetColumnCFL(oMAT2, "ITEMNAME", "TITEMNM",  "ItemName");
        SetColumnCFL(oMAT2, "WHSCODE",  "TWHSCODE", "WhsCode");
        SetColumnCFL(oMAT2, "BINCODE",  "TBINCODE", "BinCode");
    }

    //************************************************************************************************************
    // [Main Form] : ResetHeaderCFL - 헤더 품번/대체품번 항목의 CFL('...' 버튼) 바인딩 재설정 2026.09.04.LKC
    // ResetMatrixCFL()과 동일한 이유(Enabled=false 시 ChooseFromListUID 바인딩 파괴)로,
    // 헤더의 품번/품명/대체품번/대체품명 EditText 항목도 재활성화 시 명시적으로 재설정함
    //************************************************************************************************************
    private void ResetHeaderCFL()
    {
        SetEditTextCFL("ITEMCODE", "HITEMCD",  "ItemCode");
        SetEditTextCFL("ITEMNAME", "HITEMNM",  "ItemName");
        SetEditTextCFL("TITEMCD",  "HTITEMCD", "ItemCode");
        SetEditTextCFL("TITEMNM",  "HTITEMNM", "ItemName");
    }

    // 헤더 EditText 항목의 ChooseFromListUID/Alias 설정
    private void SetEditTextCFL(string sItemUID, string cflUID, string cflAlias)
    {
        try
        {
            SAPbouiCOM.EditText oEditText = (SAPbouiCOM.EditText)oForm.Items.Item(sItemUID).Specific;
            oEditText.ChooseFromListUID = cflUID;
            oEditText.ChooseFromListAlias = cflAlias;
        }
        catch { }
    }

    // COM late binding으로 Matrix 컬럼의 ChooseFromListUID/Alias 설정
    private void SetColumnCFL(SAPbouiCOM.Matrix oMat, string colUID, string cflUID, string cflAlias)
    {
        try
        {
            object col = oMat.Columns.Item(colUID);
            System.Reflection.BindingFlags setFlag = System.Reflection.BindingFlags.SetProperty;
            col.GetType().InvokeMember("ChooseFromListUID",   setFlag, null, col, new object[] { cflUID });
            col.GetType().InvokeMember("ChooseFromListAlias", setFlag, null, col, new object[] { cflAlias });
        }
        catch { }
    }

    //*******************************************************************************************
    // [참조 Form] : ZMSLIST_LoadFormA
    //*******************************************************************************************
    public void ZMSLIST_LoadFormA(ref SAPbouiCOM.DBDataSource rDBDSH, ref SAPbouiCOM.DBDataSource rDBDSL, ref SAPbouiCOM.Matrix rMat)
    {
        try
        {
            SAPbouiCOM.FormCreationParams oCreationParams = null;
            System.Xml.XmlDocument oXmlDoc = null;
            string sPath = null;
            oXmlDoc = new System.Xml.XmlDocument();
            sPath = Application.StartupPath;

            //GRID 화면 표시
            oXmlDoc.Load(sPath + @"\Screen\ZMC112A.srf");
            oXmlDoc.SelectSingleNode("Application/forms/action/form/@uid").Value = "";
            oXmlDoc.SelectSingleNode("Application/forms/action/form/@top").Value
                = Convert.ToString(Convert.ToInt16(oXmlDoc.SelectSingleNode("Application/forms/action/form/@top").Value) + (ZSBO.SBO_Application.Forms.Count + 5));
            oXmlDoc.SelectSingleNode("Application/forms/action/form/@left").Value
                = Convert.ToString(Convert.ToInt16(oXmlDoc.SelectSingleNode("Application/forms/action/form/@left").Value) + (ZSBO.SBO_Application.Forms.Count + 5));

            // Add a new form
            oCreationParams = ((SAPbouiCOM.FormCreationParams)(ZSBO.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams)));
            oCreationParams.XmlData = oXmlDoc.InnerXml.ToString();

            aForm = ZSBO.SBO_Application.Forms.AddEx(oCreationParams);

            oGrd1 = (SAPbouiCOM.Grid)aForm.Items.Item("GRD1").Specific;
            aForm.DataSources.DataTables.Add("GRD1");

            //Form DataBind
            ZMSLIST_CreateItemsA();
            //부모 데이터소스 및 Matrix 참조
            //ZDBDSH = rDBDSH;
            ZDBDSL = rDBDSL;
            ZMAT1 = rMat;

            ZMSLIST_LoadDataA("B");

            //ChooseFromList 대상 칼럼 설정
            ZMSLIST_AddChooseFromListA();
            //CFL 조건 설정
            ZMSLIST_SetChooseFromListCondsA();

            aForm.Visible = true;

            //Modal Flag 설정
            aModal = true;
        }
        catch (Exception ex)
        {
            ZSBO.SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        }
    }

    //*******************************************************************************************
    // [참조 Form] : ZMSLIST_CreateItemsA(폼에 DataBind 및 기본값 설정)
    //*******************************************************************************************
    private void ZMSLIST_CreateItemsA()
    {
        //Adding a User Data Source
        SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)ZSBO.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        string qryStr = string.Empty;
        string[] oArray;

        try
        {
            // 헤더(조건) 설정
            oArray = new string[] { "ITEMCODE", "ITEMNAME", "WHSCODE", "WHSNAME", "BINCODE" };
            foreach (string oStr in oArray)
            {
                aForm.DataSources.UserDataSources.Add(oStr, SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 200);
                ((SAPbouiCOM.EditText)aForm.Items.Item(oStr).Specific).DataBind.SetBound(true, "", oStr);
            }
        }
        catch (Exception Error)
        {
            ZSBO.SBO_Application.SetStatusBarMessage(Error.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        }
        finally
        {
            if (oRecordSet != null) oRecordSet = null;
        }
    }

    //************************************************************************************************************
    // [참조 Form] : ZMSLIST_ItemEventA
    //************************************************************************************************************
    private void ZMSLIST_ItemEventA(ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
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
                    if (pVal.BeforeAction == true)
                    {
                        //품번대체 배치/시리얼선택 버튼
                        if (pVal.ItemUID == "BTN_INQ" || pVal.ItemUID == "BTN_SEL")
                        {
                            ZMSLIST_LoadDataA(pVal.ItemUID);
                            if (pVal.ItemUID == "BTN_SEL" && oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                            oMAT1.AutoResizeColumns();
                            BubbleEvent = false;
                            return;
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
                    oCFL = ((SAPbouiCOM.IChooseFromListEvent)(pVal));

                    //품목 버튼 처리
                    if (pVal.Before_Action == false & 
                        (oCFL.ChooseFromListUID == "ITEMCODE" || oCFL.ChooseFromListUID == "ITEMNAME" ||
                         oCFL.ChooseFromListUID == "WHSCODE"  || oCFL.ChooseFromListUID == "WHSNAME"))
                    {
                        if (oCFL.SelectedObjects == null) return;
                        ZMSLIST_SetChooseFromListA(oCFL.SelectedObjects, pVal, oCFL.ChooseFromListUID);
                    }

                    //////////창고 코드


                    //if (pVal.Before_Action == false & (oCFL.ChooseFromListUID == "BINCODE" || oCFL.ChooseFromListUID == "TBINCODE"))
                    if (oCFL.ChooseFromListUID == "BINCODE")
                    {
                        //if (oCFL.SelectedObjects == null) return;
                        //ZMSLIST_SetChooseFromList(pVal);

                        if (oCFL.ChooseFromListUID == "BINCODE")
                        {
                            ZMSLIST_SetChooseFromListA(oCFL.SelectedObjects, pVal, oCFL.ChooseFromListUID);
                            //SetChooseFromList("BINCODE", oDBDSL.GetValue("U_WHSCODE", pVal.Row - 1).Trim(), 0);
                        }
                        //else if (oCFL.ChooseFromListUID == "TBINCODE")
                        //    ZMSLIST_SetChooseFromListA(oCFL.SelectedObjects, pVal, oCFL.ChooseFromListUID);
                        ////SetChooseFromList("TBINCODE", oDBDSL.GetValue("U_TWHSCODE", pVal.Row - 1).Trim(), 0);                            

                        if (oCFL.SelectedObjects == null) return;
                    }
                    break;

                //========================================================================================================
                case SAPbouiCOM.BoEventTypes.et_CLICK:
                //========================================================================================================
                    if (pVal.BeforeAction == false & pVal.ItemUID == "GRD1" & pVal.ColUID != "RowsHeader" & pVal.Row >= 0)
                    {
                        oGrd1.Rows.SelectedRows.Clear();
                        oGrd1.Rows.SelectedRows.Add(pVal.Row);
                    }
                    break;

                //========================================================================================================
                case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK:
                //========================================================================================================
                    if (pVal.BeforeAction && pVal.ItemUID == "GRD1" && pVal.Row < 0 && pVal.ColUID == "선택")
                    {
                        aForm.Freeze(true);
                        string chkYN = oGrd1.DataTable.Columns.Item("선택").Cells.Item(0).Value.ToString();
                        for (int i = 0; i <= oGrd1.Rows.Count - 1; i++)
                        {
                            //선택 칼럼
                            if (chkYN == "" | chkYN == "N")
                                oGrd1.DataTable.Columns.Item(0).Cells.Item(i).Value = "Y";
                            else
                                oGrd1.DataTable.Columns.Item(0).Cells.Item(i).Value = "N";
                        }
                        aForm.Freeze(false);
                    }
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
                                oGrd1.AutoResizeColumns();
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
            if (pVal.BeforeAction == true) BubbleEvent = false;
            if (aForm != null)
            {
                aForm.Close(); aForm = null;
            }
            ZSBO.SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        }
    }

    //************************************************************************************************************
    // [참조 Form] : ZMSLIST_SetChooseFromListA
    //************************************************************************************************************
    private void ZMSLIST_SetChooseFromListA(SAPbouiCOM.DataTable oSelTable, SAPbouiCOM.ItemEvent pVal, string oItemUID)
    {
        SAPbobsCOM.Recordset oRecordset = (SAPbobsCOM.Recordset)ZSBO.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        SAPbouiCOM.IChooseFromListEvent oCFL = ((SAPbouiCOM.IChooseFromListEvent)(pVal));
        SAPbouiCOM.Conditions oConds;
        int row = pVal.Row;
        string oQry = String.Empty, binCode;
        var oBinCodesList = new List<string>();
        string[] oBinCodes;
        string oSelBin = string.Empty, oItemCode, oWhsCode;

        try
        {
            if (pVal.Before_Action) //선택 전
            {
                switch (pVal.ItemUID + pVal.ColUID)
                {
                    case "GRD1출고BIN":
                        oConds = (SAPbouiCOM.Conditions)ZSBO.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_Conditions);
                        ////창고 조건
                        ////oConds.Add();
                        ////oConds.Item(oConds.Count - 1).Alias = "WhsCode";
                        ////oConds.Item(oConds.Count - 1).Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                        ////oConds.Item(oConds.Count - 1).CondVal = oGrd1.DataTable.GetValue(oGrd1.Columns.Item("출고창고").UniqueID, pVal.Row).ToString();

                        ////출고BIN 가져오기(2개 이상인 경우 빈칸 후, GRID에서 선택할 수 있게 설정)
                        //oQry = " SELECT T3.[BinCode] AS BINCODE                                    "
                        //     + "   FROM OIGN T0 INNER JOIN IGN1 T1 ON T0.DocEntry   = T1.DocEntry  "
                        //     + "                INNER JOIN OILM T4 ON T4.DocEntry   = T1.DocEntry  "
                        //     + "                                  AND T4.DocLineNum = T1.LineNum   "
                        //     + "                INNER JOIN OBTL T2 ON T4.MessageID  = T2.MessageID "
                        //     + "                INNER JOIN OBIN T3 ON T2.BinAbs     = T3.AbsEntry  "
                        //     + "                INNER JOIN OIBQ T5 ON T5.BinAbs     = T2.BinAbs    " // 실시간 빈 재고 테이블 조인
                        //     + "                                  AND T5.ItemCode   = T1.ItemCode  "
                        //     + "                                  AND T5.WhsCode    = T1.WhsCode   "
                        //     + "  WHERE T1.[ItemCode] = '" + oGrd1.DataTable.GetValue(oGrd1.Columns.Item("품번").UniqueID, pVal.Row).ToString() + "' "
                        //     + "    AND T1.[WhsCode]  = '" + oGrd1.DataTable.GetValue(oGrd1.Columns.Item("출고창고").UniqueID, pVal.Row).ToString() + "' "
                        //     + "    AND T4.[TransType] = 59                                        "
                        //     + "    AND T5.[OnHandQty] > 0                                         " // 2026.03.10.LKC. 재고수량 체크 추가
                        //     + "  ORDER BY T0.[DocNum] DESC, T1.[LineNum] ASC                      ";

                        //oRecordset.DoQuery(oQry);

                        oRecordset.DoQuery("EXEC [DBO].[ZTMSP_ZMC112_BIN] @TYPE     = 'WO' "
                                                                    + " , @ITEMCODE = '" + oGrd1.DataTable.GetValue(oGrd1.Columns.Item("품번").UniqueID, pVal.Row).ToString()
                                                                    + "', @WHSCODE  = '" + oGrd1.DataTable.GetValue(oGrd1.Columns.Item("출고창고").UniqueID, pVal.Row).ToString() + "' ");
                        while (oRecordset.EoF == false)
                        {
                            //같은 값은 제외함
                            binCode = oRecordset.Fields.Item(0).Value.ToString();
                            if (!oBinCodesList.Contains(binCode)) oBinCodesList.Add(binCode); //중복 제외
                            oRecordset.MoveNext();
                        }

                        if (oBinCodesList.Count > 0)
                        {
                            oBinCodes = oBinCodesList.ToArray();
                            //oConds.Item(oConds.Count - 1).Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;

                            for (int i = 0; i < oBinCodes.Length; i++)
                            {
                                oConds.Add();
                                if (i == 0) oConds.Item(oConds.Count - 1).BracketOpenNum = 1;
                                oConds.Item(oConds.Count - 1).Alias = "BinCode";
                                oConds.Item(oConds.Count - 1).Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                                oConds.Item(oConds.Count - 1).CondVal = oBinCodes[i];

                                if (i != oBinCodes.Length - 1)
                                    oConds.Item(oConds.Count - 1).Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR;

                                if (i == oBinCodes.Length - 1)
                                    oConds.Item(oConds.Count - 1).BracketCloseNum = 1;
                            }
                        }
                        else
                        {
                            oConds.Add();
                            oConds.Item(oConds.Count - 1).Alias = "BinCode";
                            oConds.Item(oConds.Count - 1).Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                            oConds.Item(oConds.Count - 1).CondVal = "";
                        }

                        aForm.ChooseFromLists.Item("BINCODE").SetConditions(oConds);
                        break;

                    //case "GRD1입고BIN":
                    //    oConds = (SAPbouiCOM.Conditions)ZSBO.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_Conditions);
                    //    //창고 기준으로 입고BIN 조건 설정함 2026.03.13.LKC.
                    //    oConds.Add();
                    //    oConds.Item(oConds.Count - 1).Alias = "WhsCode";
                    //    oConds.Item(oConds.Count - 1).Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                    //    oConds.Item(oConds.Count - 1).CondVal = oGrd1.DataTable.GetValue(oGrd1.Columns.Item("입고창고").UniqueID, pVal.Row).ToString();

                    //    //2026.03.13.LKC. 입고BIN은 창고 기준으로 표시하기로 함
                    //    //oRecordset.DoQuery("EXEC [DBO].[ZTMSP_ZMC112_BIN] @TYPE     = 'I' " //2026.03.11.LKC. 쿼리문에서 프로시저로 변경
                    //    //                                            + " , @ITEMCODE = '" + oGrd1.DataTable.GetValue(oGrd1.Columns.Item("대체품번").UniqueID, pVal.Row).ToString()
                    //    //                                            + "', @WHSCODE  = '" + oGrd1.DataTable.GetValue(oGrd1.Columns.Item("입고창고").UniqueID, pVal.Row).ToString() + "' ");

                    //    //while (oRecordset.EoF == false)
                    //    //{
                    //    //    //같은 값은 제외함
                    //    //    binCode = oRecordset.Fields.Item(0).Value.ToString();
                    //    //    if (!oBinCodesList.Contains(binCode)) oBinCodesList.Add(binCode); //중복 제외
                    //    //    oRecordset.MoveNext();
                    //    //}

                    //    //if (oBinCodesList.Count > 0)
                    //    //{
                    //    //    oBinCodes = oBinCodesList.ToArray();
                    //    //    //oConds.Item(oConds.Count - 1).Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;

                    //    //    for (int i = 0; i < oBinCodes.Length; i++)
                    //    //    {
                    //    //        oConds.Add();
                    //    //        if (i == 0) oConds.Item(oConds.Count - 1).BracketOpenNum = 1;
                    //    //        oConds.Item(oConds.Count - 1).Alias = "BinCode";
                    //    //        oConds.Item(oConds.Count - 1).Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                    //    //        oConds.Item(oConds.Count - 1).CondVal = oBinCodes[i];

                    //    //        if (i != oBinCodes.Length - 1)
                    //    //            oConds.Item(oConds.Count - 1).Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR;

                    //    //        if (i == oBinCodes.Length - 1)
                    //    //            oConds.Item(oConds.Count - 1).BracketCloseNum = 1;
                    //    //    }
                    //    //}
                    //    //else
                    //    //{
                    //    //    oConds.Add();
                    //    //    oConds.Item(oConds.Count - 1).Alias = "BinCode";
                    //    //    oConds.Item(oConds.Count - 1).Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                    //    //    oConds.Item(oConds.Count - 1).CondVal = "";
                    //    //}

                    //    aForm.ChooseFromLists.Item("TBINCODE").SetConditions(oConds);
                    //    break;
                }
            }

            if (!pVal.Before_Action) //선택 후
            {
                switch (pVal.ItemUID)
                {
                    case "ITEMCODE":
                    case "ITEMNAME":
                        for (int i = 0; i < oCFL.SelectedObjects.Rows.Count; i++)
                        {
                            ZMSLIST_FlushToItemValueA(pVal.ItemUID, "ITEMCODE", row, oCFL.SelectedObjects.GetValue("ItemCode", i).ToString());
                            row = row + 1;
                        }
                        break;

                    case "WHSCODE":
                    case "WHSNAME":
                        for (int i = 0; i < oCFL.SelectedObjects.Rows.Count; i++)
                        {
                            ZMSLIST_FlushToItemValueA(pVal.ItemUID, "WHSCODE", row, oCFL.SelectedObjects.GetValue("WhsCode", i).ToString());
                            row = row + 1;
                        }
                        break;
                }

                switch (pVal.ItemUID + pVal.ColUID)
                {
                    case "GRD1출고BIN":
                        if (oSelTable == null || oSelTable.Rows.Count == 0)
                            oSelBin = "";
                        else
                            oSelBin = CheckIsNull(oSelTable.GetValue("BinCode", 0));

                        //oMAT1.GetLineData(oRow);
                        //oDBDSL.Offset = oRow - 1;
                        //oDBDSL.SetValue("U_TITEMCD", oRow - 1, oSelTable.GetValue("ItemCode", 0).ToString().Trim());
                        //oDBDSL.SetValue("U_TITEMNM", oRow - 1, oSelTable.GetValue("ItemName", 0).ToString().Trim());
                        //oDBDSL.SetValue("U_TITEMNM1", oRow - 1, oSelTable.GetValue("FrgnName", 0).ToString().Trim());
                        //oMAT1.SetLineData(oRow);

                        oGrd1.DataTable.SetValue("출고BIN", pVal.Row, oSelBin); //oSelTable.GetValue("BinCode", 0).ToString().Trim());
                        //BIN로케이션 재고 표시
                        oItemCode = oGrd1.DataTable.GetValue("품번", pVal.Row).ToString();
                        oWhsCode = oGrd1.DataTable.GetValue("출고창고", pVal.Row).ToString();
                        oGrd1.DataTable.SetValue("BIN재고",   pVal.Row, GetOnHand("B", oItemCode, oWhsCode, oSelBin).ToString());
                        oGrd1.DataTable.SetValue("창고재고", pVal.Row, GetOnHand("W", oItemCode, oWhsCode, oSelBin).ToString());
                        break;

                    //case "GRD1입고BIN":
                    //    oGrd1.DataTable.SetValue("입고BIN", pVal.Row, oSelBin); //oSelTable.GetValue("BinCode", 0).ToString().Trim());
                    //    break;
                }
            }
        }
        catch (Exception Error)
        {
            ZSBO.SBO_Application.SetStatusBarMessage(Error.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        }
    }

    //************************************************************************************************************
    // [참조 Form] : ZMSLIST_FlushToItemValueA(validate, combo_box, chooseFromList 등 사용)
    //************************************************************************************************************
    private void ZMSLIST_FlushToItemValueA(string sUID, string sCol, int sRow, string sValue)
    {
        oForm.Freeze(true);
        SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)ZSBO.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        try
        {
            switch (sUID)
            {
                case "ITEMCODE":
                    if (sValue == "")
                    {
                        aForm.DataSources.UserDataSources.Item("ITEMCODE").ValueEx = "";
                        aForm.DataSources.UserDataSources.Item("ITEMNAME").ValueEx = "";
                    }
                    else
                    {
                        oRecordSet.DoQuery("SELECT ItemCode, ItemName FROM OITM WHERE ItemCode = '" + sValue + "'");
                        aForm.DataSources.UserDataSources.Item("ITEMCODE").ValueEx = oRecordSet.Fields.Item(0).Value.ToString();
                        aForm.DataSources.UserDataSources.Item("ITEMNAME").ValueEx = oRecordSet.Fields.Item(1).Value.ToString();
                    }
                    aForm.Items.Item("ITEMCODE").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    break;

                case "ITEMNAME":
                    if (sValue == "")
                    {
                        aForm.DataSources.UserDataSources.Item("ITEMCODE").ValueEx = "";
                        aForm.DataSources.UserDataSources.Item("ITEMNAME").ValueEx = "";
                    }
                    else
                    {
                        oRecordSet.DoQuery("SELECT ItemCode, ItemName FROM OITM WHERE ItemCode = '" + sValue + "'");
                        aForm.DataSources.UserDataSources.Item("ITEMCODE").ValueEx = oRecordSet.Fields.Item(0).Value.ToString();
                        aForm.DataSources.UserDataSources.Item("ITEMNAME").ValueEx = oRecordSet.Fields.Item(1).Value.ToString();
                    }
                    aForm.Items.Item("ITEMNAME").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    break;

                case "WHSCODE":
                    if (sValue == "")
                    {
                        aForm.DataSources.UserDataSources.Item("WHSCODE").ValueEx = "";
                        aForm.DataSources.UserDataSources.Item("WHSNAME").ValueEx = "";
                    }
                    else
                    {
                        oRecordSet.DoQuery("SELECT WhsCode, WhsName FROM OWHS WHERE WhsCode = '" + sValue + "'");
                        aForm.DataSources.UserDataSources.Item("WHSCODE").ValueEx = oRecordSet.Fields.Item(0).Value.ToString();
                        aForm.DataSources.UserDataSources.Item("WHSNAME").ValueEx = oRecordSet.Fields.Item(1).Value.ToString();
                    }
                    aForm.Items.Item("WHSCODE").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    break;

                case "WHSNAME":
                    if (sValue == "")
                    {
                        aForm.DataSources.UserDataSources.Item("WHSCODE").ValueEx = "";
                        aForm.DataSources.UserDataSources.Item("WHSNAME").ValueEx = "";
                    }
                    else
                    {
                        oRecordSet.DoQuery("SELECT WhsCode, WhsName FROM OWHS WHERE WhsCode = '" + sValue + "'");
                        aForm.DataSources.UserDataSources.Item("WHSCODE").ValueEx = oRecordSet.Fields.Item(0).Value.ToString();
                        aForm.DataSources.UserDataSources.Item("WHSNAME").ValueEx = oRecordSet.Fields.Item(1).Value.ToString();
                    }
                    aForm.Items.Item("WHSNAME").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    break;
            }
        }
        catch (Exception ex)
        {
            ZSBO.SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        }
        finally
        {
            oForm.Freeze(false);
        }
    }

    //************************************************************************************************************
    // [참조 Form] : ZMSLIST_LoadDataA
    //************************************************************************************************************
    private void ZMSLIST_LoadDataA(string value)
    {
        SAPbobsCOM.Recordset aRecordSet = (SAPbobsCOM.Recordset)ZSBO.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        string qString = "", oGRD = "GRD1", oItemCode, oWhsCode, oBinCode, sType;
        int i;
        string[] oArray;
        SAPbouiCOM.RowHeaders oHeader = null;
        SAPbouiCOM.EditTextColumn oCol;

        try
        {
            if (CheckIsNull(oDBDSH.GetValue("U_ITEMCODE", 0).Trim()) != "" &&
                CheckIsNull(aForm.DataSources.UserDataSources.Item("ITEMCODE").Value) == "")
            {
                aForm.DataSources.UserDataSources.Item("ITEMCODE").Value = oDBDSH.GetValue("U_ITEMCODE", 0).Trim();
                aForm.DataSources.UserDataSources.Item("ITEMNAME").Value = oDBDSH.GetValue("U_ITEMNAME", 0).Trim();
            }

            oItemCode = aForm.DataSources.UserDataSources.Item("ITEMCODE").ValueEx;		// 품목
            oWhsCode  = aForm.DataSources.UserDataSources.Item("WHSCODE").ValueEx;      // 창고코드
            oBinCode  = aForm.DataSources.UserDataSources.Item("BINCODE").ValueEx;      // BIN위치

            //if (CheckIsNull(oItemCode.ToString().Trim()) == "")
            //{
            //    aForm.Items.Item("ITEMCODE").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
            //    ZSBO.SBO_Application.StatusBar.SetText("품번이 누락되었습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            //    return;
            //}

            //oIOType = aForm.DataSources.UserDataSources.Item("IOTYPE").Value;           // 입출고유형
            ////oCardCode = aForm.DataSources.UserDataSources.Item("CARDCODE").Value;       // 거래처

            aForm.Freeze(true);
            if (value == "BTN_INQ" || value == "B")
            {
                if (value == "B")
                {
                    sType = "B";
                    aForm.DataSources.UserDataSources.Item("ITEMCODE").Value = CheckIsNull(oDBDSH.GetValue("U_ITEMCODE", 0));
                    aForm.DataSources.UserDataSources.Item("ITEMNAME").Value = CheckIsNull(oDBDSH.GetValue("U_ITEMNAME", 0));
                    oWhsCode = ""; oBinCode = "";
                }
                else
                {
                    sType = "S";
                }

                //품번대체요청 조회
                qString = "EXEC [dbo].[ZTMSP_ZMC112A]  @STYPE    = '" + sType + "' "
                                                  + ", @ITEMCODE = '" + oItemCode + "' "
                                                  + ", @WHSCODE  = '" + oWhsCode + "' "
                                                  + ", @BINCODE  = '" + oBinCode + "' ";

                aForm.DataSources.DataTables.Item(oGRD).ExecuteQuery(qString);
                oGrd1.DataTable = aForm.DataSources.DataTables.Item(oGRD);

                //if (value != "B")
                //{
                // 선택 버튼 제외 Grid 항목 비활성화
                for (int n = 1; n < oGrd1.Columns.Count; n++)
                {
                    oCol = (SAPbouiCOM.EditTextColumn)oGrd1.Columns.Item(n);
                    if (n == 4) //출고BIN인 경우 수정 가능하게 함
                        oCol.Editable = true;
                    else
                        oCol.Editable = false;
                    oCol.TitleObject.Sortable = true;
                }

                // 체크박스
                if (CheckIsNull(oGrd1.DataTable.GetValue(oGrd1.Columns.Item("품번").UniqueID, 0).ToString()) != "")
                {
                    oGrd1.Columns.Item("선택").Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox;
                    oGrd1.Columns.Item("선택").Editable = true;
                }
                else
                {
                    oGrd1.Columns.Item("선택").Editable = false;
                    oGrd1.Columns.Item("출고BIN").Editable = false;
                    //oGrd1.Columns.Item("입고BIN").Editable = false;
                }

                //LINK버튼
                oCol = (SAPbouiCOM.EditTextColumn)oGrd1.Columns.Item("품번"); oCol.LinkedObjectType = "4";
                //oCol = (SAPbouiCOM.EditTextColumn)oGrd1.Columns.Item("대체품번"); oCol.LinkedObjectType = "4";

                //숫자 칼럼 우측 정렬
                oGrd1.Columns.Item("BIN재고").RightJustified = true;
                oGrd1.Columns.Item("창고재고").RightJustified = true;
                oGrd1.Columns.Item("수량").RightJustified = true;
                //}
                //else
                //{
                //    // 선택 버튼 제외 Grid 항목 비활성화
                //    for (int n = 0; n < oGrd1.Columns.Count; n++)
                //    {
                //        oCol = (SAPbouiCOM.EditTextColumn)oGrd1.Columns.Item(n); oCol.Editable = false;
                //    }
                //}
                oGrd1.AutoResizeColumns();

                //데이터를 조회했을 경우에만 표시
                if (sType == "S")
                    ZSBO.SBO_Application.StatusBar.SetText("데이터를 조회하였습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                aForm.Freeze(false);
            }
            else if (value == "BTN_SEL")
            {
                int aRow = 0;
                string chkY = "N";
                for (aRow = 0; aRow < oGrd1.Rows.Count; aRow++)
                {
                    chkY = oGrd1.DataTable.Columns.Item(0).Cells.Item(aRow).Value.ToString();
                    if (chkY == "Y") break;
                }

                if (chkY == "N")
                {
                    ZSBO.SBO_Application.MessageBox("선택한 내역이 없습니다!", 1, "확인", "", "");
                    aForm.Freeze(false);
                    return;
                }

                //선택 품번에 대한 Matrix 반영
                ZMAT1.Clear();
                ZMAT1.FlushToDataSource();
                int oRow = 0;
                for (int n = 0; n < oGrd1.Rows.Count; n++)
                {
                    string selY = oGrd1.DataTable.Columns.Item(0).Cells.Item(n).Value.ToString();
                    if (selY == "Y")
                    {
                        ZDBDSL.Offset = ZDBDSL.Size - 1;
                        ZDBDSL.InsertRecord(ZDBDSL.Size);

                        ZDBDSL.SetValue("U_ITEMCODE", oRow, oGrd1.DataTable.Columns.Item("품번").Cells.Item(n).Value.ToString());
                        ZDBDSL.SetValue("U_ITEMNAME", oRow, oGrd1.DataTable.Columns.Item("품명").Cells.Item(n).Value.ToString());
                        ZDBDSL.SetValue("U_WHSCODE", oRow, oGrd1.DataTable.Columns.Item("출고창고").Cells.Item(n).Value.ToString());
                        ZDBDSL.SetValue("U_BINCODE", oRow, oGrd1.DataTable.Columns.Item("출고BIN").Cells.Item(n).Value.ToString());
                        ZDBDSL.SetValue("U_ONHAND", oRow, oGrd1.DataTable.Columns.Item("BIN재고").Cells.Item(n).Value.ToString());
                        ZDBDSL.SetValue("U_WONHAND", oRow, oGrd1.DataTable.Columns.Item("창고재고").Cells.Item(n).Value.ToString());
                        ZDBDSL.SetValue("U_QUANTITY", oRow, oGrd1.DataTable.Columns.Item("수량").Cells.Item(n).Value.ToString());

                        ZDBDSL.SetValue("U_SERIALNO", oRow, oGrd1.DataTable.Columns.Item("시리얼번호").Cells.Item(n).Value.ToString());
                        ZDBDSL.SetValue("U_BATCHNO", oRow, oGrd1.DataTable.Columns.Item("배치번호").Cells.Item(n).Value.ToString());
                        ZDBDSL.SetValue("U_TAGNO", oRow, oGrd1.DataTable.Columns.Item("TAG번호").Cells.Item(n).Value.ToString());
                        ZDBDSL.SetValue("U_SHMBER", oRow, oGrd1.DataTable.Columns.Item("호선번호").Cells.Item(n).Value.ToString());

                        oRow++;
                    }
                }

                if (oRow > 0)
                {
                    for (i = 0; i <= ZDBDSL.Size - 1; i++)
                    {
                        ZDBDSL.Offset = i;
                        ZDBDSL.SetValue("LineId", i, (i + 1).ToString());
                    }
                    ZMAT1.LoadFromDataSource();

                    // 체크박스
                    oGrd1.Columns.Item(0).Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox;

                    if (oGrd1.DataTable != null) oGrd1.DataTable = null;
                    if (aRecordSet != null) aRecordSet = null;

                    aForm.Freeze(false);
                    //ZMAT1.AutoResizeColumns();
                    aForm.Close();
                }
                else
                {
                    // 체크박스
                    oGrd1.Columns.Item(0).Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox;
                    ZSBO.SBO_Application.MessageBox("선택한 품번이 없습니다.", 1, "확인", "", "");
                    aForm.Freeze(false);
                }
            }
        }
        catch (Exception ex)
        {
            aForm.Freeze(false);
            if (aRecordSet != null)
            {
                oGrd1.DataTable = null;
                aRecordSet = null;
            }
            ZSBO.SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        }
    }

    //************************************************************************************************************
    // [참조 Form - 입력데이터체크] : ZMSLIST_ChkInDataA(헤더 및 라인)
    //************************************************************************************************************
    private bool ZMSLIST_ChkInDataA()
    {
        oForm.Freeze(true);

        try
        {
            //헤더 체크
            if (CheckIsNull(oDBDSH.GetValue("U_IOTYPE", 0).ToString().Trim()) == "")
            {
                oForm.Items.Item("IOTYPE").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                ZSBO.SBO_Application.StatusBar.SetText("입출고유형이 누락되었습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                return false;
            }
            //if (CheckIsNull(oDBDSH.GetValue("U_ITEMCODE", 0).ToString().Trim()) == "")
            //{
            //    oForm.Items.Item("ITEMCODE").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
            //    ZSBO.SBO_Application.StatusBar.SetText("품번이 누락되었습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            //    return false;
            //}
            return true;
        }
        catch (Exception ex)
        {
            ZSBO.SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            return false;
        }
        finally
        {
            oForm.Freeze(false);
        }
    }

    //************************************************************************************************************
    // [참조 Form] : ZMSLIST_AddChooseFromListA(CFL 사용 대상 칼럼 지정)
    //************************************************************************************************************
    private void ZMSLIST_AddChooseFromListA()
    {
        try
        {
            SAPbouiCOM.EditText oEditCode = null;
            SAPbouiCOM.EditTextColumn oEditTCode = null;

            SAPbouiCOM.ChooseFromListCreationParams oCFLCreationParams;
            SAPbouiCOM.ChooseFromListCollection oCFLs;
            SAPbouiCOM.ChooseFromList oCFL;

            //품번 
            oEditCode = (SAPbouiCOM.EditText)aForm.Items.Item("ITEMCODE").Specific;
            oEditCode.DataBind.SetBound(true, "", "ITEMCODE");
            oEditCode.ChooseFromListUID = "ITEMCODE";
            oEditCode.ChooseFromListAlias = "ItemCode";

            oEditCode = (SAPbouiCOM.EditText)aForm.Items.Item("ITEMNAME").Specific;
            oEditCode.DataBind.SetBound(true, "", "ITEMNAME");
            oEditCode.ChooseFromListUID = "ITEMNAME";
            oEditCode.ChooseFromListAlias = "ItemName";

            //창고 
            oEditCode = (SAPbouiCOM.EditText)aForm.Items.Item("WHSCODE").Specific;
            oEditCode.DataBind.SetBound(true, "", "WHSCODE");
            oEditCode.ChooseFromListUID = "WHSCODE";
            oEditCode.ChooseFromListAlias = "WhsCode";

            oEditCode = (SAPbouiCOM.EditText)aForm.Items.Item("WHSNAME").Specific;
            oEditCode.DataBind.SetBound(true, "", "WHSNAME");
            oEditCode.ChooseFromListUID = "WHSNAME";
            oEditCode.ChooseFromListAlias = "WhsName";

            ////BIN코드
            //oEditCode = (SAPbouiCOM.EditText)aForm.Items.Item("BINCODE").Specific;
            //oEditCode.DataBind.SetBound(true, "", "BINCODE");
            //oEditCode.ChooseFromListUID = "BINCODE";
            //oEditCode.ChooseFromListAlias = "BinCode";

            //Grid
            oEditTCode = (SAPbouiCOM.EditTextColumn)oGrd1.Columns.Item("출고BIN");
            oEditTCode.ChooseFromListUID = "BINCODE";
            oEditTCode.ChooseFromListAlias = "BinCode";

            //oEditTCode = (SAPbouiCOM.EditTextColumn)oGrd1.Columns.Item("입고BIN");
            //oEditTCode.ChooseFromListUID   = "TBINCODE";
            //oEditTCode.ChooseFromListAlias = "BinCode";
        }
        catch (Exception ex)
        {
            ZSBO.SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        }
    }

    //************************************************************************************************************
    // [참조 Form] : ZMSLIST_SetChooseFromListCondsA(CFL 조건 설정)
    //************************************************************************************************************
    private void ZMSLIST_SetChooseFromListCondsA()
    {
        SAPbouiCOM.Conditions oConds;

        //if (sUID == "ITEMCODE" || sUID == "ITEMNAME" || CheckIsNull(sUID) == "")
        //{
        //품목 조건 설정
        oConds = (SAPbouiCOM.Conditions)ZSBO.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_Conditions);

        //상품
        oConds.Add();
        oConds.Item(oConds.Count - 1).BracketOpenNum = 1;
        oConds.Item(oConds.Count - 1).Alias = "ItmsGrpCod";
        oConds.Item(oConds.Count - 1).Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
        oConds.Item(oConds.Count - 1).CondVal = "100";
        oConds.Item(oConds.Count - 1).BracketCloseNum = 1;

        //OR
        oConds.Item(oConds.Count - 1).Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR;

        //원자재
        oConds.Add();
        oConds.Item(oConds.Count - 1).BracketOpenNum = 1;
        oConds.Item(oConds.Count - 1).Alias = "ItmsGrpCod";
        oConds.Item(oConds.Count - 1).Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
        oConds.Item(oConds.Count - 1).CondVal = "103";
        oConds.Item(oConds.Count - 1).BracketCloseNum = 1;

        aForm.ChooseFromLists.Item("ITEMCODE").SetConditions(oConds);
        aForm.ChooseFromLists.Item("ITEMNAME").SetConditions(oConds);
        //}
    }

    //************************************************************************************************************
    // [MakeDI] : DI 처리
    //************************************************************************************************************
    private static bool MakeDI()
    {
        //2024.11.06  로직 추가
        SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)ZSBO.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        SAPbobsCOM.Recordset pRecordSet = (SAPbobsCOM.Recordset)ZSBO.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        SAPbobsCOM.Documents oOIGE = (SAPbobsCOM.Documents)ZSBO.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenExit);     //출고
        SAPbobsCOM.Documents oOIGN = (SAPbobsCOM.Documents)ZSBO.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenEntry);    //입고

        int retcode;
        int iRetcode, oAbsEntry;
        string sQry, sRetKey = "", sRetType = "", retKey = "", retType = "", sMassage, oDOCENTRY, pQry;
        bool bChk;

        try
        {
            if (ZSBO.oCompany.InTransaction)
            {
                ZSBO.SBO_Application.SetStatusBarMessage("트랜잭션 상태 오류가 발생했습니다.\n"
                      + ZSBO.oCompany.GetLastErrorDescription().Replace("'", "''"), SAPbouiCOM.BoMessageTime.bmt_Short, true);
                ZSBO.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                return false;
            }

            ZSBO.oCompany.StartTransaction();
            ZSBO.SBO_Application.SetStatusBarMessage("출고 처리중 입니다...", SAPbouiCOM.BoMessageTime.bmt_Long, false);

            oDOCENTRY = oDBDSH.GetValue("DocEntry", 0).Trim();
            if (oDBDSH.GetValue("U_DOCSTAT", 0).ToString().Trim() == "C")
            {
                ZSBO.SBO_Application.SetStatusBarMessage("이미 승인 처리된 문서입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                return false;
            }

            sQry = "SELECT P.DOCENTRY FROM [IGN1] P JOIN [OIGN] Z ON P.DOCENTRY = Z.DOCENTRY WHERE Z.CANCELED = 'N' AND IsNull(P.U_BSOBJECT, '*') = 'ZMC112' AND IsNull(P.U_BSENTRY, '*') = '" + oDOCENTRY + "'";
            ZSBO.oRecordSet.DoQuery(sQry);
            if (ZSBO.oRecordSet.RecordCount > 0)
            {
                oDBDSH.SetValue("U_DOCSTAT", 0, "C");
                ZSBO.oRecordSet.DoQuery("UPDATE [@ZMC112H] SET U_DOCSTAT = 'C' WHERE DOCENTRY = '" + oDOCENTRY + "' ");
                ZSBO.SBO_Application.SetStatusBarMessage("승인이 완료되었습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, false);
                return true;
            }

            ///////////////////////////////////제품 출고 DI처리///////////////////////////////////////////
            oOIGE.HandWritten = SAPbobsCOM.BoYesNoEnum.tNO;
            oOIGE.DocDate = DateTime.Parse(ZFN.CDateDI(oDBDSH.GetValue("U_DOCDATE", 0).ToString().Trim()));
            //oOIGE.TaxDate = oOIGE.DocDate;
            oOIGE.Comments = oDBDSH.GetValue("U_REMARK", 0).ToString().Trim();
            //사업장지정
            oOIGE.BPL_IDAssignedToInvoice = 1;
            oRecordSet.DoQuery("SELECT U_VALUE FROM [@ZPGMD] WHERE CODE = 'ZMC112' AND U_TABLE = 'IOTYPE_E' AND U_VALID = 'Y'"); //품번대체-출고
            oOIGE.UserFields.Fields.Item("U_IOTYPE").Value = oRecordSet.Fields.Item(0).Value.ToString();

            int eLineCnt = 0;  // DI 문서 실제 라인 번호
            for (int i = 0; i < oDBDSL.Size; i++)
            {
                oDBDSL.Offset = i;
                if (oDBDSL.GetValue("U_ITEMCODE", i).ToString().Trim() != "")
                {
                    oOIGE.Lines.ItemCode        = oDBDSL.GetValue("U_ITEMCODE", i).ToString().Trim();
                    oOIGE.Lines.ItemDescription = oDBDSL.GetValue("U_ITEMNAME", i).ToString().Trim();
                    oOIGE.Lines.Quantity        = ZFN.CDbl(oDBDSL.GetValue("U_QUANTITY", i).Trim());
                    oOIGE.Lines.WarehouseCode   = oDBDSL.GetValue("U_WHSCODE", i);
                    //oOIGE.Lines.FreeText        = oDBDSL.GetValue("U_LREMARK", i);
                    //Line User Fields
                    //oOIGE.Lines.UserFields.Fields.Item("U_INVGRP").Value = oDBDSL.GetValue("U_TITEMCD", i).ToString().Trim();
                    //oOIGE.Lines.UserFields.Fields.Item("U_PBATCH").Value = oDBDSL.GetValue("U_TBATCH", i).ToString().Trim();
                    oOIGE.Lines.UserFields.Fields.Item("U_BSOBJECT").Value = oForm.BusinessObject.Type;
                    oOIGE.Lines.UserFields.Fields.Item("U_BSENTRY").Value  = oDBDSH.GetValue("DocEntry", 0).ToString();
                    oOIGE.Lines.UserFields.Fields.Item("U_BSLINE").Value   = (i + 1).ToString();

                    //BinCode로 AbsEntry 가져와서 처리 
                    pQry = "SELECT TOP 1 AbsEntry FROM OBIN WHERE WhsCode = '" + oDBDSL.GetValue("U_WHSCODE", i).ToString() + "' "
                                                          + " AND BinCode = '" + oDBDSL.GetValue("U_BINCODE", i).ToString() + "' ";
                    pRecordSet.DoQuery(pQry);

                    //배치/시리얼 처리 (BinAllocations보다 먼저 설정)
                    oRecordSet.DoQuery("SELECT ManBtchNum, ManSerNum FROM OITM WHERE ITEMCODE = '" + oDBDSL.GetValue("U_ITEMCODE", i).Trim() + "'");
                    bool eBatchYN = (oRecordSet.Fields.Item("ManBtchNum").Value.ToString() == "Y");
                    bool eSerialYN = (oRecordSet.Fields.Item("ManSerNum").Value.ToString() == "Y");

                    if (eBatchYN) //배치
                    {
                        oOIGE.Lines.BatchNumbers.SetCurrentLine(0);
                        oOIGE.Lines.BatchNumbers.BatchNumber = oDBDSL.GetValue("U_BATCHNO", i);
                        oOIGE.Lines.BatchNumbers.Quantity    = oOIGE.Lines.Quantity;
                        //2026.09.07.LKC. "Illegal value entered" 오류 발생(U_LOTNO 필드의 Valid Values/길이 제약으로 추정, OIGN 신규배치 생성 시에도 동일 오류 발생해 배치 신규/기존 여부와는 무관).
                        //U_LOTNO 필드 정의 확인 후 재적용 필요하여 우선 제거함
                        oOIGE.Lines.BatchNumbers.Add();
                    }
                    else if (eSerialYN) //시리얼
                    {
                        oOIGE.Lines.SerialNumbers.InternalSerialNumber     = oDBDSL.GetValue("U_SERIALNO", i);
                        oOIGE.Lines.SerialNumbers.ManufacturerSerialNumber = oDBDSL.GetValue("U_TAGNO", i).ToString(); //MnfSerial
                        oOIGE.Lines.SerialNumbers.BatchID                  = oDBDSL.GetValue("U_SHMBER", i).ToString(); //LotNumber
                        oOIGE.Lines.SerialNumbers.Add();
                    } 

                    //BinAllocations 처리 (배치/시리얼 설정 후)
                    if (pRecordSet.RecordCount > 0)
                    {
                        oAbsEntry = int.Parse(pRecordSet.Fields.Item(0).Value.ToString());
                        oOIGE.Lines.BinAllocations.BinAbsEntry = oAbsEntry;
                        oOIGE.Lines.BinAllocations.Quantity = ZFN.CDbl(oDBDSL.GetValue("U_QUANTITY", i).Trim());
                        oOIGE.Lines.BinAllocations.BaseLineNumber = eLineCnt;
                        if (eBatchYN || eSerialYN)
                        {
                            oOIGE.Lines.BinAllocations.SerialAndBatchNumbersBaseLine = 0;
                        }
                        oOIGE.Lines.BinAllocations.Add();
                    }

                    oOIGE.Lines.Add();
                    eLineCnt++;
                }
            }
            retcode = oOIGE.Add();
            if (retcode != 0)
            {
                if (ZSBO.oCompany.InTransaction) ZSBO.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                ZSBO.SBO_Application.SetStatusBarMessage("출고DI 처리 중 오류발생 : " + ZSBO.oCompany.GetLastErrorDescription() + "(" + retcode.ToString() + ")" + " [BSNC]", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                return false;
            }
            else
            {
                retKey = ZSBO.oCompany.GetNewObjectKey();
                retType = ZSBO.oCompany.GetNewObjectType();
                //Set 출고DI 번호
                oDBDSH.SetValue("U_DINUM", 0, retKey.ToString());
            }


            /////////////////////////////입고 DI처리//////////////////////////////////////
            ZSBO.SBO_Application.SetStatusBarMessage("입고 처리중 입니다...", SAPbouiCOM.BoMessageTime.bmt_Long, false);

            oOIGN.HandWritten = SAPbobsCOM.BoYesNoEnum.tNO;
            oOIGN.DocDate = DateTime.Parse(ZFN.CDateDI(oDBDSH.GetValue("U_DOCDATE", 0).ToString().Trim()));
            //oOIGN.TaxDate = oOIGN.DocDate;
            oOIGN.Comments = oDBDSH.GetValue("U_REMARK", 0).ToString().Trim();
            //사업장지정
            oOIGN.BPL_IDAssignedToInvoice = 1;
            //Set the Last Evaluated Price
            oOIGN.PaymentGroupCode = -2;
            //User Fields IOTYPE 환경변수 세팅
            oRecordSet.DoQuery("SELECT U_VALUE FROM [@ZPGMD] WHERE CODE = 'ZMC112' AND U_TABLE = 'IOTYPE_N' AND U_VALID = 'Y'"); //품번대체-입고
            oOIGN.UserFields.Fields.Item("U_IOTYPE").Value = oRecordSet.Fields.Item(0).Value.ToString();

            int nLineCnt = 0;  // DI 문서 실제 라인 번호
            for (int i = 0; i < oDBDSM.Size; i++)
            {
                oDBDSM.Offset = i;
                if (oDBDSM.GetValue("U_ITEMCODE", i).ToString().Trim() != "")
                {
                    oOIGN.Lines.ItemCode        = oDBDSM.GetValue("U_ITEMCODE", i).ToString().Trim();
                    oOIGN.Lines.ItemDescription = oDBDSM.GetValue("U_ITEMNAME", i).ToString().Trim();
                    oOIGN.Lines.Quantity        = ZFN.CDbl(oDBDSM.GetValue("U_QUANTITY", i).Trim());
                    oOIGN.Lines.WarehouseCode   = oDBDSM.GetValue("U_WHSCODE", i);              
                    //oOIGN.Lines.FreeText = oDBDSM.GetValue("U_LREMARK", i);
                    //Line User Fields // DI 상위품목, 상위 배치 적용
                    //oOIGN.Lines.UserFields.Fields.Item("U_INVGRP").Value = oDBDSM.GetValue("U_ITEMCODE", i).ToString().Trim();
                    //oOIGN.Lines.UserFields.Fields.Item("U_PBATCH").Value = oDBDSM.GetValue("U_BATCHNUM", i).ToString().Trim();
                    oOIGN.Lines.UserFields.Fields.Item("U_BSOBJECT").Value = oForm.BusinessObject.Type;
                    oOIGN.Lines.UserFields.Fields.Item("U_BSENTRY").Value  = oDBDSH.GetValue("DocEntry", 0).ToString();
                    oOIGN.Lines.UserFields.Fields.Item("U_BSLINE").Value   = (i + 1).ToString();

                    //BinCode로 AbsEntry 가져와서 처리.2026.02.19
                    pQry = "SELECT TOP 1 AbsEntry FROM OBIN WHERE WhsCode = '" + oDBDSM.GetValue("U_WHSCODE", i) + "' "   
                                                          + " AND BinCode = '" + oDBDSM.GetValue("U_BINCODE", i) + "' ";  
                    pRecordSet.DoQuery(pQry);

                    //배치/시리얼 처리 (BinAllocations보다 먼저 설정)
                    oRecordSet.DoQuery("SELECT ManBtchNum, ManSerNum FROM OITM WHERE ITEMCODE = '" + oDBDSM.GetValue("U_ITEMCODE", i).ToString() + "'");
                    bool nBatchYN = (oRecordSet.Fields.Item("ManBtchNum").Value.ToString() == "Y");
                    bool nSerialYN = (oRecordSet.Fields.Item("ManSerNum").Value.ToString() == "Y");

                    if (nBatchYN) //배치
                    {
                        oOIGN.Lines.BatchNumbers.SetCurrentLine(0);
                        oOIGN.Lines.BatchNumbers.BatchNumber = oDBDSM.GetValue("U_BATCHNO", i).ToString();
                        oOIGN.Lines.BatchNumbers.Quantity    = oOIGN.Lines.Quantity;
                        //2026.09.07.LKC. 신규 배치 생성 시점에도 "Illegal value entered" 오류 발생(U_LOTNO 필드의 Valid Values/길이 제약으로 추정).
                        //U_LOTNO 필드 정의 확인 후 재적용 필요하여 우선 제거함
                        oOIGN.Lines.BatchNumbers.Add();
                    }
                    else if (nSerialYN) //시리얼
                    {
                        oOIGN.Lines.SerialNumbers.InternalSerialNumber     = oDBDSM.GetValue("U_SERIALNO", i).ToString();
                        oOIGN.Lines.SerialNumbers.ManufacturerSerialNumber = oDBDSM.GetValue("U_TAGNO",    i).ToString(); //MnfSerial
                        oOIGN.Lines.SerialNumbers.BatchID                  = oDBDSM.GetValue("U_SHMBER",   i).ToString(); //LotNumber
                        
                        oOIGN.Lines.SerialNumbers.Add();
                    } 

                    //BinAllocations 처리 (배치/시리얼 설정 후)
                    if (pRecordSet.RecordCount > 0)
                    {
                        oAbsEntry = int.Parse(pRecordSet.Fields.Item(0).Value.ToString());
                        oOIGN.Lines.BinAllocations.BinAbsEntry = oAbsEntry;
                        oOIGN.Lines.BinAllocations.Quantity = ZFN.CDbl(oDBDSM.GetValue("U_QUANTITY", i).Trim());
                        oOIGN.Lines.BinAllocations.BaseLineNumber = nLineCnt;
                        if (nBatchYN || nSerialYN)
                        {
                            oOIGN.Lines.BinAllocations.SerialAndBatchNumbersBaseLine = 0;
                        }
                        oOIGN.Lines.BinAllocations.Add();
                    }

                    oOIGN.Lines.Add();
                    nLineCnt++;
                }
            }
            
            iRetcode = oOIGN.Add();
            if (iRetcode != 0)
            {
                if (ZSBO.oCompany.InTransaction) ZSBO.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                oDBDSH.SetValue("U_DINUM", 0, "");
                ZSBO.SBO_Application.SetStatusBarMessage("입고DI 처리 중 오류발생" + ZSBO.oCompany.GetLastErrorDescription() + "(" + iRetcode.ToString() + ")" + " [BSNC]", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                return false;
            }
            else
            {
                sRetKey = ZSBO.oCompany.GetNewObjectKey();
                sRetType = ZSBO.oCompany.GetNewObjectType();
                //Set 입고DI 번호
                oDBDSH.SetValue("U_DINUM2", 0, sRetKey.ToString());
                oDBDSH.SetValue("U_DOCSTAT", 0, "C");
            }

            if (ZSBO.oCompany.InTransaction) ZSBO.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
            ZSBO.SBO_Application.SetStatusBarMessage("Success.", SAPbouiCOM.BoMessageTime.bmt_Short, false);
            return true;
        }
        catch (Exception Error)
        {
            if (ZSBO.oCompany.InTransaction) ZSBO.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
            oDBDSH.SetValue("U_DINUM",   0, "");
            oDBDSH.SetValue("U_DINUM2" , 0, "");
            oDBDSH.SetValue("U_DOCSTAT", 0, "O");
            ZSBO.SBO_Application.SetStatusBarMessage("DI 처리중 오류 발생 : " + ZSBO.oCompany.GetLastErrorDescription() + Error.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            return false;
        }
        finally
        {
            if (oRecordSet != null) oRecordSet = null;
            if (pRecordSet != null) pRecordSet = null;
            if (oOIGE != null) oOIGE = null;
            if (oOIGN != null) oOIGN = null;
        }
    }

    //************************************************************************************************************
    // [CanceDI] : DI 취소
    //************************************************************************************************************
    private static bool CancelDI()
    {
        //2024.11.06  로직 추가
        SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)ZSBO.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        SAPbobsCOM.Recordset pRecordSet = (SAPbobsCOM.Recordset)ZSBO.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        SAPbobsCOM.Documents oCOIGE     = (SAPbobsCOM.Documents)ZSBO.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenEntry);    //출고취소(입고)
        SAPbobsCOM.Documents oCOIGN     = (SAPbobsCOM.Documents)ZSBO.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenExit);     //입고취소(출고)
        int retcode;
        int iRetcode, oAbsEntry;
        string sQry, sRetKey = "", sRetType = "", retKey = "", retType = "", sMassage, oDOCENTRY = "", pQry;
        bool bChk;

        if (ZSBO.oCompany.InTransaction)
        {
            ZSBO.SBO_Application.SetStatusBarMessage("트랜잭션 상태 오류가 발생했습니다.\n"
                  + ZSBO.oCompany.GetLastErrorDescription().Replace("'", "''"), SAPbouiCOM.BoMessageTime.bmt_Short, true);
            ZSBO.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
            return false;
        }

        try
        {
            ZSBO.oCompany.StartTransaction();
            ZSBO.SBO_Application.SetStatusBarMessage("출고 취소 처리중 입니다...", SAPbouiCOM.BoMessageTime.bmt_Long, false);
            oDOCENTRY = oDBDSH.GetValue("DocEntry", 0).Trim();
            ///////////////////////////////////출고취소 DI처리///////////////////////////////////////////
            oCOIGE.HandWritten = SAPbobsCOM.BoYesNoEnum.tNO;
            oCOIGE.DocDate = DateTime.Parse(ZFN.CDateDI(oDBDSH.GetValue("U_DOCDATE", 0).ToString().Trim()));
            //oOIGE.TaxDate = oOIGE.DocDate;
            oCOIGE.Comments = "품목대체등록(출고취소)";
            //사업장지정
            oCOIGE.BPL_IDAssignedToInvoice = 1;
            //Header User Fields IOTYPE 환경변수 세팅
            oRecordSet.DoQuery("SELECT U_VALUE FROM [@ZPGMD] WHERE CODE = 'ZMC112' AND U_TABLE = 'IOTYPE_E' AND U_VALID = 'Y'");
            oCOIGE.UserFields.Fields.Item("U_IOTYPE").Value = oRecordSet.Fields.Item(0).Value.ToString();

            int eLineCnt = 0;  // DI 문서 실제 라인 번호
            for (int i = 0; i < oDBDSL.Size; i++)
            {
                oDBDSL.Offset = i;
                if (oDBDSL.GetValue("U_ITEMCODE", i).ToString().Trim() != "")
                {
                    oCOIGE.Lines.ItemCode        = oDBDSL.GetValue("U_ITEMCODE", i).ToString().Trim();             //ITEMCODE
                    oCOIGE.Lines.ItemDescription = oDBDSL.GetValue("U_ITEMNAME", i).ToString().Trim();
                    oCOIGE.Lines.Quantity        = ZFN.CDbl(oDBDSL.GetValue("U_QUANTITY", i).Trim());
                    oCOIGE.Lines.WarehouseCode   = oDBDSL.GetValue("U_WHSCODE", i);                          //WHSCODE
                    //oCOIGE.Lines.FreeText = oDBDSL.GetValue("U_LREMARK", i);

                    //Line User Fields // DI 상위품목, 상위 배치 적용 (2025.06.13.KDI)
                    //oCOIGE.Lines.UserFields.Fields.Item("U_INVGRP").Value = oDBDSL.GetValue("U_ITEMCODE", i).ToString().Trim();
                    //oCOIGE.Lines.UserFields.Fields.Item("U_PBATCH").Value = oDBDSL.GetValue("U_TBATCH", i).ToString().Trim();
                    oCOIGE.Lines.UserFields.Fields.Item("U_BSOBJECT").Value = oForm.BusinessObject.Type;
                    oCOIGE.Lines.UserFields.Fields.Item("U_BSENTRY").Value  = oDBDSH.GetValue("DocEntry", 0).ToString();
                    oCOIGE.Lines.UserFields.Fields.Item("U_BSLINE").Value   = (i + 1).ToString();

                    //BinCode로 AbsEntry 가져와서 처리.2026.02.19
                    pQry = "SELECT TOP 1 AbsEntry FROM OBIN WHERE WhsCode = '" + oDBDSL.GetValue("U_WHSCODE", i) + "' "   //WHSCODE
                                                          + " AND BinCode = '" + oDBDSL.GetValue("U_BINCODE", i) + "' ";  //BINCODE
                    pRecordSet.DoQuery(pQry);

                    //배치/시리얼 처리 (BinAllocations보다 먼저 설정)
                    oRecordSet.DoQuery("SELECT ManBtchNum, ManSerNum FROM OITM WHERE ITEMCODE = '" + oDBDSL.GetValue("U_ITEMCODE", i).Trim() + "'");
                    bool eBatchYN = (oRecordSet.Fields.Item("ManBtchNum").Value.ToString() == "Y");
                    bool eSerialYN = (oRecordSet.Fields.Item("ManSerNum").Value.ToString() == "Y");

                    if (eBatchYN) //배치
                    {
                        oCOIGE.Lines.BatchNumbers.SetCurrentLine(0);
                        oCOIGE.Lines.BatchNumbers.BatchNumber = oDBDSL.GetValue("U_BATCHNO", i);
                        oCOIGE.Lines.BatchNumbers.Quantity    = oCOIGE.Lines.Quantity;
                        oCOIGE.Lines.BatchNumbers.Add();
                    }
                    else if (eSerialYN) //시리얼
                    {
                        oCOIGE.Lines.SerialNumbers.InternalSerialNumber     = oDBDSL.GetValue("U_SERIALNO", i);
                        oCOIGE.Lines.SerialNumbers.ManufacturerSerialNumber = oDBDSL.GetValue("U_TAGNO", i).ToString(); //MnfSerial
                        oCOIGE.Lines.SerialNumbers.BatchID                  = oDBDSL.GetValue("U_SHMBER", i).ToString(); //LotNumber
                        oCOIGE.Lines.SerialNumbers.Add();
                    } 

                    //BinAllocations 처리 (배치/시리얼 설정 후)
                    if (pRecordSet.RecordCount > 0)
                    {
                        oAbsEntry = int.Parse(pRecordSet.Fields.Item(0).Value.ToString());
                        oCOIGE.Lines.BinAllocations.BinAbsEntry    = oAbsEntry;
                        oCOIGE.Lines.BinAllocations.Quantity       = ZFN.CDbl(oDBDSL.GetValue("U_QUANTITY", i).Trim());
                        oCOIGE.Lines.BinAllocations.BaseLineNumber = eLineCnt;
                        if (eBatchYN || eSerialYN)
                        {
                            oCOIGE.Lines.BinAllocations.SerialAndBatchNumbersBaseLine = 0;
                        }
                        oCOIGE.Lines.BinAllocations.Add();
                    }

                    oCOIGE.Lines.Add();
                    eLineCnt++;
                }
            }
            retcode = oCOIGE.Add();
            if (retcode != 0)
            {
                if (ZSBO.oCompany.InTransaction) ZSBO.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                ZSBO.SBO_Application.SetStatusBarMessage("출고취소DI 처리 중 오류발생 : " + ZSBO.oCompany.GetLastErrorDescription() + "(" + retcode.ToString() + ")" + " [BSNC]", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                return false;
            }
            else
            {
                retKey = ZSBO.oCompany.GetNewObjectKey();
                retType = ZSBO.oCompany.GetNewObjectType();
                //Set 출고취소DI 번호
                oDBDSH.SetValue("U_DICNUM", 0, retKey.ToString());
            }



            /////////////////////////////입고취소 DI처리//////////////////////////////////////
            ZSBO.SBO_Application.SetStatusBarMessage("입고취소 처리중 입니다...", SAPbouiCOM.BoMessageTime.bmt_Long, false);

            oCOIGN.HandWritten = SAPbobsCOM.BoYesNoEnum.tNO;
            oCOIGN.DocDate = DateTime.Parse(ZFN.CDateDI(oDBDSH.GetValue("U_DOCDATE", 0).ToString().Trim()));
            //oCOIGN.TaxDate = oOIGN.DocDate;
            oCOIGN.Comments = "품목대체등록(입고취소)";
            //사업장지정
            oCOIGN.BPL_IDAssignedToInvoice = 1;
            //Set the Last Evaluated Price
            oCOIGN.PaymentGroupCode = -2;
            //User Fields IOTYPE 환경변수 세팅
            oRecordSet.DoQuery("SELECT U_VALUE FROM [@ZPGMD] WHERE CODE = 'ZMC112' AND U_TABLE = 'IOTYPE_N' AND U_VALID = 'Y'");
            oCOIGN.UserFields.Fields.Item("U_IOTYPE").Value = oRecordSet.Fields.Item(0).Value.ToString();

            int nLineCnt = 0;  // DI 문서 실제 라인 번호
            for (int i = 0; i < oDBDSM.Size; i++)
            {
                oDBDSM.Offset = i;
                if (oDBDSM.GetValue("U_ITEMCODE", i).ToString().Trim() != "")
                {
                    oCOIGN.Lines.ItemCode        = oDBDSM.GetValue("U_ITEMCODE", i).ToString().Trim();            //TITEMCD
                    oCOIGN.Lines.ItemDescription = oDBDSM.GetValue("U_ITEMNAME", i).ToString().Trim();
                    oCOIGN.Lines.Quantity        = ZFN.CDbl(oDBDSM.GetValue("U_QUANTITY", i).Trim());
                    oCOIGN.Lines.WarehouseCode   = oDBDSM.GetValue("U_WHSCODE", i);                        //TWHSCODE
                    //oCOIGN.Lines.FreeText = oDBDSM.GetValue("U_LREMARK", i);
                    //Line User Fields // DI 상위품목, 상위 배치 적용 (2025.06.13.KDI)
                    //oCOIGN.Lines.UserFields.Fields.Item("U_INVGRP").Value = oDBDSM.GetValue("U_TITEMCD", i).ToString().Trim();
                    //oCOIGN.Lines.UserFields.Fields.Item("U_PBATCH").Value = oDBDSM.GetValue("U_BATCHNUM", i).ToString().Trim();
                    oCOIGN.Lines.UserFields.Fields.Item("U_BSOBJECT").Value = oForm.BusinessObject.Type;
                    oCOIGN.Lines.UserFields.Fields.Item("U_BSENTRY").Value  = oDBDSH.GetValue("DocEntry", 0).ToString();
                    oCOIGN.Lines.UserFields.Fields.Item("U_BSLINE").Value   = (i + 1).ToString();

                    //BinCode로 AbsEntry 가져와서 처리.2026.02.19
                    pQry = "SELECT TOP 1 AbsEntry FROM OBIN WHERE WhsCode = '" + oDBDSM.GetValue("U_WHSCODE", i) + "' "   //TWHSCODE
                                                          + " AND BinCode = '" + oDBDSM.GetValue("U_BINCODE", i) + "' ";  //TBINCODE
                    pRecordSet.DoQuery(pQry);

                    //배치/시리얼 처리 (BinAllocations보다 먼저 설정)
                    oRecordSet.DoQuery("SELECT ManBtchNum, ManSerNum FROM OITM WHERE ITEMCODE = '" + oDBDSM.GetValue("U_ITEMCODE", i).ToString() + "'");
                    bool nBatchYN = (oRecordSet.Fields.Item("ManBtchNum").Value.ToString() == "Y");
                    bool nSerialYN = (oRecordSet.Fields.Item("ManSerNum").Value.ToString() == "Y");

                    if (nBatchYN) //배치
                    {
                        oCOIGN.Lines.BatchNumbers.SetCurrentLine(0);
                        oCOIGN.Lines.BatchNumbers.BatchNumber = oDBDSM.GetValue("U_BATCHNO", i).ToString();
                        oCOIGN.Lines.BatchNumbers.Quantity    = oCOIGN.Lines.Quantity;
                        oCOIGN.Lines.BatchNumbers.Add();
                    }
                    else if (nSerialYN) //시리얼
                    {
                        oCOIGN.Lines.SerialNumbers.InternalSerialNumber     = oDBDSM.GetValue("U_SERIALNO", i).ToString();
                        oCOIGN.Lines.SerialNumbers.ManufacturerSerialNumber = oDBDSM.GetValue("U_TAGNO", i).ToString(); //MnfSerial
                        oCOIGN.Lines.SerialNumbers.BatchID                  = oDBDSM.GetValue("U_SHMBER", i).ToString(); //LotNumber
                        oCOIGN.Lines.SerialNumbers.Add();
                    } 

                    //BinAllocations 처리 (배치/시리얼 설정 후)
                    if (pRecordSet.RecordCount > 0)
                    {
                        oAbsEntry = int.Parse(pRecordSet.Fields.Item(0).Value.ToString());
                        oCOIGN.Lines.BinAllocations.BinAbsEntry    = oAbsEntry;
                        oCOIGN.Lines.BinAllocations.Quantity       = ZFN.CDbl(oDBDSM.GetValue("U_QUANTITY", i).Trim());
                        oCOIGN.Lines.BinAllocations.BaseLineNumber = nLineCnt;
                        if (nBatchYN || nSerialYN)
                        {
                            oCOIGN.Lines.BinAllocations.SerialAndBatchNumbersBaseLine = 0;
                        }
                        oCOIGN.Lines.BinAllocations.Add();
                    }

                    oCOIGN.Lines.Add();
                    nLineCnt++;
                }
            }
            iRetcode = oCOIGN.Add();
            if (iRetcode != 0)
            {
                if (ZSBO.oCompany.InTransaction) ZSBO.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                oDBDSH.SetValue("U_DICNUM", 0, "");
                ZSBO.SBO_Application.SetStatusBarMessage("입고취소DI 처리중 오류발생." + ZSBO.oCompany.GetLastErrorDescription() + "(" + iRetcode.ToString() + ")" + " [BSNC]", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                return false;
            }
            else
            {
                sRetKey = ZSBO.oCompany.GetNewObjectKey();
                sRetType = ZSBO.oCompany.GetNewObjectType();
                //Set 입고DI 번호
                oDBDSH.SetValue("U_DICNUM2", 0, sRetKey.ToString());
                oDBDSH.SetValue("CANCELED", 0, "Y");
            }

            if (ZSBO.oCompany.InTransaction) ZSBO.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
            ZSBO.SBO_Application.SetStatusBarMessage("Success.", SAPbouiCOM.BoMessageTime.bmt_Short, false);
            return true;
        }
        catch (Exception Error)
        {
            if (ZSBO.oCompany.InTransaction) ZSBO.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
            oDBDSH.SetValue("U_DICNUM", 0, "");
            oDBDSH.SetValue("U_DICNUM2", 0, "");
            oDBDSH.SetValue("CANCELED", 0, "N");
            ZSBO.SBO_Application.SetStatusBarMessage("DI 처리중 오류 발생 : " + ZSBO.oCompany.GetLastErrorDescription() + Error.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            return false;
        }
        finally
        {
            if (oRecordSet != null) oRecordSet = null;
            if (pRecordSet != null) pRecordSet = null;
            if (oCOIGE != null) oCOIGE = null;
            if (oCOIGN != null) oCOIGN = null;
        }
    }

    //************************************************************************************************************
    // [사용자정의함수 - CheckIsNull] :  데이터 Null값 체크하여 Null 일 경우 "" 값 반환
    //************************************************************************************************************
    private string CheckIsNull(object value)
    {
        if (value == null)
            return "";
        else
        {
            if (String.IsNullOrEmpty(value.ToString()) || value.ToString().Trim().Length == 0)
                return "";
            else
                return value.ToString();
        }
    }

    //************************************************************************************************************
    // [사용자정의함수 - GetOnHand] :  품목/창고/빈별로 재고수량 반환
    //************************************************************************************************************
    private double GetOnHand(string oType, string oItemCode, string oWhsCode, string oBinCode) // oType : "B" -> "BIN" 기준, "W" -> "WareHouse" 기준
    {
        SAPbobsCOM.Recordset oRecordset = (SAPbobsCOM.Recordset)ZSBO.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        string oQry = string.Empty;
        try
        {
            if (oType == "B") //BIN기준
            {
                   oQry = "SELECT ISNULL(SUM(T0.OnHandQty), 0)                                  "
                        + "  FROM OIBQ T0 INNER JOIN OWHS T1 ON T0.WhsCode = T1.WhsCode         "
                        + "               INNER JOIN OBIN T2 ON T0.BinAbs = T2.AbsEntry         "
                        + "               INNER JOIN OITM T3 ON T0.ItemCode = T3.ItemCode       "
                        + " WHERE ISNULL(T0.[OnHandQty], 0) <> 0                                "
                        + "   AND T0.ItemCode = '" + oItemCode + "' AND T1.WhsCode = '" + oWhsCode + "'  AND T2.BinCode = '" + oBinCode + "' ";
            }
            else //창고기준. 2026.04.24.LKC.
            {
                   oQry = "SELECT ISNULL(T0.OnHand, 0) FROM OITW T0                             "
                        + " WHERE T0.ItemCode = '" + oItemCode + "' AND T0.WhsCode = '" + oWhsCode + "' ";
            }
            oRecordset.DoQuery(oQry);
        }
        catch (Exception ex)
        {
            ZSBO.SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
        }

        return double.Parse(oRecordset.Fields.Item(0).Value.ToString());
    }   
}