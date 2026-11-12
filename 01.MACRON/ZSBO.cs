//****************************************************************************
//  File            : ZSBO.cs
//  Creator         : 
//  Create Date     : 
//  Memo            : 
//      2025.01.22.KDI. SAP Process Kill
//      2025.05.12.SYM. Kill Process 관련 item Event 수정
//      2025.10.26.SYM. 애드온로그생성 (ZSBO(). AddOnLog())
//  Copyright (c) BS&C
//****************************************************************************
using System;
using System.Windows.Forms;
//using System.Collections.Generic;
//using System.Text;
using System.Runtime.InteropServices;
using System.Diagnostics;           // Kill관련 (2025.02.03.KDI)
using System.Linq;                  // Kill관련 (2025.02.03.KDI)


public class ZSBO
{
    public static SAPbouiCOM.Application SBO_Application;
    public static SAPbobsCOM.Company oCompany;
    public static SAPbobsCOM.Recordset oRecordSet;

    //public static System.Collections.ArrayList sapForms = new System.Collections.ArrayList();
    //public static Hashtable mForm = new Hashtable();

    //-------------------------------------------------------------------------------------------
    public static string oBtnYES;                 // 버튼(언어별) : 예       :   [@ZMSGH].U_Class = 'ZCOM'   //   [@ZMSGH].U_TxtId = 'CXTYES'
    public static string oBtnNO;                  // 버튼(언어별) : 아니요   :   [@ZMSGH].U_Class = 'ZCOM'   //   [@ZMSGH].U_TxtId = 'CXTNO'
    public static string oBtnADD;                 // 버튼(언어별) : 추가     :   [@ZMSGH].U_Class = 'ZCOM'   //   [@ZMSGH].U_TxtId = 'CXTADD'
    public static string oBtnUPD;                 // 버튼(언어별) : 갱신     :   [@ZMSGH].U_Class = 'ZCOM'   //   [@ZMSGH].U_TxtId = 'CXTUPD'
    public static string oBtnOK;                  // 버튼(언어별) : 확인     :   [@ZMSGH].U_Class = 'ZCOM'   //   [@ZMSGH].U_TxtId = 'CXTOK'
    public static string oBtnCNL;                 // 버튼(언어별) : 취소     :   [@ZMSGH].U_Class = 'ZCOM'   //   [@ZMSGH].U_TxtId = 'CXTCNL'
    public static string oBtnDSP;                 // 버튼(언어별) : 조회     :   [@ZMSGH].U_Class = 'ZCOM'   //   [@ZMSGH].U_TxtId = 'CXTDSP'
    public static string oBtnPRT;                 // 버튼(언어별) : 출력     :   [@ZMSGH].U_Class = 'ZCOM'   //   [@ZMSGH].U_TxtId = 'CXTPRT'
    public static string oBtnSelY;                // 버튼(언어별) : 전체선택 :   [@ZMSGH].U_Class = 'ZCOM'   //   [@ZMSGH].U_TxtId = 'CXT0095'
    public static string oBtnSelN;                // 버튼(언어별) : 전체해제 :   [@ZMSGH].U_Class = 'ZCOM'   //   [@ZMSGH].U_TxtId = 'CXT0096'
    public static string oBtnCLSYN;               // 버튼(언어별) : 저장되지 않은 데이타가 있습니다. 종료하시겠습니까? :   [@ZMSGH].U_Class = 'ZCOM'   //   [@ZMSGH].U_TxtId = 'CXT0002'
    public static string oLocalCurrency;
    public static string oLocalCountry;
    public static string oToDay;                  // 현재 날짜
    public static string oDBName;                 // DB Name
    public static int oADEC;                      // 금액 라운딩 자리수
    public static int oQDEC;                      // 수량 라운딩 자리수
    public static int oPDEC;                      // 단가 라운딩 자리수
    public static int oRDEC;                      // 비율 라운딩 자리수
    public static int oCDEC;                      // 퍼센트 라운딩 자리수
    public static string oLoginUser;
    public static string oNavigate;               // URL 주소
    public static string oDateSep;                // 날짜 구분자
    public static string oDftGrp;                  // 
    public static string oBPLId;
    public static string oPlant;
    public static string oCCenter;
    public static string oPrcCode;

    //-------------------------------------------------------------------------------------------
    [DllImport("kernel32.dll", SetLastError = true)]                                            // KILL관련 (2025.02.03.KDI)
    private static extern bool ProcessIdToSessionId(int dwProcessId, out uint pSessionId);      // KILL관련 (2025.02.03.KDI)
    private static bool oCompanyChanging = false;                                               // 사업장 변경 플래그 (2025.02.03.KDI)

    //-------------------------------------------------------------------------------------------
    public static ZAPI ZAPI = new ZAPI();                                           // ZAPI
    public static ZBSFC ZBSFC = new ZBSFC();                                        // ZBSFC
    private static ZCHLIST oZCHLIST = new ZCHLIST();                                // ZCHLIST
    private static ZOPENF oZOPENF = new ZOPENF();                                   // ZOPENF
    private static ZOPENEXL oZOPENEXL = new ZOPENEXL();                             // ZOPENEXL
    public static ZSetEF oZSetEF = new ZSetEF();                                    // EventFilters

    //-------------------------------------------------------------------------------------------
    //private static ZSAMPLE oZSAMPLE = new ZSAMPLE();                              // Sample
    //-------------------------------------------------------------------------------------------



    private static ZMC905 oZMC905 = new ZMC905();                                   // 마스터 스케쥴


    //판매관리
    private static ZMC918 oZMC918 = new ZMC918();                                   // 납품_일련번호/배치번호별 납품 현황조회(TAG,호선번호)
    private static ZMC919 oZMC919 = new ZMC919();                                   // 물류_배치번호별 판매오더조회
    private static ZMC920 oZMC920 = new ZMC920();                                   // 물류_피킹리스트 생성 현황조회(TAG,배치번호)
    private static ZMC921 oZMC921 = new ZMC921();                                   // 물류_TAG 일련번호별 판매오더조회
    private static ZMC922 oZMC922 = new ZMC922();                                   // 수주잔액현황조회(순액/선매출)
    private static ZMC923 oZMC923 = new ZMC923();                                   // 판매오더_단가 누락 품목조회(가격리스트01기준)
    private static ZMC924 oZMC924 = new ZMC924();                                   // 판매오더_호선별 재고현황조회
    private static ZMC925 oZMC925 = new ZMC925();                                   // 현재고 판매오더 대사 현황조회(품목코드 존재유무)
    private static ZMC926 oZMC926 = new ZMC926();                                   // AR송장_매출·미청구채권 상세 현황조회
    private static ZMC927 oZMC927 = new ZMC927();                                   // AR송장_선매출·미청구채권 잔액 현황조회
    private static ZMC936 oZMC936 = new ZMC936();                                   // [영업마감] 매출반품/AR대변메모 집계
    private static ZMC937 oZMC937 = new ZMC937();                                   // [영업마감] 미결납품 집계
    private static ZMC938 oZMC938 = new ZMC938();                                   // [영업마감] AR송장 집계
    private static ZMC939 oZMC939 = new ZMC939();                                   // [영업마감] AR송장 집계(거래처,프로젝트,호선,제품군)

    private static ZMC202 oZMC202 = new ZMC202();                                   // 프로젝트 공정관리
    private static ZMC112 oZMC112 = new ZMC112();                                   // 품목대체등록(배치시리얼)

    //재고관리
    private static ZMC928 oZMC928 = new ZMC928();                                   // 창고별 재고실사 조회(호선번호, TAG번호, 관리번호, PO번호)
    private static ZMC944 oZMC944 = new ZMC944();                                   // [재고마감] 일련/배치 관리규칙 오류 집계
    private static ZMC945 oZMC945 = new ZMC945();                                   // [재고마감] 일련/배치 재고 정합성
    private static ZMC946 oZMC946 = new ZMC946();                                   // [재고마감] 재고실사 조정현황
    private static ZMC947 oZMC947 = new ZMC947();                                   // [재고마감] 품목대체 입출고 수량 비교

    //구매관리
    private static ZMC929 oZMC929 = new ZMC929();                                   // [구매마감] 구매반품 집계
    private static ZMC930 oZMC930 = new ZMC930();                                   // [구매마감] 기타입출고 집계
    private static ZMC931 oZMC931 = new ZMC931();                                   // [구매마감] 이상 매입단가 확인
    private static ZMC932 oZMC932 = new ZMC932();                                   // [구매마감] AP대변메모 집계
    private static ZMC933 oZMC933 = new ZMC933();                                   // [구매마감] AP송장 집계
    private static ZMC934 oZMC934 = new ZMC934();                                   // [구매마감] AP예약송장(미착) 집계

    //생산관리
    private static ZMC935 oZMC935 = new ZMC935();                                   // 생산관리_구성품 재고확보 현황조회
    private static ZMC940 oZMC940 = new ZMC940();                                   // [생산마감] 생산오더 계획대비 생산입고 차이현황
    private static ZMC941 oZMC941 = new ZMC941();                                   // [생산마감] 생산오더 미완료 집계
    private static ZMC942 oZMC942 = new ZMC942();                                   // [생산마감] 생산오더 미종료 현황(생산입고완료건)
    private static ZMC943 oZMC943 = new ZMC943();                                   // [생산마감] 생산출고 정합성 집계


    //*******************************************************************************************
    private void SetApplication()
    {

        // *******************************************************************
        // Use an SboGuiApi object to establish connection
        // with the SAP Business One application and return an
        // initialized appliction object
        // *******************************************************************

        SAPbouiCOM.SboGuiApi SboGuiApi = null;
        string sConnectionString = null;

        SboGuiApi = new SAPbouiCOM.SboGuiApi();

        // by following the steps specified above, the following
        // statment should be suficient for either development or run mode
        sConnectionString = System.Convert.ToString(Environment.GetCommandLineArgs().GetValue(1));
        //sConnectionString = "0030002C0030002C00530041005000420044005F00440061007400650076002C0050004C006F006D0056004900490056";

        // connect to a running SBO Application
        SboGuiApi.Connect(sConnectionString);

        // get an initialized application object
        SBO_Application = SboGuiApi.GetApplication(-1);

        oCompany = (SAPbobsCOM.Company)SBO_Application.Company.GetDICompany();
        oRecordSet = ((SAPbobsCOM.Recordset)(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)));

        // events handled by SBO_Application_AppEvent 
        SBO_Application.AppEvent += new SAPbouiCOM._IApplicationEvents_AppEventEventHandler(SBO_Application_AppEvent);
        //// events handled by SBO_Application_ProgressBarEvent
        //SBO_Application.ProgressBarEvent += new SAPbouiCOM._IApplicationEvents_ProgressBarEventEventHandler(SBO_Application_ProgressBarEvent); 
        // events handled by SBO_Application_MenuEvent 
        SBO_Application.MenuEvent += new SAPbouiCOM._IApplicationEvents_MenuEventEventHandler(SBO_Application_MenuEvent);
        // events handled by SBO_Application_ItemEvent
        SBO_Application.ItemEvent += new SAPbouiCOM._IApplicationEvents_ItemEventEventHandler(SBO_Application_ItemEvent);
        // events handled by SBO_Application_FormDataEvent
        SBO_Application.FormDataEvent += new SAPbouiCOM._IApplicationEvents_FormDataEventEventHandler(SBO_Application_FormDataEvent);
        // events handled by SBO_Application_RightClickEvent
        SBO_Application.RightClickEvent += new SAPbouiCOM._IApplicationEvents_RightClickEventEventHandler(SBO_Application_RightClickEvent);
    }

    //*******************************************************************************************
    private void AddingMenusWithXML()
    {       
        System.Xml.XmlDocument oXmlDoc = null;
        SAPbouiCOM.Menus oMenu;
        SAPbouiCOM.BoLanguages oLang = ZSBO.SBO_Application.Language;       // 연결된 언어
        string sPath = null;
        string tmpStr;
        string oMenuXml = "";
        
        oXmlDoc = new System.Xml.XmlDocument();
        oDBName = oCompany.CompanyDB.ToUpper();
        sPath = Application.StartupPath;

        oMenuXml = "MENU_MACRON_LO_KR";
        if (oMenuXml == "") { return; }
        oXmlDoc.Load(sPath + @"\Image\" + oMenuXml + ".xml");

        // load the form to the SBO application in one batch
        tmpStr = oXmlDoc.InnerXml;
        SBO_Application.LoadBatchActions(ref tmpStr);

        // 메뉴 이미지
        //oMenu = SBO_Application.Menus.Item("43520").SubMenus;
        //if (oMenu.Exists("MSD_00")) { oMenu.Item("M_00").Image = sPath + @"\Image\MSD.png"; }             // 기준정보관리
    }

    //*******************************************************************************************
    // 애드온로그생성 (2025.10.26.SYM)
    private void AddonLog()
    {
        ZSBO.oRecordSet.DoQuery("SELECT * FROM OUTB WHERE TABLENAME = 'ZACS003H'");
        if (ZSBO.oRecordSet.RecordCount == 0) { return; }

        string @ZUSER = "", @ANAME = "", @WKTITLE = "", @WKFILE = "", @SERVERID = "", @REMARK = "";
        @ZUSER = ZSBO.oCompany.UserName;        // 작업자ID
        @ANAME = "BSNC_MACRON_LO_X64";          // 애드온이름
        @WKTITLE = "";                          // 작업명 (애드온이름)
        @WKFILE = "";                           // 작업파일명 (EXE 파일)
        @SERVERID = "";                         // 서비스번호

        // 작업내용 (최대200자까지만입력)
        @REMARK = "메뉴명-xxx";

        try
        {
            ZSBO.oRecordSet.DoQuery("EXEC	[DBO].[ZTMSP_ZACS003]	@TYPE = 'C', @ZUSER = '" + @ZUSER + "', @ANAME = '" + @ANAME
                                                                + "', @AVERSION = '', @WKTITLE = '" + @WKTITLE + "', @WKFILE = '" + @WKFILE
                                                                + "', @SERVERID = '" + @SERVERID + "', @REMARK = '" + @REMARK + "'");

            if (ZSBO.oRecordSet.Fields.Item(0).Value.ToString().ToUpper() == "OK")
            {
                ZSBO.SBO_Application.StatusBar.SetText("AddOnLog Success", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
            }
        }
        catch (Exception Error)
        {
            //ZSBO.SBO_Application.StatusBar.SetText("AddOnLog Error : " + Error.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
        }
    }

    //*******************************************************************************************
    private void SetCommonVariableValue()
    {
        SAPbobsCOM.CompanyService oCompanyService = oCompany.GetCompanyService();

        //oBtnYES = ZAPI.GetMsg("", "CXTYES");     // 버튼(언어별) : 예       :   [@ZMSGH].U_Class = 'ZCOM'   //   [@ZMSGH].U_TxtId = 'CXTYES'
        //oBtnNO = ZAPI.GetMsg("", "CXTNO");       // 버튼(언어별) : 아니요   :   [@ZMSGH].U_Class = 'ZCOM'   //   [@ZMSGH].U_TxtId = 'CXTNO'
        //oBtnADD = ZAPI.GetMsg("", "CXTADD");     // 버튼(언어별) : 추가     :   [@ZMSGH].U_Class = 'ZCOM'   //   [@ZMSGH].U_TxtId = 'CXTADD'
        //oBtnUPD = ZAPI.GetMsg("", "CXTUPD");     // 버튼(언어별) : 갱신     :   [@ZMSGH].U_Class = 'ZCOM'   //   [@ZMSGH].U_TxtId = 'CXTUPD'
        //oBtnOK = ZAPI.GetMsg("", "CXTOK");       // 버튼(언어별) : 확인     :   [@ZMSGH].U_Class = 'ZCOM'   //   [@ZMSGH].U_TxtId = 'CXTOK'
        //oBtnCNL = ZAPI.GetMsg("", "CXTCNL");     // 버튼(언어별) : 취소     :   [@ZMSGH].U_Class = 'ZCOM'   //   [@ZMSGH].U_TxtId = 'CXTCNL'
        //oBtnDSP = ZAPI.GetMsg("", "CXTDSP");     // 버튼(언어별) : 조회     :   [@ZMSGH].U_Class = 'ZCOM'   //   [@ZMSGH].U_TxtId = 'CXTDSP'
        //oBtnPRT = ZAPI.GetMsg("", "CXTPRT");     // 버튼(언어별) : 출력     :   [@ZMSGH].U_Class = 'ZCOM'   //   [@ZMSGH].U_TxtId = 'CXTPRT'
        //oBtnSelY = ZAPI.GetMsg("", "CXT0095");   // 버튼(언어별) : 전체선택 :   [@ZMSGH].U_Class = 'ZCOM'   //   [@ZMSGH].U_TxtId = 'CXT0095'
        //oBtnSelN = ZAPI.GetMsg("", "CXT0096");   // 버튼(언어별) : 전체해제 :   [@ZMSGH].U_Class = 'ZCOM'   //   [@ZMSGH].U_TxtId = 'CXT0096'
        //oBtnCLSYN = ZAPI.GetMsg("", "CXT0002");  // 버튼(언어별) : 저장되지 않은 데이타가 있습니다. 종료하시겠습니까? :   [@ZMSGH].U_Class = 'ZCOM'   //   [@ZMSGH].U_TxtId = 'CXT0002'

        oLoginUser = oCompany.UserName;
        oQDEC = oCompanyService.GetAdminInfo().AccuracyofQuantities;
        oADEC = oCompanyService.GetAdminInfo().TotalsAccuracy;
        oPDEC = oCompanyService.GetAdminInfo().PriceAccuracy;
        oRDEC = oCompanyService.GetAdminInfo().RateAccuracy;
        oLocalCurrency = oCompanyService.GetAdminInfo().LocalCurrency;
        oLocalCountry = oCompanyService.GetAdminInfo().Country;
        oDateSep = oCompanyService.GetAdminInfo().DateSeparator;
        oToDay = ZSBO.SBO_Application.Company.ServerDate.ToString().Trim().Replace(oDateSep, "");       //oToDay = DateTime.Today.ToString().Substring(0, 10).Replace("-", "");
        //////////oRecordSet.DoQuery("SELECT U_DFTGRP FROM OUSR WHERE USER_CODE = '" + oCompany.UserName + "'");
        //////////oDftGrp = oRecordSet.Fields.Item(0).Value.ToString().Trim();

        oRecordSet.DoQuery("SELECT H.U_BPLID, H.U_PLATCODE, L.U_CCENTER, L.U_PRCCODE " +
                                    " FROM [@ZG01H] AS H JOIN [@ZG01L] AS L ON H.DocEntry = L.DocEntry " +
                                    " WHERE H.CANCELED='N' AND L.U_USER='" + oCompany.UserName + "'");

        if (oRecordSet.RecordCount > 0)
        {
            oBPLId = oRecordSet.Fields.Item(0).Value.ToString().Trim();
            oPlant = oRecordSet.Fields.Item(1).Value.ToString().Trim();
            oCCenter = oRecordSet.Fields.Item(2).Value.ToString().Trim();
            oPrcCode = oRecordSet.Fields.Item(3).Value.ToString().Trim();
        }
    }

    //*******************************************************************************************
    private void SBO_Application_AppEvent(SAPbouiCOM.BoAppEventTypes EventType)
    {
        switch (EventType)
        {
            case SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged:
                AddingMenusWithXML();
                break;
            case SAPbouiCOM.BoAppEventTypes.aet_FontChanged:
                AddingMenusWithXML();
                break;
            case SAPbouiCOM.BoAppEventTypes.aet_LanguageChanged:
                AddingMenusWithXML();
                break;
            case SAPbouiCOM.BoAppEventTypes.aet_ShutDown:
            case SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition:
                if (SBO_Application.ClientType == SAPbouiCOM.BoClientType.ct_Desktop && oCompanyChanging)
                {
                    oCompanyChanging = false; // 플래그 리셋
                    GC.Collect();                               // 추가 (2025.10.16.SYM)
                    System.Windows.Forms.Application.Exit();    // 추가 (2025.10.16.SYM)
                }
                else if (SBO_Application.ClientType == SAPbouiCOM.BoClientType.ct_Desktop && oCompanyChanging == false)
                {
                    GC.Collect();
                    KillProcess();
                    System.Windows.Forms.Application.Exit();

                }
                break;
                //////////// 수정전원본 (2025.02.06.SYM)
                //////////GC.Collect();
                //////////KillSAPBusinessOneProcess();            // 추가 (2025.01.22.KDI)
                //////////System.Windows.Forms.Application.Exit();
                //////////break;
        }
    }

    //*******************************************************************************************
    private static void KillProcess()       // Kill (2025.02.03.KDI)
    {
        try
        {
            System.Diagnostics.Process[] PROC;
            string[] oArray;

            // 엑셀(EXCEL.EXE) 프로세스 종료
            PROC = System.Diagnostics.Process.GetProcessesByName("EXCEL");
            foreach (System.Diagnostics.Process PK in PROC)
            {
                if (PK.MainWindowTitle.Length == 0)     // 안죽은 프로세스는 타이틀이 없음
                {
                    PK.Kill();
                }
            }

            // SAP 관련 프로세스 종료(세션)
            uint sessionId = GetCurrentSessionId();
            KillSAPProcessesInSession(sessionId);

        }
        catch (Exception ex)
        {
        }
    }

    //*******************************************************************************************
    private static uint GetCurrentSessionId()           // Kill (2025.02.03.KDI)
    {
        uint sessionId;
        int processId = Process.GetCurrentProcess().Id; // 현재 프로세스 ID 가져오기

        if (ProcessIdToSessionId(processId, out sessionId))
        {
            return sessionId; // 성공 시 세션 ID 반환
        }
        else
        {
            throw new Exception("현재 세션 ID를 가져올 수 없습니다. 오류 코드: " + Marshal.GetLastWin32Error());
        }
    }

    //*******************************************************************************************
    private static void KillSAPProcessesInSession(uint sessionId)       // Kill (2025.02.03.KDI)
    {
        string[] oArray;

        oArray = new string[] { "BSNC_MACRON_COA", "BSNC_MACRON_FI", "BSNC_TAX_KR", "BSNC_HR_KR", "SAP Business One", "SAPbouiCOM", "BSNC_MACRON_LO" };
        foreach (string oStr in oArray)
        {
            // 모든 SAPBusinessOne 프로세스 확인
            var processes = Process.GetProcessesByName(oStr);
            foreach (var process in processes)
            {
                uint processSessionId;

                // 프로세스의 세션 ID 가져오기
                if (ProcessIdToSessionId(process.Id, out processSessionId))
                {
                    // 현재 세션과 동일한 경우 종료
                    if (processSessionId == sessionId)
                    {
                        process.Kill();         // 프로세스 종료
                        process.WaitForExit();  // 종료 완료 대기
                    }
                }
            }
        }
    }

    //*******************************************************************************************
    private void SBO_Application_ProgressBarEvent(ref SAPbouiCOM.ProgressBarEvent pVal, out bool BubbleEvent)
    {
        BubbleEvent = true;

        switch (pVal.EventType)
        {
            case SAPbouiCOM.BoProgressBarEventTypes.pbet_ProgressBarCreated:
            case SAPbouiCOM.BoProgressBarEventTypes.pbet_ProgressBarStopped:
            case SAPbouiCOM.BoProgressBarEventTypes.pbet_ProgressBarReleased:
                break;
        }
    } 

    //*******************************************************************************************
    private void SBO_Application_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
    {
        object oTmpObj;
        BubbleEvent = true;

        try
        {
            switch (SBO_Application.Forms.Item(FormUID).TypeEx)
            {
                //---------------------------------------------------------------------------------------------------------
                case "820":
                    if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK && pVal.ItemUID == "1" && pVal.BeforeAction == true)             // "Choose Company" 창 (2025.02.03.KDI)
                    {
                        oCompanyChanging = true; // 사업장 변경
                    }
                    break;
                case "60058":
                    if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED && pVal.ItemUID == "16" && pVal.BeforeAction == true)     // 애드온관리-애드온등록버튼 (2025.05.12.SYM)
                    {
                        oCompanyChanging = true; // 사업장 변경
                    }
                    break;
                case "60059":
                    if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED && pVal.ItemUID == "10000026" && pVal.BeforeAction == true)     // 애드온관리자-설치버튼 (2025.05.12.SYM)
                    {
                        oCompanyChanging = true; // 사업장 변경
                    }
                    if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED && pVal.ItemUID == "11" && pVal.BeforeAction == true)         // 애드온관리자-중지버튼 (2025.05.12.SYM)
                    {
                        oCompanyChanging = true; // 사업장 변경
                    }
                    break;

                //---------------------------------------------------------------------------------------------------------
                //case "SAMPLE": oSAMPLE.SBO_Application_ItemEvent(FormUID, ref pVal, out BubbleEvent); break;                      // SAMPLE
                //---------------------------------------------------------------------------------------------------------
                case "ZCHLIST": oZCHLIST.SBO_Application_ItemEvent(FormUID, ref pVal, out BubbleEvent); break;                      // ZCHLIST
                case "ZOPENF": oZOPENF.SBO_Application_ItemEvent(FormUID, ref pVal, out BubbleEvent); break;                        // ZOPENF
                case "ZOPENEXL": oZOPENEXL.SBO_Application_ItemEvent(FormUID, ref pVal, out BubbleEvent); break;                    // ZOPENEXL

                //---------------------------------------------------------------------------------------------------------


                case "ZMC905": oZMC905.SBO_Application_ItemEvent(FormUID, ref pVal, out BubbleEvent); break;                        // 마스터 스케쥴

                //판매관리
                case "ZMC918": oZMC918.SBO_Application_ItemEvent(FormUID, ref pVal, out BubbleEvent); break;                        // 납품_일련번호/배치번호별 납품 현황조회(TAG,호선번호)
                case "ZMC919": oZMC919.SBO_Application_ItemEvent(FormUID, ref pVal, out BubbleEvent); break;                        // 물류_배치번호별 판매오더조회
                case "ZMC920": oZMC920.SBO_Application_ItemEvent(FormUID, ref pVal, out BubbleEvent); break;                        // 물류_피킹리스트 생성 현황조회(TAG,배치번호)
                case "ZMC921": oZMC921.SBO_Application_ItemEvent(FormUID, ref pVal, out BubbleEvent); break;                        // 물류_TAG 일련번호별 판매오더조회
                case "ZMC922": oZMC922.SBO_Application_ItemEvent(FormUID, ref pVal, out BubbleEvent); break;                        // 수주잔액현황조회(순액/선매출)
                case "ZMC923": oZMC923.SBO_Application_ItemEvent(FormUID, ref pVal, out BubbleEvent); break;                        // 판매오더_단가 누락 품목조회(가격리스트01기준)
                case "ZMC924": oZMC924.SBO_Application_ItemEvent(FormUID, ref pVal, out BubbleEvent); break;                        // 판매오더_호선별 재고현황조회
                case "ZMC925": oZMC925.SBO_Application_ItemEvent(FormUID, ref pVal, out BubbleEvent); break;                        // 현재고 판매오더 대사 현황조회(품목코드 존재유무)
                case "ZMC926": oZMC926.SBO_Application_ItemEvent(FormUID, ref pVal, out BubbleEvent); break;                        // AR송장_매출·미청구채권 상세 현황조회
                case "ZMC927": oZMC927.SBO_Application_ItemEvent(FormUID, ref pVal, out BubbleEvent); break;                        // AR송장_선매출·미청구채권 잔액 현황조회
                case "ZMC936": oZMC936.SBO_Application_ItemEvent(FormUID, ref pVal, out BubbleEvent); break;                        // [영업마감] 매출반품/AR대변메모 집계
                case "ZMC937": oZMC937.SBO_Application_ItemEvent(FormUID, ref pVal, out BubbleEvent); break;                        // [영업마감] 미결납품 집계
                case "ZMC938": oZMC938.SBO_Application_ItemEvent(FormUID, ref pVal, out BubbleEvent); break;                        // [영업마감] AR송장 집계
                case "ZMC939": oZMC939.SBO_Application_ItemEvent(FormUID, ref pVal, out BubbleEvent); break;                        // [영업마감] AR송장 집계(거래처,프로젝트,호선,제품군)

                case "ZMC202A": // 프로젝트공정관리 판매오더 팝업창
                case "ZMC202B": // 프로젝트공정관리 공정단계관리 팝업창
                case "ZMC202": oZMC202.SBO_Application_ItemEvent(FormUID, ref pVal, out BubbleEvent); break;                        // 프로젝트공정관리

                case "ZMC112A":
                case "ZMC112": oZMC112.SBO_Application_ItemEvent(FormUID, ref pVal, out BubbleEvent); break;                        // 품목대체등록(배치시리얼)

                //재고관리
                case "ZMC928": oZMC928.SBO_Application_ItemEvent(FormUID, ref pVal, out BubbleEvent); break;                        // 창고별 재고실사 조회(호선번호, TAG번호, 관리번호, PO번호)
                case "ZMC944": oZMC944.SBO_Application_ItemEvent(FormUID, ref pVal, out BubbleEvent); break;                        // [재고마감] 일련/배치 관리규칙 오류 집계
                case "ZMC945": oZMC945.SBO_Application_ItemEvent(FormUID, ref pVal, out BubbleEvent); break;                        // [재고마감] 일련/배치 재고 정합성
                case "ZMC946": oZMC946.SBO_Application_ItemEvent(FormUID, ref pVal, out BubbleEvent); break;                        // [재고마감] 재고실사 조정현황
                case "ZMC947": oZMC947.SBO_Application_ItemEvent(FormUID, ref pVal, out BubbleEvent); break;                        // [재고마감] 품목대체 입출고 수량 비교


                //구매관리 
                case "ZMC929": oZMC929.SBO_Application_ItemEvent(FormUID, ref pVal, out BubbleEvent); break;                        // [구매마감] 구매반품 집계
                case "ZMC930": oZMC930.SBO_Application_ItemEvent(FormUID, ref pVal, out BubbleEvent); break;                        // [구매마감] 기타입출고 집계
                case "ZMC931": oZMC931.SBO_Application_ItemEvent(FormUID, ref pVal, out BubbleEvent); break;                        // [구매마감] 이상 매입단가 확인
                case "ZMC932": oZMC932.SBO_Application_ItemEvent(FormUID, ref pVal, out BubbleEvent); break;                        // [구매마감] AP대변메모 집계
                case "ZMC933": oZMC933.SBO_Application_ItemEvent(FormUID, ref pVal, out BubbleEvent); break;                        // [구매마감] AP송장 집계
                case "ZMC934": oZMC934.SBO_Application_ItemEvent(FormUID, ref pVal, out BubbleEvent); break;                        // [구매마감] AP예약송장(미착) 집계

                //생산관리
                case "ZMC935": oZMC935.SBO_Application_ItemEvent(FormUID, ref pVal, out BubbleEvent); break;                        // 생산관리_구성품 재고확보 현황조회
                case "ZMC940": oZMC940.SBO_Application_ItemEvent(FormUID, ref pVal, out BubbleEvent); break;                        // [생산마감] 생산오더 계획대비 생산입고 차이현황
                case "ZMC941": oZMC941.SBO_Application_ItemEvent(FormUID, ref pVal, out BubbleEvent); break;                        // [생산마감] 생산오더 미완료 집계
                case "ZMC942": oZMC942.SBO_Application_ItemEvent(FormUID, ref pVal, out BubbleEvent); break;                        // [생산마감] 생산오더 미종료 현황(생산입고완료건)
                case "ZMC943": oZMC943.SBO_Application_ItemEvent(FormUID, ref pVal, out BubbleEvent); break;                        // [생산마감] 생산출고 정합성 집계







                //---------------------------------------------------------------------------------------------------------
            }
            //oTmpObj = mForm[FormUID];
            //oTmpObj.SBO_Application_ItemEvent(FormUID, ref pVal, out BubbleEvent);

        }
        catch (Exception Error)
        {
            //SBO_Application.SetStatusBarMessage(Error.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        }
    }

    //*******************************************************************************************
    private void SBO_Application_MenuEvent(ref SAPbouiCOM.MenuEvent pVal, out bool BubbleEvent)
    {

        // ********************************************************************************
        //  in order to activate your own forms instead of SAP Business One system forms
        //  process the menu event by your self
        //  change BubbleEvent to False so that SAP Business One won't process it
        // ********************************************************************************
        string oTypeUID = "";

        BubbleEvent = true;

        if (pVal.MenuUID.Substring(0, 1) == "Z")
            oTypeUID = pVal.MenuUID;
        else if (SBO_Application.Forms.ActiveForm.TypeEx.Substring(0, 1) == "Z")
            oTypeUID = SBO_Application.Forms.ActiveForm.TypeEx;
        else
            oTypeUID = SBO_Application.Forms.ActiveForm.TypeEx;     // 추가 (2023.03.20.YYJ)

        // 수정전원본 (2022.03.21.SYM)
        //if (SBO_Application.Forms.ActiveForm.TypeEx.Substring(0, 1) == "Z")
        //    oTypeUID = SBO_Application.Forms.ActiveForm.TypeEx;
        //else if (pVal.MenuUID.Substring(0, 1) == "Z") oTypeUID = pVal.MenuUID;
        
        try
        {
            switch (oTypeUID)
            {
                //---------------------------------------------------------------------------------------------------------
                //case "SAMPLE": oSAMPLE.SBO_Application_MenuEvent(ref pVal, out BubbleEvent); break;                               // SAMPLE
                //---------------------------------------------------------------------------------------------------------
                case "ZOPENF": oZOPENF.SBO_Application_MenuEvent(ref pVal, out BubbleEvent); break;                                 // ZOPENF
                case "ZOPENEXL": oZOPENEXL.SBO_Application_MenuEvent(ref pVal, out BubbleEvent); break;                             // ZOPENEXL

                //---------------------------------------------------------------------------------------------------------


                case "ZMC905": oZMC905.SBO_Application_MenuEvent(ref pVal, out BubbleEvent); break;                                 // 마스터 스케쥴

                //판매관리                
                case "ZMC918": oZMC918.SBO_Application_MenuEvent(ref pVal, out BubbleEvent); break;                                 // 납품_일련번호/배치번호별 납품 현황조회(TAG,호선번호)
                case "ZMC919": oZMC919.SBO_Application_MenuEvent(ref pVal, out BubbleEvent); break;                                 // 물류_배치번호별 판매오더조회
                case "ZMC920": oZMC920.SBO_Application_MenuEvent(ref pVal, out BubbleEvent); break;                                 // 물류_피킹리스트 생성 현황조회(TAG,배치번호)
                case "ZMC921": oZMC921.SBO_Application_MenuEvent(ref pVal, out BubbleEvent); break;                                 // 물류_TAG 일련번호별 판매오더조회
                case "ZMC922": oZMC922.SBO_Application_MenuEvent(ref pVal, out BubbleEvent); break;                                 // 수주잔액현황조회(순액/선매출)
                case "ZMC923": oZMC923.SBO_Application_MenuEvent(ref pVal, out BubbleEvent); break;                                 // 판매오더_단가 누락 품목조회(가격리스트01기준)
                case "ZMC924": oZMC924.SBO_Application_MenuEvent(ref pVal, out BubbleEvent); break;                                 // 판매오더_호선별 재고현황조회
                case "ZMC925": oZMC925.SBO_Application_MenuEvent(ref pVal, out BubbleEvent); break;                                 // 현재고 판매오더 대사 현황조회(품목코드 존재유무)
                case "ZMC926": oZMC926.SBO_Application_MenuEvent(ref pVal, out BubbleEvent); break;                                 // AR송장_매출·미청구채권 상세 현황조회
                case "ZMC927": oZMC927.SBO_Application_MenuEvent(ref pVal, out BubbleEvent); break;                                 // AR송장_선매출·미청구채권 잔액 현황조회

                case "ZMC936": oZMC936.SBO_Application_MenuEvent(ref pVal, out BubbleEvent); break;                                 // [영업마감] 매출반품/AR대변메모 집계
                case "ZMC937": oZMC937.SBO_Application_MenuEvent(ref pVal, out BubbleEvent); break;                                 // [영업마감] 미결납품 집계
                case "ZMC938": oZMC938.SBO_Application_MenuEvent(ref pVal, out BubbleEvent); break;                                 // [영업마감] AR송장 집계
                case "ZMC939": oZMC939.SBO_Application_MenuEvent(ref pVal, out BubbleEvent); break;                                 // [영업마감] AR송장 집계(거래처,프로젝트,호선,제품군)

                case "ZMC202": oZMC202.SBO_Application_MenuEvent(ref pVal, out BubbleEvent); break;                                 // 프로젝트공정관리
                case "ZMC202A": // 프로젝트공정관리 판매오더 팝업창 
                case "ZMC202B": // 프로젝트공정관리 공정단계관리 팝업창 

                case "ZMC112": oZMC112.SBO_Application_MenuEvent(ref pVal, out BubbleEvent); break;                                 // 품목대체등록(배치시리얼)
                case "ZMC112A":

                //재고관리
                case "ZMC928": oZMC928.SBO_Application_MenuEvent(ref pVal, out BubbleEvent); break;                                 // 창고별 재고실사 조회(호선번호, TAG번호, 관리번호, PO번호)
                case "ZMC944": oZMC944.SBO_Application_MenuEvent(ref pVal, out BubbleEvent); break;                                 // [재고마감] 일련/배치 관리규칙 오류 집계
                case "ZMC945": oZMC945.SBO_Application_MenuEvent(ref pVal, out BubbleEvent); break;                                 // [재고마감] 일련/배치 재고 정합성
                case "ZMC946": oZMC946.SBO_Application_MenuEvent(ref pVal, out BubbleEvent); break;                                 // [재고마감] 재고실사 조정현황
                case "ZMC947": oZMC947.SBO_Application_MenuEvent(ref pVal, out BubbleEvent); break;                                 // [재고마감] 품목대체 입출고 수량 비교


                //구매관리 
                case "ZMC929": oZMC929.SBO_Application_MenuEvent(ref pVal, out BubbleEvent); break;                                 // [구매마감] 구매반품 집계
                case "ZMC930": oZMC930.SBO_Application_MenuEvent(ref pVal, out BubbleEvent); break;                                 // [구매마감] 기타입출고 집계
                case "ZMC931": oZMC931.SBO_Application_MenuEvent(ref pVal, out BubbleEvent); break;                                 // [구매마감] 이상 매입단가 확인
                case "ZMC932": oZMC932.SBO_Application_MenuEvent(ref pVal, out BubbleEvent); break;                                 // [구매마감] AP대변메모 집계
                case "ZMC933": oZMC933.SBO_Application_MenuEvent(ref pVal, out BubbleEvent); break;                                 // [구매마감] AP송장 집계
                case "ZMC934": oZMC934.SBO_Application_MenuEvent(ref pVal, out BubbleEvent); break;                                 // [구매마감] AP예약송장(미착) 집계

                //생산관리 
                case "ZMC935": oZMC935.SBO_Application_MenuEvent(ref pVal, out BubbleEvent); break;                                 // 생산관리_구성품 재고확보 현황조회
                case "ZMC940": oZMC940.SBO_Application_MenuEvent(ref pVal, out BubbleEvent); break;                                 // [생산마감] 생산오더 계획대비 생산입고 차이현황
                case "ZMC941": oZMC941.SBO_Application_MenuEvent(ref pVal, out BubbleEvent); break;                                 // [생산마감] 생산오더 미완료 집계
                case "ZMC942": oZMC942.SBO_Application_MenuEvent(ref pVal, out BubbleEvent); break;                                 // [생산마감] 생산오더 미종료 현황(생산입고완료건)
                case "ZMC943": oZMC943.SBO_Application_MenuEvent(ref pVal, out BubbleEvent); break;                                 // [생산마감] 생산출고 정합성 집계
                    

                //---------------------------------------------------------------------------------------------------------
            }
        }
        catch (Exception Error)
        {
            SBO_Application.SetStatusBarMessage(Error.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        }
    }

    //*******************************************************************************************
    private void SBO_Application_FormDataEvent(ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, out bool BubbleEvent)
    {
        string oTypeUID = BusinessObjectInfo.FormTypeEx;
        BubbleEvent = true;

        if (ZSBO.SBO_Application.Forms.Item(BusinessObjectInfo.FormUID).IsSystem) { return; }
/*
        try
        {
            switch (oTypeUID)
            {
                //---------------------------------------------------------------------------------------------------------
                //case "SAMPLE": oSAMPLE.SBO_Application_FormDataEvent(ref BusinessObjectInfo, out BubbleEvent); break;                 // SAMPLE
                //---------------------------------------------------------------------------------------------------------

                //---------------------------------------------------------------------------------------------------------


                // 작성부분


                //---------------------------------------------------------------------------------------------------------
            }
        }
        catch (Exception Error)
        {
            SBO_Application.SetStatusBarMessage(Error.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        }
 */ 
    }
    //*******************************************************************************************
    private void SBO_Application_RightClickEvent(ref SAPbouiCOM.ContextMenuInfo oEventInfo, out bool BubbleEvent)
    {
        BubbleEvent = true;

        try
        {
            if (ZSBO.SBO_Application.Forms.Item(oEventInfo.FormUID).IsSystem) { return; }

/*
            switch (SBO_Application.Forms.Item(oEventInfo.FormUID).TypeEx)
            {
                //---------------------------------------------------------------------------------------------------------
                //case "SAMPLE": oSAMPLE.SBO_Application_RightClickEvent(ref oEventInfo, out BubbleEvent); break;                           // SAMPLE
                //---------------------------------------------------------------------------------------------------------

                //---------------------------------------------------------------------------------------------------------


                // 작성부분


                //---------------------------------------------------------------------------------------------------------
            }
  */ 
        }
        catch (Exception Error)
        {
            SBO_Application.SetStatusBarMessage(Error.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        }

    }

    //*******************************************************************************************
    public ZSBO()
    {
        SetApplication();
        CheckLicense();
        SetCommonVariableValue();
        AddingMenusWithXML();
        //oZSetEF.SetEventFilter();
        AddonLog();     // 애드온로그생성(2025.10.26.SYM)
    }

    //*******************************************************************************************
    private void CheckLicense()
    {
        bool oAccessYN;
        string oDBName;
        string oServerName;
        string oVersion;

        oAccessYN = false;
        oDBName = oCompany.CompanyDB.ToString().ToUpper();
        oServerName = SBO_Application.Company.ServerName.ToString().ToUpper();
        oVersion = oCompany.Version.ToString().ToUpper();


        oAccessYN = true;


        ////////////-------------------------------------------------------------------------------------
        //////////// 회사별 체크
        ////////////-------------------------------------------------------------------------------------
        //////////// 매크론
        //////////if ((oDBName.IndexOf("MACRON_SBO") >= 0) & (oVersion == "1000180"))
        //////////{
        //////////    //SBO_Application.SetStatusBarMessage("ServerName:" + oServerName, SAPbouiCOM.BoMessageTime.bmt_Short, false);
        //////////    oAccessYN = true;
        //////////}

        ////////////-------------------------------------------------------------------------------------
        //////////// 개발자 PC Name 기록
        //////////if (oServerName == "DESKTOP-83R2RPF" | oServerName == "DESKTOP-3HF7Q8H")     // SYM:DESKTOP-83R2RPF, JMK:DESKTOP-3HF7Q8H
        //////////{
        //////////    oAccessYN = true;
        //////////}

        //-------------------------------------------------------------------------------------
        if (oAccessYN == false)
        {
            //SBO_Application.SetStatusBarMessage("No license.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
            SBO_Application.SetStatusBarMessage("라이선스가 없습니다", SAPbouiCOM.BoMessageTime.bmt_Short, true);
            oCompany.Disconnect();
            GC.Collect();
            System.Windows.Forms.Application.Exit();
            Application.ExitThread();
            Environment.Exit(0);            //프로그램 종료
        }
        //-------------------------------------------------------------------------------------
    }

    //*******************************************************************************************
}


//*******************************************************************************************
internal class WindowWrapper : System.Windows.Forms.IWin32Window
{
    [DllImport("user32.dll")]
    private static extern IntPtr GetForegroundWindow();
    private IntPtr handle;

    public virtual IntPtr Handle
    {
        get
        {
            return handle;
        }
    }

    public WindowWrapper(IntPtr handle)
    {
        this.handle = handle;
    }

    public static WindowWrapper GetForegroundWindowWrapper()
    {
        return new WindowWrapper(GetForegroundWindow());
    }
}