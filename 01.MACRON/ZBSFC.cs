//************************************************************************************************************
//  File             : ZBSFC.cls
//  Module           : 공통
//  Desc             : 함수
//  Create Date      : 
//  Modified Date    : 
//  Creator          : Seo Young Mi
//  Copyright (c) BSNC
//************************************************************************************************************
using System;
using System.Windows.Forms;

public class ZBSFC
{
    //************************************************************************************************************
    // 일자 Format
    //************************************************************************************************************
    public string FChgDate(string sDate, string sDateSep)
    {
        if (sDate.Length > 10) { sDate = sDate.Substring(0, 10).Replace("-", ""); }
        if (sDate.Length == 10) { sDate = sDate.Replace(sDate.Substring(4, 1).ToString(), ""); }        // ZSBO.oCompany.GetCompanyService().GetAdminInfo().DateSeparator
        if (sDate.Length == 8) { sDate = sDate.Substring(0, 4) + sDateSep + sDate.Substring(4, 2) + sDateSep + sDate.Substring(6, 2); }
        else { sDate = ""; }
        return sDate;
    }

    //************************************************************************************************************
    // 필드의 기본값 가져오기 (숫자형식 = "0", 문자형식 = "", 날짜형식 = 값을 보낼때는 ""이고 타입을 보낼때는 "D")
    // 필드의 유형을 찾고자할 때는 sGubn = "T"로 받음
    //************************************************************************************************************
    public string FGetDfltValue(SAPbouiCOM.BoFieldsType sType, string sGubn)
    {
        string oStr = "";
        switch (sType)
        {
            case SAPbouiCOM.BoFieldsType.ft_Float:
            case SAPbouiCOM.BoFieldsType.ft_Integer:
            case SAPbouiCOM.BoFieldsType.ft_Percent:
            case SAPbouiCOM.BoFieldsType.ft_Price:
            case SAPbouiCOM.BoFieldsType.ft_Quantity:
            case SAPbouiCOM.BoFieldsType.ft_Rate:
            case SAPbouiCOM.BoFieldsType.ft_ShortNumber:
            case SAPbouiCOM.BoFieldsType.ft_Sum:
                oStr = "0";
                break;
            case SAPbouiCOM.BoFieldsType.ft_Date:
                if (sGubn == "T") { oStr = "D"; } else { oStr = ""; }
                break;
        }
        return oStr;
    }

    //************************************************************************************************************
    // 계정값 받기  ([0]:Segment, [1]:FormatCode, [2]:AcctCode, [3]:AcctName
    //************************************************************************************************************
    public string[] FGetAccont(string sField, string sValue)
    {
        SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)ZSBO.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        string[] oAccount = new string[4] { "", "", "", "" };
        int oMaxAbsId = 0;
        string oActSep = "-", sQry = "", oSegment = "";

        // 계정분할(최고차수, 0부터 시작)
        oRecordSet.DoQuery("SELECT MAX(AbsId) FROM OASG");
        if (oRecordSet.RecordCount > 0) { oMaxAbsId = (int)oRecordSet.Fields.Item(0).Value; }
        // 구분자
        oRecordSet.DoQuery("SELECT TOP 1 ActSep FROM OADM ORDER BY HldCode DESC");
        if (oRecordSet.RecordCount > 0) { oActSep = oRecordSet.Fields.Item(0).Value.ToString(); }

        // 계정내역
        sQry = "SELECT AcctCode, AcctName, FormatCode, FrgnName, Details, Segment_0, Segment_1, Segment_2, "
                    + "Segment_3, Segment_4, Segment_5, Segment_6, Segment_7, Segment_8, Segment_9 "
                + "FROM [OACT] ";
        if (sField == "Segment_0+'-'+Segment_1")
        {
            sQry += "WHERE FormatCode = '" + sValue.Replace("-", "") + "' AND AcctCode <> FormatCode ";
        }
        else if (sField == "FormatCode")
        {
            sQry += "WHERE " + sField + " = '" + sValue + "'";
        }
        else
        {
            sQry += "WHERE " + sField + " = '" + sValue + "' AND AcctCode <> FormatCode ";
        }
        oRecordSet.DoQuery(sQry);
        if (oRecordSet.RecordCount == 1)
        {
            for (int i = 0; i <= oMaxAbsId; i++)
            {
                if (oSegment != "") { oSegment += oActSep; };
                oSegment += oRecordSet.Fields.Item("Segment_" + i.ToString()).Value.ToString();
            }
            if (oSegment == "") oSegment = oRecordSet.Fields.Item("FormatCode").Value.ToString();
            oAccount[0] = oSegment;
            oAccount[1] = oRecordSet.Fields.Item("FormatCode").Value.ToString();
            oAccount[2] = oRecordSet.Fields.Item("AcctCode").Value.ToString();
            oAccount[3] = oRecordSet.Fields.Item("AcctName").Value.ToString();
        }

        return oAccount;
    }

    //************************************************************************************************************
    // 문서번호
    //************************************************************************************************************
    public string FGetDocNum(string sObjType)
    {
        SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)ZSBO.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        oRecordSet.DoQuery("SELECT Autokey FROM [ONNM] WHERE ObjectCode = '" + sObjType + "'");
        return oRecordSet.Fields.Item(0).Value.ToString();
    }

    //************************************************************************************************************
    // 현재 Matrix 내 Focus 위치
    //************************************************************************************************************
    public int FGetIndexRow(SAPbouiCOM.Matrix sMatrix)
    {
        int oRow = 0, oSelRow = 0;

        oSelRow = sMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_SelectionOrder);
        if ((sMatrix.GetCellFocus() == null) == false) { oRow = Convert.ToInt16(sMatrix.GetCellFocus().rowIndex.ToString()); }
        if (oSelRow > 0 & oSelRow < sMatrix.VisualRowCount) { oRow = oSelRow; }

        return oRow;
    }

    //************************************************************************************************************
    // 일수계산(시작일, 종료일)
    //************************************************************************************************************
    public string FDifDay(string oFDate, string oTDate)
    {
        SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)ZSBO.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        if (oFDate != "" & oTDate != "")
        {
            oRecordSet.DoQuery("SELECT 1 + DATEDIFF(D, '" + oFDate + "', '" + oTDate + "')");
            return oRecordSet.Fields.Item(0).Value.ToString();
        }
        else
        {
            return "0";
        }
    }

    //************************************************************************************************************
    // 년도체크
    //************************************************************************************************************
    public bool FChkYear(string oYear)
    {
        SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)ZSBO.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        oRecordSet.DoQuery("SELECT ISDATE('" + oYear + "-01-01')");

        if (oYear.Length != 4) { return false; }
        else if (oRecordSet.Fields.Item(0).Value.ToString() == "0") { return false; }
        else { return true; }
    }

    //************************************************************************************************************
    // 년월체크
    //************************************************************************************************************
    public bool FChkYM(string oYearMon)
    {
        SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)ZSBO.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        oRecordSet.DoQuery("SELECT ISDATE('" + oYearMon + "01')");

        if (oYearMon.Length != 6) { return false; }
        else if (oRecordSet.Fields.Item(0).Value.ToString() == "0") { return false; }
        else { return true; }
    }

    //************************************************************************************************************
    // 해당월 마지막 일자 
    // sDate, ReturnValue -> YYYYMMDD 형식
    // sGubn -> 파라메타값 -> D : 일자, YM : 년월
    //************************************************************************************************************
    public string FMonLastDate(string sGubn, string sDate, string sYear, string SMonth)
    {
        SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)ZSBO.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        string oDate; // 해당월의 첫날

        if (sGubn.ToUpper() == "D")
        {
            sDate = ZSBO.ZBSFC.FChgDate(sDate, "");
            sYear = sDate.Substring(0, 4);
            SMonth = sDate.Substring(4, 2);
        }
        oDate = sYear + "-" + SMonth + "-01";
        ZSBO.oRecordSet.DoQuery("SELECT CONVERT(NVARCHAR, DATEADD(D, -1, DATEADD(M, 1, '" + oDate + "')), 112)");
        //FMonLastDate() = ZSBO.oRecordSet.Fields.Item(0).Value;
        return oDate;
    }

    //************************************************************************************************************
    // 프린터 셋팅 데이터 받기 ([0]:SERVER, [1]:DBNAME, [2]:DBUSER, [3]:DBPW
    //************************************************************************************************************
    public string[] SetDBData()
    {
        SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)ZSBO.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        string[] oSetDBData = new string[4] { "", "", "", "" };

        oRecordSet.DoQuery("SELECT U_FldValue, U_FldDesc FROM [@ZPSDD] WHERE Code = 'PRINT' AND U_FldCode = 'CRYSET'");
        while (oRecordSet.EoF == false)
        {
            switch (oRecordSet.Fields.Item(0).Value.ToString().ToUpper())
            {
                case "SERVER": oSetDBData[0] = oRecordSet.Fields.Item(1).Value.ToString(); break;
                case "DBNAME": oSetDBData[1] = oRecordSet.Fields.Item(1).Value.ToString(); break;
                case "DBUSER": oSetDBData[2] = oRecordSet.Fields.Item(1).Value.ToString(); break;
                case "DBPW": oSetDBData[3] = oRecordSet.Fields.Item(1).Value.ToString(); break;
            }
            oRecordSet.MoveNext();
        }
        // DBName은 연결된 DB로 처리하도록 변경 (2023.11.22.SYM)
        oSetDBData[1] = ZSBO.oCompany.CompanyDB.ToString().ToUpper();

        return oSetDBData;
    }

    //************************************************************************************************************
    // EMAIL ID, PW ([0]:ID, [1]:PW
    //************************************************************************************************************
    public string[] SetEmailInfo()
    {
        SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)ZSBO.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        string[] oSetEmailInfo = new string[4] { "", "", "", "" };

        oRecordSet.DoQuery("SELECT U_FldValue, U_FldDesc FROM [@ZPSDD] WHERE Code = 'EMAIL'");
        while (oRecordSet.EoF == false)
        {
            switch (oRecordSet.Fields.Item(0).Value.ToString().ToUpper())
            {
                case "ID": oSetEmailInfo[0] = oRecordSet.Fields.Item(1).Value.ToString(); break;
                case "PW": oSetEmailInfo[1] = oRecordSet.Fields.Item(1).Value.ToString(); break;

            }
            oRecordSet.MoveNext();
        }
        return oSetEmailInfo;
    }
}
