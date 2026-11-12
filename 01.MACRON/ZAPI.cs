//************************************************************************************************************
//  File             : ZAPI.cls
//  Module           : 공통
//  Desc             : 함수
//  Create Date      : 
//  Modified Date    : 
//  Creator          : 
//  Copyright (c) BSNC
//************************************************************************************************************
using System;
using System.Windows.Forms;
using System.Net.Mail;
using System.Net;
using System.IO;

////Automanagerd Enum----------------------------------
public enum eAttrib { Editable, Visible };
public enum eFormMode { All, Add, Find, Ok, View };
public enum eBehavior { True, False, Default };
////---------------------------------------------------

public class ZAPI
{
    //************************************************************************************************************
    // 
    //************************************************************************************************************
    public int GetBplId()
    {
        //ZSBO.oCompany.UserSignature
        string bplid;

        bplid = GetOneValueQuery("SELECT BPLID FROM OUSR A INNER JOIN OUDG B ON A.DFLTSGROUP = B.CODE WHERE A.USERID='" + ZSBO.oCompany.UserSignature.ToString() + "'");
        if (bplid == null)
        {
            ZSBO.SBO_Application.SetStatusBarMessage("사용자의 BPLID가 없습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
            return -1;
        }

        return int.Parse(bplid);
    }

    //************************************************************************************************************
    // 
    //************************************************************************************************************
    public string GetOneValueQuery(string queryString)
    {
        SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)ZSBO.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        string result = null;
        try
        {

            oRecordSet.DoQuery(queryString);
            if (oRecordSet.Fields.Count != 0)
                result = oRecordSet.Fields.Item(0).Value.ToString();
            oRecordSet = null;
            return result;
        }
        catch (Exception Error)
        {
            oRecordSet = null;
            ZSBO.SBO_Application.SetStatusBarMessage(Error.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            return "error";
        }
    }

    //************************************************************************************************************
    // 언어별 메시지 가져오기
    //************************************************************************************************************
    public string GetMsg(string oClass, string oMsgNum)
    {
        try
        {
            string sQry = "SELECT  [dbo].[Fn_Get_MESSAGE] ('" + ZSBO.SBO_Application.Language + "','" + oClass + "','" + oMsgNum + "')";

            ZSBO.oRecordSet.DoQuery(sQry);
            if (ZSBO.oRecordSet.RecordCount == 0) return "";
            else return (string)ZSBO.oRecordSet.Fields.Item(0).Value;
        }
        catch (Exception Error)
        {
            ZSBO.SBO_Application.SetStatusBarMessage(Error.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            return "";
        }

    }

    //************************************************************************************************************
    // 
    //************************************************************************************************************
    public void AutoManaged(ref SAPbouiCOM.Form oForm, string itemsName, eAttrib Attrib, eFormMode autoFormMode, eBehavior Behavior)
    {
        SAPbouiCOM.BoAutoManagedAttr autoManagedAttr;
        SAPbouiCOM.BoModeVisualBehavior modeVisualBehavior;
        SAPbouiCOM.BoAutoFormMode FormMode;

        try
        {
            if (itemsName == "") { return; }

            string[] itemList = itemsName.Trim().Split(',');
            oForm.AutoManaged = true;

            //Attribute
            if (Attrib == global::eAttrib.Editable)
            {
                autoManagedAttr = SAPbouiCOM.BoAutoManagedAttr.ama_Editable;
            }
            else //if (Attrib ==global::eAttrib.Visible)
            {
                autoManagedAttr = SAPbouiCOM.BoAutoManagedAttr.ama_Visible;
            }

            //FormMode
            if (autoFormMode == global::eFormMode.Add)
            {
                FormMode = SAPbouiCOM.BoAutoFormMode.afm_Add;
            }
            else if (autoFormMode == global::eFormMode.Find)
            {
                FormMode = SAPbouiCOM.BoAutoFormMode.afm_Find;
            }
            else if (autoFormMode == global::eFormMode.Ok)
            {
                FormMode = SAPbouiCOM.BoAutoFormMode.afm_Ok;
            }
            else if (autoFormMode == global::eFormMode.View)
            {
                FormMode = SAPbouiCOM.BoAutoFormMode.afm_View;
            }
            else
            {
                FormMode = SAPbouiCOM.BoAutoFormMode.afm_All;
            }

            //Behavior
            if (Behavior == global::eBehavior.True)
            {
                modeVisualBehavior = SAPbouiCOM.BoModeVisualBehavior.mvb_True;
            }
            else if (Behavior == global::eBehavior.False)
            {
                modeVisualBehavior = SAPbouiCOM.BoModeVisualBehavior.mvb_False;
            }
            else
            {
                modeVisualBehavior = SAPbouiCOM.BoModeVisualBehavior.mvb_Default;
            }

            foreach (string str in itemList)
            {
                oForm.Items.Item(str).SetAutoManagedAttribute(autoManagedAttr, Convert.ToInt32(FormMode), modeVisualBehavior);
            }
        }
        catch (Exception Error)
        {
            ZSBO.SBO_Application.SetStatusBarMessage("ZFN.AutoManaged Error : " + Error.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        }
    }

    //************************************************************************************************************
    // Form Load 시 FormTitle 변경
    //************************************************************************************************************
    public string FGetFormTitle(string oFormTitle, string oMenuUID, string oFormType, string oObjType)
    {
        string oString, oReturn = "";

        try
        {
            if (ZSBO.SBO_Application.Menus.Exists(oMenuUID) == true)
                oString = ZSBO.SBO_Application.Menus.Item(oMenuUID).String;
            else
                return oFormTitle;

            //if (oObjType == "-1")
            //{
            if (oMenuUID == oFormType) oReturn = oString + " (" + oMenuUID + ")";
            else oReturn = oString + " (" + oMenuUID + ":" + oFormType + ")";
            //}
            //else
            //{
            //    if (oMenuUID == oFormType) oReturn = oString + " (" + oMenuUID + ":" + oObjType + ")";
            //    else oReturn = oString + " (" + oMenuUID + ":" + oFormType + ":" + oObjType + ")";
            //}

            return oReturn;

        }
        catch (Exception Error)
        {
            ZSBO.SBO_Application.SetStatusBarMessage(Error.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            return "";
        }

    }

    //************************************************************************************************************
    // Form Load 시 FormTitle 변경
    //************************************************************************************************************
    public string GetDeptName(string oDeptCode)
    {
        SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)ZSBO.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        string oQry, oReturn = "";

        try
        {
            oQry = "SELECT U_DEPNAME FROM [@ZHR007H] WHERE U_DEPCODE = '" + oDeptCode;
            oQry += "' AND ISNULL(U_NUse, 'N') <> 'Y' AND U_ORGNUM = (SELECT TOP 1 U_ORGNUM FROM [@ZHR000L] WHERE U_CHKUSE = 'Y')";
            oRecordSet.DoQuery(oQry);

            oReturn = oRecordSet.Fields.Item(0).Value.ToString();

            return oReturn;

        }
        catch (Exception Error)
        {
            ZSBO.SBO_Application.SetStatusBarMessage(Error.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            return "";
        }

    }

    //************************************************************************************************************
    // Convert_Date : 일자 형태 변환 
    //************************************************************************************************************
    public string ConvertDate(string oDate, string oDec)
    {
        string oConvertDate = "";

        try
        {
            switch (oDate.Length.ToString())
            {
                case "8":
                    if (oDec == "YMD")
                        oConvertDate = oDate.Substring(0, 4) + "년" + oDate.Substring(4, 2) + "월" + oDate.Substring(6, 2) + "일";
                    else if (oDec == "Y M D")
                        oConvertDate = oDate.Substring(0, 4) + "년 " + oDate.Substring(4, 2) + "월 " + oDate.Substring(6, 2) + "일";
                    else
                        oConvertDate = oDate.Substring(0, 4) + oDec + oDate.Substring(4, 2) + oDec + oDate.Substring(6, 2);
                    break;

                case "10":
                    if (oDec == "YMD")
                        oConvertDate = oDate.Substring(0, 4) + "년" + oDate.Substring(5, 2) + "월" + oDate.Substring(8, 2) + "일";
                    else if (oDec == "Y M D")
                        oConvertDate = oDate.Substring(0, 4) + "년 " + oDate.Substring(5, 2) + "월 " + oDate.Substring(8, 2) + "일";
                    else
                        oConvertDate = oDate.Substring(0, 4) + oDec + oDate.Substring(5, 2) + oDec + oDate.Substring(8, 2);
                    break;

            }
            return oConvertDate;
        }
        catch (Exception Error)
        {
            ZSBO.SBO_Application.SetStatusBarMessage(Error.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            return oConvertDate;
        }
    }

    //************************************************************************************************************
    // 문자 받침 여부 체크
    //************************************************************************************************************
    public string GetWa(string oStr, string oGubn)   // ReturnFinalConsonant
    {
        string oReturn = oStr + oGubn;
        bool oBool = true;

        if (ZSBO.SBO_Application.Language == SAPbouiCOM.BoLanguages.ln_Korean_Kr)   // 언어가 한국어인 경우
        {
            foreach (char c in oStr.Substring(oStr.Length - 1, 1))
            {
                if (((c - 16) % 28) == 0)
                    oBool = false;  // 받침 없음
                else
                    oBool = true;   // 받침 있음
            }

            if (oGubn == "이" | oGubn == "가") oReturn = oStr + (oBool ? "이" : "가");  // (oBool ? True : False)
            else if (oGubn == "은" | oGubn == "는") oReturn = oStr + (oBool ? "은" : "는");
            else if (oGubn == "와" | oGubn == "과") oReturn = oStr + (oBool ? "과" : "와");
        }
        else
        {
            oReturn = oStr;
        }

        return oReturn;
    }


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
    // 프린터 셋팅 데이터 받기 ([0]:SERVER, [1]:DBNAME, [2]:DBUSER, [3]:DBPW
    //************************************************************************************************************
    public string[] SetDBData()
    {
        SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)ZSBO.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        string[] oSetDBData = new string[4] { "", "", "", "" };

        oRecordSet.DoQuery("SELECT U_FLDVALUE,U_FLDDESC FROM [@ZPSDD]  WHERE U_FLDCODE = 'CRYSET'");
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
        return oSetDBData;
    }

    //************************************************************************************************************
    // 
    //************************************************************************************************************
    public string Left(string sItem, int iLength)
    {
        try
        {
            if (sItem.Length > 0)
            {
                if (iLength == 0)
                {
                    return sItem;
                }
                else if (iLength > 0)
                {
                    if (sItem.Length <= iLength)
                    {
                        return sItem;
                    }
                    else
                    {
                        return sItem.Substring(0, iLength); ;
                    }
                }
                else
                {
                    return "";
                }
            }
            else
            {
                return "";
            }
        }
        catch (Exception Error)
        {
            ZSBO.SBO_Application.SetStatusBarMessage("Left Error : " + Error.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            return "";
        }
    }

    //************************************************************************************************************
    // 
    //************************************************************************************************************
    public string Right(string Str, int Length)
    {
        try
        {
            string result;
            if (Str.Length < Length)
            {
                Length = Str.Length;
            }

            result = Str.Substring(Str.Length - Length, Length);
            return result;
        }
        catch (Exception Error)
        {
            ZSBO.SBO_Application.SetStatusBarMessage("Right Error : " + Error.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            return "";
        }
    }

    //************************************************************************************************************
    // 
    //************************************************************************************************************
    public void Load_ZPI02(string oKey)
    {
        SAPbobsCOM.Recordset oRecordset;
        string sQry;

        oRecordset = (SAPbobsCOM.Recordset)ZSBO.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        try
        {
            sQry = "SELECT DocEntry FROM [@ZPI02H] WHERE CANCELED='N' AND DocEntry=" + oKey + "";
            oRecordset.DoQuery(sQry);
            if (oRecordset.RecordCount == 0)
            {
                ZSBO.SBO_Application.SetStatusBarMessage(oKey + "의 공정 마스타를 찾을 수가 없습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            else
            {
                //ZPI002N oZPI002N = new ZPI002N();          //공정마스터
                //oZPI002N.LoadForm("ZPI002N", oRecordset.Fields.Item(0).Value.ToString().Trim());
            }
        }
        catch (Exception Error)
        {
            ZSBO.SBO_Application.SetStatusBarMessage("Load_ZPI02 Error - 공정 마스타를 여는 동안에 에러가 발생했습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
        }
    }

    //************************************************************************************************************
    // 
    //************************************************************************************************************
    public void Load_XPBHD(string oDocEntry)
    {
        SAPbobsCOM.Recordset oRecordset;
        string sQry;

        oRecordset = (SAPbobsCOM.Recordset)ZSBO.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        try
        {
            //sQry = "SELECT DocEntry FROM [@XPBHD] WHERE CANCELED='N' AND DocEntry=" + oDocEntry + "";
            //oRecordset.DoQuery(sQry);
            //if (oRecordset.RecordCount == 0)
            //{
            //    ZSBO.SBO_Application.SetStatusBarMessage(oDocEntry +"의 배합비 문서를 찾을 수가 없습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
            //}
            //else
            //{
            //    ZPF059N oZPF059N = new ZPF059N();          //배합비 문서
            //    oZPF059N.LoadForm("ZPF059N", oRecordset.Fields.Item(0).Value.ToString().Trim());

            //}
        }
        catch (Exception Error)
        {
            ZSBO.SBO_Application.SetStatusBarMessage("Load_XPBHD Error - 배합비 문서를 여는 동안에 에러가 발생했습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
        }
    }

    //************************************************************************************************************
    // 
    //************************************************************************************************************
    public void Load_ZPI04(string oDocEntry)
    {
        SAPbobsCOM.Recordset oRecordset;
        string sQry;

        oRecordset = (SAPbobsCOM.Recordset)ZSBO.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        try
        {
            sQry = "SELECT DocEntry FROM [@ZPI04H] WHERE CANCELED='N' AND DocEntry=" + oDocEntry + "";
            oRecordset.DoQuery(sQry);
            if (oRecordset.RecordCount == 0)
            {
                ZSBO.SBO_Application.SetStatusBarMessage(oDocEntry + "의 BOM 마스타를 찾을 수가 없습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            else
            {
                //ZPI004 oZPI004 = new ZPI004();          //BOM 마스터
                //oZPI004.LoadForm("ZPI004", oRecordset.Fields.Item(0).Value.ToString().Trim(), "", "");
            }
        }
        catch (Exception Error)
        {
            ZSBO.SBO_Application.SetStatusBarMessage("Load_ZPI04 Error - BOM 마스타를 여는 동안에 에러가 발생했습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
        }
    }

    //************************************************************************************************************
    // 
    //************************************************************************************************************
    public string ConvertTime(double oMin)
    {
        // 90(숫자) >> 0130 (1시간30분)
        double oHour;
        double sMin;
        string sTime = "";

        try
        {
            if (oMin == 0)
            {
                return "";
            }
            else if (oMin < 60)
            {
                return oMin.ToString();
            }
            else
            {
                oHour = Math.Truncate(oMin / 60);
                sMin = Math.Truncate(oMin % 60);
                if (sMin.ToString().Length == 0) sTime = oHour.ToString() + "00";
                else if (sMin.ToString().Length == 1) sTime = oHour.ToString() + "0" + sMin.ToString();
                else if (sMin.ToString().Length >= 2) sTime = oHour.ToString() + sMin.ToString().Substring(0, 2);


                return sTime;
            }

        }
        catch (Exception Error)
        {
            ZSBO.SBO_Application.SetStatusBarMessage("ZFN.ConvertTime Error - " + Error.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            return "";
        }
    }

    //************************************************************************************************************
    // 
    //************************************************************************************************************
    public string CDateDI(string sDate)
    {
        try
        {
            if (sDate.Length >= 8)
            {
                return sDate.Substring(0, 4) + "-" + sDate.Substring(4, 2) + "-" + sDate.Substring(6, 2);
            }
            else
            {
                ZSBO.SBO_Application.SetStatusBarMessage("CDateDI Error - 날짜길이체크", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                return sDate;
            }
        }
        catch (Exception Error)
        {
            ZSBO.SBO_Application.SetStatusBarMessage("ZFN.CDateDI Error : " + Error.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            return sDate;
        }
    }

    //************************************************************************************************************
    // 
    //************************************************************************************************************
    public double GetHour(double oMin)
    {
        double sTime = 0;

        try
        {
            if (oMin >= 60)
            {
                sTime = Math.Truncate(oMin / 60);
                return sTime;
            }
            else
            {
                return 0;
            }
        }
        catch (Exception Error)
        {
            ZSBO.SBO_Application.SetStatusBarMessage("ZFN.GetHour Error - " + Error.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            return 0;
        }
    }

    //************************************************************************************************************
    // 
    //************************************************************************************************************
    public double GetMin(double oMin)
    {
        double sTime = 0;

        try
        {
            if (oMin == 0)
            {
                return 0;
            }
            else if (oMin < 60)
            {
                return oMin;
            }
            else
            {
                sTime = Math.Truncate(oMin % 60);
                return sTime;
            }

        }
        catch (Exception Error)
        {
            ZSBO.SBO_Application.SetStatusBarMessage("ZFN.GetMin Error - " + Error.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            return 0;
        }
    }

    //************************************************************************************************************
    // 이메일 SMTP Server 생성
    //************************************************************************************************************
    public SmtpClient CreateSmtpServer(string username, string password)
    {
        string emailDomain = username.Split('@').GetValue(1).ToString();
        SmtpClient SmtpServer = null;

        switch (emailDomain)
        {
            case "gmail.com":
                SmtpServer = new SmtpClient("smtp.gmail.com");
                // SMTP 포트
                SmtpServer.Port = 587;
                // SSL 사용 여부
                SmtpServer.EnableSsl = true;

                SmtpServer.UseDefaultCredentials = false;
                SmtpServer.DeliveryMethod = System.Net.Mail.SmtpDeliveryMethod.Network;
                SmtpServer.Credentials = new System.Net.NetworkCredential(username, password);

                break;

            case "naver.com":
                SmtpServer = new SmtpClient("smtp.naver.com");
                // SMTP 포트
                SmtpServer.Port = 587;
                // SSL 사용 여부
                SmtpServer.EnableSsl = true;

                SmtpServer.UseDefaultCredentials = false;
                SmtpServer.DeliveryMethod = System.Net.Mail.SmtpDeliveryMethod.Network;
                SmtpServer.Credentials = new System.Net.NetworkCredential(username, password);

                break;
            case "bsncglobal.com":
                SmtpServer = new SmtpClient("smtps.hiworks.com");
                // SMTP 포트
                SmtpServer.Port = 587;
                // SSL 사용 여부
                SmtpServer.EnableSsl = true;

                SmtpServer.UseDefaultCredentials = false;
                SmtpServer.DeliveryMethod = System.Net.Mail.SmtpDeliveryMethod.Network;
                SmtpServer.Credentials = new System.Net.NetworkCredential(username, password);

                break;
            case "pointmobile.com":
                SmtpServer = new SmtpClient("smtps.pointmobile.com");

                // SMTP 포트
                SmtpServer.Port = 465;
                // SSL 사용 여부
                SmtpServer.EnableSsl = true;

                SmtpServer.UseDefaultCredentials = false;
                SmtpServer.DeliveryMethod = System.Net.Mail.SmtpDeliveryMethod.Network;
                SmtpServer.Credentials = new System.Net.NetworkCredential(username, password);
                ServicePointManager.SecurityProtocol = (SecurityProtocolType)3072;


                break;
        }

        return SmtpServer;
    }

    //************************************************************************************************************
    // 
    //************************************************************************************************************
    public void setComboList(SAPbouiCOM.ComboBox pCombo, string queryString, bool EmptyLine)
    {
        //ex : ZAPI.setComboList((SAPbouiCOM.ComboBox)oForm.Items.Item("uId").Specific, "SELECT CODE, NAME FROM [@ZMD101] ORDER BY CODE", false);
        SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)ZSBO.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        try
        {
            oRecordSet.DoQuery(queryString);
            if (EmptyLine & oRecordSet.RecordCount > 0) pCombo.ValidValues.Add("", "");

            while (!oRecordSet.EoF)
            {
                pCombo.ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString(), oRecordSet.Fields.Item(1).Value.ToString());
                oRecordSet.MoveNext();
            }
        }
        catch (Exception Error)
        {
            ZSBO.SBO_Application.SetStatusBarMessage(Error.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        }
    }

    //************************************************************************************************************
    // 
    //************************************************************************************************************
    public void RightClickEvent_Common(SAPbouiCOM.Form pForm, ref SAPbouiCOM.ContextMenuInfo oEventInfo, bool BubbleEvent)
    {
        SAPbouiCOM.CellPosition oCellPosition;
        try
        {
            if (oEventInfo.BeforeAction)
            {
                if (oEventInfo.ItemUID != "")
                {
                    SAPbouiCOM.BoFormItemTypes ItemType = pForm.Items.Item(oEventInfo.ItemUID).Type;
                    SAPbouiCOM.Matrix oMat;
                    SAPbouiCOM.Grid oGrid;

                    // 복사
                    pForm.EnableMenu("772",
                        ItemType == SAPbouiCOM.BoFormItemTypes.it_EDIT
                        | ItemType == SAPbouiCOM.BoFormItemTypes.it_STATIC
                        | ItemType == SAPbouiCOM.BoFormItemTypes.it_EXTEDIT
                        | ItemType == SAPbouiCOM.BoFormItemTypes.it_MATRIX
                        | ItemType == SAPbouiCOM.BoFormItemTypes.it_GRID
                        | ItemType == SAPbouiCOM.BoFormItemTypes.it_BUTTON
                        );
                    // 테이블복사
                    pForm.EnableMenu("784",
                        ItemType == SAPbouiCOM.BoFormItemTypes.it_MATRIX
                        | ItemType == SAPbouiCOM.BoFormItemTypes.it_GRID
                        );

                    int oRow = oEventInfo.Row;

                    if (oRow > 0)
                    {
                        // 클릭한 라인 행선택
                        switch (ItemType)
                        {
                            case SAPbouiCOM.BoFormItemTypes.it_MATRIX:
                                oMat = (SAPbouiCOM.Matrix)pForm.Items.Item(oEventInfo.ItemUID).Specific;
                                oMat.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single;
                                oMat.ClearSelections();
                                oMat.SelectRow(oRow, true, false);

                                break;

                            case SAPbouiCOM.BoFormItemTypes.it_GRID:
                                oGrid = (SAPbouiCOM.Grid)pForm.Items.Item(oEventInfo.ItemUID).Specific;
                                oGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single;
                                oGrid.Rows.SelectedRows.Clear();
                                oGrid.Rows.SelectedRows.Add(oRow);
                                break;
                        }
                    }
                }
            }
        }
        catch (Exception Error)
        {
            ZSBO.SBO_Application.SetStatusBarMessage("ZAPI.RightClickEvent_Common Error : " + Error.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        }
    }


    //************************************************************************************************************
    // 서버 오늘날짜
    //************************************************************************************************************
    public string getToday()
    {
        SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)ZSBO.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        string result = null;
        try
        {
            oRecordSet.DoQuery("SELECT CONVERT(NVARCHAR, GETDATE(),112) 'TODAY'");
            result = oRecordSet.Fields.Item(0).Value.ToString();
            oRecordSet = null;
            return result;
        }
        catch (Exception Error)
        {
            oRecordSet = null;
            ZSBO.SBO_Application.SetStatusBarMessage(Error.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            return null;
        }
    }

    //************************************************************************************************************
}
