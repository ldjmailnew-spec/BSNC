//****************************************************************************
//  File       :   ZFN.cls
//  Module     :   공통함수
//  Desc       :   공통함수
//  FormType   :   ZFN  
//  Create Date   : 2014.05.20
//  Creator       : Kim Kyung Ho
//****************************************************************************

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Net;
using System.Data;
using System.IO;
using ExcelDataReader;



////Automanagerd Enum----------------------------------
//public enum eAttrib { Editable, Visible };
//public enum eFormMode { All, Add, Find, Ok, View };
//public enum eBehavior { True, False, Default };
////---------------------------------------------------

public class ZFN
{
    public static void ComboBoxSet(ref SAPbouiCOM.ComboBox oCombo, string sQry)
    {
        SAPbobsCOM.Recordset oRecordset;

        try
        {
            oRecordset = (SAPbobsCOM.Recordset)ZSBO.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            oRecordset.DoQuery(sQry);
            if (oRecordset.RecordCount > 0)
            {
                do
                {
                    oCombo.ValidValues.Add(oRecordset.Fields.Item(0).Value.ToString().Trim(), oRecordset.Fields.Item(1).Value.ToString().Trim());
                    oRecordset.MoveNext();
                } while (oRecordset.EoF == false);
            }

        }
        catch (Exception Error)
        {
            ZSBO.SBO_Application.SetStatusBarMessage("ZFN.ComboBoxSet Error : " + Error.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        }
    }
    public static void ComboBoxSet(ref SAPbouiCOM.Column oColumn, string sQry)
    {
        SAPbobsCOM.Recordset oRecordset;

        try
        {
            oRecordset = (SAPbobsCOM.Recordset)ZSBO.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            oRecordset.DoQuery(sQry);
            if (oRecordset.RecordCount > 0)
            {
                do
                {
                    oColumn.ValidValues.Add(oRecordset.Fields.Item(0).Value.ToString().Trim(), oRecordset.Fields.Item(1).Value.ToString().Trim());
                    oRecordset.MoveNext();
                } while (oRecordset.EoF == false);
            }
            oColumn.DisplayDesc = true;

        }
        catch (Exception Error)
        {
            ZSBO.SBO_Application.SetStatusBarMessage("ZFN.ComboBoxSet Error : " + Error.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        }
    }

    public static void ComboBoxDelete(SAPbouiCOM.Form oForm, string sItem)
    {
        SAPbouiCOM.ComboBox oCombo;

        try
        {
            oCombo = (SAPbouiCOM.ComboBox)oForm.Items.Item(sItem).Specific;

            if (oCombo.ValidValues.Count > 0)
            {
                for (int i = oCombo.ValidValues.Count - 1; i >= 0; i--)
                {
                    oCombo.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
                }
            }
        }
        catch (Exception Error)
        {
            ZSBO.SBO_Application.SetStatusBarMessage("ZFN.ComboBoxDelete Error : " + Error.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        }
    }
    public static void ComboBoxDelete(ref SAPbouiCOM.ComboBox oCombo)
    {
        try
        {
            if (oCombo.ValidValues.Count > 0)
            {
                for (int i = oCombo.ValidValues.Count - 1; i >= 0; i--)
                {
                    oCombo.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
                }
            }
        }
        catch(Exception Error)
        {
            ZSBO.SBO_Application.SetStatusBarMessage("ZFN.ComboBoxDelete Error : " + Error.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        }

    }

    public static void ComboBoxDelete(ref SAPbouiCOM.Column oColumn)
    {
        try
        {
            if (oColumn.ValidValues.Count > 0)
            {
                for (int i = oColumn.ValidValues.Count - 1; i >= 0; i--)
                {
                    oColumn.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
                }
            }
        }
        catch (Exception Error)
        {
            ZSBO.SBO_Application.SetStatusBarMessage("ZFN.ComboBoxDelete Error : " + Error.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        }
    }

    public static string CDateUI(string sDate)
    {
        try
        {
            if (sDate.Length >= 10)
            {
                return sDate.Substring(0, 10).Replace("-", "");
            }
            else
            {
                ZSBO.SBO_Application.SetStatusBarMessage("CDateUI Error - 날짜길이체크", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                return sDate;
            }
        }
        catch (Exception Error)
        {
            ZSBO.SBO_Application.SetStatusBarMessage("ZFN.CDateUI Error : " + Error.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            return sDate;
        }
    }
    public static string CDateDI(string sDate)
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

    public static double CDbl(string sStr)
    {
        SAPbobsCOM.Recordset oRecordset;
        string sQry;

        try
        {
            oRecordset = (SAPbobsCOM.Recordset)ZSBO.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            if (sStr.Trim() == "" || sStr.Trim() == null)
            {
                return 0;
            }
            else
            {
                return Convert.ToDouble(sStr.Trim());
            }
        }
        catch (Exception Error)
        {
            ZSBO.SBO_Application.SetStatusBarMessage("ZFN.CDbl Error - " + Error.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            return 0;
        }
    }

    public static int CInt32(double sStr)
    {
        try
        {
            return CInt32(String.Format("{0:0}", sStr));
        }
        catch (Exception Error)
        {
            ZSBO.SBO_Application.SetStatusBarMessage("ZFN.CInt32 Error : " + Error.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            return 0;
        }
    }

    public static int CInt32(string sStr)
    {
        try
        {
            if (sStr.Trim() == "" || sStr.Trim() == null || sStr.Trim() == "0.0")
            {
                return 0;
            }
            else
            {
                return Convert.ToInt32(sStr.Trim());
            }
        }
        catch (Exception Error)
        {
            ZSBO.SBO_Application.SetStatusBarMessage("ZFN.CInt32 Error - " + Error.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            return 0;
        }
    }

    public static Int64 CInt64(double sStr)
    {
        try
        {
            return CInt64(String.Format("{0:0}", sStr));
        }
        catch (Exception Error)
        {
            ZSBO.SBO_Application.SetStatusBarMessage("ZFN.CInt32 Error : " + Error.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            return 0;
        }
    }

    public static Int64 CInt64(string sStr)
    {
        try
        {
            if (sStr.Trim() == "" || sStr.Trim() == null || sStr.Trim() == "0.0")
            {
                return 0;
            }
            else
            {
                return Convert.ToInt64(sStr.Trim());
            }
        }
        catch (Exception Error)
        {
            ZSBO.SBO_Application.SetStatusBarMessage("ZFN.CInt64 Error - " + Error.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            return 0;
        }
    }

    public static object GetReData(string selectColumn, string tableName)
    {
        try
        {
            return GetReData(selectColumn, tableName, "", "");
        }
        catch (Exception Error)
        {
            ZSBO.SBO_Application.SetStatusBarMessage("ZFN.GetReData Error : " + Error.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            return null;
        }
    }
    public static object GetReData(string selectColumn, string tableName, string conditionColumn, string conditionValue)
    {
        try
        {
            return GetReData(selectColumn, tableName, conditionColumn, conditionValue, "");
        }
        catch (Exception Error)
        {
            ZSBO.SBO_Application.SetStatusBarMessage("ZFN.GetReData Error : " + Error.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            return null;
        }

    }
    public static string GetReData(string selectColumn, string tableName, string conditionColumn, string conditionValue, string andSentence)
    {
        string sQry = "";
        string Value = "";
        SAPbobsCOM.Recordset oRecordset = null;

        try
        {
            oRecordset = (SAPbobsCOM.Recordset)ZSBO.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            if (conditionColumn == "" && conditionValue == "")
            {
                sQry = string.Format(@"SELECT {0} FROM {1} ", selectColumn, tableName);
            }
            else
            {
                sQry = string.Format(@"SELECT {0} FROM {1} WHERE {2} = {3}{4}", selectColumn, tableName, conditionColumn, conditionValue, andSentence);
            }

            oRecordset.DoQuery(sQry);

            oRecordset.MoveLast();

            Value = oRecordset.Fields.Item(0).Value.ToString();
            return Value;
        }
        catch (Exception Error)
        {
            ZSBO.SBO_Application.SetStatusBarMessage("ZFN.GetReData Error : " + Error.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            return "";
        }
    }

    public static void MatrixKeyMove(ref SAPbouiCOM.Form oForm, string oItem, int CharPressed, string oCol, int oRow)
    {

        SAPbouiCOM.Matrix oMatrix;

        try
        {
            oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item(oItem).Specific;

            if (CharPressed == 38)           // 커서를 한 컬럼 위에 놓는다.(38)
            {
                if (oMatrix.RowCount == 1) { }
                else if (oRow == 1) { }
                else if (oRow <= oMatrix.VisualRowCount)
                {
                    oMatrix.Columns.Item(oCol).Cells.Item(oRow - 1).Click(SAPbouiCOM.BoCellClickType.ct_Regular, 0);
                }
            }
            else if (CharPressed == 40)         // 커서를 한 컬럼 아래에 놓는다.(40)
            {
                if (oMatrix.RowCount == 1) { }
                else if (oRow == oMatrix.VisualRowCount) { }
                else if (oRow <= oMatrix.VisualRowCount - 1)
                {
                    oMatrix.Columns.Item(oCol).Cells.Item(oRow + 1).Click(SAPbouiCOM.BoCellClickType.ct_Regular, 0);
                }
            }

        }
        catch (Exception Error)
        {
            ZSBO.SBO_Application.SetStatusBarMessage("MatrixKeyMove Error - " + Error.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        }
    }

    public static void AutoManaged(ref SAPbouiCOM.Form oForm, string itemsName, eAttrib Attrib, eFormMode autoFormMode, eBehavior Behavior)
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

    /// <summary>
    /// SetEnableMenus
    /// </summary>
    /// <param name="bPreview">인쇄[미리보기]</param>
    /// <param name="bPrint">인쇄[출력]</param>
    /// <param name="bFind">문서찾기</param>
    /// <param name="bAdd">문서추가</param>
    /// <param name="bNextRecord">다음레코드</param>
    /// <param name="bPreviousRecord">이전레코드</param>
    /// <param name="bFirstRecord">처음레코드</param>
    /// <param name="bLastRecord">마지막레코드</param>
    /// <param name="bRemove">문서제거</param>
    /// <param name="bCancel">문서취소</param>
    /// <param name="bClose">문서닫기</param>
    /// <param name="bDuplicate">문서복제</param>
    /// <param name="bRowAdd">행추가</param>
    /// <param name="bDeleteRow">행삭제</param>
    /// <param name="bRowClose">행닫기</param>
    /// <param name="bRowDuplicate">행복제</param>
    //'-----------------------------------------------------------------------------------------
    //'   메뉴활성화
    //'   네비게이션 컨트롤 관련 보이기/감추기 함수
    //'   -> 미리보기, 출력, 행삭제, 찾기, 추가, 다음, 이전, 맨처음, 맨끝, 취소
    //'-----------------------------------------------------------------------------------------
    public static void SetEnableMenus(SAPbouiCOM.Form eForm,
                                      bool bPreview,
                                      bool bPrint,
                                      bool bFind,
                                      bool bAdd,
                                      bool bNextRecord,
                                      bool bPreviousRecord,
                                      bool bFirstRecord,
                                      bool bLastRecord,
                                      bool bRemove,
                                      bool bCancel,
                                      bool bClose,
                                      bool bDuplicate,
                                      bool bRowAdd,
                                      bool bDeleteRow,
                                      bool bRowClose,
                                      bool bRowDuplicate)
    {
        try
        {
            eForm.EnableMenu("519", bPreview);          // 인쇄[미리보기]
            eForm.EnableMenu("520", bPrint);            // 인쇄[출력]        
            eForm.EnableMenu("1281", bFind);            // 문서찾기
            eForm.EnableMenu("1282", bAdd);             // 문서추가        
            eForm.EnableMenu("1288", bNextRecord);      // 다음
            eForm.EnableMenu("1289", bPreviousRecord);  // 이전
            eForm.EnableMenu("1290", bFirstRecord);     // 맨처음
            eForm.EnableMenu("1291", bLastRecord);      // 맨끝
            eForm.EnableMenu("1283", bRemove);          // 문서제거(데이터 삭제시 사용)        
            eForm.EnableMenu("1284", bCancel);          // 문서취소
            eForm.EnableMenu("1286", bClose);           // 문서닫기
            eForm.EnableMenu("1287", bDuplicate);       // 문서복제
            eForm.EnableMenu("1292", bRowAdd);          // 행추가        
            eForm.EnableMenu("1293", bDeleteRow);       // 행삭제
            eForm.EnableMenu("1299", bRowClose);        // 행닫기
            eForm.EnableMenu("1294", bRowDuplicate);    // 행복제
        }
        catch (Exception Error)
        {
            ZSBO.SBO_Application.SetStatusBarMessage("ZFN.SetEnableMenus Error : " + Error.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        }
    }


    public static string Left(string sItem, int iLength)
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
            ZSBO.SBO_Application.SetStatusBarMessage("ZFN.Left Error : " + Error.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            return "";
        }
    }

    public static string Right(string Str, int Length)
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
            ZSBO.SBO_Application.SetStatusBarMessage("ZFN.Right Error : " + Error.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            return "";
        }
    }


    public static double GetHour(double oMin)
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

    public static double GetMin(double oMin)
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

    public static string ConvertTime(double oMin)
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
    /// <summary>
    /// ConvertTime2
    /// </summary>
    /// <param name="oTime"> (0000) 문자시간 (분으로 환산[숫자])</param>
    public static double ConvertTime2(string oTime)
    {
        // 0130 (시간) >> 90(숫자) 
        
        double dTime;
        string sTime = "";

        try
        {
            sTime = Convert.ToDouble(oTime).ToString("0000");
            dTime = Convert.ToDouble(Left(sTime, 2)) * 60 + Convert.ToDouble(Right(sTime, 2));
            return dTime;

        }
        catch (Exception Error)
        {
            ZSBO.SBO_Application.SetStatusBarMessage("ZFN.ConvertTime Error - " + Error.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            return 0;
        }
    }

    public static string CZero(double Value)
    {
        try
        {
            if (Value == 0) return "0";
            else return Value.ToString();
        }
        catch (Exception Error)
        {
            ZSBO.SBO_Application.SetStatusBarMessage("ZFN.CZero Error : " + Error.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            return "0";
        }
    }

    public static double PointCheck(double value, string FieldName, string sTable, string sFlag)
    {
        //////////////////////////////////////////////////////////////////////////
        //FieldName >>> U_ 제외, 
        //sTable    >>> @포함, 
        //sFlag     >>> 버림:F, 반올림:R, 올림:C , 세금규칙:V
        //////////////////////////////////////////////////////////////////////////

        SAPbobsCOM.Recordset oRecordset;
        try
        {
            oRecordset = (SAPbobsCOM.Recordset)ZSBO.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            string RndRule;
            int iDigit = 0;

            if (sFlag == "V")
            {
                RndRule = GetReData("TaxRndRule", "[OADM]").ToString();
            }
            else
            {
                RndRule = sFlag;
            }

            oRecordset.DoQuery("SELECT TypeID, EditType,editSize FROM [CUFD] WHERE aliasid = '" + FieldName + "' AND tableid = '" + sTable + "'");
            if (oRecordset.RecordCount > 0)
            {
                switch (oRecordset.Fields.Item("TypeID").Value.ToString() + oRecordset.Fields.Item("EditType").Value.ToString())
                {
                    case "BS":  //금액
                        iDigit = Convert.ToInt32(GetReData("SumDec", "[OADM]"));
                        break;
                    case "BP":  //가격
                        iDigit = Convert.ToInt32(GetReData("PriceDec", "[OADM]"));
                        break;
                    case "BR":  //환율
                        iDigit = Convert.ToInt32(GetReData("RateDec", "[OADM]"));
                        break;
                    case "BQ":  //수량
                        iDigit = Convert.ToInt32(GetReData("QtyDec", "[OADM]"));
                        break;
                    case "B%":  //백분율
                        iDigit = Convert.ToInt32(GetReData("PercentDec", "[OADM]"));
                        break;
                    case "BM":  //단위
                        iDigit = Convert.ToInt32(GetReData("MeasureDec", "[OADM]"));
                        break;

                }
            }
            return PointCalc(value, iDigit, RndRule);
        }
        catch (Exception Error)
        {
            ZSBO.SBO_Application.SetStatusBarMessage("ZFN.PointCheck Error : " + Error.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            return 0;
        }
    }

    /// <summary>
    /// PointCalc
    /// </summary>
    /// <param name="dValue"> 값</param>
    /// <param name="dValue"> 자릿수</param>
    /// <param name="dValue"> R:반올림/C:올림/F:버림</param>
    public static double PointCalc(double dValue, int iDigits, string sRule)
    {
        double dCoef = System.Math.Pow(10, iDigits);
        double dResult = 0;
        try
        {
            switch (sRule)
            {
                case "R":      //반올림
                    if (dValue > 0)
                    {
                        dResult = Math.Floor((dValue * dCoef) + 0.5) / dCoef;
                    }
                    else
                    {
                        dResult = Math.Ceiling((dValue * dCoef) - 0.5) / dCoef;
                    }
                    break;

                case "C":      //올림
                    if (dValue > 0)
                    {
                        dResult = Math.Ceiling(dValue * dCoef) / dCoef;
                    }
                    else
                    {
                        dResult = Math.Floor(dValue * dCoef) / dCoef;
                    }
                    break;

                case "F":      //버림
                    if (dValue > 0)
                    {
                        dResult = Math.Floor(dValue * dCoef) / dCoef;
                    }
                    else
                    {
                        dResult = Math.Ceiling(dValue * dCoef) / dCoef;
                    }
                    break;

            }
            return dResult;
        }
        catch (Exception Error)
        {
            ZSBO.SBO_Application.SetStatusBarMessage("ZFN.PointCalc Error : " + Error.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            return 0;
        }
    }
    public static bool Authority(string Code, string FldCode)
    {
        SAPbobsCOM.Recordset oRecordset;
        string sQry;
        bool bCheck;

        try
        {
            oRecordset = (SAPbobsCOM.Recordset)ZSBO.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            bCheck = false;

            sQry = "SELECT U_FldValue "
                 + " FROM [@ZPSDD] L INNER JOIN [@ZPSDH] H ON L.Code = H.Code "
                 + " WHERE H.Code = '" + Code + "'"
                 + " AND L.U_FldCode = '" + FldCode + "'";

            oRecordset.DoQuery(sQry);
            if (oRecordset.RecordCount > 0)
            {
                for (int i = 0; i <= oRecordset.RecordCount - 1; i++)
                {
                    if (oRecordset.Fields.Item("U_FldValue").Value.ToString().Trim().ToUpper() == ZSBO.oCompany.UserName.ToString().ToUpper())
                    {
                        bCheck = true;
                    }
                    oRecordset.MoveNext();
                }
            }

            if (bCheck == true)
            {
                return true;
            }
            else
            {
                return false;
            }
        }
        catch (Exception Error)
        {
            ZSBO.SBO_Application.SetStatusBarMessage("Authority Error - " + Error.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            return false;
        }
    }

    /// <summary>  날짜사이의시간간격의숫자를반환합니다.
    /// </summary>
    /// <param name="Interval">y-m-d h:n:s:ms</param>
    /// <param name="Date1"></param>
    /// <param name="Date2"></param>
    /// <returns></returns>
    public static double DateDiff(string Interval, DateTime Date1, DateTime Date2)
    {
        double diff = 0;
        TimeSpan ts = Date2 - Date1;

        switch (Interval.ToLower())
        {
            case "y":
                ts = DateTime.Parse(Date2.ToString("yyyy-01-01")) - DateTime.Parse(Date1.ToString("yyyy-01-01"));
                diff = Convert.ToDouble(ts.TotalDays / 365);
                break;
            case "m":
                ts = DateTime.Parse(Date2.ToString("yyyy-MM-01")) - DateTime.Parse(Date1.ToString("yyyy-MM-01"));
                diff = Convert.ToDouble((ts.TotalDays / 365) * 12);
                break;
            case "d":
                ts = DateTime.Parse(Date2.ToString("yyyy-MM-dd")) - DateTime.Parse(Date1.ToString("yyyy-MM-dd"));
                diff = ts.Days;
                break;
            case "h":
                ts = DateTime.Parse(Date2.ToString("yyyy-MM-dd HH:00:00")) - DateTime.Parse(Date1.ToString("yyyy-MM-dd HH:00:00"));
                diff = ts.TotalHours;
                break;
            case "n":
                ts = DateTime.Parse(Date2.ToString("yyyy-MM-dd HH:mm:00")) - DateTime.Parse(Date1.ToString("yyyy-MM-dd HH:mm:00"));
                diff = ts.TotalMinutes;
                break;
            case "s":
                ts = DateTime.Parse(Date2.ToString("yyyy-MM-dd HH:mm:ss")) - DateTime.Parse(Date1.ToString("yyyy-MM-dd HH:mm:ss"));
                diff = ts.TotalSeconds;
                break;
            case "ms":
                diff = ts.TotalMilliseconds;
                break;
        }

        return diff;
    }

    /// <summary>  전기기간 잠금 체크
    /// </summary>
    /// <param name="Date">년월체크 'YYYYMM'</param>
    /// <returns></returns>
    public static bool PostingCheck(string Date)
    {
        SAPbobsCOM.Recordset oRecordset;
        string sQry;

        try
        {
            oRecordset = (SAPbobsCOM.Recordset)ZSBO.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            sQry = "SELECT PeriodStat FROM OFPR WHERE replace(Code, '-', '') = SUBSTRING(REPLACE(" + Date + ", '-', '' ), 1, 6)";
            oRecordset.DoQuery(sQry);

            if (oRecordset.RecordCount > 0)
            {
                if (oRecordset.Fields.Item(0).Value.ToString().Trim() != "Y")
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            else
            {
                return false;
            }

        }
        catch (Exception Error)
        {
            ZSBO.SBO_Application.SetStatusBarMessage("SetBPLId Error - " + Error.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            return false;
        }
    }

    public static string Left_Zero(string sValue, int iDigit)
    {
        int Length;
        string Zero = "";
        try
        {
            Length = sValue.Length;
            if (Length < iDigit)
            {
                for (int i = 1; i <= (iDigit - Length); i++)
                {
                    Zero = Zero + "0";
                }
                return Zero + sValue;
            }
            else
            {
                return sValue;
            }

        }
        catch (Exception Error)
        {
            ZSBO.SBO_Application.SetStatusBarMessage("SetBPLId Error - " + Error.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            return "";
        }
    }

    /// <summary>  G/L 계정(고급) - 계정가져오기
    /// </summary>
    /// <param name="sItemCode"> 품목코드 </param>
    /// <param name="sGubn">계정구분 "1" - 비용계정, "2" - Wip계정, "3" - 비용임시계정</param>
    /// <returns></returns>
    public static string OGAR_AcctCode(string sItemCode, string sGubn)
    {
        SAPbobsCOM.Recordset oRecordset;
        SAPbobsCOM.Recordset pRecordset;
        string sQry;

        oRecordset = (SAPbobsCOM.Recordset)ZSBO.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        pRecordset = (SAPbobsCOM.Recordset)ZSBO.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        try
        {
            sQry = "select ItmsGrpCod from oitm where ItemCode = '" + sItemCode + "'";
            oRecordset.DoQuery(sQry);

            if (oRecordset.RecordCount > 0)
            {
                sQry = " SELECT DFLTEXPN AS'비용계정', WIPACCT AS'WIP계정', EXPCLRACT AS'비용임시계정'"
                     + " FROM [OGAR]"
                     + " WHERE Year = '" + DateTime.Today.Year.ToString() + "'"
                     + " AND ItmsGrpCod = '" + oRecordset.Fields.Item(0).Value.ToString().Trim() + "'";
                pRecordset.DoQuery(sQry);

                if (pRecordset.RecordCount > 0)
                {
                    if (sGubn == "1")
                    {
                        if (oRecordset.Fields.Item(0).Value.ToString().Trim() == "100" ||
                            oRecordset.Fields.Item(0).Value.ToString().Trim() == "101" )
                        {
                            return pRecordset.Fields.Item(0).Value.ToString().Trim();
                        }
                        else
                        {
                            return pRecordset.Fields.Item(1).Value.ToString().Trim();
                        }
                    }
                    else if (sGubn == "2")
                    {
                        return pRecordset.Fields.Item(1).Value.ToString().Trim();
                    }
                    else if (sGubn == "3")
                    {
                        return pRecordset.Fields.Item(2).Value.ToString().Trim();
                    }
                    else
                    {
                        ZSBO.SBO_Application.SetStatusBarMessage("해당구분이 없습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                        return "";
                    }
                }
                else
                {
                    ZSBO.SBO_Application.SetStatusBarMessage("G/L계정 계정정보가 없습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
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
            ZSBO.SBO_Application.SetStatusBarMessage("OGAR_AcctCode Error - " + Error.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            return "";
        }
    }

    public static double CalcAmt(double oAmt1, string oOperation, double oAmt2, int oDec, string oCurr)
    {

        SAPbobsCOM.Recordset oRecordset;
        string sQry;

        oRecordset = (SAPbobsCOM.Recordset)ZSBO.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        try
        {
            sQry = "SELECT [dbo].[Fn_GetDIVIDE2] (" + oAmt1 + ", " + oAmt2 + ", '" + oOperation + "'," + oDec + ", '" + oCurr + "')";
            oRecordset.DoQuery(sQry);

            if (oRecordset.RecordCount == 1)
            {
                if (oRecordset.Fields.Item(0).Value.ToString().Trim() == "")
                {
                    return 0;
                }
                else
                {
                    return ZFN.CDbl(oRecordset.Fields.Item(0).Value.ToString().Trim());
                }
            }
            else
            {
                return 0;
            }
        }
        catch (Exception Error)
        {
            ZSBO.SBO_Application.SetStatusBarMessage("CalcAmt Error - " + Error.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            return 0;
        }
    }

    /// <summary>
    /// UserDataSource 타입 별 바인딩
    /// </summary>
    /// <param name="oForm">대상 폼</param>
    public static void BindUserDataSource(ref SAPbouiCOM.Form oForm)
    {
        SAPbouiCOM.EditText oEditText;
        SAPbouiCOM.ComboBox oCombo;
        SAPbouiCOM.CheckBox oCheck;
        SAPbouiCOM.Item oItem;
        SAPbouiCOM.UserDataSources oUDS = null;

        try
        {
            // UserDataSource 할당
            oUDS = oForm.DataSources.UserDataSources;

            // 반복문으로 타입 별 바인딩
            foreach (SAPbouiCOM.UserDataSource uds in oUDS)
            {
                if (uds.UID != "FolderDS") // 폴더 관리용 소스 제외
                {
                    oItem = oForm.Items.Item(uds.UID);

                    if (oItem.Type == SAPbouiCOM.BoFormItemTypes.it_EDIT
                        | oItem.Type == SAPbouiCOM.BoFormItemTypes.it_EXTEDIT)
                    {
                        oEditText = oItem.Specific as SAPbouiCOM.EditText;
                        oEditText.DataBind.SetBound(true, "", uds.UID);
                    }
                    else if (oItem.Type == SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX)
                    {
                        oCombo = oItem.Specific as SAPbouiCOM.ComboBox;
                        oCombo.DataBind.SetBound(true, "", uds.UID);
                    }
                    else if (oItem.Type == SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX)
                    {
                        oCheck = oItem.Specific as SAPbouiCOM.CheckBox;
                        oCheck.DataBind.SetBound(true, "", uds.UID);
                    }
                }
            }
        }
        catch (Exception Error)
        {
            ZSBO.SBO_Application.SetStatusBarMessage("ZFN.BindUserdataSource Error - " + Error.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        }
    }

    /// <summary>
    /// Check Before Release Object
    /// </summary>
    /// <param name="obj">Target Object</param>
    public static void ReleaseComObject(object obj)
    {
        try
        {
            if (obj != null)
            {
                Marshal.ReleaseComObject(obj);
                obj = null;
            }
        }
        catch (Exception ex)
        {
            obj = null;
        }
        //finally
        //{
        //    GC.Collect();
        //}
    }

    //************************************************************************************************************
    // 파일 다운로드(파일이름)
    //************************************************************************************************************
    public static void OpenFileExcell(object fileName)
    {
        var pFileDlg = new SaveFileDialog();
        string strFullPathFile = "";

        pFileDlg.Filter = "Excel Files(*.xlsx)|*.xlsx|All Files(*.*)|*.*";
        pFileDlg.FileName = fileName.ToString();

        // 경로 가져오기
        if (pFileDlg.ShowDialog(WindowWrapper.GetForegroundWindowWrapper()) == DialogResult.OK)
        {
            strFullPathFile = pFileDlg.FileName;

            string sPath = null;
            sPath = Application.StartupPath;

            string path = strFullPathFile;

            string str, fileName1;
            string remoteUri = sPath + @"\XLS_Template\";
            string myStringWebResource = null;

            WebClient myWebClient = new WebClient();
            // Concatenate the domain with the Web resource filename.
            myStringWebResource = remoteUri + fileName;

            try
            {
                ZSBO.SBO_Application.StatusBar.SetText(@"다운로드 완료되었습니다. " + path, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                myWebClient.DownloadFile(myStringWebResource, path);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, ex.GetType().FullName, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }
    }

    public static DataTable GetExcelDataSet(string filePath)
    {
        var ResDs = new DataSet();
        var ResDt = new DataTable();

        try
        {
            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
            {
                var readerConf = new ExcelReaderConfiguration()
                {
                    FallbackEncoding = Encoding.GetEncoding(1252)
                };

                using (var reader = ExcelReaderFactory.CreateReader(stream, readerConf))
                {
                    ResDs = reader.AsDataSet(new ExcelDataSetConfiguration()
                    {
                        UseColumnDataType = true,
                        ConfigureDataTable = (tableReader) => new ExcelDataTableConfiguration()
                        {
                            UseHeaderRow = true
                        }
                    });

                    ResDt = ResDs.Tables[0];
                }
            }

            return ResDt;
        }
        catch (Exception Error)
        {
            ZSBO.SBO_Application.SetStatusBarMessage("GetExcelDataSet Error : " + Error.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            return null;
        }
    }

}