//************************************************************************************************************
//  File             : ZMC201.cs
//  Module           : 판매오더 불러오기 팝업창
//  Desc             : 판매오더 불러오기 팝업창
//  FormType         : ZMC201
//  Create Date      : 2026.10.12
//  Creator          : LEE DAE JIN
//  memo             : 
//  Copyright (c) BSNC
//************************************************************************************************************

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

class ZMC202A
{

    private static SAPbouiCOM.Form oForm;
    private static SAPbouiCOM.Grid oGRD1;


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
                        // if (ChkInData() == true)
                        // {
                        if (ZSBO.SBO_Application.MessageBox("조회하시겠습니까?", 1, "예", "아니오", "") == 1)
                        {
                           // LoadData("S");
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
                    //if (pVal.Before_Action & pVal.ItemUID == "1")
                    //{
                    //}
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
            //SetFormBase();
            //LoadData("S");

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

            // 옵션 버튼 컨트롤 가져오기 (OptionBtn은 SAPbouiCOM.OptionBtn 타입)
            SAPbouiCOM.OptionBtn oOptYes = (SAPbouiCOM.OptionBtn)oForm.Items.Item("OPT1").Specific;

            // 기본값으로 체크 설정
            oOptYes.Selected = true;

            // 오더기간 설정 
            oForm.DataSources.UserDataSources.Add("DOCSDATE", SAPbouiCOM.BoDataType.dt_DATE, 10);
            oForm.DataSources.UserDataSources.Add("DOCEDATE", SAPbouiCOM.BoDataType.dt_DATE, 10);

            // Binding (스크핀페인터랑 소스랑 연결하기)
            ((SAPbouiCOM.EditText)oForm.Items.Item("DOCSDATE").Specific).DataBind.SetBound(true, "", "DOCSDATE");
            ((SAPbouiCOM.EditText)oForm.Items.Item("DOCEDATE").Specific).DataBind.SetBound(true, "", "DOCEDATE");

            //초기화 
            oForm.DataSources.UserDataSources.Item("DOCSDATE").ValueEx = "";
            oForm.DataSources.UserDataSources.Item("DOCEDATE").ValueEx = "";


        }
        catch (Exception Error)
        {
            ZSBO.SBO_Application.SetStatusBarMessage("CreateItems_Error : " + Error.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        }

    }


}