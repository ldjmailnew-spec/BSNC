//************************************************************************************************************
//  File            : ZSetEF.cls
//  Module          : 공통
//  Desc            : EventFilter
//  Create Date     : 2024.10.31
//  Creator         : Seo Young Mi
//  Memo            :
//
//  Copyright (c) BSNC
//************************************************************************************************************
using System;
using System.Windows.Forms;

public class ZSetEF
{
    //************************************************************************************************************
    // 
    //************************************************************************************************************
    public void SetEventFilter()
    {
        SAPbouiCOM.EventFilters oFilters = new SAPbouiCOM.EventFilters();

        try
        {
            //----------------------------------------------------------------------------
            // [ZPMKL3370] 재고이전요청등록
            (oFilters.Add(SAPbouiCOM.BoEventTypes.et_MENU_CLICK)).AddEx("ZPMKL3370");
            (oFilters.Add(SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED)).AddEx("ZPMKL3370");
            (oFilters.Add(SAPbouiCOM.BoEventTypes.et_COMBO_SELECT)).AddEx("ZPMKL3370");
            (oFilters.Add(SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST)).AddEx("ZPMKL3370");
            (oFilters.Add(SAPbouiCOM.BoEventTypes.et_KEY_DOWN)).AddEx("ZPMKL3370");
            (oFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_RESIZE)).AddEx("ZPMKL3370");
        
            //----------------------------------------------------------------------------
            ZSBO.SBO_Application.SetFilter(oFilters);
            //----------------------------------------------------------------------------
        }
        catch (Exception Error)
        {
            ZSBO.SBO_Application.SetStatusBarMessage("setEventFilter Error : " + Error.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        }
        finally
        {
            oFilters = null;
        }
    }

    //************************************************************************************************************
}
