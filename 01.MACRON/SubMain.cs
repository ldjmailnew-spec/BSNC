//****************************************************************************
//
//  File:      SubMain.cs
//
//  Copyright (c) BS&C
//
//****************************************************************************

using System;
using System.Windows.Forms;
using System.Collections;

class SubMain {
    //public static System.Collections.ArrayList sapForms = new System.Collections.ArrayList();
    //public static Hashtable mForm = new Hashtable();

    [STAThreadAttribute()]
    //[STAThread]
    static void Main()
    {
        try
        {
            ZSBO oZSBO = null;
            oZSBO = new ZSBO();
            System.Windows.Forms.Application.Run();
        }
        catch (Exception E)
        {
            MessageBox.Show(E.Message);
        }
    }

    //public static void AddForms(object oClass, string oFormUID)
    //{
    //    //mForm.Add(oFormUID, oClass);
    //}
} 
