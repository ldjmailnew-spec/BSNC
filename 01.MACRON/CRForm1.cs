//****************************************************************************
//  File            : CRForm1.cs
//  Module          : 공통
//  Desc            : 
//  Create Date     : 
//  Creator         :
//  Memo            :
//  Copyright (c) BS&C
//****************************************************************************
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
//using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace CRPRT
{
    public partial class CRForm1 : Form
    {
        public CRForm1()
        {
            InitializeComponent();
            CrView1.Zoom(100);
        }

        //public CRForm1(object CrRep)
        //{
        //    InitializeComponent();
        //    this.CrView1.ReportSource = CrRep;
        //    this.CrView1.Zoom(100);
        //}
    }
}
