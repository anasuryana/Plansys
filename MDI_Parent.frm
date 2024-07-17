VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "ACTSKIN4.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MDI_Parent 
   BackColor       =   &H8000000C&
   Caption         =   "PLANNING SYSTEM - INJECTION"
   ClientHeight    =   8700
   ClientLeft      =   915
   ClientTop       =   1620
   ClientWidth     =   11280
   Icon            =   "MDI_Parent.frx":0000
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   Picture         =   "MDI_Parent.frx":4C4A
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin ACTIVESKINLibCtl.Skin skinFD 
      Left            =   0
      OleObjectBlob   =   "MDI_Parent.frx":8A0F
      Top             =   480
   End
   Begin VB.PictureBox picBar 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   11280
      TabIndex        =   3
      Top             =   8100
      Width           =   11280
      Begin VB.Label lblUser 
         Caption         =   "User"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   720
         TabIndex        =   5
         Top             =   0
         Width           =   2175
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "User :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   615
      End
   End
   Begin MSComctlLib.StatusBar statusBarMdi 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   2
      Top             =   8400
      Visible         =   0   'False
      Width           =   11280
      _ExtentX        =   19897
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4304
            MinWidth        =   4304
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            Bevel           =   0
            Object.Width           =   31397
            MinWidth        =   31397
            Text            =   "PT. Banshu Plastic Indonesia"
            TextSave        =   "PT. Banshu Plastic Indonesia"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox picBoxTab 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   11280
      TabIndex        =   0
      Top             =   0
      Width           =   11280
      Begin MSComctlLib.TabStrip tabWindow 
         Height          =   375
         Left            =   -15
         TabIndex        =   1
         Top             =   0
         Width           =   15435
         _ExtentX        =   27226
         _ExtentY        =   661
         MultiRow        =   -1  'True
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   1
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               ImageVarType    =   2
            EndProperty
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
      Begin VB.Menu mnuSettingConf 
         Caption         =   "Disetujui <?>, Diperiksa <?> , Dibuat <?>"
      End
   End
   Begin VB.Menu mnuSO 
      Caption         =   "SO"
      Visible         =   0   'False
      Begin VB.Menu mnuDeliveryCust 
         Caption         =   "Delivery Customer"
      End
      Begin VB.Menu mnuOSDelivery 
         Caption         =   "O/S Delivery"
      End
   End
   Begin VB.Menu mnuForecast 
      Caption         =   "Forecast"
      Begin VB.Menu mnuFCIN 
         Caption         =   "In"
      End
      Begin VB.Menu mnuRepForecast 
         Caption         =   "Report"
         Begin VB.Menu mnuHistoryFC 
            Caption         =   "History"
         End
         Begin VB.Menu mnuActFC 
            Caption         =   "Actual Forecast"
         End
         Begin VB.Menu mnuactualpercust 
            Caption         =   "Actual Forecast /Customer"
         End
         Begin VB.Menu mnuMovAVGFC 
            Caption         =   "Moving Average Forecast"
         End
         Begin VB.Menu mnuDelvSoFC 
            Caption         =   "Delivery vs SO vs Forecast"
         End
      End
      Begin VB.Menu mnuWIP 
         Caption         =   "WIP"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuLtpp 
      Caption         =   "&LTPP"
      Begin VB.Menu mnuStockatthattime 
         Caption         =   "Stock at that Time"
      End
      Begin VB.Menu mnuSetSparepart 
         Caption         =   "Setting Sparepart"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSetSstock 
         Caption         =   "Setting Safety Stock"
      End
      Begin VB.Menu mnuGenerateLtpp 
         Caption         =   "Generate LTPP"
      End
   End
   Begin VB.Menu mnuGenerateLoadCap 
      Caption         =   "Loadcap"
      Begin VB.Menu mnuMstMachine 
         Caption         =   "Master Machine"
      End
      Begin VB.Menu mnuMasterPurging 
         Caption         =   "Master Material Purging"
      End
      Begin VB.Menu mnuMstMachineB 
         Caption         =   "Master Machine B"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuMasterProduct 
         Caption         =   "Master Product"
      End
      Begin VB.Menu mstprodlist 
         Caption         =   "Master Product List"
      End
      Begin VB.Menu mnuUnregedProd 
         Caption         =   "Unregistered Product List"
      End
      Begin VB.Menu mnuMstSubcont 
         Caption         =   "Master Subcont"
      End
      Begin VB.Menu mnuGenerateLoadC 
         Caption         =   "Generate"
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRlcmachine 
         Caption         =   "Report of LoadCap /Machine"
      End
      Begin VB.Menu mnuRneedMoldMach 
         Caption         =   "Report of Overload"
      End
      Begin VB.Menu mnuRofNeedMoMa 
         Caption         =   "Report of Need for Mold and Machine"
      End
      Begin VB.Menu mnuRofGeneratedLC 
         Caption         =   "Report of Generated LoadCap"
      End
      Begin VB.Menu mnuReportOfMenuloading 
         Caption         =   "Report of Menu Loading per Customer"
      End
   End
   Begin VB.Menu amnuMPP 
      Caption         =   "MPS"
      Begin VB.Menu mnuMinMaxRun 
         Caption         =   "Min-Max Run"
         Visible         =   0   'False
      End
      Begin VB.Menu amnuDelvPlan 
         Caption         =   "Delivery Plan"
         Visible         =   0   'False
      End
      Begin VB.Menu amnuSettingOff 
         Caption         =   "Setting Off Day"
      End
      Begin VB.Menu amnuSetOvr 
         Caption         =   "Setting Overtime"
      End
      Begin VB.Menu mnuENGtrialSCH 
         Caption         =   "ENG Trial Schedule"
      End
      Begin VB.Menu amnuGenMpp 
         Caption         =   "Generate Menu Loading"
      End
      Begin VB.Menu amnuGeneratR 
         Caption         =   "Generate"
      End
      Begin VB.Menu mnuRepWO 
         Caption         =   "Reprint WO"
      End
      Begin VB.Menu mnuInsertion 
         Caption         =   "Insertion"
      End
      Begin VB.Menu mnuCancelWO 
         Caption         =   "Cancel WO"
      End
      Begin VB.Menu pemisah 
         Caption         =   "-"
      End
      Begin VB.Menu mnuUnprocessed 
         Caption         =   "Report of Unprocessed Item"
      End
      Begin VB.Menu mnuLoadingReport 
         Caption         =   "Report of Menu Loading"
      End
      Begin VB.Menu mnuLoadingReportChart 
         Caption         =   "Report of Menu Loading (Chart)"
      End
      Begin VB.Menu mnuRoOverloadingMPS 
         Caption         =   "Report of Overloading"
      End
      Begin VB.Menu mnuPrcntMchperton 
         Caption         =   "Report of Machine's Percentage per Tonage"
      End
      Begin VB.Menu mnuDocComparis 
         Caption         =   "Document Comparison"
      End
      Begin VB.Menu mnuPlanvsWO 
         Caption         =   "WO vs Actual"
      End
   End
   Begin VB.Menu mnuMPP 
      Caption         =   "MPP"
      Visible         =   0   'False
      Begin VB.Menu mnuSettOffDay 
         Caption         =   "Setting Off Day"
      End
      Begin VB.Menu mnuDeliverySche 
         Caption         =   "Delivery Schedule"
      End
      Begin VB.Menu mnuGenerateMPP 
         Caption         =   "Generate MPP"
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "Window"
      Visible         =   0   'False
      Begin VB.Menu mnuClose 
         Caption         =   "Close"
      End
      Begin VB.Menu mnuCloseAll 
         Caption         =   "Close All"
      End
      Begin VB.Menu mnuCloseAllBT 
         Caption         =   "Close All But This"
      End
   End
   Begin VB.Menu mnuPopDet 
      Caption         =   "PopUpDetails"
      Visible         =   0   'False
      Begin VB.Menu mnuDetals 
         Caption         =   "Details"
      End
   End
   Begin VB.Menu mnuFreezColumn 
      Caption         =   "Freeze Column"
      Visible         =   0   'False
   End
   Begin VB.Menu mnuFontSize 
      Caption         =   "Font Size"
      Visible         =   0   'False
   End
End
Attribute VB_Name = "MDI_Parent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub mouse_event Lib "user32.dll" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)

Private Sub amnuDelvPlan_Click()
    F_DelvSchedule.Show
    F_DelvSchedule.SetFocus
End Sub

Private Sub amnuGeneratR_Click()
    Form_GenMPPr.Show
    Form_GenMPPr.SetFocus
End Sub

Private Sub amnuGenMpp_Click()
    Form_GenMPP.Show
    Form_GenMPP.SetFocus
End Sub

Private Sub amnuSetOvr_Click()
    Form_settingOverTime.Show
    Form_settingOverTime.SetFocus
End Sub

Private Sub amnuSettingOff_Click()
    Form_SettingOffMPP.Show
    Form_SettingOffMPP.SetFocus
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    End
End Sub

Private Sub MDIForm_Load()
On Error GoTo ErrHandler
    tabWindow.Tabs.Clear
    tabWindow.Width = Screen.Width
    Call activeTheme(skinFD, Me)
    'Call BukaKoneksi
    Set RsDB = Con.Execute("SELECT empno, full_name FROM hr_employee WHERE empno='" & Form_Login.txtUsername.Text & "'")
    If Not RsDB.EOF Then
        lblUser.Caption = RsDB!full_name
        pUserName = RsDB!full_name
        pUserId = RTrim(RsDB!empno)
    Else
        lblUser.Caption = "?"
        pUserName = "?"
        pUserId = "?"
    End If
    RsDB.Close
    Exit Sub
ErrHandler:
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbInformation
        End
    End If
End Sub

Private Sub mnuActFC_Click()
    Form_ReportFC1.Show
    Form_ReportFC1.SetFocus
End Sub

Private Sub mnuactualpercust_Click()
    Form_ReportFC4.Show
    Form_ReportFC4.SetFocus
End Sub

Private Sub mnuAge_Click()
    Form_MSTDC.Show
    Form_MSTDC.SetFocus
End Sub

Private Sub mnuCancelWO_Click()
    Form_CancelWO.Show
    Form_CancelWO.SetFocus
End Sub

Private Sub mnuClose_Click()
    If Not ActiveForm Is Nothing Then Unload ActiveForm
End Sub

Private Sub mnuCloseAll_Click()
    If tabWindow.Tabs.Count = 0 Then Exit Sub
    Do Until ActiveForm Is Nothing
        Unload ActiveForm
    Loop
End Sub

Private Sub mnuCloseAllBT_Click()
    If tabWindow.Tabs.Count <= 1 Then Exit Sub
    
    Dim i As Integer
    i = tabWindow.SelectedItem.Tag

    Dim f As Form
    For Each f In Forms
        If f.Tag <> "" Then
            If f.Tag <> i Then Unload f
        End If
    Next
End Sub

Private Sub mnuDeliveryCust_Click()
    Form_CreateSJOut.Show
    Form_CreateSJOut.SetFocus
End Sub

Private Sub mnuDeliverySche_Click()
'    Form_DelSchedule.Show 1
End Sub

Private Sub mnuDelvSoFC_Click()
    Form_ReportFC3.Show
    Form_ReportFC3.SetFocus
End Sub

Private Sub mnuDetals_Click()
    popUp_HistoryProc.Show 1 'F_Mst_Product_v2.idPROC
End Sub

Private Sub mnuDocComparis_Click()
    F_Report3var.Show
    F_Report3var.SetFocus
End Sub

Private Sub mnuENGtrialSCH_Click()
    Form_STE.Show
    Form_STE.SetFocus
End Sub

Private Sub mnuExit_Click()
    End
End Sub

Private Sub mnuFCIN_Click()
    Form_Forecast.Show
    Form_Forecast.SetFocus
End Sub

Private Sub mnuFontSize_Click()
    Dim ib As Variant
    ib = InputBox("Font Size :", "Change Font Size", 9)
    If IsNumeric(ib) Then
        Dim a As Long
        Dim k As Byte
        With Form_GenMPPr
            For a = 0 To .anaGrid.rows - 1
                .anaGrid.Row = a
                For k = 0 To .anaGrid.Cols - 1
                    .anaGrid.Col = k
                    .anaGrid.cellFontSize = CInt(ib)
                Next
            Next
        End With
    End If
End Sub

Private Sub mnuFreezColumn_Click()
    Dim ib As Variant
    ib = InputBox("Total Freezed column(s) :", "Freeze Column", 3)
    If IsNumeric(ib) Then
        If ib <= 21 Then
            If Form_GenMPPr.cmbType.Text = "Machine Inj" Then
                Form_GenMPPr.anaGrid.FixedCols = ib
            Else
                Form_GenMPPr.anaSubcont.FixedCols = ib
            End If
        End If
    End If
End Sub

Private Sub mnuGenerateLoadC_Click()
    F_GenerateLoadCap_V2.Show
    F_GenerateLoadCap_V2.SetFocus
End Sub

Private Sub mnuGenerateLtpp_Click()
    Form_GenerateLTPP.Show
    Form_GenerateLTPP.SetFocus
End Sub

Private Sub mnuGenerateMPP_Click()
    Form_GenerateMPP.Show
    Form_GenerateMPP.SetFocus
End Sub

Private Sub mnuHistoryFC_Click()
    Form_ReportFC2.Show
    Form_ReportFC2.SetFocus
End Sub

Private Sub mnuInsertion_Click()
    Form_Insertion.Show
    Form_Insertion.SetFocus
End Sub

Private Sub mnuLoadingReport_Click()
   F_ReportMnuLoading.Show
   F_ReportMnuLoading.SetFocus
End Sub

Private Sub mnuLoadingReportChart_Click()
    Form_RC_menuloading.Show
    Form_RC_menuloading.SetFocus
End Sub

Private Sub mnuMasterProduct_Click()
    F_Mst_Product_v2.Show
    F_Mst_Product_v2.SetFocus
End Sub

Private Sub mnuMasterPurging_Click()
    Form_MaterialPurging.Show
    Form_MaterialPurging.SetFocus
End Sub

Private Sub mnuMinMaxRun_Click()
    Dim crunmax As Variant, crunmin As Variant
    crunmax = InputBox("Maximum Jalan ?", "Tentukan", GetINI("MPP", "maxrun", vbNullString))
    crunmin = InputBox("Minimum Jalan ?", "Tentukan", GetINI("MPP", "minrun", vbNullString))
    If Len(crunmax) > 0 And Len(crunmin) Then
        SaveINI "MPP", "maxrun", crunmax
        SaveINI "MPP", "minrun", crunmin
    End If
End Sub

Private Sub mnuMovAVGFC_Click()
    Form_MovAVG.Show
    Form_MovAVG.SetFocus
End Sub

Private Sub mnuMstMachine_Click()
    F_Mst_Mesin.Show
    F_Mst_Mesin.SetFocus
End Sub

Private Sub mnuMstSubcont_Click()
    Form_MasterSubcont.Show
    Form_MasterSubcont.SetFocus
End Sub

Private Sub mnuPlanvsWO_Click()
    Form_PlanVsWO.Show
    Form_PlanVsWO.SetFocus
End Sub

Private Sub mnuPrcntMchperton_Click()
    Form_RekapMPPMCH.Show
    Form_RekapMPPMCH.SetFocus
End Sub

Private Sub mnuReportOfMenuloading_Click()
    Form_RLoading_c1.Show
    Form_RLoading_c1.SetFocus
End Sub

Private Sub mnuRepWO_Click()
    F_ReprintWO.Show
    F_ReprintWO.SetFocus
End Sub

Private Sub mnuRlcmachine_Click()
    F_ReportLCMach.Show
    F_ReportLCMach.SetFocus
End Sub

Private Sub mnuRneedMoldMach_Click()
    F_ReportNeedMM.Show
    F_ReportNeedMM.SetFocus
End Sub

Private Sub mnuRofGeneratedLC_Click()
    F_ReportofG.Show
    F_ReportofG.Show
End Sub

Private Sub mnuRofNeedMoMa_Click()
    F_ReportofNeedMM.Show
    F_ReportofNeedMM.SetFocus
End Sub

Private Sub mnuRoOverloadingMPS_Click()
    F_ReportNeedMMmps.Show
    F_ReportNeedMMmps.SetFocus
End Sub

Private Sub mnuSetSparepart_Click()
    Form_SetSparepart.Show
    Form_SetSparepart.SetFocus
End Sub

Private Sub mnuSetSstock_Click()
    Form_SafetyStock.Show
    Form_SafetyStock.SetFocus
End Sub

Private Sub mnuSettingConf_Click()
    Dim ib_diperika, ib_dibuat, ib_disetujui As String
    ib_disetujui = InputBox("Disetujui oleh ?", "Tentukan", GetINI("LTPP", "diketahui", vbNullString))
    ib_diperika = InputBox("Diperiksa oleh ?", "Tentukan", GetINI("LTPP", "diperiksa", vbNullString))
    ib_dibuat = InputBox("Dibuat oleh ?", "Tentukan", GetINI("LTPP", "dibuat", vbNullString))
    If Len(ib_dibuat) > 0 And Len(ib_diperika) > 0 And Len(ib_disetujui) > 0 Then
        SaveINI "LTPP", "diketahui", ib_disetujui
        SaveINI "LTPP", "diperiksa", ib_diperika
        SaveINI "LTPP", "dibuat", ib_dibuat
    End If
End Sub

Private Sub mnuSettOffDay_Click()
'    Form_SettingOffDay.Show 1
End Sub

Private Sub mnuStockatthattime_Click()
    Form_StockpCutoff.Show
    Form_StockpCutoff.SetFocus
End Sub

Private Sub mnuUnprocessed_Click()
    Form_UnprcFull.Show
    Form_UnprcFull.SetFocus
End Sub

Private Sub mnuUnregedProd_Click()
    Form_Unregprodlist.Show
    Form_Unregprodlist.SetFocus
End Sub

Private Sub mstprodlist_Click()
    Form_Search.Show
    Form_Search.SetFocus
End Sub

Private Sub tabWindow_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
On Error Resume Next
    If Button = 2 Then
        mouse_event 2, x, Y, 0, 0
        mouse_event 4, x, Y, 0, 0
    End If
End Sub

Private Sub tabWindow_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
On Error Resume Next
    Dim f As Form
    For Each f In Forms
        If f.Tag = tabWindow.SelectedItem.Tag Then
            f.SetFocus
            Exit For
        End If
    Next
    
    If Button = 2 Then PopupMenu mnuWindow
End Sub

