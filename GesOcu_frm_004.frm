VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frm_RepLav_01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form2"
   ClientHeight    =   2310
   ClientLeft      =   8505
   ClientTop       =   4350
   ClientWidth     =   3840
   Icon            =   "GesOcu_frm_004.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2310
   ScaleWidth      =   3840
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   2355
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   3885
      _Version        =   65536
      _ExtentX        =   6853
      _ExtentY        =   4154
      _StockProps     =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   5
         Top             =   30
         Width           =   3795
         _Version        =   65536
         _ExtentX        =   6694
         _ExtentY        =   1191
         _StockProps     =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
         Begin Threed.SSPanel SSPanel7 
            Height          =   270
            Left            =   630
            TabIndex        =   6
            Top             =   60
            Width           =   2745
            _Version        =   65536
            _ExtentX        =   4842
            _ExtentY        =   476
            _StockProps     =   15
            Caption         =   "Reporte de Lavado de Activos"
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            Font3D          =   2
            Alignment       =   1
         End
         Begin VB.Image Image2 
            Height          =   480
            Left            =   120
            Picture         =   "GesOcu_frm_004.frx":000C
            Top             =   120
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   645
         Left            =   30
         TabIndex        =   7
         Top             =   750
         Width           =   3795
         _Version        =   65536
         _ExtentX        =   6694
         _ExtentY        =   1138
         _StockProps     =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   3180
            Picture         =   "GesOcu_frm_004.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_ExpExc 
            Height          =   585
            Left            =   30
            Picture         =   "GesOcu_frm_004.frx":0758
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Exportar a Excel"
            Top             =   30
            Width           =   585
         End
         Begin Crystal.CrystalReport crp_Imprim 
            Left            =   1530
            Top             =   120
            _ExtentX        =   741
            _ExtentY        =   741
            _Version        =   348160
            WindowTitle     =   "Presentación Preliminar"
            WindowControlBox=   -1  'True
            WindowMaxButton =   -1  'True
            WindowMinButton =   -1  'True
            WindowState     =   2
            PrintFileLinesPerPage=   60
            WindowShowPrintSetupBtn=   -1  'True
            WindowShowRefreshBtn=   -1  'True
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   855
         Left            =   30
         TabIndex        =   8
         Top             =   1440
         Width           =   3795
         _Version        =   65536
         _ExtentX        =   6694
         _ExtentY        =   1508
         _StockProps     =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
         Begin EditLib.fpDateTime ipp_FecIni 
            Height          =   315
            Left            =   1380
            TabIndex        =   0
            Top             =   90
            Width           =   1425
            _Version        =   196608
            _ExtentX        =   2514
            _ExtentY        =   556
            Enabled         =   -1  'True
            MousePointer    =   0
            Object.TabStop         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            ThreeDInsideStyle=   1
            ThreeDInsideHighlightColor=   -2147483637
            ThreeDInsideShadowColor=   -2147483642
            ThreeDInsideWidth=   1
            ThreeDOutsideStyle=   1
            ThreeDOutsideHighlightColor=   -2147483628
            ThreeDOutsideShadowColor=   -2147483632
            ThreeDOutsideWidth=   1
            ThreeDFrameWidth=   0
            BorderStyle     =   0
            BorderColor     =   -2147483642
            BorderWidth     =   1
            ButtonDisable   =   0   'False
            ButtonHide      =   0   'False
            ButtonIncrement =   1
            ButtonMin       =   0
            ButtonMax       =   100
            ButtonStyle     =   3
            ButtonWidth     =   0
            ButtonWrap      =   -1  'True
            ButtonDefaultAction=   -1  'True
            ThreeDText      =   0
            ThreeDTextHighlightColor=   -2147483637
            ThreeDTextShadowColor=   -2147483632
            ThreeDTextOffset=   1
            AlignTextH      =   0
            AlignTextV      =   0
            AllowNull       =   0   'False
            NoSpecialKeys   =   0
            AutoAdvance     =   0   'False
            AutoBeep        =   0   'False
            CaretInsert     =   0
            CaretOverWrite  =   3
            UserEntry       =   0
            HideSelection   =   -1  'True
            InvalidColor    =   -2147483637
            InvalidOption   =   0
            MarginLeft      =   3
            MarginTop       =   3
            MarginRight     =   3
            MarginBottom    =   3
            NullColor       =   -2147483637
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   0   'False
            OnFocusPosition =   0
            ControlType     =   0
            Text            =   "28/09/2004"
            DateCalcMethod  =   0
            DateTimeFormat  =   0
            UserDefinedFormat=   ""
            DateMax         =   "00000000"
            DateMin         =   "00000000"
            TimeMax         =   "000000"
            TimeMin         =   "000000"
            TimeString1159  =   ""
            TimeString2359  =   ""
            DateDefault     =   "00000000"
            TimeDefault     =   "000000"
            TimeStyle       =   0
            BorderGrayAreaColor=   -2147483637
            ThreeDOnFocusInvert=   0   'False
            ThreeDFrameColor=   -2147483637
            Appearance      =   2
            BorderDropShadow=   0
            BorderDropShadowColor=   -2147483632
            BorderDropShadowWidth=   3
            PopUpType       =   0
            DateCalcY2KSplit=   60
            CaretPosition   =   0
            IncYear         =   1
            IncMonth        =   1
            IncDay          =   1
            IncHour         =   1
            IncMinute       =   1
            IncSecond       =   1
            ButtonColor     =   -2147483637
            AutoMenu        =   0   'False
            StartMonth      =   4
            ButtonAlign     =   0
            BoundDataType   =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin EditLib.fpDateTime ipp_FecFin 
            Height          =   315
            Left            =   1380
            TabIndex        =   1
            Top             =   420
            Width           =   1425
            _Version        =   196608
            _ExtentX        =   2514
            _ExtentY        =   556
            Enabled         =   -1  'True
            MousePointer    =   0
            Object.TabStop         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            ThreeDInsideStyle=   1
            ThreeDInsideHighlightColor=   -2147483637
            ThreeDInsideShadowColor=   -2147483642
            ThreeDInsideWidth=   1
            ThreeDOutsideStyle=   1
            ThreeDOutsideHighlightColor=   -2147483628
            ThreeDOutsideShadowColor=   -2147483632
            ThreeDOutsideWidth=   1
            ThreeDFrameWidth=   0
            BorderStyle     =   0
            BorderColor     =   -2147483642
            BorderWidth     =   1
            ButtonDisable   =   0   'False
            ButtonHide      =   0   'False
            ButtonIncrement =   1
            ButtonMin       =   0
            ButtonMax       =   100
            ButtonStyle     =   3
            ButtonWidth     =   0
            ButtonWrap      =   -1  'True
            ButtonDefaultAction=   -1  'True
            ThreeDText      =   0
            ThreeDTextHighlightColor=   -2147483637
            ThreeDTextShadowColor=   -2147483632
            ThreeDTextOffset=   1
            AlignTextH      =   0
            AlignTextV      =   0
            AllowNull       =   0   'False
            NoSpecialKeys   =   0
            AutoAdvance     =   0   'False
            AutoBeep        =   0   'False
            CaretInsert     =   0
            CaretOverWrite  =   3
            UserEntry       =   0
            HideSelection   =   -1  'True
            InvalidColor    =   -2147483637
            InvalidOption   =   0
            MarginLeft      =   3
            MarginTop       =   3
            MarginRight     =   3
            MarginBottom    =   3
            NullColor       =   -2147483637
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   0   'False
            OnFocusPosition =   0
            ControlType     =   0
            Text            =   "28/09/2004"
            DateCalcMethod  =   0
            DateTimeFormat  =   0
            UserDefinedFormat=   ""
            DateMax         =   "00000000"
            DateMin         =   "00000000"
            TimeMax         =   "000000"
            TimeMin         =   "000000"
            TimeString1159  =   ""
            TimeString2359  =   ""
            DateDefault     =   "00000000"
            TimeDefault     =   "000000"
            TimeStyle       =   0
            BorderGrayAreaColor=   -2147483637
            ThreeDOnFocusInvert=   0   'False
            ThreeDFrameColor=   -2147483637
            Appearance      =   2
            BorderDropShadow=   0
            BorderDropShadowColor=   -2147483632
            BorderDropShadowWidth=   3
            PopUpType       =   0
            DateCalcY2KSplit=   60
            CaretPosition   =   0
            IncYear         =   1
            IncMonth        =   1
            IncDay          =   1
            IncHour         =   1
            IncMinute       =   1
            IncSecond       =   1
            ButtonColor     =   -2147483637
            AutoMenu        =   0   'False
            StartMonth      =   4
            ButtonAlign     =   0
            BoundDataType   =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin VB.Label Label4 
            Caption         =   "Fecha Fin:"
            Height          =   225
            Left            =   60
            TabIndex        =   10
            Top             =   480
            Width           =   1035
         End
         Begin VB.Label Label1 
            Caption         =   "Fecha Inicio:"
            Height          =   255
            Left            =   60
            TabIndex        =   9
            Top             =   120
            Width           =   1065
         End
      End
   End
End
Attribute VB_Name = "frm_RepLav_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim r_str_FecIni     As String
Dim r_str_FecFin     As String

Private Sub ipp_FecIni_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_FecFin)
   End If
End Sub

Private Sub ipp_FecFin_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_ExpExc)
   End If
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Limpia
   Call gs_CentraForm(Me)
   
   Call gs_SetFocus(ipp_FecIni)
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   ipp_FecIni.Text = (date - 1)
   ipp_FecFin.Text = (date)
End Sub

Private Sub fs_Limpia()
   ipp_FecIni.Text = (date)
   ipp_FecFin.Text = (date)
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub
   
Private Sub cmd_Imprim_Click()
   'Confirmacion
   If MsgBox("¿Está seguro de imprimir el Reporte?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   'Proceso
   Screen.MousePointer = 11
      
   'Se envia la cadena de conexión
   crp_Imprim.Connect = "DSN=" & moddat_g_str_NomEsq & "; UID=" & moddat_g_str_EntDat & "; PWD=" & moddat_g_str_ClaDat
   crp_Imprim.DataFiles(0) = "CRE_HIPMAE"
   crp_Imprim.DataFiles(1) = "CLI_DATGEN"
   crp_Imprim.SelectionFormula = "{CRE_HIPMAE.HIPMAE_TDOCLI} = {CLI_DATGEN.DATGEN_TIPDOC} AND "
   crp_Imprim.SelectionFormula = crp_Imprim.SelectionFormula & "{CRE_HIPMAE.HIPMAE_NDOCLI} = {CLI_DATGEN.DATGEN_NUMDOC} AND "
   crp_Imprim.SelectionFormula = crp_Imprim.SelectionFormula & "{CRE_HIPMAE.HIPMAE_FECDES} >= " & r_str_FecIni & " AND "
   crp_Imprim.SelectionFormula = crp_Imprim.SelectionFormula & "{CRE_HIPMAE.HIPMAE_FECDES} >= " & r_str_FecIni & " AND "
   crp_Imprim.SelectionFormula = crp_Imprim.SelectionFormula & "{CRE_HIPMAE.HIPMAE_FECDES} <= " & r_str_FecFin & " "
      
   'Se hace la invocación y llamado del Reporte en la ubicación correspondiente
   crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "OFI_REPLAV_01.RPT"
      
   'Se le envia el destino a una ventana de crystal report
   crp_Imprim.Destination = crptToWindow
   crp_Imprim.Action = 1
   
   'El puntero del mouse regresa al estado normal
   Screen.MousePointer = 0
End Sub

Private Sub cmd_ExpExc_Click()
   'Confirmacion
   If MsgBox("¿Está seguro de exportar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   Call fs_GenExc
   Screen.MousePointer = 0
End Sub

Private Sub fs_GenExc()
Dim r_obj_Excel      As Excel.Application
Dim r_int_ConVer     As Integer
Dim r_str_PerMes     As String
Dim r_str_TipMon     As String
   
   g_str_Parame = "SELECT * FROM CRE_HIPMAE WHERE "
   g_str_Parame = g_str_Parame & "HIPMAE_FECDES >= " & Format(CDate(ipp_FecIni.Text), "yyyymmdd") & " AND "
   g_str_Parame = g_str_Parame & "HIPMAE_FECDES <= " & Format(CDate(ipp_FecFin.Text), "yyyymmdd") & " AND "
   g_str_Parame = g_str_Parame & "HIPMAE_SITUAC = 2 "
   g_str_Parame = g_str_Parame & "ORDER BY HIPMAE_FECDES ASC "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      MsgBox "No se encontraron Movimientos registrados.", vbInformation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
   
   With r_obj_Excel.ActiveSheet
      '.Pictures.Insert ("\\Server_micasita\COMUN\FIRMAS\Micasita_Especialistas.gif")
      '.DrawingObjects(1).Left = 5
      '.DrawingObjects(1).Top = 5
                  
      .Range(.Cells(1, 13), .Cells(2, 13)).HorizontalAlignment = xlHAlignRight
      .Cells(1, 13) = "Dpto. de Tecnología e Informática"
      .Cells(2, 13) = "Desarrollo de Sistemas"
       
      .Range(.Cells(6, 1), .Cells(6, 13)).Merge
       
      .Range(.Cells(6, 1), .Cells(6, 1)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(6, 1), .Cells(6, 1)).Font.Bold = True
      .Range(.Cells(6, 1), .Cells(6, 1)).Font.Underline = xlUnderlineStyleSingle
      .Cells(6, 1) = "Reporte de Operaciones sujetas a Revisión por Lavado de Activos "
      
      .Cells(10, 1) = "ITEM"
      .Cells(10, 2) = "NRO OPERACION"
      .Cells(10, 3) = "DOC. IDENTIDAD"
      .Cells(10, 4) = "NOMBRE DEL CLIENTE"
      .Cells(10, 5) = "F. DESEMBOLSO"
      .Cells(10, 6) = "TIPO DE MONEDA"
      .Cells(10, 7) = "V. COMPRA VENTA US$"
      .Cells(10, 8) = "CUOTA INICIAL US$"
      .Cells(10, 9) = "MTO PRESTAMO US$."
      .Cells(10, 10) = "V. COMPRA VENTA S/."
      .Cells(10, 11) = "CUOTA INICIAL S/."
      .Cells(10, 12) = "MTO PRESTAMO S/."
      .Cells(10, 13) = "% INICIAL."
            
      .Range(.Cells(10, 1), .Cells(10, 13)).Font.Bold = True
      .Range(.Cells(10, 1), .Cells(10, 13)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(10, 1), .Cells(10, 13)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(10, 1), .Cells(10, 13)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(10, 1), .Cells(10, 13)).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range(.Cells(10, 1), .Cells(10, 13)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range(.Cells(10, 1), .Cells(10, 13)).Borders(xlEdgeRight).LineStyle = xlContinuous
      .Range(.Cells(10, 1), .Cells(10, 13)).Borders(xlInsideVertical).LineStyle = xlContinuous
       
      .Columns("A").ColumnWidth = 4.57
      .Columns("B").ColumnWidth = 16
      .Columns("B").HorizontalAlignment = xlHAlignCenter
      .Columns("C").ColumnWidth = 15.43
      .Columns("C").HorizontalAlignment = xlHAlignCenter
      .Columns("D").ColumnWidth = 40.86
      .Columns("E").ColumnWidth = 17.29
      .Columns("E").HorizontalAlignment = xlHAlignCenter
      .Columns("F").ColumnWidth = 30
      .Columns("F").HorizontalAlignment = xlHAlignCenter
      .Columns("G").ColumnWidth = 21
      .Columns("G").NumberFormat = "###,###,##0.00"
      .Columns("H").ColumnWidth = 21
      .Columns("H").NumberFormat = "###,###,##0.00"
      .Columns("I").ColumnWidth = 21
      .Columns("I").NumberFormat = "###,###,##0.00"
      .Columns("J").ColumnWidth = 21
      .Columns("J").NumberFormat = "###,###,##0.00"
      .Columns("K").ColumnWidth = 21
      .Columns("K").NumberFormat = "###,###,##0.00"
      .Columns("L").ColumnWidth = 21
      .Columns("L").NumberFormat = "###,###,##0.00"
      .Columns("M").ColumnWidth = 10
      .Columns("M").NumberFormat = "###,###,##0.00"
   End With
   
   g_rst_Princi.MoveFirst
   r_int_ConVer = 11
   
   Do While Not g_rst_Princi.EOF
      
      If (g_rst_Princi!HIPMAE_APODOL > 10000) Or (Format(g_rst_Princi!HIPMAE_APODOL / g_rst_Princi!HIPMAE_CVTDOL, "###,###,##0.00") * 100 > 30) Then
         'Buscando datos
         r_str_TipMon = moddat_gf_Consulta_ParDes("204", CStr(g_rst_Princi!HIPMAE_MONEDA))
            
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 1) = r_int_ConVer - 10
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 2) = CStr(gf_Formato_NumOpe(Trim(g_rst_Princi!HIPMAE_NUMOPE)))
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 3) = CStr(g_rst_Princi!HIPMAE_TDOCLI) & "-" & Trim(g_rst_Princi!HIPMAE_NDOCLI)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 4) = moddat_gf_Buscar_NomCli(CStr(g_rst_Princi!HIPMAE_TDOCLI), Trim(g_rst_Princi!HIPMAE_NDOCLI))
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 5) = CDate(gf_FormatoFecha(CStr(g_rst_Princi!HIPMAE_FECDES)))
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 6) = r_str_TipMon
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 7) = CDbl(Format(g_rst_Princi!HIPMAE_CVTDOL, "###,###,##0.00"))
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 8) = CDbl(Format(g_rst_Princi!HIPMAE_APODOL, "###,###,##0.00"))
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 9) = CDbl(Format(g_rst_Princi!HIPMAE_CVTDOL - g_rst_Princi!HIPMAE_APODOL, "###,###,##0.00"))
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 10) = CDbl(Format(g_rst_Princi!HIPMAE_CVTSOL, "###,###,##0.00"))
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 11) = CDbl(Format(g_rst_Princi!HIPMAE_APOSOL, "###,###,##0.00"))
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 12) = CDbl(Format(g_rst_Princi!HIPMAE_CVTSOL - g_rst_Princi!HIPMAE_APOSOL, "###,###,##0.00"))
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 13) = CDbl(Format((g_rst_Princi!HIPMAE_APODOL / g_rst_Princi!HIPMAE_CVTDOL) * 100, "###,###,##0.00"))
         r_int_ConVer = r_int_ConVer + 1
      End If
      
      g_rst_Princi.MoveNext
      DoEvents
   Loop
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub
