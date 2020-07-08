VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frm_RegLav_01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   6720
   ClientLeft      =   2325
   ClientTop       =   2550
   ClientWidth     =   13650
   Icon            =   "GesOcu_frm_006.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6720
   ScaleWidth      =   13650
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   6765
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   13695
      _Version        =   65536
      _ExtentX        =   24156
      _ExtentY        =   11933
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
         Height          =   705
         Left            =   30
         TabIndex        =   13
         Top             =   30
         Width           =   13575
         _Version        =   65536
         _ExtentX        =   23945
         _ExtentY        =   1244
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
            Height          =   585
            Left            =   630
            TabIndex        =   14
            Top             =   30
            Width           =   8325
            _Version        =   65536
            _ExtentX        =   14684
            _ExtentY        =   1032
            _StockProps     =   15
            Caption         =   "F0501-01 Registro de Operaciones de los Sujetos Obligados del Sector Financiero Bancario"
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
            Left            =   90
            Picture         =   "GesOcu_frm_006.frx":000C
            Top             =   120
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   4065
         Left            =   30
         TabIndex        =   15
         Top             =   2610
         Width           =   13575
         _Version        =   65536
         _ExtentX        =   23945
         _ExtentY        =   7170
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
         Begin MSFlexGridLib.MSFlexGrid grd_Listad 
            Height          =   3345
            Left            =   30
            TabIndex        =   11
            Top             =   660
            Width           =   13485
            _ExtentX        =   23786
            _ExtentY        =   5900
            _Version        =   393216
            Rows            =   30
            Cols            =   8
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin Threed.SSPanel SSPanel8 
            Height          =   285
            Left            =   2460
            TabIndex        =   16
            Top             =   360
            Width           =   1515
            _Version        =   65536
            _ExtentX        =   2672
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "ID Cliente"
            ForeColor       =   16777215
            BackColor       =   16384
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel SSPanel9 
            Height          =   285
            Left            =   3960
            TabIndex        =   17
            Top             =   360
            Width           =   4205
            _Version        =   65536
            _ExtentX        =   7417
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Apellidos y Nombres"
            ForeColor       =   16777215
            BackColor       =   16384
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel SSPanel10 
            Height          =   285
            Left            =   8130
            TabIndex        =   18
            Top             =   360
            Width           =   1515
            _Version        =   65536
            _ExtentX        =   2672
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Fecha Operación"
            ForeColor       =   16777215
            BackColor       =   16384
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel SSPanel11 
            Height          =   285
            Left            =   1200
            TabIndex        =   19
            Top             =   360
            Width           =   1275
            _Version        =   65536
            _ExtentX        =   2249
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Nro. Operación"
            ForeColor       =   16777215
            BackColor       =   16384
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel SSPanel3 
            Height          =   285
            Left            =   60
            TabIndex        =   26
            Top             =   360
            Width           =   1155
            _Version        =   65536
            _ExtentX        =   2037
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Nro. Interno"
            ForeColor       =   16777215
            BackColor       =   16384
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel SSPanel4 
            Height          =   285
            Left            =   11730
            TabIndex        =   27
            Top             =   360
            Width           =   1395
            _Version        =   65536
            _ExtentX        =   2461
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Mto Operación"
            ForeColor       =   16777215
            BackColor       =   16384
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel SSPanel14 
            Height          =   285
            Left            =   9630
            TabIndex        =   28
            Top             =   360
            Width           =   2115
            _Version        =   65536
            _ExtentX        =   3731
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Tipo de Moneda"
            ForeColor       =   16777215
            BackColor       =   16384
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin VB.Label Label12 
            Caption         =   "Lista de Clientes con Registro de Operaciones Sospechosas"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   60
            TabIndex        =   25
            Top             =   60
            Width           =   6615
         End
      End
      Begin Threed.SSPanel SSPanel12 
         Height          =   1095
         Left            =   30
         TabIndex        =   20
         Top             =   1470
         Width           =   13575
         _Version        =   65536
         _ExtentX        =   23945
         _ExtentY        =   1931
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
         Begin VB.ComboBox cmb_PerMes 
            Height          =   315
            Left            =   1560
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   390
            Width           =   2775
         End
         Begin VB.ComboBox cmb_Sucurs 
            Height          =   315
            Left            =   1560
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   60
            Width           =   11895
         End
         Begin EditLib.fpLongInteger ipp_PerAno 
            Height          =   315
            Left            =   1560
            TabIndex        =   2
            Top             =   720
            Width           =   825
            _Version        =   196608
            _ExtentX        =   1455
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
            ButtonStyle     =   1
            ButtonWidth     =   0
            ButtonWrap      =   -1  'True
            ButtonDefaultAction=   -1  'True
            ThreeDText      =   0
            ThreeDTextHighlightColor=   -2147483637
            ThreeDTextShadowColor=   -2147483632
            ThreeDTextOffset=   1
            AlignTextH      =   2
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
            Text            =   "0"
            MaxValue        =   "9999"
            MinValue        =   "0"
            NegFormat       =   1
            NegToggle       =   0   'False
            Separator       =   ""
            UseSeparator    =   0   'False
            IncInt          =   1
            BorderGrayAreaColor=   -2147483637
            ThreeDOnFocusInvert=   0   'False
            ThreeDFrameColor=   -2147483637
            Appearance      =   0
            BorderDropShadow=   0
            BorderDropShadowColor=   -2147483632
            BorderDropShadowWidth=   3
            ButtonColor     =   -2147483637
            AutoMenu        =   0   'False
            ButtonAlign     =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin VB.Label Label3 
            Caption         =   "Sucursal:"
            Height          =   225
            Left            =   60
            TabIndex        =   23
            Top             =   60
            Width           =   945
         End
         Begin VB.Label Label1 
            Caption         =   "Año:"
            Height          =   255
            Left            =   90
            TabIndex        =   22
            Top             =   750
            Width           =   1365
         End
         Begin VB.Label Label10 
            Caption         =   "Periodo:"
            Height          =   315
            Left            =   60
            TabIndex        =   21
            Top             =   390
            Width           =   1365
         End
      End
      Begin Threed.SSPanel SSPanel13 
         Height          =   645
         Left            =   30
         TabIndex        =   24
         Top             =   780
         Width           =   13575
         _Version        =   65536
         _ExtentX        =   23945
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
         Begin VB.CommandButton cmd_Borrar 
            Height          =   585
            Left            =   2460
            Picture         =   "GesOcu_frm_006.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Borrar Registro"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Editar 
            Height          =   585
            Left            =   1860
            Picture         =   "GesOcu_frm_006.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Editar Registro"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_ExpArc 
            Height          =   585
            Left            =   3660
            Picture         =   "GesOcu_frm_006.frx":092A
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Exportar a Archivo Texto"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_ExpExc 
            Height          =   585
            Left            =   3060
            Picture         =   "GesOcu_frm_006.frx":0C34
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Exportar a Excel"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Agrega 
            Height          =   585
            Left            =   1230
            Picture         =   "GesOcu_frm_006.frx":0F3E
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Nuevo Registro"
            Top             =   30
            Width           =   615
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   12960
            Picture         =   "GesOcu_frm_006.frx":1248
            Style           =   1  'Graphical
            TabIndex        =   10
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Limpia 
            Height          =   585
            Left            =   630
            Picture         =   "GesOcu_frm_006.frx":168A
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Limpiar Datos"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Buscar 
            Height          =   585
            Left            =   30
            Picture         =   "GesOcu_frm_006.frx":1994
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Buscar Registros"
            Top             =   30
            Width           =   585
         End
      End
   End
End
Attribute VB_Name = "frm_RegLav_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim l_arr_Sucurs()      As moddat_tpo_Genera

Private Sub cmd_Buscar_Click()
   If cmb_Sucurs.ListIndex = -1 Then
      MsgBox "Debe seleccionar la Sucursal.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_Sucurs)
      Exit Sub
   End If
   If cmb_PerMes.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Periodo.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_PerMes)
      Exit Sub
   End If
   If ipp_PerAno.Text = 0 Then
      MsgBox "Debe seleccionar un Año.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_PerAno)
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   Call fs_Buscar
   Call fs_Activa(False)
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Limpia_Click()
   Call fs_Limpia
   Call fs_Activa(True)
   Call gs_SetFocus(cmb_Sucurs)
End Sub

Private Sub cmd_Agrega_Click()
   moddat_g_int_FlgGrb = 1
   Screen.MousePointer = 11
   frm_RegLav_02.Show 1
   Call fs_Buscar
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Editar_Click()
Dim r_int_PerMes     As Integer
Dim r_int_PerAno     As Integer
      
   r_int_PerMes = CInt(cmb_PerMes.ItemData(cmb_PerMes.ListIndex))
   r_int_PerAno = CInt(ipp_PerAno.Text)
      
   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If
   
   If MsgBox("¿Está seguro que desea editar el registro?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   grd_Listad.Col = 0
   modtac_g_str_NroInt = Left(grd_Listad.Text, 4) & Mid(grd_Listad.Text, 6, 3) & Right(grd_Listad.Text, 3)
   
   grd_Listad.Col = 1
   moddat_g_str_NumOpe = Left(Trim(grd_Listad.Text), 3) & Mid(Trim(grd_Listad.Text), 5, 2) & Right(Trim(grd_Listad.Text), 5)
   
   grd_Listad.Col = 7
   modtac_g_int_Moneda = grd_Listad.Text
   
   moddat_g_int_FlgGrb = 2
   Screen.MousePointer = 11
   
   Call gs_UbiIniGrid(grd_Listad)
   Call gs_SetFocus(grd_Listad)
   
   frm_RegLav_03.Show 1
   Call fs_Buscar
   
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Borrar_Click()
   Dim r_int_PerMes     As Integer
   Dim r_int_PerAno     As Integer
   Dim r_str_NumOpe     As String
   
   r_int_PerMes = CInt(cmb_PerMes.ItemData(cmb_PerMes.ListIndex))
   r_int_PerAno = CInt(ipp_PerAno.Text)
      
   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If
   
   If MsgBox("¿Está seguro que desea borrar el registro?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   grd_Listad.Col = 1
   r_str_NumOpe = Left(Trim(grd_Listad.Text), 3) & Mid(Trim(grd_Listad.Text), 5, 2) & Right(Trim(grd_Listad.Text), 5)
   
   'Instrucción SQL
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "DELETE FROM OPE_REGLAV WHERE "
   g_str_Parame = g_str_Parame & "REGLAV_PERMES = " & r_int_PerMes & " AND "
   g_str_Parame = g_str_Parame & "REGLAV_PERANO = " & r_int_PerAno & " AND "
   g_str_Parame = g_str_Parame & "REGLAV_NUMOPE = '" & CStr(r_str_NumOpe) & "' "
      
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   Call gs_UbiIniGrid(grd_Listad)
   Call gs_SetFocus(grd_Listad)
   Call fs_Buscar
   Screen.MousePointer = 0
End Sub

Private Sub cmd_ExpExc_Click()
   If cmb_Sucurs.ListIndex = -1 Then
      MsgBox "Debe seleccionar la Sucursal.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_Sucurs)
      Exit Sub
   End If
   If cmb_PerMes.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Periodo.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_PerMes)
      Exit Sub
   End If
   If ipp_PerAno.Text = 0 Then
      MsgBox "Debe seleccionar un Año.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_PerAno)
      Exit Sub
   End If
   If grd_Listad.Rows = 0 Then
      MsgBox "No existe datos.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   'Confirmacion
   If MsgBox("¿Está seguro de exportar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   Call fs_GenExc
   Screen.MousePointer = 0
End Sub

Private Sub cmd_ExpArc_Click()
   'confirma
   If MsgBox("¿Está seguro que desea exportar el archivo?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   'Verifica que exista ruta
   If Dir$(moddat_g_str_RutLoc, vbDirectory) = "" Then
      MsgBox "Debe crear el siguente directorio " & moddat_g_str_RutLoc, vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   'Call fs_GenArcPla
   Call fs_GenArcPlaUlt
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Limpia
   Call fs_Inicia
   Call fs_Activa(True)
   Call gs_CentraForm(Me)
   
   Call gs_SetFocus(cmb_Sucurs)
   Screen.MousePointer = 0
End Sub

Private Sub fs_Limpia()
   cmb_Sucurs.ListIndex = -1
   cmb_PerMes.ListIndex = -1
   ipp_PerAno.Text = Year(date)
   Call gs_LimpiaGrid(grd_Listad)
End Sub

Private Sub fs_Inicia()
   Call moddat_gs_Carga_LisIte_Combo(cmb_PerMes, 1, "033")
   grd_Listad.Cols = 9
   grd_Listad.ColWidth(0) = 1150
   grd_Listad.ColWidth(1) = 1265
   grd_Listad.ColWidth(2) = 1490
   grd_Listad.ColWidth(3) = 4205
   grd_Listad.ColWidth(4) = 1470
   grd_Listad.ColWidth(5) = 2115
   grd_Listad.ColWidth(6) = 1395
   grd_Listad.ColWidth(7) = 0
   grd_Listad.ColWidth(8) = 0
   grd_Listad.ColAlignment(0) = flexAlignCenterCenter
   grd_Listad.ColAlignment(1) = flexAlignCenterCenter
   grd_Listad.ColAlignment(2) = flexAlignCenterCenter
   grd_Listad.ColAlignment(3) = flexAlignLeftCenter
   grd_Listad.ColAlignment(4) = flexAlignCenterCenter
   grd_Listad.ColAlignment(5) = flexAlignCenterCenter
   grd_Listad.ColAlignment(6) = flexAlignRightCenter
   grd_Listad.ColAlignment(7) = flexAlignRightCenter
   grd_Listad.ColAlignment(8) = flexAlignRightCenter
   
   moddat_g_str_Codigo = "000001"
   Call moddat_gs_Carga_SucAge(cmb_Sucurs, l_arr_Sucurs, moddat_g_str_Codigo)
End Sub

Private Sub fs_Activa(ByVal p_Activa As Integer)
   cmb_Sucurs.Enabled = p_Activa
   cmb_PerMes.Enabled = p_Activa
   cmd_Buscar.Enabled = p_Activa
   ipp_PerAno.Enabled = p_Activa
   grd_Listad.Enabled = Not p_Activa
   cmd_ExpArc.Enabled = Not p_Activa
   cmd_ExpExc.Enabled = Not p_Activa
   cmd_Agrega.Enabled = Not p_Activa
   cmd_Editar.Enabled = Not p_Activa
   cmd_Borrar.Enabled = Not p_Activa
End Sub

Private Sub cmb_Sucurs_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_Sucurs_Click
   End If
End Sub

Private Sub cmb_PerMes_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_PerMes_Click
   End If
End Sub

Private Sub ipp_PerAno_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Buscar)
   End If
End Sub

Private Sub cmb_PerMes_Click()
   Call gs_SetFocus(ipp_PerAno)
End Sub

Private Sub cmb_Sucurs_Click()
   Call gs_SetFocus(cmb_PerMes)
End Sub

Private Sub grd_Listad_SelChange()
   If grd_Listad.Rows > 2 Then
      grd_Listad.RowSel = grd_Listad.Row
   End If
End Sub

Private Sub fs_Buscar()
   moddat_g_str_CodGrp = l_arr_Sucurs(cmb_Sucurs.ListIndex + 1).Genera_Codigo
   modtac_g_int_PerMes = cmb_PerMes.ItemData(cmb_PerMes.ListIndex)
   modtac_g_int_PerAno = ipp_PerAno.Text

   Call gs_LimpiaGrid(grd_Listad)
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM OPE_REGLAV "
   g_str_Parame = g_str_Parame & " WHERE REGLAV_CODEMP = '" & Format(moddat_g_str_CodGrp, "000000") & "'  "
   g_str_Parame = g_str_Parame & "   AND REGLAV_PERMES = " & Format(cmb_PerMes.ItemData(cmb_PerMes.ListIndex), "00") & "  "
   g_str_Parame = g_str_Parame & "   AND REGLAV_PERANO = " & ipp_PerAno.Text & " "
   g_str_Parame = g_str_Parame & " ORDER BY REGLAV_PERANO, REGLAV_PERMES, REGLAV_NROINT ASC "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If

   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      MsgBox "No se han encontrado registros.", vbExclamation, modgen_g_con_PltPar
      Call gs_SetFocus(cmb_PerMes)
      Exit Sub
   End If
   
   Call fs_Activa(False)
   grd_Listad.Redraw = False
   
   g_rst_Princi.MoveFirst
   Do While Not g_rst_Princi.EOF
      grd_Listad.Rows = grd_Listad.Rows + 1
      grd_Listad.Row = grd_Listad.Rows - 1
      
      grd_Listad.Col = 0
      grd_Listad.Text = Left(g_rst_Princi!REGLAV_NROINT, 4) & "-" & Mid(g_rst_Princi!REGLAV_NROINT, 5, 3) & "-" & Right(g_rst_Princi!REGLAV_NROINT, 3)
      
      grd_Listad.Col = 1
      grd_Listad.Text = gf_Formato_NumOpe(Trim(g_rst_Princi!REGLAV_NUMOPE & ""))
      
      grd_Listad.Col = 2
      grd_Listad.Text = CStr(g_rst_Princi!REGLAV_TCLIPE) & "-" & Trim(g_rst_Princi!REGLAV_NCLIPE & "")
      
      grd_Listad.Col = 3
      grd_Listad.Text = moddat_gf_Buscar_NomCli(g_rst_Princi!REGLAV_TCLIPE, Trim(g_rst_Princi!REGLAV_NCLIPE & ""))
      
      grd_Listad.Col = 4
      grd_Listad.Text = gf_FormatoFecha(CStr(g_rst_Princi!REGLAV_FECOPE))
            
      grd_Listad.Col = 5
      If g_rst_Princi!REGLAV_TIPMON = 1 Then
         grd_Listad.Text = "SOLES"
      ElseIf g_rst_Princi!REGLAV_TIPMON = 2 Then
         grd_Listad.Text = "DOLARES AMERICANOS"
      Else
         grd_Listad.Text = "EUROS"
      End If
      
      grd_Listad.Col = 6
      grd_Listad.Text = Format(g_rst_Princi!REGLAV_IMPOPE, "###,###,##0.00")
      
      grd_Listad.Col = 7
      grd_Listad.Text = g_rst_Princi!REGLAV_TIPMON
      
      grd_Listad.Col = 8
      grd_Listad.Text = Trim(g_rst_Princi!REGLAV_DESORG)
      
      g_rst_Princi.MoveNext
   Loop
   grd_Listad.Redraw = True
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
     
   If grd_Listad.Rows > 0 Then
      cmd_Agrega.Enabled = True
      grd_Listad.Enabled = True
   End If
   
   Call gs_UbiIniGrid(grd_Listad)
   Call gs_SetFocus(grd_Listad)
End Sub

Private Sub fs_GenArcPla()
   Dim r_int_NumRes     As Integer
   Dim r_str_PerMes     As Integer
   Dim r_str_PerAno     As Integer
   Dim r_int_CodEmp     As Integer
   Dim r_int_Contad     As Integer
   Dim r_str_Cadena     As String
   Dim r_str_NomRes     As String
   Dim r_str_FecRpt     As String
   Dim r_str_TipMon     As String
   Dim r_str_Direcc     As String
   Dim r_dbl_TipCam     As Double
   
   r_str_PerMes = CInt(cmb_PerMes.ItemData(cmb_PerMes.ListIndex))
   r_str_PerAno = CInt(ipp_PerAno.Text)
   
   r_str_FecRpt = "01/" & Format(r_str_PerMes, "00") & "/" & r_str_PerAno
   r_str_NomRes = "C:\01" & Right(r_str_PerAno, 2) & Format(r_str_PerMes, "00") & Format(modsec_gf_Fin_Del_Mes(r_str_FecRpt), "dd") & ".501"
      
   g_str_Parame = "SELECT * FROM MNT_EMPGRP "
   g_str_Parame = g_str_Parame & "WHERE EMPGRP_SITUAC = 1"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      r_int_CodEmp = g_rst_Princi!EMPGRP_CODSBS
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   'Creando Archivo
   r_int_NumRes = FreeFile
   Open r_str_NomRes For Output As r_int_NumRes
      
   g_str_Parame = "SELECT * FROM OPE_REGLAV WHERE "
   g_str_Parame = g_str_Parame & "REGLAV_CODEMP = '" & Format(moddat_g_str_CodGrp, "000000") & "' AND "
   g_str_Parame = g_str_Parame & "REGLAV_PERMES = " & Format(cmb_PerMes.ItemData(cmb_PerMes.ListIndex), "00") & " AND "
   g_str_Parame = g_str_Parame & "REGLAV_PERANO = " & ipp_PerAno.Text & " "
   g_str_Parame = g_str_Parame & "ORDER BY REGLAV_PERANO, REGLAV_PERMES, REGLAV_NROINT ASC "
      
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   r_int_Contad = 1
      
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      Print #r_int_NumRes, Format(501, "0000") & Format(1, "00") & Format(r_int_CodEmp, "00000") & r_str_PerAno & Format(r_str_PerMes, "00") & Format(modsec_gf_Fin_Del_Mes(r_str_FecRpt), "dd") & Format(12, "000")
                      
      Do While Not g_rst_Princi.EOF
         If g_rst_Princi!REGLAV_TIPMON = 1 Then
            r_str_TipMon = "S"
         ElseIf g_rst_Princi!REGLAV_TIPMON = 2 Then
            r_str_TipMon = "D"
         Else
            r_str_TipMon = "E"
         End If
         
         r_str_Cadena = Format(r_int_Contad, "00000000")
         r_str_Cadena = r_str_Cadena & gs_modsec_Genera(Format(g_rst_Princi!REGLAV_CODEMP, "###########00"), 1, "0", 4)
         r_str_Cadena = r_str_Cadena & gs_modsec_Genera(Format(g_rst_Princi!REGLAV_NROINT, "###########00"), 1, "0", 20)
         r_str_Cadena = r_str_Cadena & gs_modsec_Genera(Left(g_rst_Princi!REGLAV_MODOPE, 1), 2, " ", 1)
         r_str_Cadena = r_str_Cadena & gs_modsec_Genera(Format(g_rst_Princi!REGLAV_UBIGEO, "###########00"), 1, "0", 6)
         r_str_Cadena = r_str_Cadena & gs_modsec_Genera(Format(g_rst_Princi!REGLAV_FECOPE, "###########00"), 1, "0", 8)
         r_str_Cadena = r_str_Cadena & gs_modsec_Genera(g_rst_Princi!REGLAV_HOROPE, 2, "0", 6)
         r_str_Cadena = r_str_Cadena & gs_modsec_Genera(Format(g_rst_Princi!REGLAV_TCLIFI, "###########00"), 1, "0", 1)
         r_str_Cadena = r_str_Cadena & gs_modsec_Genera(Trim(g_rst_Princi!REGLAV_NCLIFI), 2, " ", 12)
         r_str_Cadena = r_str_Cadena & gs_modsec_Genera(g_rst_Princi!REGLAV_PATFIS, 2, " ", 40)
         r_str_Cadena = r_str_Cadena & gs_modsec_Genera(g_rst_Princi!REGLAV_MATFIS, 2, " ", 40)
         r_str_Cadena = r_str_Cadena & gs_modsec_Genera(g_rst_Princi!REGLAV_NOMFIS, 2, " ", 40)
         '''''r_str_Cadena = r_str_Cadena & gs_modsec_Genera(Trim(g_rst_Princi!REGLAV_OCUFIS), 2, " ", 4)
         
         r_str_Direcc = Mid(gs_modsec_Genera(moddat_gf_Consulta_ParDes("201", CStr(g_rst_Princi!REGLAV_TIPVIA_FIS)) & " " & Trim(g_rst_Princi!REGLAV_NOMVIA_FIS) & " " & Trim(g_rst_Princi!REGLAV_NUMERO_FIS) & IIf(Len(Trim(g_rst_Princi!REGLAV_INTDPT_FIS)) > 0, " (" & Trim(g_rst_Princi!REGLAV_INTDPT_FIS) & ")", "") & IIf(Len(Trim(g_rst_Princi!REGLAV_NOMZON_FIS)) > 0, " - " & moddat_gf_Consulta_ParDes("202", CStr(g_rst_Princi!REGLAV_TIPZON_FIS)) & " " & Trim(g_rst_Princi!REGLAV_NOMZON_FIS), ""), 2, " ", 80), 1, 80)
         r_str_Cadena = r_str_Cadena & r_str_Direcc
         r_str_Cadena = r_str_Cadena & gs_modsec_Genera(Trim(g_rst_Princi!REGLAV_TELFIS), 2, " ", 10)
         r_str_Cadena = r_str_Cadena & gs_modsec_Genera(Format(g_rst_Princi!REGLAV_TCLIPE, "###########00"), 1, "0", 1)
         r_str_Cadena = r_str_Cadena & gs_modsec_Genera(Trim(g_rst_Princi!REGLAV_NCLIPE), 2, " ", 12)
         r_str_Cadena = r_str_Cadena & IIf(Len(Trim(g_rst_Princi!REGLAV_NRUCPE)) <> Null, " (" & 0 & ")", gs_modsec_Genera(" ", 2, " ", 11))
         r_str_Cadena = r_str_Cadena & gs_modsec_Genera(g_rst_Princi!REGLAV_PATPER, 2, " ", 120)
         r_str_Cadena = r_str_Cadena & gs_modsec_Genera(g_rst_Princi!REGLAV_MATPER, 2, " ", 40)
         r_str_Cadena = r_str_Cadena & gs_modsec_Genera(g_rst_Princi!REGLAV_NOMPER, 2, " ", 40)
         '''''r_str_Cadena = r_str_Cadena & gs_modsec_Genera(g_rst_Princi!REGLAV_OCUPER, 2, " ", 4)
         
         r_str_Direcc = Mid(gs_modsec_Genera(moddat_gf_Consulta_ParDes("201", CStr(g_rst_Princi!REGLAV_TIPVIA_PER)) & " " & Trim(g_rst_Princi!REGLAV_NOMVIA_PER) & " " & Trim(g_rst_Princi!REGLAV_NUMERO_PER) & IIf(Len(Trim(g_rst_Princi!REGLAV_INTDPT_PER)) > 0, " (" & Trim(g_rst_Princi!REGLAV_INTDPT_PER) & ")", "") & IIf(Len(Trim(g_rst_Princi!REGLAV_NOMZON_PER)) > 0, " - " & moddat_gf_Consulta_ParDes("202", CStr(g_rst_Princi!REGLAV_TIPZON_PER)) & " " & Trim(g_rst_Princi!REGLAV_NOMZON_PER), ""), 2, " ", 80), 1, 80)
         r_str_Cadena = r_str_Cadena & r_str_Direcc
         r_str_Cadena = r_str_Cadena & gs_modsec_Genera(Trim(g_rst_Princi!REGLAV_TELPER), 2, " ", 10)
         r_str_Cadena = r_str_Cadena & gs_modsec_Genera(Format(g_rst_Princi!REGLAV_TCLIFA, "###########00"), 1, "0", 1)
         r_str_Cadena = r_str_Cadena & gs_modsec_Genera(Trim(g_rst_Princi!REGLAV_NCLIFA), 2, " ", 12)
         r_str_Cadena = r_str_Cadena & IIf(Len(Trim(g_rst_Princi!REGLAV_NRUCFA)) <> Null, " (" & 0 & ")", gs_modsec_Genera(" ", 2, " ", 11))
         r_str_Cadena = r_str_Cadena & gs_modsec_Genera(g_rst_Princi!REGLAV_PATFAV, 2, " ", 120)
         r_str_Cadena = r_str_Cadena & gs_modsec_Genera(g_rst_Princi!REGLAV_MATFAV, 2, " ", 40)
         r_str_Cadena = r_str_Cadena & gs_modsec_Genera(g_rst_Princi!REGLAV_NOMFAV, 2, " ", 40)
         '''''r_str_Cadena = r_str_Cadena & gs_modsec_Genera(g_rst_Princi!REGLAV_OCUFAV, 2, " ", 4)
         
         r_str_Direcc = Mid(gs_modsec_Genera(moddat_gf_Consulta_ParDes("201", CStr(g_rst_Princi!REGLAV_TIPVIA_FAV)) & " " & Trim(g_rst_Princi!REGLAV_NOMVIA_FAV) & " " & Trim(g_rst_Princi!REGLAV_NUMERO_FAV) & IIf(Len(Trim(g_rst_Princi!REGLAV_INTDPT_FAV)) > 0, " (" & Trim(g_rst_Princi!REGLAV_INTDPT_FAV) & ")", "") & IIf(Len(Trim(g_rst_Princi!REGLAV_NOMZON_FAV)) > 0, " - " & moddat_gf_Consulta_ParDes("202", CStr(g_rst_Princi!REGLAV_TIPZON_FAV)) & " " & Trim(g_rst_Princi!REGLAV_NOMZON_FAV), ""), 2, " ", 80), 1, 80)
         r_str_Cadena = r_str_Cadena & r_str_Direcc
         r_str_Cadena = r_str_Cadena & gs_modsec_Genera(Trim(g_rst_Princi!REGLAV_TELFAV), 2, " ", 10)
         ' por confirmar este dato si es correcto
         'r_str_Cadena = r_str_Cadena & gs_modsec_Genera(Format(g_rst_Princi!REGLAV_TIPEFE, "###########00"), 1, "0", 1)
         r_str_Cadena = r_str_Cadena & gs_modsec_Genera(g_rst_Princi!REGLAV_TIPFON & "", 2, " ", 1)
         '*************
         r_str_Cadena = r_str_Cadena & Trim(g_rst_Princi!REGLAV_TIPOPE)
         r_str_Cadena = r_str_Cadena & IIf(Len(Trim(g_rst_Princi!REGLAV_DESORG)) <> 0, gs_modsec_Genera(Trim(g_rst_Princi!REGLAV_DESORG), 2, " ", 80), gs_modsec_Genera(" ", 2, " ", 80))
         r_str_Cadena = r_str_Cadena & gs_modsec_Genera(r_str_TipMon, 2, " ", 1)
         r_str_Cadena = r_str_Cadena & gs_modsec_Genera(Format(g_rst_Princi!REGLAV_IMPOPE, "###########0.00"), 1, "0", 18)
         r_str_Cadena = r_str_Cadena & gs_modsec_Genera(Format(g_rst_Princi!REGLAV_TIPCAM, "###########0.00"), 1, "0", 6)
         '''''r_str_Cadena = r_str_Cadena & gs_modsec_Genera(Format(g_rst_Princi!REGLAV_CTABAN, "###########00"), 2, " ", 30)
         
         Print #r_int_NumRes, r_str_Cadena
         r_int_Contad = r_int_Contad + 1
         g_rst_Princi.MoveNext
         DoEvents
      Loop
    
   End If
               
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
    
   'Cerrando Archivo Resumen
   Close #r_int_NumRes
   Screen.MousePointer = 0
   MsgBox "Archivo creado en " & r_str_NomRes & ".", vbInformation, modgen_g_str_NomPlt
End Sub

Private Sub fs_GenArcPlaUlt()
   Dim r_int_NumRes     As Integer
   Dim r_str_PerMes     As Integer
   Dim r_str_PerAno     As Integer
   Dim r_int_CodEmp     As Integer
   Dim r_int_Contad     As Integer
   Dim r_str_Cadena     As String
   Dim r_str_NomRes     As String
   Dim r_str_FecRpt     As String
   Dim r_str_TipMon     As String
   Dim r_str_NomMon     As String
   Dim r_str_Direcc     As String
   Dim r_dbl_TipCam     As Double
   
   r_str_PerMes = CInt(cmb_PerMes.ItemData(cmb_PerMes.ListIndex))
   r_str_PerAno = CInt(ipp_PerAno.Text)
   
   r_str_FecRpt = "01/" & Format(r_str_PerMes, "00") & "/" & r_str_PerAno
   r_str_NomRes = moddat_g_str_RutLoc & "\01" & Right(r_str_PerAno, 2) & Format(r_str_PerMes, "00") & Format(modsec_gf_Fin_Del_Mes(r_str_FecRpt), "dd") & ".501"
      
   g_str_Parame = "SELECT * FROM MNT_EMPGRP "
   g_str_Parame = g_str_Parame & "WHERE EMPGRP_SITUAC = 1"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      r_int_CodEmp = g_rst_Princi!EMPGRP_CODSBS
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   'Creando Archivo
   r_int_NumRes = FreeFile
   Open r_str_NomRes For Output As r_int_NumRes
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT A.*, C.DATGEN_OCUPAC, C.DATGEN_UBIGEO, D.SOLINM_UBIGEO   "
   g_str_Parame = g_str_Parame & " FROM OPE_REGLAV A "
   g_str_Parame = g_str_Parame & "INNER JOIN CRE_HIPMAE B ON B.HIPMAE_NUMOPE = A.REGLAV_NUMOPE "
   g_str_Parame = g_str_Parame & "INNER JOIN CLI_DATGEN C ON C.DATGEN_TIPDOC = B.HIPMAE_TDOCLI AND TRIM(C.DATGEN_NUMDOC) = TRIM(B.HIPMAE_NDOCLI) "
   g_str_Parame = g_str_Parame & "INNER JOIN CRE_SOLINM D ON D.SOLINM_NUMSOL = B.HIPMAE_NUMSOL "
   g_str_Parame = g_str_Parame & "WHERE REGLAV_CODEMP = '" & Format(moddat_g_str_CodGrp, "000000") & "'  "
   g_str_Parame = g_str_Parame & "  AND REGLAV_PERMES = " & Format(cmb_PerMes.ItemData(cmb_PerMes.ListIndex), "00") & " "
   g_str_Parame = g_str_Parame & "  AND REGLAV_PERANO = " & ipp_PerAno.Text & " "
   g_str_Parame = g_str_Parame & "ORDER BY REGLAV_PERANO, REGLAV_PERMES, REGLAV_NROINT ASC "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   r_int_Contad = 1
      
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      
      Print #r_int_NumRes, Format(501, "0000") & Format(1, "00") & Format(r_int_CodEmp, "00000") & r_str_PerAno & Format(r_str_PerMes, "00") & Format(modsec_gf_Fin_Del_Mes(r_str_FecRpt), "dd") & Format(12, "000")
                      
      Do While Not g_rst_Princi.EOF
      
         If g_rst_Princi!REGLAV_TIPMON = 1 Then
            r_str_TipMon = "S": r_str_NomMon = "Soles"
         ElseIf g_rst_Princi!REGLAV_TIPMON = 2 Then
            r_str_TipMon = "D": r_str_NomMon = "Dólares Americanos"
         Else
            r_str_TipMon = "E": r_str_NomMon = "Euros"
         End If
         
         '*************** DATOS DE IDENTIFICACION DEL REGISTRO DE OPERACION ***************
         
         '01. Codigo de fila
         r_str_Cadena = Format(r_int_Contad, "00000000")
         '02. Codigo empresa informante
         r_str_Cadena = r_str_Cadena & gs_modsec_Genera(Format(g_rst_Princi!REGLAV_CODEMP, "###########00"), 1, "0", 4)
         '03. Nro. registro de operacion
         r_str_Cadena = r_str_Cadena & Format(r_int_Contad, "00000000")
         '04. Nro. registro interno de operacion
         r_str_Cadena = r_str_Cadena & gs_modsec_Genera(Format(g_rst_Princi!REGLAV_NROINT, "###########00"), 1, "0", 20)
         '05. Modalidad de la operacion
         r_str_Cadena = r_str_Cadena & gs_modsec_Genera(Left(g_rst_Princi!REGLAV_MODOPE, 1), 2, " ", 1)
         '06. Codigo ubigeo de la oficina que informa
         r_str_Cadena = r_str_Cadena & "150131"
         '07. Fecha de operacion
         r_str_Cadena = r_str_Cadena & gs_modsec_Genera(Format(g_rst_Princi!REGLAV_FECOPE, "######00"), 1, "0", 8)
         '08. Hora de operacion
         r_str_Cadena = r_str_Cadena & gs_modsec_Genera(Format(g_rst_Princi!REGLAV_HOROPE, "000000"), 2, "0", 6)
         
         '*************** INFORMACION DE LA PERSONA QUE SOLICITA O FISICAMENTE REALIZA LA OPERACION ***************
         
         '09. Tipo de relacion de la persona que solicita
         r_str_Cadena = r_str_Cadena & gs_modsec_Genera("1", 2, " ", 1)
         '10. Condicion de residencia de la persona que solicita
         r_str_Cadena = r_str_Cadena & gs_modsec_Genera(" ", 2, " ", 1)
         '11. Tipo de persona que solicita
         r_str_Cadena = r_str_Cadena & gs_modsec_Genera("2", 2, " ", 1)
         '12. Tipo de documento de identidad de la persona que solicita
         r_str_Cadena = r_str_Cadena & gs_modsec_Genera(Format(g_rst_Princi!REGLAV_TCLIFI, "#00"), 1, "0", 1)
         '13. Nro. de documento de identidad de la persona que solicita
         r_str_Cadena = r_str_Cadena & gs_modsec_Genera(Trim(g_rst_Princi!REGLAV_NCLIFI), 2, " ", 12)
         '14. Nro. RUC de la persona que solicita.
         r_str_Cadena = r_str_Cadena & gs_modsec_Genera(" ", 2, " ", 11)
         '15. Ape. paterno de la persona que solicita.
         r_str_Cadena = r_str_Cadena & gs_modsec_Genera(g_rst_Princi!REGLAV_PATFIS, 2, " ", 40)
         '16. Ape. materno de la persona que solicita.
         r_str_Cadena = r_str_Cadena & gs_modsec_Genera(g_rst_Princi!REGLAV_MATFIS, 2, " ", 40)
         '17. Nombres de la persona que solicita.
         r_str_Cadena = r_str_Cadena & gs_modsec_Genera(g_rst_Princi!REGLAV_NOMFIS, 2, " ", 40)
         '18. Ocupacion de la persona que solicita.
         r_str_Cadena = r_str_Cadena & gs_modsec_Genera(g_rst_Princi!DATGEN_OCUPAC, 2, " ", 4)
         '19. Codigo CIIU de la ocupacion de la persona que solicita.
         r_str_Cadena = r_str_Cadena & gs_modsec_Genera(" ", 2, " ", 6)
         '20. Descripcion de la ocupacion de la persona que solicita.
         r_str_Cadena = r_str_Cadena & gs_modsec_Genera(" ", 2, " ", 104)
         '21. Cargo de la persona que solicita.
         r_str_Cadena = r_str_Cadena & gs_modsec_Genera("99", 2, " ", 2)
         '22. Nombre y numero de via de la direccion de la persona que solicita.
         r_str_Direcc = Mid(gs_modsec_Genera(moddat_gf_Consulta_ParDes("201", CStr(g_rst_Princi!REGLAV_TIPVIA_FIS)) & " " & Trim(g_rst_Princi!REGLAV_NOMVIA_FIS) & " " & Trim(g_rst_Princi!REGLAV_NUMERO_FIS) & IIf(Len(Trim(g_rst_Princi!REGLAV_INTDPT_FIS)) > 0, " (" & Trim(g_rst_Princi!REGLAV_INTDPT_FIS) & ")", "") & IIf(Len(Trim(g_rst_Princi!REGLAV_NOMZON_FIS)) > 0, " - " & moddat_gf_Consulta_ParDes("202", CStr(g_rst_Princi!REGLAV_TIPZON_FIS)) & " " & Trim(g_rst_Princi!REGLAV_NOMZON_FIS), ""), 2, " ", 150), 1, 150)
         r_str_Cadena = r_str_Cadena & r_str_Direcc
         '23. Dpto. correspondiente a la direccion de la persona que solitica.
         r_str_Cadena = r_str_Cadena & gs_modsec_Genera(Mid(g_rst_Princi!SOLINM_UBIGEO, 1, 2), 2, " ", 2)
         '24. Provincia. correspondiente a la direccion de la persona que solitica.
         r_str_Cadena = r_str_Cadena & gs_modsec_Genera(Mid(g_rst_Princi!SOLINM_UBIGEO, 3, 2), 2, " ", 2)
         '25. Distrito correspondiente a la direccion de la persona que solitica.
         r_str_Cadena = r_str_Cadena & gs_modsec_Genera(Mid(g_rst_Princi!SOLINM_UBIGEO, 5, 2), 2, " ", 2)
         '26. Telefono de la persona que solicita.
         r_str_Cadena = r_str_Cadena & gs_modsec_Genera(Trim(g_rst_Princi!REGLAV_TELFIS), 2, " ", 40)
         
         '*************** INFORMACION DE LA PERSONA EN CUYO NOMBRE SE REALIZA LA OPERACION ***************
         
         '27. Tipo de relacion de la persona en cuyo nombre se realiza la operacion.
         r_str_Cadena = r_str_Cadena & gs_modsec_Genera("1", 2, " ", 1)
         '28. Condicion de residencia de la persona en cuyo nombre se realiza la operacion.
         r_str_Cadena = r_str_Cadena & gs_modsec_Genera("1", 2, " ", 1)
         '29. Tipo de persona en cuyo nombre se realiza la operacion.
         r_str_Cadena = r_str_Cadena & gs_modsec_Genera("2", 2, " ", 1)
         '30. Tipo de documento de identidad de la persona en cuyo nombre se realiza la operacion.
         r_str_Cadena = r_str_Cadena & gs_modsec_Genera(Format(g_rst_Princi!REGLAV_TCLIPE, "#00"), 1, "0", 1)
         '31. Nro. de identidad de la persona en cuyo nombre se realiza la operacion.
         r_str_Cadena = r_str_Cadena & gs_modsec_Genera(Trim(g_rst_Princi!REGLAV_NCLIPE), 2, " ", 12)
         '32. Nro. de RUC de la persona en cuyo nombre se realiza la operacion.
         r_str_Cadena = r_str_Cadena & IIf(Len(Trim(g_rst_Princi!REGLAV_NRUCPE)) <> Null, " (" & 0 & ")", gs_modsec_Genera(" ", 2, " ", 11))
         '33. Ape. Paterno o razon social de la persona en cuyo nombre se realiza la operacion.
         r_str_Cadena = r_str_Cadena & gs_modsec_Genera(g_rst_Princi!REGLAV_PATPER, 2, " ", 120)
         '34. Ape. Materno de la persona en cuyo nombre se realiza la operacion.
         r_str_Cadena = r_str_Cadena & gs_modsec_Genera(g_rst_Princi!REGLAV_MATPER, 2, " ", 40)
         '35. Nombres de la persona en cuyo nombre se realiza la operacion.
         r_str_Cadena = r_str_Cadena & gs_modsec_Genera(g_rst_Princi!REGLAV_NOMPER, 2, " ", 40)
         '36. Ocupacion de la persona en cuyo nombre se realiza la operacion.
         r_str_Cadena = r_str_Cadena & gs_modsec_Genera(g_rst_Princi!DATGEN_OCUPAC, 2, " ", 4)
         '37. Codigo CIIU de la ocupacion de la persona en cuyo nombre se realiza la operacion.
         r_str_Cadena = r_str_Cadena & gs_modsec_Genera(" ", 2, " ", 6)
         '38. Descripcion de la ocupacion de la persona en cuyo nombre se realiza la operacion.
         r_str_Cadena = r_str_Cadena & gs_modsec_Genera(" ", 2, " ", 104)
         '39. Cargo de la persona en cuyo nombre se realiza la operacion.
         r_str_Cadena = r_str_Cadena & gs_modsec_Genera(" ", 2, " ", 2)
         '40. Nombre y numero de via de la direccion de la persona en cuyo nombre se realiza la operacion.
         r_str_Direcc = Mid(gs_modsec_Genera(moddat_gf_Consulta_ParDes("201", CStr(g_rst_Princi!REGLAV_TIPVIA_PER)) & " " & Trim(g_rst_Princi!REGLAV_NOMVIA_PER) & " " & Trim(g_rst_Princi!REGLAV_NUMERO_PER) & IIf(Len(Trim(g_rst_Princi!REGLAV_INTDPT_PER)) > 0, " (" & Trim(g_rst_Princi!REGLAV_INTDPT_PER) & ")", "") & IIf(Len(Trim(g_rst_Princi!REGLAV_NOMZON_PER)) > 0, " - " & moddat_gf_Consulta_ParDes("202", CStr(g_rst_Princi!REGLAV_TIPZON_PER)) & " " & Trim(g_rst_Princi!REGLAV_NOMZON_PER), ""), 2, " ", 150), 1, 150)
         r_str_Cadena = r_str_Cadena & r_str_Direcc
         '41. Dpto. correspondiente a la direccion de la persona en cuyo nombre se realiza la operacion.
         r_str_Cadena = r_str_Cadena & gs_modsec_Genera(Mid(g_rst_Princi!SOLINM_UBIGEO, 1, 2), 2, " ", 2)
         '42. Provincia. correspondiente a la direccion de la persona en cuyo nombre se realiza la operacion.
         r_str_Cadena = r_str_Cadena & gs_modsec_Genera(Mid(g_rst_Princi!SOLINM_UBIGEO, 3, 2), 2, " ", 2)
         '43. Distrito correspondiente a la direccion de la persona en cuyo nombre se realiza la operacion.
         r_str_Cadena = r_str_Cadena & gs_modsec_Genera(Mid(g_rst_Princi!SOLINM_UBIGEO, 5, 2), 2, " ", 2)
         '44. Telefono de la persona en cuyo nombre se realiza la operacion.
         r_str_Cadena = r_str_Cadena & gs_modsec_Genera(Trim(g_rst_Princi!REGLAV_TELPER), 2, " ", 40)
         
         '*************** INFORMACION DE LA PERSONA A FAVOR DE QUIEN SE REALIZA LA OPERACION ***************
         
         '45. Tipo de relacion de la persona a favor de quien se realiza la operacion.
         r_str_Cadena = r_str_Cadena & gs_modsec_Genera("1", 2, " ", 1)
         '46. Condicion de residencia de la persona a favor de quien se realiza la operacion.
         r_str_Cadena = r_str_Cadena & gs_modsec_Genera("1", 2, " ", 1)
         '47. Tipo de persona a favor de quien se realiza la operacion.
         r_str_Cadena = r_str_Cadena & gs_modsec_Genera("2", 2, " ", 1)
         '48. Tipo de documento de identidad de la persona a favor de quien se realiza la operacion.
         r_str_Cadena = r_str_Cadena & gs_modsec_Genera(Format(g_rst_Princi!REGLAV_TCLIFA, "#00"), 1, "0", 1)
         '49. Nro. de identidad de la persona a favor de quien se realiza la operacion.
         r_str_Cadena = r_str_Cadena & gs_modsec_Genera(Trim(g_rst_Princi!REGLAV_NCLIFA), 2, " ", 12)
         '50. Nro. de RUC de la persona a favor de quien se realiza la operacion.
         r_str_Cadena = r_str_Cadena & IIf(Len(Trim(g_rst_Princi!REGLAV_NRUCFA)) <> Null, " (" & 0 & ")", gs_modsec_Genera(" ", 2, " ", 11))
         '51. Ape. Paterno o razon social de la persona a favor de quien se realiza la operacion.
         r_str_Cadena = r_str_Cadena & gs_modsec_Genera(g_rst_Princi!REGLAV_PATFAV, 2, " ", 120)
         '52. Ape. Materno de la persona a favor de quien se realiza la operacion.
         r_str_Cadena = r_str_Cadena & gs_modsec_Genera(g_rst_Princi!REGLAV_MATFAV, 2, " ", 40)
         '53. Nombres de la persona a favor de quien se realiza la operacion.
         r_str_Cadena = r_str_Cadena & gs_modsec_Genera(g_rst_Princi!REGLAV_NOMFAV, 2, " ", 40)
         '54. Ocupacion de la persona a favor de quien se realiza la operacion.
         r_str_Cadena = r_str_Cadena & gs_modsec_Genera(g_rst_Princi!DATGEN_OCUPAC, 2, " ", 4)
         '55. Codigo CIIU de la ocupacion de la persona a favor de quien se realiza la operacion.
         r_str_Cadena = r_str_Cadena & gs_modsec_Genera(" ", 2, " ", 6)
         '56. Descripcion de la ocupacion de la persona a favor de quien se realiza la operacion.
         r_str_Cadena = r_str_Cadena & gs_modsec_Genera(" ", 2, " ", 104)
         '57. Cargo de la persona a favor de quien se realiza la operacion.
         r_str_Cadena = r_str_Cadena & gs_modsec_Genera(" ", 2, " ", 2)
         '58. Nombre y numero de via de la direccion de la persona a favor de quien se realiza la operacion.
         r_str_Direcc = Mid(gs_modsec_Genera(moddat_gf_Consulta_ParDes("201", CStr(g_rst_Princi!REGLAV_TIPVIA_FAV)) & " " & Trim(g_rst_Princi!REGLAV_NOMVIA_FAV) & " " & Trim(g_rst_Princi!REGLAV_NUMERO_FAV) & IIf(Len(Trim(g_rst_Princi!REGLAV_INTDPT_FAV)) > 0, " (" & Trim(g_rst_Princi!REGLAV_INTDPT_FAV) & ")", "") & IIf(Len(Trim(g_rst_Princi!REGLAV_NOMZON_FAV)) > 0, " - " & moddat_gf_Consulta_ParDes("202", CStr(g_rst_Princi!REGLAV_TIPZON_FAV)) & " " & Trim(g_rst_Princi!REGLAV_NOMZON_FAV), ""), 2, " ", 150), 1, 150)
         r_str_Cadena = r_str_Cadena & r_str_Direcc
         '59. Dpto. correspondiente a la direccion de la persona a favor de quien se realiza la operacion.
         r_str_Cadena = r_str_Cadena & gs_modsec_Genera(Mid(g_rst_Princi!SOLINM_UBIGEO, 1, 2), 2, " ", 2)
         '60. Provincia. correspondiente a la direccion de la persona a favor de quien se realiza la operacion.
         r_str_Cadena = r_str_Cadena & gs_modsec_Genera(Mid(g_rst_Princi!SOLINM_UBIGEO, 3, 2), 2, " ", 2)
         '61. Distrito correspondiente a la direccion de la persona a favor de quien se realiza la operacion.
         r_str_Cadena = r_str_Cadena & gs_modsec_Genera(Mid(g_rst_Princi!SOLINM_UBIGEO, 5, 2), 2, " ", 2)
         '62. Telefono de la persona a favor de quien se realiza la operacion.
         r_str_Cadena = r_str_Cadena & gs_modsec_Genera(Trim(g_rst_Princi!REGLAV_TELFAV), 2, " ", 40)
         
         '*************** INFORMACION DE LA OPERACION ***************
         
         '63. Tipo de fondo con que se realizo la operacion.
         If Not g_rst_Princi!REGLAV_TIPFON Then
            r_str_Cadena = r_str_Cadena & gs_modsec_Genera(CInt(g_rst_Princi!REGLAV_TIPFON), 2, " ", 1)
         Else
            r_str_Cadena = r_str_Cadena & gs_modsec_Genera(" ", 2, " ", 1)
         End If
         '64. Tipo de operacion con que se realizo.
         r_str_Cadena = r_str_Cadena & Trim(g_rst_Princi!REGLAV_TIPOPE)
         '65. Descripcion del tipo de operacion que se realizo.
         r_str_Cadena = r_str_Cadena & gs_modsec_Genera(" ", 2, " ", 40)
         '66. Origen de los fondos involucrados en la operacion.
         r_str_Cadena = r_str_Cadena & IIf(Len(Trim(g_rst_Princi!REGLAV_DESORG)) <> 0, gs_modsec_Genera(Trim(g_rst_Princi!REGLAV_DESORG), 2, " ", 80), gs_modsec_Genera(" ", 2, " ", 80))
         '67. Moneda que se realizo la operacion.
         r_str_Cadena = r_str_Cadena & gs_modsec_Genera(r_str_TipMon, 2, " ", 1)
         '68. Descripcion de la moneda que se realizo la operacion.
         r_str_Cadena = r_str_Cadena & gs_modsec_Genera(r_str_NomMon, 2, " ", 40)
         '69. Monto que se realizo la operacion.
          r_str_Cadena = r_str_Cadena & Format(g_rst_Princi!REGLAV_IMPOPE, "000000000000000.00")
         '70. Tipo de cambio que se realizo la operacion.
         r_str_Cadena = r_str_Cadena & Format(g_rst_Princi!REGLAV_TIPCAM, "00.000")
         
         '--------------------------
         'Entidad/Cuenta involucrada en la operación/producto/servicio utilizado por la persona en cuyo nombre se realiza la operación
         '71. Codigo de empresa supervisada
         r_str_Cadena = r_str_Cadena & gs_modsec_Genera("00006", 2, " ", 5)
         '72. Tipo de cuenta supervisada
         r_str_Cadena = r_str_Cadena & gs_modsec_Genera("1", 2, " ", 1)
         '73. Cod. cuenta interbancario (CCI) o Nro. de Cheque, numero de cuenta o codigo de producto involucrado en la operacion.
         r_str_Cadena = r_str_Cadena & gs_modsec_Genera("001103690200090532", 2, " ", 20)
         '74. Entidad del exterior involucrada en la operacion.
         r_str_Cadena = r_str_Cadena & gs_modsec_Genera(" ", 2, " ", 150)
         
         '--------------------------
         'Entidad/Cuenta involucrada en la operación/producto/servicio utilizado por la persona a favor de quien se realiza la operación
         '75. Codigo de empresa supervisada
         r_str_Cadena = r_str_Cadena & gs_modsec_Genera("00240", 2, " ", 5)
         '76. Tipo de cuenta supervisada
         r_str_Cadena = r_str_Cadena & gs_modsec_Genera("1", 2, " ", 1)
         '77. Cod. cuenta interbancario (CCI) o Nro. de Cheque, numero de cuenta o codigo de producto involucrado en la operacion.
         r_str_Cadena = r_str_Cadena & gs_modsec_Genera("001103690200090532", 2, " ", 20)
         '78. Entidad del exterior involucrada en la operacion.
         r_str_Cadena = r_str_Cadena & gs_modsec_Genera(" ", 2, " ", 150)
         '--------------------------
         
         '79. Alcance de la operacion.
         r_str_Cadena = r_str_Cadena & gs_modsec_Genera("1", 2, " ", 1)
         '80. Codigo del pais origen.
         r_str_Cadena = r_str_Cadena & gs_modsec_Genera(" ", 2, " ", 2)
         '81. Codigo del pais destino.
         r_str_Cadena = r_str_Cadena & gs_modsec_Genera(" ", 2, " ", 2)
         '82. Intermediario de la operacion.
         r_str_Cadena = r_str_Cadena & gs_modsec_Genera("2", 2, " ", 1)
         '83. Forma a traves del cual se realizo la operacion.
         r_str_Cadena = r_str_Cadena & gs_modsec_Genera("1", 2, " ", 1)
         '84. Descripcion de la forma a traves del cual se realizo la operacion.
         r_str_Cadena = r_str_Cadena & gs_modsec_Genera(" ", 2, " ", 40)
         
         Print #r_int_NumRes, r_str_Cadena
         r_int_Contad = r_int_Contad + 1
         g_rst_Princi.MoveNext
         DoEvents
      Loop
    
   End If
               
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
    
   'Cerrando Archivo Resumen
   Close #r_int_NumRes
   Screen.MousePointer = 0
   MsgBox "Archivo creado en " & r_str_NomRes & ".", vbInformation, modgen_g_str_NomPlt
End Sub

Private Function modtac_gf_Buscar_Ocup(ByVal p_CodOcu As String) As String
   modtac_gf_Buscar_Ocup = ""
   
   p_CodOcu = Format(p_CodOcu, "000000")
   
   g_str_Parame = "SELECT * FROM CTB_PRFMSB WHERE "
   g_str_Parame = g_str_Parame & "PRFMSB_CODPSB = '" & p_CodOcu & "' "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
      Exit Function
   End If

   If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
      modtac_gf_Buscar_Ocup = Trim(g_rst_Listas!PrfMsb_Descri)
   End If
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Function

Private Sub fs_GenExc()
Dim r_obj_Excel      As Excel.Application
Dim r_int_nrofil     As Integer
Dim r_int_Contad     As Integer
Dim r_str_TipMon     As String
Dim r_str_NomMon     As String
Dim r_str_Direcc     As String
   
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 2
   r_obj_Excel.Workbooks.Add
      
   '*********************************************************************************************************
   '*********************************************************************************************************
   r_obj_Excel.Sheets(1).Name = "RESUMEN"
   With r_obj_Excel.Sheets(1)
      .Range(.Cells(1, 9), .Cells(2, 9)).HorizontalAlignment = xlHAlignRight
      .Cells(1, 9) = "Dpto. de Tecnología e Informática"
      .Cells(2, 9) = "Desarrollo de Sistemas"
      
      .Range(.Cells(5, 1), .Cells(5, 9)).Merge
      .Range(.Cells(5, 1), .Cells(5, 1)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(5, 1), .Cells(5, 1)).Font.Bold = True
      .Range(.Cells(5, 1), .Cells(5, 1)).Font.Underline = xlUnderlineStyleSingle
      .Cells(5, 1) = "Lista de Clientes con Registro de Operaciones Sospechosas"
      
      .Cells(7, 1) = "ITEM"
      .Cells(7, 2) = "NRO. INTERNO"
      .Cells(7, 3) = "NRO. OPERACION"
      .Cells(7, 4) = "ID CLIENTE"
      .Cells(7, 5) = "APELLIDOS Y NOMBRES"
      .Cells(7, 6) = "FECHA OPERACION"
      .Cells(7, 7) = "TIPO DE MONEDA"
      .Cells(7, 8) = "MONTO OPERACION"
      .Cells(7, 9) = "ORIGEN DE LA OPERACION"
      
      .Range(.Cells(7, 1), .Cells(7, 9)).Font.Bold = True
      .Range(.Cells(7, 1), .Cells(7, 9)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(7, 1), .Cells(7, 9)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(7, 1), .Cells(7, 9)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(7, 1), .Cells(7, 9)).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range(.Cells(7, 1), .Cells(7, 9)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range(.Cells(7, 1), .Cells(7, 9)).Borders(xlEdgeRight).LineStyle = xlContinuous
      .Range(.Cells(7, 1), .Cells(7, 9)).Borders(xlInsideVertical).LineStyle = xlContinuous
       
      .Columns("A").ColumnWidth = 6
      .Columns("A").HorizontalAlignment = xlHAlignCenter
      .Columns("B").ColumnWidth = 16
      .Columns("B").HorizontalAlignment = xlHAlignCenter
      .Columns("C").ColumnWidth = 16
      .Columns("C").HorizontalAlignment = xlHAlignCenter
      .Columns("D").ColumnWidth = 14
      .Columns("D").HorizontalAlignment = xlHAlignCenter
      .Columns("E").ColumnWidth = 40
      .Columns("F").ColumnWidth = 18
      .Columns("F").HorizontalAlignment = xlHAlignCenter
      .Columns("G").ColumnWidth = 22
      .Columns("G").HorizontalAlignment = xlHAlignCenter
      .Columns("H").ColumnWidth = 20
      .Columns("H").NumberFormat = "###,###,##0.00"
      .Columns("I").ColumnWidth = 50
      .Columns("I").HorizontalAlignment = xlHAlignCenter
      
      For r_int_nrofil = 0 To grd_Listad.Rows - 1
         r_obj_Excel.ActiveSheet.Cells(r_int_nrofil + 8, 1) = Format(r_int_nrofil + 1, "000")
         r_obj_Excel.ActiveSheet.Cells(r_int_nrofil + 8, 2) = grd_Listad.TextMatrix(r_int_nrofil, 0)
         r_obj_Excel.ActiveSheet.Cells(r_int_nrofil + 8, 3) = grd_Listad.TextMatrix(r_int_nrofil, 1)
         r_obj_Excel.ActiveSheet.Cells(r_int_nrofil + 8, 4) = grd_Listad.TextMatrix(r_int_nrofil, 2)
         r_obj_Excel.ActiveSheet.Cells(r_int_nrofil + 8, 5) = grd_Listad.TextMatrix(r_int_nrofil, 3)
         r_obj_Excel.ActiveSheet.Cells(r_int_nrofil + 8, 6) = "'" & grd_Listad.TextMatrix(r_int_nrofil, 4)
         r_obj_Excel.ActiveSheet.Cells(r_int_nrofil + 8, 7) = grd_Listad.TextMatrix(r_int_nrofil, 5)
         r_obj_Excel.ActiveSheet.Cells(r_int_nrofil + 8, 8) = grd_Listad.TextMatrix(r_int_nrofil, 6)
         r_obj_Excel.ActiveSheet.Cells(r_int_nrofil + 8, 9) = grd_Listad.TextMatrix(r_int_nrofil, 8)
      Next r_int_nrofil
   End With
   
   '*********************************************************************************************************
   '*********************************************************************************************************
   r_obj_Excel.Sheets(2).Name = "DATA - SUCAVE"
   With r_obj_Excel.Sheets(2)
      .Range(.Cells(1, 84), .Cells(2, 84)).HorizontalAlignment = xlHAlignRight
      .Cells(1, 84) = "Dpto. de Tecnología e Informática"
      .Cells(2, 84) = "Desarrollo de Sistemas"
      
      .Range(.Cells(5, 1), .Cells(5, 84)).Merge
      .Range(.Cells(5, 1), .Cells(5, 1)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(5, 1), .Cells(5, 1)).Font.Bold = True
      .Range(.Cells(5, 1), .Cells(5, 1)).Font.Underline = xlUnderlineStyleSingle
      .Cells(5, 1) = "Lista de Clientes con Registro de Operaciones Sospechosas"
      
      'INFORMACION GENERAL
      .Cells(7, 1) = "01. Codigo de fila"
      .Cells(7, 2) = "02. Codigo empresa informante"
      .Cells(7, 3) = "03. Nro. registro de operacion"
      .Cells(7, 4) = "04. Nro. registro interno de operacion"
      .Cells(7, 5) = "05. Modalidad de la operacion"
      .Cells(7, 6) = "06. Codigo ubigeo de la oficina que informa"
      .Cells(7, 7) = "07. Fecha de operacion"
      .Cells(7, 8) = "08. Hora de operacion"
         
      'INFORMACION DE LA PERSONA QUE SOLICITA O FISICAMENTE REALIZA LA OPERACION
      .Cells(7, 9) = "09. Tipo de relacion de la persona que solicita"
      .Cells(7, 10) = "10. Condicion de residencia de la persona que solicita"
      .Cells(7, 11) = "11. Tipo de persona que solicita"
      .Cells(7, 12) = "12. Tipo de documento de identidad de la persona que solicita"
      .Cells(7, 13) = "13. Nro. de documento de identidad de la persona que solicita"
      .Cells(7, 14) = "14. Nro. RUC de la persona que solicita."
      .Cells(7, 15) = "15. Ape. paterno de la persona que solicita."
      .Cells(7, 16) = "16. Ape. materno de la persona que solicita."
      .Cells(7, 17) = "17. Nombres de la persona que solicita."
      .Cells(7, 18) = "18. Ocupacion de la persona que solicita."
      .Cells(7, 19) = "19. Codigo CIIU de la ocupacion de la persona que solicita."
      .Cells(7, 20) = "20. Descripcion de la ocupacion de la persona que solicita."
      .Cells(7, 21) = "21. Cargo de la persona que solicita."
      .Cells(7, 22) = "22. Nombre y numero de via de la direccion de la persona que solicita."
      .Cells(7, 23) = "23. Dpto. correspondiente a la direccion de la persona que solitica."
      .Cells(7, 24) = "24. Provincia. correspondiente a la direccion de la persona que solitica."
      .Cells(7, 25) = "25. Distrito correspondiente a la direccion de la persona que solitica."
      .Cells(7, 26) = "26. Telefono de la persona que solicita."
         
      'INFORMACION DE LA PERSONA EN CUYO NOMBRE SE REALIZA LA OPERACION
      .Cells(7, 27) = "27. Tipo de relacion de la persona en cuyo nombre se realiza la operacion."
      .Cells(7, 28) = "28. Condicion de residencia de la persona en cuyo nombre se realiza la operacion."
      .Cells(7, 29) = "29. Tipo de persona en cuyo nombre se realiza la operacion."
      .Cells(7, 30) = "30. Tipo de documento de identidad de la persona en cuyo nombre se realiza la operacion."
      .Cells(7, 31) = "31. Nro. de identidad de la persona en cuyo nombre se realiza la operacion."
      .Cells(7, 32) = "32. Nro. de RUC de la persona en cuyo nombre se realiza la operacion."
      .Cells(7, 33) = "33. Ape. Paterno o razon social de la persona en cuyo nombre se realiza la operacion."
      .Cells(7, 34) = "34. Ape. Materno de la persona en cuyo nombre se realiza la operacion."
      .Cells(7, 35) = "35. Nombres de la persona en cuyo nombre se realiza la operacion."
      .Cells(7, 36) = "36. Ocupacion de la persona en cuyo nombre se realiza la operacion."
      .Cells(7, 37) = "37. Codigo CIIU de la ocupacion de la persona en cuyo nombre se realiza la operacion."
      .Cells(7, 38) = "38. Descripcion de la ocupacion de la persona en cuyo nombre se realiza la operacion."
      .Cells(7, 39) = "39. Cargo de la persona en cuyo nombre se realiza la operacion."
      .Cells(7, 40) = "40. Nombre y numero de via de la direccion de la persona en cuyo nombre se realiza la operacion."
      .Cells(7, 41) = "41. Dpto. correspondiente a la direccion de la persona en cuyo nombre se realiza la operacion."
      .Cells(7, 42) = "42. Provincia. correspondiente a la direccion de la persona en cuyo nombre se realiza la operacion."
      .Cells(7, 43) = "43. Distrito correspondiente a la direccion de la persona en cuyo nombre se realiza la operacion."
      .Cells(7, 44) = "44. Telefono de la persona en cuyo nombre se realiza la operacion."
      
      '*************** INFORMACION DE LA PERSONA A FAVOR DE QUIEN SE REALIZA LA OPERACION ***************
      .Cells(7, 45) = "45. Tipo de relacion de la persona a favor de quien se realiza la operacion."
      .Cells(7, 46) = "46. Condicion de residencia de la persona a favor de quien se realiza la operacion."
      .Cells(7, 47) = "47. Tipo de persona a favor de quien se realiza la operacion."
      .Cells(7, 48) = "48. Tipo de documento de identidad de la persona a favor de quien se realiza la operacion."
      .Cells(7, 49) = "49. Nro. de identidad de la persona a favor de quien se realiza la operacion."
      .Cells(7, 50) = "50. Nro. de RUC de la persona a favor de quien se realiza la operacion."
      .Cells(7, 51) = "51. Ape. Paterno o razon social de la persona a favor de quien se realiza la operacion."
      .Cells(7, 52) = "52. Ape. Materno de la persona a favor de quien se realiza la operacion."
      .Cells(7, 53) = "53. Nombres de la persona a favor de quien se realiza la operacion."
      .Cells(7, 54) = "54. Ocupacion de la persona a favor de quien se realiza la operacion."
      .Cells(7, 55) = "55. Codigo CIIU de la ocupacion de la persona a favor de quien se realiza la operacion."
      .Cells(7, 56) = "56. Descripcion de la ocupacion de la persona a favor de quien se realiza la operacion."
      .Cells(7, 57) = "57. Cargo de la persona a favor de quien se realiza la operacion."
      .Cells(7, 58) = "58. Nombre y numero de via de la direccion de la persona a favor de quien se realiza la operacion."
      .Cells(7, 59) = "59. Dpto. correspondiente a la direccion de la persona a favor de quien se realiza la operacion."
      .Cells(7, 60) = "60. Provincia. correspondiente a la direccion de la persona a favor de quien se realiza la operacion."
      .Cells(7, 61) = "61. Distrito correspondiente a la direccion de la persona a favor de quien se realiza la operacion."
      .Cells(7, 62) = "62. Telefono de la persona a favor de quien se realiza la operacion."
      
      '*************** INFORMACION DE LA OPERACION ***************
      .Cells(7, 63) = "63. Tipo de fondo con que se realizo la operacion."
      .Cells(7, 64) = "64. Tipo de operacion con que se realizo."
      .Cells(7, 65) = "65. Descripcion del tipo de operacion que se realizo."
      .Cells(7, 66) = "66. Origen de los fondos involucrados en la operacion."
      .Cells(7, 67) = "67. Moneda que se realizo la operacion."
      .Cells(7, 68) = "68. Descripcion de la moneda que se realizo la operacion."
      .Cells(7, 69) = "69. Monto que se realizo la operacion."
      .Cells(7, 70) = "70. Tipo de cambio que se realizo la operacion."
      
      '--------------------------
      'Entidad/Cuenta involucrada en la operación/producto/servicio utilizado por la persona en cuyo nombre se realiza la operación
      .Cells(7, 71) = "71. Codigo de empresa supervisada"
      .Cells(7, 72) = "72. Tipo de cuenta supervisada"
      .Cells(7, 73) = "73. Cod. cuenta interbancario (CCI) o Nro. de Cheque, numero de cuenta o codigo de producto involucrado en la operacion."
      .Cells(7, 74) = "74. Entidad del exterior involucrada en la operacion."
      
      'Entidad/Cuenta involucrada en la operación/producto/servicio utilizado por la persona a favor de quien se realiza la operación
      .Cells(7, 75) = "75. Codigo de empresa supervisada"
      .Cells(7, 76) = "76. Tipo de cuenta supervisada"
      .Cells(7, 77) = "77. Cod. cuenta interbancario (CCI) o Nro. de Cheque, numero de cuenta o codigo de producto involucrado en la operacion."
      .Cells(7, 78) = "78. Entidad del exterior involucrada en la operacion."
      .Cells(7, 79) = "79. Alcance de la operacion."
      .Cells(7, 80) = "80. Codigo del pais origen."
      .Cells(7, 81) = "81. Codigo del pais destino."
      .Cells(7, 82) = "82. Intermediario de la operacion."
      .Cells(7, 83) = "83. Forma a traves del cual se realizo la operacion."
      .Cells(7, 84) = "84. Descripcion de la forma a traves del cual se realizo la operacion."
      
      .Range(.Cells(7, 1), .Cells(7, 84)).Font.Bold = True
      .Range(.Cells(7, 1), .Cells(7, 84)).RowHeight = 60
      .Range(.Cells(7, 1), .Cells(7, 84)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(7, 1), .Cells(7, 84)).VerticalAlignment = xlHAlignCenter
      .Range(.Cells(7, 1), .Cells(7, 84)).WrapText = True
      .Range(.Cells(7, 1), .Cells(7, 84)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(7, 1), .Cells(7, 84)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(7, 1), .Cells(7, 84)).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range(.Cells(7, 1), .Cells(7, 84)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range(.Cells(7, 1), .Cells(7, 84)).Borders(xlEdgeRight).LineStyle = xlContinuous
      .Range(.Cells(7, 1), .Cells(7, 84)).Borders(xlInsideVertical).LineStyle = xlContinuous
      
      .Columns("A").ColumnWidth = 18
      .Columns("B").ColumnWidth = 18
      .Columns("C").ColumnWidth = 18
      .Columns("D").ColumnWidth = 18
      .Columns("E").ColumnWidth = 18
      .Columns("F").ColumnWidth = 18
      .Columns("G").ColumnWidth = 18
      .Columns("H").ColumnWidth = 18
      .Columns("I").ColumnWidth = 18
      .Columns("J").ColumnWidth = 18
      .Columns("K").ColumnWidth = 18
      .Columns("L").ColumnWidth = 18
      .Columns("M").ColumnWidth = 18
      .Columns("N").ColumnWidth = 18
      .Columns("O").ColumnWidth = 18
      .Columns("P").ColumnWidth = 18
      .Columns("Q").ColumnWidth = 18
      .Columns("R").ColumnWidth = 18
      .Columns("S").ColumnWidth = 18
      .Columns("T").ColumnWidth = 18
      .Columns("U").ColumnWidth = 18
      .Columns("V").ColumnWidth = 18
      .Columns("W").ColumnWidth = 18
      .Columns("X").ColumnWidth = 18
      .Columns("Y").ColumnWidth = 18
      .Columns("Z").ColumnWidth = 18
      .Columns("AA").ColumnWidth = 18
      .Columns("AB").ColumnWidth = 18
      .Columns("AC").ColumnWidth = 18
      .Columns("AD").ColumnWidth = 18
      .Columns("AE").ColumnWidth = 18
      .Columns("AF").ColumnWidth = 18
      .Columns("AG").ColumnWidth = 18
      .Columns("AH").ColumnWidth = 18
      .Columns("AI").ColumnWidth = 18
      .Columns("AJ").ColumnWidth = 18
      .Columns("AK").ColumnWidth = 18
      .Columns("AL").ColumnWidth = 18
      .Columns("AM").ColumnWidth = 18
      .Columns("AN").ColumnWidth = 18
      .Columns("AO").ColumnWidth = 18
      .Columns("AP").ColumnWidth = 18
      .Columns("AQ").ColumnWidth = 18
      .Columns("AR").ColumnWidth = 18
      .Columns("AS").ColumnWidth = 18
      .Columns("AT").ColumnWidth = 18
      .Columns("AU").ColumnWidth = 18
      .Columns("AV").ColumnWidth = 18
      .Columns("AW").ColumnWidth = 18
      .Columns("AX").ColumnWidth = 18
      .Columns("AY").ColumnWidth = 18
      .Columns("AZ").ColumnWidth = 18
      .Columns("BA").ColumnWidth = 18
      .Columns("BB").ColumnWidth = 18
      .Columns("BC").ColumnWidth = 18
      .Columns("BD").ColumnWidth = 18
      .Columns("BE").ColumnWidth = 18
      .Columns("BF").ColumnWidth = 18
      .Columns("BG").ColumnWidth = 18
      .Columns("BH").ColumnWidth = 18
      .Columns("BI").ColumnWidth = 18
      .Columns("BJ").ColumnWidth = 18
      .Columns("BK").ColumnWidth = 18
      .Columns("BL").ColumnWidth = 18
      .Columns("BM").ColumnWidth = 18
      .Columns("BN").ColumnWidth = 18
      .Columns("BO").ColumnWidth = 18
      .Columns("BP").ColumnWidth = 18
      .Columns("BQ").ColumnWidth = 18
      .Columns("BR").ColumnWidth = 18
      .Columns("BS").ColumnWidth = 18
      .Columns("BT").ColumnWidth = 18
      .Columns("BU").ColumnWidth = 18
      .Columns("BV").ColumnWidth = 18
      .Columns("BW").ColumnWidth = 18
      .Columns("BX").ColumnWidth = 18
      .Columns("BY").ColumnWidth = 18
      .Columns("BZ").ColumnWidth = 18
      .Columns("CA").ColumnWidth = 18
      .Columns("CB").ColumnWidth = 18
      .Columns("CC").ColumnWidth = 18
      .Columns("CD").ColumnWidth = 18
      .Columns("CE").ColumnWidth = 18
      .Columns("CF").ColumnWidth = 18
      
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "SELECT A.*, C.DATGEN_OCUPAC, C.DATGEN_UBIGEO, D.SOLINM_UBIGEO   "
      g_str_Parame = g_str_Parame & " FROM OPE_REGLAV A "
      g_str_Parame = g_str_Parame & "INNER JOIN CRE_HIPMAE B ON B.HIPMAE_NUMOPE = A.REGLAV_NUMOPE "
      g_str_Parame = g_str_Parame & "INNER JOIN CLI_DATGEN C ON C.DATGEN_TIPDOC = B.HIPMAE_TDOCLI AND TRIM(C.DATGEN_NUMDOC) = TRIM(B.HIPMAE_NDOCLI) "
      g_str_Parame = g_str_Parame & "INNER JOIN CRE_SOLINM D ON D.SOLINM_NUMSOL = B.HIPMAE_NUMSOL "
      g_str_Parame = g_str_Parame & "WHERE REGLAV_CODEMP = '" & Format(moddat_g_str_CodGrp, "000000") & "'  "
      g_str_Parame = g_str_Parame & "  AND REGLAV_PERMES = " & Format(cmb_PerMes.ItemData(cmb_PerMes.ListIndex), "00") & " "
      g_str_Parame = g_str_Parame & "  AND REGLAV_PERANO = " & ipp_PerAno.Text & " "
      g_str_Parame = g_str_Parame & "ORDER BY REGLAV_PERANO, REGLAV_PERMES, REGLAV_NROINT ASC "
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         Exit Sub
      End If
      
      r_int_Contad = 1
      If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
         g_rst_Princi.MoveFirst
         Do While Not g_rst_Princi.EOF
         
            If g_rst_Princi!REGLAV_TIPMON = 1 Then
               r_str_TipMon = "S": r_str_NomMon = "Soles"
            ElseIf g_rst_Princi!REGLAV_TIPMON = 2 Then
               r_str_TipMon = "D": r_str_NomMon = "Dólares Americanos"
            Else
               r_str_TipMon = "E": r_str_NomMon = "Euros"
            End If
            
            '*************** DATOS DE IDENTIFICACION DEL REGISTRO DE OPERACION ***************
            '01. Codigo de fila
            .Cells(r_int_Contad + 7, 1) = Format(r_int_Contad, "00000000")
            '02. Codigo empresa informante
            .Cells(r_int_Contad + 7, 2) = gs_modsec_Genera(Format(g_rst_Princi!REGLAV_CODEMP, "###########00"), 1, "0", 4)
            '03. Nro. registro de operacion
            .Cells(r_int_Contad + 7, 3) = Format(r_int_Contad, "00000000")
            '04. Nro. registro interno de operacion
            .Cells(r_int_Contad + 7, 4) = gs_modsec_Genera(Format(g_rst_Princi!REGLAV_NROINT, "###########00"), 1, "0", 20)
            '05. Modalidad de la operacion
            .Cells(r_int_Contad + 7, 5) = gs_modsec_Genera(Left(g_rst_Princi!REGLAV_MODOPE, 1), 2, " ", 1)
            '06. Codigo ubigeo de la oficina que informa
            .Cells(r_int_Contad + 7, 6) = "150131"
            '07. Fecha de operacion
            .Cells(r_int_Contad + 7, 7) = gs_modsec_Genera(Format(g_rst_Princi!REGLAV_FECOPE, "######00"), 1, "0", 8)
            '08. Hora de operacion
            .Cells(r_int_Contad + 7, 8) = gs_modsec_Genera(Format(g_rst_Princi!REGLAV_HOROPE, "000000"), 2, "0", 6)
            
            '*************** INFORMACION DE LA PERSONA QUE SOLICITA O FISICAMENTE REALIZA LA OPERACION ***************
            '09. Tipo de relacion de la persona que solicita
            .Cells(r_int_Contad + 7, 9) = gs_modsec_Genera("1", 2, " ", 1)
            '10. Condicion de residencia de la persona que solicita
            .Cells(r_int_Contad + 7, 10) = gs_modsec_Genera(" ", 2, " ", 1)
            '11. Tipo de persona que solicita
            .Cells(r_int_Contad + 7, 11) = gs_modsec_Genera("2", 2, " ", 1)
            '12. Tipo de documento de identidad de la persona que solicita
            .Cells(r_int_Contad + 7, 12) = gs_modsec_Genera(Format(g_rst_Princi!REGLAV_TCLIFI, "#00"), 1, "0", 1)
            '13. Nro. de documento de identidad de la persona que solicita
            .Cells(r_int_Contad + 7, 13) = gs_modsec_Genera(Trim(g_rst_Princi!REGLAV_NCLIFI), 2, " ", 12)
            '14. Nro. RUC de la persona que solicita.
            .Cells(r_int_Contad + 7, 14) = gs_modsec_Genera(" ", 2, " ", 11)
            '15. Ape. paterno de la persona que solicita.
            .Cells(r_int_Contad + 7, 15) = gs_modsec_Genera(g_rst_Princi!REGLAV_PATFIS, 2, " ", 40)
            '16. Ape. materno de la persona que solicita.
            .Cells(r_int_Contad + 7, 16) = gs_modsec_Genera(g_rst_Princi!REGLAV_MATFIS, 2, " ", 40)
            '17. Nombres de la persona que solicita.
            .Cells(r_int_Contad + 7, 17) = gs_modsec_Genera(g_rst_Princi!REGLAV_NOMFIS, 2, " ", 40)
            '18. Ocupacion de la persona que solicita.
            .Cells(r_int_Contad + 7, 18) = gs_modsec_Genera(g_rst_Princi!DATGEN_OCUPAC, 2, " ", 4)
            '19. Codigo CIIU de la ocupacion de la persona que solicita.
            .Cells(r_int_Contad + 7, 19) = gs_modsec_Genera(" ", 2, " ", 6)
            '20. Descripcion de la ocupacion de la persona que solicita.
            .Cells(r_int_Contad + 7, 20) = gs_modsec_Genera(" ", 2, " ", 104)
            '21. Cargo de la persona que solicita.
            .Cells(r_int_Contad + 7, 21) = gs_modsec_Genera("99", 2, " ", 2)
            '22. Nombre y numero de via de la direccion de la persona que solicita.
            r_str_Direcc = Mid(gs_modsec_Genera(moddat_gf_Consulta_ParDes("201", CStr(g_rst_Princi!REGLAV_TIPVIA_FIS)) & " " & Trim(g_rst_Princi!REGLAV_NOMVIA_FIS) & " " & Trim(g_rst_Princi!REGLAV_NUMERO_FIS) & IIf(Len(Trim(g_rst_Princi!REGLAV_INTDPT_FIS)) > 0, " (" & Trim(g_rst_Princi!REGLAV_INTDPT_FIS) & ")", "") & IIf(Len(Trim(g_rst_Princi!REGLAV_NOMZON_FIS)) > 0, " - " & moddat_gf_Consulta_ParDes("202", CStr(g_rst_Princi!REGLAV_TIPZON_FIS)) & " " & Trim(g_rst_Princi!REGLAV_NOMZON_FIS), ""), 2, " ", 150), 1, 150)
            .Cells(r_int_Contad + 7, 22) = r_str_Direcc
            '23. Dpto. correspondiente a la direccion de la persona que solitica.
            .Cells(r_int_Contad + 7, 23) = gs_modsec_Genera(Mid(g_rst_Princi!SOLINM_UBIGEO, 1, 2), 2, " ", 2)
            '24. Provincia. correspondiente a la direccion de la persona que solitica.
            .Cells(r_int_Contad + 7, 24) = gs_modsec_Genera(Mid(g_rst_Princi!SOLINM_UBIGEO, 3, 2), 2, " ", 2)
            '25. Distrito correspondiente a la direccion de la persona que solitica.
            .Cells(r_int_Contad + 7, 25) = gs_modsec_Genera(Mid(g_rst_Princi!SOLINM_UBIGEO, 5, 2), 2, " ", 2)
            '26. Telefono de la persona que solicita.
            .Cells(r_int_Contad + 7, 26) = gs_modsec_Genera(Trim(g_rst_Princi!REGLAV_TELFIS), 2, " ", 40)
            
            '*************** INFORMACION DE LA PERSONA EN CUYO NOMBRE SE REALIZA LA OPERACION ***************
            '27. Tipo de relacion de la persona en cuyo nombre se realiza la operacion.
            .Cells(r_int_Contad + 7, 27) = gs_modsec_Genera("1", 2, " ", 1)
            '28. Condicion de residencia de la persona en cuyo nombre se realiza la operacion.
            .Cells(r_int_Contad + 7, 28) = gs_modsec_Genera(" ", 2, " ", 1)
            '29. Tipo de persona en cuyo nombre se realiza la operacion.
            .Cells(r_int_Contad + 7, 29) = gs_modsec_Genera("2", 2, " ", 1)
            '30. Tipo de documento de identidad de la persona en cuyo nombre se realiza la operacion.
            .Cells(r_int_Contad + 7, 30) = gs_modsec_Genera(Format(g_rst_Princi!REGLAV_TCLIPE, "#00"), 1, "0", 1)
            '31. Nro. de identidad de la persona en cuyo nombre se realiza la operacion.
            .Cells(r_int_Contad + 7, 31) = gs_modsec_Genera(Trim(g_rst_Princi!REGLAV_NCLIPE), 2, " ", 12)
            '32. Nro. de RUC de la persona en cuyo nombre se realiza la operacion.
            .Cells(r_int_Contad + 7, 32) = IIf(Len(Trim(g_rst_Princi!REGLAV_NRUCPE)) <> Null, " (" & 0 & ")", gs_modsec_Genera(" ", 2, " ", 11))
            '33. Ape. Paterno o razon social de la persona en cuyo nombre se realiza la operacion.
            .Cells(r_int_Contad + 7, 33) = gs_modsec_Genera(g_rst_Princi!REGLAV_PATPER, 2, " ", 120)
            '34. Ape. Materno de la persona en cuyo nombre se realiza la operacion.
            .Cells(r_int_Contad + 7, 34) = gs_modsec_Genera(g_rst_Princi!REGLAV_MATPER, 2, " ", 40)
            '35. Nombres de la persona en cuyo nombre se realiza la operacion.
            .Cells(r_int_Contad + 7, 35) = gs_modsec_Genera(g_rst_Princi!REGLAV_NOMPER, 2, " ", 40)
            '36. Ocupacion de la persona en cuyo nombre se realiza la operacion.
            .Cells(r_int_Contad + 7, 36) = gs_modsec_Genera(g_rst_Princi!DATGEN_OCUPAC, 2, " ", 4)
            '37. Codigo CIIU de la ocupacion de la persona en cuyo nombre se realiza la operacion.
            .Cells(r_int_Contad + 7, 37) = gs_modsec_Genera(" ", 2, " ", 6)
            '38. Descripcion de la ocupacion de la persona en cuyo nombre se realiza la operacion.
            .Cells(r_int_Contad + 7, 38) = gs_modsec_Genera(" ", 2, " ", 104)
            '39. Cargo de la persona en cuyo nombre se realiza la operacion.
            .Cells(r_int_Contad + 7, 39) = gs_modsec_Genera(" ", 2, " ", 2)
            '40. Nombre y numero de via de la direccion de la persona en cuyo nombre se realiza la operacion.
            r_str_Direcc = Mid(gs_modsec_Genera(moddat_gf_Consulta_ParDes("201", CStr(g_rst_Princi!REGLAV_TIPVIA_PER)) & " " & Trim(g_rst_Princi!REGLAV_NOMVIA_PER) & " " & Trim(g_rst_Princi!REGLAV_NUMERO_PER) & IIf(Len(Trim(g_rst_Princi!REGLAV_INTDPT_PER)) > 0, " (" & Trim(g_rst_Princi!REGLAV_INTDPT_PER) & ")", "") & IIf(Len(Trim(g_rst_Princi!REGLAV_NOMZON_PER)) > 0, " - " & moddat_gf_Consulta_ParDes("202", CStr(g_rst_Princi!REGLAV_TIPZON_PER)) & " " & Trim(g_rst_Princi!REGLAV_NOMZON_PER), ""), 2, " ", 150), 1, 150)
            .Cells(r_int_Contad + 7, 40) = r_str_Direcc
            '41. Dpto. correspondiente a la direccion de la persona en cuyo nombre se realiza la operacion.
            .Cells(r_int_Contad + 7, 41) = gs_modsec_Genera(Mid(g_rst_Princi!SOLINM_UBIGEO, 1, 2), 2, " ", 2)
            '42. Provincia. correspondiente a la direccion de la persona en cuyo nombre se realiza la operacion.
            .Cells(r_int_Contad + 7, 42) = gs_modsec_Genera(Mid(g_rst_Princi!SOLINM_UBIGEO, 3, 2), 2, " ", 2)
            '43. Distrito correspondiente a la direccion de la persona en cuyo nombre se realiza la operacion.
            .Cells(r_int_Contad + 7, 43) = gs_modsec_Genera(Mid(g_rst_Princi!SOLINM_UBIGEO, 5, 2), 2, " ", 2)
            '44. Telefono de la persona en cuyo nombre se realiza la operacion.
            .Cells(r_int_Contad + 7, 44) = gs_modsec_Genera(Trim(g_rst_Princi!REGLAV_TELPER), 2, " ", 40)
            
            '*************** INFORMACION DE LA PERSONA A FAVOR DE QUIEN SE REALIZA LA OPERACION ***************
            '45. Tipo de relacion de la persona a favor de quien se realiza la operacion.
            .Cells(r_int_Contad + 7, 45) = gs_modsec_Genera("1", 2, " ", 1)
            '46. Condicion de residencia de la persona a favor de quien se realiza la operacion.
            .Cells(r_int_Contad + 7, 46) = gs_modsec_Genera(" ", 2, " ", 1)
            '47. Tipo de persona a favor de quien se realiza la operacion.
            .Cells(r_int_Contad + 7, 47) = gs_modsec_Genera("2", 2, " ", 1)
            '48. Tipo de documento de identidad de la persona a favor de quien se realiza la operacion.
            .Cells(r_int_Contad + 7, 48) = gs_modsec_Genera(Format(g_rst_Princi!REGLAV_TCLIFA, "#00"), 1, "0", 1)
            '49. Nro. de identidad de la persona a favor de quien se realiza la operacion.
            .Cells(r_int_Contad + 7, 49) = gs_modsec_Genera(Trim(g_rst_Princi!REGLAV_NCLIFA), 2, " ", 12)
            '50. Nro. de RUC de la persona a favor de quien se realiza la operacion.
            .Cells(r_int_Contad + 7, 50) = IIf(Len(Trim(g_rst_Princi!REGLAV_NRUCFA)) <> Null, " (" & 0 & ")", gs_modsec_Genera(" ", 2, " ", 11))
            '51. Ape. Paterno o razon social de la persona a favor de quien se realiza la operacion.
            .Cells(r_int_Contad + 7, 51) = gs_modsec_Genera(g_rst_Princi!REGLAV_PATFAV, 2, " ", 120)
            '52. Ape. Materno de la persona a favor de quien se realiza la operacion.
            .Cells(r_int_Contad + 7, 52) = gs_modsec_Genera(g_rst_Princi!REGLAV_MATFAV, 2, " ", 40)
            '53. Nombres de la persona a favor de quien se realiza la operacion.
            .Cells(r_int_Contad + 7, 53) = gs_modsec_Genera(g_rst_Princi!REGLAV_NOMFAV, 2, " ", 40)
            '54. Ocupacion de la persona a favor de quien se realiza la operacion.
            .Cells(r_int_Contad + 7, 54) = gs_modsec_Genera(g_rst_Princi!DATGEN_OCUPAC, 2, " ", 4)
            '55. Codigo CIIU de la ocupacion de la persona a favor de quien se realiza la operacion.
            .Cells(r_int_Contad + 7, 55) = gs_modsec_Genera(" ", 2, " ", 6)
            '56. Descripcion de la ocupacion de la persona a favor de quien se realiza la operacion.
            .Cells(r_int_Contad + 7, 56) = gs_modsec_Genera(" ", 2, " ", 104)
            '57. Cargo de la persona a favor de quien se realiza la operacion.
            .Cells(r_int_Contad + 7, 57) = gs_modsec_Genera(" ", 2, " ", 2)
            '58. Nombre y numero de via de la direccion de la persona a favor de quien se realiza la operacion.
            r_str_Direcc = Mid(gs_modsec_Genera(moddat_gf_Consulta_ParDes("201", CStr(g_rst_Princi!REGLAV_TIPVIA_FAV)) & " " & Trim(g_rst_Princi!REGLAV_NOMVIA_FAV) & " " & Trim(g_rst_Princi!REGLAV_NUMERO_FAV) & IIf(Len(Trim(g_rst_Princi!REGLAV_INTDPT_FAV)) > 0, " (" & Trim(g_rst_Princi!REGLAV_INTDPT_FAV) & ")", "") & IIf(Len(Trim(g_rst_Princi!REGLAV_NOMZON_FAV)) > 0, " - " & moddat_gf_Consulta_ParDes("202", CStr(g_rst_Princi!REGLAV_TIPZON_FAV)) & " " & Trim(g_rst_Princi!REGLAV_NOMZON_FAV), ""), 2, " ", 150), 1, 150)
            .Cells(r_int_Contad + 7, 58) = r_str_Direcc
            '59. Dpto. correspondiente a la direccion de la persona a favor de quien se realiza la operacion.
            .Cells(r_int_Contad + 7, 59) = gs_modsec_Genera(Mid(g_rst_Princi!SOLINM_UBIGEO, 1, 2), 2, " ", 2)
            '60. Provincia. correspondiente a la direccion de la persona a favor de quien se realiza la operacion.
            .Cells(r_int_Contad + 7, 60) = gs_modsec_Genera(Mid(g_rst_Princi!SOLINM_UBIGEO, 3, 2), 2, " ", 2)
            '61. Distrito correspondiente a la direccion de la persona a favor de quien se realiza la operacion.
            .Cells(r_int_Contad + 7, 61) = gs_modsec_Genera(Mid(g_rst_Princi!SOLINM_UBIGEO, 5, 2), 2, " ", 2)
            '62. Telefono de la persona a favor de quien se realiza la operacion.
            .Cells(r_int_Contad + 7, 62) = gs_modsec_Genera(Trim(g_rst_Princi!REGLAV_TELFAV), 2, " ", 40)
            
            '*************** INFORMACION DE LA OPERACION ***************
            '63. Tipo de fondo con que se realizo la operacion.
            If Not g_rst_Princi!REGLAV_TIPFON Then
               .Cells(r_int_Contad + 7, 63) = gs_modsec_Genera(CInt(g_rst_Princi!REGLAV_TIPFON), 2, " ", 1)
            Else
               .Cells(r_int_Contad + 7, 63) = gs_modsec_Genera(" ", 2, " ", 1)
            End If
            '64. Tipo de operacion con que se realizo.
            .Cells(r_int_Contad + 7, 64) = Trim(g_rst_Princi!REGLAV_TIPOPE)
            '65. Descripcion del tipo de operacion que se realizo.
            .Cells(r_int_Contad + 7, 65) = gs_modsec_Genera(" ", 2, " ", 40)
            '66. Origen de los fondos involucrados en la operacion.
            .Cells(r_int_Contad + 7, 66) = IIf(Len(Trim(g_rst_Princi!REGLAV_DESORG)) <> 0, gs_modsec_Genera(Trim(g_rst_Princi!REGLAV_DESORG), 2, " ", 80), gs_modsec_Genera(" ", 2, " ", 80))
            '67. Moneda que se realizo la operacion.
            .Cells(r_int_Contad + 7, 67) = gs_modsec_Genera(r_str_TipMon, 2, " ", 1)
            '68. Descripcion de la moneda que se realizo la operacion.
            .Cells(r_int_Contad + 7, 68) = gs_modsec_Genera(r_str_NomMon, 2, " ", 40)
            '69. Monto que se realizo la operacion.
            .Cells(r_int_Contad + 7, 69) = Format(g_rst_Princi!REGLAV_IMPOPE, "000000000000000.00")
            '70. Tipo de cambio que se realizo la operacion.
            .Cells(r_int_Contad + 7, 70) = Format(g_rst_Princi!REGLAV_TIPCAM, "00.000")
            
            'Entidad/Cuenta involucrada en la operación/producto/servicio utilizado por la persona en cuyo nombre se realiza la operación
            '71. Codigo de empresa supervisada
            .Cells(r_int_Contad + 7, 71) = gs_modsec_Genera("00006", 2, " ", 5)
            '72. Tipo de cuenta supervisada
            .Cells(r_int_Contad + 7, 72) = gs_modsec_Genera("1", 2, " ", 1)
            '73. Cod. cuenta interbancario (CCI) o Nro. de Cheque, numero de cuenta o codigo de producto involucrado en la operacion.
            .Cells(r_int_Contad + 7, 73) = gs_modsec_Genera("001103690200090532", 2, " ", 20)
            '74. Entidad del exterior involucrada en la operacion.
            .Cells(r_int_Contad + 7, 74) = gs_modsec_Genera(" ", 2, " ", 150)
            
            'Entidad/Cuenta involucrada en la operación/producto/servicio utilizado por la persona a favor de quien se realiza la operación
            '75. Codigo de empresa supervisada
            .Cells(r_int_Contad + 7, 75) = gs_modsec_Genera("00240", 2, " ", 5)
            '76. Tipo de cuenta supervisada
            .Cells(r_int_Contad + 7, 76) = gs_modsec_Genera("1", 2, " ", 1)
            '77. Cod. cuenta interbancario (CCI) o Nro. de Cheque, numero de cuenta o codigo de producto involucrado en la operacion.
            .Cells(r_int_Contad + 7, 77) = gs_modsec_Genera("001103690200090532", 2, " ", 20)
            '78. Entidad del exterior involucrada en la operacion.
            .Cells(r_int_Contad + 7, 78) = gs_modsec_Genera(" ", 2, " ", 150)
            
            '79. Alcance de la operacion.
            .Cells(r_int_Contad + 7, 79) = gs_modsec_Genera("1", 2, " ", 1)
            '80. Codigo del pais origen.
            .Cells(r_int_Contad + 7, 80) = gs_modsec_Genera(" ", 2, " ", 2)
            '81. Codigo del pais destino.
            .Cells(r_int_Contad + 7, 81) = gs_modsec_Genera(" ", 2, " ", 2)
            '82. Intermediario de la operacion.
            .Cells(r_int_Contad + 7, 82) = gs_modsec_Genera("2", 2, " ", 1)
            '83. Forma a traves del cual se realizo la operacion.
            .Cells(r_int_Contad + 7, 83) = gs_modsec_Genera("1", 2, " ", 1)
            '84. Descripcion de la forma a traves del cual se realizo la operacion.
            .Cells(r_int_Contad + 7, 84) = gs_modsec_Genera(" ", 2, " ", 40)
            
            r_int_Contad = r_int_Contad + 1
            g_rst_Princi.MoveNext
            DoEvents
         Loop
      End If
      
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
   End With
   
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub

