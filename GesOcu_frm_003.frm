VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.MDIForm frm_MnuPri_01 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   6795
   ClientLeft      =   4560
   ClientTop       =   3225
   ClientWidth     =   13425
   Icon            =   "GesOcu_frm_003.frx":0000
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin Threed.SSPanel SSPanel1 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13425
      _Version        =   65536
      _ExtentX        =   23680
      _ExtentY        =   1138
      _StockProps     =   15
      BackColor       =   -2147483633
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
      Begin VB.CommandButton cmd_CamCon 
         Height          =   585
         Left            =   30
         Picture         =   "GesOcu_frm_003.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Salir"
         Top             =   30
         Width           =   585
      End
      Begin VB.CommandButton cmd_Salida 
         Height          =   585
         Left            =   630
         Picture         =   "GesOcu_frm_003.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Salir"
         Top             =   30
         Width           =   585
      End
   End
   Begin Threed.SSPanel SSPanel2 
      Align           =   2  'Align Bottom
      Height          =   390
      Left            =   0
      TabIndex        =   3
      Top             =   6405
      Width           =   13425
      _Version        =   65536
      _ExtentX        =   23680
      _ExtentY        =   688
      _StockProps     =   15
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   0
      BevelOuter      =   0
      BevelInner      =   1
      Begin Threed.SSPanel pnl_EntDat 
         Height          =   315
         Left            =   30
         TabIndex        =   4
         Top             =   30
         Width           =   3900
         _Version        =   65536
         _ExtentX        =   6879
         _ExtentY        =   556
         _StockProps     =   15
         Caption         =   "lm_db_db1 - prod1"
         ForeColor       =   32768
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.26
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
         Font3D          =   2
      End
      Begin Threed.SSPanel pnl_NumVer 
         Height          =   315
         Left            =   3960
         TabIndex        =   5
         Top             =   30
         Width           =   2100
         _Version        =   65536
         _ExtentX        =   3704
         _ExtentY        =   556
         _StockProps     =   15
         Caption         =   "rev. 008-1028.1"
         ForeColor       =   32768
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.26
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
         Font3D          =   2
      End
   End
   Begin VB.Menu mnuOpe 
      Caption         =   "Operaciones"
      Begin VB.Menu mnuOpe_Opcion 
         Caption         =   "F0501-01 Registro de Operaciones - RO"
         Index           =   1
      End
   End
   Begin VB.Menu mnuRep 
      Caption         =   "Reportes"
      Begin VB.Menu mnuRep_Opcion 
         Caption         =   "Reporte de Lavado de Activos"
         Index           =   1
      End
      Begin VB.Menu mnuRep_Opcion 
         Caption         =   "Reporte Inspektor"
         Index           =   2
      End
   End
End
Attribute VB_Name = "frm_MnuPri_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_CamCon_Click()
   If modgen_g_str_CodUsu <> "DESARROLLO" Then
      frm_IdeUsu_02.Show 1
   End If
End Sub

Private Sub cmd_Salida_Click()
   If MsgBox("¿Está seguro de salir de la Plataforma?", vbExclamation + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) = vbYes Then
      Call gs_Desconecta_Servidor
      End
   End If
End Sub

Private Sub MDIForm_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   
'   Call fs_HabSeg
   
   Screen.MousePointer = 0
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
   If MsgBox("¿Está seguro de salir de la Plataforma?", vbExclamation + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) = vbYes Then
      Call gs_Desconecta_Servidor
      End
   Else
      Cancel = True
   End If
End Sub

Private Sub mnuOpe_Opcion_Click(Index As Integer)
   Select Case Index
      Case 1
         'F0501-01 Registro de Operaciones - RO
         frm_RegLav_01.Show 1
   End Select
End Sub

Private Sub mnuRep_Opcion_Click(Index As Integer)
   Select Case Index
      Case 1
         'Reporte de Lavado de Activos
         frm_RepLav_01.Show 1
         
       Case 2
         'Reporte de Lavado de Activos
         'frm_RepLav_01.Show 1
         frm_RepIns_01.Show 1
         
            
         
   End Select
End Sub

Private Sub fs_HabSeg()
Dim r_int_Posici     As Integer
Dim r_str_CodMen     As String
   
   'pnl_Seg_NomUsu.Caption = modgen_g_str_CodUsu
   pnl_NumVer.Caption = modgen_g_str_NumRev
   pnl_EntDat.Caption = moddat_g_str_NomEsq & " - " & UCase(moddat_g_str_EntDat)
   
   'Desactivando todas las opciones
   For r_int_Posici = 1 To mnuOpe_Opcion.Count
      If mnuOpe_Opcion(r_int_Posici).Caption <> "-" Then
         mnuOpe_Opcion(r_int_Posici).Enabled = False
      End If
   Next r_int_Posici
   
   For r_int_Posici = 1 To mnuRep_Opcion.Count
      If mnuRep_Opcion(r_int_Posici).Caption <> "-" Then
         mnuRep_Opcion(r_int_Posici).Enabled = False
      End If
   Next r_int_Posici
   
   
   'Verificando si todas las Opciones están habilitadas
   g_str_Parame = "SELECT * FROM SEG_PLTOPC WHERE "
   g_str_Parame = g_str_Parame & "PLTOPC_CODPLT = '" & UCase(App.EXEName) & "' AND "
   g_str_Parame = g_str_Parame & "PLTOPC_FLGMEN = 2 "
   g_str_Parame = g_str_Parame & "ORDER BY PLTOPC_CODMEN ASC, PLTOPC_CODSUB ASC "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If

   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      
      Do While Not g_rst_Princi.EOF
         r_str_CodMen = Trim(g_rst_Princi!PLTOPC_CODMEN)
         
         Do While Not g_rst_Princi.EOF And r_str_CodMen = Trim(g_rst_Princi!PLTOPC_CODMEN)
            Select Case r_str_CodMen
               Case "MNUOPE": mnuOpe_Opcion(CInt(g_rst_Princi!PLTOPC_CODSUB)).Visible = IIf(g_rst_Princi!PLTOPC_SITUAC = 1, True, False)
               Case "MNUREP": mnuRep_Opcion(CInt(g_rst_Princi!PLTOPC_CODSUB)).Visible = IIf(g_rst_Princi!PLTOPC_SITUAC = 1, True, False)
            End Select
            
            g_rst_Princi.MoveNext
            
            If g_rst_Princi.EOF Then
               Exit Do
            End If
         Loop
      Loop
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   'Verificando por Plantilla de Acceso
   g_str_Parame = "SELECT * FROM SEG_PLTPLA WHERE "
   g_str_Parame = g_str_Parame & "PLTPLA_CODPLT = '" & UCase(App.EXEName) & "' AND "
   g_str_Parame = g_str_Parame & "PLTPLA_TIPUSU = '" & CStr(modgen_g_int_TipUsu) & "' "
   g_str_Parame = g_str_Parame & "ORDER BY PLTPLA_CODMEN ASC, PLTPLA_CODSUB ASC "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If

   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      
      Do While Not g_rst_Princi.EOF
         r_str_CodMen = Trim(g_rst_Princi!PLTPLA_CODMEN)
         
         Do While Not g_rst_Princi.EOF And r_str_CodMen = Trim(g_rst_Princi!PLTPLA_CODMEN)
            Select Case r_str_CodMen
               Case "MNUOPE": mnuOpe_Opcion(CInt(g_rst_Princi!PLTPLA_CODSUB)).Enabled = True
               Case "MNUREP": mnuRep_Opcion(CInt(g_rst_Princi!PLTPLA_CODSUB)).Enabled = True
            End Select
            
            g_rst_Princi.MoveNext
            
            If g_rst_Princi.EOF Then
               Exit Do
            End If
         Loop
      Loop
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing

   'Verificando por Personalización de Opciones
   g_str_Parame = "SELECT * FROM SEG_PLTUSU WHERE "
   g_str_Parame = g_str_Parame & "PLTUSU_CODPLT = '" & UCase(App.EXEName) & "' "
   g_str_Parame = g_str_Parame & "ORDER BY PLTUSU_CODMEN ASC, PLTUSU_CODSUB ASC "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If

   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      
      Do While Not g_rst_Princi.EOF
         r_str_CodMen = Trim(g_rst_Princi!PLTUSU_CODMEN)
         
         Do While Not g_rst_Princi.EOF And r_str_CodMen = Trim(g_rst_Princi!PLTUSU_CODMEN)
            Select Case r_str_CodMen
               Case "MNUOPE": mnuOpe_Opcion(CInt(g_rst_Princi!PLTUSU_CODSUB)).Enabled = True
               Case "MNUREP": mnuRep_Opcion(CInt(g_rst_Princi!PLTUSU_CODSUB)).Enabled = True
            End Select
            
            g_rst_Princi.MoveNext
            
            If g_rst_Princi.EOF Then
               Exit Do
            End If
         Loop
      Loop
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

