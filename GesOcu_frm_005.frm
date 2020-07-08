VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frm_RegLav_03 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   10155
   ClientLeft      =   3000
   ClientTop       =   1320
   ClientWidth     =   11910
   Icon            =   "GesOcu_frm_005.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10155
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   10215
      Left            =   0
      TabIndex        =   72
      Top             =   0
      Width           =   11925
      _Version        =   65536
      _ExtentX        =   21034
      _ExtentY        =   18018
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
         TabIndex        =   73
         Top             =   30
         Width           =   11835
         _Version        =   65536
         _ExtentX        =   20876
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
            TabIndex        =   74
            Top             =   150
            Width           =   3795
            _Version        =   65536
            _ExtentX        =   6694
            _ExtentY        =   476
            _StockProps     =   15
            Caption         =   "F0501-01 Registro de Operaciones - RO"
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
            Picture         =   "GesOcu_frm_005.frx":000C
            Top             =   120
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   645
         Left            =   30
         TabIndex        =   75
         Top             =   750
         Width           =   11835
         _Version        =   65536
         _ExtentX        =   20876
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
         Begin VB.CommandButton cmd_ComGra 
            Height          =   585
            Left            =   30
            Picture         =   "GesOcu_frm_005.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   70
            ToolTipText     =   "Grabar Comprobante Contable"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   11220
            Picture         =   "GesOcu_frm_005.frx":0758
            Style           =   1  'Graphical
            TabIndex        =   71
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel9 
         Height          =   810
         Left            =   30
         TabIndex        =   76
         Top             =   1440
         Width           =   11835
         _Version        =   65536
         _ExtentX        =   20876
         _ExtentY        =   1429
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
         Begin Threed.SSPanel pnl_Empres 
            Height          =   315
            Left            =   1530
            TabIndex        =   77
            Top             =   60
            Width           =   3465
            _Version        =   65536
            _ExtentX        =   6112
            _ExtentY        =   556
            _StockProps     =   15
            ForeColor       =   32768
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
            Alignment       =   1
         End
         Begin Threed.SSPanel pnl_Sucurs 
            Height          =   315
            Left            =   6675
            TabIndex        =   78
            Top             =   60
            Width           =   1695
            _Version        =   65536
            _ExtentX        =   2990
            _ExtentY        =   556
            _StockProps     =   15
            ForeColor       =   32768
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
            Alignment       =   1
         End
         Begin Threed.SSPanel pnl_Period 
            Height          =   315
            Left            =   10005
            TabIndex        =   79
            Top             =   60
            Width           =   1740
            _Version        =   65536
            _ExtentX        =   3069
            _ExtentY        =   556
            _StockProps     =   15
            ForeColor       =   32768
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
            Alignment       =   1
         End
         Begin Threed.SSPanel pnl_Client 
            Height          =   315
            Left            =   1530
            TabIndex        =   88
            Top             =   435
            Width           =   10215
            _Version        =   65536
            _ExtentX        =   18018
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "1-07521154 / IKEHARA PUNK MIGUEL ANGEL"
            ForeColor       =   32768
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
            Alignment       =   1
         End
         Begin VB.Label Label7 
            Caption         =   "Cliente:"
            Height          =   315
            Left            =   60
            TabIndex        =   87
            Top             =   465
            Width           =   1125
         End
         Begin VB.Label Label1 
            Caption         =   "Sucursal:"
            Height          =   255
            Left            =   5610
            TabIndex        =   82
            Top             =   90
            Width           =   1425
         End
         Begin VB.Label lbl_NomEti 
            Caption         =   "Empresa:"
            Height          =   255
            Index           =   0
            Left            =   60
            TabIndex        =   81
            Top             =   60
            Width           =   1335
         End
         Begin VB.Label Label33 
            Caption         =   "Período:"
            Height          =   255
            Left            =   8940
            TabIndex        =   80
            Top             =   90
            Width           =   1215
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   435
         Left            =   30
         TabIndex        =   83
         Top             =   2295
         Width           =   11820
         _Version        =   65536
         _ExtentX        =   20849
         _ExtentY        =   767
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
         Begin VB.ComboBox cmb_MonOpe 
            Height          =   315
            Left            =   2040
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   60
            Width           =   2025
         End
         Begin EditLib.fpDateTime ipp_FecCtb 
            Height          =   315
            Left            =   6075
            TabIndex        =   1
            Top             =   60
            Width           =   1635
            _Version        =   196608
            _ExtentX        =   2884
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
         Begin MSMask.MaskEdBox msk_HorOpe 
            Height          =   315
            Left            =   9480
            TabIndex        =   169
            Top             =   60
            Width           =   1050
            _ExtentX        =   1852
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   8
            Mask            =   "##:##:##"
            PromptChar      =   " "
         End
         Begin VB.Label Label2 
            Caption         =   "(hh:mm:ss)"
            Height          =   255
            Left            =   10635
            TabIndex        =   171
            Top             =   120
            Width           =   795
         End
         Begin VB.Label Label6 
            Caption         =   "Hora Operación:"
            Height          =   255
            Left            =   8130
            TabIndex        =   86
            Top             =   120
            Width           =   1515
         End
         Begin VB.Label Label4 
            Caption         =   "Fecha Operación:"
            Height          =   255
            Left            =   4425
            TabIndex        =   85
            Top             =   120
            Width           =   1515
         End
         Begin VB.Label Label3 
            Caption         =   "Modalidad de Operación:"
            Height          =   255
            Left            =   60
            TabIndex        =   84
            Top             =   120
            Width           =   2025
         End
      End
      Begin Threed.SSPanel SSPanel16 
         Height          =   5145
         Left            =   30
         TabIndex        =   89
         Top             =   2775
         Width           =   11820
         _Version        =   65536
         _ExtentX        =   20849
         _ExtentY        =   9075
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
         Begin TabDlg.SSTab SSTab1 
            Height          =   5025
            Left            =   60
            TabIndex        =   90
            Top             =   60
            Width           =   11715
            _ExtentX        =   20664
            _ExtentY        =   8864
            _Version        =   393216
            Style           =   1
            TabHeight       =   520
            TabCaption(0)   =   "Persona que fisicamente realiza la Operación"
            TabPicture(0)   =   "GesOcu_frm_005.frx":0B9A
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "SSPanel11"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).Control(1)=   "SSPanel12"
            Tab(0).Control(1).Enabled=   0   'False
            Tab(0).Control(2)=   "SSPanel10"
            Tab(0).Control(2).Enabled=   0   'False
            Tab(0).ControlCount=   3
            TabCaption(1)   =   "Persona en cuyo nombre se realiza la Operación"
            TabPicture(1)   =   "GesOcu_frm_005.frx":0BB6
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "SSPanel8"
            Tab(1).Control(1)=   "SSPanel13"
            Tab(1).Control(2)=   "SSPanel14"
            Tab(1).Control(3)=   "SSPanel15"
            Tab(1).ControlCount=   4
            TabCaption(2)   =   "Persona a favor de quien se realiza la Operación"
            TabPicture(2)   =   "GesOcu_frm_005.frx":0BD2
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "SSPanel21"
            Tab(2).Control(1)=   "SSPanel20"
            Tab(2).Control(2)=   "SSPanel18"
            Tab(2).Control(3)=   "SSPanel17"
            Tab(2).ControlCount=   4
            Begin Threed.SSPanel SSPanel10 
               Height          =   465
               Left            =   60
               TabIndex        =   93
               Top             =   360
               Width           =   11595
               _Version        =   65536
               _ExtentX        =   20452
               _ExtentY        =   820
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
               Begin VB.TextBox txt_DocFis 
                  Height          =   315
                  Left            =   8190
                  MaxLength       =   12
                  TabIndex        =   3
                  Text            =   "Text1"
                  Top             =   60
                  Width           =   3315
               End
               Begin VB.ComboBox cmb_TipFis 
                  Height          =   315
                  Left            =   1920
                  Style           =   2  'Dropdown List
                  TabIndex        =   2
                  Top             =   90
                  Width           =   3315
               End
               Begin VB.Label Label17 
                  Caption         =   "Nro. Doc. Id.:"
                  Height          =   285
                  Left            =   6180
                  TabIndex        =   95
                  Top             =   150
                  Width           =   1065
               End
               Begin VB.Label Label18 
                  Caption         =   "Tipo Docum. Identidad:"
                  Height          =   315
                  Left            =   60
                  TabIndex        =   94
                  Top             =   120
                  Width           =   1845
               End
            End
            Begin Threed.SSPanel SSPanel12 
               Height          =   1455
               Left            =   60
               TabIndex        =   96
               Top             =   870
               Width           =   11595
               _Version        =   65536
               _ExtentX        =   20452
               _ExtentY        =   2566
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
               Begin VB.TextBox txt_PatFis 
                  Height          =   315
                  Left            =   1920
                  MaxLength       =   30
                  TabIndex        =   4
                  Text            =   "Text1"
                  Top             =   60
                  Width           =   3315
               End
               Begin VB.TextBox txt_MatFis 
                  Height          =   315
                  Left            =   1920
                  MaxLength       =   30
                  TabIndex        =   5
                  Text            =   "Text1"
                  Top             =   390
                  Width           =   3315
               End
               Begin VB.TextBox txt_NomFis 
                  Height          =   315
                  Left            =   1920
                  MaxLength       =   30
                  TabIndex        =   7
                  Text            =   "Text1"
                  Top             =   720
                  Width           =   3315
               End
               Begin VB.TextBox txt_CasFis 
                  Height          =   315
                  Left            =   8160
                  MaxLength       =   30
                  TabIndex        =   6
                  Text            =   "Text1"
                  Top             =   420
                  Width           =   3315
               End
               Begin VB.ComboBox cmb_OcuFis 
                  Height          =   315
                  Left            =   1920
                  Style           =   2  'Dropdown List
                  TabIndex        =   8
                  Top             =   1050
                  Width           =   3315
               End
               Begin VB.TextBox txt_TelFis 
                  Height          =   315
                  Left            =   8160
                  MaxLength       =   12
                  TabIndex        =   9
                  Text            =   "Text1"
                  Top             =   1050
                  Width           =   3315
               End
               Begin VB.Label Label37 
                  Caption         =   "Apellido Paterno:"
                  Height          =   285
                  Left            =   90
                  TabIndex        =   102
                  Top             =   90
                  Width           =   1485
               End
               Begin VB.Label Label36 
                  Caption         =   "Apellido Materno:"
                  Height          =   285
                  Left            =   90
                  TabIndex        =   101
                  Top             =   450
                  Width           =   1485
               End
               Begin VB.Label Label32 
                  Caption         =   "Nombres:"
                  Height          =   285
                  Left            =   90
                  TabIndex        =   100
                  Top             =   780
                  Width           =   1485
               End
               Begin VB.Label Label31 
                  Caption         =   "Apellido Casada:"
                  Height          =   285
                  Left            =   6180
                  TabIndex        =   99
                  Top             =   420
                  Width           =   1485
               End
               Begin VB.Label Label27 
                  Caption         =   "Ocupación:"
                  Height          =   285
                  Left            =   90
                  TabIndex        =   98
                  Top             =   1110
                  Width           =   1485
               End
               Begin VB.Label Label16 
                  Caption         =   "Teléfono:"
                  Height          =   285
                  Left            =   6180
                  TabIndex        =   97
                  Top             =   1110
                  Width           =   1485
               End
            End
            Begin Threed.SSPanel SSPanel11 
               Height          =   1785
               Left            =   60
               TabIndex        =   103
               Top             =   2370
               Width           =   11595
               _Version        =   65536
               _ExtentX        =   20452
               _ExtentY        =   3149
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
               Begin VB.ComboBox cmb_ViaFis 
                  Height          =   315
                  Left            =   1920
                  Style           =   2  'Dropdown List
                  TabIndex        =   10
                  Top             =   60
                  Width           =   3315
               End
               Begin VB.TextBox txt_ViaFis 
                  Height          =   315
                  Left            =   1920
                  MaxLength       =   120
                  TabIndex        =   11
                  Text            =   "txt_ViaFis"
                  Top             =   390
                  Width           =   3315
               End
               Begin VB.TextBox txt_NroFis 
                  Height          =   315
                  Left            =   8160
                  MaxLength       =   30
                  TabIndex        =   12
                  Text            =   "Text1"
                  Top             =   390
                  Width           =   1640
               End
               Begin VB.TextBox txt_IntFis 
                  Height          =   315
                  Left            =   9840
                  MaxLength       =   30
                  TabIndex        =   13
                  Text            =   "Text1"
                  Top             =   390
                  Width           =   1640
               End
               Begin VB.ComboBox cmb_ZonFis 
                  Height          =   315
                  Left            =   1920
                  Style           =   2  'Dropdown List
                  TabIndex        =   14
                  Top             =   720
                  Width           =   3315
               End
               Begin VB.TextBox txt_ZonFis 
                  Height          =   315
                  Left            =   8160
                  MaxLength       =   120
                  TabIndex        =   15
                  Text            =   "txt_ZonFis"
                  Top             =   720
                  Width           =   3315
               End
               Begin VB.ComboBox cmb_DptFis 
                  Height          =   315
                  ItemData        =   "GesOcu_frm_005.frx":0BEE
                  Left            =   1920
                  List            =   "GesOcu_frm_005.frx":0BF0
                  Style           =   2  'Dropdown List
                  TabIndex        =   16
                  Top             =   1050
                  Width           =   3315
               End
               Begin VB.ComboBox cmb_PrvFis 
                  Height          =   315
                  Left            =   8160
                  Style           =   2  'Dropdown List
                  TabIndex        =   17
                  Top             =   1050
                  Width           =   3315
               End
               Begin VB.ComboBox cmb_DstFis 
                  Height          =   315
                  Left            =   1920
                  Style           =   2  'Dropdown List
                  TabIndex        =   18
                  Top             =   1380
                  Width           =   3315
               End
               Begin VB.TextBox txt_RefFis 
                  Height          =   315
                  Left            =   8160
                  MaxLength       =   250
                  TabIndex        =   19
                  Text            =   "Text1"
                  Top             =   1380
                  Width           =   3315
               End
               Begin VB.Label Label19 
                  Caption         =   "Tipo de Vía:"
                  Height          =   315
                  Left            =   60
                  TabIndex        =   112
                  Top             =   60
                  Width           =   1905
               End
               Begin VB.Label Label20 
                  Caption         =   "Nombre Vía:"
                  Height          =   285
                  Left            =   60
                  TabIndex        =   111
                  Top             =   390
                  Width           =   1485
               End
               Begin VB.Label Label21 
                  Caption         =   "Nro - Int/Dpto/Mza/Lote:"
                  Height          =   285
                  Left            =   6180
                  TabIndex        =   110
                  Top             =   390
                  Width           =   2055
               End
               Begin VB.Label Label22 
                  Caption         =   "Tipo de Zona:"
                  Height          =   315
                  Left            =   60
                  TabIndex        =   109
                  Top             =   720
                  Width           =   1905
               End
               Begin VB.Label Label23 
                  Caption         =   "Nombre Zona:"
                  Height          =   285
                  Left            =   6180
                  TabIndex        =   108
                  Top             =   720
                  Width           =   1485
               End
               Begin VB.Label Label24 
                  Caption         =   "Departamento:"
                  Height          =   315
                  Left            =   60
                  TabIndex        =   107
                  Top             =   1050
                  Width           =   1905
               End
               Begin VB.Label Label25 
                  Caption         =   "Provincia:"
                  Height          =   315
                  Left            =   6180
                  TabIndex        =   106
                  Top             =   1050
                  Width           =   1905
               End
               Begin VB.Label Label26 
                  Caption         =   "Distrito:"
                  Height          =   315
                  Left            =   60
                  TabIndex        =   105
                  Top             =   1380
                  Width           =   1905
               End
               Begin VB.Label Label28 
                  Caption         =   "Referencia:"
                  Height          =   285
                  Left            =   6180
                  TabIndex        =   104
                  Top             =   1380
                  Width           =   1485
               End
            End
            Begin Threed.SSPanel SSPanel8 
               Height          =   795
               Left            =   -74940
               TabIndex        =   118
               Top             =   840
               Width           =   11595
               _Version        =   65536
               _ExtentX        =   20452
               _ExtentY        =   1402
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
               Begin VB.ComboBox cmb_FlgPer 
                  Height          =   315
                  Left            =   1920
                  Style           =   2  'Dropdown List
                  TabIndex        =   22
                  Top             =   60
                  Width           =   1065
               End
               Begin VB.ComboBox cmb_JurPer 
                  Height          =   315
                  Left            =   1920
                  Style           =   2  'Dropdown List
                  TabIndex        =   23
                  Top             =   420
                  Width           =   3315
               End
               Begin VB.TextBox txt_JurPer 
                  Height          =   315
                  Left            =   8190
                  MaxLength       =   12
                  TabIndex        =   24
                  Text            =   "Text1"
                  Top             =   390
                  Width           =   3315
               End
               Begin VB.Label Label13 
                  Caption         =   "Persona Jurídica:"
                  Height          =   315
                  Left            =   90
                  TabIndex        =   121
                  Top             =   90
                  Width           =   1845
               End
               Begin VB.Label Label12 
                  Caption         =   "Tipo Docum. Identidad:"
                  Height          =   315
                  Left            =   60
                  TabIndex        =   120
                  Top             =   450
                  Width           =   1845
               End
               Begin VB.Label Label10 
                  Caption         =   "Nro. Doc. Id.:"
                  Height          =   285
                  Left            =   6180
                  TabIndex        =   119
                  Top             =   480
                  Width           =   1065
               End
            End
            Begin Threed.SSPanel SSPanel13 
               Height          =   1455
               Left            =   -74940
               TabIndex        =   122
               Top             =   1680
               Width           =   11595
               _Version        =   65536
               _ExtentX        =   20452
               _ExtentY        =   2566
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
               Begin VB.TextBox txt_TelPer 
                  Height          =   315
                  Left            =   8160
                  MaxLength       =   25
                  TabIndex        =   30
                  Text            =   "Text1"
                  Top             =   1050
                  Width           =   3315
               End
               Begin VB.ComboBox cmb_OcuPer 
                  Height          =   315
                  Left            =   1950
                  Style           =   2  'Dropdown List
                  TabIndex        =   29
                  Top             =   1050
                  Width           =   3315
               End
               Begin VB.TextBox txt_CasPer 
                  Height          =   315
                  Left            =   8160
                  MaxLength       =   30
                  TabIndex        =   27
                  Text            =   "Text1"
                  Top             =   420
                  Width           =   3315
               End
               Begin VB.TextBox txt_NomPer 
                  Height          =   315
                  Left            =   1920
                  MaxLength       =   30
                  TabIndex        =   28
                  Text            =   "Text1"
                  Top             =   720
                  Width           =   3315
               End
               Begin VB.TextBox txt_MatPer 
                  Height          =   315
                  Left            =   1920
                  MaxLength       =   30
                  TabIndex        =   26
                  Text            =   "Text1"
                  Top             =   390
                  Width           =   3315
               End
               Begin VB.TextBox txt_PatPer 
                  Height          =   315
                  Left            =   1920
                  MaxLength       =   30
                  TabIndex        =   25
                  Text            =   "Text1"
                  Top             =   60
                  Width           =   3315
               End
               Begin VB.Label Label40 
                  Caption         =   "Teléfono:"
                  Height          =   285
                  Left            =   6180
                  TabIndex        =   128
                  Top             =   1110
                  Width           =   1485
               End
               Begin VB.Label Label39 
                  Caption         =   "Ocupación:"
                  Height          =   285
                  Left            =   90
                  TabIndex        =   127
                  Top             =   1110
                  Width           =   1485
               End
               Begin VB.Label Label38 
                  Caption         =   "Apellido Casada:"
                  Height          =   285
                  Left            =   6180
                  TabIndex        =   126
                  Top             =   420
                  Width           =   1485
               End
               Begin VB.Label Label30 
                  Caption         =   "Nombres:"
                  Height          =   285
                  Left            =   90
                  TabIndex        =   125
                  Top             =   780
                  Width           =   1485
               End
               Begin VB.Label Label29 
                  Caption         =   "Apellido Materno:"
                  Height          =   285
                  Left            =   90
                  TabIndex        =   124
                  Top             =   450
                  Width           =   1485
               End
               Begin VB.Label Label14 
                  Caption         =   "Apellido Paterno:"
                  Height          =   285
                  Left            =   90
                  TabIndex        =   123
                  Top             =   90
                  Width           =   1485
               End
            End
            Begin Threed.SSPanel SSPanel14 
               Height          =   1785
               Left            =   -74940
               TabIndex        =   129
               Top             =   3180
               Width           =   11595
               _Version        =   65536
               _ExtentX        =   20452
               _ExtentY        =   3149
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
               Begin VB.TextBox txt_RefPer 
                  Height          =   315
                  Left            =   8160
                  MaxLength       =   250
                  TabIndex        =   40
                  Text            =   "txt_RefPer"
                  Top             =   1380
                  Width           =   3315
               End
               Begin VB.ComboBox cmb_DstPer 
                  Height          =   315
                  Left            =   1920
                  Style           =   2  'Dropdown List
                  TabIndex        =   39
                  Top             =   1380
                  Width           =   3315
               End
               Begin VB.ComboBox cmb_PrvPer 
                  Height          =   315
                  Left            =   8160
                  Style           =   2  'Dropdown List
                  TabIndex        =   38
                  Top             =   1050
                  Width           =   3315
               End
               Begin VB.ComboBox cmb_DptPer 
                  Height          =   315
                  Left            =   1920
                  Style           =   2  'Dropdown List
                  TabIndex        =   37
                  Top             =   1050
                  Width           =   3315
               End
               Begin VB.TextBox txt_ZonPer 
                  Height          =   315
                  Left            =   8160
                  MaxLength       =   120
                  TabIndex        =   36
                  Text            =   "Text1"
                  Top             =   720
                  Width           =   3315
               End
               Begin VB.ComboBox cmb_ZonPer 
                  Height          =   315
                  Left            =   1920
                  Style           =   2  'Dropdown List
                  TabIndex        =   35
                  Top             =   720
                  Width           =   3315
               End
               Begin VB.TextBox txt_intPer 
                  Height          =   315
                  Left            =   9840
                  MaxLength       =   30
                  TabIndex        =   34
                  Text            =   "Text1"
                  Top             =   390
                  Width           =   1640
               End
               Begin VB.TextBox txt_NroPer 
                  Height          =   315
                  Left            =   8160
                  MaxLength       =   30
                  TabIndex        =   33
                  Text            =   "Text1"
                  Top             =   390
                  Width           =   1640
               End
               Begin VB.TextBox txt_ViaPer 
                  Height          =   315
                  Left            =   1920
                  MaxLength       =   120
                  TabIndex        =   32
                  Text            =   "Text1"
                  Top             =   390
                  Width           =   3315
               End
               Begin VB.ComboBox cmb_ViaPer 
                  Height          =   315
                  Left            =   1920
                  Style           =   2  'Dropdown List
                  TabIndex        =   31
                  Top             =   60
                  Width           =   3315
               End
               Begin VB.Label Label49 
                  Caption         =   "Referencia:"
                  Height          =   285
                  Left            =   6180
                  TabIndex        =   138
                  Top             =   1380
                  Width           =   1485
               End
               Begin VB.Label Label48 
                  Caption         =   "Distrito:"
                  Height          =   315
                  Left            =   60
                  TabIndex        =   137
                  Top             =   1380
                  Width           =   1905
               End
               Begin VB.Label Label47 
                  Caption         =   "Provincia:"
                  Height          =   315
                  Left            =   6180
                  TabIndex        =   136
                  Top             =   1050
                  Width           =   1905
               End
               Begin VB.Label Label46 
                  Caption         =   "Departamento:"
                  Height          =   315
                  Left            =   60
                  TabIndex        =   135
                  Top             =   1050
                  Width           =   1905
               End
               Begin VB.Label Label45 
                  Caption         =   "Nombre Zona:"
                  Height          =   285
                  Left            =   6180
                  TabIndex        =   134
                  Top             =   720
                  Width           =   1485
               End
               Begin VB.Label Label44 
                  Caption         =   "Tipo de Zona:"
                  Height          =   315
                  Left            =   60
                  TabIndex        =   133
                  Top             =   720
                  Width           =   1905
               End
               Begin VB.Label Label43 
                  Caption         =   "Nro - Int/Dpto/Mza/Lote:"
                  Height          =   285
                  Left            =   6180
                  TabIndex        =   132
                  Top             =   390
                  Width           =   2055
               End
               Begin VB.Label Label42 
                  Caption         =   "Nombre Vía:"
                  Height          =   285
                  Left            =   60
                  TabIndex        =   131
                  Top             =   390
                  Width           =   1485
               End
               Begin VB.Label Label41 
                  Caption         =   "Tipo de Vía:"
                  Height          =   315
                  Left            =   60
                  TabIndex        =   130
                  Top             =   60
                  Width           =   1905
               End
            End
            Begin Threed.SSPanel SSPanel15 
               Height          =   435
               Left            =   -74940
               TabIndex        =   139
               Top             =   360
               Width           =   11595
               _Version        =   65536
               _ExtentX        =   20452
               _ExtentY        =   767
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
               Begin VB.TextBox txt_DocPer 
                  Height          =   315
                  Left            =   8190
                  MaxLength       =   12
                  TabIndex        =   21
                  Text            =   "Text1"
                  Top             =   60
                  Width           =   3315
               End
               Begin VB.ComboBox cmb_TipPer 
                  Height          =   315
                  Left            =   1920
                  Style           =   2  'Dropdown List
                  TabIndex        =   20
                  Top             =   60
                  Width           =   3315
               End
               Begin VB.Label Label51 
                  Caption         =   "Nro. Doc. Id.:"
                  Height          =   285
                  Left            =   6180
                  TabIndex        =   141
                  Top             =   90
                  Width           =   1065
               End
               Begin VB.Label Label50 
                  Caption         =   "Tipo Docum. Identidad:"
                  Height          =   315
                  Left            =   60
                  TabIndex        =   140
                  Top             =   90
                  Width           =   1845
               End
            End
            Begin Threed.SSPanel SSPanel17 
               Height          =   795
               Left            =   -74940
               TabIndex        =   142
               Top             =   840
               Width           =   11595
               _Version        =   65536
               _ExtentX        =   20452
               _ExtentY        =   1402
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
               Begin VB.ComboBox cmb_FlgFav 
                  Height          =   315
                  Left            =   1920
                  Style           =   2  'Dropdown List
                  TabIndex        =   43
                  Top             =   60
                  Width           =   1065
               End
               Begin VB.ComboBox cmb_JurFav 
                  Height          =   315
                  Left            =   1920
                  Style           =   2  'Dropdown List
                  TabIndex        =   44
                  Top             =   420
                  Width           =   3315
               End
               Begin VB.TextBox txt_JurFav 
                  Height          =   315
                  Left            =   8190
                  MaxLength       =   12
                  TabIndex        =   45
                  Text            =   "Text1"
                  Top             =   390
                  Width           =   3315
               End
               Begin VB.Label Label54 
                  Caption         =   "Persona Jurídica:"
                  Height          =   315
                  Left            =   90
                  TabIndex        =   145
                  Top             =   90
                  Width           =   1845
               End
               Begin VB.Label Label53 
                  Caption         =   "Tipo Docum. Identidad:"
                  Height          =   315
                  Left            =   60
                  TabIndex        =   144
                  Top             =   450
                  Width           =   1845
               End
               Begin VB.Label Label52 
                  Caption         =   "Nro. Doc. Id.:"
                  Height          =   285
                  Left            =   6180
                  TabIndex        =   143
                  Top             =   480
                  Width           =   1065
               End
            End
            Begin Threed.SSPanel SSPanel18 
               Height          =   1455
               Left            =   -74940
               TabIndex        =   146
               Top             =   1680
               Width           =   11595
               _Version        =   65536
               _ExtentX        =   20452
               _ExtentY        =   2566
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
               Begin VB.TextBox txt_TelFav 
                  Height          =   315
                  Left            =   8160
                  MaxLength       =   25
                  TabIndex        =   51
                  Text            =   "Text1"
                  Top             =   1050
                  Width           =   3315
               End
               Begin VB.ComboBox cmb_OcuFav 
                  Height          =   315
                  Left            =   1920
                  Style           =   2  'Dropdown List
                  TabIndex        =   50
                  Top             =   1050
                  Width           =   3315
               End
               Begin VB.TextBox txt_CasFav 
                  Height          =   315
                  Left            =   8160
                  MaxLength       =   30
                  TabIndex        =   48
                  Text            =   "Text1"
                  Top             =   420
                  Width           =   3315
               End
               Begin VB.TextBox txt_NomFav 
                  Height          =   315
                  Left            =   1920
                  MaxLength       =   30
                  TabIndex        =   49
                  Text            =   "Text1"
                  Top             =   720
                  Width           =   3315
               End
               Begin VB.TextBox txt_MatFav 
                  Height          =   315
                  Left            =   1920
                  MaxLength       =   30
                  TabIndex        =   47
                  Text            =   "Text1"
                  Top             =   390
                  Width           =   3315
               End
               Begin VB.TextBox txt_PatFav 
                  Height          =   315
                  Left            =   1920
                  MaxLength       =   30
                  TabIndex        =   46
                  Text            =   "Text1"
                  Top             =   60
                  Width           =   3315
               End
               Begin VB.Label Label60 
                  Caption         =   "Teléfono:"
                  Height          =   285
                  Left            =   6180
                  TabIndex        =   152
                  Top             =   1110
                  Width           =   1485
               End
               Begin VB.Label Label59 
                  Caption         =   "Ocupación:"
                  Height          =   285
                  Left            =   90
                  TabIndex        =   151
                  Top             =   1110
                  Width           =   1485
               End
               Begin VB.Label Label58 
                  Caption         =   "Apellido Casada:"
                  Height          =   285
                  Left            =   6180
                  TabIndex        =   150
                  Top             =   420
                  Width           =   1485
               End
               Begin VB.Label Label57 
                  Caption         =   "Nombres:"
                  Height          =   285
                  Left            =   90
                  TabIndex        =   149
                  Top             =   780
                  Width           =   1485
               End
               Begin VB.Label Label56 
                  Caption         =   "Apellido Materno:"
                  Height          =   285
                  Left            =   90
                  TabIndex        =   148
                  Top             =   450
                  Width           =   1485
               End
               Begin VB.Label Label55 
                  Caption         =   "Apellido Paterno:"
                  Height          =   285
                  Left            =   90
                  TabIndex        =   147
                  Top             =   90
                  Width           =   1485
               End
            End
            Begin Threed.SSPanel SSPanel20 
               Height          =   1785
               Left            =   -74940
               TabIndex        =   153
               Top             =   3180
               Width           =   11595
               _Version        =   65536
               _ExtentX        =   20452
               _ExtentY        =   3149
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
               Begin VB.TextBox txt_RefFav 
                  Height          =   315
                  Left            =   8160
                  MaxLength       =   250
                  TabIndex        =   61
                  Text            =   "Text1"
                  Top             =   1380
                  Width           =   3315
               End
               Begin VB.ComboBox cmb_DstFav 
                  Height          =   315
                  Left            =   1920
                  Style           =   2  'Dropdown List
                  TabIndex        =   60
                  Top             =   1380
                  Width           =   3315
               End
               Begin VB.ComboBox cmb_PrvFav 
                  Height          =   315
                  Left            =   8160
                  Style           =   2  'Dropdown List
                  TabIndex        =   59
                  Top             =   1050
                  Width           =   3315
               End
               Begin VB.ComboBox cmb_DptFav 
                  Height          =   315
                  Left            =   1920
                  Style           =   2  'Dropdown List
                  TabIndex        =   58
                  Top             =   1050
                  Width           =   3315
               End
               Begin VB.TextBox txt_ZonFav 
                  Height          =   315
                  Left            =   8160
                  MaxLength       =   120
                  TabIndex        =   57
                  Text            =   "Text1"
                  Top             =   720
                  Width           =   3315
               End
               Begin VB.ComboBox cmb_ZonFav 
                  Height          =   315
                  Left            =   1920
                  Style           =   2  'Dropdown List
                  TabIndex        =   56
                  Top             =   720
                  Width           =   3315
               End
               Begin VB.TextBox txt_IntFav 
                  Height          =   315
                  Left            =   9840
                  MaxLength       =   30
                  TabIndex        =   55
                  Text            =   "Text1"
                  Top             =   390
                  Width           =   1640
               End
               Begin VB.TextBox txt_NroFav 
                  Height          =   315
                  Left            =   8160
                  MaxLength       =   30
                  TabIndex        =   54
                  Text            =   "Text1"
                  Top             =   390
                  Width           =   1640
               End
               Begin VB.TextBox txt_ViaFav 
                  Height          =   315
                  Left            =   1920
                  MaxLength       =   120
                  TabIndex        =   53
                  Text            =   "Text1"
                  Top             =   390
                  Width           =   3315
               End
               Begin VB.ComboBox cmb_ViaFav 
                  Height          =   315
                  Left            =   1920
                  Style           =   2  'Dropdown List
                  TabIndex        =   52
                  Top             =   60
                  Width           =   3315
               End
               Begin VB.Label Label73 
                  Caption         =   "Referencia:"
                  Height          =   285
                  Left            =   6180
                  TabIndex        =   162
                  Top             =   1380
                  Width           =   1485
               End
               Begin VB.Label Label72 
                  Caption         =   "Distrito:"
                  Height          =   315
                  Left            =   60
                  TabIndex        =   161
                  Top             =   1380
                  Width           =   1905
               End
               Begin VB.Label Label67 
                  Caption         =   "Provincia:"
                  Height          =   315
                  Left            =   6180
                  TabIndex        =   160
                  Top             =   1050
                  Width           =   1905
               End
               Begin VB.Label Label66 
                  Caption         =   "Departamento:"
                  Height          =   315
                  Left            =   60
                  TabIndex        =   159
                  Top             =   1050
                  Width           =   1905
               End
               Begin VB.Label Label65 
                  Caption         =   "Nombre Zona:"
                  Height          =   285
                  Left            =   6180
                  TabIndex        =   158
                  Top             =   720
                  Width           =   1485
               End
               Begin VB.Label Label64 
                  Caption         =   "Tipo de Zona:"
                  Height          =   315
                  Left            =   60
                  TabIndex        =   157
                  Top             =   720
                  Width           =   1905
               End
               Begin VB.Label Label63 
                  Caption         =   "Nro - Int/Dpto/Mza/Lote:"
                  Height          =   285
                  Left            =   6180
                  TabIndex        =   156
                  Top             =   390
                  Width           =   2055
               End
               Begin VB.Label Label62 
                  Caption         =   "Nombre Vía:"
                  Height          =   285
                  Left            =   60
                  TabIndex        =   155
                  Top             =   390
                  Width           =   1485
               End
               Begin VB.Label Label61 
                  Caption         =   "Tipo de Vía:"
                  Height          =   315
                  Left            =   60
                  TabIndex        =   154
                  Top             =   60
                  Width           =   1905
               End
            End
            Begin Threed.SSPanel SSPanel21 
               Height          =   435
               Left            =   -74940
               TabIndex        =   163
               Top             =   360
               Width           =   11595
               _Version        =   65536
               _ExtentX        =   20452
               _ExtentY        =   767
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
               Begin VB.TextBox txt_DocFav 
                  Height          =   315
                  Left            =   8190
                  MaxLength       =   12
                  TabIndex        =   42
                  Text            =   "Text1"
                  Top             =   60
                  Width           =   3315
               End
               Begin VB.ComboBox cmb_TipFav 
                  Height          =   315
                  Left            =   1920
                  Style           =   2  'Dropdown List
                  TabIndex        =   41
                  Top             =   60
                  Width           =   3315
               End
               Begin VB.Label Label75 
                  Caption         =   "Nro. Doc. Id.:"
                  Height          =   285
                  Left            =   6180
                  TabIndex        =   165
                  Top             =   90
                  Width           =   1065
               End
               Begin VB.Label Label74 
                  Caption         =   "Tipo Docum. Identidad:"
                  Height          =   315
                  Left            =   60
                  TabIndex        =   164
                  Top             =   90
                  Width           =   1845
               End
            End
            Begin VB.Label Label11 
               Caption         =   "Observaciones:"
               Height          =   300
               Left            =   -74910
               TabIndex        =   92
               Top             =   1200
               Width           =   1155
            End
            Begin VB.Label Label15 
               Caption         =   "Observaciones:"
               Height          =   300
               Left            =   -74910
               TabIndex        =   91
               Top             =   1200
               Width           =   1155
            End
         End
      End
      Begin Threed.SSPanel SSPanel19 
         Height          =   2115
         Left            =   30
         TabIndex        =   113
         Top             =   7965
         Width           =   11835
         _Version        =   65536
         _ExtentX        =   20876
         _ExtentY        =   3731
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
         Begin VB.ComboBox cmb_TipFon 
            Height          =   315
            Left            =   2070
            Style           =   2  'Dropdown List
            TabIndex        =   68
            Top             =   1080
            Width           =   3315
         End
         Begin VB.ComboBox cmb_CodBan 
            Height          =   315
            Left            =   2070
            Style           =   2  'Dropdown List
            TabIndex        =   64
            Top             =   420
            Width           =   3315
         End
         Begin VB.ComboBox cmb_CtaBan 
            Height          =   315
            Left            =   8310
            Style           =   2  'Dropdown List
            TabIndex        =   65
            Top             =   420
            Width           =   3315
         End
         Begin VB.ComboBox cmb_TipOpe 
            Height          =   315
            Left            =   8310
            Style           =   2  'Dropdown List
            TabIndex        =   63
            Top             =   90
            Width           =   3315
         End
         Begin VB.ComboBox cmb_ValOpe 
            Height          =   315
            Left            =   2070
            Style           =   2  'Dropdown List
            TabIndex        =   62
            Top             =   90
            Width           =   3315
         End
         Begin VB.TextBox txt_OrgOpe 
            Height          =   315
            Left            =   2070
            MaxLength       =   80
            TabIndex        =   69
            Text            =   "Text1"
            Top             =   1440
            Width           =   9555
         End
         Begin EditLib.fpDoubleSingle ipp_OrgMto 
            Height          =   315
            Left            =   2070
            TabIndex        =   66
            Top             =   750
            Width           =   1365
            _Version        =   196608
            _ExtentX        =   2408
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
            ButtonStyle     =   0
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
            Text            =   "0.00"
            DecimalPlaces   =   2
            DecimalPoint    =   "."
            FixedPoint      =   -1  'True
            LeadZero        =   0
            MaxValue        =   "9000000000"
            MinValue        =   "-9000000000"
            NegFormat       =   1
            NegToggle       =   0   'False
            Separator       =   ","
            UseSeparator    =   -1  'True
            IncInt          =   1
            IncDec          =   1
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
         Begin EditLib.fpDoubleSingle ipp_TipCam 
            Height          =   315
            Left            =   8310
            TabIndex        =   67
            Top             =   780
            Width           =   1365
            _Version        =   196608
            _ExtentX        =   2408
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
            ButtonStyle     =   0
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
            Text            =   "0.000"
            DecimalPlaces   =   3
            DecimalPoint    =   "."
            FixedPoint      =   -1  'True
            LeadZero        =   0
            MaxValue        =   "9000000000"
            MinValue        =   "-9000000000"
            NegFormat       =   1
            NegToggle       =   0   'False
            Separator       =   ","
            UseSeparator    =   -1  'True
            IncInt          =   1
            IncDec          =   1
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
         Begin VB.Label Label5 
            Caption         =   "Tipo de Fondo:"
            Height          =   315
            Left            =   60
            TabIndex        =   172
            Top             =   1110
            Width           =   1845
         End
         Begin VB.Label pnl_TipMon 
            Height          =   285
            Left            =   3480
            TabIndex        =   170
            Top             =   780
            Width           =   615
         End
         Begin VB.Label Label35 
            Caption         =   "Tipo de Cambio:"
            Height          =   285
            Left            =   6300
            TabIndex        =   168
            Top             =   810
            Width           =   1815
         End
         Begin VB.Label Label9 
            Caption         =   "Nro. Cuenta:"
            Height          =   285
            Left            =   6300
            TabIndex        =   167
            Top             =   450
            Width           =   1335
         End
         Begin VB.Label Label8 
            Caption         =   "Banco:"
            Height          =   315
            Left            =   60
            TabIndex        =   166
            Top             =   450
            Width           =   1455
         End
         Begin VB.Label Label70 
            Caption         =   "Monto:"
            Height          =   285
            Left            =   90
            TabIndex        =   117
            Top             =   780
            Width           =   1815
         End
         Begin VB.Label Label34 
            Caption         =   "Origen de la Operación:"
            Height          =   285
            Left            =   60
            TabIndex        =   116
            Top             =   1470
            Width           =   1815
         End
         Begin VB.Label Label69 
            Caption         =   "Valor de la Operación:"
            Height          =   315
            Left            =   60
            TabIndex        =   115
            Top             =   120
            Width           =   1845
         End
         Begin VB.Label Label68 
            Caption         =   "Tipo de Operación:"
            Height          =   285
            Left            =   6300
            TabIndex        =   114
            Top             =   120
            Width           =   1815
         End
      End
   End
End
Attribute VB_Name = "frm_RegLav_03"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim l_arr_Ubigeo()   As moddat_tpo_Genera
Dim l_arr_OcuPer()   As modtac_tpo_Genera
Dim l_arr_OcuFis()   As modtac_tpo_Genera
Dim l_arr_OcuFav()   As modtac_tpo_Genera
Dim l_arr_CodBan()   As moddat_tpo_Genera
Dim l_arr_CtaBan()   As moddat_tpo_Genera
Dim r_arr_Ocupac()   As modtac_tpo_Ocupac

Dim l_str_Ubigeo     As String
Dim l_str_DptPer     As String
Dim l_str_DstPer     As String
Dim l_str_DptFis     As String
Dim l_str_DstFis     As String
Dim l_str_DptFav     As String
Dim l_str_DstFav     As String
Dim l_str_ProOcu     As String
Dim r_str_NumReg     As String
Dim l_int_FlgCmb     As Integer

' Declaración APi
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long

'Mensaje - Constante para establecer el ancho
Private Const CB_SETDROPPEDWIDTH = &H160

Private Sub txt_DocFis_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_PatFis)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
   End If
End Sub

Private Sub txt_TelFis_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_ViaFis)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
   End If
End Sub

Private Sub txt_PatFis_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_MatFis)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & "- '")
   End If
End Sub

Private Sub txt_MatFis_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_CasFis)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & "- '")
   End If
End Sub

Private Sub txt_CasFis_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_NomFis)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & "- '")
   End If
End Sub

Private Sub txt_NomFis_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_OcuFis)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & "- '")
   End If
End Sub

Private Sub txt_ViaFis_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_NroFis)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_. ,;:()/º")
   End If
End Sub

Private Sub cmb_PrvFis_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_DstFis)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & "- '")
   End If
End Sub

Private Sub txt_ZonFis_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_DptFis)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_. ,;:()/º")
   End If
End Sub

Private Sub txt_RefFis_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_TipPer)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_. ,;:()/º")
   End If
End Sub

Private Sub txt_NroFis_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_IntFis)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_. ,;:()/º")
   End If
End Sub

Private Sub txt_IntFis_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_ZonFis)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_. ,;:()/º")
   End If
End Sub

Private Sub txt_DocPer_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_PatPer)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
   End If
End Sub

Private Sub txt_TelPer_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_ViaPer)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
   End If
End Sub

Private Sub txt_PatPer_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_MatPer)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & "- '")
   End If
End Sub

Private Sub txt_MatPer_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_CasPer)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & "- '")
   End If
End Sub

Private Sub txt_CasPer_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_NomPer)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & "- '")
   End If
End Sub

Private Sub txt_NomPer_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_OcuPer)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & "- '")
   End If
End Sub

Private Sub txt_ViaPer_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_NroPer)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_. ,;:()/º")
   End If
End Sub

Private Sub cmb_PrvPer_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_DstPer)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & "- '")
   End If
End Sub

Private Sub txt_ZonPer_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_DptPer)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_. ,;:()/º")
   End If
End Sub

Private Sub txt_RefPer_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_TipFav)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_. ,;:()/º")
   End If
End Sub

Private Sub txt_NroPer_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_intPer)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_. ,;:()/º")
   End If
End Sub

Private Sub txt_IntPer_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_ZonPer)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_. ,;:()/º")
   End If
End Sub

Private Sub txt_DocFav_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_PatFav)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
   End If
End Sub

Private Sub txt_TelFav_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_ViaFav)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
   End If
End Sub

Private Sub txt_PatFav_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_MatFav)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & "- '")
   End If
End Sub

Private Sub txt_MatFav_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_CasFav)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & "- '")
   End If
End Sub

Private Sub txt_CasFav_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_NomFav)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & "- '")
   End If
End Sub

Private Sub txt_NomFav_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_OcuFav)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & "- '")
   End If
End Sub

Private Sub txt_ViaFav_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_NroFav)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_. ,;:()/º")
   End If
End Sub

Private Sub cmb_PrvFav_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_DstFav)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & "- '")
   End If
End Sub

Private Sub txt_ZonFav_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_DptFav)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_. ,;:()/º")
   End If
End Sub

Private Sub txt_RefFav_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_ValOpe)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_. ,;:()/º")
   End If
End Sub

Private Sub txt_NroFav_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_IntFav)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_. ,;:()/º")
   End If
End Sub

Private Sub txt_IntFav_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_ZonFav)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_. ,;:()/º")
   End If
End Sub

Private Sub cmd_ComGra_Click()

   If cmb_MonOpe.ListIndex = -1 Then
      MsgBox "Debe seleccionar la Modalidad de la Operación.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_MonOpe)
      Exit Sub
   End If
   If Len(Trim(msk_HorOpe)) = 0 Then
      MsgBox "Debe ingresar la Hora de la Operación.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(msk_HorOpe)
      Exit Sub
   End If
   
   If cmb_TipFis.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Documento de Identidad.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipFis)
      Exit Sub
   End If
   If Len(Trim(txt_DocFis)) = 0 Then
      MsgBox "Debe ingresar el Número de Documento de Identidad.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_DocFis)
      Exit Sub
   End If
   
   If cmb_OcuFis.ListIndex = -1 Then
      MsgBox "Debe seleccionar la ocupación de la persona que fisicamente realiza la operación.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_OcuFis)
      Exit Sub
   End If
   If cmb_OcuPer.ListIndex = -1 Then
      MsgBox "Debe seleccionar la ocupación de la persona en cuyo nombre se realiza la operación.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_OcuPer)
      Exit Sub
   End If
   If cmb_OcuFav.ListIndex = -1 Then
      MsgBox "Debe seleccionar la ocupación de la persona a favor de quien se realiza la operación.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_OcuFav)
      Exit Sub
   End If
   
   'Datos generales
   If cmb_ValOpe.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Valor de la Operación.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_ValOpe)
      Exit Sub
   End If
   If cmb_TipOpe.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Operación.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipOpe)
      Exit Sub
   End If
   If cmb_TipFon.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Fondo.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipFon)
      Exit Sub
   End If
   If cmb_CodBan.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Banco.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_CodBan)
      Exit Sub
   End If
   If cmb_CtaBan.ListIndex = -1 Then
      MsgBox "Debe seleccionar la Cuenta Bancaria.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_CtaBan)
      Exit Sub
   End If
   If ipp_OrgMto.Text <= 0 Then
      MsgBox "Debe ingresar el Mto. de la Operación.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_OrgMto)
      Exit Sub
   End If
               
   If MsgBox("¿Está seguro de grabar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   r_str_NumReg = ff_Genera_NumReg()
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
   
   Do While moddat_g_int_FlgGOK = False
      g_str_Parame = "USP_OPE_REGLAV ("
      g_str_Parame = g_str_Parame & modtac_g_int_PerMes & ", "
      g_str_Parame = g_str_Parame & modtac_g_int_PerAno & ", "
      g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumOpe & "', "
      g_str_Parame = g_str_Parame & 1 & ", "
      g_str_Parame = g_str_Parame & "'" & "000001" & "', "
      If moddat_g_int_FlgGrb = 1 Then
         g_str_Parame = g_str_Parame & "'" & CStr(r_str_NumReg) & "', "
      Else
         g_str_Parame = g_str_Parame & "'" & modtac_g_str_NroInt & "', "
      End If
      g_str_Parame = g_str_Parame & "'" & CStr(cmb_MonOpe.Text) & "', "
      g_str_Parame = g_str_Parame & "'" & 150131 & "', "
      g_str_Parame = g_str_Parame & Format(CDate(ipp_FecCtb.Text), "yyyymmdd") & ", "
      g_str_Parame = g_str_Parame & msk_HorOpe.Text & ", "
      g_str_Parame = g_str_Parame & CInt(cmb_TipFis.ItemData(cmb_TipFis.ListIndex)) & ", "
      g_str_Parame = g_str_Parame & "'" & Trim(txt_DocFis.Text) & "', "
      g_str_Parame = g_str_Parame & "'" & Trim(txt_PatFis.Text) & "', "
      g_str_Parame = g_str_Parame & "'" & Trim(txt_MatFis.Text) & "', "
      g_str_Parame = g_str_Parame & "'" & Trim(txt_NomFis.Text) & "', "
      g_str_Parame = g_str_Parame & "'" & cmb_OcuFis.ItemData(cmb_OcuFis.ListIndex) & "', "
      g_str_Parame = g_str_Parame & CInt(cmb_ViaFis.ItemData(cmb_ViaFis.ListIndex)) & ", "
      g_str_Parame = g_str_Parame & "'" & Trim(txt_ViaFis.Text) & "', "
      g_str_Parame = g_str_Parame & "'" & Trim(txt_NroFis.Text) & "', "
      g_str_Parame = g_str_Parame & "'" & Trim(txt_IntFis.Text) & "', "
      g_str_Parame = g_str_Parame & CInt(cmb_ZonFis.ItemData(cmb_ZonFis.ListIndex)) & ", "
      g_str_Parame = g_str_Parame & "'" & Trim(txt_ZonFis.Text) & "', "
      
      If Len(Trim(txt_TelFis.Text)) = 0 Then
         g_str_Parame = g_str_Parame & "'" & " " & "', "
      Else
         g_str_Parame = g_str_Parame & "'" & Trim(txt_TelFis.Text) & "', "
      End If
      
      g_str_Parame = g_str_Parame & cmb_TipPer.ItemData(cmb_TipPer.ListIndex) & ", "
      g_str_Parame = g_str_Parame & "'" & Trim(txt_DocPer.Text) & "', "
      
      If CStr(cmb_FlgPer.ItemData(cmb_FlgPer.ListIndex)) = 1 Then
         g_str_Parame = g_str_Parame & cmb_JurPer.ItemData(cmb_JurPer.ListIndex) & ", "
         g_str_Parame = g_str_Parame & "'" & Trim(txt_JurPer.Text) & "', "
      Else
         g_str_Parame = g_str_Parame & 0 & ", "
         g_str_Parame = g_str_Parame & "'" & " " & "', "
      End If
      
      g_str_Parame = g_str_Parame & "'" & Trim(txt_PatPer.Text) & "', "
      g_str_Parame = g_str_Parame & "'" & Trim(txt_MatPer.Text) & "', "
      g_str_Parame = g_str_Parame & "'" & Trim(txt_NomPer.Text) & "', "
      g_str_Parame = g_str_Parame & "'" & cmb_OcuPer.ItemData(cmb_OcuPer.ListIndex) & "', "
      g_str_Parame = g_str_Parame & cmb_ViaPer.ItemData(cmb_ViaPer.ListIndex) & ", "
      g_str_Parame = g_str_Parame & "'" & txt_ViaPer.Text & "', "
      g_str_Parame = g_str_Parame & "'" & txt_NroPer.Text & "', "
      g_str_Parame = g_str_Parame & "'" & txt_intPer.Text & "', "
      g_str_Parame = g_str_Parame & cmb_ZonPer.ItemData(cmb_ZonPer.ListIndex) & ", "
      g_str_Parame = g_str_Parame & "'" & txt_ZonPer.Text & "', "
      
      If Len(Trim(txt_TelPer.Text)) = 0 Then
         g_str_Parame = g_str_Parame & "'" & " " & "', "
      Else
         g_str_Parame = g_str_Parame & "'" & txt_TelPer.Text & "', "
      End If
      g_str_Parame = g_str_Parame & cmb_TipFav.ItemData(cmb_TipFav.ListIndex) & ", "
      g_str_Parame = g_str_Parame & "'" & Trim(txt_DocFav.Text) & "', "
      
      If CStr(cmb_FlgPer.ItemData(cmb_FlgFav.ListIndex)) = 1 Then
         g_str_Parame = g_str_Parame & cmb_JurFav.ItemData(cmb_JurFav.ListIndex) & ", "
         g_str_Parame = g_str_Parame & "'" & Trim(txt_JurFav.Text) & "', "
      Else
         g_str_Parame = g_str_Parame & 0 & ", "
         g_str_Parame = g_str_Parame & "'" & " " & "', "
      End If
      
      g_str_Parame = g_str_Parame & "'" & Trim(txt_PatFav.Text) & "', "
      g_str_Parame = g_str_Parame & "'" & Trim(txt_MatFav.Text) & "', "
      g_str_Parame = g_str_Parame & "'" & Trim(txt_NomFav.Text) & "', "
      g_str_Parame = g_str_Parame & "'" & cmb_OcuFav.ItemData(cmb_OcuFav.ListIndex) & "', "
      g_str_Parame = g_str_Parame & cmb_ViaFav.ItemData(cmb_ViaFav.ListIndex) & ", "
      g_str_Parame = g_str_Parame & "'" & Trim(txt_ViaFav.Text) & "', "
      g_str_Parame = g_str_Parame & "'" & Trim(txt_NroFav.Text) & "', "
      g_str_Parame = g_str_Parame & "'" & Trim(txt_IntFav.Text) & "', "
      g_str_Parame = g_str_Parame & cmb_ZonFav.ItemData(cmb_ZonFav.ListIndex) & ", "
      g_str_Parame = g_str_Parame & "'" & Trim(txt_ZonFav.Text) & "', "
      
      If Len(Trim(txt_TelFav.Text)) = 0 Then
         g_str_Parame = g_str_Parame & "'" & " " & "', "
      Else
         g_str_Parame = g_str_Parame & "'" & Trim(txt_TelFav.Text) & "', "
      End If
      
      g_str_Parame = g_str_Parame & "'" & cmb_ValOpe.ItemData(cmb_ValOpe.ListIndex) & "', "
      g_str_Parame = g_str_Parame & "'" & Format(cmb_TipOpe.ItemData(cmb_TipOpe.ListIndex), "00") & "', "
      g_str_Parame = g_str_Parame & "'" & Trim(txt_OrgOpe.Text) & " " & "', "
      g_str_Parame = g_str_Parame & modtac_g_int_Moneda & ", "
      
      If modtac_g_int_Moneda = 1 Then
         g_str_Parame = g_str_Parame & 0 & ", "
      Else
         g_str_Parame = g_str_Parame & CDbl(ipp_TipCam.Text) & ", "
      End If
      
      g_str_Parame = g_str_Parame & CDbl(ipp_OrgMto.Text) & ", "
      g_str_Parame = g_str_Parame & "'" & l_arr_CtaBan(cmb_CtaBan.ListIndex + 1).Genera_Codigo & "', "
         
      'Datos de Auditoria
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "
      g_str_Parame = g_str_Parame & "'" & l_arr_CodBan(cmb_CodBan.ListIndex + 1).Genera_Codigo & "', "
      g_str_Parame = g_str_Parame & "'" & Format(cmb_TipFon.ItemData(cmb_TipFon.ListIndex), "00") & "', "
      g_str_Parame = g_str_Parame & CStr(moddat_g_int_FlgGrb) & ") "
   
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
         moddat_g_int_CntErr = moddat_g_int_CntErr + 1
      Else
         moddat_g_int_FlgGOK = True
      End If
   
      If moddat_g_int_CntErr = 6 Then
         If MsgBox("No se pudo completar el procedimiento. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
            Exit Sub
         Else
            moddat_g_int_CntErr = 0
         End If
      End If
   Loop
   
   'If moddat_g_int_FlgGrb = 1 Then
   '   Call cmd_Agrega_Click
   'End If
   
   moddat_g_int_FlgAct = 2
   MsgBox "Los datos se grabaron correctamente", vbInformation
   Unload Me
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Inicia
   Call fs_Limpia
   Call fs_Buscar
   
   Call gs_CentraForm(Me)
   Screen.MousePointer = 0
End Sub

Private Sub fs_Limpia()
   msk_HorOpe.Text = ""
   txt_DocFis.Text = ""
   txt_PatFis.Text = ""
   txt_MatFis.Text = ""
   txt_CasFis.Text = ""
   txt_NomFis.Text = ""
   txt_TelFis.Text = ""
   txt_ViaFis.Text = ""
   txt_NroFis.Text = ""
   txt_IntFis.Text = ""
   txt_ZonFis.Text = ""
   txt_RefFis.Text = ""
   
   txt_DocPer.Text = ""
   txt_PatPer.Text = ""
   txt_MatPer.Text = ""
   txt_CasPer.Text = ""
   txt_NomPer.Text = ""
   txt_TelPer.Text = ""
   txt_ViaPer.Text = ""
   txt_NroPer.Text = ""
   txt_intPer.Text = ""
   txt_ZonPer.Text = ""
   txt_RefPer.Text = ""
   txt_JurPer.Text = ""
   
   txt_DocFav.Text = ""
   txt_PatFav.Text = ""
   txt_MatFav.Text = ""
   txt_CasFav.Text = ""
   txt_NomFav.Text = ""
   txt_TelFav.Text = ""
   txt_ViaFav.Text = ""
   txt_NroFav.Text = ""
   txt_IntFav.Text = ""
   txt_ZonFav.Text = ""
   txt_RefFav.Text = ""
   txt_JurFav.Text = ""
   txt_OrgOpe.Text = ""
End Sub

Private Sub fs_Buscar()
   Dim r_int_TipVia As Integer
   Dim r_str_NomVia As String
   Dim r_str_NumVia As String
   Dim r_str_IntDpt As String
   Dim r_int_TipZon As Integer
   Dim r_str_NomZon As String
   Dim r_str_datcli As String
   Dim r_int_Contad As Integer

   'Para leer Equivalencias en Ocupaciones
   ReDim r_arr_Ocupac(0)
      
   If moddat_g_int_FlgGrb = 1 Then
      g_str_Parame = "SELECT * FROM CTB_PRFMSB "
      g_str_Parame = g_str_Parame & "ORDER BY PRFMSB_CODPSB ASC "
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
         Exit Sub
      End If
      
      If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
         g_rst_Listas.MoveFirst
         Do While Not g_rst_Listas.EOF
             
            ReDim Preserve r_arr_Ocupac(UBound(r_arr_Ocupac) + 1)
            r_arr_Ocupac(UBound(r_arr_Ocupac)).Ocupac_CodPmc = Trim(g_rst_Listas!PrfMsb_CodPmc)
            r_arr_Ocupac(UBound(r_arr_Ocupac)).Ocupac_CodPsb = Trim(g_rst_Listas!PrfMsb_CodPsb)
            r_arr_Ocupac(UBound(r_arr_Ocupac)).Ocupac_Descri = Trim(g_rst_Listas!PrfMsb_Descri)
            g_rst_Listas.MoveNext
         Loop
      End If
      
      g_rst_Listas.Close
      Set g_rst_Listas = Nothing
      
      pnl_Sucurs.Caption = moddat_g_str_Codigo
      pnl_Empres.Caption = "MI CASITA HIPOTECARIA"
   
      'Buscando Información del Crédito
      g_str_Parame = "SELECT * FROM CRE_HIPMAE, CLI_DATGEN WHERE "
      g_str_Parame = g_str_Parame & "HIPMAE_NUMOPE = '" & moddat_g_str_NumOpe & "' AND "
      g_str_Parame = g_str_Parame & "HIPMAE_TDOCLI = DATGEN_TIPDOC AND "
      g_str_Parame = g_str_Parame & "HIPMAE_NDOCLI = DATGEN_NUMDOC AND "
      g_str_Parame = g_str_Parame & "(HIPMAE_SITUAC = 2 OR HIPMAE_SITUAC = 9)"
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         Exit Sub
      End If
      
      If g_rst_Princi.BOF And g_rst_Princi.EOF Then
         g_rst_Princi.Close
         Set g_rst_Princi = Nothing
         Exit Sub
      End If
      
      pnl_Client.Caption = g_rst_Princi!DATGEN_TIPDOC & " - " & Trim(g_rst_Princi!DATGEN_NUMDOC) & " / " & Trim(g_rst_Princi!DatGen_ApePat & "") & " " & Trim(g_rst_Princi!DATGEN_APECAS & "") & " " & Trim(g_rst_Princi!DatGen_ApeMat & "") & " " & Trim(g_rst_Princi!DatGen_Nombre & "")
      pnl_Period.Caption = modtac_g_int_PerAno & " - " & Format(modtac_g_int_PerMes, "00")
      r_str_datcli = ff_DatCli(g_rst_Princi!HIPMAE_NUMSOL, r_int_TipVia, r_str_NomVia, r_str_NumVia, r_str_IntDpt, r_int_TipZon, r_str_NomZon)
      
      'Informacion de la persona que fisicamente realiza la operacion
      Call gs_BuscarCombo_Item(cmb_TipFis, g_rst_Princi!DATGEN_TIPDOC)
      
      txt_DocFis.Text = Trim(g_rst_Princi!DATGEN_NUMDOC & "")
      txt_PatFis.Text = Trim(g_rst_Princi!DatGen_ApePat & "")
      txt_MatFis.Text = Trim(g_rst_Princi!DatGen_ApeMat & "")
      txt_CasFis.Text = Trim(g_rst_Princi!DATGEN_APECAS & "")
      txt_NomFis.Text = Trim(g_rst_Princi!DatGen_Nombre & "")
      
      If IsNull(g_rst_Princi!DATGEN_TELEFO) Then
         If IsNull(g_rst_Princi!DATGEN_NUMCEL) Then
            txt_TelFis.Text = ""
         Else
            txt_TelFis.Text = Trim(g_rst_Princi!DATGEN_NUMCEL & "")
         End If
      Else
         txt_TelFis.Text = Trim(g_rst_Princi!DATGEN_TELEFO & "")
      End If
      
      Call gs_BuscarCombo_Item(cmb_ViaFis, r_int_TipVia)
      txt_ViaFis.Text = Trim(r_str_NomVia & "")
      txt_NroFis.Text = Trim(r_str_NumVia & "")
      txt_IntFis.Text = Trim(r_str_IntDpt & "")
      Call gs_BuscarCombo_Item(cmb_ZonFis, r_int_TipZon)
      txt_ZonFis.Text = Trim(r_str_NomZon & "")
      
      Call gs_BuscarCombo_Item(cmb_DptFis, CInt(Left(g_rst_Princi!DatGen_Ubigeo, 2)))
      Call moddat_gs_Carga_Provin(cmb_PrvFis, Left(g_rst_Princi!DatGen_Ubigeo, 2))
      Call gs_BuscarCombo_Item(cmb_PrvFis, CInt(Mid(g_rst_Princi!DatGen_Ubigeo, 3, 2)))
      Call moddat_gs_Carga_Distri(cmb_DstFis, Left(g_rst_Princi!DatGen_Ubigeo, 2), Mid(g_rst_Princi!DatGen_Ubigeo, 3, 2))
      Call gs_BuscarCombo_Item(cmb_DstFis, CInt(Right(g_rst_Princi!DatGen_Ubigeo, 2)))
      txt_RefFis.Text = Trim(g_rst_Princi!DATGEN_REFERE & "")
      
      'Informacion de la persona en cuyo nombre se realiza la operación
      Call gs_BuscarCombo_Item(cmb_TipPer, g_rst_Princi!DATGEN_TIPDOC)
      txt_DocPer.Text = Trim(g_rst_Princi!DATGEN_NUMDOC & "")
      Call gs_BuscarCombo_Item(cmb_FlgPer, 2)
      txt_PatPer.Text = Trim(g_rst_Princi!DatGen_ApePat & "")
      txt_MatPer.Text = Trim(g_rst_Princi!DatGen_ApeMat & "")
      txt_CasPer.Text = Trim(g_rst_Princi!DATGEN_APECAS & "")
      txt_NomPer.Text = Trim(g_rst_Princi!DatGen_Nombre & "")
      If IsNull(g_rst_Princi!DATGEN_TELEFO) Then
         If IsNull(g_rst_Princi!DATGEN_NUMCEL) Then
            txt_TelPer.Text = ""
         Else
            txt_TelPer.Text = Trim(g_rst_Princi!DATGEN_NUMCEL & "")
         End If
      Else
         txt_TelPer.Text = Trim(g_rst_Princi!DATGEN_TELEFO & "")
      End If
      
      Call gs_BuscarCombo_Item(cmb_ViaPer, r_int_TipVia)
      txt_ViaPer.Text = Trim(r_str_NomVia & "")
      txt_NroPer.Text = Trim(r_str_NumVia & "")
      txt_intPer.Text = Trim(r_str_IntDpt & "")
      Call gs_BuscarCombo_Item(cmb_ZonPer, r_int_TipZon)
      txt_ZonPer.Text = Trim(r_str_NomZon & "")
      
      Call gs_BuscarCombo_Item(cmb_DptPer, CInt(Left(g_rst_Princi!DatGen_Ubigeo, 2)))
      Call moddat_gs_Carga_Provin(cmb_PrvPer, Left(g_rst_Princi!DatGen_Ubigeo, 2))
      Call gs_BuscarCombo_Item(cmb_PrvPer, CInt(Mid(g_rst_Princi!DatGen_Ubigeo, 3, 2)))
      Call moddat_gs_Carga_Distri(cmb_DstPer, Left(g_rst_Princi!DatGen_Ubigeo, 2), Mid(g_rst_Princi!DatGen_Ubigeo, 3, 2))
      Call gs_BuscarCombo_Item(cmb_DstPer, CInt(Right(g_rst_Princi!DatGen_Ubigeo, 2)))
      txt_RefPer.Text = Trim(g_rst_Princi!DATGEN_REFERE & "")
      
      'Informacion de la persona a favor de quien se realiza la operación
      Call gs_BuscarCombo_Item(cmb_TipFav, g_rst_Princi!DATGEN_TIPDOC)
      txt_DocFav.Text = Trim(g_rst_Princi!DATGEN_NUMDOC & "")
      Call gs_BuscarCombo_Item(cmb_FlgFav, 2)
      
      txt_PatFav.Text = Trim(g_rst_Princi!DatGen_ApePat & "")
      txt_MatFav.Text = Trim(g_rst_Princi!DatGen_ApeMat & "")
      txt_CasFav.Text = Trim(g_rst_Princi!DATGEN_APECAS & "")
      txt_NomFav.Text = Trim(g_rst_Princi!DatGen_Nombre & "")
      If IsNull(g_rst_Princi!DATGEN_TELEFO) Then
         If IsNull(g_rst_Princi!DATGEN_NUMCEL) Then
            txt_TelFav.Text = ""
         Else
            txt_TelFav.Text = Trim(g_rst_Princi!DATGEN_NUMCEL & "")
         End If
      Else
         txt_TelFav.Text = Trim(g_rst_Princi!DATGEN_TELEFO & "")
      End If
      
      Call gs_BuscarCombo_Item(cmb_ViaFav, r_int_TipVia)
      txt_ViaFav.Text = Trim(r_str_NomVia & "")
      txt_NroFav.Text = Trim(r_str_NumVia & "")
      txt_IntFav.Text = Trim(r_str_IntDpt & "")
      Call gs_BuscarCombo_Item(cmb_ZonFav, r_int_TipZon)
      txt_ZonFav.Text = Trim(r_str_NomZon & "")
      
      Call gs_BuscarCombo_Item(cmb_DptFav, CInt(Left(g_rst_Princi!DatGen_Ubigeo, 2)))
      Call moddat_gs_Carga_Provin(cmb_PrvFav, Left(g_rst_Princi!DatGen_Ubigeo, 2))
      Call gs_BuscarCombo_Item(cmb_PrvFav, CInt(Mid(g_rst_Princi!DatGen_Ubigeo, 3, 2)))
      Call moddat_gs_Carga_Distri(cmb_DstFav, Left(g_rst_Princi!DatGen_Ubigeo, 2), Mid(g_rst_Princi!DatGen_Ubigeo, 3, 2))
      Call gs_BuscarCombo_Item(cmb_DstFav, CInt(Right(g_rst_Princi!DatGen_Ubigeo, 2)))
      txt_RefFav.Text = Trim(g_rst_Princi!DATGEN_REFERE & "")
      
      For r_int_Contad = 1 To UBound(r_arr_Ocupac)
         If Format(g_rst_Princi!Datgen_Profes, "000000") = r_arr_Ocupac(r_int_Contad).Ocupac_CodPmc Then
             Call gs_BuscarCombo_Item(cmb_OcuFis, CInt(r_arr_Ocupac(r_int_Contad).Ocupac_CodPsb))
             Call gs_BuscarCombo_Item(cmb_OcuPer, CInt(r_arr_Ocupac(r_int_Contad).Ocupac_CodPsb))
             Call gs_BuscarCombo_Item(cmb_OcuFav, CInt(r_arr_Ocupac(r_int_Contad).Ocupac_CodPsb))
         End If
      Next r_int_Contad
      
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
       
   ElseIf moddat_g_int_FlgGrb = 2 Then
   
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "SELECT * "
      g_str_Parame = g_str_Parame & "  FROM OPE_REGLAV "
      g_str_Parame = g_str_Parame & " WHERE REGLAV_PERMES = " & modtac_g_int_PerMes & " "
      g_str_Parame = g_str_Parame & "   AND REGLAV_PERANO = " & modtac_g_int_PerAno & " "
      g_str_Parame = g_str_Parame & "   AND REGLAV_NUMOPE = '" & moddat_g_str_NumOpe & "' "
      g_str_Parame = g_str_Parame & "   AND REGLAV_NROINT = '" & modtac_g_str_NroInt & "' "
      g_str_Parame = g_str_Parame & "ORDER BY REGLAV_NUMOPE ASC "
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         Exit Sub
      End If
      
      If g_rst_Princi.BOF And g_rst_Princi.EOF Then
         g_rst_Princi.Close
         Set g_rst_Princi = Nothing
         Exit Sub
      End If
   
      pnl_Sucurs.Caption = moddat_g_str_Codigo
      pnl_Empres.Caption = "MI CASITA HIPOTECARIA"
      pnl_Client.Caption = g_rst_Princi!REGLAV_TCLIFA & " - " & Trim(g_rst_Princi!REGLAV_NCLIFA) & " / " & Trim(g_rst_Princi!REGLAV_PATFAV & "") & " " & Trim(g_rst_Princi!REGLAV_MATFAV & "") & " " & Trim(g_rst_Princi!REGLAV_NOMFAV & "")
      pnl_Period.Caption = modtac_g_int_PerAno & " - " & Format(modtac_g_int_PerMes, "00")
            
      'Informacion de la persona que fisicamente realiza la operacion
      Call gs_BuscarCombo_Item(cmb_TipFis, g_rst_Princi!REGLAV_TCLIFI)
      txt_DocFis.Text = Trim(g_rst_Princi!REGLAV_NCLIFI & "")
      txt_PatFis.Text = Trim(g_rst_Princi!REGLAV_PATFIS & "")
      txt_MatFis.Text = Trim(g_rst_Princi!REGLAV_MATFIS & "")
      'txt_CasFis.Text = Trim(g_rst_Princi!REGLAV_APECAS & "")
      txt_NomFis.Text = Trim(g_rst_Princi!REGLAV_NOMFIS & "")
      txt_TelFis.Text = Trim(g_rst_Princi!REGLAV_TELFIS & "")
      
      Call gs_BuscarCombo_Item(cmb_ViaFis, g_rst_Princi!REGLAV_TIPVIA_FIS)
      txt_ViaFis.Text = Trim(g_rst_Princi!REGLAV_NOMVIA_FIS & "")
      txt_NroFis.Text = Trim(g_rst_Princi!REGLAV_NUMERO_FIS & "")
      txt_IntFis.Text = Trim(g_rst_Princi!REGLAV_INTDPT_FIS & "")
      Call gs_BuscarCombo_Item(cmb_ZonFis, g_rst_Princi!REGLAV_TIPZON_FIS)
      txt_ZonFis.Text = Trim(g_rst_Princi!REGLAV_NOMZON_FIS & "")
   
      Call gs_BuscarCombo_Item(cmb_DptFis, CInt(Left(g_rst_Princi!REGLAV_UBIGEO, 2)))
      Call moddat_gs_Carga_Provin(cmb_PrvFis, Left(g_rst_Princi!REGLAV_UBIGEO, 2))
      Call gs_BuscarCombo_Item(cmb_PrvFis, CInt(Mid(g_rst_Princi!REGLAV_UBIGEO, 3, 2)))
      Call moddat_gs_Carga_Distri(cmb_DstFis, Left(g_rst_Princi!REGLAV_UBIGEO, 2), Mid(g_rst_Princi!REGLAV_UBIGEO, 3, 2))
      Call gs_BuscarCombo_Item(cmb_DstFis, CInt(Right(g_rst_Princi!REGLAV_UBIGEO, 2)))
            
      'txt_RefFis.Text = Trim(g_rst_Princi!DATGEN_REFERE & "")
         
      'Informacion de la persona en cuyo nombre se realiza la operación
      Call gs_BuscarCombo_Item(cmb_TipPer, g_rst_Princi!REGLAV_TCLIPE)
      txt_DocPer.Text = Trim(g_rst_Princi!REGLAV_NCLIPE & "")
      Call gs_BuscarCombo_Item(cmb_FlgPer, 2)
      
      txt_PatPer.Text = Trim(g_rst_Princi!REGLAV_PATPER & "")
      txt_MatPer.Text = Trim(g_rst_Princi!REGLAV_MATPER & "")
      'txt_CasPer.Text = Trim(g_rst_Princi!DATGEN_APECAS & "")
      txt_NomPer.Text = Trim(g_rst_Princi!REGLAV_NOMPER & "")
      txt_TelPer.Text = Trim(g_rst_Princi!REGLAV_TELPER & "")
      
      Call gs_BuscarCombo_Item(cmb_ViaPer, g_rst_Princi!REGLAV_TIPVIA_PER)
      txt_ViaPer.Text = Trim(g_rst_Princi!REGLAV_NOMVIA_PER & "")
      txt_NroPer.Text = Trim(g_rst_Princi!REGLAV_NUMERO_PER & "")
      txt_intPer.Text = Trim(g_rst_Princi!REGLAV_INTDPT_PER & "")
      Call gs_BuscarCombo_Item(cmb_ZonPer, g_rst_Princi!REGLAV_TIPZON_PER)
      txt_ZonPer.Text = Trim(g_rst_Princi!REGLAV_NOMZON_PER & "")
      
      Call gs_BuscarCombo_Item(cmb_DptPer, CInt(Left(g_rst_Princi!REGLAV_UBIGEO, 2)))
      Call moddat_gs_Carga_Provin(cmb_PrvPer, Left(g_rst_Princi!REGLAV_UBIGEO, 2))
      Call gs_BuscarCombo_Item(cmb_PrvPer, CInt(Mid(g_rst_Princi!REGLAV_UBIGEO, 3, 2)))
      Call moddat_gs_Carga_Distri(cmb_DstPer, Left(g_rst_Princi!REGLAV_UBIGEO, 2), Mid(g_rst_Princi!REGLAV_UBIGEO, 3, 2))
      Call gs_BuscarCombo_Item(cmb_DstPer, CInt(Right(g_rst_Princi!REGLAV_UBIGEO, 2)))
            
      'txt_RefPer.Text = Trim(g_rst_Princi!DATGEN_REFERE & "")
      
      'Informacion de la persona a favor de quien se realiza la operación
      Call gs_BuscarCombo_Item(cmb_TipFav, g_rst_Princi!REGLAV_TCLIFA)
      txt_DocFav.Text = Trim(g_rst_Princi!REGLAV_NCLIFA & "")
      Call gs_BuscarCombo_Item(cmb_FlgFav, 2)
      
      txt_PatFav.Text = Trim(g_rst_Princi!REGLAV_PATFAV & "")
      txt_MatFav.Text = Trim(g_rst_Princi!REGLAV_MATFAV & "")
      'txt_CasFav.Text = Trim(g_rst_Princi!DATGEN_APECAS & "")
      txt_NomFav.Text = Trim(g_rst_Princi!REGLAV_NOMFAV & "")
      txt_TelFav.Text = Trim(g_rst_Princi!REGLAV_TELFAV & "")
      
      Call gs_BuscarCombo_Item(cmb_ViaFav, g_rst_Princi!REGLAV_TIPVIA_FAV)
      txt_ViaFav.Text = Trim(g_rst_Princi!REGLAV_NOMVIA_FAV & "")
      txt_NroFav.Text = Trim(g_rst_Princi!REGLAV_NUMERO_FAV & "")
      txt_IntFav.Text = Trim(g_rst_Princi!REGLAV_INTDPT_FAV & "")
      Call gs_BuscarCombo_Item(cmb_ZonFav, g_rst_Princi!REGLAV_TIPZON_FAV)
      txt_ZonFav.Text = Trim(g_rst_Princi!REGLAV_NOMZON_FAV & "")
   
      Call gs_BuscarCombo_Item(cmb_DptFav, CInt(Left(g_rst_Princi!REGLAV_UBIGEO, 2)))
      Call moddat_gs_Carga_Provin(cmb_PrvFav, Left(g_rst_Princi!REGLAV_UBIGEO, 2))
      Call gs_BuscarCombo_Item(cmb_PrvFav, CInt(Mid(g_rst_Princi!REGLAV_UBIGEO, 3, 2)))
      Call moddat_gs_Carga_Distri(cmb_DstFav, Left(g_rst_Princi!REGLAV_UBIGEO, 2), Mid(g_rst_Princi!REGLAV_UBIGEO, 3, 2))
      Call gs_BuscarCombo_Item(cmb_DstFav, CInt(Right(g_rst_Princi!REGLAV_UBIGEO, 2)))
      'txt_RefFav.Text = Trim(g_rst_Princi!DATGEN_REFERE & "")
      
      Call gs_BuscarCombo_Item(cmb_ValOpe, g_rst_Princi!REGLAV_TIPEFE)
      Call gs_BuscarCombo_Item(cmb_TipOpe, g_rst_Princi!REGLAV_TIPOPE)
      If Not Trim(g_rst_Princi!REGLAV_TIPFON) Then
         Call gs_BuscarCombo_Item(cmb_TipFon, g_rst_Princi!REGLAV_TIPFON)
      End If
      txt_OrgOpe.Text = Trim(g_rst_Princi!REGLAV_DESORG & "")
      ipp_OrgMto.Value = g_rst_Princi!REGLAV_IMPOPE
      ipp_TipCam.Value = g_rst_Princi!REGLAV_TIPCAM
      
      If Trim(g_rst_Princi!REGLAV_MODOPE) = "UNICA" Then
         Call gs_BuscarCombo_Item(cmb_MonOpe, 1)
      Else
         Call gs_BuscarCombo_Item(cmb_MonOpe, 2)
      End If
      
      msk_HorOpe.Text = Format(g_rst_Princi!REGLAV_HOROPE, "000000")
      ipp_FecCtb.Text = CDate(gf_FormatoFecha(CStr(g_rst_Princi!REGLAV_FECOPE)))
      
      Call gs_BuscarCombo_Item(cmb_OcuFis, g_rst_Princi!REGLAV_OCUFIS)
      Call gs_BuscarCombo_Item(cmb_OcuPer, g_rst_Princi!REGLAV_OCUFIS)
      Call gs_BuscarCombo_Item(cmb_OcuFav, g_rst_Princi!REGLAV_OCUFIS)
      
      'Call moddat_gs_Carga_LisIte(cmb_CodBan, l_arr_CodBan, 1, "516")
      'Call gs_BuscarCombo_Item(cmb_CodBan, CInt(g_rst_Princi!REGLAV_CODBAN))
      
      For r_int_Contad = 1 To UBound(l_arr_CodBan)
         If Format(l_arr_CodBan(r_int_Contad).Genera_Codigo, "000000") = Format(g_rst_Princi!REGLAV_CODBAN, "000000") Then
             cmb_CodBan.ListIndex = r_int_Contad - 1
             Exit For
         End If
      Next r_int_Contad
      
      Call moddat_gs_Carga_CtaBan(l_arr_CodBan(cmb_CodBan.ListIndex + 1).Genera_Codigo, cmb_CtaBan, l_arr_CtaBan)
      
      For r_int_Contad = 1 To UBound(l_arr_CtaBan)
         If Trim(l_arr_CtaBan(r_int_Contad).Genera_Codigo) = Trim(g_rst_Princi!REGLAV_CTABAN) Then
             cmb_CtaBan.ListIndex = r_int_Contad - 1
             Exit For
         End If
      Next r_int_Contad
      
   End If
End Sub

Private Function ff_Genera_NumReg() As String
   ff_Genera_NumReg = ""
   g_str_Parame = "SELECT MAX(REGLAV_NROINT) AS DATGEN FROM OPE_REGLAV"
  
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
      Exit Function
   End If
  
   If IsNull(g_rst_Genera!DATGEN) Then
      ff_Genera_NumReg = Format(Now, "yyyy") & Mid(g_rst_Genera!DATGEN, 5, 3) & Format(1, "000")
   ElseIf Format(Now, "yyyy") = Left(g_rst_Genera!DATGEN, 4) Then
      ff_Genera_NumReg = Left(g_rst_Genera!DATGEN, 4) & Mid(g_rst_Genera!DATGEN, 5, 3) & Format(Right(g_rst_Genera!DATGEN, 3) + 1, "000")
   ElseIf Format(Now, "yyyy") <> Left(g_rst_Genera!DATGEN, 4) Then
      ff_Genera_NumReg = Format(Now, "yyyy") & Mid(g_rst_Genera!DATGEN, 5, 3) & Format(1, "000")
   End If
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
End Function

Private Sub fs_Inicia()
   cmb_ValOpe.Clear
   cmb_MonOpe.Clear
   cmb_MonOpe.AddItem "UNICA"
   cmb_MonOpe.ItemData(cmb_MonOpe.NewIndex) = 1
   cmb_MonOpe.AddItem "MULTIPLE"
   cmb_MonOpe.ItemData(cmb_MonOpe.NewIndex) = 2
   cmb_ValOpe.AddItem "OPERACION EN EFECTIVO"
   cmb_ValOpe.ItemData(cmb_ValOpe.NewIndex) = 1
   cmb_ValOpe.AddItem "OTRO TIPO DE OPER. QUE NO ES EFEC."
   cmb_ValOpe.ItemData(cmb_ValOpe.NewIndex) = 2
   
   'msk_HorOpe.Text = "000000"
      
   Call moddat_gs_Carga_LisIte_Combo(cmb_TipFis, 1, "203")
   Call moddat_gs_Carga_LisIte_Combo(cmb_ViaFis, 1, "201")
   Call moddat_gs_Carga_LisIte_Combo(cmb_ZonFis, 1, "202")
   Call modtac_gs_Carga_ProOcu(cmb_OcuFis, l_arr_OcuFis)
   Call moddat_gs_Carga_Depart(cmb_DptFis)
   
   Call moddat_gs_Carga_LisIte_Combo(cmb_FlgPer, 1, "214")
   Call moddat_gs_Carga_LisIte_Combo(cmb_TipPer, 1, "203")
   Call moddat_gs_Carga_LisIte_Combo(cmb_JurPer, 1, "232")
   Call moddat_gs_Carga_LisIte_Combo(cmb_ViaPer, 1, "201")
   Call moddat_gs_Carga_LisIte_Combo(cmb_ZonPer, 1, "202")
   Call modtac_gs_Carga_ProOcu(cmb_OcuPer, l_arr_OcuPer)
   Call moddat_gs_Carga_Depart(cmb_DptPer)
   
   Call moddat_gs_Carga_LisIte_Combo(cmb_FlgFav, 1, "214")
   Call moddat_gs_Carga_LisIte_Combo(cmb_TipFav, 1, "203")
   Call moddat_gs_Carga_LisIte_Combo(cmb_JurFav, 1, "232")
   Call moddat_gs_Carga_LisIte_Combo(cmb_ViaFav, 1, "201")
   Call moddat_gs_Carga_LisIte_Combo(cmb_ZonFav, 1, "202")
   Call modtac_gs_Carga_ProOcu(cmb_OcuFav, l_arr_OcuFav)
   Call moddat_gs_Carga_Depart(cmb_DptFav)
   Call moddat_gs_Carga_LisIte_Combo(cmb_TipOpe, 1, "070")
   Call moddat_gs_Carga_LisIte(cmb_CodBan, l_arr_CodBan, 1, "516")
   Call moddat_gs_Carga_LisIte_Combo(cmb_TipFon, 1, "072")
   
   If modtac_g_int_Moneda = 1 Then
      pnl_TipMon.Caption = "(S/.)"
   ElseIf modtac_g_int_Moneda = 2 Then
      pnl_TipMon.Caption = "(US$.)"
   End If
   
   ' se le pasa el combobox y el formulario como parámetro para ajustar el texto
   Establecer_Ancho cmb_ValOpe, Me
   Establecer_Ancho cmb_TipOpe, Me
   Establecer_Ancho cmb_OcuFis, Me
   Establecer_Ancho cmb_OcuPer, Me
   Establecer_Ancho cmb_OcuFav, Me
   Establecer_Ancho cmb_TipFon, Me
End Sub

Private Sub cmb_DstPer_Click()
   If cmb_DstPer.ListIndex > -1 Then
      If l_int_FlgCmb Then
         Call gs_SetFocus(txt_RefPer)
      End If
   End If
End Sub

Private Sub cmb_PrvPer_Click()
   If cmb_PrvPer.ListIndex > -1 Then
      If l_int_FlgCmb Then
         cmb_DstPer.Clear
         Screen.MousePointer = 11
         Call moddat_gs_Carga_Distri(cmb_DstPer, Format(cmb_DptPer.ItemData(cmb_DptPer.ListIndex), "00"), Format(cmb_PrvPer.ItemData(cmb_PrvPer.ListIndex), "00"))
         Screen.MousePointer = 0
         Call gs_SetFocus(cmb_DstPer)
      End If
   End If
End Sub

Private Sub cmb_FlgPer_Click()
   If cmb_FlgPer.ListIndex = -1 Then
      cmb_JurPer.ListIndex = -1
      txt_JurPer.Text = ""
      cmb_FlgPer.Enabled = False
      txt_DocPer.Enabled = False
   Else
      If cmb_FlgPer.ItemData(cmb_FlgPer.ListIndex) = 1 Then
         cmb_JurPer.Enabled = True
         txt_JurPer.Enabled = True
         Call gs_SetFocus(cmb_JurPer)
      Else
         cmb_JurPer.ListIndex = -1
         txt_JurPer.Text = ""
         cmb_JurPer.Enabled = False
         txt_JurPer.Enabled = False
         Call gs_SetFocus(txt_PatPer)
      End If
   End If
End Sub

Private Sub cmb_FlgPer_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_FlgPer_Click
   End If
End Sub

Private Sub cmb_DptPer_Change()
   l_str_DptPer = cmb_DptPer.Text
End Sub

Private Sub cmb_DptPer_Click()
   If cmb_DptPer.ListIndex > -1 Then
      If l_int_FlgCmb Then
         cmb_PrvPer.Clear
         cmb_DstPer.Clear
         Screen.MousePointer = 11
         Call moddat_gs_Carga_Provin(cmb_PrvPer, Format(cmb_DptPer.ItemData(cmb_DptPer.ListIndex), "00"))
         Screen.MousePointer = 0
         Call gs_SetFocus(cmb_PrvPer)
      End If
   End If
End Sub

Private Sub cmb_DptPer_GotFocus()
   l_int_FlgCmb = True
   l_str_DptPer = cmb_DptPer.Text
End Sub

Private Sub cmb_DptPer_KeyPress(KeyAscii As Integer)
   
   If KeyAscii <> 13 Then
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS + "-_ ./*+#,()" + Chr(34))
   Else
      l_int_FlgCmb = False
      Call gs_BuscarCombo(cmb_DptPer, l_str_DptPer)
      l_int_FlgCmb = True
      
      cmb_PrvPer.Clear
      cmb_DstPer.Clear
      If cmb_DptPer.ListIndex > -1 Then
         l_str_DptPer = ""
         Screen.MousePointer = 11
         Call moddat_gs_Carga_Provin(cmb_PrvPer, Format(cmb_DptPer.ItemData(cmb_DptPer.ListIndex), "00"))
         Screen.MousePointer = 0
      End If
      
      Call gs_SetFocus(cmb_PrvPer)
   End If
End Sub

Private Sub cmb_DstPer_Change()
   l_str_DstPer = cmb_DstPer.Text
End Sub

Private Sub cmb_DstPer_GotFocus()
   l_int_FlgCmb = True
   l_str_DstPer = cmb_DstPer.Text
End Sub

Private Sub cmb_DstPer_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS + "-_ ./*+#,()" + Chr(34))
   Else
      l_int_FlgCmb = False
      Call gs_BuscarCombo(cmb_DstPer, l_str_DstPer)
      l_int_FlgCmb = True
      
      If cmb_DstPer.ListIndex > -1 Then
         l_str_DstPer = ""
      End If
      
      Call gs_SetFocus(txt_RefPer)
   End If
End Sub

Private Sub cmb_DstFis_Click()
   If cmb_DstFis.ListIndex > -1 Then
      If l_int_FlgCmb Then
         Call gs_SetFocus(txt_RefFis)
      End If
   End If
End Sub

Private Sub cmb_PrvFis_Click()
If cmb_PrvFis.ListIndex > -1 Then
      If l_int_FlgCmb Then
         cmb_DstFis.Clear
         Screen.MousePointer = 11
         Call moddat_gs_Carga_Distri(cmb_DstFis, Format(cmb_DptFis.ItemData(cmb_DptFis.ListIndex), "00"), Format(cmb_PrvFis.ItemData(cmb_PrvFis.ListIndex), "00"))
         Screen.MousePointer = 0
         Call gs_SetFocus(cmb_DstFis)
      End If
   End If
End Sub

Private Sub cmb_DptFis_Change()
   l_str_DptFis = cmb_DptFis.Text
End Sub

Private Sub cmb_DptFis_Click()
   If cmb_DptFis.ListIndex > -1 Then
      If l_int_FlgCmb Then
         cmb_PrvFis.Clear
         cmb_DstFis.Clear
         Screen.MousePointer = 11
         Call moddat_gs_Carga_Provin(cmb_PrvFis, Format(cmb_DptFis.ItemData(cmb_DptFis.ListIndex), "00"))
         Screen.MousePointer = 0
         Call gs_SetFocus(cmb_PrvFis)
      End If
   End If
End Sub

Private Sub cmb_DptFis_GotFocus()
   l_int_FlgCmb = True
   l_str_DptFis = cmb_DptFis.Text
End Sub

Private Sub cmb_DptFis_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS + "-_ ./*+#,()" + Chr(34))
   Else
      l_int_FlgCmb = False
      Call gs_BuscarCombo(cmb_DptFis, l_str_DptFis)
      l_int_FlgCmb = True
      
      cmb_PrvFis.Clear
      cmb_DstFis.Clear
      If cmb_DptFis.ListIndex > -1 Then
         l_str_DptFis = ""
         
         Screen.MousePointer = 11
         Call moddat_gs_Carga_Provin(cmb_PrvFis, Format(cmb_DptFis.ItemData(cmb_DptFis.ListIndex), "00"))
         Screen.MousePointer = 0
      End If
      
      Call gs_SetFocus(cmb_PrvFis)
   End If
End Sub

Private Sub cmb_DstFis_Change()
   l_str_DstFis = cmb_DstFis.Text
End Sub

Private Sub cmb_DstFis_GotFocus()
   l_int_FlgCmb = True
   l_str_DstFis = cmb_DstFis.Text
End Sub

Private Sub cmb_DstFis_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS + "-_ ./*+#,()" + Chr(34))
   Else
      l_int_FlgCmb = False
      Call gs_BuscarCombo(cmb_DstFis, l_str_DstFis)
      l_int_FlgCmb = True
      
      If cmb_DstFis.ListIndex > -1 Then
         l_str_DstFis = ""
      End If
      
      Call gs_SetFocus(txt_RefFis)
   End If
End Sub

Private Sub cmb_DstFav_Click()
   If cmb_DstFav.ListIndex > -1 Then
      If l_int_FlgCmb Then
         Call gs_SetFocus(txt_RefFav)
      End If
   End If
End Sub

Private Sub cmb_PrvFav_Click()
If cmb_PrvFav.ListIndex > -1 Then
      If l_int_FlgCmb Then
         cmb_DstFav.Clear
         Screen.MousePointer = 11
         Call moddat_gs_Carga_Distri(cmb_DstFav, Format(cmb_DptFav.ItemData(cmb_DptFav.ListIndex), "00"), Format(cmb_PrvFav.ItemData(cmb_PrvFav.ListIndex), "00"))
         Screen.MousePointer = 0
         Call gs_SetFocus(cmb_DstFav)
      End If
   End If
End Sub

Private Sub cmb_FlgFav_Click()
   If cmb_FlgFav.ListIndex = -1 Then
      cmb_JurFav.ListIndex = -1
      txt_JurFav.Text = ""
      cmb_FlgFav.Enabled = False
      txt_DocFav.Enabled = False
   Else
      If cmb_FlgFav.ItemData(cmb_FlgFav.ListIndex) = 1 Then
         cmb_JurFav.Enabled = True
         txt_JurFav.Enabled = True
         Call gs_SetFocus(cmb_JurFav)
      Else
         cmb_JurFav.ListIndex = -1
         txt_JurFav.Text = ""
         cmb_JurFav.Enabled = False
         txt_JurFav.Enabled = False
         Call gs_SetFocus(txt_PatFav)
      End If
   End If
End Sub

Private Sub cmb_FlgFav_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_FlgFav_Click
   End If
End Sub

Private Sub cmb_DptFav_Change()
   l_str_DptFav = cmb_DptFav.Text
End Sub

Private Sub cmb_DptFav_Click()
   If cmb_DptFav.ListIndex > -1 Then
      If l_int_FlgCmb Then
         cmb_PrvFav.Clear
         cmb_DstFav.Clear
         
         Screen.MousePointer = 11
         Call moddat_gs_Carga_Provin(cmb_PrvFav, Format(cmb_DptFav.ItemData(cmb_DptFav.ListIndex), "00"))
         Screen.MousePointer = 0
         
         Call gs_SetFocus(cmb_PrvFav)
      End If
   End If
End Sub

Private Sub cmb_DptFav_GotFocus()
   l_int_FlgCmb = True
   l_str_DptFav = cmb_DptFav.Text
End Sub

Private Sub cmb_DptFav_KeyPress(KeyAscii As Integer)
   
   If KeyAscii <> 13 Then
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS + "-_ ./*+#,()" + Chr(34))
   Else
      l_int_FlgCmb = False
      Call gs_BuscarCombo(cmb_DptFav, l_str_DptFav)
      l_int_FlgCmb = True
      
      cmb_PrvFav.Clear
      cmb_DstFav.Clear
      If cmb_DptFav.ListIndex > -1 Then
         l_str_DptFav = ""
         
         Screen.MousePointer = 11
         Call moddat_gs_Carga_Provin(cmb_PrvFav, Format(cmb_DptFav.ItemData(cmb_DptFav.ListIndex), "00"))
         Screen.MousePointer = 0
      End If
      
      Call gs_SetFocus(cmb_PrvFav)
   End If
End Sub

Private Sub cmb_DstFav_Change()
   l_str_DstFav = cmb_DstFav.Text
End Sub

Private Sub cmb_DstFav_GotFocus()
   l_int_FlgCmb = True
   l_str_DstFav = cmb_DstFav.Text
End Sub

Private Sub cmb_DstFav_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS + "-_ ./*+#,()" + Chr(34))
   Else
      l_int_FlgCmb = False
      Call gs_BuscarCombo(cmb_DstFav, l_str_DstFav)
      l_int_FlgCmb = True
      
      If cmb_DstPer.ListIndex > -1 Then
         l_str_DstFav = ""
      End If
      
      Call gs_SetFocus(txt_RefFav)
   End If
End Sub


'Metodo para obtener Datos del Inmueble del Cliente Reportante
Private Function ff_DatCli(ByVal p_NroSol As String, Optional ByRef p_TipVia As Integer, Optional ByRef p_NomVia As String, Optional ByRef p_NumVia As String, Optional ByRef p_IntDpt As String, Optional ByRef p_TipZon As Integer, Optional ByRef p_NomZon As String) As String
                     
   g_str_Parame = "SELECT * FROM CRE_SOLINM WHERE "
   g_str_Parame = g_str_Parame & "SOLINM_NUMSOL = " & p_NroSol & " "
   g_str_Parame = g_str_Parame & "ORDER BY SOLINM_NUMSOL ASC"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
      Exit Function
   End If
   
   If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
      g_rst_Listas.MoveFirst
      
      Do While Not g_rst_Listas.EOF
         p_TipVia = g_rst_Listas!SOLINM_TIPVIA
         p_NomVia = Trim(g_rst_Listas!SOLINM_NOMVIA)
         p_NumVia = Trim(g_rst_Listas!SOLINM_NUMVIA)
         p_IntDpt = Trim(g_rst_Listas!SOLINM_INTDPT) & " "
         p_TipZon = g_rst_Listas!SOLINM_TIPZON
         p_NomZon = Trim(g_rst_Listas!SOLINM_NOMZON & " ")
         g_rst_Listas.MoveNext
      Loop
   End If
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Function

Sub Establecer_Ancho(combo As Object, Form As Form)
Dim k As Integer, Escala As Byte
Dim Maximo As Long
    
    'Guarda la escala del form para luego reestablecerla
    Escala = Form.ScaleMode
    
    'Cambia la escala
    Form.ScaleMode = vbPixels
    
    'Recorre los elementos para obtener el mas ancho
    For k = 0 To combo.ListCount - 1
        If Maximo < Form.TextWidth(combo.List(k)) Then
            Maximo = Form.TextWidth(combo.List(k))
        End If
    Next
    'Aplica el cambio pasandole el hwnd del combo y el valor
    SendMessage combo.hWnd, CB_SETDROPPEDWIDTH, Maximo + 18, 0
    'Reestablece la escala del form
    Form.ScaleMode = Escala

End Sub

Private Sub cmb_MonOpe_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If cmb_MonOpe.ListIndex > -1 Then
         Call gs_SetFocus(ipp_FecCtb)
      End If
   End If
End Sub

Private Sub ipp_FecCtb_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(msk_HorOpe)
   End If
End Sub

Private Sub msk_HorOpe_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_PatFis)
   End If
End Sub

Private Sub cmb_ValOpe_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If cmb_ValOpe.ListIndex > -1 Then
         Call gs_SetFocus(cmb_TipOpe)
      End If
   End If
End Sub

Private Sub cmb_TipOpe_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If cmb_TipOpe.ListIndex > -1 Then
         Call gs_SetFocus(cmb_CodBan)
      End If
   End If
End Sub

Private Sub cmb_CodBan_Click()
   If cmb_CodBan.ListIndex > -1 Then
      Screen.MousePointer = 11
      Call moddat_gs_Carga_CtaBan(l_arr_CodBan(cmb_CodBan.ListIndex + 1).Genera_Codigo, cmb_CtaBan, l_arr_CtaBan)
      Screen.MousePointer = 0
      Call gs_SetFocus(cmb_CtaBan)
   Else
      cmb_CtaBan.Clear
   End If
End Sub

Private Sub cmb_CodBan_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_CodBan_Click
   End If
End Sub

Private Sub cmb_CtaBan_Click()
   If cmb_CtaBan.ListIndex > -1 Then
      Call gs_SetFocus(ipp_OrgMto)
   End If
End Sub

Private Sub cmb_CtaBan_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_CtaBan_Click
   End If
End Sub

Private Sub ipp_OrgMto_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If CDbl(ipp_OrgMto.Text) > 0 Then
         Call gs_SetFocus(ipp_TipCam)
      End If
   End If
End Sub

Private Sub ipp_TipCam_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_TipFon)
   End If
End Sub

Private Sub cmb_TipFon_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If cmb_TipFon.ListIndex > -1 Then
         Call gs_SetFocus(txt_OrgOpe)
      End If
   End If
End Sub

Private Sub txt_OrgOpe_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_ComGra)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & "- '")
   End If
End Sub
