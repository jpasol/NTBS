VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmCYSCCR 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CY Special Services CCR Issuance"
   ClientHeight    =   10185
   ClientLeft      =   75
   ClientTop       =   660
   ClientWidth     =   15225
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "CYSCCR.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10185
   ScaleWidth      =   15225
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   " Last CCR Issuance "
      ForeColor       =   &H00004080&
      Height          =   840
      Left            =   3345
      TabIndex        =   109
      Top             =   7650
      Width           =   4845
      Begin VB.Label lblCCRLastIssue 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   390
         Left            =   120
         TabIndex        =   173
         Top             =   300
         Width           =   4575
      End
   End
   Begin VB.Frame fraRemarks 
      Caption         =   " Remarks "
      ForeColor       =   &H00004080&
      Height          =   915
      Left            =   150
      TabIndex        =   108
      Top             =   6675
      Width           =   8040
      Begin VB.TextBox txtRemark 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   150
         MaxLength       =   50
         TabIndex        =   158
         Top             =   300
         Width           =   7740
      End
   End
   Begin VB.Frame fraControl 
      Height          =   840
      Left            =   150
      TabIndex        =   89
      Top             =   7650
      Width           =   3090
      Begin VB.CheckBox chkNewCCR 
         Appearance      =   0  'Flat
         Caption         =   "&Add To Detail"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   390
         Left            =   225
         TabIndex        =   159
         Top             =   300
         Width           =   2715
      End
   End
   Begin TabDlg.SSTab tabTran 
      Height          =   4740
      Left            =   150
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1875
      Width           =   8040
      _ExtentX        =   14182
      _ExtentY        =   8361
      _Version        =   393216
      Tabs            =   6
      TabsPerRow      =   6
      TabHeight       =   706
      ForeColor       =   16512
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   16.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "ARR"
      TabPicture(0)   =   "CYSCCR.frx":014A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label8"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label9"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblARREntryNo"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblARRRegNo"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label73"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lblARRCCRNo"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lblArrPrevAmt"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label60"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtARRCCRNo"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtARRContSz"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtARRContNo"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "cboDanger"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Frame6"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Frame7"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).ControlCount=   16
      TabCaption(1)   =   "STO"
      TabPicture(1)   =   "CYSCCR.frx":0166
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame4"
      Tab(1).Control(1)=   "Frame2"
      Tab(1).Control(2)=   "txtStoContNo"
      Tab(1).Control(3)=   "txtStoContSz"
      Tab(1).Control(4)=   "txtSTOCCRNo"
      Tab(1).Control(5)=   "txtStoExtDate"
      Tab(1).Control(6)=   "lblStoPrevPay"
      Tab(1).Control(7)=   "Label75"
      Tab(1).Control(8)=   "Label57"
      Tab(1).Control(9)=   "Label13"
      Tab(1).Control(10)=   "Label49"
      Tab(1).Control(11)=   "Label47"
      Tab(1).Control(12)=   "lblStoEntryNo"
      Tab(1).Control(13)=   "lblStoRegNo"
      Tab(1).Control(14)=   "lblStoCCRNo"
      Tab(1).Control(15)=   "Label18"
      Tab(1).Control(16)=   "lblStoValidUntil"
      Tab(1).Control(17)=   "Label12"
      Tab(1).ControlCount=   18
      TabCaption(2)   =   "RFR"
      TabPicture(2)   =   "CYSCCR.frx":0182
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "txtRfrContNo"
      Tab(2).Control(1)=   "txtRFRCCRNo"
      Tab(2).Control(2)=   "txtRfrExtDate"
      Tab(2).Control(3)=   "txtRfrContSz"
      Tab(2).Control(4)=   "txtRfrEntryNo"
      Tab(2).Control(5)=   "txtRfrRegNo"
      Tab(2).Control(6)=   "txtRfrPlugInDate"
      Tab(2).Control(7)=   "lblRfrHrs"
      Tab(2).Control(8)=   "Label78"
      Tab(2).Control(9)=   "lblRfrPrevPay"
      Tab(2).Control(10)=   "Label59"
      Tab(2).Control(11)=   "Label54"
      Tab(2).Control(12)=   "Label37"
      Tab(2).Control(13)=   "Label27"
      Tab(2).Control(14)=   "lblRfrValidUntil"
      Tab(2).Control(15)=   "Label10"
      Tab(2).Control(16)=   "Label14"
      Tab(2).Control(17)=   "Label15"
      Tab(2).Control(18)=   "Label35"
      Tab(2).ControlCount=   19
      TabCaption(3)   =   "SO"
      TabPicture(3)   =   "CYSCCR.frx":019E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "txtSOVessel"
      Tab(3).Control(1)=   "txtSOContNo"
      Tab(3).Control(2)=   "txtSOCCRNo"
      Tab(3).Control(3)=   "Label64"
      Tab(3).Control(4)=   "lblSOFulEmp"
      Tab(3).Control(5)=   "lblSOContSz"
      Tab(3).Control(6)=   "Label39"
      Tab(3).Control(7)=   "Label38"
      Tab(3).Control(8)=   "Label21"
      Tab(3).Control(9)=   "Label19"
      Tab(3).ControlCount=   10
      TabCaption(4)   =   "MSC"
      TabPicture(4)   =   "CYSCCR.frx":01BA
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "txtMscContNo"
      Tab(4).Control(1)=   "txtMscContSz"
      Tab(4).Control(2)=   "txtMscQty"
      Tab(4).Control(3)=   "txtMscCCRNo"
      Tab(4).Control(4)=   "txtMscRateCode"
      Tab(4).Control(5)=   "Label36"
      Tab(4).Control(6)=   "lblMscAmount"
      Tab(4).Control(7)=   "lblMscRateUOM"
      Tab(4).Control(8)=   "Label28"
      Tab(4).Control(9)=   "lblMscRateAmt"
      Tab(4).Control(10)=   "Label26"
      Tab(4).Control(11)=   "Label11"
      Tab(4).Control(12)=   "lblMScRateDesc"
      Tab(4).Control(13)=   "Label5"
      Tab(4).Control(14)=   "Label4"
      Tab(4).Control(15)=   "Label16"
      Tab(4).ControlCount=   16
      TabCaption(5)   =   "OTH"
      TabPicture(5)   =   "CYSCCR.frx":01D6
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "txtOthVessel"
      Tab(5).Control(1)=   "Frame3"
      Tab(5).Control(2)=   "txtOthContSz"
      Tab(5).Control(3)=   "txtOthContNo"
      Tab(5).Control(4)=   "txtOTHCCRNo"
      Tab(5).Control(5)=   "txtOthAmount"
      Tab(5).Control(6)=   "txtOthEntryNo"
      Tab(5).Control(7)=   "txtOthRegNo"
      Tab(5).Control(8)=   "txtOthFulEmp"
      Tab(5).Control(9)=   "Label72"
      Tab(5).Control(10)=   "Label41"
      Tab(5).Control(11)=   "Label40"
      Tab(5).Control(12)=   "Label20"
      Tab(5).Control(13)=   "Label7"
      Tab(5).Control(14)=   "Label63"
      Tab(5).Control(15)=   "Label42"
      Tab(5).ControlCount=   16
      Begin VB.Frame Frame4 
         Enabled         =   0   'False
         Height          =   3090
         Left            =   -70125
         TabIndex        =   145
         Top             =   525
         Width           =   3015
         Begin VB.TextBox txtStoUOM 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1350
            MaxLength       =   1
            TabIndex        =   25
            Text            =   "I"
            Top             =   2100
            Width           =   315
         End
         Begin VB.CheckBox chkStoOvz 
            Appearance      =   0  'Flat
            Caption         =   "Oversize?"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   600
            TabIndex        =   21
            Top             =   300
            Width           =   1965
         End
         Begin MSMask.MaskEdBox txtStoOvzLen 
            Height          =   390
            Left            =   1350
            TabIndex        =   22
            Top             =   750
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   688
            _Version        =   393216
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox txtStoOvzWid 
            Height          =   390
            Left            =   1350
            TabIndex        =   23
            Top             =   1200
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   688
            _Version        =   393216
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox txtStoOvzHgt 
            Height          =   390
            Left            =   1350
            TabIndex        =   24
            Top             =   1650
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   688
            _Version        =   393216
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin VB.Label lblStoRevTon 
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   1350
            TabIndex        =   183
            Top             =   2550
            Width           =   1515
         End
         Begin VB.Label Label51 
            BackStyle       =   0  'Transparent
            Caption         =   "RevTon"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   150
            TabIndex        =   182
            Top             =   2625
            Width           =   1140
         End
         Begin VB.Label Label74 
            BackStyle       =   0  'Transparent
            Caption         =   "C/I"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1875
            TabIndex        =   150
            Top             =   2175
            Width           =   765
         End
         Begin VB.Label Label33 
            BackStyle       =   0  'Transparent
            Caption         =   "Length"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   150
            TabIndex        =   149
            Top             =   825
            Width           =   1215
         End
         Begin VB.Label Label32 
            BackStyle       =   0  'Transparent
            Caption         =   "Width"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   150
            TabIndex        =   148
            Top             =   1275
            Width           =   1215
         End
         Begin VB.Label Label31 
            BackStyle       =   0  'Transparent
            Caption         =   "Height"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   150
            TabIndex        =   147
            Top             =   1725
            Width           =   1215
         End
         Begin VB.Label Label30 
            BackStyle       =   0  'Transparent
            Caption         =   "UOM"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   150
            TabIndex        =   146
            Top             =   2175
            Width           =   1215
         End
      End
      Begin VB.TextBox txtOthVessel 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   -73050
         MaxLength       =   6
         TabIndex        =   48
         Top             =   3300
         Width           =   1665
      End
      Begin VB.Frame Frame3 
         Height          =   3090
         Left            =   -70125
         TabIndex        =   135
         Top             =   525
         Width           =   3015
         Begin VB.TextBox txtOthUOM 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1350
            MaxLength       =   1
            TabIndex        =   53
            Text            =   "I"
            Top             =   2100
            Width           =   315
         End
         Begin VB.CheckBox chkOthOvz 
            Appearance      =   0  'Flat
            Caption         =   "Oversize?"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   600
            TabIndex        =   49
            Top             =   300
            Width           =   1965
         End
         Begin MSMask.MaskEdBox txtOthOvzLen 
            Height          =   390
            Left            =   1350
            TabIndex        =   50
            Top             =   750
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   688
            _Version        =   393216
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox txtOthOvzWid 
            Height          =   390
            Left            =   1350
            TabIndex        =   51
            Top             =   1200
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   688
            _Version        =   393216
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox txtOthOvzHgt 
            Height          =   390
            Left            =   1350
            TabIndex        =   52
            Top             =   1650
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   688
            _Version        =   393216
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin VB.Label Label80 
            BackStyle       =   0  'Transparent
            Caption         =   "RevTon"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   150
            TabIndex        =   185
            Top             =   2625
            Width           =   1140
         End
         Begin VB.Label lblOthRevTon 
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   1350
            TabIndex        =   184
            Top             =   2550
            Width           =   1515
         End
         Begin VB.Label Label70 
            BackStyle       =   0  'Transparent
            Caption         =   "C/I"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1875
            TabIndex        =   140
            Top             =   2175
            Width           =   765
         End
         Begin VB.Label Label69 
            BackStyle       =   0  'Transparent
            Caption         =   "Length"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   150
            TabIndex        =   139
            Top             =   825
            Width           =   1215
         End
         Begin VB.Label Label61 
            BackStyle       =   0  'Transparent
            Caption         =   "Width"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   150
            TabIndex        =   138
            Top             =   1275
            Width           =   1215
         End
         Begin VB.Label Label44 
            BackStyle       =   0  'Transparent
            Caption         =   "Height"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   150
            TabIndex        =   137
            Top             =   1725
            Width           =   1215
         End
         Begin VB.Label Label43 
            BackStyle       =   0  'Transparent
            Caption         =   "UOM"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   150
            TabIndex        =   136
            Top             =   2175
            Width           =   1215
         End
      End
      Begin VB.Frame Frame2 
         Height          =   765
         Left            =   -74775
         TabIndex        =   111
         Top             =   525
         Width           =   4590
         Begin VB.OptionButton optStoImpExp 
            Appearance      =   0  'Flat
            Caption         =   "E&xport"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   315
            Index           =   1
            Left            =   2100
            TabIndex        =   17
            Top             =   300
            Width           =   1890
         End
         Begin VB.OptionButton optStoImpExp 
            Appearance      =   0  'Flat
            Caption         =   "&Import"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   315
            Index           =   0
            Left            =   150
            TabIndex        =   16
            Top             =   300
            Value           =   -1  'True
            Width           =   1740
         End
      End
      Begin VB.Frame Frame7 
         Height          =   765
         Left            =   225
         TabIndex        =   106
         Top             =   525
         Width           =   4590
         Begin VB.OptionButton optArrImpExp 
            Appearance      =   0  'Flat
            Caption         =   "&Import"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   315
            Index           =   0
            Left            =   150
            TabIndex        =   5
            Top             =   300
            Value           =   -1  'True
            Width           =   1740
         End
         Begin VB.OptionButton optArrImpExp 
            Appearance      =   0  'Flat
            Caption         =   "E&xport"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   315
            Index           =   1
            Left            =   2100
            TabIndex        =   6
            Top             =   300
            Width           =   1890
         End
      End
      Begin VB.Frame Frame6 
         Enabled         =   0   'False
         Height          =   2940
         Left            =   4875
         TabIndex        =   100
         Top             =   525
         Width           =   3015
         Begin VB.CheckBox chkARROvz 
            Appearance      =   0  'Flat
            Caption         =   "Oversize?"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   600
            TabIndex        =   10
            Top             =   225
            Width           =   1965
         End
         Begin VB.TextBox txtARRUOM 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1350
            MaxLength       =   1
            TabIndex        =   14
            Text            =   "I"
            Top             =   1950
            Width           =   315
         End
         Begin MSMask.MaskEdBox txtARROvzLen 
            Height          =   390
            Left            =   1350
            TabIndex        =   11
            Top             =   600
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   688
            _Version        =   393216
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox txtARROvzWid 
            Height          =   390
            Left            =   1350
            TabIndex        =   12
            Top             =   1050
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   688
            _Version        =   393216
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox txtARROvzHgt 
            Height          =   390
            Left            =   1350
            TabIndex        =   13
            Top             =   1500
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   688
            _Version        =   393216
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin VB.Label lblArrRevTon 
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   1350
            TabIndex        =   187
            Top             =   2400
            Width           =   1515
         End
         Begin VB.Label Label79 
            BackStyle       =   0  'Transparent
            Caption         =   "RevTon"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   150
            TabIndex        =   186
            Top             =   2475
            Width           =   1140
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "UOM"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   150
            TabIndex        =   105
            Top             =   2025
            Width           =   1215
         End
         Begin VB.Label Label22 
            BackStyle       =   0  'Transparent
            Caption         =   "Height"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   150
            TabIndex        =   104
            Top             =   1575
            Width           =   1215
         End
         Begin VB.Label Label23 
            BackStyle       =   0  'Transparent
            Caption         =   "Width"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   150
            TabIndex        =   103
            Top             =   1125
            Width           =   1215
         End
         Begin VB.Label Label24 
            BackStyle       =   0  'Transparent
            Caption         =   "Length"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   150
            TabIndex        =   102
            Top             =   675
            Width           =   1215
         End
         Begin VB.Label Label62 
            BackStyle       =   0  'Transparent
            Caption         =   "C/I"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1875
            TabIndex        =   101
            Top             =   2025
            Width           =   765
         End
      End
      Begin VB.ComboBox cboDanger 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   2175
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   4125
         Width           =   5715
      End
      Begin MSMask.MaskEdBox txtARRContNo 
         Height          =   390
         Left            =   2175
         TabIndex        =   8
         Top             =   2025
         Width           =   2580
         _ExtentX        =   4551
         _ExtentY        =   688
         _Version        =   393216
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         Enabled         =   0   'False
         MaxLength       =   12
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">AAAAAAAAAAAA"
         PromptChar      =   " "
      End
      Begin VB.TextBox txtSOVessel 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   -72825
         MaxLength       =   6
         TabIndex        =   36
         Top             =   3300
         Width           =   1665
      End
      Begin MSMask.MaskEdBox txtARRContSz 
         Height          =   390
         Left            =   2175
         TabIndex        =   9
         Top             =   2550
         Width           =   465
         _ExtentX        =   820
         _ExtentY        =   688
         _Version        =   393216
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         Enabled         =   0   'False
         MaxLength       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox txtARRCCRNo 
         Height          =   390
         Left            =   2175
         TabIndex        =   7
         Top             =   1500
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   688
         _Version        =   393216
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         Enabled         =   0   'False
         MaxLength       =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "########"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox txtStoContNo 
         Height          =   390
         Left            =   -72825
         TabIndex        =   19
         Top             =   1875
         Width           =   2580
         _ExtentX        =   4551
         _ExtentY        =   688
         _Version        =   393216
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         Enabled         =   0   'False
         MaxLength       =   12
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">AAAAAAAAAAAA"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox txtStoContSz 
         Height          =   390
         Left            =   -72825
         TabIndex        =   20
         Top             =   2325
         Width           =   465
         _ExtentX        =   820
         _ExtentY        =   688
         _Version        =   393216
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         Enabled         =   0   'False
         MaxLength       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox txtSTOCCRNo 
         Height          =   390
         Left            =   -72825
         TabIndex        =   18
         Top             =   1425
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   688
         _Version        =   393216
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         Enabled         =   0   'False
         MaxLength       =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "########"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox txtRfrContNo 
         Height          =   390
         Left            =   -72525
         TabIndex        =   28
         Top             =   1275
         Width           =   2580
         _ExtentX        =   4551
         _ExtentY        =   688
         _Version        =   393216
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         Enabled         =   0   'False
         MaxLength       =   12
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">&&&&&&&&&&&&"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox txtSOContNo 
         Height          =   390
         Left            =   -72825
         TabIndex        =   35
         Top             =   1500
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   688
         _Version        =   393216
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         Enabled         =   0   'False
         MaxLength       =   12
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox txtMscContNo 
         Height          =   390
         Left            =   -72750
         TabIndex        =   38
         Top             =   1275
         Width           =   2340
         _ExtentX        =   4128
         _ExtentY        =   688
         _Version        =   393216
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   12
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">AAAAAAAAAAAA"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox txtMscContSz 
         Height          =   390
         Left            =   -71040
         TabIndex        =   40
         Top             =   1800
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   688
         _Version        =   393216
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox txtMscQty 
         Height          =   390
         Left            =   -72750
         TabIndex        =   41
         Top             =   3375
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   688
         _Version        =   393216
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox txtOthContSz 
         Height          =   390
         Left            =   -73050
         TabIndex        =   44
         Top             =   1725
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   688
         _Version        =   393216
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox txtStoExtDate 
         Height          =   390
         Left            =   -72825
         TabIndex        =   26
         Top             =   4125
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   688
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "####-##-##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox txtOthContNo 
         Height          =   390
         Left            =   -73050
         TabIndex        =   43
         Top             =   1200
         Width           =   2580
         _ExtentX        =   4551
         _ExtentY        =   688
         _Version        =   393216
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   12
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">AAAAAAAAAAAA"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox txtSOCCRNo 
         Height          =   390
         Left            =   -72825
         TabIndex        =   34
         Top             =   900
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   688
         _Version        =   393216
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         Enabled         =   0   'False
         MaxLength       =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "########"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox txtOTHCCRNo 
         Height          =   390
         Left            =   -73050
         TabIndex        =   42
         Top             =   675
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   688
         _Version        =   393216
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "########"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox txtOthAmount 
         Height          =   390
         Left            =   -73050
         TabIndex        =   54
         Top             =   4050
         Width           =   2340
         _ExtentX        =   4128
         _ExtentY        =   688
         _Version        =   393216
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   12
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,###,##0.00"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox txtOthEntryNo 
         Height          =   390
         Left            =   -73050
         TabIndex        =   46
         Top             =   2250
         Width           =   2340
         _ExtentX        =   4128
         _ExtentY        =   688
         _Version        =   393216
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   12
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">AAAAAAAAAAAA"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox txtOthRegNo 
         Height          =   390
         Left            =   -73050
         TabIndex        =   47
         Top             =   2775
         Width           =   2340
         _ExtentX        =   4128
         _ExtentY        =   688
         _Version        =   393216
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   12
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">AAAAAAAAAAAA"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox txtOthFulEmp 
         Height          =   390
         Left            =   -72240
         TabIndex        =   45
         Top             =   1725
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   688
         _Version        =   393216
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">&"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox txtRFRCCRNo 
         Height          =   390
         Left            =   -72525
         TabIndex        =   27
         Top             =   750
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   688
         _Version        =   393216
         AutoTab         =   -1  'True
         Enabled         =   0   'False
         MaxLength       =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "########"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox txtMscCCRNo 
         Height          =   390
         Left            =   -72750
         TabIndex        =   37
         Top             =   750
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   688
         _Version        =   393216
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "########"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox txtMscRateCode 
         Height          =   390
         Left            =   -72750
         TabIndex        =   39
         Top             =   1800
         Width           =   1530
         _ExtentX        =   2699
         _ExtentY        =   688
         _Version        =   393216
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">AAAAAA"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox txtRfrExtDate 
         Height          =   390
         Left            =   -72525
         TabIndex        =   33
         Top             =   3900
         Width           =   2865
         _ExtentX        =   5054
         _ExtentY        =   688
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   16
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "####-##-## ##:##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox txtRfrContSz 
         Height          =   390
         Left            =   -69300
         TabIndex        =   29
         Top             =   1275
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   688
         _Version        =   393216
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         Enabled         =   0   'False
         MaxLength       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox txtRfrEntryNo 
         Height          =   390
         Left            =   -72525
         TabIndex        =   30
         Top             =   1800
         Width           =   2580
         _ExtentX        =   4551
         _ExtentY        =   688
         _Version        =   393216
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         Enabled         =   0   'False
         MaxLength       =   12
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">&&&&&&&&&&&&"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox txtRfrRegNo 
         Height          =   390
         Left            =   -72525
         TabIndex        =   31
         Top             =   2325
         Width           =   2580
         _ExtentX        =   4551
         _ExtentY        =   688
         _Version        =   393216
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         Enabled         =   0   'False
         MaxLength       =   12
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">&&&&&&&&&&&&"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox txtRfrPlugInDate 
         Height          =   390
         Left            =   -72525
         TabIndex        =   32
         Top             =   2850
         Width           =   2865
         _ExtentX        =   5054
         _ExtentY        =   688
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   16
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "####-##-## ##:##"
         PromptChar      =   " "
      End
      Begin VB.Label lblRfrHrs 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   -69525
         TabIndex        =   188
         Top             =   3900
         Width           =   2340
      End
      Begin VB.Label Label78 
         BackStyle       =   0  'Transparent
         Caption         =   "Paid"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -68850
         TabIndex        =   181
         Top             =   2925
         Width           =   840
      End
      Begin VB.Label lblRfrPrevPay 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   -69525
         TabIndex        =   180
         Top             =   3375
         Width           =   2340
      End
      Begin VB.Label Label36 
         BackStyle       =   0  'Transparent
         Caption         =   "Amount"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -74775
         TabIndex        =   157
         Top             =   3975
         Width           =   1965
      End
      Begin VB.Label lblMscAmount 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   -72750
         TabIndex        =   156
         Top             =   3900
         Width           =   2265
      End
      Begin VB.Label lblMscRateUOM 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   -70350
         TabIndex        =   155
         Top             =   2850
         Width           =   3240
      End
      Begin VB.Label Label28 
         BackStyle       =   0  'Transparent
         Caption         =   "Rate / UMS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -74775
         TabIndex        =   154
         Top             =   2925
         Width           =   1965
      End
      Begin VB.Label lblMscRateAmt 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   -72750
         TabIndex        =   153
         Top             =   2850
         Width           =   2265
      End
      Begin VB.Label Label26 
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -74775
         TabIndex        =   152
         Top             =   2400
         Width           =   1965
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Gatepass #"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -74775
         TabIndex        =   151
         Top             =   825
         Width           =   1965
      End
      Begin VB.Label lblStoPrevPay 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   -69600
         TabIndex        =   144
         Top             =   4125
         Width           =   2490
      End
      Begin VB.Label Label75 
         BackStyle       =   0  'Transparent
         Caption         =   "Paid"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -70500
         TabIndex        =   143
         Top             =   4200
         Width           =   765
      End
      Begin VB.Label Label59 
         BackStyle       =   0  'Transparent
         Caption         =   "Gatepass No"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -74700
         TabIndex        =   142
         Top             =   825
         Width           =   2115
      End
      Begin VB.Label Label72 
         BackStyle       =   0  'Transparent
         Caption         =   "Vessel"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -74775
         TabIndex        =   141
         Top             =   3375
         Width           =   1665
      End
      Begin VB.Label Label41 
         BackStyle       =   0  'Transparent
         Caption         =   "Entry No."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -74775
         TabIndex        =   134
         Top             =   2325
         Width           =   1665
      End
      Begin VB.Label Label40 
         BackStyle       =   0  'Transparent
         Caption         =   "Registry No."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -74775
         TabIndex        =   133
         Top             =   2850
         Width           =   1665
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "Amount"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -74775
         TabIndex        =   132
         Top             =   4125
         Width           =   1140
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "GPS / CCR"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -74775
         TabIndex        =   131
         Top             =   750
         Width           =   1665
      End
      Begin VB.Label Label60 
         BackStyle       =   0  'Transparent
         Caption         =   "Paid"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4875
         TabIndex        =   130
         Top             =   3675
         Width           =   765
      End
      Begin VB.Label lblArrPrevAmt 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   5625
         TabIndex        =   129
         Top             =   3600
         Width           =   2265
      End
      Begin VB.Label Label64 
         BackStyle       =   0  'Transparent
         Caption         =   "CCR No."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -74700
         TabIndex        =   128
         Top             =   975
         Width           =   1740
      End
      Begin VB.Label Label63 
         BackStyle       =   0  'Transparent
         Caption         =   "Container No."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -74775
         TabIndex        =   127
         Top             =   1275
         Width           =   1665
      End
      Begin VB.Label lblSOFulEmp 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   -72825
         TabIndex        =   123
         Top             =   2700
         Width           =   540
      End
      Begin VB.Label lblSOContSz 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   -72825
         TabIndex        =   122
         Top             =   2100
         Width           =   540
      End
      Begin VB.Label Label42 
         BackStyle       =   0  'Transparent
         Caption         =   "Size / FE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -74775
         TabIndex        =   126
         Top             =   1800
         Width           =   1665
      End
      Begin VB.Label Label39 
         BackStyle       =   0  'Transparent
         Caption         =   "Container"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -74700
         TabIndex        =   125
         Top             =   1575
         Width           =   1815
      End
      Begin VB.Label Label38 
         BackStyle       =   0  'Transparent
         Caption         =   "Size"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -74700
         TabIndex        =   124
         Top             =   2175
         Width           =   1815
      End
      Begin VB.Label Label57 
         BackStyle       =   0  'Transparent
         Caption         =   "Extend To"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -74775
         TabIndex        =   121
         Top             =   4200
         Width           =   1890
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Size"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -74775
         TabIndex        =   120
         Top             =   2400
         Width           =   1665
      End
      Begin VB.Label Label54 
         BackStyle       =   0  'Transparent
         Caption         =   "Entry No."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -74700
         TabIndex        =   119
         Top             =   1875
         Width           =   2115
      End
      Begin VB.Label Label37 
         BackStyle       =   0  'Transparent
         Caption         =   "Registry No."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   -74700
         TabIndex        =   118
         Top             =   2400
         Width           =   2115
      End
      Begin VB.Label Label27 
         BackStyle       =   0  'Transparent
         Caption         =   "Container No."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -74700
         TabIndex        =   117
         Top             =   1350
         Width           =   2115
      End
      Begin VB.Label Label49 
         BackStyle       =   0  'Transparent
         Caption         =   "Entry No."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -74775
         TabIndex        =   116
         Top             =   2850
         Width           =   1665
      End
      Begin VB.Label Label47 
         BackStyle       =   0  'Transparent
         Caption         =   "Registry No."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -74775
         TabIndex        =   115
         Top             =   3300
         Width           =   1665
      End
      Begin VB.Label lblStoEntryNo 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   -72825
         TabIndex        =   114
         Top             =   2775
         Width           =   2040
      End
      Begin VB.Label lblStoRegNo 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   -72825
         TabIndex        =   113
         Top             =   3225
         Width           =   2040
      End
      Begin VB.Label lblStoCCRNo 
         BackStyle       =   0  'Transparent
         Caption         =   "Gatepass No"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -74775
         TabIndex        =   112
         Top             =   1500
         Width           =   1890
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "Container #"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -74775
         TabIndex        =   110
         Top             =   1950
         Width           =   1890
      End
      Begin VB.Label lblARRCCRNo 
         BackStyle       =   0  'Transparent
         Caption         =   "Gatepass No"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   225
         TabIndex        =   107
         Top             =   1575
         Width           =   1890
      End
      Begin VB.Label Label73 
         BackStyle       =   0  'Transparent
         Caption         =   "Danger Class"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   225
         TabIndex        =   99
         Top             =   4200
         Width           =   1905
      End
      Begin VB.Label lblRfrValidUntil 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   -72525
         TabIndex        =   98
         Top             =   3375
         Width           =   2865
      End
      Begin VB.Label lblStoValidUntil 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   -72825
         TabIndex        =   97
         Top             =   3675
         Width           =   2040
      End
      Begin VB.Label lblARRRegNo 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   2175
         TabIndex        =   96
         Top             =   3600
         Width           =   2580
      End
      Begin VB.Label lblARREntryNo 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   2175
         TabIndex        =   95
         Top             =   3075
         Width           =   2580
      End
      Begin VB.Label lblMScRateDesc 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   -72750
         TabIndex        =   94
         Top             =   2325
         Width           =   5640
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Container"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -74775
         TabIndex        =   88
         Top             =   1350
         Width           =   1965
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Rate Code"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -74775
         TabIndex        =   87
         Top             =   1875
         Width           =   1965
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Quantity"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -74775
         TabIndex        =   86
         Top             =   3450
         Width           =   1965
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "Vessel"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -74700
         TabIndex        =   85
         Top             =   3375
         Width           =   1815
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "Full/Empty"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -74700
         TabIndex        =   84
         Top             =   2775
         Width           =   1815
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Extend Until"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -74700
         TabIndex        =   83
         Top             =   3975
         Width           =   2115
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Valid Until"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -74700
         TabIndex        =   82
         Top             =   3450
         Width           =   2115
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Plug-In"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -74700
         TabIndex        =   81
         Top             =   2925
         Width           =   2115
      End
      Begin VB.Label Label35 
         BackStyle       =   0  'Transparent
         Caption         =   "Sz"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -69825
         TabIndex        =   80
         Top             =   1350
         Width           =   390
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Valid Until"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -74775
         TabIndex        =   79
         Top             =   3750
         Width           =   1890
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Registry No."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   225
         TabIndex        =   77
         Top             =   3675
         Width           =   1665
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Entry No."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   225
         TabIndex        =   76
         Top             =   3150
         Width           =   1665
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Size"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   225
         TabIndex        =   75
         Top             =   2625
         Width           =   1665
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Container #"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   225
         TabIndex        =   74
         Top             =   2100
         Width           =   1965
      End
   End
   Begin VB.Frame fraCustomer 
      Caption         =   " Customer "
      ForeColor       =   &H00004080&
      Height          =   1515
      Left            =   150
      TabIndex        =   73
      Top             =   225
      Width           =   8040
      Begin VB.CheckBox chkVAT 
         Appearance      =   0  'Flat
         Caption         =   "&VAT"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   450
         TabIndex        =   1
         Top             =   975
         Width           =   1320
      End
      Begin VB.CheckBox chkWTax 
         Appearance      =   0  'Flat
         Caption         =   "&W/Tax"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   2550
         TabIndex        =   2
         Top             =   975
         Width           =   1650
      End
      Begin VB.CheckBox chkGuarantee 
         Appearance      =   0  'Flat
         Caption         =   "&Under Guarantee"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   4650
         TabIndex        =   3
         Top             =   975
         Width           =   3015
      End
      Begin MSMask.MaskEdBox txtCusName 
         Height          =   390
         Left            =   150
         TabIndex        =   0
         Top             =   300
         Width           =   7740
         _ExtentX        =   13653
         _ExtentY        =   688
         _Version        =   393216
         AutoTab         =   -1  'True
         MaxLength       =   50
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grdCCRTran 
      Height          =   3180
      Left            =   8325
      TabIndex        =   55
      Top             =   675
      Width           =   6765
      _ExtentX        =   11933
      _ExtentY        =   5609
      _Version        =   393216
      Cols            =   30
      ForeColorFixed  =   16512
      BackColorSel    =   65535
      ForeColorSel    =   -2147483640
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      Enabled         =   0   'False
      FocusRect       =   2
      SelectionMode   =   1
      AllowUserResizing=   1
      FormatString    =   "  | RATE      |  AMOUNT  |  VAT  |WTAX |  TOTAL  |#"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame fraPayment 
      Enabled         =   0   'False
      Height          =   5790
      Left            =   8325
      TabIndex        =   90
      Top             =   4275
      Width           =   6765
      Begin VB.TextBox txtLog 
         Height          =   855
         Left            =   600
         MultiLine       =   -1  'True
         TabIndex        =   189
         Text            =   "CYSCCR.frx":01F2
         Top             =   4680
         Visible         =   0   'False
         Width           =   3255
      End
      Begin MSMask.MaskEdBox txtCshAmt 
         Height          =   390
         Left            =   2400
         TabIndex        =   56
         Top             =   675
         Width           =   2190
         _ExtentX        =   3863
         _ExtentY        =   688
         _Version        =   393216
         ForeColor       =   16711680
         AutoTab         =   -1  'True
         MaxLength       =   12
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,###,##0.00"
         PromptChar      =   " "
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save|Print"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   4725
         TabIndex        =   72
         Top             =   4800
         Width           =   1890
      End
      Begin MSMask.MaskEdBox txtChkAmt 
         Height          =   390
         Index           =   0
         Left            =   150
         TabIndex        =   57
         Top             =   1800
         Width           =   2190
         _ExtentX        =   3863
         _ExtentY        =   688
         _Version        =   393216
         AutoTab         =   -1  'True
         MaxLength       =   12
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,###,##0.00"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox txtChkAmt 
         Height          =   390
         Index           =   1
         Left            =   150
         TabIndex        =   60
         Top             =   2250
         Width           =   2190
         _ExtentX        =   3863
         _ExtentY        =   688
         _Version        =   393216
         AutoTab         =   -1  'True
         MaxLength       =   12
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,###,##0.00"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox txtChkAmt 
         Height          =   390
         Index           =   2
         Left            =   150
         TabIndex        =   63
         Top             =   2700
         Width           =   2190
         _ExtentX        =   3863
         _ExtentY        =   688
         _Version        =   393216
         AutoTab         =   -1  'True
         MaxLength       =   12
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,###,##0.00"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox txtChkAmt 
         Height          =   390
         Index           =   3
         Left            =   150
         TabIndex        =   66
         Top             =   3150
         Width           =   2190
         _ExtentX        =   3863
         _ExtentY        =   688
         _Version        =   393216
         AutoTab         =   -1  'True
         MaxLength       =   12
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,###,##0.00"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox txtChkAmt 
         Height          =   390
         Index           =   4
         Left            =   150
         TabIndex        =   69
         Top             =   3600
         Width           =   2190
         _ExtentX        =   3863
         _ExtentY        =   688
         _Version        =   393216
         AutoTab         =   -1  'True
         MaxLength       =   12
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,###,##0.00"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox txtChkBank 
         Height          =   390
         Index           =   0
         Left            =   4650
         TabIndex        =   59
         Top             =   1800
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   688
         _Version        =   393216
         AutoTab         =   -1  'True
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">&&&&&&&&&&"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox txtChkBank 
         Height          =   390
         Index           =   1
         Left            =   4650
         TabIndex        =   62
         Top             =   2250
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   688
         _Version        =   393216
         AutoTab         =   -1  'True
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">&&&&&&&&&&"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox txtChkBank 
         Height          =   390
         Index           =   2
         Left            =   4650
         TabIndex        =   65
         Top             =   2700
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   688
         _Version        =   393216
         AutoTab         =   -1  'True
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">&&&&&&&&&&"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox txtChkBank 
         Height          =   390
         Index           =   3
         Left            =   4650
         TabIndex        =   68
         Top             =   3150
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   688
         _Version        =   393216
         AutoTab         =   -1  'True
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">&&&&&&&&&&"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox txtChkBank 
         Height          =   390
         Index           =   4
         Left            =   4650
         TabIndex        =   71
         Top             =   3600
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   688
         _Version        =   393216
         AutoTab         =   -1  'True
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">&&&&&&&&&&"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox txtChkNo 
         Height          =   390
         Index           =   0
         Left            =   2400
         TabIndex        =   58
         Top             =   1800
         Width           =   2190
         _ExtentX        =   3863
         _ExtentY        =   688
         _Version        =   393216
         AutoTab         =   -1  'True
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">&&&&&&&&&&"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox txtChkNo 
         Height          =   390
         Index           =   1
         Left            =   2400
         TabIndex        =   61
         Top             =   2250
         Width           =   2190
         _ExtentX        =   3863
         _ExtentY        =   688
         _Version        =   393216
         AutoTab         =   -1  'True
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">&&&&&&&&&&"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox txtChkNo 
         Height          =   390
         Index           =   2
         Left            =   2400
         TabIndex        =   64
         Top             =   2700
         Width           =   2190
         _ExtentX        =   3863
         _ExtentY        =   688
         _Version        =   393216
         AutoTab         =   -1  'True
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">&&&&&&&&&&"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox txtChkNo 
         Height          =   390
         Index           =   3
         Left            =   2400
         TabIndex        =   67
         Top             =   3150
         Width           =   2190
         _ExtentX        =   3863
         _ExtentY        =   688
         _Version        =   393216
         AutoTab         =   -1  'True
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">&&&&&&&&&&"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox txtChkNo 
         Height          =   390
         Index           =   4
         Left            =   2400
         TabIndex        =   70
         Top             =   3600
         Width           =   2190
         _ExtentX        =   3863
         _ExtentY        =   688
         _Version        =   393216
         AutoTab         =   -1  'True
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">&&&&&&&&&&"
         PromptChar      =   " "
      End
      Begin VB.Label Label45 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Caption         =   "CHANGE"
         ForeColor       =   &H00004080&
         Height          =   315
         Left            =   4650
         TabIndex        =   177
         Top             =   300
         Width           =   1965
      End
      Begin VB.Label Label50 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Caption         =   "AMOUNT DUE"
         ForeColor       =   &H00004080&
         Height          =   315
         Left            =   150
         TabIndex        =   175
         Top             =   300
         Width           =   2190
      End
      Begin VB.Label Label48 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Caption         =   "CASH AMT"
         ForeColor       =   &H00004080&
         Height          =   315
         Left            =   2400
         TabIndex        =   174
         Top             =   300
         Width           =   2190
      End
      Begin VB.Label Label34 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Caption         =   "CHECK AMT"
         ForeColor       =   &H00004080&
         Height          =   315
         Left            =   150
         TabIndex        =   162
         Top             =   1425
         Width           =   2190
      End
      Begin VB.Label Label29 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Caption         =   "CHECK BANK"
         ForeColor       =   &H00004080&
         Height          =   315
         Left            =   4650
         TabIndex        =   161
         Top             =   1425
         Width           =   1965
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Caption         =   "CHECK NO"
         ForeColor       =   &H00004080&
         Height          =   315
         Left            =   2400
         TabIndex        =   160
         Top             =   1425
         Width           =   2190
      End
      Begin VB.Label lblChkTot 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   150
         TabIndex        =   93
         Top             =   4050
         Width           =   2190
      End
      Begin VB.Label lblChange 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   390
         Left            =   4650
         TabIndex        =   92
         Top             =   675
         Width           =   1965
      End
      Begin VB.Label lblAmtDue 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   150
         TabIndex        =   91
         Top             =   675
         Width           =   2190
      End
   End
   Begin VB.Label Label53 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "F12"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   480
      Left            =   4230
      TabIndex        =   179
      Top             =   9675
      Width           =   675
   End
   Begin VB.Label Label52 
      Appearance      =   0  'Flat
      BackColor       =   &H80000010&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " SAVE && PRINT CCR/s"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   480
      Left            =   5025
      TabIndex        =   178
      Top             =   9675
      Width           =   3105
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000010&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "PAYMENT DETAILS"
      ForeColor       =   &H00C0E0FF&
      Height          =   240
      Left            =   8325
      TabIndex        =   176
      Top             =   3975
      Width           =   6765
   End
   Begin VB.Label Label77 
      Appearance      =   0  'Flat
      BackColor       =   &H80000010&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " PREPARE PAYMENT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   465
      Left            =   5025
      TabIndex        =   172
      Top             =   9150
      Width           =   3105
   End
   Begin VB.Label Label76 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "F11"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   465
      Left            =   4230
      TabIndex        =   171
      Top             =   9150
      Width           =   675
   End
   Begin VB.Label Label71 
      Appearance      =   0  'Flat
      BackColor       =   &H80000010&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " VOID CCR DETAIL"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   465
      Left            =   900
      TabIndex        =   170
      Top             =   9150
      Width           =   3105
   End
   Begin VB.Label Label68 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "F9"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   465
      Left            =   225
      TabIndex        =   169
      Top             =   9150
      Width           =   555
   End
   Begin VB.Label Label67 
      Appearance      =   0  'Flat
      BackColor       =   &H80000010&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " EXIT CCR ISSUANCE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   480
      Left            =   900
      TabIndex        =   168
      Top             =   9675
      Width           =   3105
   End
   Begin VB.Label Label66 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "F3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   480
      Left            =   225
      TabIndex        =   167
      Top             =   9675
      Width           =   555
   End
   Begin VB.Label Label65 
      Appearance      =   0  'Flat
      BackColor       =   &H80000010&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " SELECT TRANSACTION TYPE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   480
      Left            =   5025
      TabIndex        =   166
      Top             =   8625
      Width           =   3105
   End
   Begin VB.Label Label58 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "F5"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   480
      Left            =   4230
      TabIndex        =   165
      Top             =   8625
      Width           =   675
   End
   Begin VB.Label Label56 
      Appearance      =   0  'Flat
      BackColor       =   &H80000010&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " SELECT CCR DETAIL"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   480
      Left            =   900
      TabIndex        =   164
      Top             =   8625
      Width           =   3105
   End
   Begin VB.Label Label46 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "F6"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   480
      Left            =   225
      TabIndex        =   163
      Top             =   8625
      Width           =   555
   End
   Begin VB.Label Label25 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000010&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CCR DETAILS"
      ForeColor       =   &H00C0E0FF&
      Height          =   240
      Left            =   8325
      TabIndex        =   78
      Top             =   375
      Width           =   6765
   End
   Begin VB.Menu mnuMeu 
      Caption         =   "&Menu"
      Begin VB.Menu mnuMenuGrid 
         Caption         =   "Select from &Grid"
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnuMenuEdit 
         Caption         =   "&Edit selected item"
         Enabled         =   0   'False
         Shortcut        =   {F7}
      End
      Begin VB.Menu mnuMenuDelete 
         Caption         =   "&Delete selected item"
         Enabled         =   0   'False
         Shortcut        =   {F9}
      End
      Begin VB.Menu mnuF1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMenuTab 
         Caption         =   "Select &Next Tab"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuF3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMenuPayment 
         Caption         =   "Prepare &Payment"
         Enabled         =   0   'False
         Shortcut        =   {F11}
      End
      Begin VB.Menu mnuF4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMenuSave 
         Caption         =   "&Save / Print transaction"
         Enabled         =   0   'False
         Shortcut        =   {F12}
      End
      Begin VB.Menu mnuF2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMenuExit 
         Caption         =   "E&xit"
         Shortcut        =   {F3}
      End
   End
End
Attribute VB_Name = "frmCYSCCR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    
Private Enum enGridCol          ' grid column identifiers
    enCounter = 0
    enRateCode = 1
    enAmount = 2
    enVATAmt = 3
    enWTaxAmt = 4
    enTotalAmt = 5
    enCCRTag = 6
    enCCRNo = 7
    enContNo = 8
    enContSz = 9
    enFulEmp = 10
    enEntryNo = 11
    enRegNo = 12
    enOvzLen = 13
    enOvzWid = 14
    enOvzHgt = 15
    enOvzUom = 16
    enRevTon = 17
    enDangerCode = 18
    enStoValidUntil = 19
    enRfrValidUntil = 20
    enStoDays = 21
    enQuantity = 22
    enVessel = 23
    enDangerAmt = 24
    enOvzAmt = 25
    enRemark = 26
    enShipLine = 27
    enGuaranty = 28
    enRfrHours = 29
End Enum

' constants
Const cVoid = "*VOID*"
Const cEmptyRfrDate = "    -  -     :  "

Const cNullDate = #12:00:00 AM#
Const cRTon20 As Currency = 27.95       ' fixed revenue ton for 20-footer
Const cRTon40 As Currency = 63.75       ' fixed revenue ton for 40-footer
Const cRTon45 As Currency = 76.38       ' fixed revenue ton for 45-footer

' variables
Dim vRevTon As Single
Dim vRevTonRateArr, vRevTonRateSto As Currency
Dim vRevTonRateArrExp As Currency
Dim cRateCode, vCusCodeUnderG As String
Dim vStoDay, vTabOn As Integer
Dim vRfrHours, vCYMStoDay As Long

Dim bNewCCR, bArrImp, bStoImp, bVAT, bWTax, bUnderG, bARROversize, bStoOversize, bOTHOversize As Boolean
Dim nPtr, nCCRCounter As Integer
Dim nAmount, nVATAmount, nWTaxAmount, nTotalAmount As Currency
Dim bEscaped As Boolean

Dim clsCCRPrinter As Object

Private Sub cboDanger_GotFocus()
    cboDanger.BackColor = vbInfoBackground
End Sub

Private Sub cboDanger_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyEscape
            SendKeys ("+{TAB}")
        Case vbKeySpace
            SendKeys ("{F4}")
        Case vbKeyReturn
            txtRemark.SetFocus
        Case Else
    End Select
End Sub

Private Sub cboDanger_LostFocus()
    cboDanger.BackColor = vbWindowBackground
End Sub
Private Sub WriteLogError()
    Dim fso As New FileSystemObject
    Dim txtStream As TextStream
    Dim str
    Dim strErr As String
    str = fso.BuildPath("c:\BillingLog", Day(Now) & Month(Now) & Year(Now) & ".txt")
    
    If fso.FileExists(str) Then
        Set txtStream = fso.OpenTextFile(str)
        With txtStream
                Do Until .AtEndOfStream
                        strErr = strErr & .ReadLine & vbCrLf
                    Loop
            End With

            Set txtStream = fso.OpenTextFile(str, 2)
            txtStream.Write strErr & txtLog.Text
    Else
        fso.CreateFolder ("c:\BillingLog")
        Set txtStream = fso.CreateTextFile(str, True)
            txtStream.Write strErr & txtLog.Text
    End If
    txtLog.Text = ""



End Sub
Private Sub chkARROvz_GotFocus()
    chkARROvz.BackColor = vbInfoBackground
End Sub

Private Sub chkARROvz_LostFocus()
    chkARROvz.BackColor = vbButtonFace
End Sub

Private Sub chkGuarantee_Click()
    bUnderG = (chkGuarantee.Value = 1)
    SendKeys "{TAB}"
    If bUnderG Then Call lzCustomerUnderG
End Sub

Private Sub chkGuarantee_GotFocus()
    chkGuarantee.BackColor = vbInfoBackground
End Sub

Private Sub chkGuarantee_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyEscape
            SendKeys ("+{TAB}")
            KeyAscii = 0
        Case vbKeyReturn
            tabTran.SetFocus
            'SendKeys ("{TAB}")
            'KeyAscii = 0
        Case Else
    End Select
End Sub

Private Sub chkGuarantee_LostFocus()
    chkGuarantee.BackColor = vbButtonFace
End Sub

Private Sub chkNewCCR_Click()
    bNewCCR = (chkNewCCR.Value = 1)
End Sub

Private Sub chkNewCCR_GotFocus()
    chkNewCCR.BackColor = vbInfoBackground
    If txtMscRateCode = "WEIGHT" Then
        chkNewCCR.Value = 1
        bNewCCR = True
    End If
End Sub

Private Sub chkNewCCR_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyEscape
            txtRemark.SetFocus
            KeyAscii = 0
        Case vbKeyReturn
            Call lzAddTran
        Case Else
    End Select
End Sub

Private Sub chkARROvz_Click()
    bARROversize = (chkARROvz.Value = 1)
    txtARROvzLen.Enabled = bARROversize
    txtARROvzWid.Enabled = bARROversize
    txtARROvzHgt.Enabled = bARROversize
    txtARRUOM.Enabled = bARROversize
    If Not bARROversize Then lblArrRevTon = ""
    SendKeys ("{TAB}")
End Sub

Private Sub chkARROvz_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyEscape
            SendKeys ("+{TAB}")
            KeyAscii = 0
        Case vbKeyReturn
            SendKeys ("{TAB}")
            KeyAscii = 0
        Case Else
    End Select
End Sub

Private Sub chkNewCCR_LostFocus()
    chkNewCCR.BackColor = vbButtonFace
End Sub

Private Sub chkOthOvz_Click()
    bOTHOversize = (chkOthOvz.Value = 1)
    txtOthOvzLen.Enabled = bOTHOversize
    txtOthOvzWid.Enabled = bOTHOversize
    txtOthOvzHgt.Enabled = bOTHOversize
    txtOthUOM.Enabled = bOTHOversize
    If Not bOTHOversize Then lblOthRevTon = ""
    SendKeys ("{TAB}")
End Sub

Private Sub chkOthOvz_GotFocus()
    chkOthOvz.BackColor = vbInfoBackground
End Sub

Private Sub chkOthOvz_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyEscape
            SendKeys ("+{TAB}")
            KeyAscii = 0
        Case vbKeyReturn
            SendKeys ("{TAB}")
            KeyAscii = 0
        Case Else
    End Select
End Sub

Private Sub chkOthOvz_LostFocus()
    chkOthOvz.BackColor = vbButtonFace
End Sub

Private Sub chkStoOvz_Click()
    bStoOversize = (chkStoOvz.Value = 1)
    txtStoOvzLen.Enabled = bStoOversize
    txtStoOvzWid.Enabled = bStoOversize
    txtStoOvzHgt.Enabled = bStoOversize
    txtStoUOM.Enabled = bStoOversize
    If Not bStoOversize Then lblStoRevTon = ""
    SendKeys ("{TAB}")
End Sub

Private Sub chkStoOvz_GotFocus()
    chkStoOvz.BackColor = vbInfoBackground
End Sub

Private Sub chkStoOvz_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyEscape
            SendKeys ("+{TAB}")
            KeyAscii = 0
        Case vbKeyReturn
            SendKeys ("{TAB}")
            KeyAscii = 0
        Case Else
    End Select
End Sub

Private Sub chkStoOvz_LostFocus()
    chkStoOvz.BackColor = vbButtonFace
End Sub

Private Sub chkVAT_Click()
    bVAT = (chkVAT.Value = 1)
    Call lzUpdateGridVAT
    SendKeys "{TAB}"
End Sub

Private Sub chkVAT_GotFocus()
    chkVAT.BackColor = vbInfoBackground
End Sub

Private Sub chkVAT_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyEscape
            SendKeys ("+{TAB}")
            KeyAscii = 0
        Case vbKeyReturn
            SendKeys ("{TAB}")
            KeyAscii = 0
        Case Else
    End Select
End Sub

Private Sub chkVAT_LostFocus()
    chkVAT.BackColor = vbButtonFace
End Sub

Private Sub chkWTax_Click()
    bWTax = (chkWTax.Value = 1)
    Call lzUpdateGridWTax
    SendKeys "{TAB}"
End Sub

Private Sub chkWTax_GotFocus()
    chkWTax.BackColor = vbInfoBackground
End Sub

Private Sub chkWTax_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyEscape
            SendKeys ("+{TAB}")
            KeyAscii = 0
        Case vbKeyReturn
            SendKeys ("{TAB}")
            KeyAscii = 0
        Case Else
    End Select
End Sub

Private Sub chkWTax_LostFocus()
    chkWTax.BackColor = vbButtonFace
End Sub

Private Sub cmdSave_Click()
    Call lzSavePrint
End Sub

Private Sub cmdSave_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyEscape
            txtCshAmt.SetFocus
        Case Else
    End Select
End Sub

Private Sub Form_Load()
    
    ' create printer component
    Set clsCCRPrinter = CreateObject("CCRPR03.clsCCRPR03")
    
    'initialize
    vTabOn = 4
    Call lzInitialize
    ' populate combo boxes
    Call lzPopulateDangerClass
    ' get user info
    Call lzGetUserInfo
    ' get rates
    vRevTonRateArr = lzGetRateInfo("RTARIM")
    vRevTonRateSto = lzGetRateInfo("RTSTIM")
    vRevTonRateArrExp = lzGetRateInfo("RTAREX")
    '
    grdCCRTran.TextMatrix(grdCCRTran.Rows - 1, enCounter) = "**"
End Sub

Private Sub Form_Terminate()
    On Error Resume Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Cancel = (MsgBox("Exit CY Special Services Data Entry?", vbYesNo, "Quit") = vbNo)
End Sub

Private Sub grdCCRTran_GotFocus()
    mnuMenuEdit.Enabled = True
    mnuMenuDelete.Enabled = True
End Sub

Private Sub grdCCRTran_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyEscape
            tabTran.SetFocus
        Case vbKeyReturn
            Call lzInitializePay
        Case Else
    End Select
End Sub

Private Sub grdCCRTran_LostFocus()
    mnuMenuEdit.Enabled = False
    mnuMenuDelete.Enabled = False
End Sub

Private Sub mnuMenuDelete_Click()
    Call lzDeleteItem
End Sub

Private Sub mnuMenuExit_Click()
    Unload Me
End Sub

Private Sub mnuMenuGrid_Click()
    If grdCCRTran.Rows > 2 Then
        grdCCRTran.SetFocus
        SendKeys ("{RIGHT}")
    End If
End Sub

Private Sub mnuMenuPayment_Click()
    Call lzInitializePay
End Sub

Private Sub mnuMenuSave_Click()
    Call lzSavePrint
End Sub

Private Sub mnuMenuTab_Click()
    With tabTran
'        If .Tab = (.Tabs - 1) Then
'            .Tab = 0
'        Else
'            .Tab = .Tab + 1
'        End If
        
        If .Tab = (.Tabs - 1) Then
            .Tab = 4
        Else
            .Tab = .Tab + 1
        End If
        
        vTabOn = .Tab
        Call lzEnableTab
        tabTran.SetFocus
    End With
End Sub

Private Sub optArrImpExp_Click(Index As Integer)
    optArrImpExp(Index).BackColor = vbButtonFace
    bArrImp = optArrImpExp(0).Value
    lblARRCCRNo = IIf(bArrImp, "Gatepass No", "CCR No")
    txtARRCCRNo.Enabled = True
End Sub

Private Sub optArrImpExp_GotFocus(Index As Integer)
    optArrImpExp(Index).BackColor = vbInfoBackground
End Sub

Private Sub optArrImpExp_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyEscape
            chkGuarantee.SetFocus
            KeyAscii = 0
        Case vbKeyReturn
            SendKeys ("{TAB}")
            KeyAscii = 0
        Case Else
    End Select
End Sub

Private Sub optArrImpExp_LostFocus(Index As Integer)
    optArrImpExp(Index).BackColor = vbButtonFace
    bArrImp = optArrImpExp(0).Value
    lblARRCCRNo = IIf(bArrImp, "Gatepass No", "CCR No")
    txtARRCCRNo.Enabled = True
End Sub

Private Sub optArrImpExp_Validate(Index As Integer, Cancel As Boolean)
    Cancel = Not (optArrImpExp(0).Value Or optArrImpExp(1).Value)
End Sub

Private Sub optStoImpExp_GotFocus(Index As Integer)
    optStoImpExp(Index).BackColor = vbInfoBackground
End Sub

Private Sub optStoImpExp_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyEscape
            chkGuarantee.SetFocus
            KeyAscii = 0
        Case vbKeyReturn
            SendKeys ("{TAB}")
            KeyAscii = 0
        Case Else
    End Select
End Sub

Private Sub optStoImpExp_LostFocus(Index As Integer)
    optStoImpExp(Index).BackColor = vbButtonFace
    bStoImp = optStoImpExp(0).Value
    lblStoCCRNo = IIf(bStoImp, "Gatepass No", "CCR No")
    txtSTOCCRNo.Enabled = True
End Sub

Private Sub optStoImpExp_Validate(Index As Integer, Cancel As Boolean)
    Cancel = Not (optStoImpExp(0).Value Or optArrImpExp(1).Value)
End Sub

Private Sub tabTran_GotFocus()
    Select Case tabTran.Tab
        Case 0
            'Call lzClearArr
        Case 1
            'Call lzClearSto
        Case 2
            'Call lzClearRfr
        Case 3
            'Call lzClearSO
        Case 4
            Call lzClearMsc
        Case 5
            Call lzClearOth
        Case Else
    End Select
End Sub

Private Sub txtARRCCRNo_GotFocus()
    With txtARRCCRNo
        .BackColor = vbInfoBackground
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub

Private Sub txtARRCCRNo_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyEscape
            SendKeys ("+{TAB}")
            KeyAscii = 0
        Case vbKeyReturn
            SendKeys ("{TAB}")
            KeyAscii = 0
        Case Else
    End Select
End Sub

Private Sub txtARRCCRNo_LostFocus()
    txtARRCCRNo.BackColor = vbWindowBackground
    If (Val(txtARRCCRNo)) > 0 Then
        If bArrImp Then
            Call lzGetCYMArr(txtARRCCRNo)
        Else
            Call lzGetCYXArr(txtARRCCRNo)
        End If
    End If
End Sub

Private Sub txtARRContNo_GotFocus()
    With txtARRContNo
        .BackColor = vbInfoBackground
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub

Private Sub txtARRContNo_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyEscape
            SendKeys ("+{TAB}")
            KeyAscii = 0
        Case vbKeyReturn
            SendKeys ("{TAB}")
            KeyAscii = 0
        Case Else
    End Select
End Sub

Private Sub txtARRContNo_LostFocus()
    txtARRContNo.BackColor = vbWindowBackground
End Sub

Private Sub txtARRContSz_GotFocus()
    With txtARRContSz
        .BackColor = vbInfoBackground
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub

Private Sub txtARRContSz_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyEscape
            SendKeys ("+{TAB}")
            KeyAscii = 0
        Case vbKeyReturn
            SendKeys ("{TAB}")
            KeyAscii = 0
        Case Else
    End Select
End Sub

Private Sub txtARRContSz_LostFocus()
    txtARRContSz.BackColor = vbWindowBackground
End Sub

Private Sub txtARRContSz_Validate(Cancel As Boolean)
    Cancel = InStr("20|40|45|", txtARRContSz & "|") = 0
    If Cancel Then MsgBox "Invalid container size. Please correct..."
End Sub

Private Sub txtARROvzHgt_GotFocus()
    With txtARROvzHgt
        .BackColor = vbInfoBackground
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub

Private Sub txtARROvzHgt_Validate(Cancel As Boolean)
    Cancel = Not IsNumeric("0" & txtARROvzHgt)
End Sub

Private Sub txtARROvzLen_GotFocus()
    With txtARROvzLen
        .BackColor = vbInfoBackground
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub

Private Sub txtARROvzLen_Validate(Cancel As Boolean)
    Cancel = Not IsNumeric("0" & txtARROvzLen)
End Sub

Private Sub txtARROvzWid_GotFocus()
    With txtARROvzWid
        .BackColor = vbInfoBackground
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub

Private Sub txtARROvzWid_Validate(Cancel As Boolean)
    Cancel = Not IsNumeric("0" & txtARROvzWid)
End Sub

Private Sub txtARRUOM_GotFocus()
    With txtARRUOM
        .BackColor = vbInfoBackground
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub

Private Sub txtARRUOM_Validate(Cancel As Boolean)
    Cancel = (txtARRUOM <> "I") And (txtARRUOM <> "C")
End Sub

Private Sub txtChkAmt_Change(Index As Integer)
  If IsNumeric(txtChkAmt(Index).Text) Then
    Call lzComputePay
  End If
End Sub

Private Sub txtChkAmt_GotFocus(Index As Integer)
    With txtChkAmt(Index)
        .BackColor = vbInfoBackground
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub

Private Sub txtChkAmt_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyEscape
            SendKeys ("+{TAB}")
            KeyAscii = 0
        Case vbKeyReturn
            If Index < 4 Then
                If (Trim(txtChkAmt(Index)) & Trim(txtChkAmt(Index + 1)) = "") Then
                    cmdSave.SetFocus
                Else
                    SendKeys ("{TAB}")
                    KeyAscii = 0
                End If
            Else
                SendKeys ("{TAB}")
                KeyAscii = 0
            End If
        Case Else
    End Select
End Sub

Private Sub txtChkAmt_LostFocus(Index As Integer)
    txtChkAmt(Index).BackColor = vbWindowBackground
End Sub

Private Sub txtChkAmt_Validate(Index As Integer, Cancel As Boolean)
Dim n As Integer
Dim curTot As Currency
    Cancel = Not IsNumeric("0" & txtChkAmt(Index))
    If Not Cancel Then
        curTot = 0
        For n = 0 To 4
            curTot = curTot + CCur("0" & txtChkAmt(n))
        Next
        lblChkTot = Format(curTot, "##,###,##0.00")
        If Trim(txtChkAmt(Index)) <> "" Then
            txtChkAmt(Index).BackColor = vbWindowBackground
            txtChkNo(Index).SetFocus
        End If
    End If
End Sub

Private Sub txtChkBank_GotFocus(Index As Integer)
    With txtChkBank(Index)
        .BackColor = vbInfoBackground
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub

Private Sub txtChkBank_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyEscape
            SendKeys ("+{TAB}")
            KeyAscii = 0
        Case vbKeyReturn
            SendKeys ("{TAB}")
            KeyAscii = 0
        Case Else
    End Select
End Sub

Private Sub txtChkBank_LostFocus(Index As Integer)
    txtChkBank(Index).BackColor = vbWindowBackground
End Sub

Private Sub txtChkBank_Validate(Index As Integer, Cancel As Boolean)
    Cancel = (Trim(txtChkBank(Index).Text) = "") And _
             (Trim(txtChkAmt(Index).Text) <> "")
    If Cancel Then
        MsgBox "Check bank code required.", vbExclamation
        With txtChkNo(Index)
            .BackColor = vbInfoBackground
            .SelStart = 0
            .SelLength = .MaxLength
        End With
    Else
        If Trim(txtChkAmt(Index)) <> "" Then
            txtChkNo(Index).BackColor = vbWindowBackground
            txtChkBank(Index).SetFocus
        End If
    End If
End Sub

Private Sub txtChkNo_GotFocus(Index As Integer)
    With txtChkNo(Index)
        .BackColor = vbInfoBackground
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub

Private Sub txtChkNo_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyEscape
            SendKeys ("+{TAB}")
            KeyAscii = 0
        Case vbKeyReturn
            SendKeys ("{TAB}")
            KeyAscii = 0
        Case Else
    End Select
End Sub

Private Sub txtChkNo_LostFocus(Index As Integer)
    txtChkNo(Index).BackColor = vbWindowBackground
End Sub

Private Sub txtChkNo_Validate(Index As Integer, Cancel As Boolean)
    Cancel = (Trim(txtChkNo(Index).Text) = "") And _
             (Trim(txtChkAmt(Index).Text) <> "")
    If Cancel Then
        MsgBox "Check number required.", vbExclamation
        With txtChkNo(Index)
            .BackColor = vbInfoBackground
            .SelStart = 0
            .SelLength = .MaxLength
        End With
    Else
        If Trim(txtChkAmt(Index)) <> "" Then
            txtChkNo(Index).BackColor = vbWindowBackground
            txtChkBank(Index).SetFocus
        End If
    End If
End Sub

Private Sub txtCshAmt_Change()
  If IsNumeric(txtCshAmt.Text) Then
    Call lzComputePay
  End If
End Sub

Private Sub txtCshAmt_GotFocus()
    With txtCshAmt
        .BackColor = vbInfoBackground
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub

Private Sub txtCshAmt_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyEscape
            SendKeys ("+{TAB}")
            KeyAscii = 0
        Case vbKeyReturn
            SendKeys ("{TAB}")
            KeyAscii = 0
        Case Else
    End Select
End Sub

Private Sub txtCshAmt_LostFocus()
    txtCshAmt.BackColor = vbWindowBackground
End Sub

Private Sub txtCshAmt_Validate(Cancel As Boolean)
    Cancel = Not IsNumeric("0" & txtCshAmt)
End Sub

Private Sub txtCusName_GotFocus()
    With txtCusName
        .Text = Trim(.Text)
        .BackColor = vbInfoBackground
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtCusName_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyEscape
            Call lzInitialize
        Case vbKeyReturn
            SendKeys ("{TAB}")
            KeyAscii = 0
        Case Else
    End Select
End Sub

Private Sub txtCusName_LostFocus()
    txtCusName.BackColor = vbWindowBackground
    txtCusName = UCase(txtCusName)
End Sub

Private Sub lzPopulateDangerClass()
    With cboDanger
        .AddItem " " & Chr(124) & " None"
        .AddItem "1" & Chr(124) & " Explosives"
        .AddItem "2" & Chr(124) & " Gases"
        .AddItem "3" & Chr(124) & " Inflammable Liquid"
        .AddItem "4" & Chr(124) & " Inflammable Solids"
        .AddItem "5" & Chr(124) & " Oxidizing Agents/Organic Peroxides"
        .AddItem "6" & Chr(124) & " Poisonous(toxic) and Infectious Substances"
        .AddItem "7" & Chr(124) & " Radioactive Substances"
        .AddItem "8" & Chr(124) & " Corrosives"
        .AddItem "9" & Chr(124) & " Miscellaneous Dangerous Substances"
    End With
End Sub

Private Sub lzCustomerUnderG()
    frmCustPick.Show 1
    If gsCusCode <> "" Then
        vCusCodeUnderG = gsCusCode
        txtCusName = gsCusName
        chkVAT.SetFocus
    Else
        vCusCodeUnderG = Space(6)
        txtCusName.Enabled = True
    End If
End Sub

Private Sub lzInitialize()
Dim n As Integer
    bVAT = False: bWTax = False: chkVAT.Value = 0: chkWTax.Value = 0
    bNewCCR = True: chkNewCCR.Value = 1
    bUnderG = False: chkGuarantee.Value = 0
    nCCRCounter = 0: bArrImp = False: bStoImp = False
    With grdCCRTran
        .Enabled = False
        For n = 7 To .Cols - 1
            .ColWidth(n) = 0
        Next n
        .Rows = 1
        .AddItem ""
        .TextMatrix(.Row, enCounter) = "**"
    End With
    
    txtRemark = ""
    
    fraPayment.Enabled = False
    lblAmtDue = ""
    txtCshAmt = lblAmtDue
    For n = 0 To 4
        txtChkAmt(n) = ""
        txtChkNo(n) = Space(txtChkNo(n).MaxLength)
        txtChkBank(n) = Space(txtChkBank(n).MaxLength)
    Next n
    lblChkTot = ""
    lblChange = ""
    cmdSave.Enabled = False
    
    Call lzEnableTab
    txtCusName = ""

End Sub

Private Sub txtARROvzHgt_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyEscape
            SendKeys ("+{TAB}")
            KeyAscii = 0
        Case vbKeyReturn
            SendKeys ("{TAB}")
            KeyAscii = 0
        Case Else
    End Select
End Sub

Private Sub txtARROvzHgt_LostFocus()
    txtARROvzHgt.BackColor = vbWindowBackground
End Sub

Private Sub txtARROvzLen_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyEscape
            SendKeys ("+{TAB}")
            KeyAscii = 0
        Case vbKeyReturn
            SendKeys ("{TAB}")
            KeyAscii = 0
        Case Else
    End Select
End Sub

Private Sub txtARROvzLen_LostFocus()
    txtARROvzLen.BackColor = vbWindowBackground
End Sub

Private Sub txtARROvzWid_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyEscape
            SendKeys ("+{TAB}")
            KeyAscii = 0
        Case vbKeyReturn
            SendKeys ("{TAB}")
            KeyAscii = 0
        Case Else
    End Select
End Sub

Private Sub txtARROvzWid_LostFocus()
    txtARROvzWid.BackColor = vbWindowBackground
End Sub

Private Sub txtMscCCRNo_GotFocus()
    With txtMscCCRNo
        .BackColor = vbInfoBackground
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtMscCCRNo_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyEscape
            chkGuarantee.SetFocus
            KeyAscii = 0
        Case vbKeyReturn
            SendKeys ("{TAB}")
            KeyAscii = 0
        Case Else
    End Select
End Sub

Private Sub txtMscCCRNo_LostFocus()
    txtMscCCRNo.BackColor = vbWindowBackground
    Call lzGetCYMMsc(txtMscCCRNo)
End Sub

Private Sub txtMscContNo_GotFocus()
    With txtMscContNo
        .BackColor = vbInfoBackground
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtMscContNo_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyEscape
            SendKeys ("+{TAB}")
            KeyAscii = 0
        Case vbKeyReturn
            SendKeys ("{TAB}")
            KeyAscii = 0
        Case Else
    End Select
End Sub

Private Sub txtMscContNo_LostFocus()
    txtMscContNo.BackColor = vbWindowBackground
End Sub

Private Sub txtMscContSz_GotFocus()
    With txtMscContSz
        .BackColor = vbInfoBackground
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtMscContSz_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyEscape
            SendKeys ("+{TAB}")
            KeyAscii = 0
            bEscaped = True
        Case vbKeyReturn
            SendKeys ("{TAB}")
            KeyAscii = 0
        Case Else
    End Select
End Sub

Private Sub txtMscContSz_LostFocus()
    txtMscContSz.BackColor = vbWindowBackground
    If bEscaped Then
        bEscaped = False
    Else
        Call lzShowRate
    End If
End Sub

Private Sub txtMscQty_GotFocus()
    With txtMscQty
        .Text = Trim(.Text)
        .BackColor = vbInfoBackground
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtMscQty_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyEscape
            SendKeys ("+{TAB}")
            KeyAscii = 0
        Case vbKeyReturn
            KeyAscii = 0
            txtRemark.SetFocus
        Case Else
    End Select
End Sub

Private Sub txtMscQty_LostFocus()
    txtMscQty.BackColor = vbWindowBackground
    lblMscAmount = Format(CCur("0" & lblMscRateAmt) * CCur("0" & txtMscQty), "###,##0.00")
End Sub

Private Sub txtMscQty_Validate(Cancel As Boolean)
    Cancel = Not IsNumeric("0" & txtMscQty)
End Sub

Private Sub txtMscRateCode_GotFocus()
    With txtMscRateCode
        .BackColor = vbInfoBackground
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtMSCRateCode_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 Then
        Select Case KeyCode
            Case vbKeyF4
                frmCYRate.Show 1
                txtMscRateCode = vRateCode
                txtMscContSz = vRateSz
                Call lzShowRate
            Case Else
        End Select
    End If
End Sub

Private Sub txtMscRateCode_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyEscape
            SendKeys ("+{TAB}")
            KeyAscii = 0
        Case vbKeyReturn
            SendKeys ("{TAB}")
            KeyAscii = 0
        Case Else
    End Select
End Sub

Private Sub txtMscRateCode_LostFocus()
txtMscRateCode.BackColor = vbWindowBackground
End Sub

Private Sub txtOthAmount_GotFocus()
    With txtOthAmount
        .BackColor = vbInfoBackground
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub

Private Sub txtOthAmount_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyEscape
            SendKeys ("+{TAB}")
            KeyAscii = 0
        Case vbKeyReturn
            KeyAscii = 0
            txtRemark.SetFocus
        Case Else
    End Select
End Sub

Private Sub txtOthAmount_LostFocus()
    txtOthAmount.BackColor = vbWindowBackground
End Sub

Private Sub txtOthAmount_Validate(Cancel As Boolean)
    Cancel = Not IsNumeric("0" & txtOthAmount)
End Sub

Private Sub txtOTHCCRNo_GotFocus()
    With txtOTHCCRNo
        .BackColor = vbInfoBackground
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub

Private Sub txtOTHCCRNo_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyEscape
            chkGuarantee.SetFocus
            KeyAscii = 0
        Case vbKeyReturn
            SendKeys ("{TAB}")
            KeyAscii = 0
        Case Else
    End Select
End Sub

Private Sub txtOTHCCRNo_LostFocus()
    txtOTHCCRNo.BackColor = vbWindowBackground
End Sub

Private Sub txtOthContNo_GotFocus()
    With txtOthContNo
        .BackColor = vbInfoBackground
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub

Private Sub txtOthContNo_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyEscape
            SendKeys ("+{TAB}")
            KeyAscii = 0
        Case vbKeyReturn
            SendKeys ("{TAB}")
            KeyAscii = 0
        Case Else
    End Select
End Sub

Private Sub txtOthContNo_LostFocus()
    txtOthContNo.BackColor = vbWindowBackground
End Sub

Private Sub txtOthContSz_GotFocus()
    With txtOthContSz
        .BackColor = vbInfoBackground
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub

Private Sub txtOthContSz_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyEscape
            SendKeys ("+{TAB}")
            KeyAscii = 0
        Case vbKeyReturn
            SendKeys ("{TAB}")
            KeyAscii = 0
        Case Else
    End Select
End Sub

Private Sub txtOthContSz_LostFocus()
    txtOthContSz.BackColor = vbWindowBackground
End Sub

Private Sub txtOthContSz_Validate(Cancel As Boolean)
    Cancel = InStr("20|40|45|  |", txtOthContSz & "|") = 0
    If Cancel Then MsgBox "Invalid container size. Please correct..."
End Sub

Private Sub txtOthEntryNo_GotFocus()
    With txtOthEntryNo
        .BackColor = vbInfoBackground
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub

Private Sub txtOthEntryNo_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyEscape
            SendKeys ("+{TAB}")
            KeyAscii = 0
        Case vbKeyReturn
            SendKeys ("{TAB}")
            KeyAscii = 0
        Case Else
    End Select
End Sub

Private Sub txtOthEntryNo_LostFocus()
    txtOthEntryNo.BackColor = vbWindowBackground
End Sub

Private Sub txtOthFulEmp_GotFocus()
    With txtOthFulEmp
        .BackColor = vbInfoBackground
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub

Private Sub txtOthFulEmp_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyEscape
            SendKeys ("+{TAB}")
            KeyAscii = 0
        Case vbKeyReturn
            SendKeys ("{TAB}")
            KeyAscii = 0
        Case Else
    End Select
End Sub

Private Sub txtOthFulEmp_LostFocus()
    txtOthFulEmp.BackColor = vbWindowBackground
End Sub

Private Sub txtOthFulEmp_Validate(Cancel As Boolean)
    If Trim(txtOthFulEmp) <> "" Then
        Cancel = (txtOthFulEmp <> "F") And (txtOthFulEmp <> "E")
    End If
End Sub

Private Sub txtOthOvzHgt_GotFocus()
    With txtOthOvzHgt
        .BackColor = vbInfoBackground
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub

Private Sub txtOthOvzHgt_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyEscape
            SendKeys ("+{TAB}")
            KeyAscii = 0
        Case vbKeyReturn
            SendKeys ("{TAB}")
            KeyAscii = 0
        Case Else
    End Select
End Sub

Private Sub txtOthOvzHgt_LostFocus()
    txtOthOvzHgt.BackColor = vbWindowBackground
End Sub

Private Sub txtOthOvzHgt_Validate(Cancel As Boolean)
    Cancel = Not IsNumeric("0" & txtOthOvzHgt)
End Sub

Private Sub txtOthOvzLen_GotFocus()
    With txtOthOvzLen
        .BackColor = vbInfoBackground
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub

Private Sub txtOthOvzLen_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyEscape
            SendKeys ("+{TAB}")
            KeyAscii = 0
        Case vbKeyReturn
            SendKeys ("{TAB}")
            KeyAscii = 0
        Case Else
    End Select
End Sub

Private Sub txtOthOvzLen_LostFocus()
    txtOthOvzLen.BackColor = vbWindowBackground
End Sub

Private Sub txtOthOvzLen_Validate(Cancel As Boolean)
    Cancel = Not IsNumeric("0" & txtOthOvzLen)
End Sub

Private Sub txtOthOvzWid_GotFocus()
    With txtOthOvzWid
        .BackColor = vbInfoBackground
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub

Private Sub txtOthOvzWid_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyEscape
            SendKeys ("+{TAB}")
            KeyAscii = 0
        Case vbKeyReturn
            SendKeys ("{TAB}")
            KeyAscii = 0
        Case Else
    End Select
End Sub

Private Sub txtOthOvzWid_LostFocus()
    txtOthOvzWid.BackColor = vbWindowBackground
End Sub

Private Sub txtOthOvzWid_Validate(Cancel As Boolean)
    Cancel = Not IsNumeric("0" & txtOthOvzWid)
End Sub

Private Sub txtOthRegNo_GotFocus()
    With txtOthRegNo
        .BackColor = vbInfoBackground
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub

Private Sub txtOthRegNo_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyEscape
            SendKeys ("+{TAB}")
            KeyAscii = 0
        Case vbKeyReturn
            SendKeys ("{TAB}")
            KeyAscii = 0
        Case Else
    End Select
End Sub

Private Sub txtOthRegNo_LostFocus()
    txtOthRegNo.BackColor = vbWindowBackground
End Sub

Private Sub txtOthUOM_GotFocus()
    With txtOthUOM
        .BackColor = vbInfoBackground
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub

Private Sub txtOthUOM_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyEscape
            SendKeys ("+{TAB}")
            KeyAscii = 0
        Case vbKeyReturn
            SendKeys ("{TAB}")
            KeyAscii = 0
        Case Else
    End Select
End Sub

Private Sub txtOthUOM_LostFocus()
    txtOthUOM.BackColor = vbWindowBackground
    Call lzOthOversize
End Sub

Private Sub txtOthUOM_Validate(Cancel As Boolean)
    Cancel = (txtOthUOM <> "C") And (txtOthUOM <> "I")
End Sub

Private Sub txtOthVessel_GotFocus()
    With txtOthVessel
        .BackColor = vbInfoBackground
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub

Private Sub txtOthVessel_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyEscape
            SendKeys ("+{TAB}")
            KeyAscii = 0
        Case vbKeyReturn
            SendKeys ("{TAB}")
            KeyAscii = 0
        Case Else
    End Select
End Sub

Private Sub txtOthVessel_LostFocus()
    txtOthVessel.BackColor = vbWindowBackground
    txtOthVessel = UCase(txtOthVessel)
End Sub

Private Sub txtRemark_GotFocus()
    With txtRemark
        If txtMscRateCode = "WEIGHT" Then
            .Text = "WEIGHING"
        Else
            Select Case tabTran.Tab
                Case 0
                    .Text = "ADD'L ARRASTRE"
                Case 1
                    .Text = "STORAGE UP TO " & Format(txtStoExtDate, "YYYY/MM/DD")
                Case 2
                    .Text = "REEFER UP TO "
                Case 3
                    .Text = "SHUTOUT"
                Case 4
                    .Text = "EQUIPMENT RENTAL / MISCELLANEOUS"
                Case 5
                    .Text = "OTHER SPECIAL SERVICE"
                Case Else
                    .Text = Space(.MaxLength)
            End Select
        End If
        .BackColor = vbInfoBackground
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub

Private Sub txtRemark_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyEscape
            tabTran.SetFocus
            KeyAscii = 0
        Case vbKeyReturn
            SendKeys "{TAB}"
            KeyAscii = 0
        Case Else
    End Select
End Sub

Private Sub txtRemark_LostFocus()
    txtRemark.BackColor = vbWindowBackground
    txtRemark = UCase(txtRemark)
End Sub

Private Sub txtRFRCCRNo_GotFocus()
    With txtRFRCCRNo
        .BackColor = vbInfoBackground
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub

Private Sub txtRFRCCRNo_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyEscape
            chkGuarantee.SetFocus
            KeyAscii = 0
        Case vbKeyReturn
            SendKeys ("{TAB}")
            KeyAscii = 0
        Case Else
    End Select
End Sub

Private Sub txtRFRCCRNo_LostFocus()
    txtRFRCCRNo.BackColor = vbWindowBackground
    Call lzGetCYMRfr(txtRFRCCRNo)
    If txtRfrPlugInDate = cEmptyRfrDate Then txtRfrPlugInDate.SetFocus
End Sub

Private Sub txtRfrContNo_GotFocus()
    With txtRfrContNo
        .BackColor = vbInfoBackground
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub

Private Sub txtRfrContNo_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyEscape
            SendKeys ("+{TAB}")
            KeyAscii = 0
        Case vbKeyReturn
            SendKeys ("{TAB}")
            KeyAscii = 0
        Case Else
    End Select
End Sub

Private Sub txtRfrContNo_LostFocus()
    txtRfrContNo.BackColor = vbWindowBackground
End Sub

Private Sub txtRfrContSz_GotFocus()
    With txtRfrContSz
        .BackColor = vbInfoBackground
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub

Private Sub txtRfrContSz_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyEscape
            SendKeys ("+{TAB}")
            KeyAscii = 0
        Case vbKeyReturn
            SendKeys ("{TAB}")
            KeyAscii = 0
        Case Else
    End Select
End Sub

Private Sub txtRfrContSz_LostFocus()
    txtRfrContSz.BackColor = vbWindowBackground
End Sub

Private Sub txtRfrEntryNo_GotFocus()
    With txtRfrEntryNo
        .BackColor = vbInfoBackground
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub

Private Sub txtRfrEntryNo_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyEscape
            SendKeys ("+{TAB}")
            KeyAscii = 0
        Case vbKeyReturn
            SendKeys ("{TAB}")
            KeyAscii = 0
        Case Else
    End Select
End Sub

Private Sub txtRfrEntryNo_LostFocus()
    txtRfrEntryNo.BackColor = vbWindowBackground
End Sub

Private Sub txtRfrExtDate_GotFocus()
    With txtRfrExtDate
        .BackColor = vbInfoBackground
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub

Private Sub txtRfrExtDate_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyEscape
            bEscaped = True
            SendKeys ("+{TAB}")
            KeyAscii = 0
        Case vbKeyReturn
            If (txtRfrPlugInDate <> cEmptyRfrDate) And (txtRfrExtDate <> cEmptyRfrDate) Then
                txtRemark.SetFocus
            End If
        Case Else
    End Select
End Sub

Private Sub txtRfrExtDate_LostFocus()
    txtRfrExtDate.BackColor = vbWindowBackground
    If bEscaped Then
'        bEscaped = False
'        SendKeys ("+{TAB}")
    Else
        Call txtRfrExtDate_Validate(True)
    End If
End Sub

Private Sub txtRfrExtDate_Validate(Cancel As Boolean)
Dim vHrs As Long
    
    Cancel = True
    If bEscaped Then
        bEscaped = False
        Cancel = False
        SendKeys ("+{TAB}")
        Exit Sub
    End If
    
    If (txtRfrExtDate <> cEmptyRfrDate) And Not IsDate(txtRfrExtDate) Then
        MsgBox "Invalid date value. Please correct."
        txtRfrExtDate.SetFocus
    Else
        If (lblRfrValidUntil <> cEmptyRfrDate) And IsDate(txtRfrExtDate) Then
            If (DateDiff("n", CDate(lblRfrValidUntil), CDate(txtRfrExtDate)) < 1) Or _
               (DateDiff("n", gzGetSysDate(), CDate(txtRfrExtDate)) < 0) Then
                MsgBox "Should be greater than system date/time.  Please correct..."
                txtRfrExtDate.SetFocus
            End If
        End If
        If (lblRfrValidUntil <> cEmptyRfrDate) And (txtRfrExtDate <> cEmptyRfrDate) _
            And IsDate(lblRfrValidUntil) And IsDate(txtRfrExtDate) Then
            Cancel = True
        End If
    End If

    vHrs = DateDiff("h", CDate(lblRfrValidUntil), CDate(txtRfrExtDate))
    If vHrs > 0 Then
        lblRfrHrs = vHrs & " hrs"
    Else
        lblRfrHrs = ""
    End If

End Sub

Private Sub txtRfrPlugInDate_GotFocus()
    With txtRfrPlugInDate
        .BackColor = vbInfoBackground
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub

Private Sub txtRfrPlugInDate_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyEscape
            SendKeys ("+{TAB}")
            KeyAscii = 0
        Case vbKeyReturn
            SendKeys ("{TAB}")
            KeyAscii = 0
        Case Else
    End Select
End Sub

Private Sub txtRfrPlugInDate_LostFocus()
    txtRfrPlugInDate.BackColor = vbWindowBackground
    If lblRfrValidUntil = cEmptyRfrDate Then lblRfrValidUntil = txtRfrPlugInDate
End Sub

Private Sub txtRfrPlugInDate_Validate(Cancel As Boolean)
    Cancel = True
    If (txtRfrPlugInDate <> cEmptyRfrDate) And Not IsDate(txtRfrPlugInDate) Then
        MsgBox "Invalid date value. Please correct."
    Else
        If (txtRfrExtDate <> cEmptyRfrDate) And (txtRfrPlugInDate <> cEmptyRfrDate) Then
            If CDate(txtRfrPlugInDate) > CDate(txtRfrExtDate) Then
                MsgBox "Should be earlier than extension date.  Please correct..."
            Else
                Cancel = False
            End If
        Else
            Cancel = False
        End If
    End If

End Sub

Private Sub txtRfrRegNo_GotFocus()
    With txtRfrRegNo
        .BackColor = vbInfoBackground
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub

Private Sub txtRfrRegNo_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyEscape
            SendKeys ("+{TAB}")
            KeyAscii = 0
        Case vbKeyReturn
            SendKeys ("{TAB}")
            KeyAscii = 0
        Case Else
    End Select
End Sub

Private Sub txtRfrRegNo_LostFocus()
    txtRfrRegNo.BackColor = vbWindowBackground
End Sub

Private Sub txtSOCCRNo_GotFocus()
    With txtSOCCRNo
        .BackColor = vbInfoBackground
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub

Private Sub txtSOCCRNo_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyEscape
            chkGuarantee.SetFocus
            KeyAscii = 0
        Case vbKeyReturn
            SendKeys ("{TAB}")
            KeyAscii = 0
        Case Else
    End Select
End Sub

Private Sub txtSOCCRNo_LostFocus()
    txtSOCCRNo.BackColor = vbWindowBackground
End Sub

Private Sub txtSOContNo_GotFocus()
    With txtSOContNo
        .BackColor = vbInfoBackground
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub

Private Sub txtSOContNo_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyEscape
            tabTran.SetFocus
        Case vbKeyReturn
            SendKeys ("{TAB}")
            KeyAscii = 0
        Case Else
    End Select
End Sub

Private Sub txtSOContNo_LostFocus()
Dim w As New CWaitCursor
    txtSOContNo.BackColor = vbWindowBackground
    w.Restore
End Sub

Private Sub txtARRUOM_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyEscape
            SendKeys ("+{TAB}")
            KeyAscii = 0
        Case vbKeyReturn
            SendKeys ("{TAB}")
            KeyAscii = 0
        Case Else
    End Select
End Sub

Private Sub txtARRUOM_LostFocus()
    txtARRUOM.BackColor = vbWindowBackground
    txtARRUOM = UCase(txtARRUOM)
    Call lzArrOversize
End Sub

Private Sub lzGetCYMArr(ByVal pGatepass As String)
Dim cmd As ADODB.Command
Dim vOvrLen, vOvrWid, vOvrHgt As Long
Dim w As New CWaitCursor
    
    ' create command
    Set cmd = New ADODB.Command
    With cmd
        Set .ActiveConnection = gcnnBilling
        .CommandText = "upnew_getcymarrastre"
        .CommandType = adCmdStoredProc
    
        ' set parameters then execute
        .Parameters(0).Direction = adParamReturnValue
        .Parameters(1).Type = adInteger
        .Parameters(1).Value = CLng(pGatepass)
        .Parameters(1).Direction = adParamInput
        .Parameters(2).Type = adChar
        .Parameters(2).Direction = adParamOutput
        .Parameters(3).Type = adSmallInt
        .Parameters(3).Direction = adParamOutput
        .Parameters(4).Type = adInteger
        .Parameters(4).Direction = adParamOutput
        .Parameters(5).Type = adChar
        .Parameters(5).Direction = adParamOutput
        .Parameters(6).Type = adCurrency
        .Parameters(6).Direction = adParamOutput
        .Parameters(7).Type = adInteger
        .Parameters(7).Direction = adParamOutput
        .Parameters(8).Type = adInteger
        .Parameters(8).Direction = adParamOutput
        .Parameters(9).Type = adInteger
        .Parameters(9).Direction = adParamOutput
       
        .Execute
        
        If .Parameters(0) = 1 Then
            txtARRContNo = .Parameters(2)
            txtARRContSz = .Parameters(3)
            lblARREntryNo = "" & .Parameters(4)
            lblARRRegNo = "" & .Parameters(5)
            lblArrPrevAmt = Format(.Parameters(6), "##,###,##0.00")
            
            If .Parameters(7) > 0 Then
                vOvrLen = .Parameters(7)
                vOvrWid = .Parameters(8)
                vOvrHgt = .Parameters(9)
                    
                If vOvrLen + vOvrWid + vOvrHgt > 0 Then
                    Select Case txtARRContSz
                        Case "20"
                            If vOvrLen <= 240 Then vOvrLen = vOvrLen + 240
                        Case "40"
                            If vOvrLen <= 480 Then vOvrLen = vOvrLen + 480
                        Case "45"
                            If vOvrLen <= 540 Then vOvrLen = vOvrLen + 540
                        Case Else
                            vOvrLen = 0
                    End Select
                    If vOvrWid <= 96 Then vOvrWid = vOvrWid + 96
                    If vOvrHgt <= 96 Then vOvrHgt = vOvrHgt + 96
                    
                    txtARROvzLen = vOvrLen
                    txtARROvzWid = vOvrWid
                    txtARROvzHgt = vOvrHgt
                    chkARROvz.Value = 1
                End If
                
            End If
            
        Else
            txtARRContNo = Space(txtARRContNo.MaxLength)
            txtARRContSz = Space(txtARRContSz.MaxLength)
            lblARREntryNo = ""
            lblARRRegNo = ""
            lblArrPrevAmt = ""
        
        End If
    
    End With
    
End Sub

Private Sub lzGetCYXArr(ByVal pGatepass As String)
Dim cmd As ADODB.Command
Dim vOvrLen, vOvrWid, vOvrHgt As Long
Dim w As New CWaitCursor
    
    ' create command
    Set cmd = New ADODB.Command
    With cmd
        Set .ActiveConnection = gcnnBilling
        .CommandText = "upnew_getcyxarrastre"
        .CommandType = adCmdStoredProc
    
        ' set parameters then execute
        .Parameters(0).Direction = adParamReturnValue
        .Parameters(1).Type = adInteger                 ' CCR number
        .Parameters(1).Value = CLng(pGatepass)
        .Parameters(1).Direction = adParamInput
        .Parameters(2).Type = adChar                    ' container number
        .Parameters(2).Direction = adParamOutput
        .Parameters(3).Type = adSmallInt                ' container size
        .Parameters(3).Direction = adParamOutput
        .Parameters(4).Type = adCurrency                ' arrastre amount
        .Parameters(4).Direction = adParamOutput
        .Parameters(5).Type = adInteger                 ' oversize length
        .Parameters(5).Direction = adParamOutput
        .Parameters(6).Type = adInteger                 ' oversize width
        .Parameters(6).Direction = adParamOutput
        .Parameters(7).Type = adInteger                 ' oversize height
        .Parameters(7).Direction = adParamOutput
       
        .Execute
        
        If .Parameters(0) = 1 Then
            txtARRContNo = .Parameters(2)
            txtARRContSz = .Parameters(3)
            lblArrPrevAmt = Format(.Parameters(4), "##,###,##0.00")
            
            If .Parameters(5) > 0 Then
                vOvrLen = .Parameters(5)
                vOvrWid = .Parameters(6)
                vOvrHgt = .Parameters(7)
                    
                If vOvrLen + vOvrWid + vOvrHgt > 0 Then
                    Select Case txtARRContSz
                        Case "20"
                            If vOvrLen <= 240 Then vOvrLen = vOvrLen + 240
                        Case "40"
                            If vOvrLen <= 480 Then vOvrLen = vOvrLen + 480
                        Case "45"
                            If vOvrLen <= 540 Then vOvrLen = vOvrLen + 540
                        Case Else
                            vOvrLen = 0
                    End Select
                    If vOvrWid <= 96 Then vOvrWid = vOvrWid + 96
                    If vOvrHgt <= 96 Then vOvrHgt = vOvrHgt + 96
                    
                    txtARROvzLen = vOvrLen
                    txtARROvzWid = vOvrWid
                    txtARROvzHgt = vOvrHgt
                    chkARROvz.Value = 1
                End If
                
            End If
            
        Else
            txtARRContNo = Space(txtARRContNo.MaxLength)
            txtARRContSz = Space(txtARRContSz.MaxLength)
            lblArrPrevAmt = ""
        
        End If
    
    End With
    
End Sub

Private Sub lzGetCYMMsc(ByVal pGatepass As String)
Dim cmd As ADODB.Command
    
    ' create command
    Set cmd = New ADODB.Command
    With cmd
        Set .ActiveConnection = gcnnBilling
        .CommandText = "up_getcymarrastre"
        .CommandType = adCmdStoredProc
    
        ' set parameters then execute
        .Parameters(0).Direction = adParamReturnValue
        .Parameters(1).Type = adInteger
        .Parameters(1).Value = CLng("0" & pGatepass)
        .Parameters(1).Direction = adParamInput
        .Parameters(2).Type = adChar
        .Parameters(2).Direction = adParamOutput
        .Parameters(3).Type = adSmallInt
        .Parameters(3).Direction = adParamOutput
        .Parameters(4).Type = adInteger
        .Parameters(4).Direction = adParamOutput
        .Parameters(5).Type = adChar
        .Parameters(5).Direction = adParamOutput
        .Parameters(6).Type = adCurrency
        .Parameters(6).Direction = adParamOutput
       
        .Execute
        
        If .Parameters(0) = 1 Then
            txtMscContNo = .Parameters(2)
            txtMscContSz = .Parameters(3)
        Else
            txtMscContNo = Space(txtMscContNo.MaxLength)
            txtMscContSz = Space(txtMscContSz.MaxLength)
        End If
    
    End With
    
End Sub

Private Sub lzClearArr()
    txtARRCCRNo = Space(txtARRCCRNo.MaxLength)
    txtARRContNo = Space(txtARRContNo.MaxLength)
    txtARRContSz = Space(txtARRContSz.MaxLength)
    lblARREntryNo = ""
    lblARRRegNo = ""
    chkARROvz.Value = 0
    txtARROvzLen = ""
    txtARROvzWid = ""
    txtARROvzHgt = ""
    txtARRUOM = "I"
    cboDanger.ListIndex = 0
    lblArrPrevAmt = ""
    optArrImpExp(0).SetFocus
End Sub

Private Sub lzClearSO()
    txtSOCCRNo = Space(txtSOCCRNo.MaxLength)
    txtSOContNo = Space(txtSOContNo.MaxLength)
    lblSOContSz = ""
    lblSOFulEmp = ""
    txtSOVessel = ""
    txtSOCCRNo.SetFocus
End Sub

Private Sub lzClearRfr()
    txtRFRCCRNo = Space(txtRFRCCRNo.MaxLength)
    txtRfrContNo = Space(txtRfrContNo.MaxLength): txtRfrContNo.Enabled = False
    txtRfrContSz = Space(txtRfrContSz.MaxLength): txtRfrContSz.Enabled = False
    txtRfrEntryNo = Space(txtRfrEntryNo.MaxLength): txtRfrEntryNo.Enabled = False
    txtRfrRegNo = Space(txtRfrRegNo.MaxLength): txtRfrRegNo.Enabled = False
    txtRfrPlugInDate = cEmptyRfrDate: txtRfrPlugInDate.Enabled = False
    lblRfrValidUntil = ""
    lblRfrPrevPay = ""
    txtRfrExtDate = cEmptyRfrDate
    txtRFRCCRNo.SetFocus
End Sub

Private Sub lzClearSto()
    txtSTOCCRNo = Space(txtSTOCCRNo.MaxLength)
    txtStoContNo = Space(txtStoContNo.MaxLength)
    txtStoContSz = Space(txtStoContSz.MaxLength)
    lblStoEntryNo = ""
    lblStoRegNo = ""
    lblStoValidUntil = ""
    txtStoExtDate = Format(gzGetSysDate(), "YYYY-MM-DD")
    chkStoOvz.Value = 0
    txtStoOvzLen = ""
    txtStoOvzWid = ""
    txtStoOvzHgt = ""
    txtStoUOM = "I"
    lblStoRevTon = ""
    lblStoPrevPay = ""
    optStoImpExp(0).SetFocus
End Sub

Private Sub lzClearMsc()
    bEscaped = False
    txtMscCCRNo = Space(txtMscCCRNo.MaxLength)
    txtMscContNo = Space(txtMscContNo.MaxLength)
    txtMscContSz = Space(txtMscContSz.MaxLength)
    txtMscRateCode = Space(txtMscRateCode.MaxLength)
    lblMScRateDesc = ""
    lblMscRateAmt = ""
    lblMscRateUOM = ""
    txtMscQty = Space(txtMscQty.MaxLength)
    lblMscAmount = ""
    txtMscCCRNo.SetFocus
End Sub

Private Sub lzClearOth()
    txtOTHCCRNo = Space(txtOTHCCRNo.MaxLength)
    txtOthContNo = Space(txtOthContNo.MaxLength)
    txtOthContSz = Space(txtOthContSz.MaxLength)
    txtOthFulEmp = Space(txtOthFulEmp.MaxLength)
    txtOthEntryNo = Space(txtOthEntryNo.MaxLength)
    txtOthRegNo = Space(txtOthRegNo.MaxLength)
    txtOthVessel = Space(txtOthVessel.MaxLength)
    chkOthOvz.Value = 0
    txtOthOvzLen = ""
    txtOthOvzWid = ""
    txtOthOvzHgt = ""
    txtOthUOM = "I"
    txtOthAmount = ""
    txtOTHCCRNo.SetFocus
End Sub

Private Sub lzGetUserInfo()
    lblCCRLastIssue = lzGetLastCCR
End Sub

Private Function lzGetLastCCR() As String
Dim cmd As ADODB.Command
    
    ' create command
    Set cmd = New ADODB.Command
    With cmd
        Set .ActiveConnection = gcnnBilling
        .CommandText = "up_getlastsplissued"
        .CommandType = adCmdStoredProc
    
        ' set parameters then execute
        '.Parameters(0).Direction = adParamReturnValue
        .Parameters(1).Type = adChar
        .Parameters(1).Value = gUserID
        .Parameters(1).Direction = adParamInput
        .Parameters(2).Type = adInteger
        .Parameters(2).Direction = adParamOutput
        .Parameters(3).Type = adDate
        .Parameters(3).Direction = adParamOutput
       
        .Execute
        
        If .Parameters(2) = 0 Then
            lzGetLastCCR = "NEW ALLOCATION"
        Else
            lzGetLastCCR = Format(.Parameters(2), "########") & " ON " & _
                           Format(.Parameters(3), "YYYY-MM-DD hh:mm")
        End If
    
    End With
    
End Function

Private Function lzGetRateInfo(ByVal pRTECDE As String, Optional ByVal pCNTSZE As String = "NIL") As Currency
Dim cmd As ADODB.Command
    
    ' create command
    Set cmd = New ADODB.Command
    With cmd
        Set .ActiveConnection = gcnnBilling
        .CommandText = "up_getcyrateinfo"
        .CommandType = adCmdStoredProc
    
        ' set parameters then execute
        '.Parameters(0).Direction = adParamReturnValue
        .Parameters(1).Type = adChar                    ' rate code
        .Parameters(1).Value = pRTECDE
        .Parameters(1).Direction = adParamInput
        .Parameters(2).Type = adChar                    ' container size
        .Parameters(2).Value = IIf(pCNTSZE = "NIL", "  ", pCNTSZE)
        .Parameters(2).Direction = adParamInput
        .Parameters(3).Type = adChar                    ' rate type
        .Parameters(3).Direction = adParamOutput
        .Parameters(4).Type = adChar                    ' rate description
        .Parameters(4).Direction = adParamOutput
        .Parameters(5).Type = adCurrency                ' rate amount
        .Parameters(5).Direction = adParamOutput
        .Parameters(6).Type = adChar                    ' unit of measure
        .Parameters(6).Direction = adParamOutput
       
        .Execute
        
        lzGetRateInfo = IIf(IsNull(.Parameters(5)), 0, .Parameters(5))
        
        vRateCode = pRTECDE
        vRateSz = pCNTSZE
        vRateDesc = "" & .Parameters(4)
        vRateAmount = lzGetRateInfo
        vRateUOM = "" & .Parameters(6)
    End With
End Function

Private Sub lzUpdateGridVAT()
Dim n, i As Integer
    n = grdCCRTran.Rows
    If n > 2 Then
        With grdCCRTran
            lblAmtDue = ""
            For i = 1 To (n - 2)
                If .TextMatrix(i, enRateCode) <> cVoid Then
                nAmount = CCur(.TextMatrix(i, enAmount))
                nTotalAmount = CCur(.TextMatrix(i, enTotalAmt))
                If bVAT Then
                    nVATAmount = nAmount * 0.1
                    .TextMatrix(i, enVATAmt) = Format(nVATAmount, "##,##0.00")
                    nTotalAmount = nTotalAmount + nVATAmount
                Else
                    nVATAmount = CCur("0" & .TextMatrix(i, enVATAmt))
                    nTotalAmount = nTotalAmount - nVATAmount
                    nVATAmount = 0
                    .TextMatrix(i, enVATAmt) = ""
                End If
                .TextMatrix(i, enTotalAmt) = Format(nTotalAmount, "##,##,##0.00")
                Call lzAddToTotal
                End If
            Next
        End With
    End If
End Sub

Private Sub lzUpdateGridWTax()
Dim n, i As Integer
    n = grdCCRTran.Rows
    If n > 2 Then
        With grdCCRTran
            lblAmtDue = ""
            For i = 1 To (n - 2)
                If .TextMatrix(i, enRateCode) <> cVoid Then
                    nAmount = CCur(.TextMatrix(i, enAmount))
                    nTotalAmount = CCur(.TextMatrix(i, enTotalAmt))
                    If bWTax Then
                        nWTaxAmount = nAmount * 0.02
                        .TextMatrix(i, enWTaxAmt) = Format(nWTaxAmount, "##,##0.00")
                        nTotalAmount = nTotalAmount - nWTaxAmount
                    Else
                        nWTaxAmount = CCur("0" & .TextMatrix(i, enWTaxAmt))
                        nTotalAmount = nTotalAmount + nWTaxAmount
                        nWTaxAmount = 0
                        .TextMatrix(i, enWTaxAmt) = ""
                    End If
                    .TextMatrix(i, enTotalAmt) = Format(nTotalAmount, "##,##,##0.00")
                    Call lzAddToTotal
                End If
            Next
        End With
    End If
End Sub

Private Sub lzAddToTotal()
Dim curTotal As Currency
    curTotal = CCur("0" & lblAmtDue)
    lblAmtDue = Format(curTotal + nTotalAmount, "#,###,##0.00")
    lblAmtDue.Refresh
End Sub

Private Sub lzComputeArr()
Dim curArrOvzAmt, curArrDanger As Currency
    ' compute
    cRateCode = IIf(bArrImp, "IMAR", "EXAR")
    nAmount = lzGetRateInfo(cRateCode, txtARRContSz)
    If bARROversize Then
        curArrOvzAmt = lzArrOversize()
        nAmount = nAmount + curArrOvzAmt
    End If
    curArrDanger = lzArrDanger(nAmount)
    nAmount = nAmount + curArrDanger
    nVATAmount = IIf(bVAT, nAmount * 0.1, 0)
    nWTaxAmount = IIf(bWTax, nAmount * 0.02, 0)
    nTotalAmount = nAmount + nVATAmount - nWTaxAmount
    nTotalAmount = nTotalAmount - CCur("0" & lblArrPrevAmt)
        
    nWTaxAmount = IIf(bWTax, nTotalAmount * 0.02, 0)
    
    nAmount = nTotalAmount
    If bVAT Then
        nAmount = nTotalAmount / 1.1
        nVATAmount = nTotalAmount - nAmount
    Else
        nVATAmount = 0
    End If
    If bWTax Then
        nWTaxAmount = nAmount - (nAmount / 1.02)
        nAmount = nAmount / 1.02
    Else
        nWTaxAmount = 0
    End If
        
    If nTotalAmount > 0 Then
        
        ' new ccr
        If bNewCCR Or (nCCRCounter > 7) Then
            bNewCCR = False
            nCCRCounter = 1
        Else
            nCCRCounter = nCCRCounter + 1
        End If
        
        With grdCCRTran
            nPtr = .Rows
            .AddItem nPtr
            .TextMatrix(nPtr - 1, enCounter) = nPtr - 1
            .TextMatrix(nPtr - 1, enRateCode) = cRateCode
            .TextMatrix(nPtr - 1, enAmount) = Format(nAmount, "###,##0.00")
            If nVATAmount > 0 Then .TextMatrix(nPtr - 1, enVATAmt) = Format(nVATAmount, "##,##0.00")
            If nWTaxAmount > 0 Then .TextMatrix(nPtr - 1, enWTaxAmt) = Format(nWTaxAmount, "##,##0.00")
            .TextMatrix(nPtr - 1, enTotalAmt) = Format(nTotalAmount, "##,##,##0.00")
            .TextMatrix(nPtr - 1, enCCRTag) = IIf(nCCRCounter = 1, "*", " ")
            
            .TextMatrix(nPtr - 1, enCCRNo) = txtARRCCRNo
            .TextMatrix(nPtr - 1, enContNo) = txtARRContNo
            .TextMatrix(nPtr - 1, enContSz) = txtARRContSz
            .TextMatrix(nPtr - 1, enEntryNo) = lblARREntryNo
            .TextMatrix(nPtr - 1, enRegNo) = lblARRRegNo
            If bARROversize Then
                .TextMatrix(nPtr - 1, enOvzLen) = txtARROvzLen
                .TextMatrix(nPtr - 1, enOvzWid) = txtARROvzWid
                .TextMatrix(nPtr - 1, enOvzHgt) = txtARROvzHgt
                .TextMatrix(nPtr - 1, enOvzUom) = txtARRUOM
                .TextMatrix(nPtr - 1, enRevTon) = lblArrRevTon
                .TextMatrix(nPtr - 1, enOvzAmt) = curArrOvzAmt
            End If
            If Left(cboDanger, 1) <> " " Then
                .TextMatrix(nPtr - 1, enDangerCode) = cboDanger
                .TextMatrix(nPtr - 1, enDangerAmt) = curArrDanger
            End If
            .TextMatrix(nPtr - 1, enRemark) = txtRemark
            
            .TextMatrix(nPtr, enCounter) = "**"
            .Row = nPtr
        End With
        Call lzAddToTotal
        Call lzClearArr
    Else
        MsgBox "No additional charges computed...", vbInformation
        tabTran.SetFocus
    End If
End Sub

Private Sub lzComputeSO()
    nPtr = grdCCRTran.Rows
    grdCCRTran.AddItem nPtr
    cRateCode = IIf(lblSOFulEmp = "F", "SOF", "SOE")
    nAmount = lzGetRateInfo(cRateCode)
    nVATAmount = IIf(bVAT, nAmount * 0.1, 0)
    nWTaxAmount = IIf(bWTax, nAmount * 0.1, 0)
    nTotalAmount = nAmount + nVATAmount - nWTaxAmount

    If nTotalAmount > 0 Then
        ' new ccr
        If bNewCCR Or (nCCRCounter > 7) Then
            bNewCCR = False
            nCCRCounter = 1
        Else
            nCCRCounter = nCCRCounter + 1
        End If
    
        With grdCCRTran
            .TextMatrix(nPtr - 1, enCounter) = nPtr - 1
            .TextMatrix(nPtr - 1, enRateCode) = cRateCode
            .TextMatrix(nPtr - 1, enAmount) = Format(nAmount, "###,##0.00")
            If nVATAmount > 0 Then .TextMatrix(nPtr - 1, enVATAmt) = Format(nVATAmount, "##,##0.00")
            If nWTaxAmount > 0 Then .TextMatrix(nPtr - 1, enWTaxAmt) = Format(nVATAmount, "##,##0.00")
            .TextMatrix(nPtr - 1, enTotalAmt) = Format(nTotalAmount, "##,##,##0.00")
            .TextMatrix(nPtr - 1, enCCRTag) = IIf(nCCRCounter = 1, "*", " ")
            
            .TextMatrix(nPtr - 1, enCCRNo) = txtSOCCRNo
            .TextMatrix(nPtr - 1, enContNo) = txtSOContNo
            .TextMatrix(nPtr - 1, enContSz) = lblSOContSz
            .TextMatrix(nPtr - 1, enFulEmp) = lblSOFulEmp
            .TextMatrix(nPtr - 1, enEntryNo) = lblARREntryNo
            .TextMatrix(nPtr - 1, enVessel) = txtSOVessel
            .TextMatrix(nPtr - 1, enRemark) = txtRemark
            
            .TextMatrix(nPtr, enCounter) = "**"
            .Row = nPtr
        End With
        Call lzAddToTotal
        Call lzClearSO
    Else
        MsgBox "No additional charges computed...", vbInformation
        txtSOCCRNo.SetFocus
    End If
End Sub

Private Function lzArrOversize() As Currency
Dim pLength, pWidth, pHeight As Single
    
    pLength = CSng(txtARROvzLen)
    pWidth = CSng(txtARROvzWid)
    pHeight = CSng(txtARROvzHgt)
    If txtARRUOM = "C" Then
       pLength = pLength / 2.54 ': If pLength <= 0 Then pLength = 1
       pWidth = pWidth / 2.54 ': If pWidth <= 0 Then pWidth = 1
       pHeight = pHeight / 2.54 ': If pHeight <= 0 Then pWidth = 1
    'Else
    '   If pLength <= 0 Then pLength = 1
    '   If pWidth <= 0 Then pWidth = 1
    '   If pHeight <= 0 Then pWidth = 1
    End If
    vRevTon = pLength * pWidth * pHeight / 1728 / 40
    Select Case txtARRContSz
        Case "20"
            If vRevTon >= cRTon20 Then vRevTon = vRevTon - cRTon20
        Case "40"
            If vRevTon >= cRTon40 Then vRevTon = vRevTon - cRTon40
        Case "45"
            If vRevTon >= cRTon45 Then vRevTon = vRevTon - cRTon45
        Case Else
            vRevTon = 0
    End Select
    vRevTon = Round(vRevTon, 2)
    If vRevTon > 0 Then
        lblArrRevTon = Format(vRevTon, "###,##0.00")
    Else
        lblArrRevTon = ""
    End If
    If bArrImp Then
        lzArrOversize = vRevTon * vRevTonRateArr
    Else
        lzArrOversize = vRevTon * vRevTonRateArrExp
    End If

End Function

Private Function lzArrDanger(ByVal pAmt As Currency) As Currency
Dim sDangerCode As String * 1
    sDangerCode = Left(cboDanger, 1)
    Select Case sDangerCode
        Case "1", "6", "8"
            lzArrDanger = pAmt * 0.5
        Case "2", "3", "4", "7"
            lzArrDanger = pAmt * 0.25
        Case "5", "9"
            lzArrDanger = pAmt * 0.1
        Case Else
            lzArrDanger = 0
    End Select
End Function

Private Sub txtSOVessel_GotFocus()
    With txtSOVessel
        .BackColor = vbInfoBackground
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub

Private Sub txtSOVessel_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyEscape
            SendKeys "+{TAB}"
            KeyAscii = 0
        Case vbKeyReturn
            txtRemark.SetFocus
            KeyAscii = 0
        Case Else
    End Select
End Sub

Private Sub txtSOVessel_LostFocus()
    txtSOVessel.BackColor = vbWindowBackground
    txtSOVessel = UCase(txtSOVessel)
End Sub

Private Sub txtSOVessel_Validate(Cancel As Boolean)
    txtSOVessel = UCase(txtSOVessel)
End Sub

Private Sub txtSTOCCRNo_GotFocus()
    With txtSTOCCRNo
        .BackColor = vbInfoBackground
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub

Private Sub txtSTOCCRNo_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyEscape
            SendKeys ("+{TAB}")
            KeyAscii = 0
        Case vbKeyReturn
            SendKeys ("{TAB}")
            KeyAscii = 0
        Case Else
    End Select
End Sub

Private Sub txtSTOCCRNo_LostFocus()
    txtSTOCCRNo.BackColor = vbWindowBackground
    If bStoImp And (Val(txtSTOCCRNo)) > 0 Then
        Call lzGetCYMSto(txtSTOCCRNo)
    End If
End Sub

Private Sub lzGetCYMSto(ByVal pGatepass As String)
Dim vOvrLen, vOvrWid, vOvrHgt As Long
Dim cmd As ADODB.Command
Dim w As New CWaitCursor
   
    ' create command
    Set cmd = New ADODB.Command
    With cmd
        Set .ActiveConnection = gcnnBilling
        .CommandText = "upnew_getcymstorage"
        .CommandType = adCmdStoredProc
    
        ' set parameters then execute
        .Parameters(0).Direction = adParamReturnValue       ' 1 if succesfull, else 0
        .Parameters(1).Type = adInteger
        .Parameters(1).Value = CLng(pGatepass)
        .Parameters(1).Direction = adParamInput             ' gatepass number
        .Parameters(2).Type = adChar
        .Parameters(2).Direction = adParamOutput            ' container number
        .Parameters(3).Type = adSmallInt
        .Parameters(3).Direction = adParamOutput            ' container size
        .Parameters(4).Type = adInteger
        .Parameters(4).Direction = adParamOutput            ' entry number
        .Parameters(5).Type = adChar
        .Parameters(5).Direction = adParamOutput            ' registry number
        .Parameters(6).Type = adDate
        .Parameters(6).Direction = adParamOutput            ' storage validity
        .Parameters(7).Type = adCurrency
        .Parameters(7).Direction = adParamOutput            ' storage amount paid
        .Parameters(8).Type = adInteger
        .Parameters(8).Direction = adParamOutput            ' oversize length
        .Parameters(9).Type = adInteger
        .Parameters(9).Direction = adParamOutput            ' oversize width
        .Parameters(10).Type = adInteger
        .Parameters(10).Direction = adParamOutput           ' oversize height
        .Parameters(11).Type = adInteger
        .Parameters(11).Direction = adParamOutput           ' storage days
       
        .Execute
        
        If .Parameters(0) = 1 Then
            txtStoContNo = .Parameters(2)
            txtStoContSz = .Parameters(3)
            lblStoEntryNo = "" & .Parameters(4)
            lblStoRegNo = "" & .Parameters(5)
            lblStoValidUntil = Format(.Parameters(6), "YYYY-MM-DD")
            lblStoPrevPay = Format(.Parameters(7), "##,###,##0.00")
            
            vCYMStoDay = .Parameters(11)
            If .Parameters(8) > 0 Then
                vOvrLen = .Parameters(8)
                vOvrWid = .Parameters(9)
                vOvrHgt = .Parameters(10)
                    
                If vOvrLen + vOvrWid + vOvrHgt > 0 Then
                    Select Case txtStoContSz
                        Case "20"
                            If vOvrLen < 240 Then vOvrLen = vOvrLen + 240
                        Case "40"
                            If vOvrLen < 480 Then vOvrLen = vOvrLen + 480
                        Case "45"
                            If vOvrLen < 540 Then vOvrLen = vOvrLen + 540
                        Case Else
                            vOvrLen = 0
                    End Select
                    If vOvrWid < 96 Then vOvrWid = vOvrWid + 96
                    If vOvrHgt < 96 Then vOvrHgt = vOvrHgt + 96
                    
                    txtStoOvzLen = vOvrLen
                    txtStoOvzWid = vOvrWid
                    txtStoOvzHgt = vOvrHgt
                    chkStoOvz.Value = 1
                End If
                vCYMStoDay = 0
            End If
            
        Else
            txtStoContNo = Space(txtStoContNo.MaxLength)
            txtStoContSz = Space(txtStoContSz.MaxLength)
            lblStoEntryNo = ""
            lblStoRegNo = ""
            lblStoValidUntil = ""
            lblStoPrevPay = ""
        End If
    End With
    Set cmd = Nothing
End Sub

Private Sub lzGetCYMRfr(ByVal pGatepass As String)
Dim cmd As ADODB.Command
    
    If Trim(txtRFRCCRNo) <> "" Then
        ' create command
        Set cmd = New ADODB.Command
        With cmd
            Set .ActiveConnection = gcnnBilling
            .CommandText = "up_getcymreefer"
            .CommandType = adCmdStoredProc
        
            ' set parameters then execute
            .Parameters(0).Direction = adParamReturnValue       ' 1 if succesfull, else 0
            .Parameters(1).Type = adInteger
            .Parameters(1).Value = CLng("0" & pGatepass)
            .Parameters(1).Direction = adParamInput             ' gatepass number
            .Parameters(2).Type = adChar
            .Parameters(2).Direction = adParamOutput            ' container number
            .Parameters(3).Type = adSmallInt
            .Parameters(3).Direction = adParamOutput            ' container size
            .Parameters(4).Type = adInteger
            .Parameters(4).Direction = adParamOutput            ' entry number
            .Parameters(5).Type = adChar
            .Parameters(5).Direction = adParamOutput            ' registry number
            .Parameters(6).Type = adDate
            .Parameters(6).Direction = adParamOutput            ' plugin date
            .Parameters(7).Type = adDate
            .Parameters(7).Direction = adParamOutput            ' valid until
            .Parameters(8).Type = adCurrency
            .Parameters(8).Direction = adParamOutput            ' reefer amount paid

            .Execute
            
            If .Parameters(0) = 1 Then
                txtRfrContNo = .Parameters(2): txtRfrContNo.Enabled = False
                txtRfrContSz = .Parameters(3): txtRfrContSz.Enabled = False
                txtRfrEntryNo = Trim(str(.Parameters(4))): txtRfrEntryNo.Enabled = False
                txtRfrRegNo = Trim(.Parameters(5)): txtRfrRegNo.Enabled = False
                If .Parameters(6) <> cNullDate Then
                    txtRfrPlugInDate = Format(.Parameters(6), "YYYY-MM-DD hh:mm")
                    txtRfrPlugInDate.Enabled = False
                Else
                    txtRfrPlugInDate = cEmptyRfrDate
                    txtRfrPlugInDate.Enabled = True
                End If
                If .Parameters(7) <> cNullDate Then
                    lblRfrValidUntil = Format(.Parameters(7), "YYYY-MM-DD hh:mm")
                Else
                    lblRfrValidUntil = txtRfrPlugInDate
                End If
                lblRfrPrevPay = Format(.Parameters(8), "##,###,##0.00")
                
                Exit Sub
            End If
        End With
        Set cmd = Nothing
    
    End If
    
    txtRfrContNo = Space(txtRfrContNo.MaxLength): txtRfrContNo.Enabled = True
    txtRfrContSz = Space(txtRfrContSz.MaxLength): txtRfrContSz.Enabled = True
    txtRfrEntryNo = Space(txtRfrEntryNo.MaxLength): txtRfrEntryNo.Enabled = True
    txtRfrRegNo = Space(txtRfrRegNo.MaxLength): txtRfrRegNo.Enabled = True
    txtRfrPlugInDate = cEmptyRfrDate: txtRfrPlugInDate.Enabled = True
    lblRfrValidUntil = ""
    lblRfrPrevPay = ""
    
End Sub

Private Sub txtStoContNo_GotFocus()
    With txtStoContNo
        .BackColor = vbInfoBackground
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub

Private Sub txtStoContNo_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyEscape
            SendKeys ("+{TAB}")
            KeyAscii = 0
        Case vbKeyReturn
            SendKeys ("{TAB}")
            KeyAscii = 0
        Case Else
    End Select
End Sub

Private Sub txtStoContNo_LostFocus()
    txtStoContNo.BackColor = vbWindowBackground
End Sub

Private Sub txtStoContSz_GotFocus()
    With txtStoContSz
        .BackColor = vbInfoBackground
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub

Private Sub txtStoContSz_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyEscape
            SendKeys ("+{TAB}")
            KeyAscii = 0
        Case vbKeyReturn
            SendKeys ("{TAB}")
            KeyAscii = 0
        Case Else
    End Select
End Sub

Private Sub txtStoContSz_LostFocus()
    txtStoContSz.BackColor = vbWindowBackground
End Sub

Private Sub txtStoContSz_Validate(Cancel As Boolean)
    Cancel = InStr("20|40|45|", txtStoContSz & "|") = 0
    If Cancel Then MsgBox "Invalid container size. Please correct..."
End Sub

Private Sub txtStoExtDate_GotFocus()
    With txtStoExtDate
        .BackColor = vbInfoBackground
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub

Private Sub txtStoExtDate_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyEscape
            SendKeys ("+{TAB}")
            KeyAscii = 0
        Case vbKeySpace
            SendKeys ("{F4}")
        Case vbKeyReturn
            If lzStoExtDateValid() Then txtRemark.SetFocus
        Case Else
    End Select
End Sub

Private Sub txtStoExtDate_LostFocus()
    txtStoExtDate.BackColor = vbWindowBackground
    Call txtStoExtDate_Validate(True)
End Sub

Private Sub txtStoExtDate_Validate(Cancel As Boolean)
    Cancel = Not lzStoExtDateValid()
End Sub

Private Sub txtStoOvzHgt_GotFocus()
    With txtStoOvzHgt
        .BackColor = vbInfoBackground
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub

Private Sub txtStoOvzHgt_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyEscape
            SendKeys ("+{TAB}")
            KeyAscii = 0
        Case vbKeyReturn
            SendKeys ("{TAB}")
            KeyAscii = 0
        Case Else
    End Select
End Sub

Private Sub txtStoOvzHgt_LostFocus()
    txtStoOvzHgt.BackColor = vbWindowBackground
End Sub

Private Sub txtStoOvzHgt_Validate(Cancel As Boolean)
    Cancel = Not IsNumeric("0" & txtStoOvzHgt)
End Sub

Private Sub txtStoOvzLen_GotFocus()
    With txtStoOvzLen
        .BackColor = vbInfoBackground
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub

Private Sub txtStoOvzLen_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyEscape
            SendKeys ("+{TAB}")
            KeyAscii = 0
        Case vbKeyReturn
            SendKeys ("{TAB}")
            KeyAscii = 0
        Case Else
    End Select
End Sub

Private Sub txtStoOvzLen_LostFocus()
    txtStoOvzLen.BackColor = vbWindowBackground
End Sub

Private Sub txtStoOvzLen_Validate(Cancel As Boolean)
    Cancel = Not IsNumeric("0" & txtStoOvzLen)
End Sub

Private Sub txtStoOvzWid_GotFocus()
    With txtStoOvzWid
        .BackColor = vbInfoBackground
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub

Private Sub txtStoOvzWid_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyEscape
            SendKeys ("+{TAB}")
            KeyAscii = 0
        Case vbKeyReturn
            SendKeys ("{TAB}")
            KeyAscii = 0
        Case Else
    End Select
End Sub

Private Sub txtStoOvzWid_LostFocus()
    txtStoOvzWid.BackColor = vbWindowBackground
End Sub

Private Sub txtStoOvzWid_Validate(Cancel As Boolean)
    Cancel = Not IsNumeric("0" & txtStoOvzWid)
End Sub

Private Sub txtStoUOM_GotFocus()
    With txtStoUOM
        .BackColor = vbInfoBackground
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub

Private Sub txtStoUOM_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyEscape
            SendKeys ("+{TAB}")
            KeyAscii = 0
        Case vbKeyReturn
            SendKeys ("{TAB}")
            KeyAscii = 0
        Case Else
    End Select
End Sub

Private Sub txtStoUOM_LostFocus()
    txtStoUOM.BackColor = vbWindowBackground
    Call lzStoOversize
End Sub

Private Sub lzComputeSto()
Dim curStoOvzAmt As Currency
    ' compute
    cRateCode = IIf(bStoImp, "IMST", "EXST")
    nAmount = lzGetRateInfo(cRateCode, txtStoContSz)
    vStoDay = DateDiff("d", CDate(lblStoValidUntil), CDate(txtStoExtDate))
    nAmount = nAmount * vStoDay
    If bStoOversize Then
        curStoOvzAmt = lzStoOversize()
        nAmount = nAmount + curStoOvzAmt
    End If
    nVATAmount = IIf(bVAT, nAmount * 0.1, 0)
    nWTaxAmount = IIf(bWTax, nAmount * 0.02, 0)
    nTotalAmount = nAmount + nVATAmount - nWTaxAmount
    'nTotalAmount = nTotalAmount - CCur("0" & lblStoPrevPay)
        
    If nTotalAmount > 0 Then
        ' new ccr
        If bNewCCR Or (nCCRCounter > 7) Then
            bNewCCR = False
            nCCRCounter = 1
        Else
            nCCRCounter = nCCRCounter + 1
        End If
        
        With grdCCRTran
            nPtr = .Rows
            .AddItem nPtr
            .TextMatrix(nPtr - 1, enCounter) = nPtr - 1
            .TextMatrix(nPtr - 1, enRateCode) = cRateCode
            .TextMatrix(nPtr - 1, enAmount) = Format(nAmount, "###,##0.00")
            If nVATAmount > 0 Then .TextMatrix(nPtr - 1, enVATAmt) = Format(nVATAmount, "##,##0.00")
            If nWTaxAmount > 0 Then .TextMatrix(nPtr - 1, enWTaxAmt) = Format(nWTaxAmount, "##,##0.00")
            .TextMatrix(nPtr - 1, enTotalAmt) = Format(nTotalAmount, "##,##,##0.00")
            .TextMatrix(nPtr - 1, enCCRTag) = IIf(nCCRCounter = 1, "*", " ")
            
            .TextMatrix(nPtr - 1, enCCRNo) = txtSTOCCRNo
            .TextMatrix(nPtr - 1, enContNo) = txtStoContNo
            .TextMatrix(nPtr - 1, enContSz) = txtStoContSz
            .TextMatrix(nPtr - 1, enEntryNo) = lblStoEntryNo
            .TextMatrix(nPtr - 1, enRegNo) = lblStoRegNo
            If bStoOversize Then
                .TextMatrix(nPtr - 1, enOvzLen) = txtStoOvzLen
                .TextMatrix(nPtr - 1, enOvzWid) = txtStoOvzWid
                .TextMatrix(nPtr - 1, enOvzHgt) = txtStoOvzHgt
                .TextMatrix(nPtr - 1, enOvzUom) = txtStoUOM
                .TextMatrix(nPtr - 1, enRevTon) = lblStoRevTon
                .TextMatrix(nPtr - 1, enOvzAmt) = curStoOvzAmt
            End If
            .TextMatrix(nPtr - 1, enStoValidUntil) = txtStoExtDate
            .TextMatrix(nPtr - 1, enStoDays) = vStoDay
            .TextMatrix(nPtr - 1, enRemark) = txtRemark
            
            .TextMatrix(nPtr, enCounter) = "**"
            .Row = nPtr
        End With
        Call lzAddToTotal
        Call lzClearSto
    Else
        MsgBox "No additional charges computed...", vbInformation
        txtSTOCCRNo.SetFocus
    End If
End Sub

Private Function lzStoOversize() As Currency
Dim pLength, pWidth, pHeight As Single
    
    pLength = CSng("0" & txtStoOvzLen)
    pWidth = CSng("0" & txtStoOvzWid)
    pHeight = CSng("0" & txtStoOvzHgt)
    If txtStoUOM = "C" Then
       pLength = pLength / 2.54
       pWidth = pWidth / 2.54
       pHeight = pHeight / 2.54
    End If
    vRevTon = pLength * pWidth * pHeight / 1728 / 40
    Select Case txtStoContSz
        Case "20"
            If vRevTon >= cRTon20 Then vRevTon = vRevTon - cRTon20
        Case "40"
            If vRevTon >= cRTon40 Then vRevTon = vRevTon - cRTon40
        Case "45"
            If vRevTon >= cRTon45 Then vRevTon = vRevTon - cRTon45
        Case Else
            vRevTon = 0
    End Select
    vRevTon = Round(vRevTon, 2)
    If vRevTon > 0 Then
        lblStoRevTon = Format(vRevTon, "###,##0.00")
    Else
        lblStoRevTon = ""
    End If
    lzStoOversize = vRevTon * vRevTonRateSto * (vCYMStoDay + vStoDay)
    
End Function

Private Function lzStoExtDateValid() As Boolean
    lzStoExtDateValid = False
    If Not IsDate(txtStoExtDate) Then
        MsgBox "Invalid date.  Please correct..."
    Else
        If (DateDiff("d", CDate(lblStoValidUntil), CDate(txtStoExtDate)) < 1) Or _
           (DateDiff("d", gzGetSysDate(), CDate(txtStoExtDate)) < 0) Then
            MsgBox "Extension date cannot be less than date today.  Please correct..."
        Else
            lzStoExtDateValid = True
        End If
    End If
End Function

Private Sub txtStoUOM_Validate(Cancel As Boolean)
    Cancel = (txtStoUOM <> "I") And (txtStoUOM <> "C")
End Sub

Private Sub lzComputeMsc()
    ' compute
    nAmount = CCur("0" & lblMscAmount)
    nVATAmount = IIf(bVAT, nAmount * 0.1, 0)
    nWTaxAmount = IIf(bWTax, nAmount * 0.02, 0)
    nTotalAmount = nAmount + nVATAmount - nWTaxAmount
        
    If nTotalAmount > 0 Then
        ' new ccr
        If bNewCCR Or (nCCRCounter > 7) Then
            bNewCCR = False
            nCCRCounter = 1
        Else
            nCCRCounter = nCCRCounter + 1
        End If
        
        With grdCCRTran
            nPtr = .Rows
            .AddItem nPtr
            .TextMatrix(nPtr - 1, enCounter) = nPtr - 1
            .TextMatrix(nPtr - 1, enRateCode) = txtMscRateCode
            .TextMatrix(nPtr - 1, enAmount) = Format(nAmount, "###,##0.00")
            If nVATAmount > 0 Then .TextMatrix(nPtr - 1, enVATAmt) = Format(nVATAmount, "##,##0.00")
            If nWTaxAmount > 0 Then .TextMatrix(nPtr - 1, enWTaxAmt) = Format(nWTaxAmount, "##,##0.00")
            .TextMatrix(nPtr - 1, enTotalAmt) = Format(nTotalAmount, "##,##,##0.00")
            .TextMatrix(nPtr - 1, enCCRTag) = IIf(nCCRCounter = 1, "*", " ")
            
            .TextMatrix(nPtr - 1, enCCRNo) = txtMscCCRNo
            .TextMatrix(nPtr - 1, enContNo) = txtMscContNo
            .TextMatrix(nPtr - 1, enContSz) = txtMscContSz
            .TextMatrix(nPtr - 1, enQuantity) = txtMscQty
            .TextMatrix(nPtr - 1, enRemark) = txtRemark
            
            .TextMatrix(nPtr, enCounter) = "**"
            .Row = nPtr
        End With
        Call lzAddToTotal
        Call lzClearMsc
    Else
        MsgBox "No additional charges computed...", vbInformation
        txtMscCCRNo.SetFocus
    End If
End Sub

Private Sub lzComputeOth()
    ' compute
    On Error GoTo err_Amt
    nAmount = CCur("0" & txtOthAmount)
    If bOTHOversize Then
        Call lzOthOversize
    End If
    nVATAmount = IIf(bVAT, nAmount * 0.1, 0)
    nWTaxAmount = IIf(bWTax, nAmount * 0.02, 0)
    nTotalAmount = nAmount + nVATAmount - nWTaxAmount
        
    If nTotalAmount > 0 Then
        ' new ccr
        If bNewCCR Or (nCCRCounter > 7) Then
            bNewCCR = False
            nCCRCounter = 1
        Else
            nCCRCounter = nCCRCounter + 1
        End If
        
        With grdCCRTran
            nPtr = .Rows
            .AddItem nPtr
            .TextMatrix(nPtr - 1, enCounter) = nPtr - 1
            .TextMatrix(nPtr - 1, enRateCode) = "OTHERS"
            .TextMatrix(nPtr - 1, enAmount) = Format(nAmount, "###,##0.00")
            If nVATAmount > 0 Then .TextMatrix(nPtr - 1, enVATAmt) = Format(nVATAmount, "##,##0.00")
            If nWTaxAmount > 0 Then .TextMatrix(nPtr - 1, enWTaxAmt) = Format(nWTaxAmount, "##,##0.00")
            .TextMatrix(nPtr - 1, enTotalAmt) = Format(nTotalAmount, "##,##,##0.00")
            .TextMatrix(nPtr - 1, enCCRTag) = IIf(nCCRCounter = 1, "*", " ")
            
            .TextMatrix(nPtr - 1, enCCRNo) = txtOTHCCRNo
            .TextMatrix(nPtr - 1, enContNo) = txtOthContNo
            .TextMatrix(nPtr - 1, enContSz) = txtOthContSz
            .TextMatrix(nPtr - 1, enFulEmp) = txtOthFulEmp
            .TextMatrix(nPtr - 1, enEntryNo) = txtOthEntryNo
            .TextMatrix(nPtr - 1, enRegNo) = txtOthRegNo
            If bOTHOversize Then
                .TextMatrix(nPtr - 1, enOvzLen) = txtOthOvzLen
                .TextMatrix(nPtr - 1, enOvzWid) = txtOthOvzWid
                .TextMatrix(nPtr - 1, enOvzHgt) = txtOthOvzHgt
                .TextMatrix(nPtr - 1, enOvzUom) = txtOthUOM
                .TextMatrix(nPtr - 1, enRevTon) = lblOthRevTon
                .TextMatrix(nPtr - 1, enOvzAmt) = 0
            End If
            .TextMatrix(nPtr - 1, enVessel) = txtOthVessel
            .TextMatrix(nPtr - 1, enRemark) = txtRemark
            
            .TextMatrix(nPtr, enCounter) = "**"
            .Row = nPtr
        End With
        Call lzAddToTotal
        Call lzClearOth
    Else
        MsgBox "No additional charges computed...", vbInformation
        txtOTHCCRNo.SetFocus
    End If
    Exit Sub

err_Amt:
        MsgBox "Invalid amount.  Please re-enter.", vbInformation
        On Error GoTo 0
        txtOthAmount.SetFocus
End Sub

Private Sub lzOthOversize()
Dim pLength, pWidth, pHeight As Single
    
    pLength = CSng(txtOthOvzLen)
    pWidth = CSng(txtOthOvzWid)
    pHeight = CSng(txtOthOvzHgt)
    If txtOthUOM = "C" Then
       pLength = pLength / 2.54
       pWidth = pWidth / 2.54
       pHeight = pHeight / 2.54
    End If
    vRevTon = pLength * pWidth * pHeight / 1728 / 40
    Select Case txtOthContSz
        Case "20"
            If vRevTon >= cRTon20 Then vRevTon = vRevTon - cRTon20
        Case "40"
            If vRevTon >= cRTon40 Then vRevTon = vRevTon - cRTon40
        Case "45"
            If vRevTon >= cRTon45 Then vRevTon = vRevTon - cRTon45
        Case Else
            vRevTon = 0
    End Select
    vRevTon = Round(vRevTon, 2)
    If vRevTon > 0 Then
        lblOthRevTon = Format(vRevTon, "###,##0.00")
    Else
        lblOthRevTon = ""
    End If

End Sub

Private Sub lzShowRate()
    Call lzGetRateInfo(txtMscRateCode, txtMscContSz)
    If vRateCode <> "" Then
        txtMscRateCode = vRateCode
        lblMScRateDesc = vRateDesc
        txtMscContSz = vRateSz
        lblMscRateAmt = Format(vRateAmount, "###,##0.00")
        lblMscRateUOM = vRateUOM
        lblMscAmount = Format(CCur("0" & lblMscRateAmt) * CCur("0" & txtMscQty), "###,##0.00")
        If (Trim(txtMscQty) = "") And (vRateAmount > 0) Then txtMscQty = 1
        txtMscQty.SetFocus
    End If
End Sub

Private Sub lzComputeRfr()
Dim h, m As Single
Dim bDateError As Boolean
    
    bDateError = (txtRfrPlugInDate = cEmptyRfrDate) Or _
                 (lblRfrValidUntil = cEmptyRfrDate) Or _
                 (txtRfrExtDate = cEmptyRfrDate)
    If bDateError Then
       MsgBox "One or more date is invalid. Please correct.", vbExclamation
       Exit Sub
    End If
    
    ' compute
    cRateCode = "IMRF"
    nAmount = lzGetRateInfo(cRateCode, txtRfrContSz)
    h = DateDiff("n", CDate(lblRfrValidUntil), CDate(txtRfrExtDate))
    h = h / 60
    m = h - Fix(h)
    vRfrHours = Fix(h / 6) * 6
    m = m + ((h / 6) - Fix(h / 6))
    If m > 0 Then vRfrHours = vRfrHours + 6
    If vRfrHours < 6 Then vRfrHours = 6
    txtRfrExtDate = Format(DateAdd("h", vRfrHours, CDate(lblRfrValidUntil)), "YYYY-MM-DD hh:mm")
    nAmount = nAmount * (vRfrHours / 6)
    nVATAmount = IIf(bVAT, nAmount * 0.1, 0)
    nWTaxAmount = IIf(bWTax, nAmount * 0.02, 0)
    nTotalAmount = nAmount + nVATAmount - nWTaxAmount
        
    If nTotalAmount > 0 Then
        ' new ccr
        If bNewCCR Or (nCCRCounter > 7) Then
            bNewCCR = False
            nCCRCounter = 1
        Else
            nCCRCounter = nCCRCounter + 1
        End If
        
        With grdCCRTran
            nPtr = .Rows
            .AddItem nPtr
            .TextMatrix(nPtr - 1, enCounter) = nPtr - 1
            .TextMatrix(nPtr - 1, enRateCode) = cRateCode
            .TextMatrix(nPtr - 1, enAmount) = Format(nAmount, "###,##0.00")
            If nVATAmount > 0 Then .TextMatrix(nPtr - 1, enVATAmt) = Format(nVATAmount, "##,##0.00")
            If nWTaxAmount > 0 Then .TextMatrix(nPtr - 1, enWTaxAmt) = Format(nWTaxAmount, "##,##0.00")
            .TextMatrix(nPtr - 1, enTotalAmt) = Format(nTotalAmount, "##,##,##0.00")
            .TextMatrix(nPtr - 1, enCCRTag) = IIf(nCCRCounter = 1, "*", " ")
            
            .TextMatrix(nPtr - 1, enCCRNo) = txtRFRCCRNo
            .TextMatrix(nPtr - 1, enContNo) = txtRfrContNo
            .TextMatrix(nPtr - 1, enContSz) = txtRfrContSz
            .TextMatrix(nPtr - 1, enEntryNo) = txtRfrEntryNo
            .TextMatrix(nPtr - 1, enRegNo) = txtRfrRegNo
            .TextMatrix(nPtr - 1, enRfrHours) = vRfrHours
            .TextMatrix(nPtr - 1, enRfrValidUntil) = txtRfrExtDate
            
            txtRemark = "REEFER UP TO " & Format(txtRfrExtDate, "YYYY/MM/DD hh:mm")
            .TextMatrix(nPtr - 1, enRemark) = txtRemark
            
            .TextMatrix(nPtr, enCounter) = "**"
            .Row = nPtr
        End With
        Call lzAddToTotal
        'MsgBox "Reefer extended by " & vRfrHours & " hours until " & txtRfrExtDate, vbInformation
        Call lzClearRfr
    Else
        MsgBox "No additional charges computed...", vbInformation
        txtRFRCCRNo.SetFocus
    End If
End Sub

Private Sub lzInitializePay()
Dim n As Integer
       
    If Not fraPayment.Enabled Then fraPayment.Enabled = True
    txtCshAmt = lblAmtDue
    For n = 0 To 4
        txtChkAmt(n) = ""
        txtChkNo(n) = Space(txtChkNo(n).MaxLength)
        txtChkBank(n) = Space(txtChkBank(n).MaxLength)
    Next n
    lblChkTot = ""
    lblChange = ""
    txtCshAmt.SetFocus
    
End Sub

Private Sub lzComputePay()
Dim n As Integer
Dim curChkTot, curChange As Currency
    
    curChkTot = 0
    For n = 0 To 4
        curChkTot = curChkTot + CCur("0" & txtChkAmt(n))
    Next
    lblChkTot = Format(curChkTot, "##,###,##0.00")
    curChange = CCur("0" & txtCshAmt) + CCur("0" & lblChkTot) - CCur("0" & lblAmtDue)
    lblChange = Format(curChange, "##,###,##0.00")
    lblChange.ForeColor = IIf(curChange < 0, vbRed, vbBlue)
    mnuMenuSave.Enabled = (curChange >= 0)
    cmdSave.Enabled = (curChange >= 0)
    
End Sub

Private Sub lzDeleteItem()
Dim bNewCCR As Boolean
Dim n As Integer
    With grdCCRTran
        If (.Row < .Rows) And (.TextMatrix(.Row, enRateCode) <> cVoid) Then
            bNewCCR = (.TextMatrix(.Row, enCCRTag) = "*")
            Call lzLessFromTotal
            .RemoveItem .Row
            .AddItem "", .Row
            .TextMatrix(.Row, enRateCode) = cVoid
            If bNewCCR Then
                n = .Row + 1
                While n < (.Rows - 1)
                    If (.TextMatrix(n, enRateCode) <> cVoid) Then
                        .TextMatrix(n, enCCRTag) = "*"
                        n = .Rows
                    Else
                        n = n + 1
                    End If
                Wend
            End If
        End If
        .SetFocus
    End With
End Sub

Private Sub lzLessFromTotal()
Dim curTotal As Currency
    curTotal = CCur("0" & lblAmtDue)
    nTotalAmount = CCur("0" & grdCCRTran.TextMatrix(grdCCRTran.Row, enTotalAmt))
    lblAmtDue = Format(curTotal - nTotalAmount, "#,###,##0.00")
    lblAmtDue.Refresh
End Sub

Private Sub lzAddTran()
    If Not grdCCRTran.Enabled Then grdCCRTran.Enabled = True
    If Not mnuMenuPayment.Enabled Then mnuMenuPayment.Enabled = True
    
    Select Case tabTran.Tab
        Case 0
            Call lzComputeArr
        Case 1
            Call lzComputeSto
        Case 2
            Call lzComputeRfr
        Case 3
            Call lzComputeSO
        Case 4
            Call lzComputeMsc
        Case 5
            Call lzComputeOth
        Case Else
    End Select
    
    chkNewCCR.Value = IIf(bNewCCR, 1, 0)
End Sub

Private Sub lzEnableTab()
Dim n As Integer
    With tabTran
        For n = 0 To .Tabs - 1
            .TabEnabled(n) = IIf(n = vTabOn, True, False)
        Next n
        .Tab = vTabOn
    End With
End Sub

Private Function lzGetNextCCR(ByVal pUserID As String) As Long
Dim cmd As ADODB.Command
    ' create command
    Set cmd = New ADODB.Command
    With cmd
        Set .ActiveConnection = gcnnBilling
        .CommandText = "up_getnextspl"
        .CommandType = adCmdStoredProc
        
        .Parameters(1).Type = adChar
        .Parameters(1).Value = pUserID
        .Parameters(1).Direction = adParamInput             ' user id
        .Parameters(2).Type = adInteger
        .Parameters(2).Direction = adParamOutput            ' next ccr
    
        .Execute
        
        lzGetNextCCR = .Parameters(2)
    
    End With
    Set cmd = Nothing
End Function

Private Function lzCCRValid(ByVal pUserID As String, ByVal pCCRNo As Long) As Boolean
Dim cmd As ADODB.Command
    
    ' create command
    Set cmd = New ADODB.Command
    With cmd
        Set .ActiveConnection = gcnnBilling
        .CommandText = "up_chkvalidspl"
        .CommandType = adCmdStoredProc
    
        ' set parameters then execute
        .Parameters(0).Direction = adParamReturnValue
        .Parameters(1).Type = adChar
        .Parameters(1).Value = pUserID
        .Parameters(1).Direction = adParamInput
        .Parameters(2).Type = adInteger
        .Parameters(2).Value = pCCRNo
        .Parameters(2).Direction = adParamInput
       
        .Execute
        
        lzCCRValid = (.Parameters(0) > 0)
    
    End With
    
End Function

Private Function lzGetControlNo() As Long
Dim cmd As ADODB.Command
    ' create command
    Set cmd = New ADODB.Command
    With cmd
        Set .ActiveConnection = gcnnBilling
        .CommandText = "up_getcontrolno"
        .CommandType = adCmdStoredProc
        
        .Parameters(1).Type = adChar
        .Parameters(1).Value = "CCR"
        .Parameters(1).Direction = adParamInput             '
        .Parameters(2).Type = adInteger
        .Parameters(2).Direction = adParamOutput            ' control number
    
        .Execute
        
        lzGetControlNo = .Parameters(2)
    
    End With
    Set cmd = Nothing
End Function
Function writeTextLab()
    
End Function
Private Sub lzSavePrint()
Dim cmd As ADODB.Command
Dim vRef, vSeq, vItem, vCCR, n As Long
Dim bValidCCR As Boolean
Dim v As Long
Dim c As New CWaitCursor
    
    ' validate required info
    If Len(Trim(txtCusName)) = 0 Then
        MsgBox "Customer name required...", vbInformation
        txtCusName.SetFocus
        Exit Sub
    End If
   
    bValidCCR = False
    vCCR = lzGetNextCCR(gUserID)
    While Not bValidCCR
        'v = CLng("0" & Trim(InputBox("Enter CCR number: ", , Str(vCCR))))
        vNextCCR = vCCR
        frmNextCCR.Show 1
        v = vNextCCR
        
        If v > 0 Then
            bValidCCR = lzCCRValid(gUserID, v)
            If bValidCCR Then vCCR = v
        Else
            bValidCCR = True
            txtCshAmt.SetFocus
            Exit Sub
        End If
    Wend
    
    vRef = lzGetControlNo
    If vRef <= 0 Then
        MsgBox "CCR control number request error.  Please try again later.", vbInformation
        Exit Sub
    End If
    
    ' create command
    Set cmd = Nothing
    Set cmd = New ADODB.Command
    With cmd
        Set .ActiveConnection = gcnnBilling
        
        ' write payment first
        .CommandText = "up_wrtccrpay"
        .CommandType = adCmdStoredProc
        
        .Parameters(1).Type = adInteger
        .Parameters(1).Value = vRef
        .Parameters(1).Direction = adParamInput             ' control number
        .Parameters(2).Type = adChar
        If bUnderG Then
            .Parameters(2).Value = vCusCodeUnderG
        Else
            .Parameters(2).Value = ""
        End If
        .Parameters(2).Direction = adParamInput             '
        .Parameters(3).Type = adChar
        .Parameters(3).Value = txtCusName
        .Parameters(3).Direction = adParamInput             '
        .Parameters(4).Type = adCurrency
        .Parameters(4).Value = CCur("0" & txtCshAmt) + CCur("0" & txtChkAmt(0)) + CCur("0" & txtChkAmt(1)) + _
                               CCur("0" & txtChkAmt(2)) + CCur("0" & txtChkAmt(3)) + CCur("0" & txtChkAmt(4))
        .Parameters(4).Direction = adParamInput             '
        .Parameters(5).Type = adCurrency
        .Parameters(5).Value = 0
        .Parameters(5).Direction = adParamInput             '
        .Parameters(6).Type = adInteger
        .Parameters(6).Value = 0
        .Parameters(6).Direction = adParamInput             '
        .Parameters(7).Type = adCurrency
        .Parameters(7).Value = CCur("0" & lblChange)
        .Parameters(7).Direction = adParamInput             '
        .Parameters(8).Type = adChar
        .Parameters(8).Value = txtChkNo(0)
        .Parameters(8).Direction = adParamInput             '
        .Parameters(9).Type = adChar
        .Parameters(9).Value = txtChkNo(1)
        .Parameters(9).Direction = adParamInput             '
        .Parameters(10).Type = adChar
        .Parameters(10).Value = txtChkNo(2)
        .Parameters(10).Direction = adParamInput             '
        .Parameters(11).Type = adChar
        .Parameters(11).Value = txtChkNo(3)
        .Parameters(11).Direction = adParamInput             '
        .Parameters(12).Type = adChar
        .Parameters(12).Value = txtChkNo(4)
        .Parameters(12).Direction = adParamInput             '
        .Parameters(13).Type = adNumeric
        .Parameters(13).Value = CCur("0" & txtChkAmt(0))
        .Parameters(13).Direction = adParamInput             '
        .Parameters(14).Type = adNumeric
        .Parameters(14).Value = CCur("0" & txtChkAmt(1))
        .Parameters(14).Direction = adParamInput             '
        .Parameters(15).Type = adNumeric
        .Parameters(15).Value = CCur("0" & txtChkAmt(2))
        .Parameters(15).Direction = adParamInput             '
        .Parameters(16).Type = adNumeric
        .Parameters(16).Value = CCur("0" & txtChkAmt(3))
        .Parameters(16).Direction = adParamInput             '
        .Parameters(17).Type = adNumeric
        .Parameters(17).Value = CCur("0" & txtChkAmt(4))
        .Parameters(17).Direction = adParamInput             '
        .Parameters(18).Type = adChar
        .Parameters(18).Value = txtChkBank(0)
        .Parameters(18).Direction = adParamInput             '
        .Parameters(19).Type = adChar
        .Parameters(19).Value = txtChkBank(1)
        .Parameters(19).Direction = adParamInput             '
        .Parameters(20).Type = adChar
        .Parameters(20).Value = txtChkBank(2)
        .Parameters(20).Direction = adParamInput             '
        .Parameters(21).Type = adChar
        .Parameters(21).Value = txtChkBank(3)
        .Parameters(21).Direction = adParamInput             '
        .Parameters(22).Type = adChar
        .Parameters(22).Value = txtChkBank(4)
        .Parameters(22).Direction = adParamInput             '
        .Parameters(23).Type = adChar
        .Parameters(23).Value = gUserID
        .Parameters(23).Direction = adParamInput             '

        .Execute
    
        ' write details next
        .CommandText = "up_wrtccrdtl"
        .CommandType = adCmdStoredProc
    
        vSeq = 0: vItem = 0: vCCR = vCCR - 1
        For n = 1 To (grdCCRTran.Rows - 2)
            If (grdCCRTran.TextMatrix(n, enRateCode) <> cVoid) Then
                If (n = 1) Or (grdCCRTran.TextMatrix(n, enCCRTag) = "*") Then
                    vSeq = vSeq + 1
                    vCCR = vCCR + 1
                    vItem = 0
                    
                    ' re check allocation
                    bValidCCR = lzCCRValid(gUserID, vCCR)
                    If Not bValidCCR Then
                        While Not bValidCCR
                            vNextCCR = vCCR
                            frmNextCCR.Show 1
                            v = vNextCCR
                            If v > 0 Then
                                bValidCCR = lzCCRValid(gUserID, v)
                                If bValidCCR Then vCCR = v
                            Else
                                bValidCCR = True
                                MsgBox "Please void this transaction!!" & vbCrLf _
                                        & "Reference No. " & Val(vRef), vbCritical
                                
                                Call lzApplyCCR(gUserID, vCCR)
                                Call lzInitialize
                                Call lzGetUserInfo
                                
                                GoTo tag_ReInit
                            End If
                        Wend
                    End If
                    
                End If
                vItem = vItem + 1
                If vSeq = 0 Then vSeq = 1
                ' set parameters then execute
                .Parameters(1).Type = adInteger
                .Parameters(1).Value = vRef
                .Parameters(1).Direction = adParamInput             '
                .Parameters(2).Type = adInteger
                .Parameters(2).Value = vSeq
                .Parameters(2).Direction = adParamInput             '
                .Parameters(3).Type = adInteger
                .Parameters(3).Value = vItem
                .Parameters(3).Direction = adParamInput             '
                .Parameters(4).Type = adInteger
                .Parameters(4).Value = vCCR
                .Parameters(4).Direction = adParamInput             '
                .Parameters(5).Type = adChar
                .Parameters(5).Value = Trim(grdCCRTran.TextMatrix(n, enRateCode)) ' Error part
                .Parameters(5).Direction = adParamInput
                .Parameters(6).Type = adNumeric
                .Parameters(6).Value = CCur("0" & grdCCRTran.TextMatrix(n, enAmount))
                .Parameters(6).Direction = adParamInput             '
                .Parameters(7).Type = adNumeric
                .Parameters(7).Value = CCur("0" & grdCCRTran.TextMatrix(n, enVATAmt))
                .Parameters(7).Direction = adParamInput             '
                .Parameters(8).Type = adNumeric
                .Parameters(8).Value = CCur("0" & grdCCRTran.TextMatrix(n, enWTaxAmt))
                .Parameters(8).Direction = adParamInput             '
                .Parameters(9).Type = adInteger
                .Parameters(9).Value = CLng("0" & grdCCRTran.TextMatrix(n, enCCRNo))
                .Parameters(9).Direction = adParamInput             '
                .Parameters(10).Type = adChar
                .Parameters(10).Value = grdCCRTran.TextMatrix(n, enContNo)    'B
                .Parameters(10).Direction = adParamInput             '
                .Parameters(11).Type = adNumeric
                .Parameters(11).Value = CLng("0" & grdCCRTran.TextMatrix(n, enContSz))
                .Parameters(11).Direction = adParamInput             '
                .Parameters(12).Type = adChar
                .Parameters(12).Value = grdCCRTran.TextMatrix(n, enFulEmp)  'C
                .Parameters(12).Direction = adParamInput             '
                .Parameters(13).Type = adInteger
                .Parameters(13).Value = CLng("0" & grdCCRTran.TextMatrix(n, enEntryNo))
                .Parameters(13).Direction = adParamInput             '
                .Parameters(14).Type = adChar
                .Parameters(14).Value = grdCCRTran.TextMatrix(n, enRegNo)  'D
                .Parameters(14).Direction = adParamInput
                .Parameters(15).Type = adNumeric
                .Parameters(15).Value = CCur("0" & grdCCRTran.TextMatrix(n, enOvzLen))
                .Parameters(15).Direction = adParamInput             '
                .Parameters(16).Type = adNumeric
                .Parameters(16).Value = CCur("0" & grdCCRTran.TextMatrix(n, enOvzWid))
                .Parameters(16).Direction = adParamInput             '
                .Parameters(17).Type = adNumeric
                .Parameters(17).Value = CCur("0" & grdCCRTran.TextMatrix(n, enOvzHgt))
                .Parameters(17).Direction = adParamInput             '
                .Parameters(18).Type = adChar
                .Parameters(18).Value = grdCCRTran.TextMatrix(n, enOvzUom) 'F
                .Parameters(18).Direction = adParamInput             '
                .Parameters(19).Type = adNumeric
                .Parameters(19).Value = CCur("0" & grdCCRTran.TextMatrix(n, enRevTon))
                .Parameters(19).Direction = adParamInput             '
                .Parameters(20).Type = adChar
                .Parameters(20).Value = Left(grdCCRTran.TextMatrix(n, enDangerCode), 1)  'G
                .Parameters(20).Direction = adParamInput             '
                .Parameters(21).Type = adDate
                .Parameters(21).Value = CDate(IIf(IsDate(grdCCRTran.TextMatrix(n, enStoValidUntil)), grdCCRTran.TextMatrix(n, enStoValidUntil), cNullDate))
                .Parameters(21).Direction = adParamInput             '
                .Parameters(22).Type = adDate
                .Parameters(22).Value = CDate(IIf(IsDate(grdCCRTran.TextMatrix(n, enRfrValidUntil)), grdCCRTran.TextMatrix(n, enRfrValidUntil), cNullDate))
                .Parameters(22).Direction = adParamInput             '
                .Parameters(23).Type = adNumeric
                .Parameters(23).Value = CLng("0" & grdCCRTran.TextMatrix(n, enStoDays))
                .Parameters(23).Direction = adParamInput             '
                .Parameters(24).Type = adNumeric
                .Parameters(24).Value = CCur("0" & grdCCRTran.TextMatrix(n, enQuantity))
                .Parameters(24).Direction = adParamInput             '
                .Parameters(25).Type = adChar
                .Parameters(25).Value = grdCCRTran.TextMatrix(n, enVessel)  'H
                .Parameters(25).Direction = adParamInput             '
                .Parameters(26).Type = adNumeric
                If CCur("0" & Trim(grdCCRTran.TextMatrix(n, enDangerAmt))) > 0 Then
                    .Parameters(26).Value = CCur("0" & grdCCRTran.TextMatrix(n, enAmount))
                    .Parameters(6).Value = 0
                Else
                    .Parameters(26).Value = 0
                End If
                .Parameters(26).Direction = adParamInput             '
                .Parameters(27).Type = adNumeric
                If Trim(grdCCRTran.TextMatrix(n, enOvzAmt)) <> "" Then
                    .Parameters(27).Value = CCur("0" & grdCCRTran.TextMatrix(n, enAmount))
                    .Parameters(6).Value = 0
                Else
                    .Parameters(27).Value = 0
                End If
                .Parameters(27).Direction = adParamInput             '
                .Parameters(28).Type = adVarChar
                .Parameters(28).Size = 30
                .Parameters(28).Value = Left(grdCCRTran.TextMatrix(n, enRemark), 30) 'I
                .Parameters(28).Direction = adParamInput             '
                .Parameters(29).Type = adChar
                .Parameters(29).Value = grdCCRTran.TextMatrix(n, enShipLine)   'J
                .Parameters(29).Direction = adParamInput             '
                .Parameters(30).Type = adChar
                .Parameters(30).Value = IIf(bUnderG, "Y", " ")
                .Parameters(30).Direction = adParamInput             '
                .Parameters(31).Type = adNumeric
                .Parameters(31).Value = CLng("0" & grdCCRTran.TextMatrix(n, enRfrHours))
                .Parameters(31).Direction = adParamInput             '
                .Parameters(32).Type = adChar
                .Parameters(32).Value = Left(Trim(gUserID), 10)
                .Parameters(32).Direction = adParamInput             '
   
                .Execute
        
            End If
        Next n
        
        Call lzApplyCCR(gUserID, vCCR)
        Call lzInitialize
        Call lzGetUserInfo
        
        ' call printing module
        Call lzPrintCCR(vRef, vCCR)
       
tag_ReInit:
    
    End With
    Set cmd = Nothing
    txtCshAmt.ForeColor = vbWindowBackground
    txtCusName.SetFocus
    
    Exit Sub

err_Save:
    MsgBox "Error in saving this transaction...", vbExclamation
    txtCshAmt.SetFocus
End Sub

Private Sub lzApplyCCR(ByVal pUserID As String, ByVal pCCRNo As Long)
Dim cmd As ADODB.Command
    
    ' create command
    Set cmd = New ADODB.Command
    With cmd
        Set .ActiveConnection = gcnnBilling
        .CommandText = "up_applyccrspl"
        .CommandType = adCmdStoredProc
    
        ' set parameters then execute
        .Parameters(1).Type = adChar
        .Parameters(1).Value = pUserID
        .Parameters(1).Direction = adParamInput
        .Parameters(2).Type = adInteger
        .Parameters(2).Value = pCCRNo
        .Parameters(2).Direction = adParamInput

        .Execute

    End With
    
End Sub

Private Sub lzPrintCCR(ByVal pREFNUM As Long, ByVal pCCR As Long)
    With clsCCRPrinter
        .CCRSupervisor vSupervisor
        .CCRNumber = pCCR
        Call .PrintCCR(pREFNUM)
        'Call .PreviewCCR(pREFNUM)
    End With
End Sub

Private Function lzGetCustomerName(ByVal pCode As String) As String
Dim cmdGetCustomer As ADODB.Command
Dim prmGetCustomer As ADODB.Parameter
    
    ' create command
    Set cmdGetCustomer = New ADODB.Command
    With cmdGetCustomer
        Set .ActiveConnection = gcnnBilling
        .CommandText = "up_getcustomerinfo"
        .CommandType = adCmdStoredProc
    
        ' set parameters then execute
        .Parameters(0).Direction = adParamReturnValue
        .Parameters(1).Type = adChar
        .Parameters(1).Value = pCode
        .Parameters(1).Direction = adParamInput
        .Parameters(2).Type = adChar
        .Parameters(2).Direction = adParamOutput
        .Parameters(3).Type = adChar
        .Parameters(3).Direction = adParamOutput
        .Parameters(4).Type = adChar
        .Parameters(4).Direction = adParamOutput
        .Parameters(5).Type = adChar
        .Parameters(5).Direction = adParamOutput
        .Parameters(6).Type = adChar
        .Parameters(6).Direction = adParamOutput
       
        .Execute
        
        lzGetCustomerName = Trim("" & .Parameters(3))
     End With
    
End Function
