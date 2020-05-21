VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form 密码修改 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "密码修改"
   ClientHeight    =   4320
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5595
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   12
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "密码修改.frx":0000
   ScaleHeight     =   4320
   ScaleWidth      =   5595
   StartUpPosition =   3  '窗口缺省
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   1680
      Top             =   0
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label Label2 
      Caption         =   "密码"
      Height          =   495
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   1335
   End
   Begin MSForms.CommandButton CommandButton1 
      Height          =   495
      Left            =   2880
      TabIndex        =   8
      Top             =   3600
      Width           =   1455
      VariousPropertyBits=   19
      Caption         =   "确  定"
      Size            =   "2566;873"
      FontName        =   "宋体"
      FontEffects     =   1073741825
      FontHeight      =   240
      FontCharSet     =   134
      FontPitchAndFamily=   34
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.TextBox TextBox4 
      Height          =   375
      Left            =   2760
      TabIndex        =   7
      Top             =   3080
      Width           =   2295
      VariousPropertyBits=   746604563
      Size            =   "4048;661"
      FontName        =   "宋体"
      FontHeight      =   240
      FontCharSet     =   134
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox TextBox3 
      Height          =   375
      Left            =   2760
      TabIndex        =   6
      Top             =   2480
      Width           =   2295
      VariousPropertyBits=   746604563
      Size            =   "4048;661"
      FontName        =   "宋体"
      FontHeight      =   240
      FontCharSet     =   134
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox TextBox2 
      Height          =   375
      Left            =   2760
      TabIndex        =   5
      Top             =   1900
      Width           =   2295
      VariousPropertyBits=   746604563
      Size            =   "4048;661"
      FontName        =   "宋体"
      FontHeight      =   240
      FontCharSet     =   134
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox TextBox1 
      Height          =   375
      Left            =   2760
      TabIndex        =   4
      Top             =   1320
      Width           =   2295
      VariousPropertyBits=   746604563
      Size            =   "4048;661"
      FontName        =   "宋体"
      FontHeight      =   240
      FontCharSet     =   134
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label1 
      Height          =   330
      Index           =   3
      Left            =   1920
      TabIndex        =   3
      Top             =   3120
      Width           =   975
      VariousPropertyBits=   8388627
      Caption         =   "新密码："
      Size            =   "1720;582"
      FontName        =   "宋体"
      FontHeight      =   240
      FontCharSet     =   134
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label1 
      Height          =   330
      Index           =   2
      Left            =   1920
      TabIndex        =   2
      Top             =   2520
      Width           =   975
      VariousPropertyBits=   8388627
      Caption         =   "新密码："
      Size            =   "1720;582"
      FontName        =   "宋体"
      FontHeight      =   240
      FontCharSet     =   134
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label1 
      Height          =   330
      Index           =   1
      Left            =   1920
      TabIndex        =   1
      Top             =   1960
      Width           =   975
      VariousPropertyBits=   8388627
      Caption         =   "原密码："
      Size            =   "1720;582"
      FontName        =   "宋体"
      FontHeight      =   240
      FontCharSet     =   134
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label1 
      Height          =   330
      Index           =   0
      Left            =   1920
      TabIndex        =   0
      Top             =   1350
      Width           =   975
      VariousPropertyBits=   8388627
      Caption         =   "用户名："
      Size            =   "1720;582"
      FontName        =   "宋体"
      FontHeight      =   240
      FontCharSet     =   134
      FontPitchAndFamily=   34
   End
End
Attribute VB_Name = "密码修改"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
