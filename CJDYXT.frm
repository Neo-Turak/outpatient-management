VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form 新增档案 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "新增档案"
   ClientHeight    =   9750
   ClientLeft      =   210
   ClientTop       =   210
   ClientWidth     =   9150
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   14.25
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "CJDYXT.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "CJDYXT.frx":058A
   ScaleHeight     =   9750
   ScaleWidth      =   9150
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   6600
      Top             =   9480
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
      Connect         =   "Provider=SQLOLEDB.1;Password=sa;Persist Security Info=True;User ID=sa;Initial Catalog=ghgl;Data Source=NURA\SQLEXPRESS"
      OLEDBString     =   "Provider=SQLOLEDB.1;Password=sa;Persist Security Info=True;User ID=sa;Initial Catalog=ghgl;Data Source=NURA\SQLEXPRESS"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "患者总表"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataGridLib.DataGrid DataGrid2 
      Bindings        =   "CJDYXT.frx":2405CC
      Height          =   2175
      Left            =   120
      TabIndex        =   19
      Top             =   7320
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   3836
      _Version        =   393216
      AllowUpdate     =   -1  'True
      AllowArrows     =   -1  'True
      BackColor       =   32896
      ColumnHeaders   =   -1  'True
      ForeColor       =   -2147483641
      HeadLines       =   1
      RowHeight       =   18
      WrapCellPointer =   -1  'True
      FormatLocked    =   -1  'True
      AllowAddNew     =   -1  'True
      AllowDelete     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   15
      BeginProperty Column00 
         DataField       =   "患者编号"
         Caption         =   "患者编号"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "合作医疗号"
         Caption         =   "合作医疗号"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "患者姓名"
         Caption         =   "患者姓名"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "身份证号"
         Caption         =   "身份证号"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "性别"
         Caption         =   "性别"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "出生日期"
         Caption         =   "出生日期"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "年龄"
         Caption         =   "年龄"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "婚姻状况"
         Caption         =   "婚姻状况"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column08 
         DataField       =   "民族"
         Caption         =   "民族"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column09 
         DataField       =   "联系电话"
         Caption         =   "联系电话"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column10 
         DataField       =   "家庭住址"
         Caption         =   "家庭住址"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column11 
         DataField       =   "有无过敏史"
         Caption         =   "有无过敏史"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column12 
         DataField       =   "结算方式"
         Caption         =   "结算方式"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column13 
         DataField       =   "建档日期"
         Caption         =   "建档日期"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column14 
         DataField       =   "备注"
         Caption         =   "备注"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   1319.811
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1425.26
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1305.071
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1184.882
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   705.26
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1124.787
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   689.953
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1305.071
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   689.953
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   1154.835
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   1124.787
         EndProperty
         BeginProperty Column11 
            ColumnWidth     =   1214.929
         EndProperty
         BeginProperty Column12 
            ColumnWidth     =   1154.835
         EndProperty
         BeginProperty Column13 
            ColumnWidth     =   1170.142
         EndProperty
         BeginProperty Column14 
            ColumnWidth     =   824.882
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "档案信息"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9375
      Left            =   120
      TabIndex        =   14
      Top             =   120
      Width           =   8895
      Begin VB.CommandButton Command3 
         Caption         =   "删 除"
         Height          =   615
         Left            =   4320
         TabIndex        =   13
         Top             =   6480
         Width           =   1695
      End
      Begin VB.TextBox Text7 
         Height          =   405
         Left            =   1680
         MaxLength       =   11
         TabIndex        =   8
         Text            =   "Text7"
         Top             =   4680
         Width           =   3015
      End
      Begin VB.TextBox Text4 
         Height          =   405
         Left            =   1680
         MaxLength       =   3
         TabIndex        =   5
         Text            =   "Text4"
         Top             =   3240
         Width           =   3015
      End
      Begin MSComctlLib.Toolbar Toolbar2 
         Height          =   330
         Left            =   7560
         TabIndex        =   32
         Top             =   360
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   582
         ButtonWidth     =   1588
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Appearance      =   1
         Style           =   1
         TextAlignment   =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "相   片"
               Object.ToolTipText     =   "插入照片"
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   3
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "从本地文件导入"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "从摄像头抓图"
                  EndProperty
                  BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "清楚相片"
                  EndProperty
               EndProperty
            EndProperty
         EndProperty
         OLEDropMode     =   1
      End
      Begin VB.PictureBox Picture1 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2295
         Left            =   4800
         ScaleHeight     =   2235
         ScaleWidth      =   2475
         TabIndex        =   31
         Top             =   360
         Width           =   2535
      End
      Begin VB.ListBox List2 
         Height          =   3195
         Left            =   5400
         TabIndex        =   30
         Top             =   2880
         Width           =   3375
      End
      Begin VB.ListBox List1 
         BackColor       =   &H80000004&
         Height          =   3195
         ItemData        =   "CJDYXT.frx":2405E1
         Left            =   4800
         List            =   "CJDYXT.frx":2405E3
         TabIndex        =   29
         Top             =   2880
         Width           =   495
      End
      Begin VB.ComboBox Combo3 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         ItemData        =   "CJDYXT.frx":2405E5
         Left            =   1680
         List            =   "CJDYXT.frx":2405EF
         TabIndex        =   10
         Top             =   5640
         Width           =   3015
      End
      Begin VB.TextBox Text8 
         Height          =   375
         Left            =   1680
         MaxLength       =   20
         TabIndex        =   9
         Text            =   "Text8"
         Top             =   5200
         Width           =   3015
      End
      Begin VB.TextBox Text6 
         Height          =   435
         Left            =   1680
         TabIndex        =   7
         Text            =   "Text6"
         Top             =   4202
         Width           =   3015
      End
      Begin VB.ComboBox Combo2 
         Height          =   405
         ItemData        =   "CJDYXT.frx":240603
         Left            =   1680
         List            =   "CJDYXT.frx":240616
         TabIndex        =   6
         Top             =   3720
         Width           =   3015
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   1680
         TabIndex        =   4
         Top             =   2760
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         OLEDropMode     =   1
         Format          =   225116163
         CurrentDate     =   42430
         MaxDate         =   54788
         MinDate         =   2
      End
      Begin VB.ComboBox Combo1 
         Height          =   405
         ItemData        =   "CJDYXT.frx":240644
         Left            =   1680
         List            =   "CJDYXT.frx":24064E
         TabIndex        =   3
         Top             =   2296
         Width           =   3015
      End
      Begin VB.TextBox Text5 
         Enabled         =   0   'False
         Height          =   405
         Left            =   1680
         TabIndex        =   21
         Text            =   "<系统自动生成>"
         Top             =   360
         Width           =   3015
      End
      Begin VB.CommandButton Command2 
         Caption         =   "打  印"
         Height          =   615
         Left            =   2280
         TabIndex        =   12
         Top             =   6480
         Width           =   1695
      End
      Begin VB.CommandButton Command1 
         Caption         =   "添  加"
         Height          =   615
         Left            =   240
         TabIndex        =   11
         Top             =   6480
         Width           =   1575
      End
      Begin VB.TextBox Text3 
         Height          =   405
         Left            =   1680
         MaxLength       =   18
         TabIndex        =   2
         Text            =   "Text3"
         Top             =   1812
         Width           =   3015
      End
      Begin VB.TextBox Text2 
         Height          =   405
         Left            =   1680
         MaxLength       =   20
         TabIndex        =   1
         Text            =   "Text2"
         Top             =   1328
         Width           =   3015
      End
      Begin VB.TextBox Text1 
         Height          =   405
         Left            =   1680
         MaxLength       =   12
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   844
         Width           =   3015
      End
      Begin MSForms.Label Label16 
         Height          =   495
         Left            =   6240
         TabIndex        =   35
         Top             =   6480
         Visible         =   0   'False
         Width           =   2415
         BackColor       =   12648384
         Size            =   "4260;873"
         FontName        =   "幼圆"
         FontHeight      =   285
         FontCharSet     =   134
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "*结算方式："
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   120
         TabIndex        =   28
         Top             =   5640
         Width           =   1575
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "家庭住址："
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   27
         Top             =   5175
         Width           =   1455
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "民族："
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   840
         TabIndex        =   26
         Top             =   4200
         Width           =   855
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "婚姻状况："
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   25
         Top             =   3643
         Width           =   1455
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "*年龄："
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   720
         TabIndex        =   24
         Top             =   3180
         Width           =   1095
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "出生日期："
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   240
         TabIndex        =   23
         Top             =   2700
         Width           =   1575
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "*性别 ："
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   600
         TabIndex        =   22
         Top             =   2236
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "*患者编号："
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   120
         TabIndex        =   20
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "合作医疗号："
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   15
         Top             =   829
         Width           =   1935
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "联系电话："
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   18
         Top             =   4701
         Width           =   1575
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "身份证号："
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   17
         Top             =   1767
         Width           =   1455
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "*患者姓名："
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   1298
         Width           =   1815
      End
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   5160
      TabIndex        =   34
      Top             =   9720
      Width           =   1935
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   600
      TabIndex        =   33
      Top             =   9720
      Width           =   2175
   End
   Begin VB.Image Image1 
      Height          =   855
      Left            =   4080
      MousePointer    =   2  'Cross
      OLEDragMode     =   1  'Automatic
      Picture         =   "CJDYXT.frx":24065A
      Stretch         =   -1  'True
      Top             =   8280
      Visible         =   0   'False
      Width           =   855
   End
End
Attribute VB_Name = "新增档案"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_GotFocus()
List1.AddItem ("1")
List1.AddItem ("2")
List2.AddItem ("男")
List2.AddItem ("女")
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys "{tab}"
End If
End Sub

Private Sub Combo1_LostFocus()
List1.Clear
List2.Clear
End Sub

Private Sub Combo2_Change()
If Mid(Combo2.Text, 1, 1) = "1" Then
Combo2.Text = "已婚"
End If
If Mid(Combo2.Text, 1, 1) = "2" Then
Combo2.Text = "离婚"
End If
If Mid(Combo2.Text, 1, 1) = "3" Then
Combo2.Text = "未婚"
End If
If Mid(Combo2.Text, 1, 1) = "4" Then
Combo2.Text = "丧偶"
End If
If Mid(Combo2.Text, 1, 1) = "5" Then
Combo2.Text = "未说明的婚姻状况"
End If
End Sub

Private Sub Combo2_GotFocus()

List1.AddItem ("1")
List1.AddItem ("2")
List1.AddItem ("3")
List1.AddItem ("4")
List1.AddItem ("5")

List2.AddItem ("已婚")
List2.AddItem ("离婚")
List2.AddItem ("未婚")
List2.AddItem ("丧偶")
List2.AddItem ("未说明的婚姻状况")

End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys "{tab}"
End If
End Sub

Private Sub Combo2_LostFocus()
List1.Clear
List2.Clear
End Sub

Private Sub Combo3_Change()
If Mid(Combo3.Text, 1, 1) = "1" Then
Combo3.Text = "合作医疗"
End If
If Mid(Combo3.Text, 1, 1) = "2" Then
Combo3.Text = "自费"
End If
End Sub

Private Sub Combo3_GotFocus()
List2.AddItem ("合作医疗")
List2.AddItem ("自费")
List1.AddItem ("1")
List1.AddItem ("2")
End Sub

Private Sub Combo3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys "{tab}"
End If
End Sub

Private Sub Combo3_LostFocus()
List1.Clear
List2.Clear
End Sub

Private Sub Command1_Click()
On Error GoTo Err1
If Text2.Text <> "" And Combo1.Text <> "" And Combo3.Text <> "" Then
Adodc1.Recordset.AddNew
Adodc1.Recordset.Fields("患者编号") = Text5.Text
Adodc1.Recordset.Fields("合作医疗号") = Text1.Text
Adodc1.Recordset.Fields("患者姓名") = Text2.Text
Adodc1.Recordset.Fields("身份证号") = Text3.Text
Adodc1.Recordset.Fields("性别") = Combo1.Text
Adodc1.Recordset.Fields("出生日期") = DTPicker1.Value
Adodc1.Recordset.Fields("年龄") = Text4.Text
Adodc1.Recordset.Fields("婚姻状况") = Combo2.Text
Adodc1.Recordset.Fields("民族") = Text6.Text
Adodc1.Recordset.Fields("联系电话") = Text7.Text
Adodc1.Recordset.Fields("家庭住址") = Text8.Text
Adodc1.Recordset.Fields("结算方式") = Combo3.Text
Adodc1.Recordset.Fields("建档日期") = Date
Adodc1.Recordset.Update
DataGrid2.Refresh
Adodc1.Recordset.Update
'MDIForm1.StatusBar1.Panels(5) = Adodc1.Recordset.RecordCount
'Set Label15.DataSource = Adodc1
 '   Label15.DataField = "患者编号"
Else
MsgBox "必填内容不能为空！请填写！"
End If
Exit Sub
Err1:
 MsgBox "出现错误！" & vbCrLf & "错误编号：" & Err.Number & " 错误描述：" & Err.Description, 56
Resume Next
End Sub
Private Sub Command2_Click()
On Error Resume Next
Printer.Height = 20
Printer.Width = 15
Printer.PaperSize = 22
Printer.ScaleMode = vbCentimeters
Printer.Orientation = 1
Printer.FontSize = 16
Printer.CurrentX = 1
Printer.CurrentY = 1
Printer.Print Space(10) & "荒地镇卫生院"


Printer.PaintPicture Image1.Picture, 1, 1, 1.5, 1.5
Printer.PaintPicture Image1.Picture, 10, 1, 1.5, 1.5
Printer.FontSize = 12
Printer.CurrentX = 5
Printer.CurrentY = 2
Printer.Print "挂号单"


Printer.DrawStyle = 0   '以实线打印，VbDash 1 虚线 VbDot 2点线
                        'VbDashDot    3         点划线
                        'VbDashDotDot 4       双点划线
                        'VbInvisible  5           无线
                        'VbInsideSolid 6        内收实线

Printer.Line (2, 2.5)-(10, 2.5)

Printer.Line (2, 2.55)-(10, 2.55)

Printer.Line (2, 8.5)-(10, 8.5)


Printer.CurrentX = 2

Printer.CurrentY = 3

Printer.Print "合作医疗号：" & Space(4) & Text1.Text


Printer.CurrentX = 2

Printer.CurrentY = 4

Printer.Print "姓名：" & Space(10) & Text2.Text


Printer.CurrentX = 2

Printer.CurrentY = 5

Printer.Print "身份证号：" & Space(6) & Text3.Text


Printer.CurrentX = 2

Printer.CurrentY = 6

Printer.Print "挂号日期：" & Space(6) & Date


Printer.CurrentX = 2

Printer.CurrentY = 7

Printer.Print "挂号时间：" & Space(6) & Time


Printer.CurrentX = 2

Printer.CurrentY = 8

Printer.Print "挂号医生：" & Space(6) & Label16.Caption


Printer.CurrentX = 2

Printer.CurrentY = 9

Printer.Print "注：挂号单仅当日有效，下午5点自动作废！"


Printer.FontSize = 16

Printer.FontBold = True


Printer.CurrentX = 2

Printer.CurrentY = 10

Printer.Print "门诊号：" & Space(6) & Text5.Text


Printer.CurrentX = 10

Printer.CurrentY = 10

Printer.Print


Printer.EndDoc
End Sub

Private Sub Command3_Click()
If Adodc1.Recordset.BOF = True Then
Command3.Enabled = False
End If
Adodc1.Recordset.Delete
DataGrid2.Refresh
End Sub
Private Sub DTPicker1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys "{tab}"
End If
End Sub

Private Sub Form_Load()
VB.Clipboard.Clear
If Adodc1.Recordset.EOF = False Then
Adodc1.Recordset.MoveLast
End If
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
'DTPicker1.Value = ""
Combo2.Text = ""
MDIForm1.StatusBar1.Panels(5) = Adodc1.Recordset.RecordCount
Label16.Caption = MDIForm1.StatusBar1.Panels(3).Text
Text5.Text = Format(Now, "YYYYMMDDHHMMSS")
End Sub

Private Sub Label11_Click()
End
End Sub



Private Sub Text1_Change()
Dim HZYLH As String
HZYLH = Text1.Text
Dim conn As ADODB.Connection
'*定义一个记录集
Dim Mrc As ADODB.Recordset
'*分别实例化
Set conn = New ADODB.Connection
Set Mrc = New ADODB.Recordset
'*定义一个连接字符串
Dim ConnectString As String
ConnectString = "Provider=SQLOLEDB.1;password=sa;Persist Security Info=true;User ID=sa;Initial Catalog=ghgl;Data Source=NURA\SQLEXPRESS"
'*打开连接
conn.Open ConnectString
'*定义游标位置
conn.CursorLocation = adUseClient

If Len(HZYLH) = 10 Or Len(HZYLH) = 11 Then
'*查询记录集(从student表中找出名子为"张三"的记录)
Mrc.Open "select * from 参合名单 where 合作医疗号 like'%" & Text1.Text & "%'", conn, adOpenKeyset, adLockOptimistic
'*现在你已经得到了你想要查询的记录集了，那就是mrc
'*你可以把此记录集与DataGrid榜定，用datagrid显示你查询的记录
'添加到名单菜单里
    List1.Clear
    List2.Clear
    
    If Mrc.RecordCount <> 0 Then
    Mrc.MoveFirst
        Do While Not Mrc.EOF
            List2.AddItem (Mrc.Fields("姓名").Value)
'FIXIT: 用 "Mid$" 函数替换 "Mid" 函数                                                             FixIT90210ae-R9757-R1B8ZE
            List1.AddItem (Mid((Mrc.Fields("合作医疗号").Value), 11, 2))
            Mrc.MoveNext
        Loop
        Else
        List2.AddItem ("没有此合作医疗号")
        List2.AddItem ("请添加或重新输入")
        Set Text2.DataSource = Nothing
        Text2.DataField = ""
         Set Text3.DataSource = Nothing
        Text3.DataField = ""
     End If
     
End If

If Len(HZYLH) = 12 Then
Clipboard.Clear '清空剪贴板
Clipboard.SetText HZYLH

End If


If Len(Text1.Text) = 12 And Left(Text5.Text, 2) = "07" Then          '村信息赋值
Dim 编号 As String
编号 = Mid(HZYLH, 3, 2) & Right(HZYLH, 5)
              Text8.Text = "荒地镇" & Mid(Text1.Text, 3, 2) & "村 组"
              Text5.Text = 编号
              Combo3.Text = "合作医疗"
              End If
              
 If Len(Text1.Text) = 12 And Left(Text5.Text, 2) = "08" Then
 编号 = Left(HZYLH, 4) & Right(HZYLH, 5)
  Text8.Text = "墩巴格乡" & Mid(Text1.Text, 3, 2) & "村 组"
  Text5.Text = 编号
  Combo3.Text = "合作医疗"
   Else
 编号 = Left(HZYLH, 4) & Right(HZYLH, 5)
  Text8.Text = "外乡镇" & Mid(Text1.Text, 3, 2) & "村 组"
  Text5.Text = 编号
  Combo3.Text = "合作医疗"
  End If
  
  
Exit Sub
Err1:
  MsgBox "出现错误！" & vbCrLf & "错误编号：" & Err.Number & " 错误描述：" & Err.Description, 56

Resume Next
End Sub



Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys "{tab}"
End If
End Sub

Private Sub Text2_Change()
If Len(Text2.Text) >= 5 Then
 Text6.Text = "维吾尔族"
 Else
 Text6.Text = ""
End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys "{tab}"
End If
End Sub


Private Sub Text3_Change()
Dim y As Integer, Z As Integer

   
If Len(Text3.Text) = 18 Then    '处理身份信息
'FIXIT: 用 "Mid$" 函数替换 "Mid" 函数                                                             FixIT90210ae-R9757-R1B8ZE
y = Val(Mid(Text3.Text, 7, 4))

'FIXIT: 用 "Mid$" 函数替换 "Mid" 函数                                                             FixIT90210ae-R9757-R1B8ZE
c = Val(Mid(Text3.Text, 7, 8))

'FIXIT: 用 "Mid$" 函数替换 "Mid" 函数                                                             FixIT90210ae-R9757-R1B8ZE
Z = Val(Mid(Text3.Text, 17, 1))

'FIXIT: 用 "Mid$" 函数替换 "Mid" 函数                                                             FixIT90210ae-R9757-R1B8ZE
Text4.Text = Val(Mid$(Date, 1, 4)) - y

 DTPicker1.Value = CDate(Format(c, "####-##-##"))
 
     If Z Mod 2 = 1 Then
     Combo1.Text = "男"
     Else
     Combo1.Text = "女"
     End If
'FIXIT: 用 "Mid$" 函数替换 "Mid" 函数                                                             FixIT90210ae-R9757-R1B8ZE
              If Val(Mid(Text3.Text, 15, 2)) = 30 Or 32 Then
                Text8.Text = "荒地镇" + Text8.Text
              End If
              
End If

End Sub


Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text4.SetFocus
End If
End Sub

Private Sub Text4_Change()
If Val(Text4.Text) < 18 Then
Combo2.Text = "未婚"
End If
If Val(Text4.Text) > 25 And Val(Text4.Text) < 65 Then
Combo2.Text = "已婚"
End If
If Val(Text4.Text) >= 65 Then
Combo2.Text = ""
End If
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys "{tab}"
End If
End Sub

Private Sub Text4_LostFocus()
If Text3.Text = "" Then  '没有身份证号时启动自动生成出生日期函数
Dim N As Long
'FIXIT: 用早期绑定的数据类型声明 "M"                                                                   FixIT90210ae-R1672-R1B8ZE
Dim M
'FIXIT: 用早期绑定的数据类型声明 "B"                                                                   FixIT90210ae-R1672-R1B8ZE
Dim b
Randomize
M = Int(Rnd * (12 - 1 + 1)) + 1 '随机生成月份

If Len(M) = 1 Then                '双位化
M = "0" & M
End If

b = Int(Rnd * (29 - 1 + 1)) + 1   '随机生成日

If Len(b) = 1 Then             '双位化
b = "0" & b
End If

'FIXIT: 用 "Mid$" 函数替换 "Mid" 函数                                                             FixIT90210ae-R9757-R1B8ZE
N = Mid(Date, 1, 4) - Val(Text4.Text) & M & b

DTPicker1.Value = CDate(Format(N, "####-##-##")) '修改日期
End If
End Sub

Private Sub Text5_Change()
If Text5.Text = "" Then
Text5.Text = Format(Now, "YYYYMMDDHHMMSS")
End If
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys "{tab}"
End If
End Sub

Private Sub Text6_GotFocus()
List1.Width = List1.Width + 250
List2.AddItem ("维吾尔族")
List2.AddItem ("汉族")
List2.AddItem ("哈萨克族")
List2.AddItem ("回族")
List2.AddItem ("柯尔克孜族")
List2.AddItem ("蒙古族")
List2.AddItem ("塔吉克族")
List2.AddItem ("锡伯族")
List2.AddItem ("满族")
List2.AddItem ("乌孜别克族")
List2.AddItem ("俄罗斯族")
List2.AddItem ("达斡尔族")
List2.AddItem ("塔塔尔族")
List2.AddItem ("其他")
List1.AddItem ("1")
List1.AddItem ("2")
List1.AddItem ("3")
List1.AddItem ("4")
List1.AddItem ("5")
List1.AddItem ("6")
List1.AddItem ("7")
List1.AddItem ("8")
List1.AddItem ("9")
List1.AddItem ("10")
List1.AddItem ("11")
List1.AddItem ("12")
List1.AddItem ("13")
List1.AddItem ("14")
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys "{tab}"
End If
End Sub

Private Sub Text6_LostFocus()
If Val(Text6.Text) = 1 Then
Text6.Text = "维吾尔族"
End If
If Val(Text6.Text) = 2 Then
Text6.Text = "汉族"
End If
If Val(Text6.Text) = 3 Then
Text6.Text = "哈萨克族"
End If
If Val(Text6.Text) = 4 Then
Text6.Text = "回族"
End If
If Val(Text6.Text) = 5 Then
Text6.Text = "柯尔克孜族"
End If
If Val(Text6.Text) = 6 Then
Text6.Text = "蒙古族"
End If
If Val(Text6.Text) = 7 Then
Text6.Text = "塔吉克族"
End If
If Val(Text6.Text) = 8 Then
Text6.Text = "锡伯族"
End If
If Val(Text6.Text) = 9 Then
Text6.Text = "满族"
End If

If Val(Text6.Text) = 10 Then
Text6.Text = "乌孜别克族"
End If
If Val(Text6.Text) = 11 Then
Text6.Text = "俄罗斯族"
End If

If Val(Text6.Text) = 12 Then
Text6.Text = "达斡尔族"
End If
If Val(Text6.Text) = 13 Then
Text6.Text = "塔塔尔族"
End If
If Val(Text6.Text) = 13 Then
Text6.Text = "其他"
End If
List1.Width = List1.Width - 250
List1.Clear
List2.Clear
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys "{tab}"
End If
End Sub

Private Sub Text8_GotFocus()
Text8.SelLength = 3
Text8.SelStart = 3
End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys "{tab}"
End If
End Sub

Private Sub Text9_Change()
If Len(Text9.Text) = 12 Then
Text1.Text = Text9.Text
End If
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
MDIForm1.StatusBar1.Panels(5) = Adodc1.Recordset.RecordCount
If MDIForm1.StatusBar1.Panels(5) = 0 Then
Command3.Enabled = False
Else
Command3.Enabled = True
End If
Call Command1_Click
If Adodc2.Recordset.EOF Then
Timer1.Interval = 0
Else
Adodc2.Recordset.MoveNext
End If
End Sub

