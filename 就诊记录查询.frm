VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form �����¼��ѯ 
   Caption         =   "�����¼��ѯ"
   ClientHeight    =   9375
   ClientLeft      =   3945
   ClientTop       =   2385
   ClientWidth     =   12060
   BeginProperty Font 
      Name            =   "����"
      Size            =   14.25
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "�����¼��ѯ.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9375
   ScaleWidth      =   12060
   Begin VB.Frame Frame4 
      Caption         =   "�����¼��"
      DragMode        =   1  'Automatic
      Height          =   8535
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   11775
      Begin RichTextLib.RichTextBox RichTextBox7 
         DataField       =   "ҽ������"
         DataSource      =   "Adodc2"
         Height          =   495
         Left            =   5640
         TabIndex        =   19
         Top             =   6360
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   873
         _Version        =   393217
         ReadOnly        =   -1  'True
         TextRTF         =   $"�����¼��ѯ.frx":08CA
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.TextBox Text1 
         DataField       =   "���"
         DataSource      =   "Adodc2"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5640
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   4
         Top             =   360
         Width           =   5655
      End
      Begin VB.CommandButton Command2 
         Caption         =   "�  ��"
         Height          =   495
         Left            =   1920
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   7560
         Width           =   1815
      End
      Begin VB.CommandButton Command1 
         Caption         =   "��  ӡ"
         Height          =   495
         Left            =   120
         TabIndex        =   2
         Top             =   7560
         Width           =   1575
      End
      Begin MSAdodcLib.Adodc Adodc2 
         Height          =   330
         Left            =   480
         Top             =   8160
         Visible         =   0   'False
         Width           =   2655
         _ExtentX        =   4683
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
         RecordSource    =   "���ﲡ��"
         Caption         =   "Adodc2"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSDataGridLib.DataGrid DataGrid5 
         Bindings        =   "�����¼��ѯ.frx":0967
         Height          =   6975
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   12303
         _Version        =   393216
         AllowUpdate     =   -1  'True
         AllowArrows     =   -1  'True
         BackColor       =   -2147483641
         ForeColor       =   8438015
         HeadLines       =   1
         RowHeight       =   15
         RowDividerStyle =   4
         FormatLocked    =   -1  'True
         AllowDelete     =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   12
         BeginProperty Column00 
            DataField       =   "���߱��"
            Caption         =   "���߱��"
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
            DataField       =   "���"
            Caption         =   "���"
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
            DataField       =   "����"
            Caption         =   "����"
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
            DataField       =   "�ֲ�ʷ"
            Caption         =   "�ֲ�ʷ"
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
            DataField       =   "����ʷ"
            Caption         =   "����ʷ"
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
            DataField       =   "�����"
            Caption         =   "�����"
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
            DataField       =   "�������"
            Caption         =   "�������"
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
            DataField       =   "�������"
            Caption         =   "�������"
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
            DataField       =   "ҽ������"
            Caption         =   "ҽ������"
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
            DataField       =   "����ҽʦ"
            Caption         =   "����ҽʦ"
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
            DataField       =   "��������"
            Caption         =   "��������"
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
            DataField       =   "����ʱ��"
            Caption         =   "����ʱ��"
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
               ColumnWidth     =   1200.189
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   720
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   959.811
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   929.764
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   2775.118
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   2775.118
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   2775.118
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   2775.118
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   2775.118
            EndProperty
            BeginProperty Column09 
               ColumnWidth     =   2775.118
            EndProperty
            BeginProperty Column10 
               ColumnWidth     =   1814.74
            EndProperty
            BeginProperty Column11 
               ColumnWidth     =   1574.929
            EndProperty
         EndProperty
      End
      Begin RichTextLib.RichTextBox RichTextBox4 
         DataField       =   "�����"
         DataSource      =   "Adodc2"
         Height          =   975
         Left            =   5640
         TabIndex        =   5
         Top             =   3840
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   1720
         _Version        =   393217
         BorderStyle     =   0
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         TextRTF         =   $"�����¼��ѯ.frx":097C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin RichTextLib.RichTextBox RichTextBox1 
         DataField       =   "�ֲ�ʷ"
         DataSource      =   "Adodc2"
         Height          =   1095
         Left            =   5640
         TabIndex        =   6
         Top             =   1680
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   1931
         _Version        =   393217
         BorderStyle     =   0
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         TextRTF         =   $"�����¼��ѯ.frx":0A19
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin RichTextLib.RichTextBox RichTextBox3 
         DataField       =   "����"
         DataSource      =   "Adodc2"
         Height          =   855
         Left            =   5640
         TabIndex        =   7
         Top             =   840
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   1508
         _Version        =   393217
         BorderStyle     =   0
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         TextRTF         =   $"�����¼��ѯ.frx":0AB6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin RichTextLib.RichTextBox RichTextBox2 
         DataField       =   "����ʷ"
         DataSource      =   "Adodc2"
         Height          =   1095
         Left            =   5640
         TabIndex        =   8
         Top             =   2760
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   1931
         _Version        =   393217
         BorderStyle     =   0
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         TextRTF         =   $"�����¼��ѯ.frx":0B53
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin RichTextLib.RichTextBox RichTextBox5 
         DataField       =   "�������"
         DataSource      =   "Adodc2"
         Height          =   975
         Left            =   5640
         TabIndex        =   9
         Top             =   4800
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   1720
         _Version        =   393217
         BorderStyle     =   0
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         TextRTF         =   $"�����¼��ѯ.frx":0BF0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin RichTextLib.RichTextBox RichTextBox6 
         DataField       =   "�������"
         DataSource      =   "Adodc2"
         Height          =   495
         Left            =   5640
         TabIndex        =   10
         Top             =   5760
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   873
         _Version        =   393217
         BorderStyle     =   0
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         TextRTF         =   $"�����¼��ѯ.frx":0C8D
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label4 
         DataField       =   "����ҽʦ"
         DataSource      =   "Adodc2"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7680
         TabIndex        =   24
         Top             =   7920
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "����ҽʦ"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6480
         TabIndex        =   23
         Top             =   7920
         Width           =   1215
      End
      Begin MSForms.Label Label11 
         Height          =   495
         Index           =   1
         Left            =   5040
         TabIndex        =   22
         Top             =   7320
         Width           =   1095
         BackColor       =   16777215
         VariousPropertyBits=   8388627
         Caption         =   "����ʱ��"
         Size            =   "1931;873"
         FontName        =   "����"
         FontHeight      =   240
         FontCharSet     =   134
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label2 
         DataField       =   "����ʱ��"
         DataSource      =   "Adodc2"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7560
         TabIndex        =   21
         Top             =   7320
         Width           =   2055
      End
      Begin VB.Label Label1 
         DataField       =   "��������"
         DataSource      =   "Adodc2"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6120
         TabIndex        =   20
         Top             =   7320
         Width           =   2055
      End
      Begin MSForms.Label Label11 
         Height          =   495
         Index           =   0
         Left            =   4800
         TabIndex        =   18
         Top             =   6360
         Width           =   615
         BackColor       =   16777215
         VariousPropertyBits=   8388627
         Caption         =   "ҽ������"
         Size            =   "1085;873"
         FontName        =   "����"
         FontHeight      =   240
         FontCharSet     =   134
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label10 
         Height          =   615
         Index           =   4
         Left            =   4800
         TabIndex        =   17
         Top             =   5760
         Width           =   615
         BackColor       =   16777215
         VariousPropertyBits=   8388627
         Caption         =   "�������"
         Size            =   "1085;1085"
         FontName        =   "����"
         FontHeight      =   240
         FontCharSet     =   134
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label10 
         Height          =   615
         Index           =   3
         Left            =   4800
         TabIndex        =   16
         Top             =   5040
         Width           =   615
         BackColor       =   16777215
         VariousPropertyBits=   8388627
         Caption         =   "�������"
         Size            =   "1085;1085"
         FontName        =   "����"
         FontHeight      =   240
         FontCharSet     =   134
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label10 
         Height          =   615
         Index           =   2
         Left            =   4800
         TabIndex        =   15
         Top             =   4080
         Width           =   615
         BackColor       =   16777215
         VariousPropertyBits=   8388627
         Caption         =   "�����"
         Size            =   "1085;1085"
         FontName        =   "����"
         FontHeight      =   240
         FontCharSet     =   134
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label10 
         Height          =   495
         Index           =   1
         Left            =   4680
         TabIndex        =   14
         Top             =   3120
         Width           =   1215
         BackColor       =   16777215
         VariousPropertyBits=   8388627
         Caption         =   "����ʷ��"
         Size            =   "2143;873"
         FontName        =   "����"
         FontHeight      =   285
         FontCharSet     =   134
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label10 
         Height          =   495
         Index           =   0
         Left            =   4680
         TabIndex        =   13
         Top             =   1920
         Width           =   1335
         BackColor       =   16777215
         VariousPropertyBits=   8388627
         Caption         =   "�ֲ�ʷ��"
         Size            =   "2355;873"
         FontName        =   "����"
         FontHeight      =   285
         FontCharSet     =   134
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label16 
         Caption         =   "���ߣ�"
         Height          =   495
         Left            =   4800
         TabIndex        =   12
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label14 
         Caption         =   "���:"
         Height          =   375
         Left            =   4800
         TabIndex        =   11
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Label Label6 
      Height          =   375
      Left            =   2160
      TabIndex        =   26
      Top             =   120
      Width           =   9615
   End
   Begin VB.Label Label5 
      Caption         =   "������Ϣ��"
      Height          =   375
      Left            =   360
      TabIndex        =   25
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "�����¼��ѯ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()
Me.Adodc2.Recordset.Update
End Sub

Private Sub Form_Load()
Me.Width = 12300
Me.Height = 9300
End Sub

Private Sub Form_Resize()
Me.Width = 12500
Me.Height = 9500
End Sub
