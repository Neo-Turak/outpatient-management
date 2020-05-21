VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form 诊断参考 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "诊断参考管理"
   ClientHeight    =   8445
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9645
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
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "诊断参考管理.frx":0000
   ScaleHeight     =   422.25
   ScaleMode       =   2  'Point
   ScaleWidth      =   482.25
   Begin VB.CommandButton Command3 
      Caption         =   "删   除"
      Height          =   615
      Left            =   3480
      TabIndex        =   7
      Top             =   6720
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "保  存"
      Height          =   615
      Left            =   5760
      TabIndex        =   6
      Top             =   6720
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "添   加"
      Height          =   615
      Left            =   1200
      TabIndex        =   5
      Top             =   6720
      Width           =   1575
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "诊断参考管理.frx":3155B
      Height          =   4455
      Left            =   240
      TabIndex        =   4
      Top             =   120
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   7858
      _Version        =   393216
      BackColor       =   32768
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
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
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   "诊断"
         Caption         =   "诊断"
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
         DataField       =   "主诉"
         Caption         =   "主诉"
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
            ColumnWidth     =   172.488
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   221.244
         EndProperty
      EndProperty
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      DataSource      =   "Adodc1"
      Height          =   975
      Left            =   240
      TabIndex        =   3
      Top             =   5640
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   1720
      _Version        =   393217
      TextRTF         =   $"诊断参考管理.frx":31570
   End
   Begin VB.Frame Frame1 
      Caption         =   "项目"
      Height          =   855
      Left            =   240
      TabIndex        =   0
      Top             =   4680
      Width           =   2535
      Begin VB.OptionButton Option2 
         Caption         =   "主诉"
         Height          =   255
         Left            =   1350
         TabIndex        =   2
         Top             =   360
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         Caption         =   "诊断"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   1575
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   7440
      Top             =   6840
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
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
      RecordSource    =   "诊断参考"
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
   Begin VB.Label Label1 
      Height          =   855
      Left            =   240
      TabIndex        =   8
      Top             =   7440
      Width           =   9015
   End
End
Attribute VB_Name = "诊断参考"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MZ As String
Private Sub Command1_Click()
Dim RS As ADODB.Recordset
Set RS = Adodc1.Recordset
RS.AddNew
If Option1.Value = True And Not RichTextBox1.Text = "" Then RS!诊断 = RichTextBox1.Text
If Option2.Value = True And Not RichTextBox1.Text = "" Then RS!主诉 = RichTextBox1.Text
If Option3.Value = True And Not RichTextBox1.Text = "" Then RS!现病史 = RichTextBox1.Text
If Option4.Value = True And Not RichTextBox1.Text = "" Then RS!既往史 = RichTextBox1.Text
If Option5.Value = True And Not RichTextBox1.Text = "" Then RS!体格检查 = RichTextBox1.Text
If Option6.Value = True And Not RichTextBox1.Text = "" Then RS!辅助检查 = RichTextBox1.Text
'RS!MZ = RichTextBox1.Text
RS.Update
End Sub

Private Sub Command2_Click()
Dim RS As ADODB.Recordset
Set RS = Adodc1.Recordset
RS.Update
If Option1.Value = True And Not RichTextBox1.Text = "" Then RS!诊断 = RichTextBox1.Text
If Option2.Value = True And Not RichTextBox1.Text = "" Then RS!主诉 = RichTextBox1.Text
If Option3.Value = True And Not RichTextBox1.Text = "" Then RS!现病史 = RichTextBox1.Text
If Option4.Value = True And Not RichTextBox1.Text = "" Then RS!既往史 = RichTextBox1.Text
If Option5.Value = True And Not RichTextBox1.Text = "" Then RS!体格检查 = RichTextBox1.Text
If Option6.Value = True And Not RichTextBox1.Text = "" Then RS!辅助检查 = RichTextBox1.Text
RS.Update
RichTextBox1.Text = ""
End Sub

Private Sub Command3_Click()
Adodc1.Recordset.Delete
End Sub

Private Sub Form_Load()
Label1.Caption = "*添加：选择项目→填写内容→ <添  加>（只添加新内容到新栏）" & vbCrLf & "*删除：选中行→<删  除>（只删除选中行）" & vbCrLf & "*保存：选中项目→选中单元格→填写内容→<保  存>（只对单元格有效！）"
Me.Width = 9840
End Sub

Private Sub Form_Unload(Cancel As Integer)
 Adodc1.Recordset.Close
End Sub

Private Sub Option1_Click()
'RichTextBox1.DataField = Option1.Caption
End Sub
Private Sub Option2_Click()
'RichTextBox1.DataField = Option2.Caption
End Sub
Private Sub Option3_Click()
'RichTextBox1.DataField = Option3.Caption
End Sub
Private Sub Option4_Click()
'RichTextBox1.DataField = Option4.Caption
End Sub
Private Sub Option5_Click()
'RichTextBox1.DataField = Option5.Caption
End Sub
Private Sub Option6_Click()
'RichTextBox1.DataField = Option6.Caption
End Sub
