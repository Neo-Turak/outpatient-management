VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form ������ѯ 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�Һ�"
   ClientHeight    =   6990
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8145
   BeginProperty Font 
      Name            =   "����"
      Size            =   12
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "������ѯ.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "������ѯ.frx":014A
   ScaleHeight     =   6990
   ScaleWidth      =   8145
   ShowInTaskbar   =   0   'False
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   5400
      Top             =   2520
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
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   375
      Left            =   5280
      Top             =   6240
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
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
      RecordSource    =   "�Һŵ�"
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
   Begin MSDataGridLib.DataGrid DataGrid2 
      Bindings        =   "������ѯ.frx":6F8B
      Height          =   1335
      Left            =   120
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   5520
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   2355
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   -2147483643
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
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
      ColumnCount     =   9
      BeginProperty Column00 
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
      BeginProperty Column01 
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
      BeginProperty Column02 
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
      BeginProperty Column03 
         DataField       =   "�Ա�"
         Caption         =   "�Ա�"
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
      BeginProperty Column05 
         DataField       =   "���㷽ʽ"
         Caption         =   "���㷽ʽ"
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
         DataField       =   "�Һ�����"
         Caption         =   "�Һ�����"
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
         DataField       =   "�Һ�ʱ��"
         Caption         =   "�Һ�ʱ��"
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
         DataField       =   "״̬"
         Caption         =   "״̬"
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
            ColumnWidth     =   615.118
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1049.953
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1049.953
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   629.858
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   615.118
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1094.74
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   30.047
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1080
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   734.74
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command2 
      Caption         =   "����"
      Enabled         =   0   'False
      Height          =   495
      Left            =   4200
      TabIndex        =   5
      Top             =   4560
      Width           =   1095
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   1695
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   2990
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   0
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
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
      ColumnCount     =   15
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
         DataField       =   "����ҽ�ƺ�"
         Caption         =   "����ҽ�ƺ�"
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
      BeginProperty Column03 
         DataField       =   "���֤��"
         Caption         =   "���֤��"
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
         DataField       =   "�Ա�"
         Caption         =   "�Ա�"
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
      BeginProperty Column06 
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
      BeginProperty Column07 
         DataField       =   "����״��"
         Caption         =   "����״��"
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
      BeginProperty Column09 
         DataField       =   "��ϵ�绰"
         Caption         =   "��ϵ�绰"
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
         DataField       =   "��ͥסַ"
         Caption         =   "��ͥסַ"
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
         DataField       =   "���޹���ʷ"
         Caption         =   "���޹���ʷ"
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
         DataField       =   "���㷽ʽ"
         Caption         =   "���㷽ʽ"
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
      BeginProperty Column14 
         DataField       =   "��ע"
         Caption         =   "��ע"
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
            ColumnWidth     =   1110.047
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1289.764
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1080
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1124.787
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   585.071
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1814.74
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1454.74
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1814.74
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   1814.74
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   1934.929
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   1814.74
         EndProperty
         BeginProperty Column11 
            ColumnWidth     =   1214.929
         EndProperty
         BeginProperty Column12 
            ColumnWidth     =   1574.929
         EndProperty
         BeginProperty Column13 
            ColumnWidth     =   1814.74
         EndProperty
         BeginProperty Column14 
            ColumnWidth     =   2775.118
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "��ѯ"
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7335
      Begin VB.CommandButton Command3 
         Caption         =   "��ӵ���"
         Height          =   495
         Left            =   6000
         TabIndex        =   6
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "����"
         Height          =   495
         Left            =   4800
         TabIndex        =   3
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   2280
         TabIndex        =   2
         Text            =   "�ؼ���"
         Top             =   480
         Width           =   2415
      End
      Begin VB.ComboBox Combo1 
         Height          =   360
         ItemData        =   "������ѯ.frx":6FA0
         Left            =   120
         List            =   "������ѯ.frx":6FB0
         TabIndex        =   1
         Text            =   "����λ��"
         Top             =   480
         Width           =   2175
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "������Ϣ"
      Height          =   2175
      Left            =   120
      TabIndex        =   7
      Top             =   3120
      Width           =   7335
      Begin VB.ComboBox Combo2 
         Height          =   360
         ItemData        =   "������ѯ.frx":6FDE
         Left            =   1800
         List            =   "������ѯ.frx":6FE8
         TabIndex        =   17
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "�����㷽ʽ��"
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   360
         TabIndex        =   16
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         DataField       =   "����"
         DataSource      =   "Adodc1"
         Height          =   255
         Left            =   5520
         TabIndex        =   15
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         DataField       =   "�Ա�"
         DataSource      =   "Adodc1"
         Height          =   255
         Left            =   5520
         TabIndex        =   14
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label5 
         DataField       =   "��������"
         DataSource      =   "Adodc1"
         Height          =   255
         Left            =   1680
         TabIndex        =   13
         Top             =   960
         Width           =   2775
      End
      Begin VB.Label Label4 
         DataField       =   "����ҽ�ƺ�"
         DataSource      =   "Adodc1"
         Height          =   255
         Left            =   1680
         TabIndex        =   12
         Top             =   480
         Width           =   2775
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "�����䡿"
         Height          =   255
         Left            =   4560
         TabIndex        =   11
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "���Ա�"
         Height          =   255
         Left            =   4560
         TabIndex        =   10
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "������������"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   9
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "������ҽ�ƺš�"
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   8
         Top             =   480
         Width           =   1695
      End
   End
End
Attribute VB_Name = "������ѯ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo2_Change()
Command2.Enabled = True
End Sub

Private Sub Command1_Click()
On Error Resume Next
Dim SS As String
SS = Combo1.Text
Dim conn As ADODB.Connection
Dim Mrc As ADODB.Recordset
Set conn = New ADODB.Connection
Set Mrc = New ADODB.Recordset
Dim ConnectString As String
ConnectString = "Provider=SQLOLEDB.1;password=sa;Persist Security Info=true;User ID=sa;Initial Catalog=ghgl;Data Source=NURA\SQLEXPRESS"
'*������
conn.Open ConnectString
'*�����α�λ��
conn.CursorLocation = adUseClient
Mrc.Open "select * from �����ܱ� where " & Combo1.Text & " like'%" & Text1.Text & "%'", conn, adOpenKeyset, adLockOptimistic
Set Adodc1.Recordset = Mrc
Set DataGrid1.DataSource = Mrc
Adodc1.Refresh
Command2.SetFocus
End Sub

Private Sub Command2_Click()
If Combo2.Text <> "" Then
����վ.Label19.Caption = Label4.Caption
����վ.Label2.Caption = Label5.Caption
����վ.Label3.Caption = Label6.Caption
����վ.Label5.Caption = Label7.Caption
����վ.Label7.Caption = Combo2.Text
Else
MsgBox "������Ϣ����Ϊ�ջ�ѡ����㷽ʽ", vbInformation, "��Ч����"
Exit Sub
End If
����վ.SSTab2.Tab = 1
'����վ.Text1.SetFocus
Unload Me
End Sub

Private Sub Command3_Click()
��������.Show
End Sub

Private Sub Form_Unload(Cancel As Integer)
'Adodc1.Recordset.Close
'Adodc2.Recordset.Close
End Sub

Private Sub Label4_Change()
Me.Command2.Enabled = True
End Sub

Private Sub Label5_Change()
Me.Command2.Enabled = True
End Sub
Private Sub Text1_Change()
If Left(Text1.Text, 1) = "0" And Len(Text1.Text) >= 12 Then
Combo1.Text = "����ҽ�ƺ�"
Combo2.Text = "����ҽ��"
Call Command1_Click
End If
If Left(Text1.Text, 1) = "6" And Len(Text1.Text) = 18 Then
Combo1.Text = "���֤��"
Combo2.Text = "�Է�"
Call Command1_Click
End If
End Sub

Private Sub Text1_GotFocus()
Text1.SelStart = 0
Text1.SelLength = Len(Text1.Text)
End Sub
