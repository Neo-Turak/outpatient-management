VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmLogin1 
   BackColor       =   &H00FFFF00&
   BorderStyle     =   0  'None
   Caption         =   "��¼"
   ClientHeight    =   4665
   ClientLeft      =   6795
   ClientTop       =   4350
   ClientWidth     =   5625
   Icon            =   "frmLogin1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmLogin1.frx":1082
   ScaleHeight     =   4665
   ScaleWidth      =   5625
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   1920
      Top             =   4320
      Visible         =   0   'False
      Width           =   2040
      _ExtentX        =   3598
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
      Connect         =   "Provider=SQLOLEDB.1;Password=sa;Persist Security Info=True;User ID=sa;Initial Catalog=ghgl;Data Source=nura\sqlexpress"
      OLEDBString     =   "Provider=SQLOLEDB.1;Password=sa;Persist Security Info=True;User ID=sa;Initial Catalog=ghgl;Data Source=nura\sqlexpress"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "�û���"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton CmdCancel 
      Caption         =   "ȡ��"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3600
      TabIndex        =   3
      Top             =   3600
      Width           =   1335
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "ȷ��"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1320
      TabIndex        =   2
      Top             =   3600
      Width           =   1335
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "����ҽ������վ"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   12
      Top             =   1200
      Width           =   2055
   End
   Begin VB.Label Label7 
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   4320
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label6 
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   4080
      Visible         =   0   'False
      Width           =   975
   End
   Begin MSForms.CommandButton CommandButton1 
      Height          =   375
      Left            =   860
      TabIndex        =   9
      Top             =   420
      Width           =   375
      VariousPropertyBits=   19
      Size            =   "661;661"
      FontName        =   "����"
      FontHeight      =   180
      FontCharSet     =   134
      FontPitchAndFamily=   34
      ParagraphAlign  =   3
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      DataField       =   "����"
      DataSource      =   "Adodc1"
      Height          =   255
      Left            =   2160
      TabIndex        =   8
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      DataField       =   "����"
      BeginProperty Font 
         Name            =   "�����п�"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      TabIndex        =   7
      Top             =   2880
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      DataField       =   "�û���"
      BeginProperty Font 
         Name            =   "��Բ"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   6
      Top             =   1811
      Width           =   2295
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      DataField       =   "ְλ"
      BeginProperty Font 
         Name            =   "�����п�"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   5
      Top             =   2880
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      DataField       =   "����"
      Height          =   375
      Left            =   3480
      TabIndex        =   4
      Top             =   2280
      Visible         =   0   'False
      Width           =   1335
   End
   Begin MSForms.TextBox TxtPassword 
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Top             =   2319
      Width           =   1695
      VariousPropertyBits=   746604563
      Size            =   "2990;661"
      PasswordChar    =   42
      SpecialEffect   =   0
      FontName        =   "����"
      FontHeight      =   240
      FontCharSet     =   134
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox TxtUserName 
      Height          =   375
      Left            =   1440
      TabIndex        =   0
      Top             =   1777
      Width           =   855
      VariousPropertyBits=   746604563
      Size            =   "1508;661"
      SpecialEffect   =   0
      FontName        =   "����"
      FontHeight      =   240
      FontCharSet     =   134
      FontPitchAndFamily=   34
   End
End
Attribute VB_Name = "frmLogin1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SkinH_AttachEx Lib "D:\Users\NURA\vb 37��Ƥ��\SkinH_VB6.dll" (ByVal lpSkinFile As String, ByVal lpPasswd As String) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
'���ڽ�CreateRoundRectRgn������Բ�����򸳸�����
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
'���ڴ���һ��Բ�Ǿ��Σ��þ�����X1��Y1-X2��Y2ȷ��������X3��Y3ȷ������Բ����Բ�ǻ��ȡ�
'���� ���ͼ�˵����
'X1,Y1 Long���������Ͻǵ�X��Y����
'X2,Y2 Long���������½ǵ�X��Y����
'X3 Long��Բ����Բ�Ŀ��䷶Χ��0��û��Բ�ǣ������ο�ȫԲ��
'Y3 Long��Բ����Բ�ĸߡ��䷶Χ��0��û��Բ�ǣ������θߣ�ȫԲ��
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
'��CreateRoundRectRgn����������ɾ�������Ǳ�Ҫ�ģ����򲻱�Ҫ��ռ�õ����ڴ�
Dim outrgn As Long
'����������һ��ȫ�ֱ���,�������������
Option Explicit
Public LoginSucceeded As Boolean
Private Sub cmdCancel_Click()
    '����ȫ�ֱ���Ϊ false
    '����ʾʧ�ܵĵ�¼
    LoginSucceeded = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
Dim a As String

If Label1.Caption = TxtPassword.Text Then

        If Label4.Caption = "����" And Label2.Caption = "ҽ��" Then
        Me.Hide
        MDIForm1.Show
        MDIForm1.StatusBar1.Panels(3) = Label3.Caption
        MDIForm1.StatusBar1.Panels(4) = Label4.Caption
        MDIForm1.StatusBar1.Panels(5) = Label2.Caption
        End If
        
        If Label4.Caption = "����" And Label2.Caption = "ҽѧӰ����" Then
        Me.Hide
        ҽ������վ.Show
        MDIForm1.StatusBar1.Panels(3) = Label3.Caption
        MDIForm1.StatusBar1.Panels(4) = Label4.Caption
        MDIForm1.StatusBar1.Panels(5) = Label2.Caption
        End If
        If Not Label4.Caption = "����" Then
         a = Label4.Caption & "" & Label2.Caption & " ����վ"
        MsgBox "��ʹ��" & a, vbExclamation, "�����Ͽͻ���"
        End If
        
  Else
  MsgBox "��Ч��������û�����������!", , "��¼"
  TxtUserName.SetFocus
  SendKeys "{Home}+{End}"
  End If

End Sub

Private Sub CommandButton1_Click()
Dim x As Integer, y As Integer, Z As Integer
Z = (Me.Width - 4755) / 2
y = Me.Width / 2
x = Me.Height / 2   '�߶�
frmADODBLogon.Left = Me.Left + Z
frmADODBLogon.Top = VB.Screen.Height / 2 + x
frmADODBLogon.Show
End Sub

Private Sub Form_Activate() '����Activate()�¼�
Call rgnform(Me, 50, 50) '�����ӹ���
'SkinH_AttachEx "D:\Users\NURA\Desktop\���Ӳ���\Ƥ��\��Ө���.she", "" 'Ƥ������
End Sub
Private Sub Form_Load()

If App.PrevInstance = True Then
On Error Resume Next
End If
SkinH_AttachEx App.Path & "/Ƥ��/��ɫ����.she", ""
Dim x As Integer, y As Integer
x = Screen.Width / Screen.TwipsPerPixelX
y = Screen.Height / Screen.TwipsPerPixelY
Label6.Caption = x
Label7.Caption = y
End Sub

Private Sub Form_Unload(Cancel As Integer) '����Unload�¼�
DeleteObject outrgn '��Բ������ʹ�õ�����ϵͳ��Դ�ͷ�
End Sub
Private Sub rgnform(ByVal frmbox As Form, ByVal fw As Long, ByVal fh As Long) '�ӹ��̣��ı����fw��fh��ֵ��ʵ��Բ��
Dim W As Long, h As Long
W = frmbox.ScaleX(frmbox.Width, vbTwips, vbPixels)
h = frmbox.ScaleY(frmbox.Height, vbTwips, vbPixels)
outrgn = CreateRoundRectRgn(0, 0, W, h, fw, fh)
Call SetWindowRgn(frmbox.hwnd, outrgn, True)
End Sub

Private Sub TxtUserName_LostFocus()
Dim conn As ADODB.Connection
Dim Mrc As ADODB.Recordset
Set conn = New ADODB.Connection
Set Mrc = New ADODB.Recordset
Dim ConnectString As String
ConnectString = "Provider=SQLOLEDB.1;Password=sa;Persist Security Info=True;User ID=sa;Initial Catalog=ghgl;Data Source=NURA\SQLEXPRESS"
conn.Open ConnectString
conn.CursorLocation = adUseClient
Mrc.Open "select * from �û��� where ID='" & TxtUserName.Text & "'", conn, adOpenKeyset, adLockOptimistic
    Set Label1.DataSource = Mrc
    Set Label2.DataSource = Mrc
    Set Label3.DataSource = Mrc
    Set Label2.DataSource = Mrc
    Set Label4.DataSource = Mrc
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim strExit As String
If UnloadMode <> vbAppWindows Then
strExit = "��ȷ��Ҫֹͣ������"
If vbNo = MsgBox(strExit, vbQuestion Or vbYesNo, "") Then
Cancel = True
Exit Sub
End If
End If
End
End Sub
