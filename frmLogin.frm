VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmLogin 
   BorderStyle     =   0  'None
   Caption         =   "��¼"
   ClientHeight    =   3210
   ClientLeft      =   2790
   ClientTop       =   3105
   ClientWidth     =   6630
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmLogin.frx":0000
   ScaleHeight     =   1896.574
   ScaleMode       =   0  'User
   ScaleWidth      =   6225.211
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.Label Label1 
      Caption         =   "��ӭʹ��ҽԺ����ϵͳ"
      Height          =   255
      Left            =   2520
      TabIndex        =   3
      Top             =   2760
      Visible         =   0   'False
      Width           =   1935
   End
   Begin MSForms.CommandButton cmdOK 
      Height          =   313
      Left            =   4080
      TabIndex        =   2
      Top             =   2280
      Width           =   855
      VariousPropertyBits=   19
      Size            =   "1508;552"
      FontName        =   "΢���ź�"
      FontHeight      =   240
      FontCharSet     =   134
      FontPitchAndFamily=   34
      ParagraphAlign  =   3
   End
   Begin MSForms.TextBox txtPassword 
      Height          =   315
      Left            =   2640
      TabIndex        =   1
      Top             =   1800
      Width           =   2295
      VariousPropertyBits=   746604563
      BackColor       =   -2147483646
      BorderStyle     =   1
      Size            =   "4048;556"
      PasswordChar    =   42
      SpecialEffect   =   0
      FontName        =   "����"
      FontHeight      =   240
      FontCharSet     =   134
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtUserName 
      Height          =   339
      Left            =   2640
      TabIndex        =   0
      Top             =   1337
      Width           =   2295
      VariousPropertyBits=   746604563
      BorderStyle     =   1
      Size            =   "4048;598"
      SpecialEffect   =   0
      FontName        =   "����"
      FontHeight      =   240
      FontCharSet     =   134
      FontPitchAndFamily=   34
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
'���ڽ�CreateRoundRectRgn������Բ�����򸳸���
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
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
Dim Voice As SpVoice
Public LoginSucceeded As Boolean

Private Sub cmdCancel_Click()
    '����ȫ�ֱ���Ϊ false
    '����ʾʧ�ܵĵ�¼
    LoginSucceeded = False
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    '�����ȷ������
    If TxtPassword = "password" Then
        '������������ﴫ��
        '�ɹ��� calling ����
        '����ȫ�ֱ���ʱ�����׵�
        LoginSucceeded = True
        Me.Hide
        ҽ������վ.Show
    Else
        MsgBox "��Ч�����룬������!", , "��¼"
        TxtPassword.SetFocus
        SendKeys "{Home}+{End}"
    End If
End Sub

Private Sub Form_Activate() '����Activate()�¼�
Call rgnform(Me, 50, 50) '�����ӹ���
End Sub


Private Sub Form_Load()
 Set Voice = New SpVoice
 Voice.Speak Label1.Caption, SVSFlagsAsync
End Sub

Private Sub Form_Unload(Cancel As Integer) '����Unload�¼�

DeleteObject outrgn '��Բ������ʹ�õ�����ϵͳ��Դ�ͷ�
End Sub
Private Sub rgnform(ByVal frmbox As Form, ByVal fw As Long, ByVal fh As Long) '�ӹ��̣��ı����fw��fh��ֵ��ʵ��Բ��
Dim w As Long, h As Long
w = frmbox.ScaleX(frmbox.Width, vbTwips, vbPixels)
h = frmbox.ScaleY(frmbox.Height, vbTwips, vbPixels)
outrgn = CreateRoundRectRgn(0, 0, w, h, fw, fh)
Call SetWindowRgn(frmbox.hwnd, outrgn, True)
End Sub

