VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmLogin 
   BorderStyle     =   0  'None
   Caption         =   "登录"
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
   StartUpPosition =   2  '屏幕中心
   Begin VB.Label Label1 
      Caption         =   "欢迎使用医院管理系统"
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
      FontName        =   "微软雅黑"
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
      FontName        =   "宋体"
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
      FontName        =   "宋体"
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
'用于将CreateRoundRectRgn创建的圆角区域赋给窗
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
'用于创建一个圆角矩形，该矩形由X1，Y1-X2，Y2确定，并由X3，Y3确定的椭圆描述圆角弧度。
'参数 类型及说明：
'X1,Y1 Long，矩形左上角的X，Y坐标
'X2,Y2 Long，矩形右下角的X，Y坐标
'X3 Long，圆角椭圆的宽。其范围从0（没有圆角）到矩形宽（全圆）
'Y3 Long，圆角椭圆的高。其范围从0（没有圆角）到矩形高（全圆）
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
'将CreateRoundRectRgn创建的区域删除，这是必要的，否则不必要的占用电脑内存
Dim outrgn As Long
'接下来声明一个全局变量,用来获得区域句柄
Option Explicit
Dim Voice As SpVoice
Public LoginSucceeded As Boolean

Private Sub cmdCancel_Click()
    '设置全局变量为 false
    '不提示失败的登录
    LoginSucceeded = False
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    '检查正确的密码
    If TxtPassword = "password" Then
        '将代码放在这里传递
        '成功到 calling 函数
        '设置全局变量时最容易的
        LoginSucceeded = True
        Me.Hide
        医技工作站.Show
    Else
        MsgBox "无效的密码，请重试!", , "登录"
        TxtPassword.SetFocus
        SendKeys "{Home}+{End}"
    End If
End Sub

Private Sub Form_Activate() '窗体Activate()事件
Call rgnform(Me, 50, 50) '调用子过程
End Sub


Private Sub Form_Load()
 Set Voice = New SpVoice
 Voice.Speak Label1.Caption, SVSFlagsAsync
End Sub

Private Sub Form_Unload(Cancel As Integer) '窗体Unload事件

DeleteObject outrgn '将圆角区域使用的所有系统资源释放
End Sub
Private Sub rgnform(ByVal frmbox As Form, ByVal fw As Long, ByVal fh As Long) '子过程，改变参数fw和fh的值可实现圆角
Dim w As Long, h As Long
w = frmbox.ScaleX(frmbox.Width, vbTwips, vbPixels)
h = frmbox.ScaleY(frmbox.Height, vbTwips, vbPixels)
outrgn = CreateRoundRectRgn(0, 0, w, h, fw, fh)
Call SetWindowRgn(frmbox.hwnd, outrgn, True)
End Sub

