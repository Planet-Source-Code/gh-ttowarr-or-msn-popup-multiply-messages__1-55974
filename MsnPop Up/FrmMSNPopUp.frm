VERSION 5.00
Begin VB.Form FrmMSNPopUp 
   BorderStyle     =   0  'None
   ClientHeight    =   1785
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3900
   LinkTopic       =   "Form1"
   ScaleHeight     =   1785
   ScaleWidth      =   3900
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer TmrUnloader 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3000
      Top             =   1320
   End
   Begin VB.Timer TmrUnload 
      Interval        =   10000
      Left            =   3000
      Top             =   840
   End
   Begin VB.Timer TmrMove 
      Interval        =   1
      Left            =   3000
      Top             =   360
   End
   Begin VB.Image ImgClose3 
      Height          =   225
      Left            =   3480
      Picture         =   "FrmMSNPopUp.frx":0000
      Top             =   120
      Width           =   225
   End
   Begin VB.Label LblOptions 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00400000&
      Height          =   255
      Left            =   1400
      MouseIcon       =   "FrmMSNPopUp.frx":0312
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   405
      Width           =   1215
   End
   Begin VB.Label LblMessage 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000000&
      Height          =   1215
      Left            =   120
      MouseIcon       =   "FrmMSNPopUp.frx":061C
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   480
      Width           =   2535
   End
   Begin VB.Label LblText 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   90
      Width           =   1935
   End
   Begin VB.Image ImgClose2 
      Height          =   225
      Left            =   3240
      Picture         =   "FrmMSNPopUp.frx":0926
      Top             =   120
      Width           =   225
   End
   Begin VB.Image ImgClose1 
      Height          =   225
      Left            =   3000
      Picture         =   "FrmMSNPopUp.frx":0C38
      Top             =   120
      Width           =   225
   End
   Begin VB.Image ImgClose 
      Height          =   225
      Left            =   2400
      Top             =   90
      Width           =   225
   End
   Begin VB.Image ImgMsnBG 
      Height          =   1740
      Left            =   0
      Picture         =   "FrmMSNPopUp.frx":0F4A
      Top             =   0
      Width           =   2715
   End
End
Attribute VB_Name = "FrmMSNPopUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Number As Integer

Private Sub Form_Load()
Me.Width = ImgMsnBG.Width
Me.Height = ImgMsnBG.Height
ImgClose.Picture = ImgClose1.Picture
Me.Top = Screen.Height
Me.Left = Screen.Width - Me.Width - 220
ImgClose.Picture = ImgClose1.Picture
LblOptions.FontUnderline = False
End Sub

Public Sub SetNumber(Number1 As String)
Number = Number1
End Sub

Private Sub ImgClose_Click()
Unload Me
End Sub

Private Sub ImgClose_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ImgClose.Picture = ImgClose3.Picture
End Sub

Private Sub ImgClose_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ImgClose.Picture = ImgClose2.Picture
End Sub

Private Sub ImgMsnBG_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ImgClose.Picture = ImgClose1.Picture
LblOptions.FontUnderline = False
End Sub

Private Sub LblMessage_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
LblOptions.FontUnderline = False
End Sub

Private Sub LblOptions_Click()
ImgClose.Picture = ImgClose1.Picture
LblOptions.FontUnderline = False
PostMessage MSN_FindMSN, WM_COMMAND, 40268, 0
End Sub

Private Sub LblOptions_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
LblOptions.FontUnderline = True
End Sub

Private Sub LblText_Click()
ImgClose.Picture = ImgClose1.Picture
LblOptions.FontUnderline = False
End Sub

Private Sub TmrMove_Timer()
If Me.Top <= Screen.Height - Me.Height - Number Then
TmrMove.Enabled = False
AlwaysOnTop.AlwaysOnTop Me, True
PlaySound (App.Path & "\" & FrmMain.LblSound.Caption)
Else
Me.Top = Me.Top - 50
End If
End Sub

Private Sub TmrUnload_Timer()
Unload Me
End Sub

