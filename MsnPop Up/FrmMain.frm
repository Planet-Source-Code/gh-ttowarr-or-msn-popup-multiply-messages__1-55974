VERSION 5.00
Begin VB.Form FrmMain 
   BackColor       =   &H8000000E&
   Caption         =   "Msn Style PopUp 6.*"
   ClientHeight    =   3180
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   11145
   ControlBox      =   0   'False
   Icon            =   "FrmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3180
   ScaleWidth      =   11145
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtNumber 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   20
      Top             =   2600
      Width           =   255
   End
   Begin VB.TextBox TxtMessage 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   1335
      Left            =   2520
      MultiLine       =   -1  'True
      TabIndex        =   16
      Top             =   1560
      Width           =   2535
   End
   Begin VB.TextBox TxtOptions 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   2520
      TabIndex        =   14
      Top             =   1080
      Width           =   2535
   End
   Begin VB.TextBox TxtText 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   2520
      TabIndex        =   13
      Top             =   600
      Width           =   2535
   End
   Begin VB.CommandButton CmdClose 
      Appearance      =   0  'Flat
      Caption         =   "Close"
      Height          =   375
      Left            =   3840
      TabIndex        =   12
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton CmdPopUp 
      Appearance      =   0  'Flat
      Caption         =   "Send"
      Height          =   375
      Left            =   2520
      TabIndex        =   11
      Top             =   120
      Width           =   1215
   End
   Begin VB.OptionButton Option5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Option1"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   2160
      Width           =   255
   End
   Begin VB.OptionButton Option4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Option1"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1800
      Width           =   255
   End
   Begin VB.OptionButton Option3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Option1"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Width           =   255
   End
   Begin VB.OptionButton Option2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Option1"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   255
   End
   Begin VB.OptionButton Option1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   255
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Number of PopUp's"
      Height          =   255
      Left            =   480
      TabIndex        =   19
      Top             =   2640
      Width           =   1455
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "You can Select Every thing wat you want..(language). If you think this is cool for a trojan... It's your own risk"
      Height          =   1215
      Left            =   7800
      TabIndex        =   18
      Top             =   1080
      Width           =   3375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Hello, This is a little program like MSN PopUp"
      Height          =   375
      Left            =   7800
      TabIndex        =   17
      Top             =   840
      Width           =   3375
   End
   Begin VB.Label LblSound 
      BackColor       =   &H8000000E&
      Height          =   375
      Left            =   2520
      TabIndex        =   15
      Top             =   600
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "New Email"
      Height          =   255
      Left            =   480
      TabIndex        =   10
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "New Alert"
      Height          =   255
      Left            =   480
      TabIndex        =   9
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Type"
      Height          =   255
      Left            =   480
      TabIndex        =   8
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Ring"
      Height          =   255
      Left            =   480
      TabIndex        =   7
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Online"
      Height          =   255
      Left            =   480
      TabIndex        =   6
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Play Sound:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   240
      Width           =   855
   End
   Begin VB.Image ImgBgMSN 
      Height          =   2475
      Left            =   4800
      Picture         =   "FrmMain.frx":13B2
      Top             =   0
      Width           =   2700
   End
   Begin VB.Image ImgBG 
      Height          =   2475
      Left            =   0
      Picture         =   "FrmMain.frx":17000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11340
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdClose_Click()
End
End Sub

Private Sub CmdPopUp_Click()
If Option1.Value = True Then
LblSound.Caption = "online.wav"
ElseIf Option2.Value = True Then
LblSound.Caption = "ring.wav"
ElseIf Option3.Value = True Then
LblSound.Caption = "type.wav"
ElseIf Option4.Value = True Then
LblSound.Caption = "newalert.wav"
ElseIf Option5.Value = True Then
LblSound.Caption = "newemail.wav"
End If
FrmMSNPopUp.LblText.Caption = TxtText.Text
FrmMSNPopUp.LblMessage.Caption = TxtMessage.Text
FrmMSNPopUp.LblOptions.Caption = TxtOptions.Text
LoadPopUp
End Sub

Private Sub Form_Load()
Me.Width = 7875
End Sub

Sub LoadPopUp()
If TxtNumber.Text = "" Then
MsgBox "Enter a number of messages"
Exit Sub
End If
If TxtNumber.Text >= 1 Then
Set new1 = New FrmMSNPopUp
new1.SetNumber 450
new1.LblText.Caption = TxtText.Text
new1.LblMessage.Caption = TxtMessage.Text
new1.LblOptions.Caption = TxtOptions.Text
new1.Visible = True
End If
If TxtNumber.Text >= 2 Then
Set new2 = New FrmMSNPopUp
new2.SetNumber 450 + 1785
new2.LblText.Caption = TxtText.Text
new2.LblMessage.Caption = TxtMessage.Text
new2.LblOptions.Caption = TxtOptions.Text
new2.Visible = True
End If
If TxtNumber.Text >= 3 Then
Set new3 = New FrmMSNPopUp
new3.SetNumber 450 + 1785 * 2
new3.LblText.Caption = TxtText.Text
new3.LblMessage.Caption = TxtMessage.Text
new3.LblOptions.Caption = TxtOptions.Text
new3.Visible = True
End If
If TxtNumber.Text >= 4 Then
Set new4 = New FrmMSNPopUp
new4.SetNumber 450 + 1785 * 3
new4.LblText.Caption = TxtText.Text
new4.LblMessage.Caption = TxtMessage.Text
new4.LblOptions.Caption = TxtOptions.Text
new4.Visible = True
End If
If TxtNumber.Text >= 5 Then
Set new5 = New FrmMSNPopUp
new5.SetNumber 450 + 1785 * 4
new5.LblText.Caption = TxtText.Text
new5.LblMessage.Caption = TxtMessage.Text
new5.LblOptions.Caption = TxtOptions.Text
new5.Visible = True
End If
If TxtNumber.Text >= 6 Then
Set new6 = New FrmMSNPopUp
new6.SetNumber 450 + 1785 * 5
new6.LblText.Caption = TxtText.Text
new6.LblMessage.Caption = TxtMessage.Text
new6.LblOptions.Caption = TxtOptions.Text
new6.Visible = True
End If
End Sub
