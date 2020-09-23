VERSION 5.00
Begin VB.Form FrmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login"
   ClientHeight    =   2580
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3720
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2580
   ScaleWidth      =   3720
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   0
      Top             =   1320
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   720
      Top             =   840
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   360
      Top             =   840
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   0
      Top             =   840
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Remember my username."
      Height          =   255
      Left            =   1320
      TabIndex        =   3
      Top             =   960
      Width           =   2175
   End
   Begin VB.Frame Frame1 
      Height          =   135
      Left            =   120
      TabIndex        =   7
      Top             =   1800
      Width           =   3495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Login"
      Height          =   255
      Left            =   2400
      TabIndex        =   4
      Top             =   1440
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Cancel"
      Height          =   255
      Left            =   1320
      TabIndex        =   5
      Top             =   1440
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1320
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   600
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1320
      TabIndex        =   0
      Top             =   240
      Width           =   2175
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   360
      Picture         =   "Form1.frx":65AA
      Top             =   1080
      Width           =   480
   End
   Begin VB.Label Label3 
      Caption         =   "Please login above to access all services.  If you are not an admin, please exit this program now!"
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   2040
      Width           =   3495
   End
   Begin VB.Label Label2 
      Caption         =   "Password:"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Username:"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
End
End Sub

Private Sub Command2_Click()

Dim UserName, Password, Remember As String

UserName = GetSetting("Accounts", "Profile", "Username")
Password = GetSetting("Accounts", "Profile", "Password")


If Text1.Text = UserName Then
    GoTo chkpasswd
Else
    MsgBox "Username is incorrect.  Please try again.", vbExclamation, "Profiles"
    Text1.SetFocus
    Exit Sub
End If

chkpasswd:

If Text2.Text = Password Then
    SaveSetting "Accounts", "Profile", "Remember", Check1.Value
    FrmAdmin.Show
    Unload Me
Else
    MsgBox "Password is incorrect. Please try again.", vbExclamation, "Profiles"
    Text2.SetFocus
End If

End Sub

Private Sub Form_Load()


If App.PrevInstance = True Then
    MsgBox "You can only run 1 instance of 'Accounts Database' at a time.", vbExclamation, "Accounts Database"
    End
Else
    GoTo resumeagain
End If

resumeagain:

If Not GetSetting("Accounts", "Profile", "Logo") = "" Then
    GoTo resumeit
Else
    MsgBox "You must run 'New user setup.reg' file before attempting to run 'Accounts Database.exe", vbCritical, "Accounts Database"
    End
End If

resumeit:

If Not GetSetting("Accounts", "Profile", "Username") = "" Then
    GoTo nextpart
Else
    MsgBox "New account setup will now start, please click OK to proceed.", vbInformation, "Profiles"
    Timer3.Enabled = True
End If

nextpart:

If GetSetting("Accounts", "Profile", "AutoLogin") = "1" Then
    Timer4.Enabled = True
End If

If Not GetSetting("Accounts", "Profile", "Remember") = "1" Then
    Exit Sub
Else
    Check1.Value = 1
    Text1.Text = GetSetting("Accounts", "Profile", "Username")
    Timer1.Enabled = True
End If

End Sub
Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 13 Then
    Command2_Click
End If

End Sub

Private Sub Timer1_Timer()

Text2.SetFocus
Timer1.Enabled = False

End Sub

Private Sub Timer2_Timer()
FrmAdmin.Show
FrmAdmin.Label25.Caption = "Automaticly Login: Yes"
Unload Me
End Sub

Private Sub Timer3_Timer()
FrmSetup.Show
Timer3.Enabled = False
Unload Me
End Sub

Private Sub Timer4_Timer()
    FrmAdmin.Show
    FrmAdmin.Label25.Caption = "Automaticly Login: No"
    Unload Me
End Sub
