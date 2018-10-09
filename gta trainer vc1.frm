VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Grand Theft Auto :: Vice City PC Trainer [www.MMTbb.com]"
   ClientHeight    =   6975
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   9330
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   14.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "gta trainer vc1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "gta trainer vc1.frx":0CCA
   ScaleHeight     =   6975
   ScaleWidth      =   9330
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   5400
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   720
      Width           =   3495
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   7920
      Top             =   3120
   End
   Begin VB.Label lblL 
      BackStyle       =   0  'Transparent
      Caption         =   "L] You'll Never Get Me Mode"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   5880
      TabIndex        =   18
      ToolTipText     =   "Keeps your wanted level at six stars."
      Top             =   6120
      Width           =   3375
   End
   Begin VB.Label lblK 
      BackStyle       =   0  'Transparent
      Caption         =   "K] Leave Me Alone Mode"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   5880
      TabIndex        =   17
      ToolTipText     =   "This mode will cause your wanted level to stay at zero."
      Top             =   5760
      Width           =   3375
   End
   Begin VB.Label lblP 
      BackStyle       =   0  'Transparent
      Caption         =   "P] Pursuit Mode"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   5880
      TabIndex        =   16
      ToolTipText     =   "This mode makes it so your tires automatically un-pop, your cars don't get destroyed, and you always have armor."
      Top             =   5400
      Width           =   3375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Special Modes:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   5880
      TabIndex        =   15
      Top             =   5040
      Width           =   3375
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "D] Blow Up All Cars"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   2880
      TabIndex        =   14
      Top             =   5040
      Width           =   3375
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "1] Basic Weapons"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   4680
      Width           =   3375
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "2] Moderate Weapons"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   5040
      Width           =   3375
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "3] Best Weapons"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   5400
      Width           =   3375
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Help / About"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   4320
      Width           =   3375
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "X] No Wanted Level"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   2880
      TabIndex        =   9
      Top             =   5760
      Width           =   3375
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Z] Full Health and Armor"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   2880
      TabIndex        =   8
      Top             =   5400
      Width           =   3375
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "N] Clear Weather"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   2880
      TabIndex        =   7
      Top             =   6120
      Width           =   3375
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "S] Tank ""Rhino"""
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   2880
      TabIndex        =   6
      Top             =   6480
      Width           =   3375
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "B] Golf Cart ""Caddy"""
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   6480
      Width           =   3375
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "A] Instant Dealership"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   6120
      Width           =   3375
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "www.MMTbb.com"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5760
      TabIndex        =   3
      Top             =   360
      Width           =   3495
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "M] Hotring Racer"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   5760
      Width           =   3375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Vice City PC Trainer"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5760
      TabIndex        =   1
      Top             =   0
      Width           =   3495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim result As Integer
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Private Sub Form_Unload(Cancel As Integer)
End

End Sub

Private Sub Label11_Click()
Form2.Show

End Sub

Private Sub Label7_Click()
x$ = Left$(UCase$(InputBox("Hotkey (One letter). Make sure it's not already in use, or the cheats/functions will double-up.", "Choose the cheats HotKey")), 1)
y$ = InputBox("Cheat Code:", "Enter the Cheat to Activate")

Label7.Caption = x$ + "] " + y$
lblhot.Caption = x$
lblcheat.Caption = y$

End Sub

Private Sub Text1_Change()
Text1.Text = Trim(Text1.Text)
'On Error Resume Next
If Text1.Text = "M" Then
    AppActivate "GTA: Vice City"
    SendKeys "getthereveryfastindeed"
    Text1.Text = ""
End If


If Text1.Text = "N" Then
    AppActivate "GTA: Vice City"
    SendKeys "alovelyday"
    Text1.Text = ""
End If

If Text1.Text = "B" Then
    AppActivate "GTA: Vice City"
    SendKeys "betterthanwalking"
    Text1.Text = ""
End If

If Text1.Text = "S" Then
    AppActivate "GTA: Vice City"
    SendKeys "panzer"
    Text1.Text = ""
End If

If Text1.Text = "Z" Then
    AppActivate "GTA: Vice City"
    SendKeys "aspirine"
    SendKeys "PRECIOUSPROTECTION"
    Text1.Text = ""
End If

If Text1.Text = "X" Then
    AppActivate "GTA: Vice City"
    SendKeys "leavemealone"
    Text1.Text = ""
End If

If Text1.Text = "1" Then
    AppActivate "GTA: Vice City"
    SendKeys "thugstools"
    Text1.Text = ""
End If

If Text1.Text = "2" Then
    AppActivate "GTA: Vice City"
    SendKeys "professionaltools"
    Text1.Text = ""
End If

If Text1.Text = "3" Then
    AppActivate "GTA: Vice City"
    SendKeys "nuttertools"
    Text1.Text = ""
End If

If Text1.Text = "D" Then
    AppActivate "GTA: Vice City"
    SendKeys "bigbang"
    Text1.Text = ""
End If

If Text1.Text = "A" Then
    AppActivate "GTA: Vice City"
    SendKeys "getthereveryfastindeed"
    SendKeys "getthereamazinglyfast"
    SendKeys "gettherefast"
    SendKeys "travelinstyle"
    SendKeys "gettherequickly"
    Text1.Text = ""
End If

If Text1.Text = "P" Then
    lblP.BackColor = &HFFFF&
    lblK.BackColor = &HFFFFFF
    lblL.BackColor = &HFFFFFF
End If

If Text1.Text = "K" Then
    lblP.BackColor = &HFFFFFF
    lblK.BackColor = &HFFFF&
    lblL.BackColor = &HFFFFFF
End If

If Text1.Text = "L" Then
    lblP.BackColor = &HFFFFFF
    lblK.BackColor = &HFFFFFF
    lblL.BackColor = &HFFFF&
End If






End Sub

Private Sub Timer1_Timer()
For i = 1 To 255
result = 0
result = GetAsyncKeyState(i)

If result = -32767 Then
Text1.Text = Chr(i)
'txtRecentASCII.Text = txtRecentASCII.Text + Str(i)
End If
Next i


End Sub
