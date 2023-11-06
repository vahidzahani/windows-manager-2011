VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Window manager 1.0"
   ClientHeight    =   5250
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3720
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5250
   ScaleMode       =   0  'User
   ScaleWidth      =   3720
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command9 
      Caption         =   "Set default"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   16
      Top             =   550
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   15
      Text            =   "Window text"
      Top             =   600
      Width           =   2175
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Always On top"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   13
      Top             =   3960
      Width           =   1575
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Show All"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   3480
      Width           =   1575
   End
   Begin VB.ListBox List1 
      Height          =   3375
      Left            =   1800
      TabIndex        =   9
      Top             =   960
      Width           =   1815
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Show normal"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   3120
      Width           =   1575
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Minimize"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   2760
      Width           =   1575
   End
   Begin VB.CommandButton Command5 
      Caption         =   "To system try"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   2400
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Hide"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   2040
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Restore"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   2160
      Top             =   120
   End
   Begin VB.CommandButton Command2 
      Caption         =   "close"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Maximize"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   1575
   End
   Begin ComctlLib.Slider Slider1 
      Height          =   375
      Left            =   0
      TabIndex        =   11
      Top             =   4440
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   661
      _Version        =   327682
      Min             =   10
      Max             =   255
      SelStart        =   255
      Value           =   255
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "2008 - 2010 delta-soft.ir"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   720
      TabIndex        =   17
      Top             =   5040
      Width           =   2055
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "vahid.zahani@gmail.com"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   720
      TabIndex        =   14
      Top             =   4800
      Width           =   2085
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Opacity:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   12
      Top             =   4200
      Width           =   690
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Handle :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   675
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00000000"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00F48039&
      Height          =   195
      Left            =   840
      TabIndex        =   0
      Top             =   360
      Width           =   840
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function InitCommonControls Lib "comctl32.dll" () As Long
Private Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function SetActiveWindow Lib "user32.dll" (ByVal hwnd As Long) As Long

Dim Handle As Long
Dim Def_text As String


Private Function Getwintext(ByVal h As Long) As String
    Dim MyStr As String
    MyStr = String(100, Chr$(0))
    GetWindowText h, MyStr, 100
    MyStr = Left$(MyStr, InStr(MyStr, Chr$(0)) - 1)
    Getwintext = MyStr
End Function


Private Sub Check1_Click()
SetWindowPos Handle, Val(Check1.Value) - 2, 0, 0, 0, 0, &H10 Or &H40 Or &H2 Or &H1 'always on top
'//-1=true -2=flase
End Sub

Private Sub Command1_Click() 'maximize
ShowWindow Handle, 3
End Sub

Private Sub Command2_Click()
'SetActiveWindow Handle
'SendKeys "%{f4}"

End Sub

Private Sub Command3_Click()
ShowWindow Handle, 3
End Sub

Private Sub Command4_Click() 'hide
ShowWindow Handle, 0
List1.AddItem Handle
End Sub

Private Sub Command6_Click() 'minimize
ShowWindow Handle, 2
End Sub

Private Sub Command7_Click()
If List1.ListIndex <> -1 Then
    ShowWindow List1.List(List1.ListIndex), 1
    List1.RemoveItem (List1.ListIndex)
End If
End Sub

Private Sub Command8_Click()
For i = List1.ListCount - 1 To 0 Step -1
    ShowWindow List1.List(i), 1
    List1.RemoveItem (i)
Next
End Sub

Private Sub Command9_Click()
Text1.Text = Def_text
End Sub

Private Sub Form_Activate()
SetWindowPos Me.hwnd, -1, 0, 0, 0, 0, &H10 Or &H40 Or &H2 Or &H1 'always on top
End Sub

Private Sub Form_Initialize()
InitCommonControls
End Sub

Private Sub Slider1_Scroll()
    SetWindowLong Handle, -20, 524288
    SetLayeredWindowAttributes Handle, 0, Slider1.Value, 2
End Sub

Private Sub Text1_Change()
    SetWindowText Handle, Text1.Text
End Sub

Private Sub Timer1_Timer()
If GetForegroundWindow <> Me.hwnd Then
    Handle = GetForegroundWindow
    Label1 = Handle
    Def_text = Getwintext(Handle)
    Text1.Text = Def_text
End If
End Sub
