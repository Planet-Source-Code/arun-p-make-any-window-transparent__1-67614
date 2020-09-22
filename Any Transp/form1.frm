VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Make all Transparent"
   ClientHeight    =   1305
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5610
   LinkTopic       =   "Form1"
   ScaleHeight     =   1305
   ScaleWidth      =   5610
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.Slider Slider1 
      Height          =   675
      Left            =   1440
      TabIndex        =   3
      ToolTipText     =   "Transparency Alpha"
      Top             =   360
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   1191
      _Version        =   393216
      Max             =   255
      SelStart        =   150
      TickStyle       =   2
      TickFrequency   =   10
      Value           =   150
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Set Opaque"
      Height          =   375
      Left            =   4320
      TabIndex        =   2
      Top             =   480
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Set Transparent"
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   480
      Width           =   1335
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   5040
      Top             =   0
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "How to use? | Home page | About"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   0
      MousePointer    =   15  'Size All
      TabIndex        =   4
      Top             =   1080
      Width           =   5655
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Const GWL_EXSTYLE = -20
    Const WS_EX_LAYERED = &H80000
    Const GWL_STYLE = (-16)
    Const WS_VISIBLE = &H10000000
    Const LWA_ALPHA = &H2
    
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Option Explicit

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long


'@ http://digitalpbk.blogspot.com/2007/01/making-any-window-transparent-on-xp.html


Private Type POINTAPI
        x As Long
        y As Long
End Type

Public xhwnd As Long

Private Sub Command1_Click()
    SetWindowLong xhwnd, GWL_EXSTYLE, WS_EX_LAYERED
    SetLayeredWindowAttributes xhwnd, 0, Slider1.Value, LWA_ALPHA
End Sub

Private Sub Command2_Click()
    SetWindowLong xhwnd, GWL_EXSTYLE, 0
End Sub


Private Sub Label2_Click()
    ShellExecute Me.hwnd, "", "http://digitalpbk.blogspot.com/2007/01/making-any-window-transparent-on-xp.html?ref=EXE", "", "", 1
End Sub

Private Sub Slider1_Change()
    SetLayeredWindowAttributes xhwnd, 0, Slider1.Value, LWA_ALPHA
End Sub

Private Sub Timer1_Timer()
    Dim xy As POINTAPI
    Dim retStr As String * 80

    If GetAsyncKeyState(vbKeyControl) Then
        Call GetCursorPos&(xy)
        xhwnd = WindowFromPoint(xy.x, xy.y)
        Call GetWindowText(xhwnd, retStr, 100)
        If xhwnd = Me.hwnd Then
            Label1.Caption = "Making me transparent is not good!"
            xhwnd = 0
        Else
            Label1.Caption = retStr & " " & xy.x & " " & xy.y & " " & xhwnd
        End If
    End If
End Sub
