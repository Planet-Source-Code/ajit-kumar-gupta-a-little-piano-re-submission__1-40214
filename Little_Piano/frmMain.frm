VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Piano.."
   ClientHeight    =   2790
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3390
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   2790
   ScaleWidth      =   3390
   StartUpPosition =   3  'Windows Default
   Begin VB.HScrollBar HScrollFrequency 
      Height          =   210
      LargeChange     =   100
      Left            =   120
      Max             =   4100
      Min             =   2600
      SmallChange     =   100
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   2100
      Value           =   2600
      Width           =   3135
   End
   Begin VB.HScrollBar HScrollTime 
      Height          =   210
      Left            =   120
      Max             =   50
      Min             =   20
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1440
      Value           =   20
      Width           =   3135
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000014&
      BorderStyle     =   0  'None
      Height          =   1155
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   3135
      Begin VB.Line Line1 
         BorderColor     =   &H80000001&
         BorderWidth     =   3
         Index           =   0
         X1              =   60
         X2              =   60
         Y1              =   180
         Y2              =   1080
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000001&
         BorderWidth     =   3
         Index           =   1
         X1              =   180
         X2              =   180
         Y1              =   180
         Y2              =   1080
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000001&
         BorderWidth     =   3
         Index           =   2
         X1              =   300
         X2              =   300
         Y1              =   180
         Y2              =   1080
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000001&
         BorderWidth     =   3
         Index           =   3
         X1              =   420
         X2              =   420
         Y1              =   180
         Y2              =   1080
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000001&
         BorderWidth     =   3
         Index           =   4
         X1              =   540
         X2              =   540
         Y1              =   180
         Y2              =   1080
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000001&
         BorderWidth     =   3
         Index           =   5
         X1              =   660
         X2              =   660
         Y1              =   180
         Y2              =   1080
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000001&
         BorderWidth     =   3
         Index           =   6
         X1              =   780
         X2              =   780
         Y1              =   180
         Y2              =   1080
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000001&
         BorderWidth     =   3
         Index           =   7
         X1              =   900
         X2              =   900
         Y1              =   180
         Y2              =   1080
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000001&
         BorderWidth     =   3
         Index           =   8
         X1              =   1020
         X2              =   1020
         Y1              =   180
         Y2              =   1080
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000001&
         BorderWidth     =   3
         Index           =   9
         X1              =   1140
         X2              =   1140
         Y1              =   180
         Y2              =   1080
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000001&
         BorderWidth     =   3
         Index           =   10
         X1              =   1260
         X2              =   1260
         Y1              =   180
         Y2              =   1080
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000001&
         BorderWidth     =   3
         Index           =   11
         X1              =   1380
         X2              =   1380
         Y1              =   180
         Y2              =   1080
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000001&
         BorderWidth     =   3
         Index           =   12
         X1              =   1500
         X2              =   1500
         Y1              =   180
         Y2              =   1080
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000001&
         BorderWidth     =   3
         Index           =   13
         X1              =   1620
         X2              =   1620
         Y1              =   180
         Y2              =   1080
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000001&
         BorderWidth     =   3
         Index           =   14
         X1              =   1740
         X2              =   1740
         Y1              =   180
         Y2              =   1080
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000001&
         BorderWidth     =   3
         Index           =   15
         X1              =   1860
         X2              =   1860
         Y1              =   180
         Y2              =   1080
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000001&
         BorderWidth     =   3
         Index           =   16
         X1              =   1980
         X2              =   1980
         Y1              =   180
         Y2              =   1080
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000001&
         BorderWidth     =   3
         Index           =   17
         X1              =   2100
         X2              =   2100
         Y1              =   180
         Y2              =   1080
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000001&
         BorderWidth     =   3
         Index           =   18
         X1              =   2220
         X2              =   2220
         Y1              =   180
         Y2              =   1080
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000001&
         BorderWidth     =   3
         Index           =   19
         X1              =   2340
         X2              =   2340
         Y1              =   180
         Y2              =   1080
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000001&
         BorderWidth     =   3
         Index           =   20
         X1              =   2460
         X2              =   2460
         Y1              =   180
         Y2              =   1080
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000001&
         BorderWidth     =   3
         Index           =   21
         X1              =   2580
         X2              =   2580
         Y1              =   180
         Y2              =   1080
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000001&
         BorderWidth     =   3
         Index           =   22
         X1              =   2700
         X2              =   2700
         Y1              =   180
         Y2              =   1080
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000001&
         BorderWidth     =   3
         Index           =   23
         X1              =   2820
         X2              =   2820
         Y1              =   180
         Y2              =   1080
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000001&
         BorderWidth     =   3
         Index           =   24
         X1              =   2940
         X2              =   2940
         Y1              =   180
         Y2              =   1080
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000001&
         BorderWidth     =   3
         Index           =   25
         X1              =   3060
         X2              =   3060
         Y1              =   180
         Y2              =   1080
      End
   End
   Begin VB.TextBox txtEdit 
      Height          =   285
      Left            =   240
      TabIndex        =   0
      Top             =   60
      Width           =   2955
   End
   Begin VB.Label LblTitleFrequency 
      AutoSize        =   -1  'True
      Caption         =   "Frequency:"
      Height          =   195
      Left            =   1800
      TabIndex        =   6
      Top             =   2400
      Width           =   795
   End
   Begin VB.Label LblFrequency 
      AutoSize        =   -1  'True
      Caption         =   "2600"
      Height          =   195
      Left            =   2700
      TabIndex        =   5
      Top             =   2400
      Width           =   360
   End
   Begin VB.Label LblTitleTime 
      AutoSize        =   -1  'True
      Caption         =   "Time in Milliseconds:"
      Height          =   195
      Left            =   1320
      TabIndex        =   4
      Top             =   1740
      Width           =   1455
   End
   Begin VB.Label LblTime 
      AutoSize        =   -1  'True
      Caption         =   "20"
      Height          =   195
      Left            =   2880
      TabIndex        =   3
      Top             =   1740
      Width           =   180
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'There is a textbox behind The Frame which always has the focus
'whenever the key is pressed.
'Make sure to make Tabstop property to False for the two Scrollbar
'The Piano will function only during the keypress provided
'there is focus to this Form.
'The keypress (a to z) and spaceBar is only accepted to make sound

'Declare API BeepAPI Function for Beep
Private Declare Function BeepAPI Lib "kernel32" Alias "Beep" (ByVal dwFrequency As Long, ByVal dwMilliseconds As Long) As Long

Dim intKey As Integer 'as global

Private Sub HScrollFrequency_Change()
LblFrequency.Caption = HScrollFrequency.Value
txtEdit.SetFocus
End Sub

Private Sub HScrollTime_Change()
LblTime.Caption = HScrollTime.Value
txtEdit.SetFocus
End Sub

Private Sub txtEdit_KeyPress(KeyAscii As Integer)
Dim i As Integer, j As Integer

'select the keypress between a to z
If KeyAscii >= 97 And KeyAscii <= 122 Then
    intKey = KeyAscii - 97
    Me.Caption = Chr$(KeyAscii)
    Line1(intKey).BorderColor = &HFF&
End If

Select Case KeyAscii
    Case 97
        BeepAPI (LblFrequency.Caption - 0), LblTime.Caption
    Case 98
        BeepAPI (LblFrequency.Caption - 100), LblTime.Caption
    Case 99
        BeepAPI (LblFrequency.Caption - 200), LblTime.Caption
    Case 100
        BeepAPI (LblFrequency.Caption - 300), LblTime.Caption
    Case 101
        BeepAPI (LblFrequency.Caption - 400), LblTime.Caption
    Case 102
        BeepAPI (LblFrequency.Caption - 500), LblTime.Caption
    Case 103
        BeepAPI (LblFrequency.Caption - 600), LblTime.Caption
    Case 104
        BeepAPI (LblFrequency.Caption - 700), LblTime.Caption
    Case 105
        BeepAPI (LblFrequency.Caption - 800), LblTime.Caption
    Case 106
        BeepAPI (LblFrequency.Caption - 900), LblTime.Caption
    Case 107
        BeepAPI (LblFrequency.Caption - 1000), LblTime.Caption
    Case 108
        BeepAPI (LblFrequency.Caption - 1100), LblTime.Caption
    Case 109
        BeepAPI (LblFrequency.Caption - 1200), LblTime.Caption
    Case 110
        BeepAPI (LblFrequency.Caption - 1300), LblTime.Caption
    Case 111
        BeepAPI (LblFrequency.Caption - 1400), LblTime.Caption
    Case 112
        BeepAPI (LblFrequency.Caption - 1500), LblTime.Caption
    Case 113
        BeepAPI (LblFrequency.Caption - 1600), LblTime.Caption
    Case 114
        BeepAPI (LblFrequency.Caption - 1700), LblTime.Caption
    Case 115
        BeepAPI (LblFrequency.Caption - 1800), LblTime.Caption
    Case 116
        BeepAPI (LblFrequency.Caption - 1900), LblTime.Caption
    Case 117
        BeepAPI (LblFrequency.Caption - 2000), LblTime.Caption
    Case 118
        BeepAPI (LblFrequency.Caption - 2100), LblTime.Caption
    Case 119
        BeepAPI (LblFrequency.Caption - 2200), LblTime.Caption
    Case 120
        BeepAPI (LblFrequency.Caption - 2300), LblTime.Caption
    Case 121
        BeepAPI (LblFrequency.Caption - 2400), LblTime.Caption
    Case 122
        BeepAPI (LblFrequency.Caption - 2500), LblTime.Caption
    Case 32 'select space bar keypress to get sound in one shot from (a to z)
        j = 97
        For i = LblFrequency.Caption To (LblFrequency.Caption - 2500) Step -100
            Me.Caption = Chr$(j)
            BeepAPI i, LblTime.Caption
            j = j + 1
        Next
    Case Else
        KeyAscii = 0
End Select
End Sub

Private Sub txtEdit_KeyUp(KeyCode As Integer, Shift As Integer)
Me.Caption = ""
txtEdit.Text = ""
Line1(intKey).BorderColor = &H80000001
End Sub

Private Sub Form_Load()
Dim i As Integer

'Avoid multiple instance
If App.PrevInstance Then
    End
End If

LblTime.Caption = HScrollTime.Value
LblFrequency.Caption = HScrollFrequency.Value
End Sub

