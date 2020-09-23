VERSION 5.00
Begin VB.Form frmClock 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   1485
   ClientLeft      =   11985
   ClientTop       =   765
   ClientWidth     =   1305
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1485
   ScaleWidth      =   1305
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   600
      Top             =   120
   End
   Begin VB.Label lblSecs 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   960
      TabIndex        =   2
      Top             =   1200
      Width           =   375
   End
   Begin VB.Label lblMins 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   1200
      Width           =   375
   End
   Begin VB.Label lblHrs 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   255
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H0000FF00&
      Height          =   135
      Index           =   37
      Left            =   120
      Shape           =   3  'Circle
      Top             =   960
      Width           =   135
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H0000FF00&
      Height          =   135
      Index           =   36
      Left            =   240
      Shape           =   3  'Circle
      Top             =   960
      Width           =   135
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H0000FF00&
      Height          =   135
      Index           =   35
      Left            =   360
      Shape           =   3  'Circle
      Top             =   960
      Width           =   135
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H0000FF00&
      Height          =   135
      Index           =   34
      Left            =   480
      Shape           =   3  'Circle
      Top             =   960
      Width           =   135
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H0000FF00&
      Height          =   135
      Index           =   33
      Left            =   600
      Shape           =   3  'Circle
      Top             =   960
      Width           =   135
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H0000FF00&
      Height          =   135
      Index           =   32
      Left            =   720
      Shape           =   3  'Circle
      Top             =   960
      Width           =   135
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H0000FF00&
      Height          =   135
      Index           =   31
      Left            =   840
      Shape           =   3  'Circle
      Top             =   960
      Width           =   135
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H0000FF00&
      Height          =   135
      Index           =   30
      Left            =   960
      Shape           =   3  'Circle
      Top             =   960
      Width           =   135
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H0000FF00&
      Height          =   135
      Index           =   29
      Left            =   1080
      Shape           =   3  'Circle
      Top             =   960
      Width           =   135
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H0000FF00&
      Height          =   135
      Index           =   28
      Left            =   120
      Shape           =   3  'Circle
      Top             =   840
      Width           =   135
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H0000FF00&
      Height          =   135
      Index           =   27
      Left            =   240
      Shape           =   3  'Circle
      Top             =   840
      Width           =   135
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H0000FF00&
      Height          =   135
      Index           =   26
      Left            =   360
      Shape           =   3  'Circle
      Top             =   840
      Width           =   135
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H0000FF00&
      Height          =   135
      Index           =   25
      Left            =   480
      Shape           =   3  'Circle
      Top             =   840
      Width           =   135
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H0000FF00&
      Height          =   135
      Index           =   24
      Left            =   600
      Shape           =   3  'Circle
      Top             =   840
      Width           =   135
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H0000FF00&
      Height          =   135
      Index           =   23
      Left            =   120
      Shape           =   3  'Circle
      Top             =   600
      Width           =   135
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H0000FF00&
      Height          =   135
      Index           =   22
      Left            =   240
      Shape           =   3  'Circle
      Top             =   600
      Width           =   135
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H0000FF00&
      Height          =   135
      Index           =   21
      Left            =   360
      Shape           =   3  'Circle
      Top             =   600
      Width           =   135
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H0000FF00&
      Height          =   135
      Index           =   20
      Left            =   480
      Shape           =   3  'Circle
      Top             =   600
      Width           =   135
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H0000FF00&
      Height          =   135
      Index           =   19
      Left            =   600
      Shape           =   3  'Circle
      Top             =   600
      Width           =   135
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H0000FF00&
      Height          =   135
      Index           =   18
      Left            =   720
      Shape           =   3  'Circle
      Top             =   600
      Width           =   135
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H0000FF00&
      Height          =   135
      Index           =   17
      Left            =   840
      Shape           =   3  'Circle
      Top             =   600
      Width           =   135
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H0000FF00&
      Height          =   135
      Index           =   16
      Left            =   960
      Shape           =   3  'Circle
      Top             =   600
      Width           =   135
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H0000FF00&
      Height          =   135
      Index           =   15
      Left            =   1080
      Shape           =   3  'Circle
      Top             =   600
      Width           =   135
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H0000FF00&
      Height          =   135
      Index           =   14
      Left            =   600
      Shape           =   3  'Circle
      Top             =   480
      Width           =   135
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H0000FF00&
      Height          =   135
      Index           =   13
      Left            =   480
      Shape           =   3  'Circle
      Top             =   480
      Width           =   135
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H0000FF00&
      Height          =   135
      Index           =   12
      Left            =   360
      Shape           =   3  'Circle
      Top             =   480
      Width           =   135
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H0000FF00&
      Height          =   135
      Index           =   11
      Left            =   240
      Shape           =   3  'Circle
      Top             =   480
      Width           =   135
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H0000FF00&
      Height          =   135
      Index           =   10
      Left            =   120
      Shape           =   3  'Circle
      Top             =   480
      Width           =   135
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H0000FF00&
      Height          =   135
      Index           =   9
      Left            =   120
      Shape           =   3  'Circle
      Top             =   120
      Width           =   135
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H0000FF00&
      Height          =   135
      Index           =   8
      Left            =   120
      Shape           =   3  'Circle
      Top             =   240
      Width           =   135
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H0000FF00&
      Height          =   135
      Index           =   7
      Left            =   240
      Shape           =   3  'Circle
      Top             =   240
      Width           =   135
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H0000FF00&
      Height          =   135
      Index           =   6
      Left            =   360
      Shape           =   3  'Circle
      Top             =   240
      Width           =   135
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H0000FF00&
      Height          =   135
      Index           =   5
      Left            =   480
      Shape           =   3  'Circle
      Top             =   240
      Width           =   135
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H0000FF00&
      Height          =   135
      Index           =   4
      Left            =   600
      Shape           =   3  'Circle
      Top             =   240
      Width           =   135
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H0000FF00&
      Height          =   135
      Index           =   3
      Left            =   720
      Shape           =   3  'Circle
      Top             =   240
      Width           =   135
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H0000FF00&
      Height          =   135
      Index           =   2
      Left            =   840
      Shape           =   3  'Circle
      Top             =   240
      Width           =   135
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H0000FF00&
      Height          =   135
      Index           =   1
      Left            =   960
      Shape           =   3  'Circle
      Top             =   240
      Width           =   135
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H0000FF00&
      Height          =   135
      Index           =   0
      Left            =   1080
      Shape           =   3  'Circle
      Top             =   240
      Width           =   135
   End
End
Attribute VB_Name = "frmClock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_DblClick()
    End
End Sub


Private Sub Timer1_Timer()
Dim x
Dim hr
Dim min1
Dim min2
Dim sec1
Dim sec2


x = Time

If Len(x) > 10 Then
    Shape1(9).FillStyle = 0
    lblHrs.Caption = Mid(x, 1, 2)
    lblMins.Caption = Mid(x, 4, 2)
    lblSecs.Caption = Mid(x, 7, 2)
    hr = Mid(x, 2, 1)
    min1 = Mid(x, 4, 1)
    min2 = Mid(x, 5, 1)
    sec1 = Mid(x, 7, 1)
    sec2 = Mid(x, 8, 1)
Else
    Shape1(9).FillStyle = 1
    lblHrs.Caption = Mid(x, 1, 1)
    lblMins.Caption = Mid(x, 3, 2)
    lblSecs.Caption = Mid(x, 6, 2)
    hr = Mid(x, 1, 1)
    min1 = Mid(x, 3, 1)
    min2 = Mid(x, 4, 1)
    sec1 = Mid(x, 6, 1)
    sec2 = Mid(x, 7, 1)
End If

If hr = 1 Then
    Shape1(8).FillStyle = 0
Else
    Shape1(8).FillStyle = 1
End If

If hr = 2 Then
    Shape1(7).FillStyle = 0
Else
    Shape1(7).FillStyle = 1
End If

If hr = 3 Then
    Shape1(6).FillStyle = 0
Else
    Shape1(6).FillStyle = 1
End If

If hr = 4 Then
    Shape1(5).FillStyle = 0
Else
    Shape1(5).FillStyle = 1
End If

If hr = 5 Then
    Shape1(4).FillStyle = 0
Else
    Shape1(4).FillStyle = 1
End If

If hr = 6 Then
    Shape1(3).FillStyle = 0
Else
    Shape1(3).FillStyle = 1
End If

If hr = 7 Then
    Shape1(2).FillStyle = 0
Else
    Shape1(2).FillStyle = 1
End If

If hr = 8 Then
    Shape1(1).FillStyle = 0
Else
    Shape1(1).FillStyle = 1
End If

If hr = 9 Then
    Shape1(0).FillStyle = 0
Else
    Shape1(0).FillStyle = 1
End If

If min1 = 0 Then
    Shape1(10).FillStyle = 1
    Shape1(11).FillStyle = 1
    Shape1(12).FillStyle = 1
    Shape1(13).FillStyle = 1
    Shape1(14).FillStyle = 1
End If

If min1 = 1 Then
    Shape1(10).FillStyle = 0
Else
    Shape1(10).FillStyle = 1
End If

If min1 = 2 Then
    Shape1(11).FillStyle = 0
Else
    Shape1(11).FillStyle = 1
End If

If min1 = 3 Then
    Shape1(12).FillStyle = 0
Else
    Shape1(12).FillStyle = 1
End If

If min1 = 4 Then
    Shape1(13).FillStyle = 0
Else
    Shape1(13).FillStyle = 1
End If

If min1 = 5 Then
    Shape1(14).FillStyle = 0
Else
    Shape1(14).FillStyle = 1
End If

If min2 = 0 Then
    Shape1(15).FillStyle = 1
    Shape1(16).FillStyle = 1
    Shape1(17).FillStyle = 1
    Shape1(18).FillStyle = 1
    Shape1(19).FillStyle = 1
    Shape1(20).FillStyle = 1
    Shape1(21).FillStyle = 1
    Shape1(22).FillStyle = 1
    Shape1(23).FillStyle = 1
End If

If min2 = 1 Then
    Shape1(23).FillStyle = 0
Else
    Shape1(23).FillStyle = 1
End If

If min2 = 2 Then
    Shape1(22).FillStyle = 0
Else
    Shape1(22).FillStyle = 1
End If

If min2 = 3 Then
    Shape1(21).FillStyle = 0
Else
    Shape1(21).FillStyle = 1
End If

If min2 = 4 Then
    Shape1(20).FillStyle = 0
Else
    Shape1(20).FillStyle = 1
End If

If min2 = 5 Then
    Shape1(19).FillStyle = 0
Else
    Shape1(19).FillStyle = 1
End If

If min2 = 6 Then
    Shape1(18).FillStyle = 0
Else
    Shape1(18).FillStyle = 1
End If

If min2 = 7 Then
    Shape1(17).FillStyle = 0
Else
    Shape1(17).FillStyle = 1
End If

If min2 = 8 Then
    Shape1(16).FillStyle = 0
Else
    Shape1(16).FillStyle = 1
End If

If min2 = 9 Then
    Shape1(15).FillStyle = 0
Else
    Shape1(15).FillStyle = 1
End If

If sec1 = 0 Then
    Shape1(24).FillStyle = 1
    Shape1(25).FillStyle = 1
    Shape1(26).FillStyle = 1
    Shape1(27).FillStyle = 1
    Shape1(28).FillStyle = 1
End If

If sec1 = 1 Then
    Shape1(28).FillStyle = 0
Else
    Shape1(28).FillStyle = 1
End If

If sec1 = 2 Then
    Shape1(27).FillStyle = 0
Else
    Shape1(27).FillStyle = 1
End If

If sec1 = 3 Then
    Shape1(26).FillStyle = 0
Else
    Shape1(26).FillStyle = 1
End If

If sec1 = 4 Then
    Shape1(25).FillStyle = 0
Else
    Shape1(25).FillStyle = 1
End If

If sec1 = 5 Then
    Shape1(24).FillStyle = 0
Else
    Shape1(24).FillStyle = 1
End If

If sec2 = 0 Then
    Shape1(29).FillStyle = 1
    Shape1(30).FillStyle = 1
    Shape1(31).FillStyle = 1
    Shape1(32).FillStyle = 1
    Shape1(33).FillStyle = 1
    Shape1(34).FillStyle = 1
    Shape1(35).FillStyle = 1
    Shape1(36).FillStyle = 1
    Shape1(37).FillStyle = 1
End If

If sec2 = 1 Then
    Shape1(37).FillStyle = 0
Else
    Shape1(37).FillStyle = 1
End If

If sec2 = 2 Then
    Shape1(36).FillStyle = 0
Else
    Shape1(36).FillStyle = 1
End If

If sec2 = 3 Then
    Shape1(35).FillStyle = 0
Else
    Shape1(35).FillStyle = 1
End If

If sec2 = 4 Then
    Shape1(34).FillStyle = 0
Else
    Shape1(34).FillStyle = 1
End If

If sec2 = 5 Then
    Shape1(33).FillStyle = 0
Else
    Shape1(33).FillStyle = 1
End If

If sec2 = 6 Then
    Shape1(32).FillStyle = 0
Else
    Shape1(32).FillStyle = 1
End If

If sec2 = 7 Then
    Shape1(31).FillStyle = 0
Else
    Shape1(31).FillStyle = 1
End If

If sec2 = 8 Then
    Shape1(30).FillStyle = 0
Else
    Shape1(30).FillStyle = 1
End If

If sec2 = 9 Then
    Shape1(29).FillStyle = 0
Else
    Shape1(29).FillStyle = 1
End If
End Sub
