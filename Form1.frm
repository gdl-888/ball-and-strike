VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Form1 
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "스트라이크"
   ClientHeight    =   3825
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6105
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3825
   ScaleWidth      =   6105
   StartUpPosition =   3  'Windows 기본값
   Begin VB.CommandButton Command2 
      Caption         =   "다시"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1920
      TabIndex        =   27
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Frame Frame3 
      Caption         =   "메모 (스트라이크, 볼)"
      Height          =   3615
      Left            =   3720
      TabIndex        =   2
      Top             =   120
      Width           =   2295
      Begin VB.TextBox Text4 
         BackColor       =   &H00FFFFFF&
         Height          =   3255
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  '수직
         TabIndex        =   17
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "결과"
      Height          =   1335
      Left            =   120
      TabIndex        =   1
      Top             =   2400
      Width           =   3495
      Begin VB.Label a3 
         Caption         =   "-"
         Height          =   255
         Left            =   1920
         TabIndex        =   26
         Top             =   480
         Width           =   495
      End
      Begin VB.Label a2 
         Caption         =   "-"
         Height          =   255
         Left            =   1440
         TabIndex        =   25
         Top             =   480
         Width           =   375
      End
      Begin VB.Label a1 
         Caption         =   "-"
         Height          =   255
         Left            =   960
         TabIndex        =   24
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Label5 
         Caption         =   "정답은"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "시도"
         Height          =   255
         Left            =   1680
         TabIndex        =   19
         Top             =   840
         Width           =   615
      End
      Begin MSForms.Label cnt2 
         Height          =   255
         Left            =   1320
         TabIndex        =   18
         Top             =   840
         Width           =   375
         Caption         =   "0"
         Size            =   "661;450"
         FontName        =   "굴림"
         FontHeight      =   180
         FontCharSet     =   129
         FontPitchAndFamily=   34
      End
      Begin VB.Label answer 
         Caption         =   "실패"
         Height          =   255
         Left            =   3000
         TabIndex        =   16
         Top             =   840
         Width           =   375
      End
      Begin VB.Label Label6 
         Caption         =   "남음"
         Height          =   255
         Left            =   360
         TabIndex        =   15
         Top             =   840
         Width           =   495
      End
      Begin VB.Label cnt 
         Caption         =   "10"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   840
         Width           =   375
      End
      Begin VB.Label bls 
         Caption         =   "0"
         Height          =   375
         Left            =   3120
         TabIndex        =   13
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label3 
         Caption         =   "볼:"
         Height          =   255
         Left            =   2640
         TabIndex        =   12
         Top             =   240
         Width           =   375
      End
      Begin VB.Label str 
         Caption         =   "0"
         Height          =   255
         Left            =   1320
         TabIndex        =   11
         Top             =   240
         Width           =   135
      End
      Begin VB.Label Label1 
         Caption         =   "스트라이크:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "추측"
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3495
      Begin VB.CommandButton Command1 
         Caption         =   "추측"
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   1560
         Width           =   1575
      End
      Begin VB.TextBox y3 
         Alignment       =   1  '오른쪽 맞춤
         BackColor       =   &H00FFFFFF&
         Height          =   270
         Left            =   2520
         TabIndex        =   7
         Text            =   "0"
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox y2 
         Alignment       =   1  '오른쪽 맞춤
         BackColor       =   &H00FFFFFF&
         Height          =   270
         Left            =   1440
         TabIndex        =   5
         Text            =   "0"
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox y1 
         Alignment       =   1  '오른쪽 맞춤
         BackColor       =   &H00FFFFFF&
         Height          =   270
         Left            =   240
         TabIndex        =   3
         Text            =   "0"
         Top             =   360
         Width           =   420
      End
      Begin MSForms.SpinButton SpinButton3 
         Height          =   255
         Left            =   2880
         TabIndex        =   8
         Top             =   360
         Width           =   255
         Size            =   "450;450"
      End
      Begin MSForms.SpinButton SpinButton2 
         Height          =   255
         Left            =   1800
         TabIndex        =   6
         Top             =   360
         Width           =   255
         Size            =   "450;450"
      End
      Begin MSForms.SpinButton SpinButton1 
         Height          =   255
         Left            =   600
         TabIndex        =   4
         Top             =   360
         Width           =   255
         Size            =   "450;450"
      End
   End
   Begin VB.Label n3 
      Caption         =   "Label10"
      Height          =   375
      Left            =   2040
      TabIndex        =   22
      Top             =   4080
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label n2 
      Caption         =   "Label9"
      Height          =   375
      Left            =   1080
      TabIndex        =   21
      Top             =   4080
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label n1 
      Caption         =   "Label8"
      Height          =   375
      Left            =   120
      TabIndex        =   20
      Top             =   4080
      Visible         =   0   'False
      Width           =   375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Function RandomInteger(Lowerbound As Integer, Upperbound As Integer) As Integer
    RandomInteger = Int((Upperbound - Lowerbound + 1) * Rnd + Lowerbound)
End Function

Private Sub Command1_Click()
    cnt.Caption = cnt.Caption - 1
    cnt2.Caption = cnt2.Caption + 1
    str.Caption = "0"
    bls.Caption = "0"
    If cnt.Caption = "0" Then
        a1.Caption = n1.Caption
        a2.Caption = n2.Caption
        a3.Caption = n3.Caption
        Command1.Enabled = False
        Command2.Enabled = True
    End If
    If y1.Text = n1.Caption Then
        str.Caption = str.Caption + 1
    End If
    If y2.Text = n2.Caption Then
        str.Caption = str.Caption + 1
    End If
    If y3.Text = n3.Caption Then
        str.Caption = str.Caption + 1
    End If
    If y1.Text = n2.Caption Then
        bls.Caption = bls.Caption + 1
    End If
    If y1.Text = n3.Caption Then
        bls.Caption = bls.Caption + 1
    End If
    If y2.Text = n1.Caption Then
        bls.Caption = bls.Caption + 1
    End If
    If y2.Text = n3.Caption Then
        bls.Caption = bls.Caption + 1
    End If
    If y3.Text = n1.Caption Then
        bls.Caption = bls.Caption + 1
    End If
    If y3.Text = n2.Caption Then
        bls.Caption = bls.Caption + 1
    End If
    If y1.Text = n1.Caption And y2.Text = n2.Caption And y3.Text = n3.Caption Then
        str.Caption = "3"
        answer.Caption = "성공"
        a1.Caption = n1.Caption
        a2.Caption = n2.Caption
        a3.Caption = n3.Caption
        Command1.Enabled = False
        Command2.Enabled = True
    End If
End Sub

Private Sub Command2_Click()
    cnt.Caption = "10"
    cnt2.Caption = "0"
    answer.Caption = "실패"
    a1.Caption = "-"
    a2.Caption = "-"
    a3.Caption = "-"
    Randomize
    n1.Caption = Int((Rnd * 9) + 1)
    n2.Caption = Int((Rnd * 9) + 1)
    If n2.Caption = n1.Caption Then
        n2.Caption = Int((Rnd * 9) + 1)
    End If
    If n2.Caption = n1.Caption Then
        n2.Caption = Int((Rnd * 9) + 1)
    End If
    If n2.Caption = n1.Caption Then
        n2.Caption = Int((Rnd * 9) + 1)
    End If
    n3.Caption = Int((Rnd * 9) + 1)
    If n3.Caption = n2.Caption Then
        n3.Caption = Int((Rnd * 9) + 1)
    End If
    If n3.Caption = n2.Caption Then
        n3.Caption = Int((Rnd * 9) + 1)
    End If
    If n3.Caption = n2.Caption Then
        n3.Caption = Int((Rnd * 9) + 1)
    End If
    If n3.Caption = n1.Caption Then
        n3.Caption = Int((Rnd * 9) + 1)
    End If
    If n3.Caption = n1.Caption Then
        n3.Caption = Int((Rnd * 9) + 1)
    End If
    If n3.Caption = n1.Caption Then
        n3.Caption = Int((Rnd * 9) + 1)
    End If
    str.Caption = "0"
    bls.Caption = "0"
    Command2.Enabled = False
    Command1.Enabled = True
End Sub

Private Sub Form_Load()
    Randomize
    n1.Caption = Int((Rnd * 9) + 1)
    n2.Caption = Int((Rnd * 9) + 1)
    If n2.Caption = n1.Caption Then
        n2.Caption = Int((Rnd * 9) + 1)
    End If
    If n2.Caption = n1.Caption Then
        n2.Caption = Int((Rnd * 9) + 1)
    End If
    If n2.Caption = n1.Caption Then
        n2.Caption = Int((Rnd * 9) + 1)
    End If
    n3.Caption = Int((Rnd * 9) + 1)
    If n3.Caption = n2.Caption Then
        n3.Caption = Int((Rnd * 9) + 1)
    End If
    If n3.Caption = n2.Caption Then
        n3.Caption = Int((Rnd * 9) + 1)
    End If
    If n3.Caption = n2.Caption Then
        n3.Caption = Int((Rnd * 9) + 1)
    End If
    If n3.Caption = n1.Caption Then
        n3.Caption = Int((Rnd * 9) + 1)
    End If
    If n3.Caption = n1.Caption Then
        n3.Caption = Int((Rnd * 9) + 1)
    End If
    If n3.Caption = n1.Caption Then
        n3.Caption = Int((Rnd * 9) + 1)
    End If
End Sub

Private Sub SpinButton1_Change()
   y1.Text = SpinButton1.Value
End Sub

Private Sub SpinButton2_Change()
    y2.Text = SpinButton2.Value
    
End Sub

Private Sub SpinButton3_Change()
    y3.Text = SpinButton3.Value
    
End Sub
