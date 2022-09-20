VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  '單線固定
   Caption         =   "ScoreRecordAndTextTool1.0"
   ClientHeight    =   3465
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7875
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3465
   ScaleWidth      =   7875
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton Command4 
      Caption         =   "關於(&I)"
      Height          =   375
      Left            =   6600
      TabIndex        =   13
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Frame Frame3 
      Caption         =   "第三步驟：輸出"
      Height          =   2535
      Left            =   3120
      TabIndex        =   9
      Top             =   120
      Width           =   2895
      Begin VB.TextBox Text2 
         Height          =   1695
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   11
         ToolTipText     =   "將此處文字除了""exit""與首行空白外，其他複製至系統"
         Top             =   720
         Width           =   2655
      End
      Begin VB.CommandButton Command3 
         Caption         =   "3.輸出記錄的成績"
         Height          =   375
         Left            =   120
         TabIndex        =   10
         ToolTipText     =   "輸出記錄成績"
         Top             =   240
         Width           =   2655
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "第二步驟：偵測"
      Height          =   1695
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   2895
      Begin VB.CommandButton Command2 
         Caption         =   "查詢分數(&E)"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1560
         TabIndex        =   8
         ToolTipText     =   "開始查詢成績"
         Top             =   1200
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   270
         Left            =   1200
         TabIndex        =   4
         ToolTipText     =   "輸入被查詢之學生座號"
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label4 
         BorderStyle     =   1  '單線固定
         Enabled         =   0   'False
         Height          =   255
         Left            =   1200
         TabIndex        =   7
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "分數(&C)："
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "座號(&N)："
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "確認成績無誤(&C)"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "第一步驟：記錄"
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2895
      Begin VB.CommandButton Command1 
         Caption         =   "1.開始記錄成績(&S)"
         Height          =   375
         Left            =   120
         TabIndex        =   1
         ToolTipText     =   "開始記錄成績"
         Top             =   240
         Width           =   2655
      End
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000000&
      X1              =   120
      X2              =   7800
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Label Label5 
      Caption         =   $"Form1.frx":0000
      Height          =   2535
      Left            =   6120
      TabIndex        =   12
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim score(1 To 70)
Private Sub Command1_Click()
    Dim X As String
    Dim zoHao As Long
    zoHao = 1
    While Text <> "exit"
        Text = InputBox("學生的成績", "記錄", "exit")
        If zoHao > 69 Then
            MsgBox ("程式不允許超過69位學生之資料")
        Else
            score(zoHao) = Text
            zoHao = zoHao + 1
        End If
    Wend
    Text1.Enabled = True
    Label4.Enabled = True
    Command2.Enabled = True
    MsgBox "完成！"
End Sub

Private Sub Command2_Click()
    If Not (Text1.Text = "") Then
        Label4.Caption = score(Val(Text1.Text))
    Else
        MsgBox "你必須輸入座號！", 16
    End If
End Sub

Private Sub Command3_Click()
    For I = 1 To 70
        Text2.Text = Text2.Text + vbNewLine + score(I)
    Next I
End Sub

Private Sub Command4_Click()
    Form2.Show
End Sub
