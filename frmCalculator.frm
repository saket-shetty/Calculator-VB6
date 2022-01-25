VERSION 5.00
Begin VB.Form frmCalculator 
   Caption         =   "Form1"
   ClientHeight    =   8775
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7305
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   15.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   8775
   ScaleWidth      =   7305
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdResult 
      Caption         =   "="
      Height          =   975
      Left            =   5160
      TabIndex        =   17
      Top             =   7320
      Width           =   990
   End
   Begin VB.CommandButton cmdDivide 
      Caption         =   "/"
      Height          =   975
      Left            =   5160
      TabIndex        =   16
      Top             =   6000
      Width           =   990
   End
   Begin VB.CommandButton cmdMultiply 
      Caption         =   "X"
      Height          =   975
      Left            =   5160
      TabIndex        =   15
      Top             =   4800
      Width           =   990
   End
   Begin VB.CommandButton cmdSubtract 
      Caption         =   "-"
      Height          =   975
      Left            =   5160
      TabIndex        =   14
      Top             =   3600
      Width           =   990
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "+"
      Height          =   975
      Left            =   5160
      TabIndex        =   13
      Top             =   2400
      Width           =   990
   End
   Begin VB.CommandButton cmdDecimal 
      Caption         =   "."
      Height          =   1215
      Left            =   3720
      TabIndex        =   12
      Top             =   7080
      Width           =   990
   End
   Begin VB.CommandButton cmdZero 
      Caption         =   "0"
      Height          =   1215
      Left            =   2160
      TabIndex        =   11
      Top             =   7080
      Width           =   990
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "C"
      Height          =   1215
      Left            =   720
      TabIndex        =   10
      Top             =   7080
      Width           =   990
   End
   Begin VB.CommandButton cmdNine 
      Caption         =   "9"
      Height          =   1215
      Left            =   3720
      TabIndex        =   9
      Top             =   5520
      Width           =   990
   End
   Begin VB.CommandButton cmdEight 
      Caption         =   "8"
      Height          =   1215
      Left            =   2160
      TabIndex        =   8
      Top             =   5520
      Width           =   990
   End
   Begin VB.CommandButton cmdSeven 
      Caption         =   "7"
      Height          =   1215
      Left            =   720
      TabIndex        =   7
      Top             =   5520
      Width           =   990
   End
   Begin VB.CommandButton cmdSix 
      Caption         =   "6"
      Height          =   1215
      Left            =   3720
      TabIndex        =   6
      Top             =   3960
      Width           =   990
   End
   Begin VB.CommandButton cmdFive 
      Caption         =   "5"
      Height          =   1215
      Left            =   2160
      TabIndex        =   5
      Top             =   3960
      Width           =   990
   End
   Begin VB.CommandButton cmdFour 
      Caption         =   "4"
      Height          =   1215
      Left            =   720
      TabIndex        =   4
      Top             =   3960
      Width           =   990
   End
   Begin VB.CommandButton cmdThree 
      Caption         =   "3"
      Height          =   1215
      Left            =   3720
      TabIndex        =   3
      Top             =   2400
      Width           =   990
   End
   Begin VB.Frame FraCalculator 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8295
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   6855
      Begin VB.TextBox txtResult 
         Height          =   1095
         Left            =   480
         TabIndex        =   18
         Top             =   600
         Width           =   5415
      End
      Begin VB.CommandButton cmdTwo 
         Caption         =   "2"
         Height          =   1215
         Left            =   1920
         TabIndex        =   2
         Top             =   2160
         Width           =   990
      End
      Begin VB.CommandButton cmdOne 
         Caption         =   "1"
         Height          =   1200
         Left            =   480
         TabIndex        =   1
         Top             =   2160
         Width           =   990
      End
   End
End
Attribute VB_Name = "frmCalculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim op As String
Dim r As Integer

Private Sub cmdAdd_Click()
    op = "+"
    r = txtResult
    txtResult = ""
End Sub

Private Sub cmdClear_Click()
    txtResult = Empty
End Sub

Private Sub cmdDivide_Click()
    op = "/"
    r = txtResult
    txtResult = ""
End Sub

Private Sub cmdEight_Click()
    txtResult = txtResult & cmdEight.Caption
End Sub

Private Sub cmdFive_Click()
    txtResult = txtResult & cmdFive.Caption
End Sub

Private Sub cmdFour_Click()
    txtResult = txtResult & cmdFour.Caption
End Sub

Private Sub cmdMultiply_Click()
    op = "*"
    r = txtResult
    txtResult = ""
End Sub

Private Sub cmdNine_Click()
    txtResult = txtResult & cmdNine.Caption
End Sub

Private Sub cmdOne_Click()
    txtResult = txtResult & cmdOne.Caption
End Sub

Private Sub cmdResult_Click()
    Call CalculateResult
End Sub

Private Sub CalculateResult()
    If op = "+" Then
        txtResult = r + txtResult
    ElseIf op = "-" Then
        txtResult = r - txtResult
    ElseIf op = "*" Then
        txtResult = r * txtResult
    ElseIf op = "/" Then
        txtResult = r / txtResult
    End If
    op = Empty
End Sub

Private Sub cmdSeven_Click()
    txtResult = txtResult & cmdSeven.Caption
End Sub

Private Sub cmdSix_Click()
    txtResult = txtResult & cmdSix.Caption
End Sub

Private Sub cmdSubtract_Click()
    op = "-"
    r = txtResult
    txtResult = ""
End Sub

Private Sub cmdThree_Click()
    txtResult = txtResult & cmdThree.Caption
End Sub

Private Sub cmdTwo_Click()
    txtResult = txtResult & cmdTwo.Caption
End Sub

Private Sub cmdZero_Click()
    txtResult = txtResult & cmdZero.Caption
End Sub


