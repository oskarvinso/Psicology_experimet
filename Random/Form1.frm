VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   9960
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15090
   LinkTopic       =   "Form1"
   ScaleHeight     =   9960
   ScaleWidth      =   15090
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   735
      Left            =   1320
      TabIndex        =   0
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   375
      Index           =   3
      Left            =   7440
      TabIndex        =   4
      Top             =   3720
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   375
      Index           =   2
      Left            =   7560
      TabIndex        =   3
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   375
      Index           =   1
      Left            =   4080
      TabIndex        =   2
      Top             =   3720
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   375
      Index           =   0
      Left            =   4080
      TabIndex        =   1
      Top             =   1440
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RndEntre4y1 As Integer

Private Sub Command1_Click()
Randy
Organizator
End Sub

Sub Randy()
Dim Listo As Boolean
RndEntre4y1 = 0
While Listo = False
    If RndEntre4y1 > 4 Or RndEntre4y1 < 1 Then
        RndEntre4y1 = Int(Right(Format(Time, "hh:mm:ss"), 1))
        RndEntre4y1 = Int(Right(RndEntre4y1 - ((4 * Rnd) + 1), 1))
    Else
        Listo = True
    End If
Wend
Print RndEntre4y1
End Sub


Sub Organizator()
Select Case RndEntre4y1
    Case 1
        Label1(0).Caption = "Opcion A1"
        Label1(1).Caption = "Opcion A2"
        Label1(2).Caption = "Opcion B1"
        Label1(3).Caption = "Opcion B2"
    Case 2
        Label1(0).Caption = "Opcion A2"
        Label1(1).Caption = "Opcion A1"
        Label1(2).Caption = "Opcion B2"
        Label1(3).Caption = "Opcion B1"
    Case 3
        Label1(0).Caption = "Opcion B1"
        Label1(1).Caption = "Opcion B2"
        Label1(2).Caption = "Opcion A1"
        Label1(3).Caption = "Opcion A2"
    Case 4
        Label1(0).Caption = "Opcion B2"
        Label1(1).Caption = "Opcion B1"
        Label1(2).Caption = "Opcion A2"
        Label1(3).Caption = "Opcion A1"
End Select
End Sub
