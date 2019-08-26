VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "bul"
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   840
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1800
      TabIndex        =   0
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "faktoriyelini bulmasýný istediðiniz sayýyý kutuya yazý butona týklayýn"
      Height          =   1815
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   1680
      TabIndex        =   2
      Top             =   1440
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim sayac As Integer
Dim toplam As Integer

toplam = 1
sayac = 1



While sayac < Text1
sayac = sayac + 1
toplam = toplam * sayac

Wend
Label1.Caption = Val(toplam)

End Sub

Private Sub Label1_Click()

End Sub
