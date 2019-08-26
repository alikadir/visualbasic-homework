VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Meltem Kayýkcý"
   ClientHeight    =   4455
   ClientLeft      =   -3015
   ClientTop       =   8970
   ClientWidth     =   5790
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   5790
   Begin VB.Frame Frame2 
      Caption         =   "Message"
      Height          =   1575
      Left            =   480
      TabIndex        =   4
      Top             =   2520
      Width           =   4935
      Begin VB.CommandButton Command3 
         Caption         =   "Mesaj ver"
         Height          =   495
         Left            =   1680
         TabIndex        =   5
         Top             =   600
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Saat"
      Height          =   1935
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   4935
      Begin VB.CommandButton Command1 
         Caption         =   "timer'ý Baþlat"
         Height          =   495
         Index           =   0
         Left            =   240
         TabIndex        =   2
         Top             =   240
         Width           =   1695
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   2160
         Top             =   240
      End
      Begin VB.CommandButton Command2 
         Caption         =   "timer'ý Durdur"
         Height          =   495
         Left            =   2880
         TabIndex        =   1
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "SaaT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   615
         Left            =   1080
         TabIndex        =   3
         Top             =   1080
         Width           =   2175
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click(Index As Integer)
'butona týklanýnca Timer çalýþmaya baþladý
Timer1.Enabled = True
End Sub

Private Sub Command2_Click()
'butona týklanýnca timer durdu
Timer1.Enabled = False
End Sub

Private Sub Command3_Click()
'mesaj kutusunyla kullanýcýya soru sorduk ve gelen cevaba göre baþka bi mesaj kutusuyla tepki verdik .
If MsgBox("Yukarýda Çalýþan Saat Doðru Mu ? ", vbExclamation + vbYesNo) = vbYes Then
MsgBox "Doðru Çalýþtýðýna Sevindim :)"
Else
MsgBox "Ýþletim Sisteminin Saat Ayarlarýný Kontrol edin!"
End If


End Sub

Private Sub Form_Load()
'Form u ekran kordinatlarýný alýp ortalýyoruz load olayýnda
Form1.Left = (Screen.Width - Form1.Width) / 2
Form1.Top = (Screen.Height - Form1.Height) / 2
End Sub

Private Sub Timer1_Timer()
'timer in içi, yani belirli zman aralýðýnda çalýþacak kodlar
Label1.Caption = Time
End Sub

