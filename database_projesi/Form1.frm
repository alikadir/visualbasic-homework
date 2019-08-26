VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "meltem kayýkcý"
   ClientHeight    =   5910
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8865
   LinkTopic       =   "Form1"
   ScaleHeight     =   11595
   ScaleWidth      =   19080
   Begin VB.CommandButton Command9 
      Caption         =   "ÝPTAL"
      Height          =   615
      Left            =   5400
      TabIndex        =   22
      Top             =   2040
      Width           =   2055
   End
   Begin VB.CommandButton Command8 
      Caption         =   "KAYIT SÝL"
      Height          =   615
      Left            =   6720
      TabIndex        =   21
      Top             =   3840
      Width           =   1935
   End
   Begin VB.CommandButton Command7 
      Caption         =   "SON KAYIT"
      Height          =   615
      Left            =   6720
      TabIndex        =   20
      Top             =   2880
      Width           =   1935
   End
   Begin VB.CommandButton Command6 
      Caption         =   "ÝLK KAYIT"
      Height          =   735
      Left            =   6720
      TabIndex        =   19
      Top             =   1080
      Width           =   1935
   End
   Begin VB.CommandButton Command5 
      Caption         =   "ÇIKIÞ"
      Height          =   615
      Left            =   6720
      TabIndex        =   18
      Top             =   240
      Width           =   1935
   End
   Begin VB.CommandButton Command4 
      Caption         =   "GÜNCELLE"
      Height          =   615
      Left            =   4320
      TabIndex        =   17
      Top             =   3840
      Width           =   1935
   End
   Begin VB.CommandButton Command3 
      Caption         =   "YENÝ KAYIT"
      Height          =   615
      Left            =   4320
      TabIndex        =   16
      Top             =   2880
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ÖNCEKÝ"
      Height          =   615
      Left            =   4320
      TabIndex        =   15
      Top             =   1200
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "SONRAKÝ"
      Height          =   615
      Left            =   4320
      TabIndex        =   14
      Top             =   240
      Width           =   1935
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "access.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   4320
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "OGRTAKIP"
      Top             =   5040
      Width           =   4335
   End
   Begin VB.TextBox Text7 
      DataField       =   "MAHKOY"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   1320
      TabIndex        =   13
      Text            =   "Text7"
      Top             =   5160
      Width           =   2655
   End
   Begin VB.TextBox Text6 
      DataField       =   "ADRES"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   1320
      TabIndex        =   12
      Text            =   "Text6"
      Top             =   4320
      Width           =   2655
   End
   Begin VB.TextBox Text5 
      DataField       =   "ILCESI"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   1320
      TabIndex        =   11
      Text            =   "Text5"
      Top             =   3480
      Width           =   2655
   End
   Begin VB.TextBox Text4 
      DataField       =   "ILI"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   1320
      TabIndex        =   10
      Text            =   "Text4"
      Top             =   2640
      Width           =   2655
   End
   Begin VB.TextBox Text3 
      DataField       =   "DOGYER"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   1320
      TabIndex        =   9
      Text            =   "Text3"
      Top             =   1800
      Width           =   2655
   End
   Begin VB.TextBox Text2 
      DataField       =   "SOYAD"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   1320
      TabIndex        =   8
      Text            =   "Text2"
      Top             =   960
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      DataField       =   "AD"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   1320
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   240
      Width           =   2655
   End
   Begin VB.Label Label7 
      Caption         =   "MAHKOY"
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   5160
      Width           =   735
   End
   Begin VB.Label Label6 
      Caption         =   "ADRES"
      Height          =   255
      Left            =   480
      TabIndex        =   5
      Top             =   4320
      Width           =   615
   End
   Begin VB.Label Label5 
      Caption         =   "ILCESI"
      Height          =   255
      Left            =   480
      TabIndex        =   4
      Top             =   3480
      Width           =   495
   End
   Begin VB.Label Label4 
      Caption         =   "ILI"
      Height          =   255
      Left            =   720
      TabIndex        =   3
      Top             =   2640
      Width           =   255
   End
   Begin VB.Label Label3 
      Caption         =   "DOGYER"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   1800
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "SOYAD"
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   960
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "AD"
      Height          =   255
      Left            =   600
      TabIndex        =   0
      Top             =   240
      Width           =   375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Data1.Recordset.EOF Then
Data1.Recordset.MoveLast
MsgBox ("Son Kayýt")
Else
Data1.Recordset.MoveNext
End If
End Sub

Private Sub Command2_Click()
If Data1.Recordset.BOF Then
Data1.Recordset.MoveFirst
MsgBox ("Ýlk Kayýt")
Else
Data1.Recordset.MovePrevious
End If
End Sub

Private Sub Command3_Click()
Data1.Recordset.AddNew
Command1.Visible = False
Command2.Visible = False
Command3.Visible = False
Command4.Visible = True
Command5.Visible = False
Command6.Visible = False
Command7.Visible = False
Command8.Visible = False
Command9.Visible = True

End Sub

Private Sub Command4_Click()
Data1.Recordset.Update
Command1.Visible = True
Command2.Visible = True
Command3.Visible = True
Command4.Visible = False
Command5.Visible = True
Command6.Visible = True
Command7.Visible = True
Command8.Visible = True
Command9.Visible = False
End Sub

Private Sub Command5_Click()
Dim cevap
cevap = MsgBox("Çýkmak Ýstiyormusunuz", vbYesNo + vbQuestion, "Çýkýþ")
If cevap = vbYes Then
MsgBox ("Ýyi Günler")
End
End If
End Sub

Private Sub Command6_Click()
Data1.Recordset.MoveFirst
End Sub

Private Sub Command7_Click()
Data1.Recordset.MoveLast
End Sub

Private Sub Command8_Click()
Dim cevap
cevap = MsgBox("Kayýt Silinecek", vbYesNo + vbQuestion, "Kayýt Sil")
If cevap = vbYes Then
Data1.Recordset.Delete
Data1.Recordset.MoveNext
End If
End Sub

Private Sub Command9_Click()
Data1.Recordset.CancelUpdate
Command1.Visible = True
Command2.Visible = True
Command3.Visible = True
Command4.Visible = False
Command5.Visible = True
Command6.Visible = True
Command7.Visible = True
Command8.Visible = True
Command9.Visible = False
End Sub

