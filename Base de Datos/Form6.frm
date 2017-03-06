VERSION 5.00
Begin VB.Form Form6 
   Caption         =   "Cliente"
   ClientHeight    =   6030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6075
   BeginProperty Font 
      Name            =   "Magneto"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form6"
   ScaleHeight     =   9950.495
   ScaleMode       =   0  'User
   ScaleWidth      =   6075
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Crear"
      Height          =   465
      Index           =   1
      Left            =   360
      TabIndex        =   15
      Top             =   5280
      Width           =   975
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Guardar"
      Height          =   495
      Index           =   1
      Left            =   1440
      TabIndex        =   14
      Top             =   5280
      Width           =   1335
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Modificar"
      Height          =   495
      Index           =   1
      Left            =   2880
      TabIndex        =   13
      Top             =   5280
      Width           =   1455
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Eliminar"
      Height          =   420
      Index           =   1
      Left            =   4560
      TabIndex        =   12
      Top             =   5280
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Inicio"
      Height          =   420
      Left            =   2040
      TabIndex        =   11
      Top             =   4560
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Siguiente"
      Height          =   420
      Left            =   3240
      TabIndex        =   10
      Top             =   3840
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Anterior"
      Height          =   465
      Left            =   1200
      TabIndex        =   9
      Top             =   3840
      Width           =   1335
   End
   Begin VB.Data Data1 
      Caption         =   "Cliente"
      Connect         =   "Access"
      DatabaseName    =   "C:\Users\GODINEZ\Desktop\Base de datos Progra\Alquiler de discos.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   405
      Left            =   1680
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Cliente"
      Top             =   3240
      Width           =   2895
   End
   Begin VB.TextBox Text4 
      DataField       =   "Teléfono"
      DataSource      =   "Data1"
      Height          =   420
      Left            =   2880
      TabIndex        =   8
      Top             =   2520
      Width           =   2775
   End
   Begin VB.TextBox Text3 
      DataField       =   "Dirección"
      DataSource      =   "Data1"
      Height          =   420
      Left            =   2880
      TabIndex        =   7
      Top             =   2040
      Width           =   2775
   End
   Begin VB.TextBox Text2 
      DataField       =   "Nombre"
      DataSource      =   "Data1"
      Height          =   420
      Left            =   2880
      TabIndex        =   6
      Top             =   1560
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      DataField       =   "Num_Membresía"
      DataSource      =   "Data1"
      Height          =   420
      Left            =   2880
      TabIndex        =   5
      Top             =   1080
      Width           =   2775
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Teléfono"
      Height          =   300
      Left            =   360
      TabIndex        =   4
      Top             =   2520
      Width           =   1125
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Dirección"
      Height          =   300
      Left            =   360
      TabIndex        =   3
      Top             =   2040
      Width           =   1275
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Nombre"
      Height          =   300
      Left            =   360
      TabIndex        =   2
      Top             =   1560
      Width           =   1005
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Num_Mebresía"
      Height          =   300
      Left            =   360
      TabIndex        =   1
      Top             =   1080
      Width           =   1980
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cliente"
      BeginProperty Font 
         Name            =   "Segoe Script"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      TabIndex        =   0
      Top             =   360
      Width           =   1200
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Me.Hide
Form5.Show

End Sub

Private Sub Command2_Click()
Me.Hide
Form1.Show
End Sub

Private Sub Command3_Click()
Me.Hide
Form7.Show
End Sub

Private Sub Command4_Click(Index As Integer)
Data1.Recordset.AddNew
End Sub

Private Sub Command5_Click(Index As Integer)
Data1.Recordset.Update
End Sub

Private Sub Command6_Click(Index As Integer)
Data1.Recordset.Edit
End Sub

Private Sub Command7_Click(Index As Integer)
Data1.Recordset.Delete
End Sub

Private Sub Data1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

End Sub

Private Sub Data1_Validate(Action As Integer, Save As Integer)
If Data1.Recordset.BOF = True Then
Data1.Recordset.MoveFirst
End If
If Data1.Recordset.EOF = True Then
Data1.Recordset.MoveLast
End If
End Sub
