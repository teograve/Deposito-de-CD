VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Tipo de Pelíicula"
   ClientHeight    =   6390
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5985
   BeginProperty Font 
      Name            =   "Magneto"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   6390
   ScaleWidth      =   5985
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command7 
      Caption         =   "Eliminar"
      Height          =   420
      Left            =   4320
      TabIndex        =   11
      Top             =   5280
      Width           =   1335
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Modificar"
      Height          =   495
      Left            =   2640
      TabIndex        =   10
      Top             =   5280
      Width           =   1455
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Guardar"
      Height          =   495
      Left            =   1200
      TabIndex        =   9
      Top             =   5280
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Crear"
      Height          =   465
      Left            =   120
      TabIndex        =   8
      Top             =   5280
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Inicio"
      Height          =   540
      Left            =   2520
      TabIndex        =   7
      Top             =   4320
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Adelante"
      Height          =   495
      Left            =   3240
      TabIndex        =   6
      Top             =   3480
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Anterior"
      Height          =   540
      Left            =   1440
      TabIndex        =   5
      Top             =   3480
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      DataField       =   "Cateogoría"
      DataSource      =   "Data1"
      Height          =   345
      Left            =   2280
      TabIndex        =   4
      Top             =   1800
      Width           =   3015
   End
   Begin VB.TextBox Text1 
      DataField       =   "Tipo"
      DataSource      =   "Data1"
      Height          =   345
      Left            =   2280
      TabIndex        =   3
      Top             =   1320
      Width           =   3015
   End
   Begin VB.Data Data1 
      Caption         =   "Tipo de Película"
      Connect         =   "Access"
      DatabaseName    =   "C:\Users\GODINEZ\Desktop\Base de datos Progra\Alquiler de discos.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   405
      Left            =   1440
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Tipo de Película"
      Top             =   2760
      Width           =   3495
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Categoría"
      Height          =   300
      Left            =   360
      TabIndex        =   2
      Top             =   1800
      Width           =   1245
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Tipo"
      Height          =   300
      Left            =   360
      TabIndex        =   1
      Top             =   1320
      Width           =   555
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo de Película"
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
      Left            =   1800
      TabIndex        =   0
      Top             =   480
      Width           =   2820
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Me.Hide
End Sub

Private Sub Command2_Click()
Me.Hide
Form2.Show
End Sub

Private Sub Command3_Click()
Me.Hide
Form7.Show
End Sub

Private Sub Command4_Click()
Data1.Recordset.AddNew
End Sub

Private Sub Command5_Click()
Data1.Recordset.Update
End Sub

Private Sub Command6_Click()
Data1.Recordset.Edit
End Sub

Private Sub Command7_Click()
Data1.Recordset.Delete
End Sub

Private Sub Data1_Validate(Action As Integer, Save As Integer)
If Data1.Recordset.BOF = True Then
Data1.Recordset.MoveFirst
End If
If Data1.Recordset.EOF = True Then
Data1.Recordset.MoveLast
End If
End Sub
