VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "Disco"
   ClientHeight    =   5925
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6015
   BeginProperty Font 
      Name            =   "Magneto"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form4"
   ScaleHeight     =   5925
   ScaleWidth      =   6015
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Crear"
      Height          =   465
      Left            =   240
      TabIndex        =   13
      Top             =   5040
      Width           =   975
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Guardar"
      Height          =   495
      Left            =   1320
      TabIndex        =   12
      Top             =   5040
      Width           =   1335
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Modificar"
      Height          =   495
      Left            =   2760
      TabIndex        =   11
      Top             =   5040
      Width           =   1455
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Eliminar"
      Height          =   420
      Left            =   4440
      TabIndex        =   10
      Top             =   5040
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Inicio"
      Height          =   375
      Left            =   2160
      TabIndex        =   9
      Top             =   4200
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      DataField       =   "Cod_Película"
      DataSource      =   "Data1"
      Height          =   420
      Left            =   3000
      TabIndex        =   8
      Top             =   1920
      Width           =   2655
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Siguiente"
      Height          =   495
      Left            =   3120
      TabIndex        =   6
      Top             =   3240
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Anterior"
      Height          =   540
      Left            =   1200
      TabIndex        =   5
      Top             =   3240
      Width           =   1575
   End
   Begin VB.Data Data1 
      Caption         =   "Disco"
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
      RecordSource    =   "Disco"
      Top             =   2640
      Width           =   3015
   End
   Begin VB.TextBox Text2 
      DataField       =   "Num_Copias"
      DataSource      =   "Data1"
      Height          =   420
      Left            =   3000
      TabIndex        =   4
      Top             =   1440
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      DataField       =   "Código"
      DataSource      =   "Data1"
      Height          =   420
      Left            =   3000
      TabIndex        =   3
      Top             =   960
      Width           =   2655
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Cod_Película"
      Height          =   300
      Left            =   360
      TabIndex        =   7
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Num_Copias"
      Height          =   300
      Left            =   360
      TabIndex        =   2
      Top             =   1560
      Width           =   1605
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Código"
      Height          =   300
      Left            =   360
      TabIndex        =   1
      Top             =   1080
      Width           =   840
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Disco"
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
      Left            =   1920
      TabIndex        =   0
      Top             =   360
      Width           =   900
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Me.Hide
Form3.Show
End Sub

Private Sub Command2_Click()
Me.Hide
Form5.Show
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
