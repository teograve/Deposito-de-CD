VERSION 5.00
Begin VB.Form Form5 
   Caption         =   "Alquiler"
   ClientHeight    =   6720
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5685
   BeginProperty Font 
      Name            =   "Magneto"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form5"
   ScaleHeight     =   6720
   ScaleWidth      =   5685
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Inicio"
      Height          =   495
      Left            =   2040
      TabIndex        =   17
      Top             =   5880
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Siguiente"
      Height          =   540
      Left            =   2760
      TabIndex        =   16
      Top             =   5160
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Anterior"
      Height          =   540
      Left            =   720
      TabIndex        =   15
      Top             =   5160
      Width           =   1695
   End
   Begin VB.Data Data1 
      Caption         =   "Alquiler"
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
      RecordSource    =   "Alquiler"
      Top             =   4560
      Width           =   2775
   End
   Begin VB.TextBox Text7 
      DataField       =   "Cantidad"
      DataSource      =   "Data1"
      Height          =   420
      Left            =   3000
      TabIndex        =   14
      Top             =   3840
      Width           =   2295
   End
   Begin VB.TextBox Text6 
      DataField       =   "Valor_Alquiler"
      DataSource      =   "Data1"
      Height          =   420
      Left            =   3000
      TabIndex        =   13
      Top             =   3360
      Width           =   2295
   End
   Begin VB.TextBox Text5 
      DataField       =   "Fecha_Devolución"
      DataSource      =   "Data1"
      Height          =   420
      Left            =   3000
      TabIndex        =   12
      Top             =   2880
      Width           =   2295
   End
   Begin VB.TextBox Text4 
      DataField       =   "Fecha_Alquiler"
      DataSource      =   "Data1"
      Height          =   420
      Left            =   3000
      TabIndex        =   11
      Top             =   2400
      Width           =   2295
   End
   Begin VB.TextBox Text3 
      DataField       =   "Cod_Cliente"
      DataSource      =   "Data1"
      Height          =   420
      Left            =   3000
      TabIndex        =   10
      Top             =   1920
      Width           =   2295
   End
   Begin VB.TextBox Text2 
      DataField       =   "Cod_Disco"
      DataSource      =   "Data1"
      Height          =   420
      Left            =   3000
      TabIndex        =   9
      Top             =   1440
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      DataField       =   "Código"
      DataSource      =   "Data1"
      Height          =   420
      Left            =   3000
      TabIndex        =   8
      Top             =   960
      Width           =   2295
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Cantidad"
      Height          =   300
      Left            =   480
      TabIndex        =   7
      Top             =   3840
      Width           =   1140
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Valor_Alquiler"
      Height          =   300
      Left            =   360
      TabIndex        =   6
      Top             =   3360
      Width           =   1965
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Fecha Devolución"
      Height          =   300
      Left            =   360
      TabIndex        =   5
      Top             =   2880
      Width           =   2295
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Fecha Alquiler"
      Height          =   300
      Left            =   360
      TabIndex        =   4
      Top             =   2400
      Width           =   2025
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Cod_Cliente"
      Height          =   300
      Left            =   360
      TabIndex        =   3
      Top             =   1920
      Width           =   1560
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Cod_Disco"
      Height          =   300
      Left            =   360
      TabIndex        =   2
      Top             =   1440
      Width           =   1350
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Código"
      Height          =   300
      Left            =   360
      TabIndex        =   1
      Top             =   960
      Width           =   840
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Alquiler"
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
      Left            =   1560
      TabIndex        =   0
      Top             =   240
      Width           =   1470
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Me.Hide
Form4.Show
End Sub

Private Sub Command2_Click()
Me.Hide
Form6.Show
End Sub

Private Sub Command3_Click()
Me.Hide
Form7.Show
End Sub

Private Sub Data1_Validate(Action As Integer, Save As Integer)
If Data1.Recordset.BOF = True Then
Data1.Recordset.MoveFirst
End If
If Data1.Recordset.EOF = True Then
Data1.Recordset.MoveLast
End If
End Sub
