VERSION 5.00
Begin VB.Form Form7 
   Caption         =   "Inicio"
   ClientHeight    =   4470
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4275
   BeginProperty Font 
      Name            =   "Magneto"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form7"
   ScaleHeight     =   4470
   ScaleWidth      =   4275
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command7 
      Caption         =   "Regresar"
      Height          =   420
      Left            =   1320
      TabIndex        =   7
      Top             =   3480
      Width           =   1695
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Cliente"
      Height          =   615
      Left            =   2400
      TabIndex        =   6
      Top             =   2520
      Width           =   1455
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Disco"
      Height          =   495
      Left            =   2400
      TabIndex        =   5
      Top             =   1800
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Película"
      Height          =   615
      Left            =   2400
      TabIndex        =   4
      Top             =   960
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Alquiler"
      Height          =   615
      Left            =   600
      TabIndex        =   3
      Top             =   2520
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Autor"
      Height          =   540
      Left            =   600
      TabIndex        =   2
      Top             =   1800
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Tipo de Película"
      Height          =   615
      Left            =   600
      TabIndex        =   1
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Inicio"
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
      Width           =   990
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Me.Hide
Form1.Show
End Sub

Private Sub Command2_Click()
Me.Hide
Form3.Show
End Sub

Private Sub Command3_Click()
Me.Hide
Form5.Show
End Sub

Private Sub Command4_Click()
Me.Hide
Form2.Show
End Sub

Private Sub Command5_Click()
Me.Hide
Form2.Show
End Sub

Private Sub Command6_Click()
Me.Hide
Form6.Show
End Sub

Private Sub Command7_Click()
Me.Hide
Form8.Show
End Sub
