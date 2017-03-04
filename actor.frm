VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H00808080&
   Caption         =   "Form3"
   ClientHeight    =   8925
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9930
   LinkTopic       =   "Form3"
   ScaleHeight     =   8925
   ScaleWidth      =   9930
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFFF00&
      Caption         =   "ELIMINAR"
      Height          =   615
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   6360
      Width           =   2175
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFF00&
      Caption         =   "MODIFICAR"
      Height          =   615
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5640
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFF00&
      Caption         =   "GUARDAR"
      Height          =   615
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5520
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFF00&
      Caption         =   "CREAR"
      Height          =   615
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6360
      Width           =   2295
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Users\estudiante.SERVERINT\Desktop\ELIZABETH\Tienda de Discos.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   2400
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Actor"
      Top             =   5640
      Width           =   1815
   End
   Begin VB.TextBox Text3 
      DataField       =   "Fecha_Nac "
      DataSource      =   "Data1"
      Height          =   735
      Left            =   3000
      TabIndex        =   6
      Top             =   4440
      Width           =   4215
   End
   Begin VB.TextBox Text2 
      Height          =   735
      Left            =   3000
      TabIndex        =   5
      Top             =   3360
      Width           =   4215
   End
   Begin VB.TextBox Text1 
      DataField       =   "Código"
      DataSource      =   "Data1"
      Height          =   735
      Left            =   3000
      TabIndex        =   4
      Top             =   2280
      Width           =   4215
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "FECHA_NAC"
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   3
      Top             =   4560
      Width           =   2895
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "NOMBRE"
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   2
      Top             =   3480
      Width           =   2895
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "CODIGO"
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      TabIndex        =   1
      Top             =   2280
      Width           =   2775
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "ACTOR"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   3000
      TabIndex        =   0
      Top             =   360
      Width           =   3735
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Data1.Recordset.AddNew
End Sub

Private Sub Command2_Click()
Data1.Recordset.Update
End Sub

Private Sub Command4_Click()
Data1.Recordset.Delete
End Sub

