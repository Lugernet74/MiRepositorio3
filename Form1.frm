VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "BD Simple :D"
   ClientHeight    =   6330
   ClientLeft      =   6450
   ClientTop       =   4020
   ClientWidth     =   5460
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6330
   ScaleWidth      =   5460
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command8 
      Caption         =   "Limpiar Lista"
      Height          =   375
      Left            =   3720
      TabIndex        =   11
      Top             =   4800
      Width           =   1455
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Buscar Todos"
      Height          =   375
      Left            =   3720
      TabIndex        =   10
      Top             =   4080
      Width           =   1455
   End
   Begin VB.ListBox List1 
      Height          =   2010
      Left            =   480
      TabIndex        =   9
      Top             =   3960
      Width           =   2895
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Borrar Registro"
      Height          =   375
      Left            =   1320
      TabIndex        =   8
      Top             =   2520
      Width           =   2775
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Buscar"
      Height          =   375
      Left            =   3720
      TabIndex        =   7
      Top             =   3240
      Width           =   1455
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   480
      TabIndex        =   6
      Text            =   "Buscar x Nombre"
      Top             =   3240
      Width           =   2895
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Siguiente"
      Height          =   375
      Left            =   4200
      TabIndex        =   5
      Top             =   1800
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Anterior"
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   1800
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Limpiar"
      Height          =   375
      Left            =   2880
      TabIndex        =   3
      Top             =   1800
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Top             =   960
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Agregar"
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   1800
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      DataSource      =   "Adodc1"
      Height          =   405
      Left            =   1320
      TabIndex        =   0
      Top             =   240
      Width           =   2775
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As Recordset
Dim db As Connection
Dim pathBD As String

Private Sub Command1_Click()
' agregar datos en las tables

If Len(Text1.Text) <> 0 And Len(Text2.Text) <> 0 Then
rs.AddNew
rs("Nombre") = Text1.Text
rs("Apellido") = Text2.Text
rs.Update
Text1.Text = ""
Text2.Text = ""
End If
Text1.SetFocus
End Sub

Private Sub Command2_Click()
Text1.Text = ""
Text2.Text = ""
Text1.SetFocus
End Sub

Private Sub Command3_Click()
rs.MovePrevious
If rs.BOF Then
rs.MoveLast
End If
Text1.Text = rs.Fields("Nombre")
Text2.Text = rs.Fields("Apellido")
End Sub

Private Sub Command4_Click()
rs.MoveNext
If rs.EOF Then
rs.MoveFirst
End If
Text1.Text = rs.Fields("Nombre")
Text2.Text = rs.Fields("Apellido")
End Sub

Private Sub Command5_Click()
rs.Close
rs.Open "select * from Datos where Nombre = '" & Text3.Text & "'", db, adOpenDynamic, adLockOptimistic

If Not (rs.EOF And rs.BOF) Then
Text1.Text = rs.Fields("Nombre")
Text2.Text = rs.Fields("Apellido")
Else
Text3.Text = "No se encontró registro"
Text1.Text = ""
Text2.Text = ""
End If
rs.Close
rs.Open "select * from Datos", db, adOpenDynamic, adLockOptimistic
End Sub

Private Sub Command6_Click()

If Not (rs.EOF And rs.BOF) Then
rs.Close
rs.Open "delete * from Datos where Nombre = '" & Text1.Text & "' and Apellido = '" & Text2.Text & "'", db, adOpenDynamic, adLockOptimistic
rs.Open "select * from Datos", db, adOpenDynamic, adLockOptimistic
rs.MoveNext
Text1.Text = rs.Fields("Nombre")
Text2.Text = rs.Fields("Apellido")
Text3.Text = "Registro Borrado"

Else
Text3.Text = "No se encontró registro"
Text1.Text = ""
Text2.Text = ""
rs.MoveFirst
End If

rs.Close
rs.Open "select * from Datos", db, adOpenDynamic, adLockOptimistic
End Sub

Private Sub Command7_Click()
rs.Close
rs.Open "select * from Datos where Nombre = '" & Text3.Text & "'", db, adOpenDynamic, adLockOptimistic

If Not (rs.EOF And rs.BOF) Then

Do Until (rs.EOF Or rs.BOF)
List1.AddItem (rs.Fields("Nombre") & " " & rs.Fields("Apellido"))
rs.MoveNext
Loop

Else
Text3.Text = "No se encontró registro"
List1.Clear
End If
rs.Close
rs.Open "select * from Datos", db, adOpenDynamic, adLockOptimistic
End Sub

Private Sub Command8_Click()
List1.Clear
End Sub

Private Sub Form_Load()
Set db = New Connection
Set rs = New Recordset

pathBD = App.Path & "\bd1.mdb"

db.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & pathBD & ";"
rs.Open "select * from Datos", db, adOpenDynamic, adLockOptimistic

' Para access 2002  ó xp.
' Anteriores usar PROVIDER=Microsoft.Jet.OLEDB.3.51;

End Sub

Private Sub Form_Unload(Cancel As Integer)
rs.Close
db.Close
End Sub

Private Sub Text3_GotFocus()
Text3.Text = ""
End Sub
