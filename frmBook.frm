VERSION 5.00
Begin VB.Form frmBook 
   Caption         =   "Form1"
   ClientHeight    =   4830
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10290
   LinkTopic       =   "Form1"
   ScaleHeight     =   4830
   ScaleWidth      =   10290
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8640
      TabIndex        =   15
      Top             =   240
      Width           =   1335
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7080
      TabIndex        =   14
      Top             =   240
      Width           =   1335
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8640
      TabIndex        =   13
      Top             =   3960
      Width           =   1455
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6840
      TabIndex        =   12
      Top             =   3960
      Width           =   1455
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Edit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5040
      TabIndex        =   11
      Top             =   3960
      Width           =   1455
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3240
      TabIndex        =   10
      Top             =   3960
      Width           =   1455
   End
   Begin VB.TextBox txtDays 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3240
      TabIndex        =   9
      Top             =   3120
      Width           =   6855
   End
   Begin VB.TextBox txtPenalty 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3240
      TabIndex        =   8
      Top             =   2400
      Width           =   6855
   End
   Begin VB.TextBox txtCopies 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3240
      TabIndex        =   7
      Top             =   1680
      Width           =   6855
   End
   Begin VB.TextBox txtTitle 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3240
      TabIndex        =   3
      Top             =   960
      Width           =   6855
   End
   Begin VB.TextBox txtID 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3240
      TabIndex        =   0
      Top             =   240
      Width           =   3615
   End
   Begin VB.Label Label5 
      Caption         =   "No. of Days Allowed:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   6
      Top             =   3240
      Width           =   2775
   End
   Begin VB.Label Label4 
      Caption         =   "Penalty Fee:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   5
      Top             =   2520
      Width           =   2775
   End
   Begin VB.Label Label3 
      Caption         =   "No. of Copies:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   1800
      Width           =   2775
   End
   Begin VB.Label Label2 
      Caption         =   "Book Name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   1080
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "Book ID:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   2775
   End
End
Attribute VB_Name = "frmBook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As Database
Dim rsBook As Recordset

Private Sub cmdAdd_Click()
    cmdDelete.Enabled = False
    cmdEdit.Enabled = False
    SQL = "select * from Book where BookID = '" & txtID.Text & "'"
    Set rsBook = db.OpenRecordset(SQL)
    If rsBook.BOF = False Then
        MsgBox "There is already an existing ID"
    Else
        If txtTitle.Text = "" Or txtCopies.Text = "" Or txtPenalty.Text = "" Or txtDays.Text = "" Then
         MsgBox "Invalid input"
        Else
            SQL = "insert into Book values ('" & txtID.Text & "', '" & txtTitle.Text & "', '" & txtCopies.Text & "', '" & txtPenalty.Text & "', '" & txtDays.Text & "')"
            db.Execute (SQL)
        End If
    End If
    txtID.Text = ""
    txtTitle.Text = ""
    txtCopies.Text = ""
    txtPenalty = ""
    txtDays = ""
    txtID.SetFocus
    
End Sub

Private Sub cmdClear_Click()
    txtID.Text = ""
    txtTitle.Text = ""
    txtCopies.Text = ""
    txtPenalty = ""
    txtDays = ""
    txtID.SetFocus
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdDelete_Click()
    cmdDelete.Enabled = False
    cmdEdit.Enabled = False
    SQL = "select * from Book where BookID = '" & txtID.Text & "'"
    Set rsBook = db.OpenRecordset(SQL)
    If rsBook.BOF = False Then
        SQL = "delete from Book where BookID = '" & txtID.Text & "'"
        db.Execute (SQL)
    End If
    txtID.Text = ""
    txtTitle.Text = ""
    txtCopies.Text = ""
    txtPenalty = ""
    txtDays = ""
    txtID.SetFocus
End Sub

Private Sub cmdEdit_Click()
    cmdDelete.Enabled = False
    cmdEdit.Enabled = False
    SQL = "select * from BookID where BookID = '" & txtID.Text & "'"
    Set rsBook = db.OpenRecordset(SQL)
    If rsBook.BOF = False Then
        SQL = " update Book set title = '" & txtTitle.Text & "', NoOfCopies = '" & txtCopies.Text & "', PenaltyFee = '" & txtPenalty.Text & "', NoOfDaysAllowed = '" & txtDays.Text & "'"
        db.Execute (SQL)
    End If
        
End Sub

Private Sub cmdsearch_Click()
    SQL = "select * from Book where BookID = '" & txtID.Text & "'"
    Set rsBook = db.OpenRecordset(SQL)
    If rsBook.BOF = True Then
        MsgBox "ID number does not exist"
    Else
        txtTitle.Text = rsBook.Fields("title")
        txtCopies.Text = rsBook.Fields("NoOfCopies")
        txtPenalty = rsBook.Fields("PenaltyFee")
        txtDays = rsBook.Fields("NoOfDaysAllowed")
        cmdDelete.Enabled = True
        cmdEdit.Enabled = True
    End If
End Sub

Private Sub Form_Load()
    cmdDelete.Enabled = False
    cmdEdit.Enabled = False
    Set db = OpenDatabase(App.Path & "\Library System.mdb")
    Set rsBook = db.OpenRecordset("Book")
End Sub
