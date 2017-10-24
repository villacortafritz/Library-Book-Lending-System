VERSION 5.00
Begin VB.Form frmBorrow 
   Caption         =   "Form1"
   ClientHeight    =   5235
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7050
   LinkTopic       =   "Form1"
   ScaleHeight     =   5235
   ScaleWidth      =   7050
   StartUpPosition =   3  'Windows Default
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
      Left            =   4800
      TabIndex        =   18
      Top             =   4440
      Width           =   1455
   End
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
      Left            =   3000
      TabIndex        =   17
      Top             =   4440
      Width           =   1455
   End
   Begin VB.CommandButton cmdBorrow 
      Caption         =   "Borrow"
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
      Left            =   5520
      TabIndex        =   16
      Top             =   2520
      Width           =   1335
   End
   Begin VB.CommandButton cmdBookSearch 
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
      Left            =   5520
      TabIndex        =   15
      Top             =   1320
      Width           =   1335
   End
   Begin VB.CommandButton cmdStudentSearch 
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
      Left            =   5520
      TabIndex        =   14
      Top             =   120
      Width           =   1335
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
      Left            =   2400
      TabIndex        =   13
      Top             =   3720
      Width           =   4455
   End
   Begin VB.TextBox txtDueDate 
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
      Left            =   2400
      TabIndex        =   12
      Top             =   3120
      Width           =   4455
   End
   Begin VB.TextBox txtDateBorrowed 
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
      Left            =   2400
      TabIndex        =   11
      Top             =   2520
      Width           =   3135
   End
   Begin VB.TextBox txtBookName 
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
      Left            =   2400
      TabIndex        =   10
      Top             =   1920
      Width           =   4455
   End
   Begin VB.TextBox txtBookID 
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
      Left            =   2400
      TabIndex        =   9
      Top             =   1320
      Width           =   3135
   End
   Begin VB.TextBox txtStudentName 
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
      Left            =   2400
      TabIndex        =   8
      Top             =   720
      Width           =   4455
   End
   Begin VB.TextBox txtStudentID 
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
      Left            =   2400
      TabIndex        =   7
      Top             =   120
      Width           =   3135
   End
   Begin VB.Label Label7 
      Caption         =   "No. of Copies Remaining"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   6
      Top             =   3720
      Width           =   2775
   End
   Begin VB.Label Label6 
      Caption         =   "Due Date:"
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
      Left            =   120
      TabIndex        =   5
      Top             =   3120
      Width           =   2775
   End
   Begin VB.Label Label5 
      Caption         =   "Date Borrowed:"
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
      Left            =   120
      TabIndex        =   4
      Top             =   2520
      Width           =   2775
   End
   Begin VB.Label Label4 
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
      Left            =   120
      TabIndex        =   3
      Top             =   1920
      Width           =   2775
   End
   Begin VB.Label Label3 
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
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   2775
   End
   Begin VB.Label Label2 
      Caption         =   "Student Name:"
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
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "Student ID:"
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
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2775
   End
End
Attribute VB_Name = "frmBorrow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As Database
Dim rsStudent As Recordset
Dim rsBook As Recordset
Dim dateBorrow As Date


Private Sub cmdBookSearch_Click()
    SQL = "select * from Book where BookID = '" & txtBookID.Text & "'"
    Set rsBook = db.OpenRecordset(SQL)
    If rsBook.BOF = True Or txtBookID = "" Then
        MsgBox "Id number does not exist."
    Else
        txtBookName.Text = rsBook.Fields("Title")
    End If
    If txtStudentName.Text <> "" And txtBookName.Text <> "" Then
        cmdBorrow.Enabled = True
    End If
End Sub

Private Sub cmdBorrow_Click()
    dateBorrow = Now()
    If rsBook.Fields("NoOfCopies") > 0 And txtStudentID.Text <> "" And txtBookID.Text <> "" Then
        txtDateBorrowed.Text = Format(DateBorrowed, "mm/dd/yy")
        txtDueDate.Text = Format(DateBorrowed + rsBook.Fields("NoOfDaysAllowed"), "mm/dd/yy")
        txtCopies.Text = rsBook.Fields("NoOfCopies")
        SQL = "insert into Borrow values('" & txtDateBorrowed.Text & "', '" & txtDueDate.Text & "', '" & txtStudentID.Text & "', '" & txtStudentName.Text & "', '" & txtBookID.Text & "', '" & txtBookName.Text & "')"
        db.Execute (SQL)
        SQL = "update book set NoOfCopies = '" & rsBook.Fields("NoOfCopies") - 1 & "' where BookID = '" & txtBookID.Text & "'"
        db.Execute (SQL)
    Else
        MsgBox "There are no more copies of the Book" + rsBook.Fields("Title")
    End If
    lblMessage.Caption = "Congrats! Nakaborrow ka ug Book! :)"
        
End Sub

Private Sub cmdClear_Click()
    txtStudentName.Text = ""
    txtBookName.Text = ""
    txtStudentID.Text = ""
    txtBookID.Text = ""
    txtDateBorrowed.Text = ""
    txtDueDate.Text = ""
    txtCopies.Text = ""
    
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdStudentSearch_Click()
    SQL = "select * from Student where StudentID = '" & txtStudentID.Text & "'"
    Set rsStudent = db.OpenRecordset(SQL)
    If rsStudent.BOF = True Or txtStudentID = "" Then
        MsgBox "Id number does not exist."
    Else
        txtStudentName.Text = rsStudent.Fields("StudentName")
    End If
    If txtStudentName.Text <> "" And txtBookName.Text <> "" Then
        cmdBorrow.Enabled = True
    End If
End Sub

Private Sub Form_Load()
    Set db = OpenDatabase(App.Path & "\Library System.mdb")
    Set rsStudent = db.OpenRecordset("Student")
    Set rsBook = db.OpenRecordset("Book")
    Set rsBorrowed = db.OpenRecordset("Borrow")
    cmdBorrow.Enabled = False
End Sub


