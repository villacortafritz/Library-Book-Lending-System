VERSION 5.00
Begin VB.Form frmReturn 
   Caption         =   "Form2"
   ClientHeight    =   3645
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6585
   LinkTopic       =   "Form2"
   ScaleHeight     =   3645
   ScaleWidth      =   6585
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
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
      Left            =   4560
      TabIndex        =   11
      Top             =   3000
      Width           =   1815
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return"
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
      Left            =   2520
      TabIndex        =   10
      Top             =   3000
      Width           =   1815
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
      Height          =   495
      Left            =   2520
      TabIndex        =   4
      Top             =   0
      Width           =   3855
   End
   Begin VB.ComboBox txtBorrowed 
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
      Left            =   2520
      TabIndex        =   3
      Top             =   600
      Width           =   3855
   End
   Begin VB.TextBox txtDue 
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
      Left            =   2520
      TabIndex        =   2
      Top             =   1200
      Width           =   3855
   End
   Begin VB.TextBox txtReturned 
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
      Left            =   2520
      TabIndex        =   1
      Top             =   1800
      Width           =   3855
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
      Height          =   495
      Left            =   2520
      TabIndex        =   0
      Top             =   2400
      Width           =   3855
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
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Borrowed Books:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   600
      Width           =   2295
   End
   Begin VB.Label Label3 
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
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "Date Returned:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   1800
      Width           =   2055
   End
   Begin VB.Label Label5 
      Caption         =   "Penalty:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   2400
      Width           =   1095
   End
End
Attribute VB_Name = "frmReturn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As Database
Dim rsStudent As Recordset
Dim rsBook As Recordset
Dim dateReturned As Date
Dim rsReturned As Recordset
Dim Borrow As Recordset

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdReturn_Click()
    SQL = "select * from Student where studentID = '" & txtID.Text & "'"
    Set rsStudent = db.OpenRecordset(SQL)
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

Private Sub Form_Load()
    Set db = OpenDatabase(App.Path & "\Library System.mdb")
    Set rsStudent = db.OpenRecordset("Student")
    Set rsBook = db.OpenRecordset("Book")
    Set rsBorrow = db.OpenRecordset("Borrow")
    Set rsReturned = db.OpenRecordset("Return")
    cmdReturn.Enabled = False
End Sub
