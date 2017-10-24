VERSION 5.00
Begin VB.Form frmStudent 
   Caption         =   "Form1"
   ClientHeight    =   5025
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8715
   LinkTopic       =   "Form1"
   ScaleHeight     =   5025
   ScaleWidth      =   8715
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
      Height          =   495
      Left            =   5040
      TabIndex        =   15
      Top             =   3840
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
      Height          =   495
      Left            =   6600
      TabIndex        =   14
      Top             =   3840
      Width           =   1335
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
      Height          =   495
      Left            =   3480
      TabIndex        =   13
      Top             =   3840
      Width           =   1335
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
      Height          =   495
      Left            =   1920
      TabIndex        =   12
      Top             =   3840
      Width           =   1335
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
      Height          =   495
      Left            =   360
      TabIndex        =   11
      Top             =   3840
      Width           =   1335
   End
   Begin VB.CommandButton cmdsearch 
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
      Height          =   495
      Left            =   6600
      TabIndex        =   10
      Top             =   3120
      Width           =   1335
   End
   Begin VB.TextBox txtAge 
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
      Left            =   2640
      TabIndex        =   9
      Top             =   3120
      Width           =   3375
   End
   Begin VB.TextBox txtYear 
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
      Left            =   2640
      TabIndex        =   8
      Top             =   2400
      Width           =   3375
   End
   Begin VB.TextBox txtCourse 
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
      Left            =   2640
      TabIndex        =   7
      Top             =   1680
      Width           =   5535
   End
   Begin VB.TextBox txtName 
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
      Left            =   2640
      TabIndex        =   6
      Top             =   960
      Width           =   5535
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
      Height          =   495
      Left            =   2640
      TabIndex        =   5
      Top             =   240
      Width           =   5535
   End
   Begin VB.Label Label5 
      Caption         =   "Age:"
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
      Left            =   360
      TabIndex        =   4
      Top             =   3120
      Width           =   1575
   End
   Begin VB.Label Label4 
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
      Left            =   360
      TabIndex        =   3
      Top             =   960
      Width           =   2055
   End
   Begin VB.Label Label3 
      Caption         =   "Course:"
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
      Left            =   360
      TabIndex        =   2
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Year:"
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
      Left            =   360
      TabIndex        =   1
      Top             =   2400
      Width           =   1575
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
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "frmStudent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As Database
Dim rsStudent As Recordset

Private Sub cmdAdd_Click()
    cmdDelete.Enabled = False
    cmdEdit.Enabled = False
    SQL = "select * from Student where studentID = '" & txtID.Text & "'"
    Set rsStudent = db.OpenRecordset(SQL)
        If rsStudent.BOF = False Then
            MsgBox "ID number already exists"
        Else
            SQL = "insert into Student values('" & txtID.Text & "','" & txtName.Text & "','" & txtCourse.Text & "','" & txtYear.Text & "','" & txtAge.Text & "')"
            db.Execute (SQL)
    End If
    txtName.Text = ""
    txtID.Text = ""
    txtAge.Text = ""
    txtCourse.Text = ""
    txtYear.Text = ""
    txtID.SetFocus
    
End Sub

Private Sub cmdClear_Click()
    txtID.Text = ""
    txtName.Text = ""
    txtCourse.Text = ""
    txtYear.Text = ""
    txtAge.Text = ""
    txtID.SetFocus
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdDelete_Click()
    cmdDelete.Enabled = False
    cmdEdit.Enabled = False
    SQL = "select * from student where studentID = '" & txtID.Text & "'"
    Set rsStudent = db.OpenRecordset(SQL)
    If rsStudent.BOF = False Then
        SQL = "delete from student where studentID = '" & txtID.Text & "'"
        db.Execute (SQL)
    End If
End Sub

Private Sub cmdEdit_Click()
    cmdDelete.Enabled = False
    cmdEdit.Enabled = False
    SQL = "select * from student where studentID = '" & txtID.Text & "'"
    Set rsStudent = db.OpenRecordset(SQL)
    If rsStudent.BOF = False Then
        SQL = "update Student set Studentname = '" & txtName.Text & "', course = '" & txtCourse.Text & "', year = '" & txtYear.Text & "', age = '" & txtAge.Text & "'"
        db.Execute (SQL)
    End If
End Sub

Private Sub cmdsearch_Click()
    SQL = "select * from Student where studentID = '" & txtID.Text & "'"
    Set rsStudent = db.OpenRecordset(SQL)
    
    If rsStudent.BOF = True Then
        MsgBox "ID No. doesn't exist"
    Else
        txtName.Text = rsStudent.Fields("StudentName")
        txtCourse.Text = rsStudent.Fields(2)
        txtYear.Text = rsStudent.Fields("Age")
        txtAge.Text = rsStudent.Fields("Age")
        cmdDelete.Enabled = True
        cmdEdit.Enabled = True
    End If
    
End Sub

Private Sub Form_Load()
    cmdDelete.Enabled = False
    cmdEdit.Enabled = False
    Set db = OpenDatabase(App.Path & "\LibrarySystem.mdb")
    Set rsStudent = db.OpenRecordset("Student")
End Sub

