VERSION 5.00
Begin VB.Form frmImport 
   Caption         =   "Import Student"
   ClientHeight    =   3528
   ClientLeft      =   60
   ClientTop       =   456
   ClientWidth     =   6024
   LinkTopic       =   "Form1"
   ScaleHeight     =   3528
   ScaleWidth      =   6024
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdImportStudents 
      Caption         =   "Import"
      BeginProperty Font 
         Name            =   "Bodoni MT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   700
      Left            =   768
      TabIndex        =   3
      Top             =   2700
      Width           =   2000
   End
   Begin VB.FileListBox File1 
      BeginProperty Font 
         Name            =   "Bodoni MT"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2088
      Left            =   1956
      Pattern         =   "*.csv"
      TabIndex        =   2
      Top             =   468
      Width           =   3972
   End
   Begin VB.DirListBox Dir1 
      BeginProperty Font 
         Name            =   "Bodoni MT"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2004
      Left            =   120
      TabIndex        =   1
      Top             =   516
      Width           =   1695
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Bodoni MT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   700
      Left            =   3900
      TabIndex        =   0
      Top             =   2712
      Width           =   2000
   End
   Begin VB.Label Label1 
      Caption         =   "Please select student import file:"
      BeginProperty Font 
         Name            =   "Bodoni MT"
         Size            =   10.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   4215
   End
End
Attribute VB_Name = "frmImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim conn As ADODB.Connection, rec As ADODB.Recordset

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdImportStudents_Click()

    If File1.FileName = "" Then
        Call MsgBox("Please Select a .CSV file", vbExclamation + vbOKOnly, "Import")
        Exit Sub
    End If
    
    'check if student table already have data
    If recCount(TB_STUDENT) > 0 Then
        If vbYes = MsgBox("Student data already exists, do you want to overwrite it ?", vbQuestion + vbYesNo, "Import") Then
            truncateStudentTable
        Else
            Exit Sub
        End If
    End If
    
    'import .csv file to student table
    If ImportTextToAccessADO(Dir1.Path, File1.FileName) Then
        Call MsgBox("Import successfully", vbInformation + vbOKOnly, "Import")
    Else
        Call MsgBox("Import failed", vbCritical + vbOKOnly, "Import")
    End If
    
End Sub

Private Sub Dir1_Change()
'refresh filelist box, only support  .csv file import

    File1.Path = Dir1.Path
    File1.Pattern = "*.csv"

End Sub

Private Sub Form_Load()
    
    Dir1.Path = "D:\cliu\Homework"
    File1.Path = Dir1.Path

End Sub
