VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Student Grades"
   ClientHeight    =   3360
   ClientLeft      =   132
   ClientTop       =   420
   ClientWidth     =   5280
   LinkTopic       =   "Form1"
   ScaleHeight     =   3360
   ScaleWidth      =   5280
   StartUpPosition =   3  'Windows Default
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
      Left            =   1704
      TabIndex        =   4
      Top             =   2340
      Width           =   2000
   End
   Begin VB.CommandButton cmdReport 
      Caption         =   "Report"
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
      Left            =   2988
      TabIndex        =   3
      Top             =   1284
      Width           =   2000
   End
   Begin VB.CommandButton cmdAddStudent 
      Caption         =   "Add Student"
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
      Left            =   252
      TabIndex        =   2
      Top             =   1296
      Width           =   2000
   End
   Begin VB.CommandButton cmdListCourse 
      Caption         =   "List Courses"
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
      Left            =   2976
      TabIndex        =   1
      Top             =   312
      Width           =   2000
   End
   Begin VB.CommandButton cmdImport 
      Caption         =   "Import Students"
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
      Left            =   264
      TabIndex        =   0
      Top             =   312
      Width           =   2000
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAddStudent_Click()
    frmAdd.Show
End Sub

Private Sub cmdExit_Click()
    Unload Me
    End
End Sub

Private Sub cmdImport_Click()
    frmImport.Show
End Sub

Private Sub cmdListCourse_Click()
    frmCourses.Show
End Sub

Private Sub cmdReport_Click()
    frmReport.Show
End Sub

Private Sub Form_Load()
    frmMain.Show modal
    DB_PATH = App.Path & "\" & DB_Name
    RPT_PATH = App.Path & "\"
End Sub
