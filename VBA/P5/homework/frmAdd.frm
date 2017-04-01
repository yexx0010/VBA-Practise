VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmAdd 
   Caption         =   "Students "
   ClientHeight    =   6960
   ClientLeft      =   60
   ClientTop       =   456
   ClientWidth     =   11160
   LinkTopic       =   "Form1"
   ScaleHeight     =   6960
   ScaleWidth      =   11160
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkEdit 
      Caption         =   "Editable"
      BeginProperty Font 
         Name            =   "Bodoni MT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   192
      Left            =   6492
      TabIndex        =   3
      Top             =   6420
      Width           =   1344
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
      Left            =   8892
      TabIndex        =   2
      Top             =   6120
      Width           =   2000
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
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
      Left            =   576
      TabIndex        =   1
      Top             =   6120
      Width           =   2000
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   648
      Left            =   3084
      Top             =   6156
      Width           =   3144
      _ExtentX        =   5546
      _ExtentY        =   1143
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=registrar.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=registrar.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Data"
      Caption         =   "Students"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Bodoni MT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataGridLib.DataGrid dgStudent 
      Bindings        =   "frmAdd.frx":0000
      Height          =   5772
      Left            =   132
      TabIndex        =   0
      Top             =   168
      Width           =   10812
      _ExtentX        =   19071
      _ExtentY        =   10181
      _Version        =   393216
      AllowUpdate     =   -1  'True
      HeadLines       =   1
      RowHeight       =   27
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   4105
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   4105
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

End Sub

Private Sub chkEdit_Click()

    If chkEdit.Value = vbChecked Then
        dgStudent.AllowUpdate = True
    Else
        dgStudent.AllowUpdate = False
    End If
    
End Sub

Private Sub cmdAdd_Click()
'enable write mode in datagrid and also move cursor to the the last row
    dgStudent.AllowAddNew = True
    dgStudent.AllowUpdate = True
    chkEdit.Value = vbChecked
    dgStudent.Row = dgStudent.ApproxCount
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdLock_Click()
'write protect
    dgStudent.AllowAddNew = False
    cmdLock.Enabled = False
End Sub

Private Sub Form_Load()
    Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DB_PATH & ";Persist Security Info=False"
    chkEdit.Value = Unchecked
    dgStudent.AllowUpdate = False
    
End Sub
