VERSION 5.00
Begin VB.Form frmReport 
   Caption         =   "Report"
   ClientHeight    =   1896
   ClientLeft      =   3552
   ClientTop       =   4500
   ClientWidth     =   5112
   LinkTopic       =   "Form1"
   ScaleHeight     =   1896
   ScaleWidth      =   5112
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
      Left            =   2880
      TabIndex        =   2
      Top             =   1008
      Width           =   2000
   End
   Begin VB.CommandButton cmdExcel 
      Caption         =   "Export Report"
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
      Top             =   1020
      Width           =   2000
   End
   Begin VB.Label lblFn 
      BeginProperty Font 
         Name            =   "Bodoni MT"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   408
      Left            =   600
      TabIndex        =   3
      Top             =   540
      Width           =   3924
   End
   Begin VB.Label Label1 
      Caption         =   "Export student mark to Excel and Word file"
      BeginProperty Font 
         Name            =   "Bodoni MT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   384
      Left            =   144
      TabIndex        =   1
      Top             =   144
      Width           =   4680
   End
End
Attribute VB_Name = "frmReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExcel_Click()
    Dim cnn As ADODB.Connection
    Dim rst As ADODB.Recordset
    Dim sSql As String
    'Vars for Excel
    Dim xlApp As Excel.Application
    Dim xlWBook As Excel.Workbook
    Dim xlWSheet As Excel.Worksheet
    Dim xlRReport As Excel.Range, xlRData As Excel.Range
    'Vars for MS Word.
    Dim wdApp As Word.Application
    Dim wdDoc As Word.Document
    
    'Instantiate the MS Excel-objects.
    Set xlApp = New Excel.Application
    Set xlWBook = xlApp.Workbooks.Open(App.Path & "\StudentMark.xls")
    Set xlWSheet = xlWBook.Worksheets("Marks")
    
    lblFn.Caption = "Processing ... "
    'Disable excel automation warning messages
    xlApp.DisplayAlerts = False
    With xlWSheet
         'Here we use a named range which holds the whole reporttable.
        Set xlRReport = .Range("MarkReport")
         'Here we use a named range that holds the data from the database.
        Set xlRData = .Range("MarkList")
    End With
     
    'Instantiate the ADO-objects.
    Set cnn = New ADODB.Connection
    Set rst = New ADODB.Recordset
    
    cnn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                   "Data Source= " & DB_PATH & ";" & _
                   "Jet OLEDB:Engine Type=4;"
    sSql = "SELECT * FROM data "
    rst.Open sSql, cnn
    'Copy the recordset to the table in StudentMark.xls.
    xlRData.CopyFromRecordset rst

    'Release ADO-objects from memory.
    rst.Close
    cnn.Close
    Set rst = Nothing
    Set cnn = Nothing
    
    'Instantiate the MS Word-objects
    Call FileCopy(App.Path & "\StudentMarkReport.dotx", App.Path & "\StudentMark.doc")
    Set wdApp = New Word.Application
    Set wdDoc = wdApp.Documents.Open(App.Path & "\StudentMark.doc")
     
    On Error Resume Next
    'Here we copy the EXCEL Report table in the worksheet.
    xlRReport.Copy
    wdDoc.Content.PasteExcelTable False, False, False
     
    
closeApp:
    On Error Resume Next
    wdDoc.Close savechanges:=True
    'Release objects from the memory.
    'Close MS Word.
    wdApp.Quit
    Set rbmReport = Nothing
    Set wdDoc = Nothing
    Set wdApp = Nothing
   
   'Close and save the Excel-workbook.
    xlWBook.Close savechanges:=True
    xlApp.Quit
    On Error GoTo 0
    lblFn.Caption = "StudentMark.xls  && StudentMark.doc"
    Call MsgBox("Excel and Word files exported!", vbInformation + vbOKOnly, "Export")
    
End Sub


Private Sub cmdExit_Click()
    Unload Me
End Sub
