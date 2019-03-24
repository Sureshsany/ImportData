VERSION 5.00
Begin VB.Form frmDownload 
   Caption         =   "Time Tracker: Data Move"
   ClientHeight    =   6435
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8310
   LinkTopic       =   "Form1"
   ScaleHeight     =   6435
   ScaleWidth      =   8310
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOpen 
      Caption         =   "SQL"
      Height          =   615
      Left            =   6720
      TabIndex        =   10
      Top             =   480
      Width           =   1095
   End
   Begin VB.ComboBox cboDB 
      Height          =   315
      ItemData        =   "DB Migration.frx":0000
      Left            =   1560
      List            =   "DB Migration.frx":000D
      OLEDropMode     =   1  'Manual
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   600
      Width           =   2295
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H0000C000&
      Caption         =   "Exit"
      Default         =   -1  'True
      Height          =   615
      Left            =   6960
      TabIndex        =   7
      Top             =   5640
      Width           =   975
   End
   Begin VB.CommandButton cmdTrasfer 
      Caption         =   "Move Data"
      Height          =   615
      Left            =   4320
      TabIndex        =   0
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
      Caption         =   "From DB"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   600
      TabIndex        =   9
      Top             =   720
      Width           =   735
   End
   Begin VB.Label lblFinal 
      Alignment       =   2  'Center
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1320
      TabIndex        =   6
      Top             =   5520
      Width           =   5175
   End
   Begin VB.Label lblTimeDiff 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   480
      TabIndex        =   5
      Top             =   3840
      Width           =   6375
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      Caption         =   "Label1"
      Height          =   735
      Left            =   1440
      TabIndex        =   4
      Top             =   4440
      Width           =   4935
   End
   Begin VB.Label lblTarget 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      Height          =   195
      Left            =   1320
      TabIndex        =   3
      Top             =   3120
      Width           =   480
   End
   Begin VB.Label lblSource 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      Height          =   195
      Left            =   1320
      TabIndex        =   2
      Top             =   2400
      Width           =   480
   End
   Begin VB.Label lblPath 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      Height          =   195
      Left            =   1320
      TabIndex        =   1
      Top             =   1680
      Width           =   480
   End
End
Attribute VB_Name = "frmDownload"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    End
End Sub


Public Function OpenDB() As Boolean
    Dim isOpen      As Boolean
    Dim ANS         As VbMsgBoxResult
    Dim CN          As ADODB.Connection
    Set CN = New ADODB.Connection
    Dim DBPath                       As String
    DBPath = App.Path & "\TimeDB.mdb"
    isOpen = False
  '  On Error GoTo err
        Do Until isOpen = True
                CN.CursorLocation = adUseClient
                'CN.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DBPath & ";Persist Security Info=False;Jet OLEDB:Database Password=philiprj"
                
                If InStr(DBPath, "DSN:") Then
                    CN.Open Split(DBPath, ":")(1), , "rAMAcHANDRA"
                Else
                    CN.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DBPath & ";Persist Security Info=False;Jet OLEDB:Database Password=rAMAcHANDRA"
                End If
            isOpen = True
        Loop
        OpenDB = isOpen
    Exit Function
err:
    ANS = MsgBox("Error # " & err.Number & vbCrLf & "Description: " & err.Description, vbCritical + vbRetryCancel)
    If ANS = vbCancel Then
        OpenDB = False
    ElseIf ANS = vbRetry Then
        Resume
    End If
End Function

Private Sub cboDB_Click()
     sDB = cboDB.List(cboDB.ListIndex) & ".mdb"
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub getDBDetails()
    Dim File As String
    Dim sf As FileSystemObject
    Set sf = New FileSystemObject
    
    File = App.Path & "\setting.ini"
    If Not sf.FileExists(App.Path & "\setting.ini") Then
        MsgBox "Database base entry not found"
    Else
        sPath = ReadIni(File, "source", "Path")
        sTabel = ReadIni(File, "dbtable", "table")
        sColDate = ReadIni(File, "Coldate", "ColDate")
        sColTime = ReadIni(File, "Coltime", "ColTime")
        sColLogID = ReadIni(File, "Colloginid", "LoginID")
    End If
    
End Sub

Private Sub cmdTrasfer_Click()
    Dim sMsg    As String
    
    If cboDB.ListIndex = -1 Then
        MsgBox "Please select the database from where data needs to imported", vbInformation, "Time Tracker"
        Exit Sub
    ElseIf cboDB.List(cboDB.ListIndex) = "File" Then
        getDBDetails
    End If
    
    cmdExit.Enabled = False
    sMsg = ReadPath
    If InStr(sMsg, "Error") Then
        With lblPath
            .BackColor = vbRed
            .FontSize = 15
            .Caption = "1. " & sMsg
        End With
        Exit Sub
    Else
        With lblPath
            .BackColor = vbGreen
            .FontSize = 15
            .Caption = "1. Source and Target database path found."
        End With
    End If
    
    sMsg = OpenSourceDB
    If InStr(sMsg, "Error") Then
        With lblSource
            .BackColor = vbRed
            .FontSize = 15
            .Caption = "2. " & sMsg
        End With
        Exit Sub
    Else
        With lblSource
            .BackColor = vbGreen
            .FontSize = 15
            .Caption = "2. Source Database Connected"
        End With
    End If
    
    sMsg = OpenTargetDB
    If InStr(sMsg, "Error") Then
        With lblTarget
            .BackColor = vbRed
            .FontSize = 15
            .Caption = "3. " & sMsg
        End With
        Exit Sub
    Else
        With lblTarget
            .BackColor = vbGreen
            .FontSize = 15
            .Caption = "3. Target Database Connected"
        End With
    End If
    
    If MsgBox("Are you sure you want to move the data from .. " & sDB & "?", vbInformation + vbYesNo, "Time Tracker") = vbYes Then
        Select Case sDB
            Case "ATT2000.mdb"
                AttDownloadData
            Case "Warden.mdb"
                WardenDownloadData
            Case "File.mdb"
                FileLoad
        End Select
    End If
    
    CloseDatabase
    lblFinal.Caption = "Completed..."
    cmdExit.Enabled = True
  '  ClearLable
    
End Sub

Private Sub ClearLable()
    lblPath.Caption = ""
    lblSource = ""
    lblTarget = ""
    lblTimeDiff = ""
    lblFinal = ""
    With lblStatus
        .FontSize = 12
        .Caption = ""
        .FontBold = True
    End With
    
End Sub

Private Sub Form_Load()
    ClearLable
'    LoadDB
End Sub

Private Sub LoadDB()
    Dim sFile       As String
    
    cboDB.Clear
    
    sFile = GetSetting
    
    
End Sub


