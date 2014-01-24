VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   ClientHeight    =   4575
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   7290
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   4575
   ScaleWidth      =   7290
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdGetURL 
      Height          =   375
      Left            =   6750
      OLEDropMode     =   1  'Manual
      Picture         =   "frmMain.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "Open script URL..."
      Top             =   780
      Width           =   435
   End
   Begin VB.CommandButton cmdAbortScript 
      Caption         =   "&Abort"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3480
      OLEDropMode     =   1  'Manual
      Picture         =   "frmMain.frx":0894
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3480
      Width           =   1095
   End
   Begin VB.CommandButton cmdExecute 
      Caption         =   "&Execute"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4800
      OLEDropMode     =   1  'Manual
      Picture         =   "frmMain.frx":155E
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3480
      Width           =   1095
   End
   Begin VB.CommandButton cmdGetScript 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6240
      OLEDropMode     =   1  'Manual
      Picture         =   "frmMain.frx":1E28
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Open script file..."
      Top             =   780
      Width           =   435
   End
   Begin VB.TextBox txtScript 
      Height          =   285
      Left            =   360
      OLEDropMode     =   1  'Manual
      TabIndex        =   0
      Top             =   840
      Width           =   5835
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Height          =   975
      Left            =   6120
      OLEDropMode     =   1  'Manual
      Picture         =   "frmMain.frx":23B2
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3480
      Width           =   1095
   End
   Begin VB.Frame fraOptions 
      Caption         =   "Options"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   120
      OLEDropMode     =   1  'Manual
      TabIndex        =   9
      Top             =   1320
      Width           =   7095
      Begin VB.CheckBox chkShowLog 
         Caption         =   "Show log after script ends"
         Height          =   255
         Left            =   240
         OLEDropMode     =   1  'Manual
         TabIndex        =   11
         Top             =   1560
         Width           =   2295
      End
      Begin VB.CheckBox chkUseScriptOptions 
         Caption         =   "Use settings specified in script for above options"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         OLEDropMode     =   1  'Manual
         TabIndex        =   4
         Top             =   1200
         Value           =   1  'Checked
         Width           =   3855
      End
      Begin VB.CheckBox chkUseRecycleBin 
         Caption         =   "Delete files to Recycle Bin when possible"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         OLEDropMode     =   1  'Manual
         TabIndex        =   3
         Top             =   720
         Width           =   3375
      End
      Begin VB.CheckBox chkUnloadShell 
         Caption         =   "Unload Explorer from memory before executing script"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         OLEDropMode     =   1  'Manual
         TabIndex        =   2
         Top             =   360
         Width           =   4095
      End
      Begin VB.Line linSeperator 
         BorderColor     =   &H80000014&
         Index           =   1
         X1              =   240
         X2              =   6840
         Y1              =   1095
         Y2              =   1095
      End
      Begin VB.Line linSeperator 
         BorderColor     =   &H80000010&
         Index           =   0
         X1              =   240
         X2              =   6840
         Y1              =   1080
         Y2              =   1080
      End
   End
   Begin VB.Frame fraProgress 
      Caption         =   "Progress"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   120
      TabIndex        =   12
      Top             =   1320
      Visible         =   0   'False
      Width           =   7095
      Begin VB.Label lblProgress 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0 %"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   360
         TabIndex        =   16
         Top             =   990
         Width           =   345
      End
      Begin VB.Shape shpProgress 
         BackColor       =   &H8000000D&
         BackStyle       =   1  'Opaque
         Height          =   495
         Left            =   240
         Top             =   840
         Width           =   1215
      End
      Begin VB.Shape shpProgressBackgrond 
         BackStyle       =   1  'Opaque
         Height          =   495
         Left            =   240
         Top             =   840
         Width           =   6615
      End
      Begin VB.Label lblCurrentAction 
         AutoSize        =   -1  'True
         Caption         =   "Nothing"
         Height          =   195
         Left            =   1440
         TabIndex        =   15
         Top             =   360
         Width           =   555
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "Current action:"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   14
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame fraLog 
      Caption         =   "Log"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   120
      OLEDropMode     =   1  'Manual
      TabIndex        =   17
      Top             =   1320
      Visible         =   0   'False
      Width           =   7095
      Begin VB.CommandButton cmdLogBack 
         Caption         =   "Back"
         Height          =   375
         Left            =   5880
         OLEDropMode     =   1  'Manual
         TabIndex        =   21
         Top             =   1200
         Width           =   1095
      End
      Begin VB.CommandButton cmdLogCopy 
         Caption         =   "Copy"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5880
         OLEDropMode     =   1  'Manual
         TabIndex        =   20
         Top             =   720
         Width           =   1095
      End
      Begin VB.CommandButton cmdLogSave 
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5880
         OLEDropMode     =   1  'Manual
         TabIndex        =   19
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtLog 
         Height          =   1695
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         OLEDropMode     =   1  'Manual
         ScrollBars      =   3  'Both
         TabIndex        =   18
         Top             =   240
         Width           =   5655
      End
   End
   Begin VB.Label lblVersion 
      AutoSize        =   -1  'True
      Caption         =   "BFU - Brute Force Uninstaller"
      Height          =   195
      Left            =   720
      OLEDropMode     =   1  'Manual
      TabIndex        =   10
      Top             =   3600
      Width           =   2070
   End
   Begin VB.Image imgLogo 
      Height          =   480
      Left            =   120
      OLEDropMode     =   1  'Manual
      Picture         =   "frmMain.frx":327C
      Top             =   3600
      Width           =   480
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "Scriptfile to execute:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   120
      OLEDropMode     =   1  'Manual
      TabIndex        =   8
      Top             =   600
      Width           =   1755
   End
   Begin VB.Label lblHeader 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H80000002&
      Caption         =   "The Brute Force Uninstaller"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   525
      Left            =   0
      OLEDropMode     =   1  'Manual
      TabIndex        =   7
      Top             =   0
      Width           =   7410
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'TODO:
'--- 1.11
'V RegSetExpandValue
'V RegDelKeyIfNameContainsText/Hex
'V execute from URL from commandline
'V fixed RegResetPermissions (and others) not being recognized,
'  by fixing horrible LoadCommandsList sub
'V fixed error in manual about RegDelValueIfNameContains..
'V FileSetAttributes op hosts file file not found na 1e keer
'V OptionSaveLog filename.log (aliases supported)
'V OptionShowLog unknown command
'V more extensive logging
'V fixed bug in CRC32 checksum procedure with leading zeroes
'? GetMatchingFiles subfolders use in Cmd optional maken?
'? runonreboot werkt niet.. path tussen quotes?
'X modCRC32 updaten om niet zo langzaam files te lezen
'--- 1.10
'V manual
'V lines with unexpanded env vars are skipped
'V Savvas email
'V Pieter/Mark email
'V wildcard support for andere Reg functions
'V RegSetMultiValue (nog testen) met \0
'V Vervelende sCmd stringmanipulaties omzetten in Split()/Select Case
'V FolderDelete fail -> FolderClear, delete on reboot
'V OptionSetStatus weghalen
'V maken: RegDelValueIfNameContains[Text/Hex]
'V RegDelValueIfDataContains[Text/Hex]
'V OptionShowLog
'V fixed bug: CRC32 was reversed
'V niet silentrun doen bij script als parameter
'V LogIfFileMD5Match / SHA1, etc
'V file mask search zoekt ook in subfolders (nu met alle wildcards!)
'V FolderClear
'V fixed WinXP 64b <-> Win2003 SBS
'V RegKeyResetPermissions
'V LogIfFileExist, LogIfRegKeyExist, LogIfRegValExist, (nog testen)
'V OptionBFUExit
'V %Favorites% alias

Option Explicit
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Sub ReleaseCapture Lib "user32" ()

Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2

Private sScript$

Private Sub chkShowLog_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Form_OLEDragDrop Data, Effect, Button, Shift, X, Y
End Sub

Private Sub chkUnloadShell_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Form_OLEDragDrop Data, Effect, Button, Shift, X, Y
End Sub

Private Sub chkUseRecycleBin_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Form_OLEDragDrop Data, Effect, Button, Shift, X, Y
End Sub

Private Sub chkUseScriptOptions_Click()
    If chkUseScriptOptions.Value = 1 Then
        chkUnloadShell.Enabled = False
        chkUseRecycleBin.Enabled = False
        GetScriptOptions sScript
    Else
        chkUnloadShell.Enabled = True
        chkUseRecycleBin.Enabled = True
    End If
End Sub

Private Sub chkUseScriptOptions_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Form_OLEDragDrop Data, Effect, Button, Shift, X, Y
End Sub

Private Sub cmdAbortScript_Click()
    bAbortScript = True
End Sub

Private Sub cmdAbortScript_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Form_OLEDragDrop Data, Effect, Button, Shift, X, Y
End Sub

Private Sub cmdExecute_Click()
    If txtScript.Text = vbNullString Then Exit Sub
    cmdExecute.Enabled = False
    bAbortScript = False
    cmdAbortScript.Enabled = True
    
    ExecuteScript sScript
    
    cmdAbortScript.Enabled = False
    cmdExecute.Enabled = True
    If chkShowLog.Value = 1 Then
        fraOptions.Visible = False
        fraLog.Visible = True
        txtLog.Text = sLog
    End If
    If bRunSilent Then End
End Sub

Private Sub cmdExecute_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Form_OLEDragDrop Data, Effect, Button, Shift, X, Y
End Sub

Private Sub cmdExit_Click()
    Close
    End
End Sub

Private Sub cmdExit_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Form_OLEDragDrop Data, Effect, Button, Shift, X, Y
End Sub

Private Sub cmdGetScript_Click()
    Dim sFile$
    sFile = CmnDialogGetFilename("BFU script files (*.bfu)|*.bfu|All files (*.*)|*.*", "Open BFU script...")
    If sFile = vbNullString Then Exit Sub
    
    GetScript sFile
End Sub

Private Sub cmdGetScript_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Form_OLEDragDrop Data, Effect, Button, Shift, X, Y
End Sub

Private Sub cmdGetURL_Click()
    Dim sMsg$, sURL$, sFile$
    sMsg = "Please enter the full URL to the script you want to download:"
    sURL = InputBox(sMsg, "Download BFU script...")
    If sURL = vbNullString Then Exit Sub
    sScript = InputURL(sURL)
    If sScript = vbNullString Then
        MsgBox "BFU was unable to download the file located at:" & vbCrLf & _
               sURL & vbCrLf & vbCrLf & "Please verify the address " & _
               "is correct and the file is available from the webserver.", vbExclamation
        Exit Sub
    End If
    txtScript.Text = sURL
    If InStr(1, sURL, ".bfu", vbTextCompare) > 0 Then
        Dim sCRC32$
        sFile = App.Path & "\" & Mid(sURL, InStrRev(sURL, "/") + 1)
        OutputFile sFile, sScript
        sCRC32 = GetScriptCRC32(sFile)
        If sCRC32 <> vbNullString Then
            lblInfo(0).Caption = "Script to execute (CRC32 " & sCRC32 & "):"
        Else
            lblInfo(0).Caption = "Script to execute:"
        End If
    End If
    cmdExecute.Enabled = True
End Sub

Private Sub cmdGetURL_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Form_OLEDragDrop Data, Effect, Button, Shift, X, Y
End Sub

Private Sub cmdLogBack_Click()
    fraLog.Visible = False
    fraOptions.Visible = True
End Sub

Private Sub cmdLogBack_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Form_OLEDragDrop Data, Effect, Button, Shift, X, Y
End Sub

Private Sub cmdLogCopy_Click()
    Clipboard.Clear
    Clipboard.SetText txtLog.Text
    MsgBox "Text copied to clipboard.", vbInformation
End Sub

Private Sub cmdLogCopy_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Form_OLEDragDrop Data, Effect, Button, Shift, X, Y
End Sub

Private Sub cmdLogSave_Click()
    Dim sFile$
    sFile = CmnDialogSaveFilename("Text files (*.txt)|*.txt|All files (*.*)|*.*", "Save log...")
    If sFile <> vbNullString Then
        On Error Resume Next
        Open sFile For Output As #1
            Print #1, txtLog.Text
        Close #1
    End If
End Sub

Private Sub cmdLogSave_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Form_OLEDragDrop Data, Effect, Button, Shift, X, Y
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Dim sFile$, sCRC32$
    lblVersion.Caption = "BFU - Brute Force Uninstaller" & vbCrLf & _
                          "Version " & App.Major & "." & _
                          Format(App.Minor, "00") & "." & App.Revision & vbCrLf & _
                          "Written by Merijn" & vbCrLf & "http://www.merijn.org/"
                          
    GetWindowsInfo
    LoadCommandsList
    shpProgress.Width = 15
    modCRC32.Init
            
    If Command$ <> vbNullString Then
        sFile = Command$
        If Left(sFile, 1) = """" Then
            'stupid Windows - adding quotes
            sFile = Mid(sFile, 2, Len(sFile) - 2)
        End If
        If LCase(Right(sFile, 4)) = ".bfu" Then
            If InStr(1, sFile, "http://", vbTextCompare) > 0 Then
                'it's an url
                sScript = InputURL(sFile)
                If sScript = vbNullString Then
                    MsgBox "BFU was unable to download the file located at:" & vbCrLf & _
                           sFile & vbCrLf & vbCrLf & "Please verify the address " & _
                           "is correct and the file is available from the webserver.", vbExclamation
                    Exit Sub
                End If
                txtScript.Text = sFile
                If InStr(1, sFile, ".bfu", vbTextCompare) > 0 Then
                    sFile = App.Path & "\" & Mid(sFile, InStrRev(sFile, "/") + 1)
                    OutputFile sFile, sScript
                    sCRC32 = GetScriptCRC32(sFile)
                    If sCRC32 <> vbNullString Then
                        lblInfo(0).Caption = "Script to execute (CRC32 " & sCRC32 & "):"
                    Else
                        lblInfo(0).Caption = "Script to execute:"
                    End If
                End If
                cmdExecute.Enabled = True
            ElseIf InStr(sFile, "\") = 0 Then sFile = BuildPath(App.Path, sFile)
                If FileExists(sFile) Then
                    txtScript.Text = sFile
                    sScript = vbNullString
                    sScript = InputFile(sFile)
                    GetScriptOptions sScript
                    sCRC32 = GetScriptCRC32(sFile)
                    If sCRC32 <> vbNullString Then
                        lblInfo(0).Caption = "Script to execute (CRC32 " & sCRC32 & "):"
                    Else
                        lblInfo(0).Caption = "Script to execute:"
                    End If
                    Me.Show
                    cmdExecute.Enabled = True
                    cmdExecute_Click
                End If
            End If
        End If
    End If
End Sub

Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    'user dragged a file onto BFU
    Dim sFile$
    On Error Resume Next
    sFile = Data.Files(1)
    If Err Then Exit Sub
    If InStr(1, sFile, ".bfu", vbTextCompare) = 0 Then Exit Sub
    Err.Clear
    GetScript sFile
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Me.WindowState = vbMinimized Then Exit Sub
    
    If Me.ScaleWidth >= 7290 Then
        lblHeader.Width = Me.ScaleWidth
        txtScript.Width = Me.ScaleWidth - 1455
        cmdGetScript.Left = Me.ScaleWidth - 1050
        cmdGetURL.Left = Me.ScaleWidth - 540
        fraOptions.Width = Me.ScaleWidth - 195
        fraProgress.Width = Me.ScaleWidth - 195
        fraLog.Width = Me.ScaleWidth - 195
        linSeperator(0).X2 = Me.ScaleWidth - 450
        linSeperator(1).X2 = Me.ScaleWidth - 450
        cmdAbortScript.Left = Me.ScaleWidth - 3810
        cmdExecute.Left = Me.ScaleWidth - 2490
        cmdExit.Left = Me.ScaleWidth - 1170
        shpProgressBackgrond.Width = Me.ScaleWidth - 675
        txtLog.Width = Me.ScaleWidth - 1635
        cmdLogSave.Left = Me.ScaleWidth - 1410
        cmdLogCopy.Left = Me.ScaleWidth - 1410
        cmdLogBack.Left = Me.ScaleWidth - 1410
    End If
    
    If Me.ScaleHeight >= 4575 Then
        fraOptions.Height = Me.ScaleHeight - 2520
        fraProgress.Height = Me.ScaleHeight - 2520
        fraLog.Height = Me.ScaleHeight - 2520
        lblVersion.Top = Me.ScaleHeight - 975
        imgLogo.Top = Me.ScaleHeight - 975
        cmdAbortScript.Top = Me.ScaleHeight - 1095
        cmdExecute.Top = Me.ScaleHeight - 1095
        cmdExit.Top = Me.ScaleHeight - 1095
        txtLog.Height = Me.ScaleHeight - 2880
    End If
End Sub

Private Sub fraLog_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Form_OLEDragDrop Data, Effect, Button, Shift, X, Y
End Sub

Private Sub fraOptions_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Form_OLEDragDrop Data, Effect, Button, Shift, X, Y
End Sub

Private Sub imgLogo_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Form_OLEDragDrop Data, Effect, Button, Shift, X, Y
End Sub

Private Sub lblHeader_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        ReleaseCapture
        SendMessage Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0
    End If
End Sub

Private Sub GetScript(sFile$)
    Dim sCRC32$
    sScript = vbNullString
    txtScript.Text = sFile
    sScript = InputFile(sFile)
    'Open sFile For Binary As #1
    '    sScript = Input(FileLen(sFile), #1)
    'Close #1
    GetScriptOptions sScript
    sCRC32 = GetScriptCRC32(sFile)
    If sCRC32 <> vbNullString Then
        lblInfo(0).Caption = "Script to execute (CRC32 " & sCRC32 & "):"
    Else
        lblInfo(0).Caption = "Script to execute:"
    End If
    cmdExecute.Enabled = True
End Sub

Private Sub lblHeader_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Form_OLEDragDrop Data, Effect, Button, Shift, X, Y
End Sub

Private Sub lblInfo_OLEDragDrop(Index As Integer, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Form_OLEDragDrop Data, Effect, Button, Shift, X, Y
End Sub

Private Sub lblVersion_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Form_OLEDragDrop Data, Effect, Button, Shift, X, Y
End Sub

Private Sub txtLog_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Form_OLEDragDrop Data, Effect, Button, Shift, X, Y
End Sub

Private Sub txtScript_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Form_OLEDragDrop Data, Effect, Button, Shift, X, Y
End Sub
