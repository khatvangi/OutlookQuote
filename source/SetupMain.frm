VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Import Quotes"
   ClientHeight    =   3000
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "SetupMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3000
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   90
      Left            =   0
      TabIndex        =   8
      Top             =   2250
      Width           =   4740
   End
   Begin VB.TextBox txtQuotesFilePath 
      BackColor       =   &H00C0C0C0&
      Height          =   315
      Left            =   1125
      TabIndex        =   7
      Top             =   1800
      Width           =   3465
   End
   Begin VB.CommandButton cmdImport 
      Caption         =   "Import Now!"
      Default         =   -1  'True
      Height          =   390
      Left            =   2025
      TabIndex        =   0
      Top             =   2475
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   390
      Left            =   3375
      TabIndex        =   2
      Top             =   2475
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Import Options:"
      Height          =   1590
      Left            =   75
      TabIndex        =   1
      Top             =   75
      Width           =   4515
      Begin VB.OptionButton optSomeQuotes 
         Caption         =   "Get me some quotes available (few seconds)"
         Height          =   240
         Left            =   150
         TabIndex        =   5
         Top             =   750
         Width           =   4215
      End
      Begin VB.OptionButton optAllQuotes 
         Caption         =   "Get me all the quotes available (takes a minute)"
         Height          =   240
         Left            =   150
         TabIndex        =   4
         Top             =   375
         Value           =   -1  'True
         Width           =   4215
      End
      Begin VB.OptionButton optNoneQuotes 
         Caption         =   "Don't get me any quotes"
         Height          =   240
         Left            =   150
         TabIndex        =   3
         Top             =   1125
         Width           =   4215
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Quotes Path:"
      Height          =   195
      Left            =   75
      TabIndex        =   6
      Top             =   1800
      Width           =   930
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const REG_APP_NAME As String = "OutlookQuotes"
Private Const REG_SECTION_SETTINGS As String = "Settings"

Private Sub cmdCancel_Click()
    Unload Me
    End
End Sub

Private Sub cmdImport_Click()

    On Error GoTo ErrorTrap

    Screen.MousePointer = vbHourglass
    
    Dim sInstallDir As String
    sInstallDir = txtQuotesFilePath.Text
    
    Dim oOutlookQuote As OutlookQuote.clsAddIn
    Set oOutlookQuote = New clsAddIn
        
    If optAllQuotes.Value = True Then
        Call oOutlookQuote.ImportQuotes(True, GetPathWithSlash(sInstallDir) & "basic_quotes.txt", vbCrLf, "|")
        Call oOutlookQuote.ImportQuotes(True, GetPathWithSlash(sInstallDir) & "more_quotes.txt", vbCrLf, "|")
    ElseIf optSomeQuotes.Value = True Then
        Call oOutlookQuote.ImportQuotes(True, GetPathWithSlash(sInstallDir) & "basic_quotes.txt", vbCrLf, "|")
    Else
        'No imports
    End If

Cleanup:
    Unload Me
    End

Exit Sub
ErrorTrap:
    Screen.MousePointer = vbDefault
    
    MsgBox "Error " & Err.Number & " [" & Err.Source & "]:" & Err.Description
End Sub

Private Sub Form_Load()

    Dim sInstallDir As String
    sInstallDir = GetSetting(REG_APP_NAME, REG_SECTION_SETTINGS, "InstallPath", vbNullString)

    If sInstallDir = vbNullString Then
        'Use current App path - in case manual setup by developers
        sInstallDir = App.Path
        'MsgBox "Warning: InstallPath in registry is not found. It should be probably in Program Files\Outlook Quote folder"
    End If
    txtQuotesFilePath.Text = sInstallDir
        
    Dim oOutlookQuote As OutlookQuote.clsAddIn
    Set oOutlookQuote = New clsAddIn
    If Command$ = vbNullString Then
        oOutlookQuote.RegisterAddIn
        Me.Show
    ElseIf InStr(1, Command$, "reg", vbTextCompare) = 1 Then
        oOutlookQuote.RegisterAddIn
        Me.Show
    ElseIf InStr(1, Command$, "un", vbTextCompare) = 1 Then
        oOutlookQuote.UnregisterAddIn
        MsgBox "If you had imported any quotes, please manually delete them from your Outlook's Notes folder.", , "Reminder"
        Unload Me
        End
    Else
        MsgBox "The command line parameter " & Command$ & " is not recognized"
        Unload Me
        End
    End If
End Sub
