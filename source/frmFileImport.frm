VERSION 5.00
Begin VB.Form frmFileImport 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Import Quotes From File"
   ClientHeight    =   2940
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3600
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2940
   ScaleWidth      =   3600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   390
      Left            =   2250
      TabIndex        =   8
      Top             =   2475
      Width           =   1215
   End
   Begin VB.CommandButton cmdImport 
      Caption         =   "Start Import"
      Default         =   -1  'True
      Height          =   390
      Left            =   900
      TabIndex        =   7
      Top             =   2475
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   90
      Left            =   75
      TabIndex        =   6
      Top             =   2250
      Width           =   3540
   End
   Begin VB.TextBox txtDelimiterQuoteAuthor 
      Height          =   285
      Left            =   1125
      TabIndex        =   5
      Top             =   1200
      Width           =   2340
   End
   Begin VB.TextBox txtDelimiterQuotes 
      Height          =   285
      Left            =   1125
      TabIndex        =   3
      Top             =   600
      Width           =   2340
   End
   Begin VB.TextBox txtQuotesFile 
      Height          =   285
      Left            =   1125
      TabIndex        =   1
      Top             =   150
      Width           =   2340
   End
   Begin VB.Label lblProgress 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "                      "
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   1125
      TabIndex        =   10
      Top             =   1875
      Width           =   2340
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Progress:"
      Height          =   195
      Left            =   150
      TabIndex        =   9
      Top             =   1875
      Width           =   660
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Quote-Author Delimiter:"
      Height          =   390
      Left            =   150
      TabIndex        =   4
      Top             =   1275
      Width           =   990
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Quotes Delimiter:"
      Height          =   390
      Left            =   150
      TabIndex        =   2
      Top             =   600
      Width           =   825
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "File:"
      Height          =   195
      Left            =   150
      TabIndex        =   0
      Top             =   150
      Width           =   360
   End
End
Attribute VB_Name = "frmFileImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mbStopImporting As Boolean

Public Sub DisplayForm(ByVal IsInvisible As Boolean, ByVal QuotesFileName As String, ByVal DelimiterBetweenQuotes As String, ByVal DelimiterBetweenQuoteAndAuthor As String)

    On Error GoTo ErrorTrap

    txtQuotesFile = QuotesFileName
    txtDelimiterQuotes = DelimiterBetweenQuotes
    txtDelimiterQuoteAuthor = DelimiterBetweenQuoteAndAuthor
    
    If IsInvisible = False Then
        Me.Show vbModal
        
        QuotesFileName = txtQuotesFile
        DelimiterBetweenQuotes = txtDelimiterQuotes
        DelimiterBetweenQuoteAndAuthor = txtDelimiterQuoteAuthor
    Else
        Me.Show
        cmdImport.Visible = False
        txtQuotesFile.Enabled = False
        txtDelimiterQuoteAuthor.Enabled = False
        txtDelimiterQuotes.Enabled = False
    End If
    
    If mbStopImporting = False Then
        Call ImportQuotes(QuotesFileName, DelimiterBetweenQuotes, DelimiterBetweenQuoteAndAuthor)
    End If
    
    Unload Me
        
Exit Sub
ErrorTrap:
    MsgBox "Error occured while importing " & Err.Number & ": " & Err.DESCRIPTION & " - Source: " & Err.Source
End Sub

Private Sub ImportQuotes(ByVal QuotesFileName As String, ByVal DelimiterBetweenQuotes As String, ByVal DelimiterBetweenQuoteAndAuthor As String)
    
    mbStopImporting = False
    
    Dim oOutlookApplication As Outlook.Application
    Set oOutlookApplication = CreateObject("Outlook.Application")
    
    Dim oNameSpace As NameSpace
    Set oNameSpace = oOutlookApplication.GetNamespace("MAPI")
    Dim oNotesFolder As MAPIFolder
    Dim oQuotesFolder As MAPIFolder
    Set oNotesFolder = oNameSpace.GetDefaultFolder(olFolderNotes)
    Set oQuotesFolder = GetSubFolder(oNotesFolder, "Quotes")
    If oQuotesFolder Is Nothing Then
        Set oQuotesFolder = oNotesFolder.Folders.Add("Quotes")
    End If
    
    Dim sQuotesFileContent As String
    Dim aQuotes As Variant
    Dim aQuoteDetail As Variant
    
    sQuotesFileContent = LoadStringFromFile(QuotesFileName)
    aQuotes = Split(sQuotesFileContent, DelimiterBetweenQuotes)
    
    Dim lQuoteIndex As Long
    Dim oQuoteItem As NoteItem
    Dim lTotalQuotes As Long
    lTotalQuotes = UBound(aQuotes) - LBound(aQuotes) + 1
    For lQuoteIndex = LBound(aQuotes) To UBound(aQuotes)
        If mbStopImporting = True Then
            Exit For
        End If
        aQuoteDetail = Split(aQuotes(lQuoteIndex), DelimiterBetweenQuoteAndAuthor)
        
        Select Case (UBound(aQuoteDetail) - LBound(aQuoteDetail))
            Case 0
                Set oQuoteItem = oQuotesFolder.Items.Add
                oQuoteItem.Body = aQuoteDetail(LBound(aQuoteDetail))
                oQuoteItem.Save
            Case Is >= 1
                Set oQuoteItem = oQuotesFolder.Items.Add
                Dim sAuthorName As String
                sAuthorName = LTrim(aQuoteDetail(UBound(aQuoteDetail)))
                If sAuthorName <> vbNullString Then
                    If Left(sAuthorName, 1) <> "-" Then
                        sAuthorName = "-" & sAuthorName
                    End If
                    sAuthorName = vbCrLf & vbCrLf & sAuthorName
                End If
                oQuoteItem.Body = aQuoteDetail(LBound(aQuoteDetail)) & sAuthorName
                oQuoteItem.Save
            Case Is < 0
                'Don't add
        End Select
        Call UpdateProgress(lTotalQuotes, lQuoteIndex - LBound(aQuotes) + 1)
    Next
    
End Sub

Private Sub UpdateProgress(ByVal TotalQuotes As Double, ByVal QuotesImported As Double)
    Dim lProgress As Long
    If QuotesImported <> 0 Then
        lProgress = (QuotesImported * 22) / TotalQuotes '22 is size of lblProgress's Caption
    Else
        lProgress = 0
    End If
    lblProgress.Caption = String(lProgress, "*")
    DoEvents
End Sub

Private Sub cmdCancel_Click()
    Me.Hide
    mbStopImporting = True
End Sub

Private Sub cmdImport_Click()
    On Error GoTo ErrorTrap
    
    Call ImportQuotes(txtQuotesFile, txtDelimiterQuotes, txtDelimiterQuoteAuthor)
    
    Me.Hide
    Unload Me
    
Exit Sub
ErrorTrap:
    MsgBox "Error occured while importing " & Err.Number & ": " & Err.DESCRIPTION & " - Source: " & Err.Source
End Sub

