VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAddIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Based on:
'http://msdn.microsoft.com/library/default.asp?URL=/library/techart/trcomad.htm
'COM Add-ins Part II: Building a COM Add-in for Outlook 2000

Implements IDTExtensibility2

Private WithEvents moApplication As Application
Attribute moApplication.VB_VarHelpID = -1
Private WithEvents moInspectors As Inspectors
Attribute moInspectors.VB_VarHelpID = -1
Private WithEvents moMailItem As MailItem
Attribute moMailItem.VB_VarHelpID = -1
Private WithEvents moInspector As Inspector
Attribute moInspector.VB_VarHelpID = -1

Private WithEvents oButtonReadMail As Office.CommandBarButton
Attribute oButtonReadMail.VB_VarHelpID = -1
Private WithEvents oButtonStopReadMail As Office.CommandBarButton
Attribute oButtonStopReadMail.VB_VarHelpID = -1
Private WithEvents oButtonPauseReadMail As Office.CommandBarButton
Attribute oButtonPauseReadMail.VB_VarHelpID = -1
Private WithEvents oButtonReasumeReadMail As Office.CommandBarButton
Attribute oButtonReasumeReadMail.VB_VarHelpID = -1
Private WithEvents oButtonForwardReadMail As Office.CommandBarButton
Attribute oButtonForwardReadMail.VB_VarHelpID = -1

Private moVoiceSynth As TextToSpeech
Private msLastMailItemUsingTTS As String

'These constants are also used by OutlookQuoteProperties.ocx and OutlookQuote.dll projects
Private Const REG_APP_NAME As String = "OutlookQuotes"
Private Const REG_SECTION_SETTINGS As String = "Settings"
Private Const REG_KEY_AFTER_QUOTE_STRING As String = "AfterQuoteString"
Private Const REG_KEY_BEFORE_QUOTE_STRING As String = "BeforeQuoteString"
Private Const REG_KEY_DELIMITER_QUOTE_AUTHOR As String = "DelimiterQuoteAuthor"
Private Const REG_KEY_DELIMITER_QUOTES As String = "DelimiterQuotes"
Private Const REG_KEY_QUOTE_FILE As String = "QuoteFile"
Private Const DEFAULT_BEFORE_INSERTE_STRING As String = vbCrLf & vbCrLf & "Regards," & vbCrLf & "(my name). <--- to change this, click on Outlook's main window (it's behind this email window), then click on Tools > Options menu and see the Quotes tab" & vbCrLf & vbCrLf & "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~" & vbCrLf
Private Const DEFAULT_AFTER_INSERTE_STRING As String = vbCrLf & "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~"

Private mbDontInsertQuote As Boolean

Private Sub moApplication_OptionsPagesAdd(ByVal Pages As Outlook.PropertyPages)
    Dim oQuoteSettings As Outlook.PropertyPage
    Set oQuoteSettings = CreateObject("OutlookQuoteProperties.SettingsPage")
    Call Pages.Add(oQuoteSettings, "Quotes")
End Sub

Private Sub moApplication_Quit()
    Call Cleanup
End Sub

Private Sub moApplication_Startup()
    Set moInspectors = moApplication.Inspectors
End Sub

Private Sub moInspector_Close()
    Set oButtonReadMail = Nothing
    Set moInspector = Nothing
End Sub

Private Sub moInspectors_NewInspector(ByVal Inspector As Outlook.Inspector)
    If TypeName(Inspector.CurrentItem) = "MailItem" Then
        Set moMailItem = Inspector.CurrentItem
        Set moInspector = Inspector
        Call AddMenu(Inspector.CommandBars)
    Else
        'No processing for other items
    End If
End Sub

Private Sub Cleanup()
    Set moApplication = Nothing
    Set moInspectors = Nothing
    Set moMailItem = Nothing
    Set oButtonReadMail = Nothing
    Set moVoiceSynth = Nothing
End Sub


Private Sub moMailItem_Close(Cancel As Boolean)
    If Not IsMailItemMovedOrDeleted Then
    If msLastMailItemUsingTTS = moMailItem.EntryID Then
        Call StopTTS
    End If
    End If
    Set moMailItem = Nothing
End Sub

Private Function IsMailItemMovedOrDeleted() As Boolean
    On Error Resume Next
    Dim sDummy As String
    sDummy = moMailItem.EntryID
    IsMailItemMovedOrDeleted = (Err.Number <> 0)
End Function


Private Sub moMailItem_Open(Cancel As Boolean)
    On Error GoTo ErrorTrap
    
    If mbDontInsertQuote = False Then
        If moMailItem.EntryID = vbNullString Then
            Dim oNameSpace As NameSpace
            Set oNameSpace = moMailItem.Application.GetNamespace("MAPI")
            Dim oNotesFolder As MAPIFolder
            Set oNotesFolder = oNameSpace.GetDefaultFolder(olFolderNotes)
            Dim oQuotesFolder As MAPIFolder
            Set oQuotesFolder = GetSubFolder(oNotesFolder, "Quotes")
            
            If oQuotesFolder Is Nothing Then
                Call frmFileImport.DisplayForm(True, GetPathWithSlash(App.Path) & "basic_quotes.txt", vbCrLf, "|")
                Set oQuotesFolder = GetSubFolder(oNotesFolder, "Quotes")
                If oQuotesFolder Is Nothing Then
                    Err.Raise 1000, , "Quotes folder in Outlook Notes item is not created. Try to import new Quotes from Tools > Options menu and Quotes tab."
                End If
            End If
            
            If oQuotesFolder.Items.Count > 0 Then
                Randomize
                Dim oQuoteNote As NoteItem
                Set oQuoteNote = oQuotesFolder.Items.Item(Int((oQuotesFolder.Items.Count * Rnd) + 1))
                
                Dim sSignature As String
                
                Dim sStringToInsertBeforeQuote As String
                Dim sStringToInsertAfterQuote As String
                sStringToInsertBeforeQuote = GetSetting(REG_APP_NAME, REG_SECTION_SETTINGS, REG_KEY_BEFORE_QUOTE_STRING, DEFAULT_BEFORE_INSERTE_STRING)
                sStringToInsertAfterQuote = GetSetting(REG_APP_NAME, REG_SECTION_SETTINGS, REG_KEY_AFTER_QUOTE_STRING, DEFAULT_AFTER_INSERTE_STRING)
                

                sSignature = sStringToInsertBeforeQuote & oQuoteNote.Body & sStringToInsertAfterQuote
            
                Select Case moMailItem.GetInspector.EditorType
                    Case olEditorText
                        moMailItem.Body = sSignature & moMailItem.Body
                    Case olEditorHTML
                        Dim sHTMLBody As String
                        sHTMLBody = moMailItem.HTMLBody
                        Call Replace(sHTMLBody, "<BODY>", "<BODY>" & vbCrLf & sSignature & vbCrLf, , vbTextCompare)
                        moMailItem.HTMLBody = sHTMLBody
                    Case Else
                        'moMailItem.Body = "This is body type " & moMailItem.GetInspector.EditorType & vbCrLf & moMailItem.Body
                        moMailItem.Body = sSignature & moMailItem.Body
                End Select
            End If
        Else
            'This is not new mail
        End If
    Else
        'Don't insert quote
    End If
    
Exit Sub
ErrorTrap:
    If Not mbDontInsertQuote Then
        mbDontInsertQuote = True
        MsgBox "Quote will not be inserted: " & Err.DESCRIPTION
    End If
End Sub


Private Sub IDTExtensibility2_OnAddInsUpdate(custom() As Variant)

End Sub

Private Sub IDTExtensibility2_OnBeginShutdown(custom() As Variant)
    Call Cleanup
End Sub

Private Sub IDTExtensibility2_OnConnection(ByVal Application As Object, ByVal ConnectMode As AddInDesignerObjects.ext_ConnectMode, ByVal AddInInst As Object, custom() As Variant)
    Set moApplication = Application
End Sub

Private Sub IDTExtensibility2_OnDisconnection(ByVal RemoveMode As AddInDesignerObjects.ext_DisconnectMode, custom() As Variant)
    Call Cleanup
End Sub

Private Sub IDTExtensibility2_OnStartupComplete(custom() As Variant)

End Sub

Private Sub AddMenu(ByVal voCommandBars As CommandBars)
    Dim oMenuBar As Office.CommandBar
    Set oMenuBar = voCommandBars.Add("OutlookWizAddIn", msoBarTop, False, True)
    oMenuBar.Visible = True
    
    Set oButtonReadMail = oMenuBar.Controls.Add(msoControlButton, , , , True)
    oButtonReadMail.Caption = "Read2Me"
    oButtonReadMail.Enabled = True
    
    Set oButtonStopReadMail = oMenuBar.Controls.Add(msoControlButton, , , , True)
    oButtonStopReadMail.Caption = "Stop Reading"
    oButtonStopReadMail.Enabled = True
    
    Set oButtonPauseReadMail = oMenuBar.Controls.Add(msoControlButton, , , , True)
    oButtonPauseReadMail.Caption = "Pause Reading"
    oButtonPauseReadMail.Enabled = True
    
    Set oButtonReasumeReadMail = oMenuBar.Controls.Add(msoControlButton, , , , True)
    oButtonReasumeReadMail.Caption = "Resume Reading"
    oButtonReasumeReadMail.Enabled = True
    
    Set oButtonForwardReadMail = oMenuBar.Controls.Add(msoControlButton, , , , True)
    oButtonForwardReadMail.Caption = "Skip >>"
    oButtonForwardReadMail.Enabled = True
    
    Set oMenuBar = Nothing
End Sub

Private Sub oButtonForwardReadMail_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)
    Call ForwardTTS
End Sub

Private Sub oButtonPauseReadMail_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)
    Call PauseTTS
End Sub

Private Sub oButtonReadMail_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)
    On Error GoTo ErrorTrap
    
    Call InitTTS
    
    Dim sTextToRead As String
    sTextToRead = moMailItem.Body
    If sTextToRead <> vbNullString Then
        Call MakeTextReadable(sTextToRead)
        Call moVoiceSynth.Speak(sTextToRead)
    Else
        'Nothing to do
    End If
    msLastMailItemUsingTTS = moMailItem.EntryID
    
Exit Sub
ErrorTrap:
    MsgBox Err.DESCRIPTION
End Sub

Private Sub MakeTextReadable(ByRef rstrText As String)
    rstrText = Left(rstrText, 3000)
    rstrText = Replace(rstrText, "http", "URL ", vbTextCompare)
    rstrText = Replace(rstrText, "://", "", vbTextCompare)
    rstrText = Replace(rstrText, "...", ". ", vbTextCompare)
    rstrText = Replace(rstrText, "www", "", vbTextCompare)
    rstrText = Replace(rstrText, ".com/", ".com", vbTextCompare)
    'rstrText = Replace(rstrText, " - ", " to ", vbTextCompare)
End Sub

Private Sub RemoveUnspeakbleChars(ByRef rstrText As String)
    Dim sNewSpeakableText As String
    Dim sThisChar As String
    Dim sNextChar As String
    Dim lCharIndex As Long
    For lCharIndex = 1 To Len(rstrText) - 1
        sThisChar = Mid(rstrText, lCharIndex, 1)
        sNextChar = Mid(rstrText, lCharIndex + 1, 1)
        If IsCharSpeakable(sThisChar) = True Then
            sNewSpeakableText = sNewSpeakableText & sThisChar
        Else
            If IsCharSpeakable(Mid(rstrText, lCharIndex + 1, 1)) = True Then
                sNewSpeakableText = sNewSpeakableText & sThisChar
            Else
                'Skip
            End If
        End If
    Next
    sNewSpeakableText = sNewSpeakableText & Right(rstrText, 1)
    rstrText = sNewSpeakableText
End Sub

Private Function IsCharSpeakable(ByVal vsChar As String) As Boolean
    Dim sLCasedChar As String
    sLCasedChar = LCase(vsChar)
    If ((sLCasedChar >= "a") And (sLCasedChar <= "z")) Or ((sLCasedChar >= "0") And (sLCasedChar <= "9")) Or (sLCasedChar = " ") Then
        IsCharSpeakable = True
    Else
        IsCharSpeakable = False
    End If
End Function

Private Sub InitTTS()
    If moVoiceSynth Is Nothing Then
        Set moVoiceSynth = New HTTSLib.TextToSpeech
    End If
    Call StopTTS
End Sub

Private Sub StopTTS()
    If Not (moVoiceSynth Is Nothing) Then
        If moVoiceSynth.IsSpeaking Then
            moVoiceSynth.StopSpeaking
        End If
    End If
End Sub

Private Sub PauseTTS()
    If Not (moVoiceSynth Is Nothing) Then
        If moVoiceSynth.IsSpeaking Then
            moVoiceSynth.Pause
        End If
    End If
End Sub

Private Sub ResumeTTS()
    If Not (moVoiceSynth Is Nothing) Then
        If Not moVoiceSynth.IsSpeaking Then
            moVoiceSynth.Resume
        End If
    End If
End Sub
Private Sub ForwardTTS()
    If Not (moVoiceSynth Is Nothing) Then
        If Not moVoiceSynth.IsSpeaking Then
            moVoiceSynth.FastForward
        End If
    End If
End Sub
Private Sub oButtonReasumeReadMail_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)
    Call ResumeTTS
End Sub

Private Sub oButtonStopReadMail_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)
    Call StopTTS
End Sub

Public Sub ImportQuotes(ByVal IsUIInvisible As Boolean, ByVal QuotesFileName As String, ByVal DelimiterBetweenQuotes As String, ByVal DelimiterBetweenQuoteAndAuthor As String)
    Call frmFileImport.DisplayForm(IsUIInvisible, QuotesFileName, DelimiterBetweenQuotes, DelimiterBetweenQuoteAndAuthor)
End Sub

Public Sub RegisterAddIn()
    Call RegisterOfficeAddin("Outlook")
End Sub

Public Sub UnregisterAddIn()
    Call UnRegisterOfficeAddin("Outlook")
End Sub
