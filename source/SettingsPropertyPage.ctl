VERSION 5.00
Begin VB.UserControl SettingsPage 
   ClientHeight    =   5910
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5625
   ForeColor       =   &H00FF8080&
   ScaleHeight     =   5910
   ScaleWidth      =   5625
   Begin VB.Frame Frame3 
      Height          =   90
      Left            =   75
      TabIndex        =   21
      Top             =   3975
      Width           =   5415
   End
   Begin VB.CommandButton cmdImport 
      Caption         =   "Import More Quotes..."
      Height          =   315
      Left            =   3600
      TabIndex        =   19
      Top             =   3150
      Width           =   1815
   End
   Begin VB.Frame Frame2 
      Height          =   90
      Left            =   75
      TabIndex        =   18
      Top             =   2925
      Width           =   5415
   End
   Begin VB.Timer tmrAnimation 
      Enabled         =   0   'False
      Interval        =   30
      Left            =   3750
      Top             =   4575
   End
   Begin VB.Frame fraShadowTravel 
      BorderStyle     =   0  'None
      Height          =   405
      Left            =   1500
      TabIndex        =   14
      Top             =   4425
      Visible         =   0   'False
      Width           =   2685
      Begin VB.Line linSignatureAnimationLine2 
         BorderColor     =   &H00000000&
         BorderWidth     =   15
         DrawMode        =   6  'Mask Pen Not
         X1              =   -510
         X2              =   1950
         Y1              =   -360
         Y2              =   480
      End
      Begin VB.Label lblSignatureAnimationLabel1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Shital Shah"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000F&
         Height          =   555
         Left            =   -195
         TabIndex        =   15
         Top             =   -90
         Width           =   2910
      End
      Begin VB.Line linSignatureAnimationLine1 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   15
         DrawMode        =   6  'Mask Pen Not
         X1              =   -1740
         X2              =   720
         Y1              =   -90
         Y2              =   750
      End
      Begin VB.Label lblSignatureAnimationLabel2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Shital Shah"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   555
         Left            =   -165
         TabIndex        =   16
         Top             =   -60
         Width           =   2910
      End
   End
   Begin VB.TextBox txtQuotesFile 
      Height          =   315
      Left            =   1425
      TabIndex        =   12
      Text            =   "QUOTES.txt"
      Top             =   6675
      Visible         =   0   'False
      Width           =   2190
   End
   Begin VB.TextBox txtLineBeforeQuotes 
      Height          =   1365
      Left            =   1425
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   11
      Text            =   "SettingsPropertyPage.ctx":0000
      Top             =   150
      Width           =   3990
   End
   Begin VB.Frame Frame1 
      Caption         =   "Advanced"
      Height          =   2040
      Left            =   150
      TabIndex        =   3
      Top             =   6900
      Visible         =   0   'False
      Width           =   3465
      Begin VB.CheckBox chkUseLineBreakeAsQuoteAuthorDelimiter 
         Caption         =   "Use line break instead"
         Height          =   240
         Left            =   1425
         TabIndex        =   9
         Top             =   675
         Width           =   1890
      End
      Begin VB.CheckBox chkUseLineBreakeAsQuoteDelimiter 
         Caption         =   "Use line break instead"
         Height          =   240
         Left            =   1425
         TabIndex        =   8
         Top             =   1575
         Width           =   1890
      End
      Begin VB.TextBox txtDelimiterQuotes 
         Height          =   390
         Left            =   1425
         TabIndex        =   7
         Text            =   ","
         Top             =   1125
         Width           =   1965
      End
      Begin VB.TextBox txtDelimiterQuoteAndAuthor 
         Height          =   390
         Left            =   1425
         TabIndex        =   5
         Text            =   "|"
         Top             =   225
         Width           =   1965
      End
      Begin VB.Label Label3 
         Caption         =   "Delimiter between each quote:"
         Height          =   615
         Left            =   75
         TabIndex        =   6
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Delimiter between a quote and author:"
         Height          =   690
         Left            =   75
         TabIndex        =   4
         Top             =   300
         Width           =   1215
      End
   End
   Begin VB.TextBox txtLineAfterQuotes 
      Height          =   840
      Left            =   1425
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   2
      Text            =   "SettingsPropertyPage.ctx":00D0
      Top             =   1950
      Width           =   3990
   End
   Begin VB.Label Label10 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tip: You can add your own quotes too! Just go to Outlook's Notes' folder and look for Quotes!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   450
      Left            =   375
      TabIndex        =   24
      Top             =   3525
      Width           =   5025
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Dedicated to my old friend, Loneliness..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   180
      Left            =   2625
      TabIndex        =   23
      Top             =   5175
      Width           =   2490
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Designed, Developed And Coded By,"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   180
      Left            =   1500
      TabIndex        =   22
      Top             =   4125
      Width           =   2280
   End
   Begin VB.Label Label7 
      Caption         =   "Not Enogh Quotes? Click  Here ------>>>>"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   315
      Left            =   375
      TabIndex        =   20
      Top             =   3150
      Width           =   3165
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "http://www.ShitalShah.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   2
      Left            =   1875
      MouseIcon       =   "SettingsPropertyPage.ctx":0115
      MousePointer    =   99  'Custom
      TabIndex        =   17
      Top             =   4875
      Width           =   2070
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label6 
      Caption         =   "Tip: Use Ctrl+Enter for new line."
      Height          =   240
      Left            =   1425
      TabIndex        =   13
      Top             =   1650
      Width           =   5265
   End
   Begin VB.Label Label5 
      Caption         =   "Text to attach before quote:"
      Height          =   465
      Index           =   0
      Left            =   75
      TabIndex        =   10
      Top             =   225
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Text to attach after quote:"
      Height          =   465
      Left            =   75
      TabIndex        =   1
      Top             =   2025
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Where's your quotes located?"
      Height          =   465
      Left            =   75
      TabIndex        =   0
      Top             =   6825
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "SettingsPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements Outlook.PropertyPage

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

Dim mbIsDirty As Boolean
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const SW_SHOW = 5

Public Property Get PropertyPageCaption()
Attribute PropertyPageCaption.VB_UserMemId = -518
    PropertyPageCaption = "Quotes"
End Property


Private Sub cmdImport_Click()

    On Error GoTo ErrorTrap

    Dim oOutlookAddIn As Object
    Set oOutlookAddIn = CreateObject("OutlookQuote.clsAddIn")
    Call oOutlookAddIn.ImportQuotes(False, GetPathWithSlash(App.Path) & "more_quotes.txt", vbCrLf, "|")
    'Call oOutlookAddIn.ImportQuotes(True, "", "", "")
    
Exit Sub
ErrorTrap:
    MsgBox "Error occured while importing " & Err.Number & ": " & Err.Description & " - Source: " & Err.Source
End Sub

Private Sub PropertyPage_Apply()
    Call SaveSetting(REG_APP_NAME, REG_SECTION_SETTINGS, REG_KEY_AFTER_QUOTE_STRING, txtLineAfterQuotes.Text)
    Call SaveSetting(REG_APP_NAME, REG_SECTION_SETTINGS, REG_KEY_BEFORE_QUOTE_STRING, txtLineBeforeQuotes.Text)
'    Call SaveSetting(REG_APP_NAME, REG_SECTION_SETTINGS, REG_KEY_DELIMITER_QUOTE_AUTHOR, txtDelimiterQuoteAndAuthor)
'    Call SaveSetting(REG_APP_NAME, REG_SECTION_SETTINGS, REG_KEY_DELIMITER_QUOTES, txtDelimiterQuotes)
'    Call SaveSetting(REG_APP_NAME, REG_SECTION_SETTINGS, REG_KEY_QUOTE_FILE, txtQuotesFile)
    IsDirty = False
End Sub

Private Property Get PropertyPage_Dirty() As Boolean
    PropertyPage_Dirty = IsDirty
End Property

Private Sub PropertyPage_GetPageInfo(HelpFile As String, HelpContext As Long)
    'No help file exist
End Sub

Private Sub txtLineAfterQuotes_LostFocus()
    IsDirty = True
End Sub

Private Sub txtLineBeforeQuotes_Change()
    IsDirty = True
End Sub

Private Property Get IsDirty() As Boolean
    IsDirty = mbIsDirty
End Property

Private Property Let IsDirty(ByVal value As Boolean)
    mbIsDirty = value
    If OutlookSite Is Nothing Then
        'MsgBox "Can not save the settings because the Outlook as a parent was not available"
    Else
        If mbIsDirty = True Then
            OutlookSite.OnStatusChange
        End If
    End If
End Property

Private Property Get OutlookSite() As Outlook.PropertyPageSite
    On Error Resume Next
    Set OutlookSite = Parent
End Property

Private Sub UserControl_InitProperties()
    txtLineBeforeQuotes.Text = GetSetting(REG_APP_NAME, REG_SECTION_SETTINGS, REG_KEY_BEFORE_QUOTE_STRING, DEFAULT_BEFORE_INSERTE_STRING)
    txtLineAfterQuotes.Text = GetSetting(REG_APP_NAME, REG_SECTION_SETTINGS, REG_KEY_AFTER_QUOTE_STRING, DEFAULT_AFTER_INSERTE_STRING)
End Sub


'------------------------------------------------------------------------
'Shital's programming signature. (C) Shital Shah
'***************************************
'NOTE: This code can be safely removed without affecting program function
Private Sub RandomizeAnimation()
    Randomize
'    If Int((100 * Rnd) Mod 5) = 0 Then
'        Label1.Caption = "Shital Shah"
'        Label1.Font.Size = 16
'        Label2.Caption = Label1.Caption
'        Label2.Font.Size = Label1.Font.Size
'    End If
    lblSignatureAnimationLabel1.ForeColor = RGB(Int(255 * Rnd), Int(255 * Rnd), Int(255 * Rnd))
    lblSignatureAnimationLabel2.ForeColor = RGB(Int(255 * Rnd), Int(255 * Rnd), Int(255 * Rnd))
    '***************************************
End Sub
Private Sub StartAnimation()

    On Error GoTo ERR_StartAnimation
    
    Call RandomizeAnimation
    
    fraShadowTravel.Visible = False

    Static x1 As Long
    Static x2 As Long
    Static bInitialValueSaved As Boolean
    If Not bInitialValueSaved Then    'Not initialised
        x1 = linSignatureAnimationLine1.x1
        x2 = linSignatureAnimationLine1.x2
        bInitialValueSaved = True
    Else
        'Restore linSignatureAnimationLine1 pos
        linSignatureAnimationLine1.x1 = x1
        linSignatureAnimationLine1.x2 = x2
    End If
    linSignatureAnimationLine2.x1 = linSignatureAnimationLine1.x1
    linSignatureAnimationLine2.x2 = linSignatureAnimationLine1.x2
    linSignatureAnimationLine2.Y1 = linSignatureAnimationLine1.Y1
    linSignatureAnimationLine2.Y2 = linSignatureAnimationLine1.Y2
    linSignatureAnimationLine1.Visible = True
    linSignatureAnimationLine2.Visible = True
    fraShadowTravel.Visible = True
    tmrAnimation.Enabled = False
    tmrAnimation.Interval = 30
    tmrAnimation.Enabled = True
Exit Sub
ERR_StartAnimation:
    Call StopAnimation
End Sub
Private Sub tmrAnimation_Timer()

    Static lAnimationPhase As Long
    
    On Error GoTo ERR_tmrAnimation_Timer
    
    Select Case lAnimationPhase
        Case -1
            lAnimationPhase = 0
            Call StartAnimation
        Case 0
            If linSignatureAnimationLine1.x1 > fraShadowTravel.Width Then
                lAnimationPhase = lAnimationPhase + 1
            Else
                'Advance lines
                linSignatureAnimationLine1.x1 = linSignatureAnimationLine1.x1 + 70
                linSignatureAnimationLine1.x2 = linSignatureAnimationLine1.x2 + 70
                linSignatureAnimationLine2.x1 = linSignatureAnimationLine1.x1
                linSignatureAnimationLine2.x2 = linSignatureAnimationLine1.x2
            End If
'        Case 1
'            If Label1.Left >= Label2.Left Then
'                Label1.Top = Label2.Top
'                lAnimationPhase = lAnimationPhase + 1
'            Else
'                Label1.Left = Label1.Left + 2
'            End If
        Case Else
            Call StopAnimation
            lAnimationPhase = -1
            tmrAnimation.Interval = 3000    'Wait for next animation
            tmrAnimation.Enabled = True
    End Select
Exit Sub
ERR_tmrAnimation_Timer:
    Call StopAnimation
End Sub
Private Sub StopAnimation()
    tmrAnimation.Enabled = False
    'fraShadowTravel.Visible = False
    linSignatureAnimationLine1.Visible = False
    linSignatureAnimationLine2.Visible = False
End Sub
Private Sub Label5_Click(Index As Integer)
    Call OpenAnyFile(Label5(Index).Caption)
End Sub
Private Function OpenAnyFile(ByVal vsFileName As String, Optional ByVal vsParameters As String = "") As Boolean
    OpenAnyFile = ShellExecute(0, "open", vsFileName, vsParameters, "", SW_SHOW) > 32
End Function
'-----------Programming signature end------------------------------------------------------------------------

Private Sub UserControl_Show()
    Call StartAnimation
End Sub
