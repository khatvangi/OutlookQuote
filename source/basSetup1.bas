Attribute VB_Name = "basSetup1"
Option Explicit
Option Compare Text

'
' Global Constants
'

'Return values for setup toolkit functions
Global Const gintRET_CONT% = 1
Global Const gintRET_CANCEL% = 2
Global Const gintRET_EXIT% = 3
Global Const gintRET_ABORT% = 4
Global Const gintRET_FATAL% = 5
Global Const gintRET_FINISHEDSUCCESS% = 6 'Used only as parameter to ExitSetup at end of successful install

'Error levels for GetAppRemovalCmdLine()
Global Const APPREMERR_NONE = 0 'no error
Global Const APPREMERR_FATAL = 1 'fatal error
Global Const APPREMERR_NONFATAL = 2 'non-fatal error, user chose to abort
Global Const APPREMERR_USERCANCEL = 3 'user chose to cancel (no error)

'Flag for Path Dialog specifying Source or Dest directory needed
Global Const gstrDIR_SRC$ = "S"
Global Const gstrDIR_DEST$ = "D"

'Beginning of lines in [Files], [Bootstrap], and [Licenses] sections of SETUP.LST
Global Const gstrINI_FILE$ = "File"
Global Const gstrINI_REMOTE$ = "Remote"
Global Const gstrINI_LICENSE$ = "License"
'
' Command line constants
'
Global Const gstrSILENTSWITCH = "s"
Global Const gstrSMSSWITCH = "q"
'
' Icon Information
'
Global Const gsGROUP As String = "Group"
Global Const gsICON As String = "Icon"
Global Const gsTITLE As String = "Title"
Global Const gsICONGROUP As String = "IconGroups"

'
'Type Definitions
'
Type FILEINFO                                               ' Setup information file line format
    intDiskNum As Integer                                   ' disk number
    fSplit As Integer                                       ' split flag
    strSrcName As String                                    ' name of source file
    strDestName As String                                   ' name of destination file
    strDestDir As String                                    ' destination directory
    strRegister As String                                   ' registration info
    fShared As Boolean                                      ' whether the file is shared or private
    fSystem As Boolean                                      ' whether the file is a system file (i.e. should be installed but never removed)
    varDate As Variant                                      ' file date
    lFileSize As Long                                       ' file size
    sVerInfo As VERINFO                                     ' file version number
    strReserved As String                                   ' Reserved. Leave empty, or error.
    strProgramIconTitle As String                                ' Caption for icon in program group
    strProgramIconCmdLine As String                         ' Command Line for icon in program group
End Type

Type DISKINFO                                               ' Disk drive information
    lAvail As Long                                          ' Bytes available on drive
    lReq As Long                                            ' Bytes required for setup
    lMinAlloc As Long                                       ' minimum allocation unit
End Type

Type DESTINFO                                               ' save dest dir for certain files
    strAppDir As String
    strAUTMGR32 As String
    strRACMGR32 As String
End Type

Type REGINFO                                                ' save registration info for files
    strFilename As String
    strRegister As String
    
    'The following are used only for remote server registration
    strNetworkAddress As String
    strNetworkProtocol As String
    intAuthentication As Integer
    fDCOM As Boolean      ' True if DCOM, otherwise False
End Type

'
'Global Variables
'
Global gstrSETMSG As String
Global gfRetVal As Integer                                  'return value for form based functions
Global gstrAppName As String                                'name of app being installed
Global gstrTitle As String                                  '"setup" name of app being installed
Public gstrDefGroup As String                               'Default name for group -- from setup.lst
Global gstrDestDir As String                                'dest dir for application files
Global gstrAppExe As String                                 'name of app .EXE being installed
Public gstrAppToUninstall As String                         ' Name of app exe/ocx/dll to be uninstalled.  Should be the same as gstrAppExe in most cases.
Global gstrSrcPath As String                                'path of source files
Global gstrSetupInfoFile As String                          'pathname of SETUP.LST file
Global gstrWinDir As String                                 'windows directory
Global gstrWinSysDir As String                              'windows\system directory
Global gsDiskSpace() As DISKINFO                            'disk space for target drives
Global gstrDrivesUsed As String                             'dest drives used by setup
Global glTotalCopied As Long                                'total bytes copied so far
Global gintCurrentDisk As Integer                           'current disk number being installed
Global gsDest As DESTINFO                                   'dest dirs for certain files
Global gstrAppRemovalLog As String                           'name of the app removal logfile
Global gstrAppRemovalEXE As String                           'name of the app removal executable
Global gfAppRemovalFilesMoved As Boolean                     'whether or not the app removal files have been moved to the application directory
Global gfForceUseDefDest As Boolean                         'If set to true, then the user will not be prompted for the destination directory
Global fMainGroupWasCreated As Boolean                     'Whether or not a main folder/group has been created
Public gfRegDAO As Boolean                                 ' If this gets set to true in the code, then
                                                           ' we need to add some registration info for DAO
                                                           ' to the registry.

Global gsCABNAME As String
Global gsTEMPDIR As String

Global Const gsINI_CABNAME As String = "Cab"
Global Const gsINI_TEMPDIR As String = "TmpDir"
'
'Form/Module Constants
'

'Possible ProgMan actions
Const mintDDE_ITEMADD% = 1                                  'AddProgManItem flag
Const mintDDE_GRPADD% = 2                                   'AddProgManGroup flag

'Special file names
Const mstrFILE_APPREMOVALLOGBASE$ = "ST6UNST"               'Base name of the app removal logfile
Const mstrFILE_APPREMOVALLOGEXT$ = ".LOG"                   'Default extension for the app removal logfile
Const mstrFILE_AUTMGR32 = "AUTMGR32.EXE"
Const mstrFILE_RACMGR32 = "RACMGR32.EXE"
Const mstrFILE_CTL3D32$ = "CTL3D32.DLL"
Const mstrFILE_RICHED32$ = "RICHED32.DLL"

'Name of temporary file used for concatenation of split files
Const mstrCONCATFILE$ = "VB5STTMP.CCT"

'setup information file registration macros
Const mstrDLLSELFREGISTER$ = "$(DLLSELFREGISTER)"
Const mstrEXESELFREGISTER$ = "$(EXESELFREGISTER)"
Const mstrTLBREGISTER$ = "$(TLBREGISTER)"
Const mstrREMOTEREGISTER$ = "$(REMOTE)"
Const mstrVBLREGISTER$ = "$(VBLREGISTER)"  ' Bug 5-8039

'
'Form/Module Variables
'
Private msRegInfo() As REGINFO                                  'files to be registered
Private mlTotalToCopy As Long                                   'total bytes to copy
Private mintConcatFile As Integer                               'handle of dest file for concatenation
Private mlSpaceForConcat As Long                                'extra space required for concatenation
Private mstrConcatDrive As String                               'drive to use for concatenation
Private mstrVerTmpName As String                                'temp file name for VerInstallFile API

' Hkey cache (used for logging purposes)
Private Type HKEY_CACHE
    hKey As Long
    strHkey As String
End Type

Private hkeyCache() As HKEY_CACHE

' Registry manipulation API's (32-bit)
Global Const HKEY_CLASSES_ROOT = &H80000000
Global Const HKEY_CURRENT_USER = &H80000001
Global Const HKEY_LOCAL_MACHINE = &H80000002
Global Const HKEY_USERS = &H80000003
Const ERROR_SUCCESS = 0&
Const ERROR_NO_MORE_ITEMS = 259&

Const REG_SZ = 1
Const REG_BINARY = 3
Const REG_DWORD = 4


Declare Function OSRegCloseKey Lib "advapi32" Alias "RegCloseKey" (ByVal hKey As Long) As Long
Declare Function OSRegCreateKey Lib "advapi32" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpszSubKey As String, phkResult As Long) As Long
Declare Function OSRegDeleteKey Lib "advapi32" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpszSubKey As String) As Long
Declare Function OSRegEnumKey Lib "advapi32" Alias "RegEnumKeyA" (ByVal hKey As Long, ByVal iSubKey As Long, ByVal lpszName As String, ByVal cchName As Long) As Long
Declare Function OSRegOpenKey Lib "advapi32" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpszSubKey As String, phkResult As Long) As Long
Declare Function OSRegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpszValueName As String, ByVal dwReserved As Long, lpdwType As Long, lpbData As Any, cbData As Long) As Long
Declare Function OSRegSetValueEx Lib "advapi32" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long

Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Declare Function ExtractFileFromCab Lib "vb6stkit.dll" (ByVal Cab As String, ByVal File As String, ByVal Dest As String) As Long


'-----------------------------------------------------------
' FUNCTION: AddQuotesToFN
'
' Given a pathname (directory and/or filename), returns
'   that pathname surrounded by double quotes if the
'   path contains spaces or commas.  This is required for
'   setting up an icon correctly, since otherwise such paths
'   would be interpreted as a pathname plus arguments.
'-----------------------------------------------------------
'
Function AddQuotesToFN(ByVal strFilename) As String
    If InStr(strFilename, " ") Or InStr(strFilename, ",") Then
        AddQuotesToFN = """" & strFilename & """"
    Else
        AddQuotesToFN = strFilename
    End If
End Function



'-----------------------------------------------------------
' SUB: CalcFinalSize
'
' Computes the space required for a file of the size
' specified on the given dest path.  This includes the
' file size plus a padding to ensure that the final size
' is a multiple of the minimum allocation unit for the
' dest drive
'-----------------------------------------------------------
'
Function CalcFinalSize(lBaseFileSize As Long, strDestPath As String) As Long
    Dim lMinAlloc As Long
    Dim intPadSize As Long

    lMinAlloc = gsDiskSpace(InStr(gstrDrivesUsed, Left$(strDestPath, 1))).lMinAlloc
    intPadSize = lMinAlloc - (lBaseFileSize Mod lMinAlloc)
    If intPadSize = lMinAlloc Then
        intPadSize = 0
    End If

    CalcFinalSize = lBaseFileSize + intPadSize
End Function

'-----------------------------------------------------------
' SUB: CenterForm
'
' Centers the passed form just above center on the screen
'-----------------------------------------------------------
'
Sub CenterForm(frm As Form)
    SetMousePtr vbHourglass

    frm.Top = (Screen.Height * 0.85) \ 2 - frm.Height \ 2
    frm.Left = Screen.Width \ 2 - frm.Width \ 2

    SetMousePtr gintMOUSE_DEFAULT
End Sub

'-----------------------------------------------------------
' FUNCTION: DecideIncrementRefCount
'
' Increments the reference count of a file under 32-bits
' if the file is a shared file.
'
' IN: [strFullPath] - full pathname of the file to reference
'                     count.  Example:
'                     'C:\MYAPP\MYAPP.DAT'
'     [fShared] - whether the file is shared or private
'     [fSystem] - The file is a system file
'     [fFileAlreadyExisted] - whether or not the file already
'                             existed on the hard drive
'                             before our setup program
'-----------------------------------------------------------
'
Sub DecideIncrementRefCount(ByVal strFullPath As String, ByVal fShared As Boolean, ByVal fSystem As Boolean, ByVal fFileAlreadyExisted As Boolean)
    'Reference counting takes place under both Windows 95 and Windows NT
    If fShared Or fSystem Then
        IncrementRefCount strFullPath, fFileAlreadyExisted
    End If
End Sub
            
'-----------------------------------------------------------
' SUB: EtchedLine
'
' Draws an 'etched' line upon the specified form starting
' at the X,Y location passed in and of the specified length.
' Coordinates are in the current ScaleMode of the passed
' in form.
'
' IN: [frmEtch] - form to draw the line upon
'     [intX1] - starting horizontal of line
'     [intY1] - starting vertical of line
'     [intLength] - length of the line
'-----------------------------------------------------------
'
Sub EtchedLine(frmEtch As Form, ByVal intX1 As Integer, ByVal intY1 As Integer, ByVal intLength As Integer)
    Const lWHITE& = vb3DHighlight
    Const lGRAY& = vb3DShadow

    frmEtch.Line (intX1, intY1)-(intX1 + intLength, intY1), lGRAY
    frmEtch.Line (frmEtch.CurrentX + 5, intY1 + 20)-(intX1 - 5, intY1 + 20), lWHITE
End Sub

'-----------------------------------------------------------
' SUB: ExeSelfRegister
'
' Synchronously runs the file passed in (which should be
' an executable file that supports the /REGSERVER switch,
' for instance, a VB5 generated ActiveX Component .EXE).
'
' IN: [strFileName] - .EXE file to register
'-----------------------------------------------------------
'
Sub ExeSelfRegister(ByVal strFilename As String)
    Const strREGSWITCH$ = " /REGSERVER"

    Dim fShell As Integer

    '
    'Synchronously shell out and run the .EXE with the self registration switch
    '
    fShell = SyncShell(AddQuotesToFN(strFilename) & strREGSWITCH, INFINITE, , True)
End Sub

'-----------------------------------------------------------
' FUNCTION: GetFileName
'
' Return the filename portion of a path
'
'-----------------------------------------------------------
'
Function GetFileName(ByVal strPath As String) As String
    Dim strFilename As String
    Dim iSep As Integer
    
    strFilename = strPath
    Do
        iSep = InStr(strFilename, gstrSEP_DIR)
        If iSep = 0 Then iSep = InStr(strFilename, gstrCOLON)
        If iSep = 0 Then
            GetFileName = strFilename
            Exit Function
        Else
            strFilename = Right(strFilename, Len(strFilename) - iSep)
        End If
    Loop
End Function

'-----------------------------------------------------------
' FUNCTION: GetFileSize
'
' Determine the size (in bytes) of the specified file
'
' IN: [strFileName] - name of file to get size of
'
' Returns: size of file in bytes, or -1 if an error occurs
'-----------------------------------------------------------
'
Function GetFileSize(strFilename As String) As Long
    On Error Resume Next

    GetFileSize = FileLen(strFilename)

    If Err > 0 Then
        GetFileSize = -1
        Err = 0
    End If
End Function

'-----------------------------------------------------------
' FUNCTION: GetAppRemovalCmdLine
'
' Returns the correct command-line arguments (including
' path to the executable for use in calling the
' application removal executable)
'
' IN: [strAppRemovalEXE] - Full path/filename of the app removal EXE
'     [strAppRemovalLog] - Full path/filename of the app removal logfile
'     [strSilentLog] - Full path/filename of the file to log messages to when in silent mode.
'                       If this is an empty string then silent mode is turned off for uninstall.
'     [fSMS] - Boolean.  If True, we have been doing an SMS install and must tell the Uninstaller
'              to also do an SMS uninstall.  SMS is the Microsoft Systems Management Server.
'     [nErrorLevel] - Error level:
'                        APPREMERR_NONE - no error
'                        APPREMERR_FATAL - fatal error
'                        APPREMERR_NONFATAL - non-fatal error, user chose to abort
'                        APPREMERR_USERCANCEL - user chose to cancel (no error)
'     [fWaitForParent] - True if the application removal utility should wait
'                        for the parent (this process) to finish before starting
'                        to remove files.  Otherwise it may not be able to remove
'                        this process' executable file, depending upon timing.
'                        Defaults to False if not specified.
'-----------------------------------------------------------
'
Function GetAppRemovalCmdLine(ByVal strAppRemovalEXE As String, ByVal strAppRemovalLog, ByVal strSilentLog As String, ByVal fSMS As Boolean, ByVal nErrorLevel As Integer, Optional fWaitForParent)
    Dim strEXE As String
    Dim strLog As String
    Dim strSilent As String
    Dim strErrLevel As String
    Dim strForce As String
    Dim strWait As String
    Dim strSMS As String

    If IsMissing(fWaitForParent) Then
        fWaitForParent = False
    End If
    
    strEXE = AddQuotesToFN(strAppRemovalEXE)
    strLog = "-n " & """" & GetLongPathName(strAppRemovalLog) & """"
    If gfSilent And strSilentLog <> vbNullString Then
        strSilent = "/s " & """" & strSilentLog & """"
    Else
        strSilent = vbNullString
    End If
    
    strSMS = IIf(fSMS, " /q ", vbNullString)
    
    strErrLevel = IIf(nErrorLevel <> APPREMERR_NONE, "-e " & Format(nErrorLevel), "")
    If nErrorLevel <> APPREMERR_NONE Then
        strForce = " -f"
    End If
    If fWaitForParent Then
        Dim curProcessId As Currency
        Dim Wrap As Currency
        Dim lProcessId As Long
        Dim cProcessId As Currency
        
        Wrap = 2 * (CCur(&H7FFFFFFF) + 1)

        'Always print as an unsigned long
        lProcessId = GetCurrentProcessId()
        cProcessId = lProcessId
        If cProcessId < 0 Then cProcessId = cProcessId + Wrap

        strWait = " -w " & str(cProcessId)
    End If
    
    GetAppRemovalCmdLine = strEXE & " " & strLog & " " & strSilent & " " & strSMS & strErrLevel & strForce & strWait
End Function

'-----------------------------------------------------------
' FUNCTION: IncrementRefCount
'
' Increments the reference count on a file in the registry
' so that it may properly be removed if the user chooses
' to remove this application.
'
' IN: [strFullPath] - FULL path/filename of the file
'     [fFileAlreadyExisted] - indicates whether the given
'                             file already existed on the
'                             hard drive
'-----------------------------------------------------------
'
Sub IncrementRefCount(ByVal strFullPath As String, ByVal fFileAlreadyExisted As Boolean)
    Dim strSharedDLLsKey As String
    strSharedDLLsKey = RegPathWinCurrentVersion() & "\SharedDLLs"
    
    'We must always use the LFN for the filename, so that we can uniquely
    'and accurately identify the file in the registry.
    strFullPath = GetLongPathName(strFullPath)
    
    'Get the current reference count for this file
    Dim fSuccess As Boolean
    Dim hKey As Long
    fSuccess = RegCreateKey(HKEY_LOCAL_MACHINE, strSharedDLLsKey, "", hKey)
    If fSuccess Then
        Dim lCurRefCount As Long
        If Not RegQueryRefCount(hKey, strFullPath, lCurRefCount) Then
            'No current reference count for this file
            If fFileAlreadyExisted Then
                'If there was no reference count, but the file was found
                'on the hard drive, it means one of two things:
                '  1) This file is shipped with the operating system
                '  2) This file was installed by an older setup program
                '     that does not do reference counting
                'In either case, the correct conservative thing to do
                'is assume that the file is needed by some application,
                'which means it should have a reference count of at
                'least 1.  This way, our application removal program
                'will not delete this file.
                lCurRefCount = 1

            Else
                lCurRefCount = 0
            End If
        End If
        
        'Increment the count in the registry
        fSuccess = RegSetNumericValue(hKey, strFullPath, lCurRefCount + 1, False)
        If Not fSuccess Then
            GoTo DoErr
        End If
        RegCloseKey hKey
    Else
        GoTo DoErr
    End If
    
    Exit Sub
    
DoErr:
    'An error message should have already been shown to the user
    Exit Sub
End Sub

'-----------------------------------------------------------
' FUNCTION: IsDisplayNameUnique
'
' Determines whether a given display name for registering
'   the application removal executable is unique or not.  This
'   display name is the title which is presented to the
'   user in Windows 95's control panel Add/Remove Programs
'   applet.
'
' IN: [hkeyAppRemoval] - open key to the path in the registry
'                       containing application removal entries
'     [strDisplayName] - the display name to test for uniqueness
'
' Returns: True if the given display name is already in use,
'          False if otherwise
'-----------------------------------------------------------
'
Function IsDisplayNameUnique(ByVal hkeyAppRemoval As Long, ByVal strDisplayName As String) As Boolean
    Dim lIdx As Long
    Dim strSubkey As String
    Dim strDisplayNameExisting As String
    Const strKEY_DISPLAYNAME$ = "DisplayName"
    
    IsDisplayNameUnique = True
    
    lIdx = 0
    Do
        Select Case RegEnumKey(hkeyAppRemoval, lIdx, strSubkey)
            Case ERROR_NO_MORE_ITEMS
                'No more keys - must be unique
                Exit Do
            Case ERROR_SUCCESS
                'We have a key to some application removal program.  Compare its
                '  display name with ours
                Dim hkeyExisting As Long
                
                If RegOpenKey(hkeyAppRemoval, strSubkey, hkeyExisting) Then
                    If RegQueryStringValue(hkeyExisting, strKEY_DISPLAYNAME, strDisplayNameExisting) Then
                        If strDisplayNameExisting = strDisplayName Then
                            'There is a match to an existing display name
                            IsDisplayNameUnique = False
                            RegCloseKey hkeyExisting
                            Exit Do
                        End If
                    End If
                    RegCloseKey hkeyExisting
                End If
            Case Else
                'Error, we must assume it's unique.  An error will probably
                '  occur later when trying to add to the registry
                Exit Do
            'End Case
        End Select
        lIdx = lIdx + 1
    Loop
End Function

'-----------------------------------------------------------
' FUNCTION: IsNewerVer
'
' Compares two file version structures and determines
' whether the source file version is newer (greater) than
' the destination file version.  This is used to determine
' whether a file needs to be installed or not
'
' IN: [sSrcVer] - source file version information
'     [sDestVer] - dest file version information
'
' Returns: True if source file is newer than dest file,
'          False if otherwise
'-----------------------------------------------------------
'
Function IsNewerVer(sSrcVer As VERINFO, sDestVer As VERINFO) As Integer
    IsNewerVer = False

    If sSrcVer.nMSHi > sDestVer.nMSHi Then GoTo INVNewer
    If sSrcVer.nMSHi < sDestVer.nMSHi Then GoTo INVOlder
    
    If sSrcVer.nMSLo > sDestVer.nMSLo Then GoTo INVNewer
    If sSrcVer.nMSLo < sDestVer.nMSLo Then GoTo INVOlder
    
    If sSrcVer.nLSHi > sDestVer.nLSHi Then GoTo INVNewer
    If sSrcVer.nLSHi < sDestVer.nLSHi Then GoTo INVOlder
    
    If sSrcVer.nLSLo > sDestVer.nLSLo Then GoTo INVNewer

    GoTo INVOlder

INVNewer:
    IsNewerVer = True
INVOlder:
End Function

'-----------------------------------------------------------
' FUNCTION: MakePath
'
' Creates the specified directory path
'
' IN: [strDirName] - name of the dir path to make
'     [fAllowIgnore] - whether or not to allow the user to
'                      ignore any encountered errors.  If
'                      false, the function only returns
'                      if successful.  If missing, this
'                      defaults to True.
'
' Returns: True if successful, False if error and the user
'          chose to ignore.  (The function does not return
'          if the user selects ABORT/CANCEL on an error.)
'-----------------------------------------------------------
'
Public Function MakePath(ByVal strDir As String, Optional ByVal fAllowIgnore) As Boolean
    If IsMissing(fAllowIgnore) Then
        fAllowIgnore = True
    End If
    
    Do
        If MakePathAux(strDir) Then
            MakePath = True
            Exit Function
        Else
            Dim strMsg As String
            Dim iRet As Integer
            
            strMsg = ResolveResString(resMAKEDIR) & vbLf & strDir
            iRet = MsgError(strMsg, IIf(fAllowIgnore, vbAbortRetryIgnore, vbRetryCancel) Or vbExclamation Or vbDefaultButton2, gstrSETMSG)
            '
            ' if we are running silent then we
            ' can't continue.  Previous MsgError
            ' took care of write silent log entry.
            '
            If gfNoUserInput = True Then
                ExitSetup
            End If
            
            Select Case iRet
                Case vbAbort, vbCancel
                    ExitSetup
                Case vbIgnore
                    MakePath = False
                    Exit Function
                Case vbRetry
                'End Case
            End Select
        End If
    Loop
End Function

Function ExitSetup()
    MsgBox "Setup can not continue"
    End
End Function
'-----------------------------------------------------------
' SUB: KillTempFolder
' BUG FIX #6-34583
'
' Deletes the temporary files stored in the temp folder
'
Private Sub KillTempFolder()

    Const sWILD As String = "*.*"
    Dim sFile As String
    
    sFile = Dir(gsTEMPDIR & sWILD)
    While sFile <> vbNullString
        SetAttr gsTEMPDIR & sFile, vbNormal
        Kill gsTEMPDIR & sFile
        sFile = Dir
    Wend
    RmDir gsTEMPDIR
End Sub

'-----------------------------------------------------------
' SUB: ParseDateTime
'
' Same as CDate with a string argument, except that it
' ignores the current localization settings.  This is
' important because SETUP.LST always uses the same
' format for dates.
'
' IN: [strDate] - string representing the date in
'                 the format mm/dd/yy or mm/dd/yyyy
' OUT: The date which strDate represents
'-----------------------------------------------------------
'
Function ParseDateTime(ByVal strDateTime As String) As Date
    Const strDATESEP$ = "/"
    Const strTIMESEP$ = ":"
    Const strDATETIMESEP$ = " "
    Dim iMonth As Integer
    Dim iDay As Integer
    Dim iYear As Integer
    Dim iHour As Integer
    Dim iMinute As Integer
    Dim iSecond As Integer
    Dim iPos As Integer
    Dim vTime As Date
    
    iPos = InStr(strDateTime, strDATESEP)
    If iPos = 0 Then GoTo Err
    iMonth = Val(Left$(strDateTime, iPos - 1))
    strDateTime = Mid$(strDateTime, iPos + 1)
    
    iPos = InStr(strDateTime, strDATESEP)
    If iPos = 0 Then GoTo Err
    iDay = Val(Left$(strDateTime, iPos - 1))
    strDateTime = Mid$(strDateTime, iPos + 1)
    
    iPos = InStr(strDateTime, strDATETIMESEP)
    If iPos = 0 Then GoTo SkipTime
    iYear = Val(Left$(strDateTime, iPos - 1))
    strDateTime = Mid$(strDateTime, iPos + 1)
    
    vTime = TimeSerial(0, 0, 0)
    
    iPos = InStr(strDateTime, strTIMESEP)
    If iPos = 0 Then GoTo SkipTime
    iHour = Val(Left$(strDateTime, iPos - 1))
    strDateTime = Mid$(strDateTime, iPos + 1)
    
    iPos = InStr(strDateTime, strTIMESEP)
    If iPos = 0 Then GoTo SkipTime
    iMinute = Val(Left$(strDateTime, iPos - 1))
    strDateTime = Mid$(strDateTime, iPos + 1)
    
    iSecond = Val(strDateTime)
    
    vTime = TimeSerial(iHour, iMinute, iSecond)
    
SkipTime:
    
    If iYear < 100 Then iYear = iYear + 1900
    
    ParseDateTime = DateSerial(iYear, iMonth, iDay) + vTime
    
    Exit Function
    
Err:
    Error 13 'Type mismatch error, same as intrinsic CDate triggers on error
End Function

Function SrcFileMissing(ByVal strSrcDir As String, ByVal strSrcFile As String, ByVal intDiskNum As Integer) As Boolean
'-----------------------------------------------------------
' FUNCTION: SrcFileMissing
'
' Tries to locate the file strSrcFile by first looking
' in the strSrcDir directory, then in the DISK(x+1)
' directory if it exists.
'
' IN: [strSrcDir] - Directory/Path where file should be.
'     [strSrcFile] - File we are looking for.
'     [intDiskNum] - Disk number we are expecting file
'                    to be on.
'
' Returns: True if file not found; otherwise, false
'-----------------------------------------------------------
    Dim fFound As Boolean
    Dim strMultDirBaseName As String
    
    fFound = False
    
    AddDirSep strSrcDir
    '
    ' First check to see if it's in the main src directory.
    ' This would happen if someone copied the contents of
    ' all the floppy disks to a single directory on the
    ' hard drive.  We should allow this to work.
    '
    ' This test would also let us know if the user inserted
    ' the wrong floppy disk or if a network connection is
    ' unavailable.
    '
    If FileExists(strSrcDir & strSrcFile) = True Then
        fFound = True
        GoTo doneSFM
    End If
    '
    ' Next try the DISK(x) subdirectory of the main src
    ' directory.  This would happen if the floppy disks
    ' were copied into directories named DISK1, DISK2,
    ' DISK3,..., DISKN, etc.
    '
    strMultDirBaseName = ResolveResString(resCOMMON_MULTDIRBASENAME)
    If FileExists(strSrcDir & ".." & gstrSEP_DIR & strMultDirBaseName & Format(intDiskNum) & gstrSEP_DIR & strSrcFile) = True Then
        fFound = True
        GoTo doneSFM
    End If
    
doneSFM:
    SrcFileMissing = Not fFound
End Function
'-----------------------------------------------------------
' FUNCTION: ReadIniFile
'
' Reads a value from the specified section/key of the
' specified .INI file
'
' IN: [strIniFile] - name of .INI file to read
'     [strSection] - section where key is found
'     [strKey] - name of key to get the value of
'
' Returns: non-zero terminated value of .INI file key
'-----------------------------------------------------------
'
Function ReadIniFile(ByVal strIniFile As String, ByVal strsection As String, ByVal strKey As String) As String
    Dim strBuffer As String
    Dim intPos As Integer

    '
    'If successful read of .INI file, strip any trailing zero returned by the Windows API GetPrivateProfileString
    '
    strBuffer = Space$(gintMAX_SIZE)
    
    If GetPrivateProfileString(strsection, strKey, vbNullString, strBuffer, gintMAX_SIZE, strIniFile) > 0 Then
        ReadIniFile = RTrim$(StripTerminator(strBuffer))
    Else
        ReadIniFile = vbNullString
    End If
End Function

'-----------------------------------------------------------
' FUNCTION: RegCloseKey
'
' Closes an open registry key.
'
' Returns: True on success, else False.
'-----------------------------------------------------------
'
Function RegCloseKey(ByVal hKey As Long) As Boolean
    Dim lResult As Long
    
    On Error GoTo 0
    lResult = OSRegCloseKey(hKey)
    RegCloseKey = (lResult = ERROR_SUCCESS)
End Function

'-----------------------------------------------------------
' FUNCTION: RegCreateKey
'
' Opens (creates if already exists) a key in the system registry.
'
' IN: [hkey] - The HKEY parent.
'     [lpszSubKeyPermanent] - The first part of the subkey of
'         'hkey' that will be created or opened.  The application
'         removal utility (32-bit only) will never delete any part
'         of this subkey.  May NOT be an empty string ("").
'     [lpszSubKeyRemovable] - The subkey of hkey\lpszSubKeyPermanent
'         that will be created or opened.  If the application is
'         removed (32-bit only), then this entire subtree will be
'         deleted, if it is empty at the time of application removal.
'         If this parameter is an empty string (""), then the entry
'         will not be logged.
'
' OUT: [phkResult] - The HKEY of the newly-created or -opened key.
'
' Returns: True if the key was created/opened OK, False otherwise
'   Upon success, phkResult is set to the handle of the key.
'
'-----------------------------------------------------------
Function RegCreateKey(ByVal hKey As Long, ByVal lpszSubKeyPermanent As String, ByVal lpszSubKeyRemovable As String, phkResult As Long) As Boolean
    Dim lResult As Long
    Dim strHkey As String
    Dim fLog As Boolean
    Dim strSubKeyFull As String

    On Error GoTo 0

    If lpszSubKeyPermanent = "" Then
        RegCreateKey = False 'Error: lpszSubKeyPermanent must not = ""
        Exit Function
    End If
    
    If Left$(lpszSubKeyRemovable, 1) = "\" Then
        lpszSubKeyRemovable = Mid$(lpszSubKeyRemovable, 2)
    End If

    If lpszSubKeyRemovable = "" Then
        fLog = False
    Else
        fLog = True
    End If
    
    If lpszSubKeyRemovable <> "" Then
        strSubKeyFull = lpszSubKeyPermanent & "\" & lpszSubKeyRemovable
    Else
        strSubKeyFull = lpszSubKeyPermanent
    End If
    strHkey = strGetHKEYString(hKey)

    If fLog Then
        NewAction _
          gstrKEY_REGKEY, _
          """" & strHkey & "\" & lpszSubKeyPermanent & """" _
            & ", " & """" & lpszSubKeyRemovable & """"
    End If

    lResult = OSRegCreateKey(hKey, strSubKeyFull, phkResult)
    If lResult = ERROR_SUCCESS Then
        RegCreateKey = True
        If fLog Then
            CommitAction
        End If
        AddHkeyToCache phkResult, strHkey & "\" & strSubKeyFull
    Else
        RegCreateKey = False
        MsgError ResolveResString(resERR_REG), vbOKOnly Or vbExclamation, gstrTitle
        If fLog Then
            AbortAction
        End If
        If gfNoUserInput Then
            ExitSetup frmSetup1, gintRET_FATAL
        End If
    End If
End Function

'-----------------------------------------------------------
' FUNCTION: RegDeleteKey
'
' Deletes an existing key in the system registry.
'
' Returns: True on success, False otherwise
'-----------------------------------------------------------
'
Function RegDeleteKey(ByVal hKey As Long, ByVal lpszSubKey As String) As Boolean
    Dim lResult As Long
    
    On Error GoTo 0
    lResult = OSRegDeleteKey(hKey, lpszSubKey)
    RegDeleteKey = (lResult = ERROR_SUCCESS)
End Function

'-----------------------------------------------------------
' SUB: RegEdit
'
' Calls REGEDIT to add the information in the specifed file
' to the system registry.  If your .REG file requires path
' information based upon the destination directory given by
' the user, then you will need to write and call a .REG fixup
' routine before performing the registration below.
'
' WARNING: Use of this functionality under Win32 is not recommended,
' WARNING: because the application removal utility does not support
' WARNING: undoing changes that occur as a result of calling
' WARNING: REGEDIT on an arbitrary .REG file.
' WARNING: Instead, it is recommended that you use the RegCreateKey(),
' WARNING: RegOpenKey(), RegSetStringValue(), etc. functions in
' WARNING: this module instead.  These make entries to the
' WARNING: application removal logfile, thus enabling application
' WARNING: removal to undo such changes.
'
' IN: [strRegFile] - name of file containing reg. info
'-----------------------------------------------------------
'
Sub RegEdit(ByVal strRegFile As String)
    Const strREGEDIT$ = "REGEDIT /S "

    Dim fShellOK As Integer

    On Error Resume Next

    If FileExists(strRegFile) = True Then
        'Because regedit is a 16-bit application, it does not accept
        'double quotes around the filename.  Thus, if strRegFile
        'contains spaces, the only way to get this to work is to pass
        'regedit the short pathname version of the filename.
        strRegFile = GetShortPathName(strRegFile)
        
        fShellOK = SyncShell(strREGEDIT & strRegFile, INFINITE, , True)
        frmSetup1.Refresh
    Else
        MsgError ResolveResString(resCANTFINDREGFILE, "|1", strRegFile), vbExclamation Or vbOKOnly, gstrTitle
        ExitSetup frmSetup1, gintRET_FATAL
    End If

    Err = 0
End Sub

' FUNCTION: RegEnumKey
'
' Enumerates through the subkeys of an open registry
' key (returns the "i"th subkey of hkey, if it exists)
'
' Returns:
'   ERROR_SUCCESS on success.  strSubkeyName is set to the name of the subkey.
'   ERROR_NO_MORE_ITEMS if there are no more subkeys (32-bit only)
'   anything else - error
'
Function RegEnumKey(ByVal hKey As Long, ByVal i As Long, strKeyName As String) As Long
    Dim strResult As String
    
    strResult = String(300, " ")
    RegEnumKey = OSRegEnumKey(hKey, i, strResult, Len(strResult))
    strKeyName = StripTerminator(strResult)
End Function
'-----------------------------------------------------------
' SUB: RegisterDAO
'
' Special keys need to be added to the registry if
' DAO is installed.  This routine adds those keys.
'
' Note, these keys will not be uninstalled.
'
Sub RegisterDAO()
    Const strDAOKey = "CLSID\{F7A9C6E0-EFF2-101A-8185-00DD01108C6B}"
    Const strDAOKeyVal = "OLE 2.0 Link"
    Const strDAOInprocHandlerKey = "CLSID\{F7A9C6E0-EFF2-101A-8185-00DD01108C6B}\InprocHandler"
    Const strDAOInprocHandlerKeyVal = "ole2.dll"
    Const strDAOProgIDKey = "CLSID\{F7A9C6E0-EFF2-101A-8185-00DD01108C6B}\ProgID"
    Const strDAOProgIDKeyVal = "Access.OLE2Link"
    
    Dim hKey As Long
    
    If Not RegCreateKey(HKEY_CLASSES_ROOT, strDAOKey, "", hKey) Then
        '
        ' RegCreateKey displays an error if something goes wrong.
        '
        GoTo REGDAOError
    End If
    '
    ' Set the key's value
    '
    If Not RegSetStringValue(hKey, "", strDAOKeyVal, False) Then
        '
        ' RegSetStringValue displays an error if something goes wrong.
        '
        GoTo REGDAOError
    End If
    '
    ' Close the key
    '
    RegCloseKey hKey
    '
    ' Repeat the same process for the other two keys.
    '
    If Not RegCreateKey(HKEY_CLASSES_ROOT, strDAOInprocHandlerKey, "", hKey) Then GoTo REGDAOError
    If Not RegSetStringValue(hKey, "", strDAOInprocHandlerKeyVal, False) Then GoTo REGDAOError
    RegCloseKey hKey
    
    If Not RegCreateKey(HKEY_CLASSES_ROOT, strDAOProgIDKey, "", hKey) Then GoTo REGDAOError
    If Not RegSetStringValue(hKey, "", strDAOProgIDKeyVal, False) Then GoTo REGDAOError
    RegCloseKey hKey

    Exit Sub
        
REGDAOError:
    '
    ' Error messages should have already been displayed.
    '
    ExitSetup frmSetup1, gintRET_FATAL
        
End Sub
'-----------------------------------------------------------
' SUB: RegisterFiles
'
' Loop through the list (array) of files to register that
' was created in the CopySection function and register
' each file therein as required
'
' Notes: msRegInfo() array created by CopySection function
'-----------------------------------------------------------
'
Sub RegisterFiles()
    Const strEXT_EXE$ = "EXE"

    Dim intIdx As Integer
    Dim intLastIdx As Integer
    Dim strFilename As String
    Dim strMsg As String

    On Error Resume Next

    '
    'Get number of items to register, if none then we can get out of here
    '
    intLastIdx = UBound(msRegInfo)
    If Err > 0 Then
        GoTo RFCleanup
    End If

    For intIdx = 0 To intLastIdx
        strFilename = msRegInfo(intIdx).strFilename

        Select Case msRegInfo(intIdx).strRegister
            Case mstrDLLSELFREGISTER
                Dim intDllSelfRegRet As Integer
                Dim intErrRes As Integer
                Const FAIL_OLE = 2
                Const FAIL_LOAD = 3
                Const FAIL_ENTRY = 4
                Const FAIL_REG = 5
                
                NewAction gstrKEY_DLLSELFREGISTER, """" & strFilename & """"
                
RetryDllSelfReg:
                Err = 0
                intErrRes = 0
                intDllSelfRegRet = DLLSelfRegister(strFilename)
                If (Err <> 49) And (Err <> 0) Then
                    intErrRes = resCOMMON_CANTREGUNEXPECTED
                Else
                    Select Case intDllSelfRegRet
                        Case 0
                            'Good - everything's okay
                        Case FAIL_OLE
                            intErrRes = resCOMMON_CANTREGOLE
                        Case FAIL_LOAD
                            intErrRes = resCOMMON_CANTREGLOAD
                        Case FAIL_ENTRY
                            intErrRes = resCOMMON_CANTREGENTRY
                        Case FAIL_REG
                            intErrRes = resCOMMON_CANTREGREG
                        Case Else
                            intErrRes = resCOMMON_CANTREGUNEXPECTED
                        'End Case
                    End Select
                End If
                
                If intErrRes Then
                    'There was some kind of error
                    
                    'Log the more technical version of the error message -
                    'this would be too confusing to show to the end user
                    LogError ResolveResString(intErrRes, "|1", strFilename)
                    
                    'Now show a general error message to the user
AskWhatToDo:
                    strMsg = ResolveResString(resCOMMON_CANTREG, "|1", strFilename)
                    
                    Select Case MsgError(strMsg, vbExclamation Or vbAbortRetryIgnore, gstrTitle)
                        Case vbAbort:
                            ExitSetup frmSetup1, gintRET_ABORT
                            GoTo AskWhatToDo
                        Case vbRetry:
                            GoTo RetryDllSelfReg
                        Case vbIgnore:
                            AbortAction
                        'End Case
                    End Select
                Else
                    CommitAction
                End If
            Case mstrEXESELFREGISTER
                '
                'Only self register EXE files
                '
                If Extension(strFilename) = strEXT_EXE Then
                    NewAction gstrKEY_EXESELFREGISTER, """" & strFilename & """"
                    Err = 0
                    ExeSelfRegister strFilename
                    If Err Then
                        AbortAction
                    Else
                        CommitAction
                    End If
                End If
            Case mstrREMOTEREGISTER
                NewAction gstrKEY_REMOTEREGISTER, """" & strFilename & """"
                Err = 0
                RemoteRegister strFilename, msRegInfo(intIdx)
                If Err Then
                    AbortAction
                Else
                    CommitAction
                End If
            Case mstrTLBREGISTER
                NewAction gstrKEY_TLBREGISTER, """" & strFilename & """"
                '
                ' Call vb6stkit.dll's RegisterTLB export which calls
                ' LoadTypeLib and RegisterTypeLib.
                '
RetryTLBReg:
                If Not RegisterTLB(strFilename) Then
                    '
                    ' Registration of the TLB file failed.
                    '
                    LogError ResolveResString(resCOMMON_CANTREGTLB, "|1", strFilename)
TLBAskWhatToDo:
                    strMsg = ResolveResString(resCOMMON_CANTREGTLB, "|1", strFilename)
                    
                    Select Case MsgError(strMsg, vbExclamation Or vbAbortRetryIgnore, gstrTitle)
                        Case vbAbort:
                            ExitSetup frmSetup1, gintRET_ABORT
                            GoTo TLBAskWhatToDo
                        Case vbRetry:
                            GoTo RetryTLBReg
                        Case vbIgnore:
                            AbortAction
                        'End Case
                    End Select
                Else
                    CommitAction
                End If
            Case mstrVBLREGISTER
                '
                ' RegisterVBLFile takes care of logging, etc.
                '

                RegisterVBLFile strFilename
            Case Else
                RegEdit msRegInfo(intIdx).strRegister
            'End Case
        End Select
    Next


    Erase msRegInfo

RFCleanup:
    Err = 0
End Sub
'-----------------------------------------------------------
' SUB: RegisterLicenses
'
' Find all the setup.lst license entries and register
' them.
'-----------------------------------------------------------
'
Sub RegisterLicenses()
    Const strINI_LICENSES = "Licenses"
    Const strREG_LICENSES = "Licenses"
    Dim iLic As Integer
    Dim strLine As String
    Dim strLicKey As String
    Dim strLicVal As String
    Dim iCommaPos As Integer
    Dim strMsg As String
    Dim hkeyLicenses As Long
    Const strCopyright$ = "Licensing: Copying the keys may be a violation of established copyrights."

    'Make sure the Licenses key exists
    If Not RegCreateKey(HKEY_CLASSES_ROOT, strREG_LICENSES, "", hkeyLicenses) Then
        'RegCreateKey will have already displayed an error message
        '  if something's wrong
        ExitSetup frmSetup1, gintRET_FATAL
    End If
    If Not RegSetStringValue(hkeyLicenses, "", strCopyright, False) Then
        RegCloseKey hkeyLicenses
        ExitSetup frmSetup1, gintRET_FATAL
    End If
    RegCloseKey hkeyLicenses
    
    iLic = 1
    Do
        strLine = ReadIniFile(gstrSetupInfoFile, strINI_LICENSES, gstrINI_LICENSE & iLic)
        If strLine = vbNullString Then
            '
            ' We've got all the licenses.
            '
            Exit Sub
        End If
        strLine = strUnQuoteString(strLine)
        '
        ' We have a license, parse it and register it.
        '
        iCommaPos = InStr(strLine, gstrCOMMA)
        If iCommaPos = 0 Then
            '
            ' Looks like the setup.lst file is corrupt.  There should
            ' always be a comma in the license information that separates
            ' the license key from the license value.
            '
            GoTo RLError
        End If
        strLicKey = Left(strLine, iCommaPos - 1)
        strLicVal = Mid(strLine, iCommaPos + 1)
        
        RegisterLicense strLicKey, strLicVal
        
        iLic = iLic + 1
    Loop While strLine <> vbNullString
    Exit Sub
        
RLError:
    strMsg = gstrSetupInfoFile & vbLf & vbLf & ResolveResString(resINVLINE) & vbLf & vbLf
    strMsg = strMsg & ResolveResString(resSECTNAME) & strINI_LICENSES & vbLf & strLine
    MsgError strMsg, vbCritical, gstrTitle
    ExitSetup frmSetup1, gintRET_FATAL
End Sub
'-----------------------------------------------------------
' SUB: RegisterLicense
'
' Register license information given the key and default
' value.  License information always goes into
' HKEY_CLASSES_ROOT\Licenses.
'-----------------------------------------------------------
'
Sub RegisterLicense(strLicKey As String, strLicVal As String)
    Const strREG_LICENSES = "Licenses"
    Dim hKey As Long
    '
    ' Create the key
    '
    If Not RegCreateKey(HKEY_CLASSES_ROOT, strREG_LICENSES, strLicKey, hKey) Then
        '
        ' RegCreateKey displays an error if something goes wrong.
        '
        GoTo REGError
    End If
    '
    ' Set the key's value
    '
    If Not RegSetStringValue(hKey, "", strLicVal, True) Then
        '
        ' RegSetStringValue displays an error if something goes wrong.
        '
        GoTo REGError
    End If
    '
    ' Close the key
    '
    RegCloseKey hKey

    Exit Sub
        
REGError:
    '
    ' Error messages should have already been displayed.
    '
    ExitSetup frmSetup1, gintRET_FATAL
End Sub
'-----------------------------------------------------------
' SUB: RegisterVBLFile
'
' Register license information in a VB License (vbl) file.
' Basically, parse out the license info and then call
' RegisterLicense.
'
' If strVBLFile is not a valid VBL file, nothing is
' registered.
'-----------------------------------------------------------
'
Sub RegisterVBLFile(strVBLFile As String)
    Dim strLicKey As String
    Dim strLicVal As String
    
    GetLicInfoFromVBL strVBLFile, strLicKey, strLicVal
    If strLicKey <> vbNullString Then
        RegisterLicense strLicKey, strLicVal
    End If
End Sub

'----------------------------------------------------------
' SUB: RegisterAppRemovalEXE
'
' Registers the application removal program (Windows 95 only)
' or else places an icon for it in the application directory.
'
' Returns True on success, False otherwise.
'----------------------------------------------------------
Function RegisterAppRemovalEXE(ByVal strAppRemovalEXE As String, ByVal strAppRemovalLog As String, ByVal strGroupName As String) As Boolean
    On Error GoTo Err
    
    Const strREGSTR_VAL_AppRemoval_APPNAMELINE = "ApplicationName"
    Const strREGSTR_VAL_AppRemoval_DISPLAYNAME = "DisplayName"
    Const strREGSTR_VAL_AppRemoval_COMMANDLINE = "UninstallString"
    Const strREGSTR_VAL_AppRemoval_APPTOUNINSTALL = "AppToUninstall"
    
    
    Dim strREGSTR_PATH_UNINSTALL As String
    strREGSTR_PATH_UNINSTALL = RegPathWinCurrentVersion() & "\Uninstall"
    
    'The command-line for the application removal executable is simply the path
    'for the installation logfile
    Dim strAppRemovalCmdLine As String
    strAppRemovalCmdLine = GetAppRemovalCmdLine(strAppRemovalEXE, strAppRemovalLog, vbNullString, False, APPREMERR_NONE)
    '
    ' Make sure that the Removal command line (including path, filename, commandline args, etc.
    ' is not longer than the max allowed, which is _MAX_PATH.
    '
    If Not fCheckFNLength(strAppRemovalCmdLine) Then
        Dim strMsg As String
        strMsg = ResolveResString(resCANTCREATEICONPATHTOOLONG) & vbLf & vbLf & ResolveResString(resCHOOSENEWDEST) & vbLf & vbLf & strAppRemovalCmdLine
        Call MsgError(strMsg, vbOKOnly, gstrSETMSG)
        ExitSetup frmCopy, gintRET_FATAL
        Exit Function
    End If
    '
    ' Create registry entries to tell Windows where the app removal executable is,
    ' how it should be displayed to the user, and what the command-line arguments are
    '
    Dim iAppend As Integer
    Dim fOk As Boolean
    Dim hkeyAppRemoval As Long
    Dim hkeyOurs As Long
    Dim i As Integer
    
    'Go ahead and create a key to the main Uninstall branch
    If Not RegCreateKey(HKEY_LOCAL_MACHINE, strREGSTR_PATH_UNINSTALL, "", hkeyAppRemoval) Then
        GoTo Err
    End If
    
    'We need a unique key.  This key is never shown to the end user.  We will use a key of
    'the form 'ST5UNST #xxx'
    Dim strAppRemovalKey As String
    Dim strAppRemovalKeyBase As String
    Dim hkeyTest As Long
    strAppRemovalKeyBase = mstrFILE_APPREMOVALLOGBASE$ & " #"
    iAppend = 1
    
    Do
        strAppRemovalKey = strAppRemovalKeyBase & Format(iAppend)
        If RegOpenKey(hkeyAppRemoval, strAppRemovalKey, hkeyTest) Then
            'This key already exists.  But we need a unique key.
            RegCloseKey hkeyTest
        Else
            'We've found a key that doesn't already exist.  Use it.
            Exit Do
        End If
        
        iAppend = iAppend + 1
    Loop
    
    '
    ' We also need a unique displayname.  This name is
    ' the only means the user has to identify the application
    ' to remove
    '
    Dim strDisplayName As String
    strDisplayName = gstrAppName 'First try... Application name
    If Not IsDisplayNameUnique(hkeyAppRemoval, strDisplayName) Then
        'Second try... Add path
        strDisplayName = strDisplayName & " (" & gstrDestDir & ")"
        If Not IsDisplayNameUnique(hkeyAppRemoval, strDisplayName) Then
            'Subsequent tries... Append a unique integer
            Dim strDisplayNameBase As String
            
            strDisplayNameBase = strDisplayName
            iAppend = 3
            Do
                strDisplayName = strDisplayNameBase & " #" & Format(iAppend)
                If IsDisplayNameUnique(hkeyAppRemoval, strDisplayName) Then
                    Exit Do
                Else
                    iAppend = iAppend + 1
                End If
            Loop
        End If
    End If
    
    'Go ahead and fill in entries for the app removal executable
    If Not RegCreateKey(hkeyAppRemoval, strAppRemovalKey, "", hkeyOurs) Then
        GoTo Err
    End If
    If Not RegSetStringValue(hkeyOurs, strREGSTR_VAL_AppRemoval_APPNAMELINE, gstrAppExe, False) Then
        GoTo Err
    End If
    If Not RegSetStringValue(hkeyOurs, strREGSTR_VAL_AppRemoval_DISPLAYNAME, strDisplayName, False) Then
        GoTo Err
    End If
    If Not RegSetStringValue(hkeyOurs, strREGSTR_VAL_AppRemoval_COMMANDLINE, strAppRemovalCmdLine, False) Then
        GoTo Err
    End If
    If gstrAppToUninstall = vbNullString Then gstrAppToUninstall = gstrAppExe
    If Not RegSetStringValue(hkeyOurs, strREGSTR_VAL_AppRemoval_APPTOUNINSTALL, gstrAppToUninstall, False) Then
        GoTo Err
    End If
    If Not TreatAsWin95() Then
        '
        ' Under NT3.51, we simply place an icon to the app removal EXE in the program manager
        '
        If fMainGroupWasCreated Then
            CreateProgManItem frmSetup1, strGroupName, strAppRemovalCmdLine, ResolveResString(resAPPREMOVALICONNAME, "|1", gstrAppName)
        Else
            'If you get this message, it means that you incorrectly customized Form_Load().
            'Under 32-bits and NT 3.51, a Program Manager group must always be created.
            MsgError ResolveResString(resNOFOLDERFORICON, "|1", strAppRemovalEXE), vbOKOnly Or vbExclamation, gstrTitle
            ExitSetup frmSetup1, gintRET_FATAL
        End If
    End If
    
    RegCloseKey hkeyAppRemoval
    RegCloseKey hkeyOurs
    
    RegisterAppRemovalEXE = True
    Exit Function
    
Err:
    If hkeyOurs Then
        RegCloseKey hkeyOurs
        RegDeleteKey hkeyAppRemoval, strAppRemovalKey
    End If
    If hkeyAppRemoval Then
        RegCloseKey hkeyAppRemoval
    End If
    
    RegisterAppRemovalEXE = False
    Exit Function
End Function

'-----------------------------------------------------------
' FUNCTION: RegOpenKey
'
' Opens an existing key in the system registry.
'
' Returns: True if the key was opened OK, False otherwise
'   Upon success, phkResult is set to the handle of the key.
'-----------------------------------------------------------
'
Function RegOpenKey(ByVal hKey As Long, ByVal lpszSubKey As String, phkResult As Long) As Boolean
    Dim lResult As Long
    Dim strHkey As String

    On Error GoTo 0

    strHkey = strGetHKEYString(hKey)

    lResult = OSRegOpenKey(hKey, lpszSubKey, phkResult)
    If lResult = ERROR_SUCCESS Then
        RegOpenKey = True
        AddHkeyToCache phkResult, strHkey & "\" & lpszSubKey
    Else
        RegOpenKey = False
    End If
End Function
'----------------------------------------------------------
' FUNCTION: RegPathWinPrograms
'
' Returns the name of the registry key
' "\HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders"
'----------------------------------------------------------
Function RegPathWinPrograms() As String
    RegPathWinPrograms = RegPathWinCurrentVersion() & "\Explorer\Shell Folders"
End Function
 
'----------------------------------------------------------
' FUNCTION: RegPathWinCurrentVersion
'
' Returns the name of the registry key
' "\HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion"
'----------------------------------------------------------
Function RegPathWinCurrentVersion() As String
    RegPathWinCurrentVersion = "SOFTWARE\Microsoft\Windows\CurrentVersion"
End Function

'----------------------------------------------------------
' FUNCTION: RegQueryIntValue
'
' Retrieves the integer data for a named
' (strValueName = name) or unnamed (strValueName = "")
' value within a registry key.  If the named value
' exists, but its data is not a REG_DWORD, this function
' fails.
'
' NOTE: There is no 16-bit version of this function.
'
' Returns: True on success, else False.
'   On success, lData is set to the numeric data value
'
'----------------------------------------------------------
Function RegQueryNumericValue(ByVal hKey As Long, ByVal strValueName As String, lData As Long) As Boolean
    Dim lResult As Long
    Dim lValueType As Long
    Dim lBuf As Long
    Dim lDataBufSize As Long
    
    RegQueryNumericValue = False
    
    On Error GoTo 0
    
    ' Get length/data type
    lDataBufSize = 4
        
    lResult = OSRegQueryValueEx(hKey, strValueName, 0&, lValueType, lBuf, lDataBufSize)
    If lResult = ERROR_SUCCESS Then
        If lValueType = REG_DWORD Then
            lData = lBuf
            RegQueryNumericValue = True
        End If
    End If
End Function

' FUNCTION: RegQueryStringValue
'
' Retrieves the string data for a named
' (strValueName = name) or unnamed (strValueName = "")
' value within a registry key.  If the named value
' exists, but its data is not a string, this function
' fails.
'
' NOTE: For 16-bits, strValueName MUST be "" (but the
' NOTE: parameter is left in for source code compatability)
'
' Returns: True on success, else False.
'   On success, strData is set to the string data value
'
Function RegQueryStringValue(ByVal hKey As Long, ByVal strValueName As String, strData As String) As Boolean
    Dim lResult As Long
    Dim lValueType As Long
    Dim strBuf As String
    Dim lDataBufSize As Long
    
    RegQueryStringValue = False
    On Error GoTo 0
    ' Get length/data type
    lResult = OSRegQueryValueEx(hKey, strValueName, 0&, lValueType, ByVal 0&, lDataBufSize)
    If lResult = ERROR_SUCCESS Then
        If lValueType = REG_SZ Then
            strBuf = String(lDataBufSize, " ")
            lResult = OSRegQueryValueEx(hKey, strValueName, 0&, 0&, ByVal strBuf, lDataBufSize)
            If lResult = ERROR_SUCCESS Then
                RegQueryStringValue = True
                strData = StripTerminator(strBuf)
            End If
        End If
    End If
End Function

'----------------------------------------------------------
' FUNCTION: RegQueryRefCount
'
' Retrieves the data inteded as a reference count for a
' particular value within a registry key.  Although
' REG_DWORD is the preferred way of storing reference
' counts, it is possible that some installation programs
' may incorrect use a string or binary value instead.
' This routine accepts the data whether it is a string,
' a binary value or a DWORD (Long).
'
' NOTE: There is no 16-bit version of this function.
'
' Returns: True on success, else False.
'   On success, lrefcount is set to the numeric data value
'
'----------------------------------------------------------
Function RegQueryRefCount(ByVal hKey As Long, ByVal strValueName As String, lRefCount As Long) As Boolean
    Dim lResult As Long
    Dim lValueType As Long
    Dim lBuf As Long
    Dim lDataBufSize As Long

    RegQueryRefCount = False

    On Error GoTo 0

    ' Get length/data type
    lDataBufSize = 4

    lResult = OSRegQueryValueEx(hKey, strValueName, 0&, lValueType, lBuf, lDataBufSize)
    If lResult = ERROR_SUCCESS Then
        Select Case lValueType
            Case REG_DWORD
                lRefCount = lBuf
                RegQueryRefCount = True
            Case REG_BINARY
                If lDataBufSize = 4 Then
                    lRefCount = lBuf
                    RegQueryRefCount = True
                End If
            Case REG_SZ
                Dim strRefCount As String
                
                If RegQueryStringValue(hKey, strValueName, strRefCount) Then
                    lRefCount = Val(strRefCount)
                    RegQueryRefCount = True
                End If
            'End Case
        End Select
    End If
End Function

' FUNCTION: RegSetNumericValue
'
' Associates a named (strValueName = name) or unnamed (strValueName = "")
'   value with a registry key.
'
' If fLog is missing or is True, then this action is logged in the logfile,
' and the value will be deleted by the application removal utility if the
' user choose to remove the installed application.
'
' NOTE: There is no 16-bit version of this function.
'
' Returns: True on success, else False.
'
Function RegSetNumericValue(ByVal hKey As Long, ByVal strValueName As String, ByVal lData As Long, Optional ByVal fLog) As Boolean
    Dim lResult As Long
    Dim strHkey As String

    On Error GoTo 0
    
    If IsMissing(fLog) Then fLog = True

    strHkey = strGetHKEYString(hKey)
    
    If fLog Then
        NewAction _
          gstrKEY_REGVALUE, _
          """" & strHkey & """" _
            & ", " & """" & strValueName & """"
    End If

    lResult = OSRegSetValueEx(hKey, strValueName, 0&, REG_DWORD, lData, 4)
    If lResult = ERROR_SUCCESS Then
        RegSetNumericValue = True
        If fLog Then
            CommitAction
        End If
    Else
        RegSetNumericValue = False
        MsgError ResolveResString(resERR_REG), vbOKOnly Or vbExclamation, gstrTitle
        If fLog Then
            AbortAction
        End If
        If gfNoUserInput Then
            ExitSetup frmSetup1, gintRET_FATAL
        End If
    End If
End Function

' FUNCTION: RegSetStringValue
'
' Associates a named (strValueName = name) or unnamed (strValueName = "")
'   value with a registry key.
'
' If fLog is missing or is True, then this action is logged in the
' logfile, and the value will be deleted by the application removal
' utility if the user choose to remove the installed application.
'
' Returns: True on success, else False.
'
Function RegSetStringValue(ByVal hKey As Long, ByVal strValueName As String, ByVal strData As String, Optional ByVal fLog) As Boolean
    Dim lResult As Long
    Dim strHkey As String
    
    On Error GoTo 0
    
    If IsMissing(fLog) Then fLog = True

    If hKey = 0 Then
        Exit Function
    End If
    
    strHkey = strGetHKEYString(hKey)

    If fLog Then
        NewAction _
          gstrKEY_REGVALUE, _
          """" & strHkey & """" _
            & ", " & """" & strValueName & """"
    End If

    'lResult = OSRegSetValueEx(hKey, strValueName, 0&, REG_SZ, ByVal strData, LenB(StrConv(strData, vbFromUnicode)) + 1)
    lResult = OSRegSetValueEx(hKey, strValueName, 0&, REG_SZ, ByVal strData, Len(strData) + 1)
    
    If lResult = ERROR_SUCCESS Then
        RegSetStringValue = True
        If fLog Then
            CommitAction
        End If
    Else
        RegSetStringValue = False
        MsgError ResolveResString(resERR_REG), vbOKOnly Or vbExclamation, gstrTitle
        If fLog Then
            AbortAction
        End If
        If gfNoUserInput Then
            ExitSetup frmSetup1, gintRET_FATAL
        End If
    End If
End Function

'-----------------------------------------------------------
' SUB: RemoteRegister
'
' Synchronously run the client registration utility on the
' given remote server registration file in order to set it
' up properly in the registry.
'
' IN: [strFileName] - .EXE file to register

'-----------------------------------------------------------
'
Sub RemoteRegister(ByVal strFilename As String, rInfo As REGINFO)
    Const strClientRegistrationUtility$ = "CLIREG32.EXE"
    Const strAddressSwitch = " /s "
    Const strProtocolSwitch = " /p "
    Const strSilentSwitch = " /q "
    Const strNoLogoSwitch = " /nologo "
    Const strAuthenticationSwitch = " /a "
    Const strTypelibSwitch = " /t "
    Const strDCOMSwitch = " /d "
    Const strEXT_REMOTE$ = "VBR"
    Const strEXT_REMOTETLB$ = "TLB"

    Dim strAddress As String
    Dim strProtocol As String
    Dim intAuthentication As Integer
    Dim strCmdLine As String
    Dim fShell As Integer
    Dim strMatchingTLB As String
    Dim fDCOM As Boolean

    'Find the name of the matching typelib file.  This should have already
    'been installed to the same directory as the .VBR file.
    strMatchingTLB = strFilename
    If Right$(strMatchingTLB, Len(strEXT_REMOTE)) = strEXT_REMOTE Then
        strMatchingTLB = Left$(strMatchingTLB, Len(strMatchingTLB) - Len(strEXT_REMOTE))
    End If
    strMatchingTLB = strMatchingTLB & strEXT_REMOTETLB

    strAddress = rInfo.strNetworkAddress
    strProtocol = rInfo.strNetworkProtocol
    intAuthentication = rInfo.intAuthentication
    fDCOM = rInfo.fDCOM
    frmRemoteServerDetails.GetServerDetails strFilename, strAddress, strProtocol, fDCOM
    frmMessage.Refresh
    strCmdLine = _
      strClientRegistrationUtility _
      & strAddressSwitch & """" & strAddress & """" _
      & IIf(fDCOM, " ", strProtocolSwitch & strProtocol) _
      & IIf(fDCOM, " ", strAuthenticationSwitch & Format$(intAuthentication) & " ") _
      & strNoLogoSwitch _
      & strTypelibSwitch & """" & strMatchingTLB & """" & " " _
      & IIf(fDCOM, strDCOMSwitch, "") _
      & IIf(gfNoUserInput, strSilentSwitch, "") _
      & """" & strFilename & """"
      
    '
    'Synchronously shell out and run the utility with the correct switches
    '
    fShell = SyncShell(strCmdLine, INFINITE, , False)

    If Not fShell Then
        MsgError ResolveResString(resCANTRUNPROGRAM, "|1", strClientRegistrationUtility), vbOKOnly Or vbExclamation, gstrTitle, gintRET_FATAL
        ExitSetup frmSetup1, gintRET_FATAL
    End If
End Sub

'-----------------------------------------------------------
' SUB: RemoveShellLink
'
' Removes a link in either Start>Programs or any of its

' immediate subfolders in the Windows 95 shell.
'
' IN: [strFolderName] - text name of the immediate folder
'                       in which the link to be removed
'                       currently exists, or else the
'                       empty string ("") to indicate that
'                       the link can be found directly in
'                       the Start>Programs menu.
'     [strLinkName] - text caption for the link
'
' This action is never logged in the app removal logfile.
'
' PRECONDITION: strFolderName has already been created and is
'               an immediate subfolder of Start>Programs, if it
'               is not equal to ""
'-----------------------------------------------------------
'
Sub RemoveShellLink(ByVal strFolderName As String, ByVal strLinkName As String)
    Dim fSuccess As Boolean
    
    ReplaceDoubleQuotes strFolderName
    ReplaceDoubleQuotes strLinkName
    
    fSuccess = OSfRemoveShellLink(strFolderName, strLinkName)
End Sub

'-----------------------------------------------------------
' FUNCTION: ResolveDestDir
'
' Given a destination directory string, equate any macro
' portions of the string to their runtime determined
' actual locations and return a string reflecting the
' actual path.
'
' IN: [strDestDir] - string containing directory macro info
'                    and/or actual dir path info
'
'     [fAssumeDir] - boolean that if true, causes this routine
'                    to assume that strDestDir contains a dir
'                    path.  If a directory isn't given it will
'                    make it the application path.  If false,
'                    this routine will return strDestDir as
'                    is after performing expansion.  Set this
'                    to False when you are not sure it is a
'                    directory but you want to expand macros
'                    if it contains any.  E.g., If this is a
'                    command line parameter, you can't be
'                    certain if it refers to a path.  In this
'                    case, set fAssumeDir = False.  Default
'                    is True.
'
' Return: A string containing the resolved dir name
'-----------------------------------------------------------
'
Function ResolveDestDir(ByVal strDestDir As String, Optional fAssumeDir As Variant) As String
    Const strMACROSTART$ = "$("
    Const strMACROEND$ = ")"

    Dim intPos As Integer
    Dim strResolved As String
    Dim hKey As Long
    Dim strPathsKey As String
    Dim fQuoted As Boolean
    
    If IsMissing(fAssumeDir) Then
        fAssumeDir = True
    End If
    
    strPathsKey = RegPathWinCurrentVersion()
    strDestDir = Trim(strDestDir)
    '
    ' If strDestDir is quoted when passed to this routine, it
    ' should be quoted when it's returned.  The quotes need
    ' to be temporarily removed, though, for processing.
    '
    If Left(strDestDir, 1) = gstrQUOTE Then
        fQuoted = True
        strDestDir = strUnQuoteString(strDestDir)
    End If
    '
    ' We take the first part of destdir, and if its $( then we need to get the portion
    ' of destdir up to and including the last paren.  We then test against this for
    ' macro expansion.  If no ) is found after finding $(, then must assume that it's
    ' just a normal file name and do no processing.  Only enter the case statement
    ' if strDestDir starts with $(.
    '
    If Left$(strDestDir, 2) = strMACROSTART Then
        intPos = InStr(strDestDir, strMACROEND)

        Select Case Left$(strDestDir, intPos)
            Case gstrAPPDEST
                If gstrDestDir <> vbNullString Then

                    strResolved = gstrDestDir
                Else
                    strResolved = "?"
                End If
            Case gstrWINDEST
                strResolved = gstrWinDir
            Case gstrWINSYSDEST, gstrWINSYSDESTSYSFILE
                strResolved = gstrWinSysDir
            Case gstrPROGRAMFILES
                If TreatAsWin95() Then
                    Const strProgramFilesKey = "ProgramFilesDir"
    
                    If RegOpenKey(HKEY_LOCAL_MACHINE, strPathsKey, hKey) Then
                        RegQueryStringValue hKey, strProgramFilesKey, strResolved
                        RegCloseKey hKey
                    End If
                End If
    
                If strResolved = "" Then
                    'If not otherwise set, let strResolved be the root of the first fixed disk
                    strResolved = strRootDrive()
                End If
            Case gstrCOMMONFILES
                'First determine the correct path of Program Files\Common Files, if under Win95
                strResolved = strGetCommonFilesPath()
                If strResolved = "" Then
                    'If not otherwise set, let strResolved be the Windows directory
                    strResolved = gstrWinDir
                End If
            Case gstrCOMMONFILESSYS
                'First determine the correct path of Program Files\Common Files, if under Win95
                Dim strCommonFiles As String
                
                strCommonFiles = strGetCommonFilesPath()
                If strCommonFiles <> "" Then
                    'Okay, now just add \System, and we're done
                    strResolved = strCommonFiles & "System\"
                Else
                    'If Common Files isn't in the registry, then map the
                    'entire macro to the Windows\{system,system32} directory
                    strResolved = gstrWinSysDir
                End If
            Case gstrDAODEST
                strResolved = strGetDAOPath()
            Case Else
                intPos = 0
            'End Case
        End Select
    End If
    
    If intPos <> 0 Then
        AddDirSep strResolved
    End If

    If fAssumeDir = True Then
        If intPos = 0 Then
            '
            'if no drive spec, and doesn't begin with any root path indicator ("\"),
            'then we assume that this destination is relative to the app dest dir
            '
            If Mid$(strDestDir, 2, 1) <> gstrCOLON Then
                If Left$(strDestDir, 1) <> gstrSEP_DIR Then
                    strResolved = gstrDestDir
                End If
            End If
        Else
            If Mid$(strDestDir, intPos + 1, 1) = gstrSEP_DIR Then
                intPos = intPos + 1
            End If
        End If
    End If

    If fQuoted = True Then
        ResolveDestDir = strQuoteString(strResolved & Mid$(strDestDir, intPos + 1), True, False)
    Else
        ResolveDestDir = strResolved & Mid$(strDestDir, intPos + 1)
    End If
End Function
'-----------------------------------------------------------
' FUNCTION: ResolveDestDirs
'
' Given a space delimited string, this routine finds all
' Destination directory macros and expands them by making
' repeated calls to ResolveDestDir.  See ResolveDestDir.
'
' Note that the macro must immediately follow a space (or
' a space followed by a quote) delimiter or else it will
' be ignored.
'
' Note that this routine does not assume that each item
' in the delimited string is actually a directory path.
' Therefore, the last parameter in the call to ResolveDestDir,
' below, is false.
'
' IN: [str] - string containing directory macro(s) info
'             and/or actual dir path info
'
' Return: str with destdir macros expanded.
'-----------------------------------------------------------
'
Function ResolveDestDirs(str As String)
    Dim intAnchor As Integer
    Dim intOffset As Integer
    Dim strField As String
    Dim strExpField As String
    Dim strExpanded As String
    
    If Len(Trim(strUnQuoteString(str))) = 0 Then
        ResolveDestDirs = str
        Exit Function
    End If
        
    intAnchor = 1
    strExpanded = ""
    
    Do
        intOffset = intGetNextFldOffset(intAnchor, str, " ")
        If intOffset = 0 Then intOffset = Len(str) + 1
        strField = Mid(str, intAnchor, intOffset - intAnchor)
        strExpField = ResolveDestDir(strField, False)
        strExpanded = strExpanded & strExpField & " "
        intAnchor = intOffset + 1
    Loop While intAnchor < Len(str)
    
    ResolveDestDirs = Trim(strExpanded)
End Function
'-----------------------------------------------------------
' FUNCTION: ResolveDir
'
' Given a pathname, resolve it to its smallest form.  If
' the pathname is invalid, then optionally warn the user.
'
' IN: [strPathName] - pathname to resolve
'     [fMustExist] - enforce that the path actually exists
'     [fWarn] - If True, warn user upon invalid path
'
' Return: A string containing the resolved dir name
'-----------------------------------------------------------
'
Function ResolveDir(ByVal strPathName As String, fMustExist As Integer, fWarn As Integer) As String
    Dim strMsg As String
    Dim fInValid As Integer
    Dim strUnResolvedPath As String
    Dim strResolvedPath As String
    Dim strIgnore As String
    Dim cbResolved As Long

    On Error Resume Next

    fInValid = False
    '
    'If the pathname is a UNC name (16-bit only), or if it's in actuality a file name, then it's invalid
    '
    If FileExists(strPathName) = True Then
        fInValid = True
        GoTo RDContinue
    End If

    strUnResolvedPath = strPathName

    If InStr(3, strUnResolvedPath, gstrSEP_DIR) > 0 Then

        strResolvedPath = Space(gintMAX_PATH_LEN * 2)
        cbResolved = GetFullPathName(strUnResolvedPath, gintMAX_PATH_LEN, strResolvedPath, strIgnore)
        If cbResolved = 0 Then
            '
            ' The path couldn't be resolved.  If we can actually
            ' switch to the directory we want, continue anyway.
            '
            ChDir strUnResolvedPath
            AddDirSep strUnResolvedPath
            If Err > 0 Then
                Err = 0
                ChDir strUnResolvedPath
                If Err > 0 Then
                    fInValid = True
                Else
                    strResolvedPath = strUnResolvedPath
                End If
            Else
                strResolvedPath = strUnResolvedPath
            End If
        Else
            '
            ' GetFullPathName returned us a NULL terminated string in
            ' strResolvedPath.  Remove the NULL.
            '
            strResolvedPath = StripTerminator(strResolvedPath)
            If CheckDrive(strResolvedPath, gstrTitle) = False Then
                fInValid = True
            Else
                AddDirSep strResolvedPath
                If fMustExist = True Then
                    Err = 0
                    
                    Dim strDummy As String
                    strDummy = Dir$(strResolvedPath & "*.*")
                    
                    If Err > 0 Then
                        strMsg = ResolveResString(resNOTEXIST) & vbLf & vbLf
                        fInValid = True
                    End If
                End If
            End If
        End If
    Else
        fInValid = True
    End If

RDContinue:
    If fInValid = True Then
        If fWarn = True Then
            strMsg = strMsg & ResolveResString(resDIRSPECIFIED) & vbLf & vbLf & strPathName & vbLf & vbLf
            strMsg = strMsg & ResolveResString(resDIRINVALID)
            MsgError strMsg, vbOKOnly Or vbExclamation, ResolveResString(resDIRINVNAME)
            If gfNoUserInput Then
                ExitSetup frmSetup1, gintRET_FATAL
            End If
        End If

        ResolveDir = vbNullString
    Else
        ResolveDir = strResolvedPath
    End If

    Err = 0
End Function

'-----------------------------------------------------------
' SUB: RestoreProgMan
'
' Restores Windows Program Manager
'-----------------------------------------------------------
'
Sub RestoreProgMan()
    Const strPMTITLE$ = "Program Manager"

    On Error Resume Next

    'Try the localized name first
    AppActivate ResolveResString(resPROGRAMMANAGER)
    
    If Err Then
        'If that doesn't work, try the English name
        AppActivate strPMTITLE
    End If

    Err = 0
End Sub

'-----------------------------------------------------------
' SUB: ShowPathDialog
'
' Display form to allow user to get either a source or
' destination path
'
' IN: [strPathRequest] - determines whether to ask for the
'                        source or destination pathname.
'                        gstrDIR_SRC for source path
'                        gstrDIR_DEST for destination path
'-----------------------------------------------------------
'
Sub ShowPathDialog(ByVal strPathRequest As String)
    frmSetup1.Tag = strPathRequest

    '
    'frmPath.Form_Load() reads frmSetup1.Tag to determine whether
    'this is a request for the source or destination path
    '
    frmPath.Show vbModal

    If strPathRequest = gstrDIR_SRC Then
        gstrSrcPath = frmSetup1.Tag
    Else
        If gfRetVal = gintRET_CONT Then
            gstrDestDir = frmSetup1.Tag
        End If
    End If
End Sub

'-----------------------------------------------------------
' FUNCTION: strExtractFilenameArg
'
' Extracts a quoted or unquoted filename from a string
'   containing command-line arguments
'
' IN: [str] - string containing a filename.  This filename
'             begins at the first character, and continues
'             to the end of the string or to the first space
'             or switch character, or, if the string begins
'             with a double quote, continues until the next
'             double quote
' OUT: Returns the filename, without quotes
'      str is set to be the remainder of the string after
'      the filename and quote (if any)
'
'-----------------------------------------------------------
'
Function strExtractFilenameArg(str As String, fErr As Boolean)
    Dim strFilename As String
    
    str = Trim$(str)
    
    Dim iEndFilenamePos As Integer
    If Left$(str, 1) = """" Then
        ' Filenames is surrounded by quotes
        iEndFilenamePos = InStr(2, str, """") ' Find matching quote
        If iEndFilenamePos > 0 Then
            strFilename = Mid$(str, 2, iEndFilenamePos - 2)
            str = Right$(str, Len(str) - iEndFilenamePos)
        Else
            fErr = True
            Exit Function
        End If
    Else
        ' Filename continues until next switch or space or quote
        Dim iSpacePos As Integer
        Dim iSwitch1 As Integer
        Dim iSwitch2 As Integer
        Dim iQuote As Integer
        
        iSpacePos = InStr(str, " ")
        iSwitch2 = InStr(str, gstrSwitchPrefix2)
        iQuote = InStr(str, """")
        
        If iSpacePos = 0 Then iSpacePos = Len(str) + 1
        If iSwitch1 = 0 Then iSwitch1 = Len(str) + 1
        If iSwitch2 = 0 Then iSwitch2 = Len(str) + 1
        If iQuote = 0 Then iQuote = Len(str) + 1
        
        iEndFilenamePos = iSpacePos
        If iSwitch2 < iEndFilenamePos Then iEndFilenamePos = iSwitch2
        If iQuote < iEndFilenamePos Then iEndFilenamePos = iQuote
        
        strFilename = Left$(str, iEndFilenamePos - 1)
        If iEndFilenamePos > Len(str) Then
            str = ""
        Else
            str = Right(str, Len(str) - iEndFilenamePos + 1)
        End If
    End If
    
    strFilename = Trim$(strFilename)
    If strFilename = "" Then
        fErr = True
        Exit Function
    End If
    
    fErr = False
    strExtractFilenameArg = strFilename
    str = Trim$(str)
End Function



'-----------------------------------------------------------
' SUB: UpdateStatus
'
' "Fill" (by percentage) inside the PictureBox and also
' display the percentage filled
'
' IN: [pic] - PictureBox used to bound "fill" region
'     [sngPercent] - Percentage of the shape to fill
'     [fBorderCase] - Indicates whether the percentage
'        specified is a "border case", i.e. exactly 0%
'        or exactly 100%.  Unless fBorderCase is True,
'        the values 0% and 100% will be assumed to be
'        "close" to these values, and 1% and 99% will
'        be used instead.
'
' Notes: Set AutoRedraw property of the PictureBox to True
'        so that the status bar and percentage can be auto-
'        matically repainted if necessary
'-----------------------------------------------------------
'
Sub UpdateStatus(pic As PictureBox, ByVal sngPercent As Single, Optional ByVal fBorderCase)
    Dim strPercent As String
    Dim intX As Integer
    Dim intY As Integer
    Dim intWidth As Integer
    Dim intHeight As Integer

    If IsMissing(fBorderCase) Then fBorderCase = False
    
    'For this to work well, we need a white background and any color foreground (blue)
    Const colBackground = &HFFFFFF ' white
    Const colForeground = &H800000 ' dark blue

    pic.ForeColor = colForeground
    pic.BackColor = colBackground
    
    '
    'Format percentage and get attributes of text
    '
    Dim intPercent
    intPercent = Int(100 * sngPercent + 0.5)
    
    'Never allow the percentage to be 0 or 100 unless it is exactly that value.  This
    'prevents, for instance, the status bar from reaching 100% until we are entirely done.
    If intPercent = 0 Then
        If Not fBorderCase Then
            intPercent = 1
        End If
    ElseIf intPercent = 100 Then
        If Not fBorderCase Then
            intPercent = 99
        End If
    End If
    
    strPercent = Format$(intPercent) & "%"
    intWidth = pic.TextWidth(strPercent)
    intHeight = pic.TextHeight(strPercent)

    '
    'Now set intX and intY to the starting location for printing the percentage
    '
    intX = pic.Width / 2 - intWidth / 2
    intY = pic.Height / 2 - intHeight / 2

    '
    'Need to draw a filled box with the pics background color to wipe out previous
    'percentage display (if any)
    '
    pic.DrawMode = 13 ' Copy Pen
    pic.Line (intX, intY)-Step(intWidth, intHeight), pic.BackColor, BF

    '
    'Back to the center print position and print the text
    '
    pic.CurrentX = intX
    pic.CurrentY = intY
    pic.Print strPercent

    '
    'Now fill in the box with the ribbon color to the desired percentage
    'If percentage is 0, fill the whole box with the background color to clear it
    'Use the "Not XOR" pen so that we change the color of the text to white
    'wherever we touch it, and change the color of the background to blue
    'wherever we touch it.
    '
    pic.DrawMode = 10 ' Not XOR Pen
    If sngPercent > 0 Then
        pic.Line (0, 0)-(pic.Width * sngPercent, pic.Height), pic.ForeColor, BF
    Else
        pic.Line (0, 0)-(pic.Width, pic.Height), pic.BackColor, BF
    End If

    pic.Refresh
End Sub

'-----------------------------------------------------------
' FUNCTION: WriteAccess
'
' Determines whether there is write access to the specified
' directory.
'
' IN: [strDirName] - directory to check for write access
'
' Returns: True if write access, False otherwise
'-----------------------------------------------------------
'
Function WriteAccess(ByVal strDirName As String) As Integer
    Dim intFileNum As Integer

    On Error Resume Next

    AddDirSep strDirName

    intFileNum = FreeFile
    Open strDirName & mstrCONCATFILE For Output As intFileNum

    WriteAccess = IIf(Err, False, True)
    
    Close intFileNum

    Kill strDirName & mstrCONCATFILE

    Err = 0
End Function
'-----------------------------------------------------------
' FUNCTION: WriteMIF
'
' If this is a SMS install, this routine writes the
' failed MIF status file if something goes wrong or
' a successful MIF if everything installs correctly.
'
' The MIF file requires a special format specified
' by SMS.  Currently, this routine implements the
' minimum requirements.  The hardcoded strings below
' that are written to the MIF should be written
' character by character as they are; except that
' status message should change depending on the
' circumstances of the install.  DO NOT LOCALIZE
' anything except the status message.
'
' IN: [strMIFFilename] - The name of the MIF file.
'                        Passed in to setup1 by
'                        setup.exe.  It is probably
'                        named <appname>.mif where
'                        <appname> is the name of the
'                        application you are installing.
'
'     [fStatus] - False to write a failed MIF (i.e. setup
'                 failed); True to write a successful MIF.
'
'     [strSMSDescription] - This is the description string
'                           to be written to the MIF file.
'                           It cannot be longer than 255
'                           characters and cannot contain
'                           carriage returns and/or line
'                           feeds.  This routine will
'                           enforce these requirements.
'
' Note, when running in SMS mode, there is no other way
' to display a message to the user than to write it to
' the MIF file.  Displaying a MsgBox will cause the
' computer to appear as if it has hung.  Therefore, this
' routine makes no attempt to display an error message.
'
'-----------------------------------------------------------
'
Sub WriteMIF(ByVal strMIFFilename As String, ByVal fStatus As Boolean, ByVal strSMSDescription As String)
    Const strSUCCESS = """SUCCESS"""                 ' Cannot be localized as per SMS
    Const strFAILED = """FAILED"""                   ' Cannot be localized as per SMS
    
    Dim fn As Integer
    Dim intOffset As Integer
    Dim fOpened As Boolean
        
    fOpened = False
        
    On Error GoTo WMIFFAILED  ' If we fail, we just return without doing anything
                              ' because there is no way to inform the user while
                              ' in SMS mode.

    '
    ' If the description string is greater than 255 characters,
    ' truncate it.  Required my SMS.
    '
    strSMSDescription = Left(strSMSDescription, MAX_SMS_DESCRIP)
    '
    ' Remove any carriage returns or line feeds and replace
    ' them with spaces.  The message must be a single line.
    '
    For intOffset = 1 To Len(strSMSDescription)
        If (Mid(strSMSDescription, intOffset, 1) = Chr(10)) Or (Mid(strSMSDescription, intOffset, 1) = Chr(13)) Then
            Mid(strSMSDescription, intOffset, 1) = " "
        End If
    Next intOffset
    '
    ' Open the MIF file for append, but first delete any existing
    ' ones with the same name.  Note, that setup.exe passed a
    ' unique name so if there is one with this name already in
    ' on the disk, it was put there by setup.exe.
    '
    If FileExists(strMIFFilename) Then
        Kill strMIFFilename
    End If
    
    fn = FreeFile
    Open strMIFFilename For Append As fn
    fOpened = True
    '
    ' We are ready to write the actual MIF file
    ' Note, none of the string below are supposed
    ' to be localized.
    '
    Print #fn, "Start Component"
        Print #fn, Tab; "Name = ""Workstation"""
        Print #fn, Tab; "Start Group"
            Print #fn, Tab; Tab; "Name = ""InstallStatus"""
            Print #fn, Tab; Tab; "ID = 1"
            Print #fn, Tab; Tab; "Class = ""MICROSOFT|JOBSTATUS|1.0"""
            Print #fn, Tab; Tab; "Start Attribute"
                Print #fn, Tab; Tab; Tab; "Name = ""Status"""
                Print #fn, Tab; Tab; Tab; "ID = 1"
                Print #fn, Tab; Tab; Tab; "Type = String(16)"
                Print #fn, Tab; Tab; Tab; "Value = "; IIf(fStatus, strSUCCESS, strFAILED)
            Print #fn, Tab; Tab; "End Attribute"
            Print #fn, Tab; Tab; "Start Attribute"
                Print #fn, Tab; Tab; Tab; "Name = ""Description"""
                Print #fn, Tab; Tab; Tab; "ID = 2"
                Print #fn, Tab; Tab; Tab; "Type = String(256)"
                Print #fn, Tab; Tab; Tab; "Value = "; strSMSDescription
            Print #fn, Tab; Tab; "End Attribute"
        Print #fn, Tab; "End Group"
    Print #fn, "End Component"

    Close fn
    '
    ' Success
    '
    Exit Sub

WMIFFAILED:
    '
    ' At this point we are unable to create the MIF file.
    ' Since we are running under SMS there is no one to
    ' tell, so we don't generate an error message at all.
    '
    If fOpened = True Then
        Close fn
    End If
    Exit Sub
End Sub

'Adds or replaces an HKEY to the list of HKEYs in cache.
'Note that it is not necessary to remove keys from
'this list.
Private Sub AddHkeyToCache(ByVal hKey As Long, ByVal strHkey As String)
    Dim intIdx As Integer
    
    intIdx = intGetHKEYIndex(hKey)
    If intIdx < 0 Then
        'The key does not already exist.  Add it to the end.
        On Error Resume Next
        ReDim Preserve hkeyCache(0 To UBound(hkeyCache) + 1)
        If Err Then
            'If there was an error, it means the cache was empty.
            On Error GoTo 0
            ReDim hkeyCache(0 To 0)
        End If
        On Error GoTo 0

        intIdx = UBound(hkeyCache)
    Else
        'The key already exists.  It will be replaced.
    End If

    hkeyCache(intIdx).hKey = hKey
    hkeyCache(intIdx).strHkey = strHkey
End Sub

'Given a predefined HKEY, return the text string representing that
'key, or else return "".
Private Function strGetPredefinedHKEYString(ByVal hKey As Long) As String
    Select Case hKey
        Case HKEY_CLASSES_ROOT
            strGetPredefinedHKEYString = "HKEY_CLASSES_ROOT"
        Case HKEY_CURRENT_USER
            strGetPredefinedHKEYString = "HKEY_CURRENT_USER"
        Case HKEY_LOCAL_MACHINE
            strGetPredefinedHKEYString = "HKEY_LOCAL_MACHINE"
        Case HKEY_USERS
            strGetPredefinedHKEYString = "HKEY_USERS"
        'End Case
    End Select
End Function

'Given an HKEY, return the text string representing that
'key.
Private Function strGetHKEYString(ByVal hKey As Long) As String
    Dim strKey As String

    'Is the hkey predefined?
    strKey = strGetPredefinedHKEYString(hKey)
    If strKey <> "" Then
        strGetHKEYString = strKey
        Exit Function
    End If
    
    'It is not predefined.  Look in the cache.
    Dim intIdx As Integer
    intIdx = intGetHKEYIndex(hKey)
    If intIdx >= 0 Then
        strGetHKEYString = hkeyCache(intIdx).strHkey
    Else
        strGetHKEYString = ""
    End If
End Function

'Searches the cache for the index of the given HKEY.
'Returns the index if found, else returns -1.
Private Function intGetHKEYIndex(ByVal hKey As Long) As Integer
    Dim intUBound As Integer
    
    On Error Resume Next
    intUBound = UBound(hkeyCache)
    If Err Then
        'If there was an error accessing the ubound of the array,
        'then the cache is empty
        GoTo NotFound
    End If
    On Error GoTo 0

    Dim intIdx As Integer
    For intIdx = 0 To intUBound
        If hkeyCache(intIdx).hKey = hKey Then
            intGetHKEYIndex = intIdx
            Exit Function
        End If
    Next intIdx
    
NotFound:
    intGetHKEYIndex = -1
End Function

'Returns the location of the Program Files\Common Files path, if
'it is present in the registry.  Otherwise, returns "".
Public Function strGetCommonFilesPath() As String
    Dim hKey As Long
    Dim strPath As String
    
    If TreatAsWin95() Then
        Const strCommonFilesKey = "CommonFilesDir"

        If RegOpenKey(HKEY_LOCAL_MACHINE, RegPathWinCurrentVersion(), hKey) Then
            RegQueryStringValue hKey, strCommonFilesKey, strPath
            RegCloseKey hKey
        End If
    End If

    If strPath <> "" Then
        AddDirSep strPath
    End If
    
    strGetCommonFilesPath = strPath
End Function
'Returns the location of the "Windows\Start Menu\Programs" Files path, if
'it is present in the registry.  Otherwise, returns "".
Public Function strGetProgramsFilesPath() As String
    Dim hKey As Long
    Dim strPath As String
    
    strPath = ""
    If TreatAsWin95() Then
        Const strProgramsKey = "Programs"

        If RegOpenKey(HKEY_CURRENT_USER, RegPathWinPrograms(), hKey) Then
            RegQueryStringValue hKey, strProgramsKey, strPath
            RegCloseKey hKey
        End If
    End If

    If strPath <> "" Then
        AddDirSep strPath
    End If
    
    strGetProgramsFilesPath = strPath
End Function

'Returns the directory where DAO is or should be installed.  If the
'key does not exist in the registry, it is created.  For instance, under
'NT 3.51 this location is normally 'C:\WINDOWS\MSAPPS\DAO'
Private Function strGetDAOPath() As String
    Const strMSAPPS$ = "MSAPPS\"
    Const strDAO3032$ = "DAO350.DLL"
    
    'first look in the registry
    Const strKey = "SOFTWARE\Microsoft\Shared Tools\DAO350"
    Const strValueName = "Path"
    Dim hKey As Long
    Dim strPath As String

    If RegOpenKey(HKEY_LOCAL_MACHINE, strKey, hKey) Then
        RegQueryStringValue hKey, strValueName, strPath
        RegCloseKey hKey
    End If

    If strPath <> "" Then
        strPath = GetPathName(strPath)
        AddDirSep strPath
        strGetDAOPath = strPath
        Exit Function
    End If
    
    'It's not yet in the registry, so we need to decide
    'where the directory should be, and then need to place
    'that location in the registry.

    If TreatAsWin95() Then
        'For Win95, use "Common Files\Microsoft Shared\DAO"
        strPath = strGetCommonFilesPath() & ResolveResString(resMICROSOFTSHARED) & "DAO\"
    Else
        'Otherwise use Windows\MSAPPS\DAO
        strPath = gstrWinDir & strMSAPPS & "DAO\"
    End If
    
    'Place this information in the registry (note that we point to DAO3032.DLL
    'itself, not just to the directory)
    If RegCreateKey(HKEY_LOCAL_MACHINE, strKey, "", hKey) Then
        RegSetStringValue hKey, strValueName, strPath & strDAO3032, False
        RegCloseKey hKey
    End If

    strGetDAOPath = strPath
End Function

' Replace all double quotes with single quotes
Public Sub ReplaceDoubleQuotes(str As String)
    Dim i As Integer
    
    For i = 1 To Len(str)
        If Mid$(str, i, 1) = """" Then
            Mid$(str, i, 1) = "'"
        End If
    Next i
End Sub

'Get the path portion of a filename
Function GetPathName(ByVal strFilename As String) As String
    Dim sPath As String
    Dim sFile As String
    
    SeparatePathAndFileName strFilename, sPath, sFile
    
    GetPathName = sPath
End Function
'Determines if a character is a path separator (\ or /).
Public Function IsSeparator(Character As String) As Boolean
    Select Case Character
        Case gstrSEP_DIR
            IsSeparator = True
        Case gstrSEP_DIRALT
            IsSeparator = True
    End Select
End Function
'Given a fully qualified filename, returns the path portion and the file
'   portion.
Public Sub SeparatePathAndFileName(FullPath As String, ByRef Path As String, _
    ByRef FileName As String)

    Dim nSepPos As Long
    Dim sSEP As String

    nSepPos = Len(FullPath)
    sSEP = Mid$(FullPath, nSepPos, 1)
    Do Until IsSeparator(sSEP)
        nSepPos = nSepPos - 1
        If nSepPos = 0 Then Exit Do
        sSEP = Mid$(FullPath, nSepPos, 1)
    Loop

    Select Case nSepPos
        Case 0
            'Separator was not found.
            Path = CurDir$
            FileName = FullPath
        Case Else
            Path = Left$(FullPath, nSepPos - 1)
            FileName = Mid$(FullPath, nSepPos + 1)
    End Select
End Sub

'Returns the path to the root of the first fixed disk
Function strRootDrive() As String
    Dim intDriveNum As Integer
    
    For intDriveNum = 0 To Asc("Z") - Asc("A") - 1
        If GetDriveType(intDriveNum) = intDRIVE_FIXED Then
            strRootDrive = Chr$(Asc("A") + intDriveNum) & gstrCOLON & gstrSEP_DIR
            Exit Function
        End If
    Next intDriveNum
    
    strRootDrive = "C:\"
End Function

'Returns "" if the path is not complete, or is a UNC pathname
Function strGetDriveFromPath(ByVal strPath As String) As String
    If Len(strPath) < 2 Then
        Exit Function
    End If
    
    If Mid$(strPath, 2, 1) <> gstrCOLON Then
        Exit Function
    End If
    
    strGetDriveFromPath = Mid$(strPath, 1, 1) & gstrCOLON & gstrSEP_DIR
End Function

Public Function fValidFilename(strFilename As String) As Boolean
'
' This routine verifies that strFileName is a valid file name.
' It checks that its length is less than the max allowed
' and that it doesn't contain any invalid characters..
'
    If Not fCheckFNLength(strFilename) Then
        '
        ' Name is too long.
        '
        fValidFilename = False
        Exit Function
    End If
    '
    ' Search through the list of invalid filename characters and make
    ' sure none of them are in the string.
    '
    Dim iInvalidChar As Integer
    Dim iFilename As Integer
    Dim strInvalidChars As String
    
    strInvalidChars = ResolveResString(resCOMMON_INVALIDFILECHARS)
    
    For iInvalidChar = 1 To Len(strInvalidChars)
        If InStr(strFilename, Mid$(strInvalidChars, iInvalidChar, 1)) <> 0 Then
            fValidFilename = False
            Exit Function
        End If
    Next iInvalidChar
    
    fValidFilename = True
    
End Function
Public Function fValidNTGroupName(strGroupName) As Boolean
'
' This routine verifies that strGroupName is a valid group name.
' It checks that its length is less than the max allowed
' and that it doesn't contain any invalid characters.
'
    If Len(strGroupName) > gintMAX_GROUPNAME_LEN Then
        fValidNTGroupName = False
        Exit Function
    End If
    '
    ' Search through the list of invalid filename characters and make
    ' sure none of them are in the string.
    '
    Dim iInvalidChar As Integer
    Dim iFilename As Integer
    Dim strInvalidChars As String
    
    strInvalidChars = ResolveResString(resGROUPINVALIDCHARS)
    
    For iInvalidChar = 1 To Len(strInvalidChars)
        If InStr(strGroupName, Mid$(strInvalidChars, iInvalidChar, 1)) <> 0 Then
            fValidNTGroupName = False
            Exit Function
        End If
    Next iInvalidChar
    
    fValidNTGroupName = True
    
End Function
'-----------------------------------------------------------
' SUB: CountGroups
'
' Determines how many groups must be installed by counting
' them in the setup information file (SETUP.LST)
'-----------------------------------------------------------
'
Function CountGroups(ByVal strsection As String) As Integer
    Dim intIdx As Integer
    Dim sGroup As String
    
    intIdx = 0
    Do
        sGroup = ReadIniFile(gstrSetupInfoFile, strsection, gsGROUP & CStr(intIdx))
        If sGroup <> vbNullString Then 'Found a group
            intIdx = intIdx + 1
        Else
            Exit Do
        End If
    Loop
    CountGroups = intIdx
End Function
'-----------------------------------------------------------
' SUB: GetGroup
'
' Returns the Groupname specified by Index
'-----------------------------------------------------------
'
Function GetGroup(ByVal strsection As String, ByVal Index As Integer)
    GetGroup = ReadIniFile(gstrSetupInfoFile, strsection, gsGROUP & CStr(Index))
End Function
'-----------------------------------------------------------
' SUB: CountIcons
'
' Determines how many icons must be installed by counting
' them in the setup information file (SETUP.LST)
'-----------------------------------------------------------
'
Function CountIcons(ByVal strsection As String) As Integer
    Dim intIdx As Integer
    Dim cIcons As Integer
    Dim sGroup As String
    Dim oCol As New Collection
    
    intIdx = 0
    cIcons = 0
    Do
        sGroup = ReadIniFile(gstrSetupInfoFile, strsection, gsGROUP & CStr(intIdx))
        If sGroup <> vbNullString Then 'Found a group
            oCol.Add sGroup
            intIdx = intIdx + 1
        Else
            Exit Do
        End If
    Loop
    Dim sGName As String
    Dim vGroup As Variant
    For Each vGroup In oCol
        intIdx = 1
        Do
            sGName = ReadIniFile(gstrSetupInfoFile, vGroup, gsICON & CStr(intIdx))
            If sGName <> vbNullString Then
                cIcons = cIcons + 1
                intIdx = intIdx + 1
            Else
                Exit Do
            End If
        Loop
    Next
    CountIcons = cIcons
    
End Function
'-----------------------------------------------------------
' SUB: CreateIcons
'
' Walks through the list of files in SETUP.LST and creates
' Icons in the Program Group for files needed it.
'-----------------------------------------------------------
'
Sub CreateIcons(ByVal strsection As String)
    Dim intIdx As Integer
    Dim sFile As FILEINFO
    Dim strProgramIconTitle As String
    Dim strProgramIconCmdLine As String
    Dim strProgramPath As String
    Dim strProgramArgs As String
    Dim intAnchor As Integer
    Dim intOffset As Integer
    Dim strGroup As String
    Dim sGroup As String
    Dim oCol As New Collection
    Const CompareBinary = 0
    '
    'For each file in the specified section, read info from the setup info file
    '
    intIdx = 0
    Do
        sGroup = ReadIniFile(gstrSetupInfoFile, strsection, gsGROUP & CStr(intIdx))
        If sGroup <> vbNullString Then 'Found a group
            oCol.Add sGroup
            intIdx = intIdx + 1
        Else
            Exit Do
        End If
    Loop
    Dim sGName As String
    Dim vGroup As Variant
    For Each vGroup In oCol
        intIdx = 0
        Do
            intIdx = intIdx + 1
            sGName = ReadIniFile(gstrSetupInfoFile, vGroup, gsICON & CStr(intIdx))
            If sGName <> vbNullString Then
                '
                ' Get the Icon's caption and command line
                '
                strProgramIconTitle = ReadIniFile(gstrSetupInfoFile, vGroup, gsTITLE & CStr(intIdx))
                strProgramIconCmdLine = ReadIniFile(gstrSetupInfoFile, vGroup, gsICON & CStr(intIdx))
                strGroup = vGroup
                '
                ' if the ProgramIcon is specified, then we create an icon,
                ' otherwise we don't.
                '
                If Trim(strUnQuoteString(strProgramIconTitle)) <> "" Then
                    '
                    ' If the command line is not specified in SETUP.LST and the icon
                    ' is, then use the files destination path as the command line.  In
                    ' this case there are no parameters.
                    '
                    If Trim(strUnQuoteString(strProgramIconCmdLine)) = "" Then
                        strProgramPath = sFile.strDestDir & gstrSEP_DIR & sFile.strDestName
                        strProgramArgs = ""
                    Else
                        '
                        ' Parse the command line, to determine what is the exe, etc. and what
                        ' are the parameters.  The first space that is not contained within
                        ' quotes, marks the end of the exe, etc..  Everything afterwards are
                        ' parameters/arguments for the exe.  NOTE: It is important that if
                        ' the exe is contained within quotes that the parameters not be
                        ' contained within the same quotes.  The arguments can themselves
                        ' each be inside quotes as long as they are not in the same quotes
                        ' with the exe.
                        '
                        intAnchor = 1
                        intOffset = intGetNextFldOffset(intAnchor, strProgramIconCmdLine, " ", CompareBinary)
                        If intOffset = 0 Then intOffset = Len(strProgramIconCmdLine) + 1
                        strProgramPath = Trim(Left(strProgramIconCmdLine, intOffset - 1))
                        '
                        ' Got the exe, now the parameters.
                        '
                        strProgramArgs = Trim(Mid(strProgramIconCmdLine, intOffset + 1))
                    End If
                    '
                    ' Expand all the Destination Directory macros that are embedded in the
                    ' Program Path and the Arguments'
                    '
                    strProgramPath = ResolveDestDir(strProgramPath)
                    strProgramArgs = ResolveDestDirs(strProgramArgs)
                    '
                    ' Finally, we have everything we need, create the icon.
                    '
                    CreateOSLink frmSetup1, strGroup, strProgramPath, strProgramArgs, strProgramIconTitle
                ElseIf Trim(strUnQuoteString(strProgramIconCmdLine)) <> "" Then
                    '
                    ' This file contained specified a command line in SETUP.LST but no icon.
                    ' This is an error.  Let the user know and skip this icon or abort.
        
                    '
                    If gfNoUserInput Or MsgWarning(ResolveResString(resICONMISSING, "|1", sFile.strDestName), vbYesNo Or vbExclamation, gstrSETMSG) = vbNo Then
                        ExitSetup frmSetup1, gintRET_FATAL
                    End If
                End If
            Else
                Exit Do
            End If
        Loop
    Next
End Sub



