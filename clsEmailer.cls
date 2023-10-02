VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsEmailer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' Class variables
Private mSendTo         As String
Private mSendCC         As String
Private mSendBCC        As String
Private mSendOnBehalf   As String
Private mSubjectContent As String
Private mBodyContent    As String
Private mAttachFile     As String
Private mBodyFormat     As BODYFORMATSTYLE
Private mFlagImportance As EMAILIMPORTANCE

Private mStopProcess    As Boolean
Private IsMultiFile     As Boolean
Private mDisplayMail    As Boolean

' This enum is for setting the body style of the email
Public Enum BODYFORMATSTYLE
    FORMAT_PLAIN = 0
    FORMAT_HTML = 1
    FORMAT_RICH = 2
End Enum

' This enum is for setting the email importance flag
Public Enum EMAILIMPORTANCE
    IMPORTANCE_HIGH = 2
    IMPORTANCE_LOW = 0
    IMPORTANCE_NORMAL = 1
End Enum

' ------------------------------------------------------
'| START THE LET AND GET PROPERTIES FOR CLASS VARIABLES |
' ------------------------------------------------------
Public Property Let SendTo(SendTo_ As String)
    mSendTo = SendTo_
End Property

Public Property Get SendTo() As String
    SendTo = mSendTo
End Property

Public Property Let SendCC(SendCC_ As String)
    mSendCC = SendCC_
End Property

Public Property Get SendCC() As String
    SendCC = mSendCC
End Property

Public Property Let SendBCC(SendBCC_ As String)
    mSendBCC = SendBCC_
End Property

Public Property Get SendBCC() As String
    SendBCC = mSendBCC
End Property

Public Property Let SendOnBehalf(SendOnBeHalf_ As String)
    mSendOnBehalf = SendOnBeHalf_
End Property

Public Property Get SendOnBehalf() As String
    SendOnBehalf = mSendOnBehalf
End Property

Public Property Let SubjectContent(SubjectContent_ As String)
    If Trim(SubjectContent_ & vbNullString) <> vbNullString Then
        mSubjectContent = SubjectContent_
    Else
        EndProcess "A subject title is required for an email to generate correctly", True
    End If
End Property

Public Property Get SubjectContent() As String
    SubjectContent = mSubjectContent
End Property

Public Property Let BodyContent(BodyContent_ As String)
    If Trim(BodyContent_ & vbNullString) <> vbNullString Then
        mBodyContent = BodyContent_
    Else
        EndProcess "A body text content is required for an email to generate correctly", True
    End If
End Property

Public Property Get BodyContent() As String
    BodyContent = mBodyContent
End Property

Public Property Let AttachFile(AttachFile_ As String)
    Dim fso As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Validate if the none empty value AttachFile_ is flagged True for IsMultiFile then assign the first file from the folder path
    ' to the class variable mAttachFile. Else we go one more step to validate that the value represents the correct file path or not.
    ' If not, we will end the process because the program is unable to locate the file or the path is incorrectly entered or only
    ' the folder path was entered.
    If Len(Trim(AttachFile_)) > 0 Then
        If IsMultiFile Then
            mAttachFile = AttachFile_
        Else
            If fso.FileExists(AttachFile_) Then
                mAttachFile = AttachFile_
            Else
                EndProcess "The provided file """ & AttachFile_ & """ is not a valid file or an incorrect path for attaching file(s)", True
            End If
        End If
    Else
        mAttachFile = Empty
    End If
    
    Set fso = Nothing
End Property

Public Property Get AttachFile() As String
    AttachFile = mAttachFile
End Property

Public Property Let BodyFormat(BodyFormat_ As BODYFORMATSTYLE)
    ' Validate to make sure that the values are within the assigned type only
    If BodyFormat_ >= 0 And BodyFormat_ <= 2 Then
        mBodyFormat = BodyFormat_
    Else
        EndProcess "The value provided for BODYFORMATSTYLE """ & BodyFormat_ & """ is not a valid value.", True
    End If
End Property

Public Property Get BodyFormat() As BODYFORMATSTYLE
    BodyFormat = mBodyFormat
End Property

Public Property Let FlagImportance(FlagImportance_ As EMAILIMPORTANCE)
    mFlagImportance = FlagImportance_
End Property

Public Property Get FlagImportance() As EMAILIMPORTANCE
    FlagImportance = mFlagImportance
End Property

Private Property Let StopProcess(StopProcess_ As Boolean)
' This should only be modified within the class itself only and not open to public
    mStopProcess = StopProcess_
End Property

Public Property Get StopProcess() As Boolean
    StopProcess = mStopProcess
End Property

Public Property Let DisplayMail(DisplayMail_ As Boolean)
    mDisplayMail = DisplayMail_
End Property

Public Property Get DisplayMail() As Boolean
    DisplayMail = mDisplayMail
End Property
' ----------------------------------------------------
'| END THE LET AND GET PROPERTIES FOR CLASS VARIABLES |
' ----------------------------------------------------

' This is the initial part of initialize the variables needed for the email creation and validate for correctness.
' If any validation failed (currently only on attachment), it will stop the entire process.
Public Sub InitEmailContent(SubjectContent As String, BodyContent As String, _
                        Optional AttachFile As String = "", _
                        Optional SendTo As String = "", Optional SendCC As String = "", Optional SendBCC As String = "", _
                        Optional BodyFormat As BODYFORMATSTYLE = FORMAT_PLAIN, Optional DisplayMail As Boolean = True, _
                        Optional SendOnBehalf As String = "", Optional FlagImportance As EMAILIMPORTANCE = IMPORTANCE_NORMAL)
    Dim fso         As Object
    Dim stFilePath  As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Verify is there is a wild card file types and names inside the AttachFile String then set up the parent folder path for use later in
    ' attaching multiple files
    If InStr(AttachFile, "*.*") > 0 Then stFilePath$ = fso.GetParentFolderName(AttachFile) & "\"
    
    ' Set the default check for multi file usage each time the procedure is called
    IsMultiFile = False

    With Me
        ' Assign values to class variables
        .SubjectContent = SubjectContent
        .BodyContent = BodyContent
        .SendTo = SendTo
        .SendCC = SendCC
        .SendBCC = SendBCC
        .SendOnBehalf = SendOnBehalf
        .BodyFormat = BodyFormat
        .FlagImportance = FlagImportance
        .DisplayMail = DisplayMail
        
        If stFilePath <> "" Then
            IsMultiFile = True
            .AttachFile = Dir(stFilePath & "*.*")
        Else
            .AttachFile = AttachFile
        End If

    End With
    
    If Not StopProcess Then
        Call CreateEmail(stFilePath)
    End If
    
    Set fso = Nothing
End Sub

' This is where the emailing begins
Private Sub CreateEmail(FilePath As String)
    Dim oOutlook    As Object
    Dim oMailItem   As Object
    
    ' Check that mStopProcess is not set to True then we can start the emailing process
    If Not mStopProcess Then
        Set oOutlook = CreateObject("Outlook.Application")
        Set oMailItem = oOutlook.CreateItem(0)
        
        With oMailItem
            ' Start composing the email
            If SendOnBehalf <> "" Then
                .SentOnBehalfOfName = SendOnBehalf
            End If
            
            .To = SendTo
            .CC = SendCC
            .bCC = SendBCC
            .Subject = SubjectContent
            .Importance = FlagImportance
            
            ' Check to see what type of body format of the email will be then set the body content text
            Select Case BodyFormat
                Case 0  ' Default/plain text
                    .BodyFormat = 1 ' olFormatPlain
                    .body = BodyContent
                Case 1  ' HTML style
                    .BodyFormat = 2 ' olFormatHTML
                    .HTMLBody = BodyContent
                Case 2  ' Rich Text style
                    .BodyFormat = 3 ' olFormatRichText
                    .RTFBody = BodyContent
            End Select
            
            ' Attach single or multiple files depending on the validation above
            If Len(Me.AttachFile) > 0 Then
                If FilePath <> "" Then
                    Do While Len(Me.AttachFile) > 0
                        .Attachments.Add (FilePath & Me.AttachFile)
                        Me.AttachFile = Dir
                    Loop
                Else
                    .Attachments.Add Me.AttachFile
                End If
            End If
            
            ' Save the email so it doesn't lose or cause issues
            .Save
            
            If DisplayMail Then
                .Display
            ElseIf Not DisplayMail Then
                ' Currently .Send does not work and prompts error code 287. Use the SendKeys function instead with ("%S") with
                ' some wait time before sending to be in sync with Outlook;However, this requires user not to do anything else.
                .Send
'                .Display
'                Pause 2
'                SendKeys ("%S")
'                Pause 1
            End If
        End With
    End If
    
    mStopProcess = False
    
    Set oMailItem = Nothing
    Set oOutlook = Nothing
End Sub

' This procedure will be called only if there is an error during any of the email validations. When called,
' it will end the current process and user would need to re-run the procedure before proceding
Private Sub EndProcess(REASON As String, StopProcess_ As Boolean)
Dim sMessageBuilder As String

    sMessageBuilder = "Process terminated due to the following reason:" & vbNewLine & vbNewLine
    sMessageBuilder = sMessageBuilder & REASON & vbNewLine & vbNewLine
    sMessageBuilder = sMessageBuilder & "Please make sure to verify and correct above reason before trying again."
    
    MsgBox sMessageBuilder, vbExclamation + vbOKOnly, "PROCESS TERMINATED"
    StopProcess = StopProcess_
End Sub