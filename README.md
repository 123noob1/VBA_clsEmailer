# VBA_clsEmailer
A simple manageable, maintainable, and updateable emailer class that can be used in any VBA project.

-- <b>To Use:</b>
- Dim the class clsEmailer either on class/form level or procedure level.
- Call the object's function <code>InitEmailContent</code> and set the parameters accordingly to generate the email.

- 'InitEmailContent(SubjectContent , BodyContent , [AttachFile], [SendTo], [SendCC], [SendBCC], [BodyFormat], [DisplayMail], [SendOnBehalf], [FlagImportance])'

- <code>SubjectContent</code> - [str] Email Subject line (required)
- <code>BodyContent</code> - [str] Email Body content (required)
- <code>AttachFile</code> - [str] Attach a file based on full file path (optional)
- <code>SendTo</code> - [str] Email To line (optional)
- <code>SendCC</code> - [str] Email CC line (optional)
- <code>SendBCC</code> - [str] Email BCC line (optional)
- <code>BodyFormat</code> - [int] Plan (0), HTML (1), Rich (2) (Optional)
- <code>DisplayMail</code> - [bool] True/False (optional)
- <code>SendOnBehalf</code> - [str] Send-on-Behalf (make sure you have this allowed) (optional)
- <code>FlagImportance</code> - [int] Low (0), Normal (1), High (2) (optional)

-- <b>Example:</b>
- Class/Form level

'''Option Explicit
Private clsEmail as New clasEmailer

Public Sub CreateEmail()
    Call clsEmail.InitEmailContent(SubjectContent:= "This is a Subject line", _
                                    BodyContent:= "This is a Body line", _
                                    SendTo:= "abc@gmail.com; user1@hotmail.com")
End Sub'''

- Within procedure/function

'''Public Sub CreateEmail()
Dim clsEmail As New clsEmailer

    Call clsEmail.InitEmailContent(SubjectContent:= "This is a Subject line", _
                                    BodyContent:= "This is a Body line", _
                                    SendTo:= "abc@gmail.com; user1@hotmail.com")
    
Set clsEmail = Nothing
End Sub'''