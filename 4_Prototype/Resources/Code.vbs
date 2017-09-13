Option Compare Database
'set a public variable for the userlevel as Userlevel
Public UserLevel As String
'declare a variable for the timer as StartTimer
Dim StartTimer As Integer

Private Sub cmdSubmit_Click()
'check if the user name is empty or not
If IsNull(Me.txtUser) Or Me.txtUser = "" Then
   'if it is show message
   MsgBox "User Name is Empty", vbInformation, "Login"
   'if it is empty the UserName text box and focus on it
   Me.txtUser.SetFocus

Exit Sub

End If
'check if the user name and password is matching
If (IsNull(DLookup("UserName", "tblUserDetails", "UserName = '" & Me.txtUser.Value & "'"))) Or _
   (IsNull(DLookup("Password", "tblUserDetails", "Password = '" & Me.txtPassword.Value & "'"))) Then
   'if it is not matching then show message
   MsgBox "Invalid User Name or Password", vbCritical, "Alert"
   'empty User name Text box
   Me.txtUser = ""
   'empty Password Text Box
   Me.txtPassword = ""
   'focus on the User Name Text Box
   Me.txtUser.SetFocus
   'disable password text box
   Me.txtPassword.Enabled = False
   'disable command button Submit
   Me.cmdSubmit.Enabled = False
   
Else
'make the UserLevel variable equal to lookup the Userlevel from the User table
UserLevel = DLookup("UserLevel", "tblUserDetails", "UserName = '" & Me.txtUser.Value & "'")
            'if the user is a manager
            If UserLevel = "Manager" Then
                ' call on the sub ClosenOpen
                Call ClosenOpen
                'Open form
                DoCmd.OpenForm "frmManagerAuthority"
                   
            Else
                'Call on the sub ClosenOpen
                Call ClosenOpen
                'Open form
                DoCmd.OpenForm "frmMainMenu", acNormal, "", "", acFormReadOnly
                
            End If
            
End If

End Sub
Sub ClosenOpen()
'logstatus is a gloabal variable set to pass on the userlevel of the user
logstatus = Me.UserLevel
'close form
DoCmd.Close acForm, "frmLogin", acSaveNo

End Sub

Private Sub Form_Load()
'set StartTimer variable to 30 seconds
StartTimer = 30
'set timerinterval equal to 1000 milliseconds
TimerInterval = 1000

End Sub

Private Sub Form_Timer()
'StartTimer is not equal to 30 second, so in this process it will count down. StartTimer = 30 - 1
StartTimer = StartTimer - 1
'if the StartTimer is equal to 0 Seconds
If StartTimer = 0 Then
 'then timerinterval is equal to 0 seconds
 TimerInterval = 0
 'close form
 DoCmd.Close acForm, "frmLogin", acSaveNo
 'open form
 DoCmd.OpenForm "frmWelcomeScreen", acNormal, "", "", acFormReadOnly, acWindowNormal
 
 End If

End Sub

Private Sub txtPassword_Change()
'if the Password Text box is empty
If txtPassword.Text = "" Then
   'then disable the command button Submit
   cmdSubmit.Enabled = False

Else
   'if the Password Text Box is not empty then enable the command button Submit
   cmdSubmit.Enabled = True
   
End If

End Sub


Private Sub txtUser_Change()
'if the UserName text box is not empty then enable the Password Text box
txtPassword.Enabled = True
'if the User Name Text box is empty
If txtUser.Text = "" Then
   'then disable the Password text box
   txtPassword.Enabled = False

Else
   'if the User Name Text box is not empty then enable the Password Text box
   txtPassword.Enabled = True

End If

End Sub
