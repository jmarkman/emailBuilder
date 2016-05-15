Private Sub btnCancel_Click()
    Unload subjectivitiesSelection
End Sub

Private Sub btnSubmit_Click()
    Dim msg As String
    Dim greeting As String
    Dim brokerName As String
    Dim StrSignature As String
    StrSignature = GetSignature(Environ("Userprofile") & "\AppData\Roaming\Microsoft\Signatures\CorRisk.htm")
    
    Dim myAttach(0 To 14) As Variant
    ' Declare the attachments that will correspond to the selections in the userform
    myAttach(0) = "[filepath]"
    myAttach(1) = "[filepath]"
    myAttach(2) = "[filepath]"
    myAttach(3) = "[filepath]"
    myAttach(4) = "[filepath]"
    myAttach(5) = "[filepath]"
    myAttach(6) = "[filepath]"
    myAttach(7) = "[filepath]"
    myAttach(8) = "[filepath]"
    myAttach(9) = "[filepath]"
    myAttach(10) = "[filepath]"
    myAttach(11) = "[filepath]"
    myAttach(12) = "[filepath]"
    myAttach(13) = "[filepath]"
    myAttach(14) = "[filepath]"
    
    ' Get currently open email                
    Dim followUp As Outlook.MailItem
    Dim myinspector As Outlook.Inspector
    Set myinspector = Application.ActiveInspector
    Set followUp = myinspector.CurrentItem.Forward
    
    ' Get original sender's email
    Set ObjSelectedItem = Outlook.ActiveExplorer.Selection.item(1)
    If TypeName(ObjSelectedItem) = "MailItem" Then
	If ObjSelectedItem.SenderEmailType = "EX" Then
	senderEmail = ObjSelectedItem.Sender.GetExchangeUser.PrimarySmtpAddress
	Else
	senderEmail = ObjSelectedItem.SenderEmailAddress
	End If
    Else
	MsgBox ("No items selected (OR) Selected item not a MailItem.")
    End If
    Set ObjSelectedItem = Nothing
    
    ' Determine greeting to use
    If time > #12:00:00 AM# And time <= #11:59:59 AM# Then
        greeting = "Good Morning"
    ElseIf time >= #12:00:00 PM# And time < #5:00:00 PM# Then
        greeting = "Good Afternoon"
    ElseIf time >= #5:00:00 PM# And time <= #11:59:59 PM# Then
        greeting = "Good Evening"
    End If
       
    With followUp
        Unload Me
        brokerName = InputBox("Please input the recipient's name")
    
        ' Compose message
        msg = "<HTML><BODY style = font-face:""Calibri(Body)"";font-size:""11pt"";color:""black"">" & greeting & " " & brokerName & ", " & "<br />" & "<br />"
           
        msg = msg & "This is a follow-up request for the following outstanding subjectivities: " & "<br />" & "<br />"
    
        For i = 0 To List1.ListCount - 1
            If List1.Selected(i) Then
                Counter = Counter + 1
                msg = msg & "<font style = 'background: yellow'>" & List1.List(i) & "<br />" & "</font>"
            Else: If Counter = 0 Then End
        End If
        Next
              
        ' Attach related form
        For i = 0 To List1.ListCount - 1
        If List1.Selected(i) = True Then
              .Attachments.Add myAttach(i)
        Else: If Counter = 0 Then End
        End If
        Next
                    
        msg = msg & "<br />" & "If these are not submitted within 3 days, a Notice of Cancellation will be sent. Please let us know if you have any questions or concerns." & "<br />" & "<br />" & "</BODY></HTML>"
                                           
        .Forward
        .To = senderEmail
        '.SentOnBehalfOfName = "[user]@corrisk.com"
        .HTMLBody = msg & vbNewLine + StrSignature + .HTMLBody
        Unload Me
        .Display
    End With
End Sub

Function GetSignature(fPath As String) As String
    Dim fso As Object
    Dim TSet As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set TSet = fso.GetFile(fPath).OpenAsTextStream(1, -2)
    GetSignature = TSet.readall
    TSet.Close
End Function


Private Sub Userform_Initialize()
    With List1
        .AddItem "5 year clean loss runs or warranty statement completed signed and dated in lieu of loss runs"
        .AddItem "A copy of the applicant’s resume"
        .AddItem "A copy of the expiring declaration page evidencing retroactive date"
        .AddItem "A copy of the standard contract including the scope of services"
        .AddItem "A fully completed, signed and dated AIG application, dated no more than 30 days prior to the effective date of coverage"
        .AddItem "A separate and fully completed, signed and dated claims supplemental application for each loss declared on the application and/ or loss runs detailing particulars of claim(s)"
        .AddItem "All subjectivities required prior to binding"
        .AddItem "All subjectivities required within 14 days of binding"
        .AddItem "All subjectivities required within 7 days of binding"
        .AddItem "Application submitted needs to be signed and dated no more than 30 days prior to the effective date of coverage"
        .AddItem "Clean loss history"
        .AddItem "Five Years of currently valued, dated within 30 days of the effective date, hard copy loss history provided by the insurance carrier on new business"
        .AddItem "Fully completed, signed and dated attached Warranty and Representation Letter"
        .AddItem "Subject to current & sound financials defined as an annual Profit & Loss Statement and Balance Sheet"
        .AddItem "The application submitted resigned and redated no more than 60 days prior to the effective date of coverage"
    End With
End Sub
