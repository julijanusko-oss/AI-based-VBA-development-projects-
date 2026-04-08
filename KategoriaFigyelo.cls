Public WithEvents FwdMail As Outlook.MailItem
Public EredetiLevel As Outlook.MailItem
Public KategoriaNev As String

Private Sub FwdMail_Send(Cancel As Boolean)
    On Error Resume Next
    If Not EredetiLevel Is Nothing And KategoriaNev <> "" Then
        EredetiLevel.Categories = KategoriaNev
        EredetiLevel.Save
    End If
    On Error GoTo 0
End Sub