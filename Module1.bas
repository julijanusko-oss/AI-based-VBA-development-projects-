Global LevelNyomkoveto As KategoriaFigyelo

Sub StartProcess()
    EmailPanel.Show
End Sub

Function GetClipboardText() As String
    On Error Resume Next
    With CreateObject("htmlfile")
        GetClipboardText = .ParentWindow.ClipboardData.GetData("text")
    End With
    On Error GoTo 0
End Function

Sub ProcessEmail(Eset As String, Optional KollegaNev As String = "")
    Dim olExplorer As Object
    Dim outMail As Object
    Dim ClipboardText As String
    Dim ToAddr As String, CcAddr As String, Greeting As String
    Dim KollegaEmail As String, Megszolitas As String
    Dim Kategoria As String
    
    ClipboardText = GetClipboardText()

    Select Case KollegaNev
        Case "Pityi Palkó": KollegaEmail = "palko@ceg.hu": Megszolitas = "Szia Palkó"
        Case "Sanyi": KollegaEmail = "sanyi@ceg.hu": Megszolitas = "Szia Sanyi"
        Case "Kovács Béla": KollegaEmail = "bela@ceg.hu": Megszolitas = "Szia Béla"
        Case "Nagy Anna": KollegaEmail = "anna@ceg.hu": Megszolitas = "Szia Anna"
        Case "Ludas Matyi": KollegaEmail = "matyi@ceg.hu": Megszolitas = "Szia Matyi"
    End Select

    Select Case Eset
        Case "MOBO"
            ToAddr = "mobo@ceg.hu": Greeting = "Sziasztok MOBO!": Kategoria = "MOBO Kategória"
        Case "MO"
            ToAddr = "mo@ceg.hu": Greeting = "Szia MO!": Kategoria = "MO Kategória"
        Case "TEL"
            ToAddr = "tel@ceg.hu": Greeting = "Szia TEL!": Kategoria = "TEL Kategória"
        Case "TS"
            ToAddr = KollegaEmail: Greeting = "Szia " & Megszolitas & "!": Kategoria = "TS Kategória"
        Case "MO_KOLLEGA_CC"
            ToAddr = "mo@ceg.hu; " & KollegaEmail: CcAddr = "fix_szemely@ceg.hu"
            Greeting = "Sziasztok, szia " & Megszolitas & "!": Kategoria = "MO_CC Kategória"
    End Select

    Set olExplorer = Application.ActiveExplorer
    
    On Error Resume Next
    olExplorer.CommandBars.ExecuteMso ("Forward")
    DoEvents
    Set outMail = olExplorer.ActiveInlineResponse
    On Error GoTo 0
    
    If outMail Is Nothing Then
        MsgBox "Nem sikerült a beágyazott továbbítás. Ellenőrizd, be van-e kapcsolva az Olvasóablak!", vbCritical
        Exit Sub
    End If

    Set LevelNyomkoveto = New KategoriaFigyelo
    Set LevelNyomkoveto.FwdMail = outMail
    Set LevelNyomkoveto.EredetiLevel = olExplorer.Selection.Item(1)
    LevelNyomkoveto.KategoriaNev = Kategoria

    With outMail
        .To = ToAddr
        If CcAddr <> "" Then .CC = CcAddr
        If Trim(ClipboardText) <> "" Then .Subject = ClipboardText
        .HTMLBody = "<p style='font-family:Calibri;font-size:11pt;'>" & Greeting & "<br><br>" & _
                    "az alábbi levelet továbbítom, jó munkát!<br>" & _
                    "Üdv,</p>" & .HTMLBody
    End With
End Sub