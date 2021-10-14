Module Module1

    Sub Main()

        For index = 1 To 10

        Next

        Const WAITTICK As Long = 100

        message = ""
        isComplete = False

        Dim oSession = GetObject("SAPGUI").GetScriptingEngine.Children(0).Children(0)
        On Error GoTo mEnd

        Dim index
        Dim isEnd
        index = WAITTICK
        isEnd = False
        Do
            index = index - 1
            If Not oSession.findbyid("wnd[1]/tbar[0]/btn[8]", False) Is Nothing Then
                isEnd = True
            End If
        Loop Until index < 0 Or isEnd = True

        If index < 0 Or isEnd = False Then
            GoTo mEnd
        Else
            oSession.findById("wnd[1]/tbar[0]/btn[8]").press
        End If

        oSession.findById("wnd[0]/usr/ctxtLIST_BRE").text = ""
        oSession.findById("wnd[0]/usr/txtMAX_SEL").text = ""
        oSession.findById("wnd[0]/usr/txtMAX_SEL").setFocus
        oSession.findById("wnd[0]/usr/txtMAX_SEL").caretPosition = 11
        oSession.findById("wnd[0]/tbar[1]/btn[8]").press

        'Wait for load
        index = WAITTICK
        isEnd = False
        Do
            index = index - 1
            If Not oSession.findbyid("wnd[0]/usr/cntlGRID1/shellcont/shell", False) Is Nothing Then
                isEnd = True
            End If
        Loop Until index < 0 Or isEnd = True

        If index < 0 Or isEnd = False Then
            GoTo mEnd
        End If

        oSession.findById("wnd[0]").sendVKey(43)

        index = WAITTICK
        isEnd = False
        Do
            index = index - 1
            If Not oSession.findbyid("wnd[1]/usr/radRB_OTHERS", False) Is Nothing Then
                isEnd = True
            End If
        Loop Until index < 0 Or isEnd = True

        If index < 0 Or isEnd = False Then
            GoTo mEnd
        Else
            oSession.findById("wnd[1]/usr/radRB_OTHERS").setFocus
            oSession.findById("wnd[1]/usr/radRB_OTHERS").select
            oSession.findById("wnd[1]/usr/cmbG_LISTBOX").setFocus
            oSession.findById("wnd[1]/usr/cmbG_LISTBOX").key = "04"
            'oSession.findById("wnd[1]/tbar[0]/btn[0]").press
        End If

        Sleep(300)
        isComplete = True

mEnd:
        If Not (isComplete) Then
            message = message & Err.Description
        End If

mEndSpecial:
        On Error GoTo 0

        'close SUP
        oSession = Nothing

    End Sub

End Module
