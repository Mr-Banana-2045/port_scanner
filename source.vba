Dim ip As String
Dim objshell, bool
    ip = TextBox1.Text
    
    Set objshell = CreateObject("Wscript.Shell")
    bool = objshell.Run("ping -n 1 -w 1000 " & ip, 0, True)
    
    If bool = 0 Then
        Label1.Caption = True & ": " & ip
    Else
        Label1.Caption = False
    End If
