Option Explicit
'Petroleum Experts Ltd - Open Server VBA Example

' These lines declare global variables
Dim Server As Object
Dim Connected As Integer
Dim lErr As Long
Dim Command As String
Dim AppName As String
Dim OSString As String
Dim typeFluide As Long

Sub Macro()

Connect

'Write Macro Here

 
 
 

Disconnect
    
End Sub

Sub Connect() 'This utility creates the OpenServer object which allows comunication between Excel and IPM tools
    
    If Connected = 0 Then
        Set Server = CreateObject("PX32.OpenServer.1")
        Connected = 1
    End If

End Sub
Sub Disconnect()
    
    If Connected = 1 Then
       Set Server = Nothing
       Connected = 0
    End If

End Sub

' This utility function extracts the application name from the tag string
Function GetAppName(Strval As String) As String
   Dim Pos
   Pos = InStr(Strval, ".")
   If Pos < 2 Then
        MsgBox "Badly formed tag string"
        End
   End If
   GetAppName = Left(Strval, Pos - 1)
  
End Function
' Perform a command, then check for errors
Sub DoCmd(Cmd As String)
    Dim lErr As Long
    lErr = Server.DoCommand(Cmd)
    If lErr > 0 Then
        MsgBox Server.GetErrorDescription(lErr)
        Set Server = Nothing
        End
    End If
End Sub
'Set a value, then check for errors
Sub DoSet(Sv As String, Val)
    Dim lErr As Long
    lErr = Server.SetValue(Sv, Val)
    AppName = GetAppName(Sv)
    lErr = Server.GetLastError(AppName)
    If lErr > 0 Then
        MsgBox Server.GetErrorDescription(lErr)
        Set Server = Nothing
        End
    End If
End Sub
' Get a value, then check for errors
Function DoGet(Gv As String) As String
    Dim lErr As Long
    DoGet = Server.GetValue(Gv)
    AppName = GetAppName(Gv)
    lErr = Server.GetLastError(AppName)
    If lErr > 0 Then
        MsgBox Server.GetLastErrorMessage(AppName)
        Set Server = Nothing
        End
    End If
End Function
' Perform a command, then wait for the command to exit
' Then check for errors
Sub DoSlowCmd(Cmd As String)
    Dim starttime As Single
    Dim endtime As Single
    Dim CurrentTime As Single
    Dim lErr As Long
    Dim bLoop As Boolean
    Dim step As Single
        
    step = 0.001
    AppName = GetAppName(Cmd)
    lErr = Server.DoCommandAsync(Cmd)
    If lErr > 0 Then
        MsgBox Server.GetErrorDescription(lErr)
        Disconnect
        End
    End If
    While Server.IsBusy(AppName) > 0
        If step < 2 Then
            step = step * 2
        End If
        starttime = Timer
        endtime = starttime + step
        Do
            CurrentTime = Timer
            'DoEvents
            bLoop = True
            Rem Check first for the case where we have gone over midnight
            Rem and the number of seconds will go back to zero
            If CurrentTime < starttime Then
                bLoop = False
            Rem Now check for the 2 second pause finishing
            ElseIf CurrentTime > endtime Then
                bLoop = False
            End If
        Loop While bLoop
    Wend
    AppName = GetAppName(Cmd)
    lErr = Server.GetLastError(AppName)
    If lErr > 0 Then
        MsgBox Server.GetErrorDescription(lErr)
        Disconnect
        End
    End If
End Sub


' Perform a function in GAP, then retrieve return value
' Finally, check for errors
Function DoGAPFunc(Gv As String) As String
    DoSlowCmd Gv
    DoGAPFunc = DoGet("GAP.LASTCMDRET")
    lErr = Server.GetLastError("GAP")
    If lErr > 0 Then
        MsgBox Server.GetErrorDescription(lErr)
        End
    End If
End Function
