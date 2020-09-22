Attribute VB_Name = "modSyntax"
'Option Explicit

Dim inputvar, SetI As Integer, SetX As Integer, Label, x, newvariable
Private VarNames As New Collection, VarData As New Collection
Private LabelNames As New Collection, LabelData As New Collection
Private InIf As Boolean, InElseIf As Boolean, InElse As Boolean, LoopEnd As Long
Private dScript As String, LoopState As String
Private iLocation As Long, SetI2 As Integer, Var, VarMin, VarMax, Step
Private ForLoops As New Collection, ForLoopDepth As Integer

Public Function CheckCode(Script As String) As Integer
On Error Resume Next
Dim i As Long, Temp As String
Dim TempArray As Variant

Dim FirstTime As Boolean

dScript = Script
'Primary Code Checking
For i = 1 To Len(dScript)
    iLocation = i
    If CheckIDE Then RemoveAll: Exit Function
    
    If i = 1 Then
        If Not FirstTime Then
            Firstime = True
        Else
            Exit Function
        End If
    End If
    
    If LCase(Mid(dScript, i, 5)) = "endif" Then
        i = i + 5
    ElseIf LCase(Mid(dScript, i, 4)) = "else" Then
        i = i + 4
    End If
    
    'output to user
    If LCase(Mid(dScript, i, 10)) = "screenout " Then
        i = i + 10
        
        CheckSemiColon i
        i = SetI
    
    'output all at once
    ElseIf LCase(Mid(dScript, i, 10)) = "screenput " Then
        i = i + 10

        CheckSemiColon i
        i = SetI
        
    'get input from user
    ElseIf LCase(Mid(dScript, i, 9)) = "screenin " Then
        i = i + 9
        
        If CheckSemiColon(i) = True Then
            inputvar = GetValue2(i, ";")
            If FindVar(inputvar) = False Then
                AddError ">Error at " & i & ": undefined identifier '" & inputvar & "'", i
            End If
        End If
        i = SetI

    'get input from user in password
    ElseIf LCase(Mid(dScript, i, 11)) = "screenpass " Then
        i = i + 11
        
        If CheckSemiColon(i) = True Then
            inputvar = GetValue2(i, ";")
            If FindVar(inputvar) = False Then
                AddError ">Error at " & i & ": undefined identifier '" & inputvar & "'", i
            End If
        End If
        i = SetI
        
    'title the application
    ElseIf LCase(Mid(dScript, i, 6)) = "title " Then
        i = i + 6
        
        CheckSemiColon i
        i = SetI
        
    'delete file
    ElseIf LCase(Mid(dScript, i, 7)) = "delete " Then
        i = i + 7

        CheckSemiColon i
        i = SetI
        
    'comment
    ElseIf Mid(dScript, i, 1) = "'" Then
        i = i + 1
        
        i = InStr(i, dScript, vbCrLf)
        If i = 0 Then GoTo FinishCode
        
    'create a message box
    ElseIf LCase(Mid(dScript, i, 4)) = "box " Then
        i = i + 4
    
        If CodeInStr(i, dScript, ",") = 0 Then
            AddError ">Error at " & i & ": expected ,", i
            i = i + 1
        Else
            i = CodeInStr(i, dScript, ",")
            If CodeInStr(i, dScript, ";") = 0 Then
                AddError ">Error at " & i & ": expected ;", i
                i = i + 1
            Else
                i = CodeInStr(i, dScript, ";")
            End If
        End If
        
    'pause for given time
    ElseIf LCase(Mid(dScript, i, 6)) = "pause " Then
        i = i + 6
        
        CheckSemiColon i
        i = SetI
    
    'Launch URL
    ElseIf LCase(Mid(dScript, i, 4)) = "web " Then
        i = i + 4
    
        CheckSemiColon i
        i = SetI
        
    'Open program
    ElseIf LCase(Mid(dScript, i, 5)) = "open " Then
        i = i + 5

        CheckSemiColon i
        i = SetI
        
    'Create new label
    ElseIf LCase(Mid(dScript, i, 6)) = "label " Then
        i = i + 6
        
        If CheckIdentifier2(i, ";") = True Then
            AddLabel GetValue2(i, ";"), i
        End If
        i = SetI
        
    'goto a label
    ElseIf LCase(Mid(dScript, i, 5)) = "goto " Then
        i = i + 5

        If CheckIdentifier2(i, ";") = True Then
            GetLabel GetValue2(i, ";")
        End If
        i = SetI
        
    'create a new variable
    ElseIf LCase(Mid(dScript, i, 7)) = "newvar " Then
        i = i + 7
        
        'this gets complicated
        TempArray = dSplit(GetValue2(i, ";"), ",") 'split it into an array
        SetI2 = SetI
        For z = LBound(TempArray) To UBound(TempArray) 'loop through array
            If z <> UBound(TempArray) Then  'if it's not the last array item
                If CheckIdentifier2(i, ",") = True Then 'if it's a valid variable
                    equals = InStr(1, TempArray(z), "=") 'look for equal sign
                    If equals = 0 Then 'nope, no equal sign
                        AddVariable Trim(TempArray(z)), "", False 'create variable
                        i = i + Len(TempArray(z)) + 2 'set length
                    Else 'yes, create variable and asign new value
                        AddVariable Trim(Mid(TempArray(z), 1, equals - 1)), Mid(TempArray(z), equals + 1), False
                        i = i + Len(TempArray(z)) + 2 'set length
                    End If
                End If
            Else 'if it's the last array item
                If CheckIdentifier2(i, ";") = True Then 'if it's a valid variable
                    equals = InStr(1, TempArray(z), "=") 'look for equal sign
                    If equals = 0 Then 'nope, no equal sign
                        AddVariable Trim(TempArray(z)), "", False 'create variable
                        i = i + Len(Trim(TempArray(z))) + 1 'set length
                    Else 'yes, create variable and asign new value
                        AddVariable Trim(Mid(TempArray(z), 1, equals - 1)), Mid(TempArray(z), equals + 1), False
                        i = i + Len(TempArray(z)) + 1 'set length
                    End If
                End If
            End If
        Next z
        i = SetI2
        
    'Loops
    ElseIf LCase(Mid(dScript, i, 9)) = "do until " Then
        i = i + 9
        
        LoopState = GetValue2(i, ";")   'loop expression
        eval LoopState
        i = SetI
        
        'check loop syntax
        FindLoopEnd i
        
    'loops
    ElseIf LCase(Mid(dScript, i, 9)) = "do while " Then
        i = i + 9
        
        LoopState = GetValue2(i, ";")   'loop expression
        eval LoopState
        i = SetI
        
        'check loop syntax
        FindLoopEnd i
        
    'more loop stuff
    ElseIf LCase(Mid(dScript, i, 4)) = "loop" Then
    
        On Error Resume Next
        If Asc(Mid(dScript, i + 4, 1)) > 64 And Asc(Mid(dScript, i + 4, 1)) < 122 Then GoTo SkipCurrent
        If Asc(Mid(dScript, i - 1, 1)) > 64 And Asc(Mid(dScript, i - 1, 1)) < 122 Then GoTo SkipCurrent
        FindLoopStart i - 1
        i = i + 4
        
    'FOR loops
    ElseIf LCase(Mid(dScript, i, 4)) = "for " Then
        i = i + 4
        
        Var = GetValue2(i, "=")
        i = SetI
        
        VarMin = GetValue(i, "to")
        i = SetI + 2
        
        If StringExist(i, dScript, "step") Then
            VarMax = GetValue(i, "step")
            i = SetI
            
            Step = GetValue(i, ";")
            i = SetI
        Else
            VarMax = GetValue(i, ";")
            
            i = SetI
            Step = 1
        End If
        
        SetVar Var, VarMin
        ForLoops.Add Var & ":" & VarMax & ":" & Step & ":" & i
        ForLoopDepth = ForLoopDepth + 1
        
    'FOR loops
    ElseIf LCase(Mid(dScript, i, 5)) = "next " Then
        i = i + 5
        
        If ForLoopDepth = 0 Then
            AddError ">Error at " & i - 5 & ": next without for", i - 5
        End If
        
        Var = GetValue2(i, ";")
        
        If Not CheckSemiColon(i, False) Then i = SetI: GoTo SkipCurrent
        
        If Var = GetCurrentLoopData(0) Then
            If FindVar(Var) = False Then
                AddError ">Error at " & i & ": undefined identifier '" & Var & "'", i
            Else
                ForLoops.Remove (ForLoopDepth)
                ForLoopDepth = ForLoopDepth - 1
            End If
        Else '
            AddError ">Error at " & i - 5 & ": invalid next reference (current=" & GetCurrentLoopData(0) & ")", i - 5
        End If
        
        i = SetI

    'if statments
    ElseIf LCase(Mid(dScript, i, 3)) = "if " Then
        i = i + 2
        
        If Mid(dScript, i - 5, 5) = "endif" Then GoTo SkipCurrent
        
        'Get expression
        ifstate = GetValue2(i, "then")
        eval ifstate
        i = SetI
        FindIfEnd i, False
    
    'expressions
    ElseIf LCase(Mid(dScript, i, 4)) = "set " Then
        i = i + 4

        TempString = GetValue2(i, ";")      'get the expression
        HandleExpression TempString         'handle it
        i = CodeInStr(i, dScript, ";")      'put our marker at the semicolon
        
    'play wav
    ElseIf LCase(Mid(dScript, i, 4)) = "wav " Then
        i = i + 4
        
        CheckSemiColon i
        i = SetI

    'these commands don't have arguments, so we don't check then
    'have to have them though, for exprssions to work.
    ElseIf LCase(Mid(dScript, i, 6)) = "clear;" Then
        i = i + 6
    ElseIf LCase(Mid(dScript, i, 5)) = "hide;" Then
        i = i + 5
    ElseIf LCase(Mid(dScript, i, 5)) = "show;" Then
        i = i + 5
    ElseIf LCase(Mid(dScript, i, 8)) = "open_cd;" Then
        i = i + 8
    ElseIf LCase(Mid(dScript, i, 9)) = "close_cd;" Then
        i = i + 9
    ElseIf LCase(Mid(dScript, i, 12)) = "disable_cad;" Then
        i = i + 12
    ElseIf LCase(Mid(dScript, i, 11)) = "enable_cad;" Then
        i = i + 11
    ElseIf LCase(Mid(dScript, i, 14)) = "show_controls;" Then
        i = i + 14
    ElseIf LCase(Mid(dScript, i, 14)) = "hide_controls;" Then
        i = i + 14
    ElseIf LCase(Mid(dScript, i, 9)) = "doevents;" Then
        i = i + 9
    ElseIf LCase(Mid(dScript, i, 7)) = "finish;" Then
        i = i + 7
    ElseIf LCase(Mid(dScript, i, 4)) = "end;" Then
        i = i + 4
    ElseIf LCase(Mid(dScript, i, 7)) = "screen;" Then
        i = i + 7
        
    Else 'if it's nothing else, we have to assume it's an expression
    
        'make sure it's not a space, return, wierd character, etc...
        If Asc(Mid(dScript, i, 1)) < 36 Then GoTo SkipCurrent
        If Asc(Mid(dScript, i, 1)) > 192 Then GoTo SkipCurrent
        
        TempString = GetValue2(i, ";")      'get the expression
        HandleExpression TempString         'handle it
        
        If CodeInStr(i, dScript, ";") = 0 Then RemoveAll: Exit Function
        i = CodeInStr(i, dScript, ";")      'put our marker at the semicolon
        
    End If
    
SkipCurrent:
Next i

FinishCode:
RemoveAll
End Function

Private Function CodeInStr(StartPos As Long, SourceText As String, ToFind As String) As Long
'This is just like InStr() only it skips over ""'s
Dim iPend As Long, i As Long
Dim sTemp As String

If StartPos > Len(SourceText) Then 'fatal error
    If Len(SourceText) > 10 Then
        AddError ">Fatal internal error! (" & Mid(SourceText, 1, 10) & "... , " & SourceText & ", " & ToFind & ")", StartPos
    ElseIf SourceText = dScript Then
        AddError ">Fatal internal error! (" & StartPos & ", dScript, " & ToFind & ")", StartPos
    Else
        AddError ">Fatal internal error! (" & StartPos & ", " & SourceText & ", " & ToFind & ")", StartPos
    End If
    CodeInStr = StartPos + 1
    Exit Function
End If

For i = StartPos To Len(SourceText)
    If CheckIDE Then RemoveAll: Exit Function
    
    If LCase(Mid(SourceText, i, Len(ToFind))) = LCase(ToFind) Then
        CodeInStr = i
        Exit Function
    End If
                          
    If Mid(SourceText, i, 1) = Chr(34) Then
        i = InStr(i + 1, SourceText, Chr(34))
        If i = 0 Then: Exit Function
    End If
Next i
End Function

Private Function HandleExpression(ByVal sExpression As String)
'This parses a string into a left value and right value, then
'asings the rightval to leftval
Dim LeftVal As String, RightVal As String
Dim sChar   As String, Operator As Integer
    
    If Right(sExpression, 2) = "++" Then 'incrament one
        LeftVal = Trim(Left(sExpression, Len(sExpression) - 2)) 'get the variable
        CheckIdentifier LeftVal 'make sure it's valid
        SetVar LeftVal, Val(GetVar(LeftVal)) + 1 'add one to vars current value
        Exit Function
    ElseIf Right(sExpression, 2) = "--" Then 'subtract one
        LeftVal = Trim(Left(sExpression, Len(sExpression) - 2)) 'get the variable
        CheckIdentifier LeftVal 'make sure it's valid
        SetVar LeftVal, Val(GetVar(LeftVal)) - 1 'subtract one from vars current value
        Exit Function
    End If

    'it's a normal expression
    Operator = InStr(1, sExpression, "=")   'find the equal sign
    If Operator = 0 Then    'this probably isn't an expression (maybe invalid command?)
        AddError ">Error at " & iLocation & ": syntax error: " & sExpression, iLocation
        Exit Function
    End If
    
    LeftVal = Trim(Left(sExpression, Operator - 1)) 'leftval is everything before equal sign
    
    CheckIdentifier LeftVal     'check it, make sure it's a valid var (no spaces, etc)
    
    RightVal = Solve(Trim(Mid(sExpression, Operator + 1))) 'rightval is everything after equal sign (solve it)
    SetVar LeftVal, RightVal    'assign the rightval to leftval
End Function

Private Function StringExist(StartPos As Long, SourceText As String, ToFind As String) As Boolean
    'This determines if a string exists in another string
If CodeInStr(StartPos, LCase(SourceText), LCase(ToFind)) = 0 Then
    StringExist = False
Else
    StringExist = True
End If
End Function

Private Function GetVar(TheVar As Variant) As Variant
Dim y As Integer
'Gets a variables value
If IsNumeric(TheVar) Then GetVar = TheVar: Exit Function
If FindVar(TheVar) = False Then
    Select Case LCase(TheVar)
        Case "dpp.ip", "dpp.host", "dpp.tick", "dpp.crlf", "true", "false", "dpp.systemfolder", "dpp.path"
            Exit Function
        Case Else
            AddError ">Error at " & iLocation & ": undefined identifier '" & TheVar & "'", iLocation
    End Select
End If
For y = 1 To VarNames.Count
    If VarNames(y) = TheVar Then
        GetVar = VarData(y)
        Exit Function
    End If
Next y
End Function

Private Function GetValue(Start As Long, Parameter As String, Optional InWhat As String) As Variant
On Error Resume Next
'this will search for the parameter, starting at the
'specified location, skipping things in quotes, then
'evalutate it.
If InWhat = "" Then InWhat = dScript
Dim FinalCode As String, WhereItIs As Long, i As Long
'Determine the location of the parameter
WhereItIs = CodeInStr(Start, InWhat, Parameter)
If WhereItIs = 0 Then
    'Not there? Flag an error.
    AddError ">Error at " & Start & ":  expected " & Parameter, Start
    SetI = Start
Else
    'Get the code between the starting location and the paramter
    FinalCode = Mid(InWhat, Start, WhereItIs - Start)
    'Evalutate it
    GetValue = Solve(FinalCode)
    'Set the location
    SetI = Start + Len(FinalCode) + Len(Parameter)
End If
End Function

Private Function GetValue2(Start As Long, Parameter As String) As Variant
On Error Resume Next
'This will search for the parameter, starting at the
'specified location, skipping things in quotes, but
'DOESN'T evaluate it.
Dim FinalCode As String, WhereItIs As Long
'Determine the location of the parameter
WhereItIs = CodeInStr(Start, dScript, Parameter)
If WhereItIs = 0 Then
    'Not there? Flag an error.
    AddError ">Error at " & Start & ": expected " & Parameter, Start
    SetI = Start
Else
    'Get the code between the starting location and the paramter
    FinalCode = Mid(dScript, Start, WhereItIs - Start)
    GetValue2 = Trim(FinalCode)
    'Set the location
    SetI = Start + Len(FinalCode) + Len(Parameter)
End If
End Function

Private Function GetValue3(Start As Long, Parameter As String, InWhat As String) As Variant
'This will search for the paramter, starting at the
'specified location, skipping things in quotes, but
'DOESN'T evaluate it.
On Error Resume Next
Dim FinalCode As Variant
Dim MakeInt As Integer
If InWhat = "" Then InWhat = dScript
If CodeInStr(Start, InWhat, Parameter) = 0 Then
    'Not there?  Flag an error.
    AddError ">Error at " & Start & ":  expected " & Parameter, Start
    SetX = Start
End If
For x = Start To Len(InWhat)
    If Mid(InWhat, x, Len(Parameter)) = Parameter Then
        FinalCode = Mid(InWhat, Start, x - Start)
        GetValue3 = Trim(FinalCode)
        SetX = Start + Len(FinalCode) + 1
        Exit Function
    ElseIf Mid(InWhat, x, 1) = Chr(34) Then
        Quote = InStr(x, InWhat, Chr(34))
        If Quote <> 0 Then
            x = Quote
        End If
    ElseIf Mid(InWhat, x, 1) = "(" Then
        Paren = InStr(x, InWhat, ")")
        If Paren <> 0 Then
            x = Paren
        End If
    End If
Next x
End Function

Private Function GetFunctions(sFunction As String) As Variant
'This sub searches for functions and solves them
On Error Resume Next
Dim TempFunction As String
Dim Arg1 As String, Arg2 As String, Arg3 As String
Dim Args As Variant, TempString As String
Dim x As Long, sChar As String, z As Long
For x = 1 To Len(sFunction)
    If CheckIDE Then RemoveAll: Exit Function
    sChar = Mid(sFunction, x, 1)
    'Lower Case
    If LCase(Mid(sFunction, x, 6)) = "lcase(" Then
        x = x + 6
        TempString = GetValue3(x, ")", sFunction)
        If TempString <> "" Then
            Args = dSplit(TempString, ",")
            x = SetX
            If UBound(Args) <> 1 Then
                AddError ">Error at " & iLocation & ": wrong number of arguments for lcase()", iLocation
            Else
                For z = LBound(Args) To UBound(Args)
                    TempString = Args(z)
                    Args(z) = Solve(TempString)
                Next z
                TempFunction = TempFunction & Chr(34) & LCase(Args(1)) & Chr(34)
            End If
        End If
    'Upper Case
    ElseIf LCase(Mid(sFunction, x, 6)) = "ucase(" Then
        x = x + 6
        TempString = GetValue3(x, ")", sFunction)
        If TempString <> "" Then
            Args = dSplit(TempString, ",")
            x = SetX
            If UBound(Args) <> 1 Then
                AddError ">Error at " & iLocation & ": wrong number of arguments for ucase()", iLocation
            Else
                For z = LBound(Args) To UBound(Args)
                    TempString = Args(z)
                    Args(z) = Solve(TempString)
                Next z
                TempFunction = TempFunction & Chr(34) & UCase(Args(1)) & Chr(34)
            End If
        End If
    'Length
    ElseIf LCase(Mid(sFunction, x, 4)) = "len(" Then
        x = x + 4
        TempString = GetValue3(x, ")", sFunction)
        If TempString <> "" Then
            Args = dSplit(TempString, ",")
            x = SetX
            If UBound(Args) <> 1 Then
                AddError ">Error at " & iLocation & ": wrong number of arguments for len()", iLocation
            Else
                For z = LBound(Args) To UBound(Args)
                    TempString = Args(z)
                    Args(z) = Solve(TempString)
                Next z
                TempFunction = TempFunction & Chr(34) & Len(Args(1)) & Chr(34)
            End If
        End If
    'Is a number
    ElseIf LCase(Mid(sFunction, x, 10)) = "isnumeric(" Then
        x = x + 10
        TempString = GetValue3(x, ")", sFunction)
        If TempString <> "" Then
            Args = dSplit(TempString, ",")
            x = SetX
            If UBound(Args) <> 1 Then
                AddError ">Error at " & iLocation & ": wrong number of arguments for isnumeric()", iLocation
            Else
                For z = LBound(Args) To UBound(Args)
                    TempString = Args(z)
                    Args(z) = Solve(TempString)
                Next z
                TempFunction = TempFunction & Chr(34) & IsNumeric(Args(1)) & Chr(34)
            End If
        End If
    'Is an operator
    ElseIf LCase(Mid(sFunction, x, 5)) = "isop(" Then
        x = x + 5
        TempString = GetValue3(x, ")", sFunction)
        If TempString <> "" Then
            Args = dSplit(TempString, ",")
            x = SetX
            If UBound(Args) <> 1 Then
                AddError ">Error at " & iLocation & ": wrong number of arguments for isop()", iLocation
            Else
                For z = LBound(Args) To UBound(Args)
                    TempString = Args(z)
                    Args(z) = Solve(TempString)
                Next z
                TempFunction = TempFunction & Chr(34) & IsOperator(Args(1)) & Chr(34)
            End If
        End If
    'make integer
    ElseIf LCase(Mid(sFunction, x, 4)) = "int(" Then
        x = x + 4
        TempString = GetValue3(x, ")", sFunction)
        If TempString <> "" Then
            Args = dSplit(TempString, ",")
            x = SetX
            If UBound(Args) <> 1 Then
                AddError ">Error at " & iLocation & ": wrong number of arguments for int()", iLocation
            Else
                For z = LBound(Args) To UBound(Args)
                    TempString = Args(z)
                    Args(z) = Solve(TempString)
                Next z
                TempFunction = TempFunction & Chr(34) & Int(Args(1)) & Chr(34)
            End If
        End If
    'FileExist
    ElseIf LCase(Mid(sFunction, x, 9)) = "filexist(" Then
        x = x + 9
        TempString = GetValue3(x, ")", sFunction)
        If TempString <> "" Then
            Args = dSplit(TempString, ",")
            x = SetX
            If UBound(Args) <> 1 Then
                AddError ">Error at " & iLocation & ": wrong number of arguments for filexist()", iLocation
            Else
                For z = LBound(Args) To UBound(Args)
                    TempString = Args(z)
                    Args(z) = Solve(TempString)
                Next z
                TempFunction = TempFunction & Chr(34) & FileExist(Args(1)) & Chr(34)
            End If
        End If
    'return text length
    ElseIf LCase(Mid(sFunction, x, 8)) = "conlen()" Then
        x = x + 8
        TempFunction = TempFunction & "10000"
        
    'Produce Random Number
    ElseIf LCase(Mid(sFunction, x, 7)) = "rndnum(" Then
        x = x + 7
        TempString = GetValue3(x, ")", sFunction)
        If TempString <> "" Then
            Args = dSplit(TempString, ",")
            x = SetX
            If UBound(Args) <> 1 Then
                AddError ">Error at " & iLocation & ": wrong number of arguments for rndnum()", iLocation
            Else
                For z = LBound(Args) To UBound(Args)
                    TempString = Args(z)
                    Args(z) = Solve(TempString)
                Next z
                Randomize
                TempFunction = TempFunction & Chr(34) & Int(Rnd * Int(Args(1))) & Chr(34)
            End If
        End If
    'InStr
    ElseIf LCase(Mid(sFunction, x, 6)) = "instr(" Then
        x = x + 6
        TempString = GetValue3(x, ")", sFunction)
        If TempString <> "" Then
            Args = dSplit(TempString, ",")
            x = SetX
            If UBound(Args) <> 3 Then
                AddError ">Error at " & iLocation & ": wrong number of arguments for instr()", iLocation
            Else
                For z = LBound(Args) To UBound(Args)
                    TempString = Args(z)
                    Args(z) = Solve(TempString)
                Next z
                TempFunction = TempFunction & Chr(34) & InStr(Args(1), Args(2), Args(3)) & Chr(34)
            End If
        End If
    'Right
    ElseIf LCase(Mid(sFunction, x, 6)) = "right(" Then
        x = x + 6
        TempString = GetValue3(x, ")", sFunction)
        If TempString <> "" Then
            Args = dSplit(TempString, ",")
            x = SetX
            If UBound(Args) <> 2 Then
                AddError ">Error at " & iLocation & ": wrong number of arguments for right()", iLocation
            Else
                For z = LBound(Args) To UBound(Args)
                    TempString = Args(z)
                    Args(z) = Solve(TempString)
                Next z
                TempFunction = TempFunction & Chr(34) & Right(Args(1), Args(2)) & Chr(34)
            End If
        End If
    'Left
    ElseIf LCase(Mid(sFunction, x, 5)) = "left(" Then
        x = x + 5
        TempString = GetValue3(x, ")", sFunction)
        If TempString <> "" Then
            Args = dSplit(TempString, ",")
            x = SetX
            If UBound(Args) <> 2 Then
                AddError ">Error at " & iLocation & ": wrong number of arguments for left()", iLocation
            Else
                For z = LBound(Args) To UBound(Args)
                    TempString = Args(z)
                    Args(z) = Solve(TempString)
                Next z
                TempFunction = TempFunction & Chr(34) & Left(Args(1), Args(2)) & Chr(34)
            End If
        End If
    'Mid
    ElseIf LCase(Mid(sFunction, x, 4)) = "mid(" Then
        x = x + 4
        TempString = GetValue3(x, ")", sFunction)
        If TempString <> "" Then
            Args = dSplit(TempString, ",")
            x = SetX
            If UBound(Args) <> 3 Then
                AddError ">Error at " & iLocation & ": wrong number of arguments for mid()", iLocation
            Else
                For z = LBound(Args) To UBound(Args)
                    TempString = Args(z)
                    Args(z) = Solve(TempString)
                Next z
                TempFunction = TempFunction & Chr(34) & Mid(Args(1), Args(2), Args(3)) & Chr(34)
            End If
        End If
    Else
        TempFunction = TempFunction & sChar
    End If
Next x
GetFunctions = TempFunction
End Function

Private Function Solve(sFunction As String) As Variant
'This is the base solver
sFunction = GetFunctions(sFunction) 'First clear all functions out
sFunction = Equation(sFunction)
Solve = SolveFunction(sFunction)    'Then solve
End Function

Private Function eval(ByVal sFunction As String) As Boolean
'This parses a string into a left value, operator, and right value
Dim LeftVal As String, RightVal As String, Operator
Dim sChar, OpFound As Boolean
OpFound = False

    For x = 1 To Len(sFunction)
        sChar = Mid(sFunction, x, 1)
        If sChar = ">" Or sChar = "<" Or sChar = "=" Then
            Operator = Operator & sChar
            OpFound = True
        Else
            If OpFound = True Then
                RightVal = RightVal & sChar
            Else
                LeftVal = LeftVal & sChar
            End If
        End If
    Next x
    
    LeftVal = Solve(LeftVal)
    RightVal = Solve(RightVal)
    
    eval = DoOperation2(LeftVal, Operator, RightVal)
End Function

Private Function Equation(ByVal sFunction As String) As Variant
'This sub basically looks for parentheses, and solves what's in them
Dim Paren1 As Integer, Paren2 As Integer, sChar As String
Do
    DoEvents
    For x = 1 To Len(sFunction)
        sChar = Mid(sFunction, x, 1)
        Select Case sChar
            Case Chr(34) 'Character 34 is the "
                x = InStr(x + 1, sFunction, Chr(34))
            Case "("
                Paren1 = x
            Case ")"
                Paren2 = x
                Exit For
        End Select
    Next x
    If Paren1 = 0 Then
        Exit Do
    Else
        sFunction = Mid(sFunction, 1, Paren1 - 1) & " " & Chr(34) & SolveFunction(Mid(sFunction, Paren1 + 1, Paren2 - (Paren1 + 1))) & Chr(34) & " " & Mid(sFunction, Paren2 + 1)
        Paren1 = 0
        Paren2 = 0
    End If
Loop
Equation = sFunction
End Function

Private Function SolveFunction(sFunction As String) As Variant
'This sub solves equations like 5 + (num1 * 3)
Dim Quote As Integer, sChar As String, variable As String
Dim Num2 As Variant, SolveOp As String, Num1
sFunction = Trim(sFunction)
For x = 1 To Len(sFunction)
    sChar = Mid(sFunction, x, 1)
    If sChar = Chr(34) Then                             'thats the "
        Quote = InStr(x + 1, sFunction, Chr(34))
        Num2 = Mid(sFunction, x + 1, Quote - (x + 1))
        x = Quote
        If SolveOp <> "" Then
            SolveFunction = DoOperation(SolveFunction, SolveOp, Num2)
            SolveOp = ""
        Else
            SolveFunction = Num2
        End If
    ElseIf IsOperator(sChar) = True Then
        If Num1 <> 0 Then
            Num2 = GetVar(Trim(Mid(sFunction, Num1, x - (Num1 + 1))))
            If SolveOp <> "" Then
                SolveFunction = DoOperation(SolveFunction, SolveOp, Num2)
                SolveOp = ""
            Else
                SolveFunction = Num2
            End If
            Num1 = 0
        End If
        SolveOp = sChar
    Else
        If Asc(sChar) <> 32 And Num1 = 0 Then Num1 = x
        If x >= Len(sFunction) Then
            Num2 = GetVar(Trim(Mid(sFunction, Num1, x)))
            If SolveOp <> "" Then
                SolveFunction = DoOperation(SolveFunction, SolveOp, Num2)
                SolveOp = ""
            Else
                SolveFunction = Num2
            End If
            Exit For
        End If
    End If
    GoTo NextX
NextX:
Next x
End Function

Private Function IsOperator(NumVal As Variant) As Boolean
'Determines if a character is an operator
Select Case NumVal
Case "+", "-", "*", "\", "/", "&", ">", "<", "="
    IsOperator = True
End Select
End Function

Private Function DoOperation(ByVal LeftVal As Variant, ByVal Operator As Variant, ByVal RightVal As Variant) As Variant
'Solves an equation
Select Case Operator
    Case "+"
        DoOperation = Val(LeftVal) + Val(RightVal)
    Case "-"
        DoOperation = Val(LeftVal) - Val(RightVal)
    Case "/"
        DoOperation = Val(LeftVal) / Val(RightVal)
    Case "\"
        DoOperation = Val(LeftVal) \ Val(RightVal)
    Case "^"
        DoOperation = Val(LeftVal) ^ Val(RightVal)
    Case "*"
        DoOperation = Val(LeftVal) * Val(RightVal)
    Case "&"
        DoOperation = LeftVal & RightVal
    Case Else
        AddError ">Error at " & iLocation & ": invalid operator, '" & Operator & "'", iLocation
End Select
End Function

Private Function DoOperation2(ByVal LeftVal As Variant, ByVal Operator As Variant, ByVal RightVal As Variant) As Boolean
'Determines if an expression is True or False
Select Case Operator
    Case ">"
        If IsNumeric(LeftVal) Or IsNumeric(RightVal) Then
            If Val(LeftVal) > Val(RightVal) Then DoOperation2 = True
        Else
            If LeftVal > RightVal Then DoOperation2 = True
        End If
    Case "<"
        If IsNumeric(LeftVal) Or IsNumeric(RightVal) Then
            If Val(LeftVal) < Val(RightVal) Then DoOperation2 = True
        Else
            If LeftVal < RightVal Then DoOperation2 = True
        End If
    Case "="
        If IsNumeric(LeftVal) Or IsNumeric(RightVal) Then
            If Val(LeftVal) = Val(RightVal) Then DoOperation2 = True
        Else
            If LeftVal = RightVal Then DoOperation2 = True
        End If
    Case "<>"
        If IsNumeric(LeftVal) Or IsNumeric(RightVal) Then
            If Val(LeftVal) <> Val(RightVal) Then DoOperation2 = True
        Else
            If LeftVal <> RightVal Then DoOperation2 = True
        End If
    Case ">="
        If IsNumeric(LeftVal) Or IsNumeric(RightVal) Then
            If Val(LeftVal) >= Val(RightVal) Then DoOperation2 = True
        Else
            If LeftVal >= RightVal Then DoOperation2 = True
        End If
    Case "<="
        If IsNumeric(LeftVal) Or IsNumeric(RightVal) Then
            If Val(LeftVal) <= Val(RightVal) Then DoOperation2 = True
        Else
            If LeftVal <= RightVal Then DoOperation2 = True
        End If
    Case Else
        AddError ">Error at " & iLocation & ": invalid operator, '" & Operator & "'", iLocation
End Select
End Function

Private Function FindVar(TheVar As Variant) As Boolean
On Error Resume Next
Dim z As Integer
For z = 1 To VarNames.Count
    If VarNames(z) = TheVar Then
        FindVar = True
        Exit Function
    End If
Next z
FindVar = False
End Function

Private Sub SetVar(TheVar As Variant, NewVal As Variant)
On Error Resume Next
'Sets the value of a variable
Dim z
'If a number...
If IsNumeric(TheVar) Then AddError ">Error at " & iLocation & ": cannot modify identifier '" & TheVar & "'", iLocation
'If variable doesn't exist...
If FindVar(TheVar) = False Then AddError ">Error at " & iLocation & ": undefined identifier '" & TheVar & "'", iLocation

If Mid(TheVar, 1, 4) = "dpp." Or TheVar = "True" Or TheVar = "False" Then
    AddError ">Error at " & iLocation & ": cannot modify identifier '" & TheVar & "' (identifier is constant)", iLocation
End If
For z = VarNames.Count To 1 Step -1
    If VarNames(z) = TheVar Then
        VarNames.Remove z
        VarData.Remove z
        VarNames.Add TheVar
        VarData.Add NewVal
        Exit Sub
    End If
Next z
End Sub

Private Sub AddVariable(VarName As Variant, VariableData As String, Optional PreSet As Boolean = True)
If IsNumeric(VarName) Then AddError "Error at " & iLocation & ": cannot modify identifier '" & VarName & "'.", iLocation
If FindVar(VarName) = True Then
    AddError ">Error at " & iLocation & ": cannot create identifier '" & VarName & "'", iLocation
Else
    VarNames.Add VarName
    If PreSet Then
        VarData.Add VariableData
    Else
        VarData.Add Solve(VariableData)
    End If
End If
End Sub

Private Function FindLabel(TheLabel As Variant) As Boolean
On Error Resume Next
Dim z
'Determines if a variable exists
For z = 1 To LabelNames.Count
    If LabelNames(z) = TheLabel Then
        FindLabel = True
        Exit Function
    End If
Next z
FindLabel = False
End Function

Private Function GetLabel(TheLabel As Variant) As Variant
On Error Resume Next
Dim z
'Gets a variables value
'MsgBox TheLabel & ": " & FindLabel(TheLabel)
If Not FindLabel(TheLabel) Then AddError ">Error at " & iLocation & ": undefined identifier '" & TheLabel & "'", iLocation
For z = 1 To LabelNames.Count
    If LabelNames(z) = TheLabel Then
        GetLabel = LabelData(z)
        Exit Function
    End If
Next z
End Function

Private Sub AddLabel(TheLabel As Variant, LabelPos As Variant)
If FindLabel(TheLabel) = True Then
    AddError ">Error at " & iLocation & ": cannot create identifier '" & TheLabel & "'", iLocation
Else
    LabelNames.Add TheLabel
    LabelData.Add LabelPos
End If
End Sub

Private Sub RemoveAll()
'On Error Resume Next
Dim z As Integer
For z = LabelNames.Count To 1 Step -1
    LabelNames.Remove z
    LabelData.Remove z
Next z
For z = VarNames.Count To 1 Step -1
    VarNames.Remove z
    VarData.Remove z
Next z
End Sub

Private Function CheckSemiColon(StartPos As Long, Optional eval As Boolean = True) As Boolean
Dim Text As String
Text = GetValue2(StartPos, ";")
If Text = "" Then Exit Function

If _
StringExist(1, Text, "screenout ") = True Or StringExist(1, Text, "screenput ") = True Or _
StringExist(1, Text, "screenin ") = True Or StringExist(1, Text, "screenpass ") = True Or _
StringExist(1, Text, "title ") = True Or StringExist(1, Text, "delete ") = True Or _
StringExist(1, Text, "box ") = True Or StringExist(1, Text, "pause ") = True Or _
StringExist(1, Text, "web ") = True Or StringExist(1, Text, "open ") = True Or _
StringExist(1, Text, "label ") = True Or StringExist(1, Text, "goto ") = True Or _
StringExist(1, Text, "newvar ") = True Or StringExist(1, Text, "do until ") = True Or _
StringExist(1, Text, "if ") = True Or StringExist(1, Text, "set ") = True Or _
StringExist(1, Text, "wav ") = True Or StringExist(1, Text, "clear;") = True Or _
StringExist(1, Text, "hide;") = True Or StringExist(1, Text, "show;") = True Or _
StringExist(1, Text, "open_cd;") = True Or StringExist(1, Text, "close_cd;") = True Or _
StringExist(1, Text, "enable_cad;") = True Or StringExist(1, Text, "disable_cad;") = True Or _
StringExist(1, Text, "show_controls;") = True Or StringExist(1, Text, "hide_controls;") = True Or _
StringExist(1, Text, "end;") = True Or StringExist(1, Text, "finsih;") = True Or _
StringExist(1, Text, "doevents;") = True Or StringExist(1, Text, "screen") = True Then
    AddError ">Error at " & StartPos & ": expected ;", StartPos
    SetI = StartPos + 1
    Exit Function
End If
If eval = False Then CheckSemiColon = True: Exit Function

Text = GetValue(StartPos, ";")
If Text = "" Then Exit Function
CheckSemiColon = True
End Function

Private Function CheckIdentifier2(StartPos As Long, Delimiter As String) As Boolean
Dim Text As String
Text = GetValue2(StartPos, Delimiter)
'MsgBox Text
If Text = "" Then Exit Function
If _
StringExist(1, Text, "screenout ") = True Or StringExist(1, Text, "screenput ") = True Or _
StringExist(1, Text, "screenin ") = True Or StringExist(1, Text, "screenpass ") = True Or _
StringExist(1, Text, "title ") = True Or StringExist(1, Text, "delete ") = True Or _
StringExist(1, Text, "box ") = True Or StringExist(1, Text, "pause ") = True Or _
StringExist(1, Text, "web ") = True Or StringExist(1, Text, "open ") = True Or _
StringExist(1, Text, "label ") = True Or StringExist(1, Text, "goto ") = True Or _
StringExist(1, Text, "newvar ") = True Or StringExist(1, Text, "do until ") = True Or _
StringExist(1, Text, "if ") = True Or StringExist(1, Text, "set ") = True Or _
StringExist(1, Text, "wav ") = True Or StringExist(1, Text, "clear;") = True Or _
StringExist(1, Text, "hide;") = True Or StringExist(1, Text, "show;") = True Or _
StringExist(1, Text, "open_cd;") = True Or StringExist(1, Text, "close_cd;") = True Or _
StringExist(1, Text, "enable_cad;") = True Or StringExist(1, Text, "disable_cad;") = True Or _
StringExist(1, Text, "show_controls;") = True Or StringExist(1, Text, "hide_controls;") = True Or _
StringExist(1, Text, "end;") = True Or StringExist(1, Text, "finsih;") = True Or _
StringExist(1, Text, "doevents;") = True Or StringExist(1, Text, "screen") = True Then

    AddError ">Error at " & StartPos & ": expected ;", StartPos
    SetI = StartPos + 1
    Exit Function
ElseIf StringExist(1, Text, ",") = True Or StringExist(1, Text, vbCrLf) = True Or _
       StringExist(1, Text, "(") Then
    AddError ">Error at " & StartPos & ": invalid identifier", StartPos
    SetI = StartPos + 1
    Exit Function
End If
CheckIdentifier2 = True
End Function

Public Function dSplit(Expression As String, Delimiter As String) As Variant
    ReDim SplitArray(1 To 1) As Variant
    Dim TempLetter As String
    Dim TempSplit As String
    Dim i As Integer
    Dim x As Integer
    Dim StartPos As Integer
    
    Expression = Expression & Delimiter
    For i = 1 To Len(Expression)
        If Mid(Expression, i, 1) = Chr(34) Then
            If InStr(i + 1, Expression, Chr(34)) = 0 Then
                AddError ">Error at " & iLocation & ": expected """, iLocation
                i = i + 1
            Else
                i = InStr(i + 1, Expression, Chr(34))
            End If
        End If
        TempLetter = Mid(Expression, i, Len(Delimiter))
        If TempLetter = Delimiter Then
            TempSplit = Mid(Expression, (StartPos + 1), (i - StartPos) - 1)
            If TempSplit <> "" Then
                x = x + 1
                ReDim Preserve SplitArray(1 To x) As Variant
                SplitArray(x) = TempSplit
            End If
            StartPos = i
        End If
    Next i
    dSplit = SplitArray
End Function

Private Function FindLoopEnd(StartPos As Long) As Long
Dim z As Long, LoopAt As Long, Quote As Long
For z = StartPos To Len(dScript)
    If Mid(dScript, z, 4) = "loop" Then
        If LoopAt = 0 Then
            FindLoopEnd = z + 4
            Exit Function
        Else
            LoopAt = LoopAt + 1
        End If
    ElseIf Mid(dScript, z, 8) = "do until" Then
        LoopAt = LoopAt - 1
    ElseIf Mid(dScript, z, 8) = "do while" Then
        LoopAt = LoopAt - 1
    ElseIf Mid(dScript, z, 1) = Chr(34) Then
        Quote = InStr(z + 1, dScript, Chr(34))
        If Quote <> 0 Then z = Quote
    End If
Next z
AddError ">Error at " & StartPos & ": expected 'loop'.", StartPos
End Function

Private Function FindLoopStart(StartPos As Long) As Long
Dim z As Long, LoopAt As Long, Quote As Boolean
LoopAt = 0
For z = StartPos To 1 Step -1
    If Mid(dScript, z, 8) = "do until" Then
        If Quote Then GoTo SkipCurrentZ
        If LoopAt = 0 Then
            FindLoopStart = z - 1
            Exit Function
        Else
            LoopAt = LoopAt - 1
        End If
    ElseIf Mid(dScript, z, 8) = "do while" Then
        If Quote Then GoTo SkipCurrentZ
        If LoopAt = 0 Then
            FindLoopStart = z - 1
            Exit Function
        Else
            LoopAt = LoopAt - 1
        End If
    ElseIf Mid(dScript, z, 4) = "loop" Then
        If Quote Then GoTo SkipCurrentZ
        LoopAt = LoopAt + 1
    ElseIf Mid(dScript, z, 1) = Chr(34) Then
        If Quote Then
            Quote = False
        Else
            Quote = True
        End If
    End If
SkipCurrentZ:
Next z
AddError ">Error at " & StartPos & ": loop without do or while", StartPos
End Function

Private Function FindIdentifier(StartPos As Long) As String
Dim i As Long
For i = StartPos To 1 Step -1
    Select Case Mid(dScript, i - 1, 1)
        Case " ", ";", vbTab
            If i > StartPos Then GoTo SkipCurrentI
            FindIdentifier = Trim(Mid(dScript, i, StartPos - (i - 1)))
            If FindIdentifier <> "" Then Exit Function
    End Select
    If Mid(dScript, i - 2, 2) = vbCrLf Then
        If i > StartPos Then GoTo SkipCurrentI
        FindIdentifier = Trim(Mid(dScript, i, StartPos - (i - 1)))
        If FindIdentifier <> "" Then Exit Function
        Exit Function
    End If
SkipCurrentI:
Next i
FindIdentifier = Trim(Mid(dScript, 1, Len(dScript) - StartPos))
End Function

Private Function FindIfEnd(StartPos As Long, UseElse As Boolean) As Long
Dim z As Long, IfAt As Long, Quote As Long
If UseElse = False Then GoTo EndIfs
For z = StartPos To Len(dScript) 'first look for an else for this if statement
    If LCase(Mid(dScript, z, 4)) = "else" Then
        If IfAt = 0 Then
            FindIfEnd = z + 3
            InIf = True
            Exit Function
        End If
    ElseIf LCase(Mid(dScript, z, 2)) = "if" Then
        If LCase(Mid(dScript, z - 3, 5)) <> "endif" Then 'make sure it's not in other statements
            If LCase(Mid(dScript, z - 4, 6)) <> "elseif" Then
                IfAt = IfAt + 1
            End If
        Else
            If IfAt = 0 Then
                FindIfEnd = z + 1
                InIf = True
                Exit Function
            Else
                IfAt = IfAt - 1
            End If
        End If
    ElseIf Mid(dScript, z, 1) = Chr(34) Then
        Quote = InStr(z + 1, dScript, Chr(34))
        If Quote <> 0 Then z = Quote
    End If
Next z
EndIfs:
IfAt = 0
For z = StartPos To Len(dScript) 'couldn't find an else statement in this if, goto endif
    If LCase(Mid(dScript, z, 5)) = "endif" Then
        If IfAt = 0 Then
            FindIfEnd = z + 6
            InIf = False
            Exit Function
        Else
            IfAt = IfAt - 1
        End If
    ElseIf LCase(Mid(dScript, z, 2)) = "if" Then
        If LCase(Mid(dScript, z - 3, 5)) <> "endif" Then 'make sure it's not in other statements
            If LCase(Mid(dScript, z - 4, 6)) <> "elseif" Then
                IfAt = IfAt + 1
            End If
        End If
    ElseIf Mid(dScript, z, 1) = Chr(34) Then
        Quote = InStr(z + 1, dScript, Chr(34))
        If Quote <> 0 Then z = Quote
    End If
Next z
AddError ">Error at " & StartPos & ": expected 'endif'", StartPos
End Function

Public Function CheckIDE() As Boolean
    DoEvents
    If StopDebug = True Then
        frmMain.ResumeIDE
        AddError ">Error, debug halted at " & iLocation & ".", iLocation
        CheckIDE = True
    End If
End Function

Private Function CountQuotes(Source As String) As Long
On Error Resume Next
Dim i As Integer
For i = 1 To Len(Source)
    If Mid(Source, i, 1) = Chr(34) Then
        CountQuotes = CountQuotes + 1
    End If
Next i
End Function

Private Function GetCurrentLoopData(Data As Integer) As Variant
Dim TempArray As Variant

TempArray = Split(ForLoops(ForLoopDepth), ":")
Select Case Data
    Case 0
        GetCurrentLoopData = TempArray(0)
    Case 1
        GetCurrentLoopData = TempArray(1)
    Case 2
        GetCurrentLoopData = TempArray(2)
    Case 3
        GetCurrentLoopData = TempArray(3)
End Select
End Function

Private Function CheckIdentifier(Identifier As String) As Boolean
Dim Text As String
Text = Trim(Identifier)
If Text = "" Then Exit Function
If _
StringExist(1, Text, "screenout ") = True Or StringExist(1, Text, "screenput ") = True Or _
StringExist(1, Text, "screenin ") = True Or StringExist(1, Text, "screenpass ") = True Or _
StringExist(1, Text, "title ") = True Or StringExist(1, Text, "delete ") = True Or _
StringExist(1, Text, "box ") = True Or StringExist(1, Text, "pause ") = True Or _
StringExist(1, Text, "web ") = True Or StringExist(1, Text, "open ") = True Or _
StringExist(1, Text, "label ") = True Or StringExist(1, Text, "goto ") = True Or _
StringExist(1, Text, "newvar ") = True Or StringExist(1, Text, "do until ") = True Or _
StringExist(1, Text, "if ") = True Or StringExist(1, Text, "set ") = True Or _
StringExist(1, Text, "wav ") = True Or StringExist(1, Text, "clear;") = True Or _
StringExist(1, Text, "hide;") = True Or StringExist(1, Text, "show;") = True Or _
StringExist(1, Text, "open_cd;") = True Or StringExist(1, Text, "close_cd;") = True Or _
StringExist(1, Text, "enable_cad;") = True Or StringExist(1, Text, "disable_cad;") = True Or _
StringExist(1, Text, "show_controls;") = True Or StringExist(1, Text, "hide_controls;") = True Or _
StringExist(1, Text, "end;") = True Or StringExist(1, Text, "finsih;") = True Or _
StringExist(1, Text, "doevents;") = True Or StringExist(1, Text, "screen;") = True Then

    AddError ">Error at " & iLocation & ": expression syntax error", iLocation
    Exit Function
ElseIf StringExist(1, Text, " ") = True Or StringExist(1, Text, vbCrLf) = True Or _
       StringExist(1, Text, "(") Then
    AddError ">Error at " & iLocation & ": expression syntax error", iLocation
    Exit Function
End If
CheckIdentifier = True
End Function

