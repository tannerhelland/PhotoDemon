Attribute VB_Name = "Evaluator"
'***************************************************************************
'Expression Evaluator (infix notation)
'Copyright 2018 by Olaf Schmidt, with additional improvements by Jason Peter Brown
'Created: 25/June/18
'Last updated: 25/June/18
'Last update: initial integration into PD; see https://github.com/tannerhelland/PhotoDemon/issues/263 for details
'
'In June 2018, Jason Peter Brown (https://www.github.com/jpbro) suggested adding arbitrary expression
' evaluation to PhotoDemon's various edit boxes.  He followed this up with a great deal of research and
' ultimately the full-blown submission of a working evaluator, based on work originally shared by
' Olaf Schmidt.  Thank you to these two individuals for enabling arbitrary expression evaluation in
' PD's input forms!
'
'Olaf's original evaluator design can be found here (link good as of June '18):
' http://www.vbforums.com/showthread.php?860225-simple-math-string-parser&p=5271805&viewfull=1#post5271805
'
'While Jason's modified version can be found here (link good as of June '18):
' https://github.com/tannerhelland/PhotoDemon/issues/263
'
'Jason's changes include:
' - Added Evaluate wrapper function to pre-evaluation processing to passed expressions
' - Simplified by removing "Function" support
' - Simplified by remove logical/boolean operator support
' - Raise errors for expressions that can't be evaluated.
' - Added "CanEvaluate" method that returns true if passed expression can be evaluated
'
'The code have been further modified for integration into PD, but any additional changes by the
 'PhotoDemon authors can be considered "public domain".
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Can a given expression be evaluated?  Note that this doesn't necessarily indicate that an expression
' can be evaluated *correctly*; rather, it just means the expression doesn't crash the parser.
' (Also note that the passed expression string *may* be modified; it must be passed ByVal.)
'
'RETURNS: TRUE if the expression looks evaluate-able; FALSE otherwise.
Public Function CanEvaluate(ByVal srcExpression As String) As Boolean
   
    On Error Resume Next
    Err.Clear
    
    Dim l_Eval As Variant
    l_Eval = Evaluator.Evaluate(srcExpression)
    CanEvaluate = (Err.Number = 0)
    
    On Error GoTo 0
   
End Function

'Attempt to evaluate any arbitrary string expression as numeric input (e.g. 1 + 2 - 3).
' Standard operators, parentheses, and order of operations should all be handled correctly.
'
'RETURNS: numeric result of evaluation, if one exists
Public Function Evaluate(ByVal srcExpression As String) As Variant
    
    'Coerce arbitrary decimal separators into the standard, locale-invariant "."
    ' (For details on why we do this, see PD's primary implementation in TextSupport.CDblCustom().)
    If (InStr(1, srcExpression, ",", vbBinaryCompare) <> 0) Then srcExpression = Replace$(srcExpression, ",", ".", , , vbBinaryCompare)
    If (InStr(1, srcExpression, ChrW$(&H66B&), vbBinaryCompare) <> 0) Then srcExpression = Replace$(srcExpression, ChrW$(&H66B&), ".", , , vbBinaryCompare)
    
    'See if the passed expression is a plain number; if it is, return it immediately
    If TextSupport.IsNumberLocaleUnaware(srcExpression) Then
       Evaluate = CDblCustom(srcExpression)
    
    'The passed string is not a plain number; attempt to evaluate it as an expression
    Else
      
        'Start by removing spaces and problematic operators (e.g. "--")
        If InStr(1, srcExpression, " ", vbBinaryCompare) Then srcExpression = Replace$(srcExpression, " ", vbNullString, , , vbBinaryCompare)
        If InStr(1, srcExpression, "--", vbBinaryCompare) Then srcExpression = Replace$(srcExpression, "--", "+", , , vbBinaryCompare)
        
        'Find any "(" with preceding numeric characters, and manually insert a * operator
        Dim p As Long
        Do
           
            p = InStr(p + 1, srcExpression, "(", vbBinaryCompare)
              
            If (p > 1) Then
                Select Case Mid$(srcExpression, p - 1, 1)
                    Case "0" To "9", "."
                        srcExpression = Left$(srcExpression, p - 1) & "*" & Mid$(srcExpression, p)
                        p = p + 1
                 End Select
            End If
            
            Loop While p > 0
         
        'Preprocessing is complete.  Pass the finished string to the actual evaluator.
        Evaluate = Eval(srcExpression)
      
    End If
   
End Function

Private Function Eval(ByVal srcExpression As String) As Variant
    
    'Preprocess any parentheses in the expression
    Do While HandleParentheses(srcExpression): Loop
    
    Dim l As String, r As String
    
    'Check for standard operators.  Note that we silently convert some VB6 patterns
    ' (e.g. "\" for "integer division") to standard usage patterns.
    If Spl(srcExpression, "Or", l, r) Then:  Eval = Eval(l) Or Eval(r): Exit Function
    If Spl(srcExpression, "And", l, r) Then: Eval = Eval(l) And Eval(r): Exit Function
    If Spl(srcExpression, ">=", l, r) Then:  Eval = Eval(l) >= Eval(r): Exit Function
    If Spl(srcExpression, "<=", l, r) Then:  Eval = Eval(l) <= Eval(r): Exit Function
    If Spl(srcExpression, "=", l, r) Then:   Eval = Eval(l) = Eval(r): Exit Function
    If Spl(srcExpression, ">", l, r) Then:   Eval = Eval(l) > Eval(r): Exit Function
    If Spl(srcExpression, "<", l, r) Then:   Eval = Eval(l) < Eval(r): Exit Function
    If Spl(srcExpression, "Like", l, r) Then Eval = Eval(l) Like Eval(r): Exit Function
    If Spl(srcExpression, "&", l, r) Then:   Eval = Eval(l) & Eval(r): Exit Function
    If Spl(srcExpression, "+", l, r) Then:   Eval = Eval(l) + Eval(r): Exit Function
    If Spl(srcExpression, "-", l, r) Then:   Eval = Eval(l) - Eval(r): Exit Function
    If Spl(srcExpression, "Mod", l, r) Then: Eval = Eval(l) Mod Eval(r): Exit Function
    If Spl(srcExpression, "\", l, r) Then:   Eval = Eval(l) \ Eval(r): Exit Function
    If Spl(srcExpression, "*", l, r) Then:   Eval = Eval(l) * Eval(r): Exit Function
    If Spl(srcExpression, "/", l, r) Then:   Eval = Eval(l) / Eval(r): Exit Function
    If Spl(srcExpression, "^", l, r) Then:   Eval = Eval(l) ^ Eval(r): Exit Function
    If Trim$(srcExpression) >= "A" Then:     Eval = EvalFnc(Trim$(srcExpression)): Exit Function
    If InStr(srcExpression, "'") Then Eval = Replace(Trim$(srcExpression), "'", ""): Exit Function
    
    'Return numeric characters as-is, with an additional check for duplicate negatives
    If (LenB(srcExpression) <> 0) Then Eval = Val(Replace(srcExpression, "--", vbNullString))
    
    'If we don't want to support characters or functions, check for invalid alpha characters here
    'If (Trim$(srcExpression) >= "A") Then: Err.Raise 5, , "Invalid expression.": Exit Function
    
    If (LenB(srcExpression) <> 0) Then

        'Test for some unevaluatable conditions
        Select Case Left$(srcExpression, 1)

          Case "-", "+"

             'Check for non-numeric character after + or -
             Select Case Mid$(srcExpression, 2, 1)
                Case "0" To "9", ".", "(", ")"
                Case Else
                   Err.Raise 5, , "Invalid expression: " & srcExpression
             End Select

          Case "0" To "9", ".", "(", ")"
             'Numeric and "." characters are OK

          Case Else

             'Unacceptable starting character (non-numeric, "+", "-", or ".")
             Err.Raise 5, , "Invalid expression: " & srcExpression

        End Select

        Eval = Val(srcExpression)

    End If
   
End Function

Private Function HandleParentheses(ByRef srcExpression As String) As Boolean
   
    Dim p As Long
    
    'Check for the presence of parentheses, and while we're at it, attempt to compensate for
    ' mismatched pairs (e.g. "(" without ")")
    p = InStr(1, srcExpression, "(", vbBinaryCompare)
    If (p < 1) Then
    
        'No opening parenthesis; check for a closing parenthesis
        If InStr(srcExpression, ")") Then
            
            'A closing parenthesis exists.  Silently insert an opening one at the start of the expression.
            srcExpression = "(" & srcExpression
            p = 1
            
        'No parenthesis
        Else
            Exit Function
        End If
        
    End If
    
    'At least one opening parenthesis exists; check for a matching closing parenthesis
    If InStrRev(srcExpression, ")", p) Then
        
        'The user is likely still typing out an expression (e.g. "(1+2").  Rather than fail,
        ' silently insert a closing parenthesis at the end of the expression.
        srcExpression = srcExpression & ")"
        
    End If
    
    'If we're still here, the function contains one or more sets of parentheses.  Find the innermost set of
    ' parentheses and evaluate it.
    Dim i As Long, c As Long, parenthesesUnbalanced As Boolean
    
    Do
        
        c = 0
        parenthesesUnbalanced = False
        
        For i = p To Len(srcExpression)
            If (Mid$(srcExpression, i, 1) = "(") Then c = c + 1
            If (Mid$(srcExpression, i, 1) = ")") Then c = c - 1
            If (c = 0) Then Exit For
        Next i

        'If parentheses are unbalanced, try to automatically insert an opening or closing parentheses,
        ' as necessary, then search again for the innermost pair.
        If (c <> 0) Then
            parenthesesUnbalanced = True
            If (c < 0) Then srcExpression = "(" & srcExpression Else srcExpression = srcExpression & ")"
        End If
    
    Loop While parenthesesUnbalanced
    
    'Replace the outermost parentheses pair with the numeric result of its interior contents.
    srcExpression = Left$(srcExpression, p - 1) & Trim$(Str$(Eval(Mid$(srcExpression, p + 1, i - p - 1)))) & Mid$(srcExpression, i + 1)
    
End Function

'Attempt to split a source expression into three parts: left and right operands, separated by a known operator.
Private Function Spl(ByRef srcExpression As String, ByRef Operator As String, ByRef l As String, ByRef r As String) As Boolean

    'Look for the requested operator in the source expression
    Dim p As Long, l_ReEval As Boolean
    p = -1

Rematch:
    p = InStrRev(srcExpression, Operator, p, vbTextCompare)
    If p Then Spl = True Else Exit Function
    
    If ((p < InStrRev(srcExpression, "'", -1, vbBinaryCompare)) And InStr(1, "*-", Operator, vbBinaryCompare)) Then p = InStrRev(srcExpression, "'", p, vbBinaryCompare) - 1
    
    'Separate out the left and right operands
    l = Trim$(Left$(srcExpression, p - 1))
    r = Mid$(srcExpression, p + Len(Operator))
    
    Do
        
        Select Case Right$(l, 1)
        
            Case "", "+", "A" To "z"
                Spl = False
                Exit Do
                
            Case "*", "/", "\"
                If (Operator = "-") Then
                    p = p - 1
                    l_ReEval = True
                    Spl = False
                    GoTo Rematch
                Else
                    Spl = False
                    'Expression contains an operator after another operator,
                    ' and the operator is not a "-" which would indicate a negative number
                    Err.Raise 5, , "Invalid expression: " & srcExpression
                    Exit Do
                End If
                 
            Case "-"
                l = Trim$(Left$(l, Len(l) - 1))
                r = "-" & r
                
            Case Else
                Exit Do
            
        End Select
        
    Loop
    
End Function

Private Function EvalFnc(ByRef srcExpr As String) As Variant
    Select Case LCase$(Left$(srcExpr, 3))
        Case "abs": EvalFnc = Abs(Val(Mid$(srcExpr, 4)))
        Case "sin": EvalFnc = Sin(Val(Mid$(srcExpr, 4)))
        Case "cos": EvalFnc = Cos(Val(Mid$(srcExpr, 4)))
        Case "atn": EvalFnc = Atn(Val(Mid$(srcExpr, 4)))
        Case "log": EvalFnc = Log(Val(Mid$(srcExpr, 4)))
        Case "exp": EvalFnc = Exp(Val(Mid$(srcExpr, 4)))
        Case Else: Err.Raise 5, , "Invalid expression: " & srcExpr
    End Select
End Function
