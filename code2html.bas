Attribute VB_Name = "modCode2Html"
Option Explicit
' ===================================================================================
' ===================================================================================
' CODE2HTML - by M Ferris of Intact Interactive Software. Copyright M Ferris 2001
' ===================================================================================
' ===================================================================================
' Okay so here's some code to convert vb sourcecode from raw text into HTML -
' The finished result has color coding similar to that of the VB IDE and
' preserves all of the tab positioning.
'
' It won't manage to translate HTML tags inside strings - what you will get is
' probably just an empty string, but - hey - I don't put a whole lot of HTML
' tags in my VB code so it works okay for me
'
' Also I have optimized the tagging a bit so that it doesn't put a font tag for
' each reserved word - instead it tries wherever possible to wrap the group of
' words up in one tag - i.e. <font color="#000088">Private Function</font>
' instead of <font color="#000088">Private</font> <font color="#000088">Private</font>
' So the resulting HTML is a bit smaller...
'
' If you like this code don't forget to come visit us on the web at
' www.intactinteractive.com
' Oh yeah - and don't forget to vote for me if you got this from a free source code
' site.
'
' Any other comments or feedback would also be appreciated - except for abuse of
' course!
' ===================================================================================
' ===================================================================================
    
Const BLU = "<FONT COLOR = ""#000088"">"
Const GRN = "<FONT COLOR = ""#008800"">"
Const CLOSEFONT = "</FONT>"
Const OPENSTR = "<!-- CODE2HTML by M Ferris - WWW.INTACTINTERACTIVE.COM --><BR><FONT COLOR=""#000000""><TT><PRE>"
Const CLOSESTR = "</FONT></PRE></TT>"
    
Dim InResWord As Boolean
Dim InComment As Boolean

Function CheckReserved(s As String) As String
    Dim RW As Variant
    Dim i As Integer
    Dim tmp As String
    Dim lnOrig As String
    Dim ln As Integer
    
    RW = Array("Const", "Else", "ElseIf", "If", "Alias", "And", "As", "Base", "Binary", "Boolean", "Byte", "ByVal", "Call", "Case", "CBool", _
               "CByte", "CCur", "CDate", "CDbl", "CDec", "CInt", "CLng", "Close", "Compare", "Const", "CSng", "CStr", "Currency", "CVar", "CVErr", _
               "Decimal", "Declare", "DefBool", "DefByte", "DefCur", "DefDate", "DefDbl", "DefDec", "DefInt", "DefLng", "DefObj", "DefSng", "DefStr", _
               "DefVar", "Dim", "Do", "Double", "Each", "Else", "ElseIf", "End", "Enum", "Eqv", "Erase", "Error", "Exit", "Explicit", "False", "For", _
               "Function", "Get", "Global", "GoSub", "GoTo", "If", "Imp", "In", "Input", "Input", "Integer", "Is", "LBound", "Let", "Lib", "Like", "Line", _
               "Lock", "Long", "Loop", "LSet", "Name", "New", "Next", "Not", "Object", "Open", "Option", "On", "Or", "Output", "Preserve", "Print", "Private", _
               "Property", "Public", "Put", "Random", "Read", "ReDim", "Resume", "Return", "RSet", "Seek", "Select", "Set", "Single", "Spc", "Static", "String", _
               "Stop", "Sub", "Tab", "Then", "True", "UBound", "Variant", "While", "Wend", "With")
    
    tmp = UCase(Trim(s))
    ln = Len(tmp)
    lnOrig = ln
    If ln > 2 Then
        If Mid(tmp, Len(tmp) - 1, 2) = "()" Then
            ln = Len(tmp) - 2
            tmp = Mid(tmp, 1, ln)
        End If
        If Mid(tmp, Len(tmp), 1) = ")" Then
            ln = Len(tmp) - 1
            tmp = Mid(tmp, 1, ln)
        End If
        If Mid(tmp, Len(tmp), 1) = ":" Then
            ln = Len(tmp) - 1
            tmp = Mid(tmp, 1, ln)
        End If
    End If
    
    For i = 0 To UBound(RW)
        If UCase(RW(i)) = tmp Then
            If InResWord Then
                If ln <> lnOrig Then
                    CheckReserved = Mid(s, 1, ln) & CLOSEFONT & Mid(s, ln + 1)
                    InResWord = False
                Else
                    CheckReserved = s
                End If
            Else
                If ln <> Len(tmp) Then
                    CheckReserved = Mid(s, 1, ln) & CLOSEFONT & Mid(s, ln + 1)
                    InResWord = False
                Else
                    CheckReserved = BLU & s
                    InResWord = True
                End If
            End If
            Exit Function
        End If
    Next i
    If InResWord Then
        CheckReserved = CLOSEFONT & s
    Else
        CheckReserved = s
    End If
    InResWord = False
End Function


Function GetLines(s As String) As String()
    Dim Lines() As String
    Dim i As Integer
    Dim EOS As Boolean
    Dim numLines As Integer
    Dim ln As Integer
    Dim tmp As String
    
    tmp = s
    EOS = False
    i = 1
    Do While Not EOS
        ln = InStr(tmp, vbCrLf)
        If ln = 0 Or ln > Len(tmp) Then EOS = True: Exit Do
        ReDim Preserve Lines(numLines)
        numLines = numLines + 1
        Lines(UBound(Lines)) = Mid(tmp, 1, IIf(ln = 1, 1, ln - 1))
        tmp = Mid(tmp, ln + 2)
    Loop
    ReDim Preserve Lines(numLines)
    Lines(UBound(Lines)) = tmp
    GetLines = Lines
End Function

Function GetWords(s As String) As String()
    Dim Words() As String
    Dim i As Integer
    Dim EOS As Boolean
    Dim numWords As Integer
    Dim ln As Integer
    Dim tmp As String
    Dim t As String
    
    tmp = s
    EOS = False
    i = 1
    Do While Not EOS
        ln = InStr(tmp, " ")
        If ln = 0 Or ln > Len(tmp) Then EOS = True: Exit Do
        ReDim Preserve Words(numWords)
        numWords = numWords + 1
        t = Mid(tmp, 1, ln)
        If Mid(t, 1, 1) = vbTab Then t = "&#09;" & Trim(t) & " "
        Words(UBound(Words)) = t
        tmp = Mid(tmp, ln + 1)
    Loop
    ReDim Preserve Words(numWords)
    Words(UBound(Words)) = tmp
    GetWords = Words
End Function

Function ProcessBlock(s As String) As String
    'process a raw block of code in the form of text and turn it into
    'html with color coding similar to that of the VB IDE ...
    Dim tmp As String
    Dim i As Integer
    Dim n As Integer
    Dim m As Integer
    Dim lns() As String
    Dim wrds() As String
    Dim t As String
    
    Screen.MousePointer = 11
    tmp = OPENSTR & vbCrLf
    InComment = False
    InResWord = False
    lns = GetLines(s)
    For i = 0 To UBound(lns)
        wrds = GetWords(lns(i))
        
        If InComment Then
            If Mid(Trim(wrds(0)), 1, 1) <> "'" Then
                InComment = False
                tmp = tmp & CLOSEFONT
            End If
        End If
        If Mid(Trim(wrds(0)), 1, 1) = "'" Then
            If InResWord Then
                InResWord = False
                tmp = tmp & CLOSEFONT
            End If
            If Not InComment Then
                tmp = tmp & GRN & lns(i)
            Else
                tmp = tmp & lns(i)
            End If
            InComment = True
        Else
            For n = 0 To UBound(wrds)
                If Mid(Trim(wrds(n)), 1, 1) = "'" Then
                     InComment = True
                     If InResWord Then
                        InResWord = False
                        tmp = tmp & CLOSEFONT
                     End If
                     tmp = tmp & GRN
                     For m = n To UBound(wrds)
                         tmp = tmp & wrds(m)
                     Next m
                     Exit For
                End If
                t = CheckReserved(wrds(n))
                tmp = tmp & t
            Next n
        End If
        Erase wrds
        tmp = tmp & vbCrLf
    Next i
    Erase lns
    If InComment Or InResWord Then
        tmp = tmp & CLOSEFONT
    End If
    tmp = tmp & CLOSESTR & vbCrLf
    ProcessBlock = tmp
    Screen.MousePointer = 0
End Function

' ===================================================================================
' CODE2HTML - by M Ferris of Intact Interactive Software. Copyright M Ferris 2001
' ===================================================================================

