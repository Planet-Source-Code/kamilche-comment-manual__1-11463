VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPrintFormatted"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'This class handles printing of formatted text.

Option Explicit
Private mPaperWidth As Double
Private mPaperHeight As Double
Private mLeftMargin As Double
Private mRightMargin As Double
Private mTopMargin As Double
Private mBottomMargin As Double
Private mBoxed As Boolean
Private mNewPage As Boolean
Private PrintObject As Object
Private Const mLineAdvance As Double = 1 / 12
Private mStyles() As typeFont

Friend Sub SetMargins(LeftMargin As Double, RightMargin As Double, _
  TopMargin As Double, BottomMargin As Double)
    'Set the margins - example, 1 inch all the way around.
    mLeftMargin = LeftMargin
    mRightMargin = RightMargin
    mTopMargin = TopMargin
    mBottomMargin = BottomMargin
End Sub

Friend Sub SetPaperSize(Width As Double, Height As Double)
    'Set physical paper size - example: 8.5x11
    mPaperWidth = Width
    mPaperHeight = Height
End Sub

Friend Sub SetDestination(Where As Object)
    'Set the print destination to either the printer or a picturebox.
    'This allows for easier 'print previewing.'
    Set PrintObject = Where
End Sub

Friend Property Let Boxed(ByVal vData As Boolean)
    'Draw a box around the page?
    mBoxed = vData
End Property

Friend Sub PrintPreview()
    'Print the first file's comment manual to the picture window.
    PrintFile Files(1)
End Sub

Friend Sub StartPrinting()
    'Set up the printer for printing.
    PrintObject.ScaleMode = vbInches
    If TypeOf PrintObject Is PictureBox Then
        PrintObject.Cls
    Else
        PrintObject.Print "";
    End If
    PrintObject.CurrentX = mLeftMargin
    PrintObject.CurrentY = mTopMargin
    mNewPage = True
    mBoxed = True
End Sub

Friend Sub EndPrinting()
    'End the current document.
    If TypeOf PrintObject Is PictureBox Then
        'skip it
    Else
        PrintObject.EndDoc
    End If
End Sub

Friend Sub PrintFile(ByVal FileName As String)
'Prints a complete file.
    Dim FileNo As Integer, LastHeader As String, TheHeader As String
    Dim s As String, s1 As String, s2() As String, c As Long
    FileNo = FreeFile
    Open FileName For Input As #FileNo
    StartPrinting
    PrintText "Methods in " & LastPart(FileName), stTitle, Center, True
    PrintText " " & vbCrLf & " ", stPlain, Left, True
    Do While Not EOF(FileNo)
        Line Input #FileNo, s: s = Trim(s)
        If IsHeading(s) Then
            'Print the method name
            c = InStr(1, s, "(", vbTextCompare)
            If c > 0 Then
                s1 = Mid$(s, 1, c - 1)
                s2 = Split(s1, " ")
                If InStr(1, s, "Public", vbTextCompare) = 1 Or InStr(1, s, "Friend", vbTextCompare) = 1 Then
                    'print it in bold
                    TheHeader = s2(UBound(s2, 1))
                    If TheHeader = LastHeader Then
                        'skip it
                    Else
                        PrintText TheHeader, stHeader, Left, True
                        LastHeader = TheHeader
                    End If
                Else
                    'print it in italics
                    PrintText s2(UBound(s2, 1)), stHeaderPrivate, Left, True
                End If
            End If
            'Print the syntax
            Do While InStr(1, s, " _", vbTextCompare) > 0
                PrintText s, stPlain, Left, True
                Line Input #FileNo, s: s = Trim(s)
            Loop
            PrintText s, stPlain, Left, True
            Line Input #FileNo, s: s = Trim(s)
            Do While InStr(1, s, "'", vbTextCompare) = 1
                PrintText s, stItalic, Left, True
                Line Input #FileNo, s: s = Trim(s)
            Loop
            PrintText " ", stPlain, Left, True
        End If
    Loop
    Close #FileNo
    EndPrinting
End Sub

Friend Sub PrintText(ByVal s As String, Style As enumStyles, Justify As enumJustification, LineFeedAfter As Boolean)
    'Print text with the specified font and justification.
    
    Dim TextWidth As Double, TextHeight As Double
    Dim StartPt As Double, TextRect As RECT
    
    'Set the new PrintObject settings
    With PrintObject.Font
        .Name = mStyles(Style).Name
        .Bold = mStyles(Style).Bold
        .Italic = mStyles(Style).Italic
        .Size = mStyles(Style).Size
        .Underline = mStyles(Style).Underline
    End With
    TextWidth = PrintObject.TextWidth(s)
    TextHeight = PrintObject.TextHeight("W")
    
    'Check to see if we should start a new page.
    If (PrintObject.CurrentY + TextHeight) > (mPaperHeight - mBottomMargin) Then
        If TypeOf PrintObject Is PictureBox Then
            'stop printing - we've exceeded the size.
            Exit Sub
        Else
            PrintObject.NewPage
        End If
        mNewPage = True
    End If
    
    'Do 'first time printing on page' stuff including boxes.
    If mNewPage = True Then
        mNewPage = False
        If TypeOf PrintObject Is PictureBox Then
            PrintObject.Cls
        End If
        If mBoxed = True Then
            PrintObject.Line (0.5, 0.5)-(mPaperWidth - 0.5, mPaperHeight - 0.5), , B
        End If
        PrintObject.CurrentX = mLeftMargin
        PrintObject.CurrentY = mTopMargin
    End If
    
    'Calculate x position
    If Justify = Left Then
        'calculate left
        PrintObject.CurrentX = mLeftMargin
        'print at left margin
        PrintIt s, LineFeedAfter
    ElseIf Justify = Center Then
        'calculate center
        PrintObject.CurrentX = ((mPaperWidth - mLeftMargin - mRightMargin) / 2) - (PrintObject.TextWidth(s) / 2) + mLeftMargin
        'print at center
        PrintIt s, LineFeedAfter
    ElseIf Justify = Right Then
        'calulate right
        PrintObject.CurrentX = (mPaperWidth - mRightMargin) - PrintObject.TextWidth(s)
        'print at right
        PrintIt s, LineFeedAfter
    End If
End Sub

Friend Sub PrintLine()
    'Prints a horizontal line going the width of the page.
    PrintObject.Line (0.5, PrintObject.CurrentY)-(mPaperWidth - 0.5, PrintObject.CurrentY), , B
    PrintObject.CurrentY = PrintObject.CurrentY + mLineAdvance
End Sub

Private Sub PrintIt(ByVal s As String, ByVal LineFeedAfter As Boolean)
    'Print the text to the PrintObject, optionally with a carriage return after it.
    Dim r As RECT, RetVal As Long
    'Set up the initial size of the rectangle.
    If LineFeedAfter = True Then
        'PrintObject.Print s
        r.Left = PrintObject.ScaleX(PrintObject.CurrentX, vbInches, vbPixels)
        r.Right = r.Left + (PrintObject.ScaleX(mPaperWidth, vbInches, vbPixels) - PrintObject.ScaleX(mLeftMargin, vbInches, vbPixels) - PrintObject.ScaleX(mRightMargin, vbInches, vbPixels))
        r.Top = PrintObject.ScaleX(PrintObject.CurrentY, vbInches, vbPixels)
        RetVal = DrawText(PrintObject.hdc, s, Len(s), r, DT_CALCRECT Or DT_WORDBREAK)
        RetVal = DrawText(PrintObject.hdc, s, Len(s), r, DT_LEFT Or DT_WORDBREAK)
        PrintObject.CurrentY = PrintObject.ScaleY(r.Bottom, vbPixels, vbInches)
    Else
        PrintObject.Print s;
    End If
End Sub

Private Function IsHeading(ByVal s As String) As Boolean
    'Determines whether or not a line is the beginning of a method declaration
    s = Trim(s)
    If InStr(1, s, "private sub ", vbTextCompare) = 1 Or _
    InStr(1, s, "private function ", vbTextCompare) = 1 Or _
    InStr(1, s, "public sub ", vbTextCompare) = 1 Or _
    InStr(1, s, "public function ", vbTextCompare) = 1 Or _
    InStr(1, s, "friend sub ", vbTextCompare) = 1 Or _
    InStr(1, s, "friend function ", vbTextCompare) = 1 Or _
    InStr(1, s, "public property ", vbTextCompare) = 1 Or _
    InStr(1, s, "friend property ", vbTextCompare) = 1 Or _
    InStr(1, s, "property ", vbTextCompare) = 1 Or _
    InStr(1, s, "sub ", vbTextCompare) = 1 Or _
    InStr(1, s, "function ", vbTextCompare) = 1 Then
        IsHeading = True
    Else
        IsHeading = False
    End If
End Function

Private Sub Class_Initialize()
    'Set default settings for font size, styles, and printer object.
    Dim DefaultFont As String
    DefaultFont = "Helvetica"
    mBoxed = False
    mNewPage = True
    SetPaperSize 8.5, 11
    SetMargins 1, 1, 1, 1
    Set PrintObject = frmFiles.picParent(2)
    
    'Set up the styles
    ReDim mStyles(1 To 6)
    
    mStyles(stTitle).Name = DefaultFont
    mStyles(stTitle).Bold = True
    mStyles(stTitle).Size = 14
    
    mStyles(stHeader).Name = DefaultFont
    mStyles(stHeader).Bold = True
    mStyles(stHeader).Size = 12
    
    mStyles(stSubHeader).Name = DefaultFont
    mStyles(stSubHeader).Bold = True
    mStyles(stSubHeader).Size = 10
    
    mStyles(stPlain).Name = DefaultFont
    mStyles(stPlain).Size = 10
    
    mStyles(stItalic).Name = DefaultFont
    mStyles(stItalic).Size = 10
    mStyles(stItalic).Italic = True
    
    mStyles(stHeaderPrivate).Name = DefaultFont
    mStyles(stHeaderPrivate).Size = 12
    mStyles(stHeaderPrivate).Italic = True
    
End Sub

