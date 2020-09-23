Attribute VB_Name = "modMain"
Option Explicit

Public Files() As String

Public MyPrinter As clsPrintFormatted
Public Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Public Const DT_CALCRECT = &H400
Public Const DT_LEFT = &H0
Public Const DT_WORDBREAK = &H10

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Type typeFont
    Name As String
    Size As Long
    Bold As Boolean
    Italic As Boolean
    Underline As Boolean
End Type

Public Styles() As typeFont

Public Enum enumStyles
    stTitle = 1
    stHeader = 2
    stSubHeader = 3
    stPlain = 4
    stItalic = 5
    stHeaderPrivate = 6
End Enum

Public Enum enumJustification
    Left = 0
    Center = 1
    Right = 2
End Enum

Public Sub Main()
    'Initialize the printer class, show the form
    Set MyPrinter = New clsPrintFormatted
    frmFiles.Show
End Sub

Public Sub ShutDown()
    'Release reference to printer class, unload the form
    Set MyPrinter = Nothing
    Unload frmFiles
End Sub

Public Function LoadFile(ByVal FileName As String) As String
    'Returns the entire file in a string
    Dim FileNo As Integer, l As Long, s As String
    FileNo = FreeFile
    l = FileLen(FileName)
    Open FileName For Input As #FileNo
    s = Input(l, #FileNo)
    Close #FileNo
    LoadFile = s
End Function

Public Function LastPart(ByVal StrData As String) As String
    'Returns the last part of a filename.
    Dim s As String, i As Long, max As Long, c As String
    max = Len(StrData)
    For i = max To 1 Step -1
        c = Mid$(StrData, i, 1)
        If c = "/" Or c = "\" Or c = ":" Then
            'found the last part
            Exit For
        End If
        s = c & s
    Next i
    LastPart = s
End Function


