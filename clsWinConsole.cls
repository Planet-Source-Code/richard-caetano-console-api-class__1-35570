' =============================================================================
'    Copyright (c) 2002 Richard Caetano - http://www.stronglytyped.com
'
'    Permission is hereby granted, free of charge, to any person obtaining a
'    copy of this software and associated documentation files (the "Software"),
'    to deal in the Software without restriction, including without limitation
'    the rights to use, copy, modify, merge, publish, distribute, sublicense,
'    and/or sell copies of the Software, and to permit persons to whom the
'    Software is furnished to do so, subject to the following conditions:
'
'    The above copyright notice and this permission notice shall be included
'    in all copies or substantial portions of the Software.
'
'    THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS
'    OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
'    FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL
'    THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
'    LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
'    OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
'    THE SOFTWARE.
'
'
'    clsWinConsole - A Visual Basic class that encapsulates calls to the Win32
'                    API for interacting with the Console
'
'
' =============================================================================


Option Explicit

Private Declare Function AllocConsole Lib "kernel32" () As Long

Private Declare Function FreeConsole Lib "kernel32" () As Long

Private Declare Function GetStdHandle Lib "kernel32" _
    (ByVal nStdHandle As Long) As Long

Private Declare Function ReadConsole Lib "kernel32" Alias "ReadConsoleA" ( _
    ByVal hConsoleInput As Long, _
    ByVal lpBuffer As String, _
    ByVal nNumberOfCharsToRead As Long, _
    lpNumberOfCharsRead As Long, _
    lpReserved As Any _
) As Long

Private Declare Function SetConsoleMode Lib "kernel32" ( _
    ByVal hConsoleOutput As Long, dwMode As Long _
) As Long

Private Declare Function SetConsoleTextAttribute Lib "kernel32" ( _
    ByVal hConsoleOutput As Long, _
    ByVal wAttributes As Long _
) As Long

Private Declare Function SetConsoleTitle Lib "kernel32" Alias "SetConsoleTitleA" ( _
    ByVal lpConsoleTitle As String _
) As Long

Private Declare Function WriteConsole Lib "kernel32" Alias "WriteConsoleA" ( _
    ByVal hConsoleOutput As Long, _
    ByVal lpBuffer As Any, _
    ByVal nNumberOfCharsToWrite As Long, _
    lpNumberOfCharsWritten As Long, _
    lpReserved As Any _
) As Long

' Console Handlers
Private Const STD_INPUT_HANDLE = -10&
Private Const STD_OUTPUT_HANDLE = -11&
Private Const STD_ERROR_HANDLE = -12&

' SetConsoleTextAttribute Color values
Private Const FOREGROUND_BLUE = &H1
Private Const FOREGROUND_GREEN = &H2
Private Const FOREGROUND_RED = &H4
Private Const FOREGROUND_INTENSITY = &H8
Private Const BACKGROUND_BLUE = &H10
Private Const BACKGROUND_GREEN = &H20
Private Const BACKGROUND_RED = &H40
Private Const BACKGROUND_INTENSITY = &H80

' For SetConsoleMode (input)
Private Const ENABLE_LINE_INPUT = &H2
Private Const ENABLE_ECHO_INPUT = &H4
Private Const ENABLE_MOUSE_INPUT = &H10
Private Const ENABLE_PROCESSED_INPUT = &H1
Private Const ENABLE_WINDOW_INPUT = &H8

' For SetConsoleMode (output)
Private Const ENABLE_PROCESSED_OUTPUT = &H1
Private Const ENABLE_WRAP_AT_EOL_OUTPUT = &H2

' Handles to the console
Private hConsoleIn As Long
Private hConsoleOut As Long
Private hConsoleErr As Long

Private Sub Class_Terminate()
    Hide
End Sub

Public Sub Show()
'
'   Show the console window
'
    AllocConsole
    SetConsoleTitle App.Title
    
    hConsoleIn = GetStdHandle(STD_INPUT_HANDLE)
    hConsoleOut = GetStdHandle(STD_OUTPUT_HANDLE)
    hConsoleErr = GetStdHandle(STD_ERROR_HANDLE)

End Sub

Public Sub Hide()
'
'   If the console window is active then free it
'
    If Active Then
        FreeConsole
    End If

End Sub

Public Function Active() As Boolean
'
'   Returns True if the console is open
'
    Active = (hConsoleIn <> 0) Or (hConsoleOut <> 0) Or (hConsoleErr <> 0)

End Function

Public Sub WriteText(text As String)
'
'   Sends the text out to the console
'
    WriteConsole hConsoleOut, text, Len(text), vbNull, vbNull

End Sub


Public Sub WriteTextLine(ByVal text As String)
'
'   Sends the text out to the console
'
    text = text & vbNewLine
    WriteConsole hConsoleOut, text, Len(text), vbNull, vbNull

End Sub

Public Function ReadText() As String
'
'   Prompts the user for input
'
    Dim sUserInput As String * 256
    
    Call ReadConsole(hConsoleIn, sUserInput, Len(sUserInput), vbNull, vbNull)
    ReadText = Left$(sUserInput, InStr(sUserInput, Chr$(0)) - 3)

End Function


