Attribute VB_Name = "Demo"
Option Explicit

Dim Nvr As NvRam

Private Sub Main()
    Dim vByte As Byte, vInteger As Integer, vLong As Long
    Dim vBoolean As Boolean, vCurrency As Currency, vDate As Date
    Dim vSingle As Single, vDouble As Double
    Dim vString As String, vByteArray() As Byte
    Set Nvr = New NvRam
    
    MsgBox "It's suggested that you step through this main routine using Shift-F8." & vbCrLf _
      & "To start, hit CTRL-Break now." & vbCrLf & vbCrLf _
      & "You may also want a hex editor to view the data file (C:\temp\test.txt)."
    
    'Connect to the data file and clear all the values
    Nvr.Connect App.Path & "\test.txt"
    Nvr.ClearAll

    'Output some sample data at the specified "memory" locations
    Nvr.WriteVar "Fourscore and seven years ago...", 0, 10  'Truncated string of specified width
    Nvr.WriteVar CByte(1), 10
    Nvr.WriteVar CInt(2), 12
    Nvr.WriteVar CLng(3), 14
    Nvr.WriteVar CCur(4), 18
    Nvr.WriteVar CSng(5), 26
    Nvr.WriteVar CDbl(6), 34
    Nvr.WriteVar True, 42
    Nvr.WriteVar False, 43
    Nvr.WriteVar "Fourscore and seven years ago...", 44  'String of unspecified width

    'Now read back the sample data from the specified "memory" locations
    Nvr.ReadVar vString, 0, 10  'String of specified width
    MsgBox "'" & vString & "'"

    Nvr.ReadVar vByte, 10
    MsgBox vByte

    Nvr.ReadVar vInteger, 12
    MsgBox vInteger

    Nvr.ReadVar vLong, 14
    MsgBox vLong

    Nvr.ReadVar vCurrency, 18
    MsgBox vCurrency
    
    Nvr.ReadVar vSingle, 26
    MsgBox vSingle
    
    Nvr.ReadVar vDouble, 34
    MsgBox vDouble
    
    Nvr.ReadVar vBoolean, 42
    MsgBox vBoolean
    
    Nvr.ReadVar vBoolean, 43
    MsgBox vBoolean
    
    Nvr.ReadVar vString, 44  'String of unspecified width
    MsgBox "'" & vString & "'"
    
    'Read a variable written as a string using a byte array
    'to modify its contents and then write it back out.
    ReDim vByteArray(3)
    Nvr.ReadVar vByteArray, 0
    vByteArray(0) = vByteArray(0) Xor &H20  'LCase("F")
    vByteArray(1) = vByteArray(1) Xor &H20  'UCase("o")
    vByteArray(2) = vByteArray(2) Xor &H20  'UCase("u")
    vByteArray(3) = vByteArray(3) Xor &H20  'UCase("r")
    Nvr.WriteVar vByteArray, 0

    'All done; disconnect from the data file
    Nvr.Disconnect
End Sub

