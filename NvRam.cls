VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NvRam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'################################################################
' Non-Volatile RAM for VB Apps
' Created 6 June 2000 by James Vincent Carnicelli
'
' Notes:
' The purpose of this class is to let your apps maintain data
' in random access disk files, instead of memory.  This has
' four primarily uses: managing large amounts of data,
' protecting critical data from accidental shutdowns,
' maintaining program state across run-time sessions, and
' sharing data among multiple processes.
'
' Refer to .ReadVar() and .WriteVar() for info on what data
' types are supported and how many bytes they take up.
'################################################################

Option Explicit


'##################### Private Declarations #####################

'Windows API Stuff
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)


'Storage for public properties
Private propMemorySize As Long
Private propFileName As String

'Application data
Private hFile As Integer



'###################### Public Declarations #####################

Public Property Get FileName() As String
    MemorySize = propMemorySize
End Property
Public Property Let FileName(newFileName As String)
    If hFile <> 0 Then Err.Raise vbObjectError, "NvRam.FileName", _
      "Can't change the file name while connected"
End Property

Public Property Get MemorySize() As Long
    CheckConnected "MemorySize"
    propMemorySize = LOF(hFile)
    MemorySize = propMemorySize
End Property
Public Property Let MemorySize(newMemorySize As Long)
    If propMemorySize = newMemorySize Then Exit Property
    Err.Raise vbObjectError, "NvRam.MemorySize", "Feature not yet implemented"
    
    'To do: resize file
'    propMemorySize = newMemorySize
End Property


'####################### Public Properties ######################


'######################## Public Methods ########################

Public Sub Connect(Optional FileName)
    Disconnect
    If Not IsMissing(FileName) Then propFileName = FileName
    hFile = FreeFile
    Open FileName For Binary As hFile Len = 1
    propMemorySize = LOF(hFile)
End Sub

Public Sub Disconnect()
    If hFile = 0 Then Exit Sub
    Close #hFile
    hFile = 0
End Sub

Public Sub ClearAll()
    Dim i As Long
    CheckConnected "ClearAll()"
    For i = 1 To LOF(hFile) - 1
        Put #hFile, i, 0
    Next
End Sub

Public Sub ReadVar(Variable, ByVal StartByte As Long, Optional Size)
    Dim bByte As Byte, i As Long, lSize As Long, sChar As String
    Dim Bytes() As Byte, nValue As Integer, lValue As Long, curValue As Currency
    Dim sValue As Single, dValue As Double, tValue As Date
    Dim nNumberOfBytes As Integer
    CheckConnected "ReadVar()"
    
    'If the size is specified for anything but a string, let's assume
    'this is for a data structure, and the user knows what he's doing.
    If Not IsMissing(Size) And TypeName(Variable) <> "String" Then
        lSize = Size
        ReDim Bytes(lSize - 1)
        ChewBytes Bytes, StartByte
        CopyMemory Variable, Bytes(0), lSize
    
    'Otherwise, we'll do this the safe way
    Else
        
        Select Case TypeName(Variable)
            
            Case "Integer"  '2 bytes
                ReDim Bytes(1)
                ChewBytes Bytes, StartByte
                CopyMemory nValue, Bytes(0), 2
                Variable = nValue
            
            Case "Long"  '4 bytes
                ReDim Bytes(3)
                ChewBytes Bytes, StartByte
                CopyMemory lValue, Bytes(0), 4
                Variable = lValue
            
            Case "Single"  '4 bytes
                ReDim Bytes(3)
                ChewBytes Bytes, StartByte
                CopyMemory sValue, Bytes(0), 4
                Variable = sValue
            
            Case "Double"  '8 bytes
                ReDim Bytes(7)
                ChewBytes Bytes, StartByte
                CopyMemory dValue, Bytes(0), 8
                Variable = dValue
            
            Case "Currency"  '8 bytes
                ReDim Bytes(7)
                ChewBytes Bytes, StartByte
                CopyMemory curValue, Bytes(0), 8
                Variable = curValue
            
            Case "Date"  '8 bytes
                ReDim Bytes(7)
                ChewBytes Bytes, StartByte
                CopyMemory tValue, Bytes(0), 8
                Variable = tValue
            
            Case "Boolean"  '1 byte
                Variable = 0
                Get #hFile, StartByte + 1, bByte
                Variable = (bByte <> 0)
            
            Case "Byte"  '1 byte
                Variable = 0
                Get #hFile, StartByte + 1, bByte
                Variable = bByte
            
            Case "Byte()"  'UBound(Variable) + 1 or Size bytes
                If Not IsMissing(Size) Then
                    ReDim Preserve Variable(Size - 1)
                End If
                ChewBytes Variable, StartByte
            
            Case "String"  'Len(Variable) or Size bytes
                If IsMissing(Size) Then  'Unspecified size
                    'If Size is not specified, the first four bytes of the
                    'block are assumed to be a Long indicating how many bytes
                    'are in the string.  So its LLLLSSSSSS... where L is
                    'one byte in a Long and S is one byte in the data string.
                    lSize = Len(Variable)
                    ReadVar lSize, StartByte
                    Variable = String(lSize, " ")
                    For i = 1 To lSize
                        Get #hFile, StartByte + i + 4, bByte
                        Mid(Variable, i, 1) = Chr(bByte)
                    Next
                Else  'Specified size
                    Variable = String(Size, " ")
                    For i = 1 To Size
                        Get #hFile, StartByte + i, bByte
                        If bByte = 0 Then
                            Variable = Left(Variable, i - 1)
                            Exit For
                        End If
                        Mid(Variable, i, 1) = Chr(bByte)
                    Next
                End If
                
            Case Else  'Unknown number of bytes or treatment
                If IsMissing(Size) Then Err.Raise vbObjectError, "NvRam.WriteVar()", _
                  "Don't know how to input a """ & TypeName(Variable) & """"
                Get #hFile, StartByte + 1, bByte
                ReDim Bytes(Size - 1)
                ChewBytes Bytes, StartByte
                CopyMemory Variable, Bytes(0), Size
        End Select
    End If
End Sub

Public Sub WriteVar(Variable, ByVal StartByte As Long, Optional Size)
    Dim bByte As Byte, i As Long, lSize As Long, Bytes() As Byte
    Dim nValue As Integer, lValue As Long, sValue As Single
    Dim dValue As Double, tValue As Date, curValue As Currency
    CheckConnected "WriteVar()"
    
    'If the size is specified for anything but a string, let's assume
    'this is for a data structure, and the user knows what he's doing.
    If Not IsMissing(Size) And TypeName(Variable) <> "String" Then
        lSize = Size
        ReDim Bytes(lSize - 1)
        CopyMemory Bytes(0), Variable, lSize
        SpewBytes Bytes, StartByte
    
    'Otherwise, we'll do this the safe way
    Else
    
        Select Case TypeName(Variable)
            
            Case "Boolean"  '1 byte
                ReDim Bytes(0)
                If Variable Then
                    Bytes(0) = 1
                Else
                    Bytes(0) = 0
                End If
                SpewBytes Bytes, StartByte
            
            Case "Byte"  '1 byte
                ReDim Bytes(0)
                Bytes(0) = Variable
                SpewBytes Bytes, StartByte
            
            Case "Integer"  '2 bytes
                lSize = 2
                ReDim Bytes(lSize - 1)
                nValue = Variable
                CopyMemory Bytes(0), nValue, lSize
                SpewBytes Bytes, StartByte
            
            Case "Long"  '4 bytes
                lSize = 4
                ReDim Bytes(lSize - 1)
                lValue = Variable
                CopyMemory Bytes(0), lValue, lSize
                SpewBytes Bytes, StartByte
                
            Case "Single"  '4 bytes
                lSize = 4
                ReDim Bytes(lSize - 1)
                sValue = Variable
                CopyMemory Bytes(0), sValue, lSize
                SpewBytes Bytes, StartByte
            
            Case "Double"  '8 bytes
                lSize = 8
                ReDim Bytes(lSize - 1)
                dValue = Variable
                CopyMemory Bytes(0), dValue, lSize
                SpewBytes Bytes, StartByte
            
            Case "Currency"  '8 bytes
                lSize = 8
                ReDim Bytes(lSize - 1)
                curValue = Variable
                CopyMemory Bytes(0), curValue, lSize
                SpewBytes Bytes, StartByte
            
            Case "Date"  '8 bytes
                lSize = 8
                ReDim Bytes(lSize - 1)
                tValue = Variable
                CopyMemory Bytes(0), tValue, lSize
                SpewBytes Bytes, StartByte
            
            Case "Byte()"  'UBound(Variable) bytes
                Bytes = Variable
                SpewBytes Bytes, StartByte
            
            Case "String"  'Len(Variable) or Size bytes
                lSize = Len(Variable)
                If IsMissing(Size) Then
                    'If Size is not specified, the first four bytes of the
                    'block will be a Long indicating how many bytes
                    'are in the string.  So its LLLLSSSSSS... where L is
                    'one byte in a Long and S is one byte in the data string.
                    'So the number of bytes used will be the length of the
                    'string plus 4 bytes.
                    WriteVar lSize, StartByte
                    For i = 1 To lSize
                        bByte = Asc(Mid(Variable, i, 1))
                        Put #hFile, StartByte + i + 4, bByte
                    Next
                Else
                    WriteVar lSize, StartByte
                    For i = 1 To Size
                        If i > lSize Then
                            bByte = 0
                        Else
                            bByte = Asc(Mid(Variable, i, 1))
                        End If
                        Put #hFile, StartByte + i, bByte
                    Next
                End If
                
            Case Else  'Unknown number of bytes or treatment
                Err.Raise vbObjectError, "NvRam.WriteVar()", _
                  "Don't know how to output a """ & TypeName(Variable) & """"
        End Select
    End If
End Sub


'######################## Private Methods ########################

Private Sub CheckConnected(ByVal Member As String)
    If hFile = 0 Then Err.Raise vbObjectError, "NvRam." & Member, _
      "Can't perform operation without being connected to file"
End Sub

Private Sub SpewBytes(Bytes, ByVal StartByte As Long)
    Dim i As Integer, bByte As Byte
    For i = UBound(Bytes) To 0 Step -1
        bByte = Bytes(i)
        Put #hFile, StartByte + i + 1, bByte
    Next
End Sub

Private Sub ChewBytes(Bytes, ByVal StartByte As Long)
    Dim i As Integer, bByte As Byte
    For i = UBound(Bytes) To 0 Step -1
        Get #hFile, StartByte + i + 1, bByte
        Bytes(i) = bByte
    Next
End Sub


'##################### Private Event Handlers ####################

Private Sub Class_Initialize()
    hFile = 0
End Sub

Private Sub Class_Terminate()
    Disconnect
End Sub
