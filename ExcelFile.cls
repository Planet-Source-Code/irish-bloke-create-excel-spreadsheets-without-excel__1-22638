VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExcelFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"

'Beginning Of File record
Private Type BOF
  opcode1 As Byte
  opcode2 As Byte
  length1 As Byte
  length2 As Byte
  version1 As Byte
  version2 As Byte
  ftype1 As Byte
  ftype2 As Byte
End Type

'End Of File record
Private Type EOF
  opcode1 As Byte
  opcode2 As Byte
  length1 As Byte
  length2 As Byte
End Type

'Integer record
Private Type tInteger
  opcode1 As Byte
  opcode2 As Byte
  length1 As Byte
  length2 As Byte
  row1 As Byte
  row2 As Byte
  col1 As Byte
  col2 As Byte
  rgbattr1 As Byte
  rgbAttr2 As Byte
  rgbAttr3 As Byte
  w1 As Byte
  w2 As Byte
End Type

'Label (Text) record
Private Type tLabel
  opcode1 As Byte
  opcode2 As Byte
  length1 As Byte
  length2 As Byte
  row1 As Byte
  row2 As Byte
  col1 As Byte
  col2 As Byte
  rgbattr1 As Byte
  rgbAttr2 As Byte
  rgbAttr3 As Byte
  length As Byte
End Type


Dim fhFile As Integer
Dim bof1 As BOF
Dim eof1 As EOF
Dim l1 As tLabel
Dim i1 As tInteger

Public Sub OpenFile(ByVal FileName As String)
fhFile = FreeFile
Open FileName For Binary As #fhFile
Put #fhFile, , bof1

End Sub

Public Sub CloseFile()

Put #fhFile, , eof1
Close #fhFile

End Sub


Private Sub Class_Initialize()
'Set up default values for records
'These should be the values that are the same for every record

With bof1
    .opcode1 = 9
    .opcode2 = 0
    .length1 = 4
    .length2 = 0
    .version1 = 2
    .version2 = 0
    .ftype1 = 10
    .ftype2 = 0
End With

With eof1
    .opcode1 = 10
End With

With l1
    .opcode1 = 4
    .opcode2 = 0
    .length1 = 10
    .length2 = 0
    .row2 = 0
    .col2 = 0
    .rgbattr1 = 0
    .rgbattr1 = 0
    .rgbattr1 = 0
    .length = 2
End With

With i1
  .opcode1 = 2
  .opcode2 = 0
  .length1 = 9
  .length2 = 0
  .row1 = 0
  .row2 = 0
  .col1 = 0
  .col2 = 0
  .rgbattr1 = 0
  .rgbAttr2 = 0
  .rgbAttr3 = 0
  .w1 = 0
  .w2 = 0
End With

End Sub

Function EWriteString(r As Byte, c As Byte, t As String)
Dim b As Byte
Dim l As Byte
stringtowrite = t
l = Len(stringtowrite)

'Length of the text portion of the record
l1.length = l

'Total length of the record
l1.length1 = 8 + l

'BIFF counts from zero
l1.row1 = r - 1
l1.col1 = c - 1

'Put record header
Put #fhFile, , l1
'Then the actual string data
For a = 1 To l
    b = Asc(Mid$(stringtowrite, a, 1))
    Put #fhFile, , b
Next

End Function

Function EWriteInteger(r As Byte, c As Byte, i As Long)
  
With i1
  .row1 = r - 1
  .col1 = c - 1
  .w1 = i - (Int(i / 256) * 256)
  .w2 = Int(i / 256)
End With

Put #fhFile, , i1

End Function
