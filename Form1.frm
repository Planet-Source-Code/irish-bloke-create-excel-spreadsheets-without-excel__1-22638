VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Excel Class Sample"
   ClientHeight    =   3795
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6270
   LinkTopic       =   "Form1"
   ScaleHeight     =   3795
   ScaleWidth      =   6270
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "create my excel spreadsheet"
      Height          =   540
      Left            =   3570
      TabIndex        =   0
      Top             =   3045
      Width           =   2325
   End
   Begin VB.Label Label2 
      Caption         =   "http://go.to/wdwtbam"
      Height          =   330
      Left            =   105
      TabIndex        =   2
      Top             =   3255
      Width           =   2010
   End
   Begin VB.Label Label1 
      Caption         =   "sample class allows you to create excel spreedsheest without using excel or any automation!"
      Height          =   960
      Left            =   105
      TabIndex        =   1
      Top             =   210
      Width           =   4740
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'http://www.vb-helper.com/HowTo/doexcel.zip
''   Purpose
'Write Excel files directly without using Excel
''   Method
'See the code and other text files.
'Thanks to Dan Gardner (nambit@hotmail.com).
'   Disclaimer
'This example program is provided "as is" with no warranty of any kind. It is
'intended for demonstration purposes only. In particular, it does no error
'handling. You can use the example in any form, but please mention
'www.vb-helper.com.

'MODIFIES BY GERRY MC DONNELL gerrymcd@oceanfree.net
'i take no credit for this code
'im just sending to psc.com b coz is very easy and handy
'usefull code

Private Sub Command1_Click()
Dim colu As Byte
Dim rw As Byte

'create new excel class called ef1
Dim ef1 As New ExcelFile

With ef1
'open it path will be c:\vbtest.xls
    .OpenFile "vbtest.xls"
    'write integer data @ col 1, row 1
    .EWriteInteger 1, 1, 100
    'write a string @ col 2, row 1
    .EWriteString 1, 2, "Test writing a string"
    'write another string @ col 3, row 1
    .EWriteString 1, 3, "gerry"
    .CloseFile
End With

End Sub




Private Sub Label2_Click()
Shell "start http://go.to/wdwtbam"
End Sub
