VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "String Demo 1"
   ClientHeight    =   3705
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9705
   LinkTopic       =   "Form1"
   ScaleHeight     =   3705
   ScaleWidth      =   9705
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command22 
      Caption         =   "StrLen"
      Height          =   375
      Left            =   2670
      TabIndex        =   25
      Top             =   3030
      Width           =   1530
   End
   Begin VB.ListBox List1 
      Height          =   2790
      Left            =   7860
      TabIndex        =   24
      Top             =   570
      Width           =   1545
   End
   Begin VB.CommandButton Command17 
      Caption         =   "Str Split"
      Height          =   390
      Left            =   7845
      TabIndex        =   23
      Top             =   75
      Width           =   1545
   End
   Begin VB.CommandButton Command21 
      Caption         =   "Exit"
      Height          =   390
      Left            =   6435
      TabIndex        =   22
      Top             =   3030
      Width           =   1350
   End
   Begin VB.CommandButton Command20 
      Caption         =   "CharToAsc(Str)"
      Height          =   375
      Left            =   2670
      TabIndex        =   21
      Top             =   2595
      Width           =   1560
   End
   Begin VB.CommandButton Command19 
      Caption         =   "StrReplace Text"
      Height          =   375
      Left            =   6240
      TabIndex        =   20
      Top             =   2115
      Width           =   1560
   End
   Begin VB.CommandButton Command18 
      Caption         =   "StrReplaceChar"
      Height          =   375
      Left            =   4380
      TabIndex        =   19
      Top             =   2130
      Width           =   1560
   End
   Begin VB.CommandButton Command16 
      Caption         =   "About"
      Height          =   390
      Left            =   6435
      TabIndex        =   18
      Top             =   2595
      Width           =   1365
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   195
      TabIndex        =   17
      Top             =   855
      Width           =   2295
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   195
      TabIndex        =   16
      Top             =   495
      Width           =   2295
   End
   Begin VB.CommandButton Command15 
      Caption         =   "StrPad"
      Height          =   375
      Left            =   6240
      TabIndex        =   15
      Top             =   1650
      Width           =   1560
   End
   Begin VB.CommandButton Command14 
      Caption         =   "StrGetToken"
      Height          =   375
      Left            =   4380
      TabIndex        =   14
      Top             =   1650
      Width           =   1560
   End
   Begin VB.CommandButton Command13 
      Caption         =   "StrReverse"
      Height          =   375
      Left            =   2670
      TabIndex        =   13
      Top             =   2160
      Width           =   1560
   End
   Begin VB.CommandButton Command12 
      Caption         =   "RemoveStrRight(Str,3)"
      Height          =   375
      Left            =   4335
      TabIndex        =   12
      Top             =   3030
      Width           =   1875
   End
   Begin VB.CommandButton Command11 
      Caption         =   "RemoveStrLeft(Str,3)"
      Height          =   375
      Left            =   4335
      TabIndex        =   11
      Top             =   2595
      Width           =   1875
   End
   Begin VB.CommandButton Command10 
      Caption         =   "InstrRight(Str,""#"",5)"
      Height          =   375
      Left            =   2670
      TabIndex        =   10
      Top             =   1650
      Width           =   1560
   End
   Begin VB.CommandButton Command9 
      Caption         =   "InstrLeft(Str,""#"",5)"
      Height          =   375
      Left            =   6240
      TabIndex        =   9
      Top             =   1110
      Width           =   1560
   End
   Begin VB.CommandButton Command8 
      Caption         =   "IsStringSame"
      Height          =   375
      Left            =   4380
      TabIndex        =   8
      Top             =   1110
      Width           =   1560
   End
   Begin VB.CommandButton Command7 
      Caption         =   "IsAllChar"
      Height          =   375
      Left            =   2670
      TabIndex        =   7
      Top             =   1110
      Width           =   1560
   End
   Begin VB.CommandButton Command6 
      Caption         =   "StrUCase(Str,8)"
      Height          =   375
      Left            =   6240
      TabIndex        =   6
      Top             =   585
      Width           =   1560
   End
   Begin VB.CommandButton Command5 
      Caption         =   "StrLCase(Str,18)"
      Height          =   375
      Left            =   4380
      TabIndex        =   5
      Top             =   585
      Width           =   1560
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Strip Numeric"
      Height          =   375
      Left            =   2670
      TabIndex        =   4
      Top             =   585
      Width           =   1560
   End
   Begin VB.CommandButton Command3 
      Caption         =   "StripNone Numeric"
      Height          =   375
      Left            =   6240
      TabIndex        =   3
      Top             =   105
      Width           =   1560
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Str ProperCase"
      Height          =   375
      Left            =   4380
      TabIndex        =   2
      Top             =   105
      Width           =   1560
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   195
      TabIndex        =   1
      Top             =   180
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Str Encrypt"
      Height          =   375
      Left            =   2670
      TabIndex        =   0
      Top             =   105
      Width           =   1560
   End
   Begin VB.Image Image1 
      Height          =   2325
      Left            =   -15
      Picture         =   "Form1.frx":0000
      Top             =   1245
      Width           =   2640
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim TArray() As String

Function StrSplit(StrBuff As String, Dilmiter As String) As String
Dim Cnt As Integer, I As Integer
Dim G As String, M As String
I = -1
If Len(StrBuff) <= 0 Then Exit Function
    If Len(Dilmiter) <= 0 Then
        StrSplit = StrBuff
        Exit Function
    Else
        If Not Right(StrBuff, 1) = Dilmiter Then StrBuff = Left(StrBuff, Len(StrBuff)) & Dilmiter
        If Left(StrBuff, 1) = Dilmiter Then StrBuff = Right(StrBuff, Len(StrBuff) - 1)
        
        For Cnt = 1 To Len(StrBuff)
            ch = Mid(StrBuff, Cnt, 1)
            M = M & ch
            If ch = Dilmiter Then
                I = I + 1
                ReDim Preserve TArray(I)
                G = G & Left(M, Len(M) - 1)
                TArray(I) = G
                G = ""
                M = ""
            End If
        Next
        End If
        
End Function


Function ChrToAsc(TChar As String) As Integer
    If Len(TChar) <= 0 Then Exit Function
    ChrToAsc = Asc(TChar)
    
End Function
Function StrReplace(StrBuff, StrFind, StrReplaceWith) As String
Dim Xpos As Integer
Dim NewStr As String
    Do While Len(StrBuff) > 0
        Xpos = InStr(StrBuff, StrFind)
        If Xpos = 0 Then
            NewStr = NewStr & StrBuff
            StrBuff = ""
        Else
            NewStr = NewStr & Left(StrBuff, Xpos - 1) & StrReplaceWith
            StrBuff = Mid(StrBuff, Xpos + Len(StrFind))
        End If
    Loop
    StrReplace = NewStr
    
End Function
Function StrEncrypt(lzStr As String, Key As String)
Dim n As Integer
    If Len(lzStr) <= 0 Then Exit Function
    If Len(Key) <= 0 Then StrEncrypt = lzStr: Exit Function
    For n = 1 To Len(lzStr)
        ch = Mid(lzStr, n, 1)
        Mid(lzStr, n, 1) = Chr(Asc(ch) + Len(Key))
    Next
    StrEncrypt = lzStr
    n = 0
    
End Function
Function StrDencrypt(lzStr As String, Key As String)
Dim n As Integer
    If Len(lzStr) <= 0 Then Exit Function
    If Len(Key) <= 0 Then StrDencrypt = lzStr: Exit Function
    For n = 1 To Len(lzStr)
        ch = Mid(lzStr, n, 1)
        Mid(lzStr, n, 1) = Chr(Asc(ch) - Len(Key))
    Next
    StrDencrypt = lzStr
    n = 0
    
End Function

Function StrProperCase(lzStr As String) As String
Dim n As Integer, I As Integer
On Error Resume Next
If Len(lzStr) <= 0 Then Exit Function

    For I = 1 To Len(lzStr)
        ch = Asc(Mid(lzStr, I, 1))
        If ch = 32 Then
            n = I
            Mid(lzStr, 1, 1) = UCase(Mid(lzStr, 1, 1))
            Mid(lzStr, n + 1, 1) = UCase(Mid(lzStr, n + 1, 1))
        End If
    Next
    StrProperCase = lzStr
    n = 0
    I = 0
    
End Function

Function IsStringSame(String1 As String, String2 As String) As Boolean

    If String1 = String2 Then
        IsStringSame = True
    Else
        IsStringSame = False
    End If
    
End Function
Function InstrRight(lzStr As String, TString As String, Length As String)
    If Len(lzStr) < 0 Then Exit Function
    If Len(TString) <= 0 Then Exit Function
    If Length <= 0 Then
        InstrRight = lzStr
        Exit Function
    Else
        InstrRight = lzStr & String(Length, TString)
    End If
    
End Function
Function InstrLeft(lzStr As String, TString As String, Length As String)
    If Len(lzStr) < 0 Then Exit Function
    If Len(TString) <= 0 Then Exit Function
    If Length <= 0 Then
        InstrLeft = lzStr
        Exit Function
    Else
        InstrLeft = String(Length, TString) & lzStr
    End If
End Function

Function StrReplaceChar(lzStr As String, RemoveChar As String, ReplaceChar As String)
Dim Xpos As Integer
Dim NewStr As String

    If Len(lzStr) <= 0 Then Exit Function
        For Xpos = 1 To Len(lzStr)
            ch = Mid(lzStr, Xpos, 1)
            If ch = RemoveChar Then
                NewStr = NewStr & ReplaceChar
            Else
                NewStr = NewStr & ch
            End If
        Next
        StrReplaceChar = NewStr
        Xpos = 0
End Function

Function StrReverse(lzStr As String) As String
Dim Xpos As Integer
Dim NewStr As String
    If Len(lzStr) <= 0 Then Exit Function
    For Xpos = Len(lzStr) To 1 Step -1
        ch = Mid(lzStr, Xpos, 1)
            NewStr = NewStr & ch
        Next
    StrReverse = NewStr
    
End Function
Function RemoveStrLeft(lzStr As String, TLength As Integer)
On Error Resume Next
    If TLength = 0 Then Exit Function
    If Len(lzStr) <= 0 Then Exit Function
    If TLength > Len(lzStr) Then
        RemoveStrLeft = lzStr
    Else
        RemoveStrLeft = Right(lzStr, Len(lzStr) - TLength)
    End If
    
End Function
Function RemoveStrRight(lzStr As String, TLength As Integer)
On Error Resume Next
    If TLength = 0 Then Exit Function
    If Len(lzStr) <= 0 Then Exit Function
    If TLength > Len(lzStr) Then
        RemoveStrRight = lzStr
    Else
        RemoveStrRight = Left(lzStr, Len(lzStr) - TLength)
    End If
    
End Function
Function StrUCase(lzStr As String, Index As Integer) As String
Dim NewStr As String
On Error Resume Next
    If Index = 0 Then Exit Function
    If Len(lzStr) <= 0 Then Exit Function
    NewStr = lzStr
    Mid(NewStr, Index, 1) = UCase(Mid(NewStr, Index, 1))
    If Err Then
        StrUCase = lzStr
        Exit Function
    Else
        StrUCase = NewStr
    End If
    NewStr = ""
End Function

Function StrLCase(lzStr As String, Index As Integer) As String
Dim NewStr As String
On Error Resume Next
    If Index = 0 Then Exit Function
    If Len(lzStr) <= 0 Then Exit Function
    NewStr = lzStr
    Mid(NewStr, Index, 1) = LCase(Mid(NewStr, Index, 1))
    If Err Then
        StrLCase = lzStr
        Exit Function
    Else
        StrLCase = NewStr
    End If
    NewStr = ""
End Function

Function IsAllChar(lzStr As String) As Integer
Dim Xpos As Integer

    If Len(lzStr) <= 0 Then Exit Function
    For Xpos = 1 To Len(lzStr)
        ch = Mid(lzStr, Xpos, 1)
            If ch Like "[A-Z a-z -.;'#=+?!£$%^&*()]" Then
                IsAllChar = 1
            Else
                IsAllChar = 0
                Exit For
            End If
        Next
        
End Function
Function StripNumeric(lzStr As String) As Variant
Dim Xpos As Integer
Dim num As String
    If Len(lzStr) <= 0 Then Exit Function
    For Xpos = 1 To Len(lzStr)
        ch = Mid(lzStr, Xpos, 1)
        If ch Like "[0-9]" Then
           num = num & ch
           
        End If
    Next
    StripNumeric = Val(num)
    
End Function

Function StripNoneNumeric(lzStr As String) As String
Dim Xpos As Integer
Dim StrB As String

    If Len(lzStr) <= 0 Then Exit Function
    For Xpos = 1 To Len(lzStr)
        ch = Mid(lzStr, Xpos, 1)
        If ch Like "[a-z A-Z -.;'#=+?!£$%^&*()]" Then
           StrB = StrB & ch
        End If
    Next
    StripNoneNumeric = StrB
    
End Function

Function StrPad(lzStr As String, PadWith As String) As String
Dim Xpos As Integer
Dim NewStr As String
    If Len(lzStr) <= 0 Then Exit Function
    If Len(PadWith) <= 0 Then Exit Function
    For Xpos = 1 To Len(lzStr)
        ch = Mid(lzStr, Xpos, 1)
        NewStr = NewStr & PadWith & ch
    Next
    
    If Left(NewStr, 1) = PadWith Then NewStr = Right(NewStr, Len(NewStr) - 1)
    StrPad = NewStr
    
End Function
Function StrGetToken(lzStr As String, Delmiter1 As String, Delmiter2 As String) As String
Dim X1, X2 As Integer

    If Len(lzStr) <= 0 Then
        Exit Function
    Else
        X1 = InStr(lzStr, Delmiter1)
        X2 = InStr(X1 + 1, lzStr, Delmiter2)
        If X1 = 0 Or X2 = 0 Then
            StrGetToken = lzStr
            Exit Function
        Else
            StrGetToken = Mid(lzStr, X1 + 1, X2 - X1 - 1)
        End If
    End If
    
    
End Function



Private Sub Command1_Click()
    Text1 = "Visual Basic 6.0"
    Text2 = StrEncrypt(Text1, "mypass")
    Text3 = StrDencrypt(Text2, "mypass")
    
End Sub

Private Sub Command10_Click()
    Text3 = ""
    Text1 = "Visual J"
    Text2 = InstrRight(Text1, "#", 5)
    
End Sub

Private Sub Command11_Click()
    Text3 = ""
    Text1 = "Visual Basic 6.0"
    Text2 = RemoveStrLeft(Text1, 3)
    
End Sub

Private Sub Command12_Click()
    Text3 = ""
    Text1 = "Visual Basic 6.0"
    Text2 = RemoveStrRight(Text1, 3)
    
End Sub

Private Sub Command13_Click()
    Text3 = ""
    Text1 = "Welcome to my Home Page"
    Text2 = StrReverse(Text1)
    
End Sub

Private Sub Command14_Click()
    Text3 = ""
    Text1 = "Value=(&H125)"
    Text2 = StrGetToken(Text1, "(", ")")
    
End Sub

Private Sub Command15_Click()
    Text3 = ""
    Text1 = "Padded String"
    Text2 = StrPad(Text1, "-")
    
End Sub



Private Sub Command16_Click()
    MsgBox "Simple String Demo by Ben Jones Please Vote", vbInformation, "About"
    
End Sub



Private Sub Command17_Click()
    List1.Clear
    Text1 = "Microsoft|Visual|Basic|For|Windows"
    Text2 = "|"
    
    StrSplit Text1, Text2
    
    For I = LBound(TArray) To UBound(TArray)
        List1.AddItem TArray(I)
    Next
    
    
End Sub

Private Sub Command18_Click()
    Text1 = "-----Visual Basic------"
    Text2 = "0"
    Text3 = StrReplaceChar(Text1, "-", Text2)
    
End Sub

Private Sub Command19_Click()
    Text1 = "Visual Basic for Windows"
    Text2 = "C++"
    Text3 = StrReplace(Text1, "Basic", Text2)
    
End Sub

Private Sub Command2_Click()
    Text1 = "this is a test"
    Text2 = StrProperCase(Text1.Text)
    Text3 = ""
    
End Sub

Private Sub Command20_Click()
    Text3 = ""
    Text1 = "A"
    Text2 = ChrToAsc(Text1)
    
End Sub

Private Sub Command21_Click()
    Unload Me: End
    
End Sub

Private Sub Command22_Click()
    Text1 = "Length Test"
    Text2 = Len(Text1)
    
End Sub

Private Sub Command23_Click()
    
    
End Sub

Private Sub Command3_Click()
    Text1 = "01252-Hello-world1258"
    Text2 = StripNoneNumeric(Text1)
    Text3 = ""
    
End Sub

Private Sub Command4_Click()
    Text1 = "125858VisualBasic--+++99911"
    Text2 = StripNumeric(Text1)
    Text3 = ""
    
End Sub

Private Sub Command5_Click()
    Text1 = "microsoft visual Basic"
    Text2 = StrLCase(Text1, 18)
    Text3 = ""
    
End Sub

Private Sub Command6_Click()
    Text1 = "visual c++"
    Text2 = StrUCase(Text1, 8)
    Text3 = ""
    
End Sub

Private Sub Command7_Click()
    Text3 = ""
    Text1 = "planet9-source-code"
    If IsAllChar(Text1) = 0 Then
        Text2 = "FALSE"
    Else
        Text2 = "TRUE"
    End If
    
    
End Sub

Private Sub Command8_Click()
    Text3 = ""
    Text1 = "borland"
    Text2 = "borLand"
    If IsStringSame(Text1, Text2) = False Then
        Text3 = "FALSE"
    Else
        Text3 = "TRUE"
    End If
    
    
End Sub

Private Sub Command9_Click()
    Text3 = ""
    Text1 = "Visual J"
    Text2 = InstrLeft(Text1, "#", 5)
    
End Sub

