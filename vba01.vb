'1.1

Global Const gKolRec As Integer = 1000   'const
Global gKolMag As Integer
Global gListMag() As String        'array
Global gListName() As String


Dim mStartMagRow As Integer
Dim mEndMagRow As Integer   
Dim mAdr As String   
Dim mMagInList As Boolean  
Dim mMagName As String 
Dim mCena As Single       
    
mMagInList = False
mMagName = ""
mItogListName = "Total"  

'---
'2.1 

If mSprMagListName <> "" Then
   Worksheets(Trim(mSprMagListName)).Select
End If


'2.2 

If Cells(i, 4).Value <> Empty Then
        '--
End If


'2.3

If mRetVal = vbNo Then
    Exit Sub      
End If

'2.4

If Trim(mTovName) = Trim(Cells(i, 2).Value) Then
  pGetMagTovRow = i
  Exit Function
End If

'2.5

If mTovInMagRow < 2 Then
    '--
End If

'2.6
If Cells(mStartRow - 1, i) = Empty And Cells(mStartRow - 2, i) = Empty Then
    '--
End If

'2.7

If Left(Trim(Cells(mEndRow + 2, 2).Value), 5) = "Total" And Left(Trim(Cells(mEndRow + 3, 2).Value), 5) = "Total" Then
    MsgBox ("...")
    Exit Sub
End If

'2.8
  
If Left(Trim(Cells(x, mIDCol)), 2) = "ds" Then
    '-- 
End If
                
'---         