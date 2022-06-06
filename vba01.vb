'2.1

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
'3.1 

If mSprMagListName <> "" Then
   Worksheets(Trim(mSprMagListName)).Select
End If


'3.2 

If Cells(i, 4).Value <> Empty Then
        '--
End If


'3.3

If mRetVal = vbNo Then
    Exit Sub      
End If

'3.4

If Trim(mTovName) = Trim(Cells(i, 2).Value) Then
  pGetMagTovRow = i
  Exit Function
End If

'3.5

If mTovInMagRow < 2 Then
    '--
End If

'3.6
If Cells(mStartRow - 1, i) = Empty And Cells(mStartRow - 2, i) = Empty Then
    '--
End If

'3.7

If Left(Trim(Cells(mEndRow + 2, 2).Value), 5) = "Total" And Left(Trim(Cells(mEndRow + 3, 2).Value), 5) = "Total" Then
    MsgBox ("...")
    Exit Sub
End If

'3.8
  
If Left(Trim(Cells(x, mIDCol)), 2) = "ds" Then
    '-- 
End If
                
'---         
                
'4.1

  For x = 1 To mNumRows
         
         If Left(Trim(Cells(x, mIDCol)), 2) = "ds" Then
             y = mStartCol
             mEndCol = mStartCol + 30
             For y = mStartCol To mEndCol
                Cells(x, y) = Empty
                Cells(x, y).Value = Empty
                Cells(x, y).ClearContents
             Next
         End If
         
  Next

        
'4.2

    Sheets(mSprMagListName).Select

    For i = mStartMagRow To mEndMagRow
       mMagName = ""
       If Cells(i, 4).Value <> Empty Then
            mMagName = Left(Trim(Cells(i, 2).Value), 30)
            checkSheetName = ""
            On Error Resume Next
            checkSheetName = Worksheets(mMagName).Name
            If checkSheetName = "" Then
               mList.Copy after:=ActiveSheet
               ActiveSheet.Name = mMagName
                With ActiveSheet.Tab
                    .ColorIndex = xlColorIndexNone
                    .TintAndShade = 0
                End With
            End If
        End If
    
    Sheets(mSprMagListName).Select
    Next

'4.3

  Worksheets(Trim(mItogListName)).Select
  For x = mStartRow To mEndRow
             For y = mStartCol To mEndCol
                Cells(x, y).Value = Empty
                Cells(x, y).Formula = Empty
             Next
  Next


'4.4

  For i = mC To 1 Step -1
    If Cells(3, i) <> "" Then
      mNameTov = Trim(Cells(3, i))
      Exit For
    End If
  Next

'4.5
For Each cen In aTovCens
    If Round(cen, 2) = Round(mCena, 2) Then
       ex = True
       Exit For
    End If
Next                