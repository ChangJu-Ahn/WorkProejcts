Const SUM_SEQ_NO = 999999

'==========================================================================================
'   Event Desc : Grid의 Max Count 를 찾는다.
'==========================================================================================
Function GoToCell(Byref oSpread, Byval Col, Byval Row)
	With oSpread
		.SetActiveCell Col, Row
		.Col = Col : .Row = Row
		
		.SelStart = 1
		.SelLength = Len(.Value)
	End With
End Function

'==========================================================================================
'   Event Desc : Grid의 Max Count 를 찾는다.
'==========================================================================================
Function MaxSpreadVal(Byref objSpread, ByVal intCol, byval Row)

	Dim iRows
	Dim MaxValue
	Dim tmpVal

	MAxValue = 0

	For iRows = 1 to  objSpread.MaxRows
		objSpread.row = iRows
	    objSpread.col = intCol

		If objSpread.Text = "" Then
		   tmpVal = 0
		Else
  		   tmpVal = cdbl(objSpread.value)
		End If

		If tmpval > MaxValue And tmpval < SUM_SEQ_NO Then
		   MaxValue = cdbl(tmpVal)
		End If
	Next

	MaxValue = MaxValue + 1
	objSpread.row	= row
	objSpread.col	= intCol
	objSpread.text	= MaxValue
	MaxSpreadVal = MaxValue
end Function

'==========================================================================================
'   Event Desc : Grid의 Max Count 를 찾는다.
'==========================================================================================
Function MaxSpreadVal3(Byref objSpread, ByVal intCol, byval Row)

	Dim iRows
	Dim MaxValue
	Dim tmpVal

	MAxValue = 0

	For iRows = 1 to  objSpread.MaxRows
		objSpread.row = iRows
	    objSpread.col = intCol

		If objSpread.Text = "" Then
		   tmpVal = 0
		Else
  		   tmpVal = cdbl(objSpread.value)
		End If

		If tmpval > MaxValue And tmpval < SUM_SEQ_NO Then
		   MaxValue = cdbl(tmpVal)
		End If
	Next

	MaxValue = MaxValue + 1
	MaxValue = Right("000000" & MaxValue, 6)
	objSpread.row	= row
	objSpread.col	= intCol
	objSpread.text	= MaxValue
	MaxSpreadVal = MaxValue
end Function

'==========================================================================================
'   Event Desc : Grid의 Max Count 를 찾는다.
'==========================================================================================
Function GetMaxSpreadVal(Byref objSpread, ByVal intCol)

	Dim iRows
	Dim MaxValue
	Dim tmpVal

	MAxValue = 0

	For iRows = 1 to  objSpread.MaxRows
		objSpread.row = iRows
	    objSpread.col = intCol

		If objSpread.Text = "" Then
		   tmpVal = 0
		Else
  		   tmpVal = cdbl(objSpread.value)
		End If

		If tmpval > MaxValue And tmpval <> SUM_SEQ_NO Then
		   MaxValue = cdbl(tmpVal)
		End If
	Next

	MaxValue = MaxValue + 1

	GetMaxSpreadVal = MaxValue
end Function

'==========================================================================================
'   Event Desc : 히든 Grid의 Max Count 를 찾는다.
'==========================================================================================
Function MaxSpreadVal2(Byref objSpread, ByVal intHeadCol, ByVal intDetailCol, byval Row, Byval pSeqNo)

	Dim iRows
	Dim MaxValue
	Dim tmpVal, iSeqNo

	MAxValue = 0

	For iRows = 1 to  objSpread.MaxRows
		objSpread.row = iRows
	    objSpread.col = intHeadCol
		iSeqNo = objSpread.Value
		
		If iSeqNo = pSeqNo Then
			objSpread.col = intDetailCol
			If objSpread.Text = "" Then
			   tmpVal = 0
			Else
  			   tmpVal = cdbl(objSpread.value)
			End If

			If tmpval > MaxValue And tmpval <> SUM_SEQ_NO Then
			   MaxValue = cdbl(tmpVal)
			End If
		End If
	Next

	MaxValue = MaxValue + 1

	objSpread.row	= row
	objSpread.col	= intDetailCol
	objSpread.text	= MaxValue
	objSpread.col	= intHeadCol
	objSpread.text	= pSeqNo	

end Function

'==========================================================================================
'   Event Desc : Sheet의 금액을 Sum한다.
'==========================================================================================
Function FncSumSheet1(pObject,pPiVot,pStart,pEnd,pBool,pTargetRow,pTargetCol,pVerHor)

    Dim iDx
    Dim iSum
    Dim iOperStatus

    iOperStatus =  True

    iSum = 0

    For iDx = pStart to pEnd

        pObject.Row = iDx
        pObject.Col = 0

        IF pObject.Text <> ggoSpread.DeleteFlag Then

			If pVerHor = "V" Then
			   pObject.Col = pPiVot
			Else
			   pObject.Row = pPiVot
			End If

			If pVerHor = "V" Then
			   pObject.Row = iDx
			Else
			   pObject.Col = iDx
			End If
			
			If Trim(pObject.Text) > ""  Then
			   If IsNumeric(parent.UNICDbl(pObject.Text)) Then
			      iSum = iSum + parent.UNICDbl(pObject.Text)
			   Else
			      iOperStatus = False
			   End If
			End If

        End If
    Next

   If iOperStatus = True Then
       If pBool =  True Then
          pObject.Col  = pTargetCol
          pObject.Row  = pTargetRow
          pObject.Text = iSum
       End If
    End If

   FncSumSheet1  = iSum

End Function

' 법인세는 그리드에 합계가 존재하므로, 삭제된 Row를 제외하고 썸을 내야한다.
'======================================================================================================
Function FncSumSheet(pObject,pPiVot,pStart,pEnd,pBool,pTargetRow,pTargetCol,pVerHor)
    Dim iDx
    Dim iSum, iSumTemp
    Dim iOperStatus, RowStatus
    
    iOperStatus =  True	: RowStatus = ""
    
    With pObject
    
		If pVerHor = "V" Then
		   .Col = pPiVot
		Else
		   .Row = pPiVot
		End If       
    
		iSum = 0
		For iDx = pStart To pEnd
		    If pVerHor = "V" Then
		       .Row = iDx 
		       
		       .Col = 0 : RowStatus = .Value : .Col = pPiVot ' 원상복귀 
		    Else
		       .Col = iDx 
		    End If
		               
		    If Trim(.Text) > ""  Then
		       If IsNumeric(.Value) Then
				  If RowStatus <> "삭제" And .RowHIdden = False Then
					iSum = iSum + UNICDbl(.Value) 
		          End If
		       Else
		          iOperStatus = False
		       End If   
		    End If   
		    
		Next

		iSumTemp = Replace(iSum    , gClientNum1000, "**"         )
		iSumTemp = Replace(iSumTemp, gClientNumDec ,  "@@"        )
		iSumTemp = Replace(iSumTemp, "@@"          ,   gComNumDec )
		iSum     = Replace(iSumTemp, "**"          ,   gComNum1000)
		    
		If iOperStatus = True Then
		   If pBool =  True Then
		      .Col  = pTargetCol
		      .Row  = pTargetRow
		      
		      .Value = iSum
		      
		   End If   
		End If   
	End With
	
    FncSumSheet  = iSum
    
End Function 

' -- 배열을 콤보스트링으로 
Function MakeSpreadCombo(pArr)
	Dim i, iLen, sTmp
	iLen = UBound(pArr)
	For i = 0 To iLen -1
		If pArr(i) <> "" Then	sTmp = sTmp & pArr(i) & Chr(9)
	Next
	MakeSpreadCombo = sTmp
End Function

' -- 그리드 컬럼을 퍼센트형으로 
Function MakePercentCol(Byref pObj, Byval pCol, Byval pDecimal, Byval pMax, Byval pMin) 
	With pObj
		If pMax = "" Then pMax = 100
		If pMin = "" Then pMin = 0
		If pDecimal = "" Then pDecimal = 2
		' 퍼센트 형 정의 
		.Col = pCol
		.Row = -1
		.CellType = 14
		'.TypePercentDecimal = pDecimal
		.TypePercentMax = pMax
		.TypePercentMin = pMin
		'.TypePercentDecPlaces = 0
		.TypeHAlign = 1
	End With
End Function

' -- 그리드 컬럼을 퍼센트형으로 
Function MakePercentType(Byref pObj, Byval pCol1, Byval pRow1, Byval pCol2, pRow2, Byval pDecimal, Byval pMax, Byval pMin) 
	With pObj
		If pMax = "" Then pMax = 100
		If pMin = "" Then pMin = 0
		If pDecimal = "" Then pDecimal = 2
		' 퍼센트 형 정의 
		
		.Col = pCol1	
		If pCol2 > -1 Then .Col2 = pCol2
		
		.Row = pRow1	
		If pRow2 > -1 Then .Row2 = pRow2
		
		.BlockMode = True
		.CellType = 14
		.TypePercentDecimal = pDecimal
		.TypePercentMax = pMax
		.TypePercentMin = pMin
		.TypePercentDecPlaces = 0
		.TypeHAlign = 1
		.TypeVAlign = 2
		.BlockMode = False
	End With
End Function

