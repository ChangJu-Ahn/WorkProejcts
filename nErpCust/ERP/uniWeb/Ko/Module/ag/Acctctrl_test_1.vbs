
'=======================================================================================================
'	관리항목 그리드 상수선언 
'=======================================================================================================
Const C_HMaxCols = 16
Const C_NoteSep  = ","							'비고 seperate.

Dim C_DtlSeq
Dim C_CtrlCd
Dim C_CtrlNm
Dim C_CtrlVal
Dim C_CtrlPB
Dim C_CtrlValNm
Dim C_Seq
Dim C_Tableid
Dim C_Colid
Dim C_ColNm
Dim C_Datatype
Dim C_DataLen
Dim C_DRFg
Dim C_HItemSeq
Dim C_MajorCd

Dim lgCashAcct									'현금계정을 미리 저장한다.
Dim lgAuthorityFlag                             '권한관리 추가 
Dim lgOrgExitFg					' 부서조회 위해 추가 
Dim lgOrgChangeId 

'========================================================================================================
' Name : initSpreadPosVariables()
' Desc : Initialize Column Const value
'========================================================================================================
Sub initCtrlSpreadPosVariables()

	C_DtlSeq    = 1
	C_CtrlCd    = 2
	C_CtrlNm    = 3
	C_CtrlVal   = 4
	C_CtrlPB    = 5
	C_CtrlValNm = 6
	C_Seq       = 7
	C_Tableid   = 8
	C_Colid     = 9
	C_ColNm     = 10
	C_Datatype  = 11
	C_DataLen   = 12
	C_DRFg      = 13
	C_HItemSeq  = 14
	C_MajorCd   = 15

End Sub
'=======================================================================================================
'   Event Name : InitCtrlSpread()
'   Event Desc : 관리항목 그리드 초기화 
'=======================================================================================================
Sub InitCtrlSpread()

	Call initCtrlSpreadPosVariables()
	
    With frm1

		ggoSpread.Source			= .vspdData2
		ggoSpread.Spreadinit "V20021217",,parent.gAllowDragDropSpread

		.vspdData2.ReDraw			= False
		
'		.vspdData2.AutoClipboard	= False
		.vspdData2.MaxCols			= C_MajorCd + 1

		Call ggoSpread.ClearSpreadData()

		Call AppendNumberPlace("6","3","0")
		Call GetCtrlSpreadColumnPos("A")

		ggoSpread.SSSetFloat	C_DtlSeq,		"NO" ,				6,	"6",	ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gComNum1000,	parent.gComNumDec,	2,	,	,	"0",	"999"
		ggoSpread.SSSetEdit		C_CtrlCd,		"관리항목",			10,	2
		ggoSpread.SSSetEdit		C_CtrlNm,		"관리항목명",		30,	3
		ggoSpread.SSSetEdit		C_CtrlVal,		"관리항목 VALUE",	32,	,		,							30,							2
		ggoSpread.SSSetButton	C_CtrlPB
		ggoSpread.SSSetEdit		C_CtrlValNm,	"관리항목 VALUE명",	45
		ggoSpread.SSSetEdit		C_Seq,			"A",				8,	,		,							3
		ggoSpread.SSSetEdit		C_Tableid,		"B",				32
		ggoSpread.SSSetEdit		C_Colid,		"C",				32
		ggoSpread.SSSetEdit		C_ColNm,		"D",				32
		ggoSpread.SSSetEdit		C_DataType,		"E",				2
		ggoSpread.SSSetFloat	C_DataLen,		"F",				3,	"6",	ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gComNum1000,	parent.gComNumDec,	,	,	,	"0",	"999"
		ggoSpread.SSSetEdit		C_DRFg,			"G",				1
		ggoSpread.SSSetEdit		C_HItemSeq,     "H",				1
		ggoSpread.SSSetEdit		C_MajorCd,		"I",				1

		Call ggoSpread.MakePairsColumn(C_CtrlVal,C_CtrlPB)

		Call ggoSpread.SSSetColHidden(C_DtlSeq,C_DtlSeq,True)
		Call ggoSpread.SSSetColHidden(C_CtrlCd,C_CtrlCd,True)
		Call ggoSpread.SSSetColHidden(C_Seq,C_Seq,True)
		Call ggoSpread.SSSetColHidden(C_Tableid,C_Tableid,True)
		Call ggoSpread.SSSetColHidden(C_Colid,C_Colid,True)
		Call ggoSpread.SSSetColHidden(C_ColNm,C_ColNm,True)
		Call ggoSpread.SSSetColHidden(C_DataType,C_DataType,True)
		Call ggoSpread.SSSetColHidden(C_DataLen,C_DataLen,True)
		Call ggoSpread.SSSetColHidden(C_DRFg,C_DRFg,True)
		Call ggoSpread.SSSetColHidden(C_HItemSeq,C_HItemSeq,True)
		Call ggoSpread.SSSetColHidden(C_MajorCd,C_MajorCd,True)
		Call ggoSpread.SSSetColHidden(.vspdData2.MaxCols,.vspdData2.MaxCols,True)


		.vspdData2.ReDraw = True

    End With
    
	Call CtrlSpreadLock("X","X", -1, -1)

End Sub

'=======================================================================================================
'   Event Name : InitCtrlHSpread()
'   Event Desc : 관리항목 그리드 초기화 
'=======================================================================================================
Sub InitCtrlHSpread()
	
	frm1.vspdData3.ReDraw = False
	ggoSpread.Source = frm1.vspdData3
	Call ggoSpread.ClearSpreadData()
	frm1.vspdData3.MaxCols = C_HMaxCols
	frm1.vspdData3.ReDraw = True

End Sub

'=======================================================================================================
' Function Name : CtrlSpreadLock
' Function Desc : 관리항목 그리드 Lock
'=======================================================================================================
Sub CtrlSpreadLock(Byval stsFg, Byval Index, ByVal lRow  , ByVal lRow2 )

	Dim objSpread

	With frm1

		ggoSpread.Source	= .vspdData2

		.vspdData2.Redraw	= False

		ggoSpread.SpreadLock  C_DtlSeq,    lRow,  C_DtlSeq,    lRow2
		ggoSpread.SpreadLock  C_CtrlCd,    lRow,  C_CtrlCd,    lRow2
		ggoSpread.SpreadLock  C_CtrlNm,    lRow,  C_CtrlNm,    lRow2
		ggoSpread.SpreadLock  C_CtrlValNm, lRow,  C_CtrlValNm, lRow2
		    		
		.vspdData2.Redraw = True

	End With

End Sub

'=======================================================================================================
'   Event Name : SetSpread2Color()
'   Event Desc : 관리항목 그리드 색상설정, Protect, Require
'=======================================================================================================
Sub SetSpread2Color()

Dim indx
Dim tmpDrCrFG
Dim strStartRow, strEndRow

    With frm1
		strStartRow = 1
		strEndRow	= .vspdData2.MaxRows

		ggoSpread.Source	= .vspdData2
		.vspdData2.ReDraw	= False		
		.vspdData.Col		= C_DrCRFG
		If lgCurrRow = "" then
			lgCurrRow = 1
		End If
		If (lgCurrRow >= 1 And .vspdData.Row > lgCurrRow) Or .vspdData.Row < 1 Then
			.vspdData.Row		= lgCurrRow
		End If
		tmpDrCrFG			= LEFT(.vspdData.Text,1)

		ggoSpread.SSSetProtected C_DtlSeq,		strStartRow,	strEndRow
		ggoSpread.SSSetProtected C_CtrlCd,		strStartRow,	strEndRow
		ggoSpread.SSSetProtected C_CtrlNm,		strStartRow,	strEndRow
		ggoSpread.SSSetProtected C_CtrlValNm,	strStartRow,	strEndRow

		For indx = 1 to .vspdData2.MaxRows

			.vspdData2.Row = indx
			.vspdData2.Col = C_DrFg

			'msgbox "lgCurrRow=" & lgCurrRow & "tmpDrCrFG=" & tmpDrCrFG & ".vspdData2.Text=" & .vspdData2.Text
			If (.vspdData2.Text = tmpDrCrFG AND .vspdData2.Text <>"" ) Or .vspdData2.Text = "Y" Or .vspdData2.Text = "DC" Then
				ggoSpread.SSSetRequired C_CtrlVal, indx, indx
			Else
				ggoSpread.SpreadUnLock  C_CtrlVal, indx, C_CtrlVal, indx
			End If
			
		Next
		
		.vspdData2.ReDraw = True

    End With

End Sub
'========================================================================================
' Function Name : GetCtrlSpreadColumnPos
' Description   : 
'========================================================================================
Sub GetCtrlSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData2
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
            C_DtlSeq    = iCurColumnPos(1)
			C_CtrlCd    = iCurColumnPos(2)
			C_CtrlNm    = iCurColumnPos(3)
			C_CtrlVal   = iCurColumnPos(4)
			C_CtrlPB    = iCurColumnPos(5)
			C_CtrlValNm = iCurColumnPos(6)
			C_Seq       = iCurColumnPos(7)
			C_Tableid   = iCurColumnPos(8)
			C_Colid     = iCurColumnPos(9)
			C_ColNm     = iCurColumnPos(10)
			C_Datatype  = iCurColumnPos(11)
			C_DataLen   = iCurColumnPos(12)
			C_DRFg      = iCurColumnPos(13)
			C_HItemSeq  = iCurColumnPos(14)
			C_MajorCd   = iCurColumnPos(15)

    End Select    
End Sub
'=======================================================================================================
'   Event Name : OpenCtrlPB
'   Event Desc : 관리항목 PopUp
'=======================================================================================================
Function OpenCtrlPB(Byval strTable, Byval strFld1 , Byval strFld2 , Byval strCode , Byval FldNm, ByVal sWhere )

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	If strFld1 = "BANK_ACCT_NO" Then
		arrParam(0) = "관리항목 VALUE 팝업"		' 팝업 명칭 
		arrParam(1) = strTable	    				' TABLE 명칭 
		arrParam(2) = strCode						' Code Condition
		arrParam(3) = ""							' Name Cindition
		arrParam(4) = sWhere						' Where Condition
		arrParam(5) = FldNm      					' 조건필드의 라벨 명칭 

		arrField(0) = strFld1	    				' Field명(0)
		arrField(1) = strFld2	    		        ' Field명(1)
		arrField(2) = "BANK_CD"	    		        ' Field명(2)

		arrHeader(0) = "관리항목 VALUE"			' Header명(0)
		arrHeader(1) = "관리항목 VALUE명"
		arrHeader(2) = "은행코드"
		
	ElseIf strFld1 = "COST_CD" Then	
	  	arrParam(0) = "관리항목 VALUE 팝업"						' 팝업 명칭 
	  	arrParam(1) = "B_ACCT_DEPT A, B_COST_CENTER B "	    	' TABLE 명칭 
	  	arrParam(2) = strCode						' Code Condition
	  	arrParam(3) = ""							' Name Cindition
	  	'arrParam(4) = sWhere						' Where Condition
		arrParam(4) = "A.ORG_CHANGE_ID =  " & FilterVar(frm1.hOrgChangeId.value , "''", "S") & ""
		'arrParam(4) = arrParam(4) & " And A.COST_CD = B.COST_CD And B.BIZ_UNIT_CD = ( Select B.BIZ_UNIT_CD"
		'arrParam(4) = arrParam(4) & " And A.DEPT_CD = B.DEPT_CD And B.BIZ_UNIT_CD = ( Select B.BIZ_UNIT_CD"
		arrParam(4) = arrParam(4) & " And A.COST_CD =* B.COST_CD And B.BIZ_UNIT_CD = ( Select B.BIZ_UNIT_CD"
		arrParam(4) = arrParam(4) & " FROM B_ACCT_DEPT A, B_COST_CENTER B WHERE A.DEPT_CD =  " & FilterVar(frm1.txtDeptCd.value , "''", "S") & ""
		arrParam(4) = arrParam(4) & " And A.DEPT_CD = B.DEPT_CD And A.ORG_CHANGE_ID =  " & FilterVar(frm1.hOrgChangeId.value , "''", "S") & ")"	  	
	  	arrParam(5) = "코스트코드"   				' 조건필드의 라벨 명칭 
	  
	  	'arrField(0) = "A.COST_CD"    				' Field명(0)
	  	arrField(0) = "B.COST_CD"    				' Field명(0)
	  	arrField(1) = "B.COST_NM"	    		    ' Field명(1)
	  
	  	arrHeader(0) = "관리항목 VALUE"				' Header명(0)
	  	arrHeader(1) = "관리항목 VALUE명"		
		
	Else
	  	arrParam(0) = "관리항목 VALUE 팝업"			' 팝업 명칭 
	  	arrParam(1) = strTable	    				' TABLE 명칭 
	  	arrParam(2) = strCode						' Code Condition
	  	arrParam(3) = ""							' Name Cindition
	  	arrParam(4) = sWhere						' Where Condition
	  	arrParam(5) = FldNm      					' 조건필드의 라벨 명칭 
	  
	  	arrField(0) = strFld1	    				' Field명(0)
	  	arrField(1) = strFld2	    		        ' Field명(1)
	  
	  	arrHeader(0) = "관리항목 VALUE"				' Header명(0)
	  	arrHeader(1) = "관리항목 VALUE명"
	End If


	'####### 관리항목이 거래처일때, 거래처정보가 상세히 보이도록 수정 #######
	'200803181431
	If strFld1 <> "BP_CD" Then
		arrRet = window.showModalDialog("../../comasp/adoAcctctrl_ko441_1_popup.asp", Array(arrParam, arrField, arrHeader), _
				"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	Else
		'관리항목이 거래처일때.
		Dim iCalledAspName
		Dim IntRetCD

		
		iCalledAspName = AskPRAspName("b1261pa1")	
		If Trim(iCalledAspName) = "" then
			IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION, "b1261pa1", "x")
			IsOpenPop = False
			Exit Function
		End if
	
		arrRet = window.showModalDialog(iCalledAspName,Array(window.parent, strCode),"dialogWidth=800px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	End If
	'200803181431


	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetCtrlPB(arrRet, strFld2, strFld1)
	End If

End Function

'=======================================================================================================
'   Event Name : SetCtrlPB
'   Event Desc : 관리항목 PopUp Data Setting
'=======================================================================================================
Function SetCtrlPB(Byval arrRet, Byval pstrFld2, Byval strFld1)
	Dim lngRows
	Dim iTempColid
	Dim iTempCtrlVal
	Dim itempColNm
	Dim strSelect
	Dim strFrom
	Dim strWhere 	
	Dim arrVal1

	With frm1
		If strFld1 = "BANK_ACCT_NO" Then
			.vspdData2.Row =  .vspdData2.ActiveRow
			.vspdData2.Col =  C_CtrlVal
			.vspdData2.Text = arrRet(0)

			If Len(Trim(pstrFld2)) > 0 Then
				.vspdData2.Col =  C_CtrlValNm
				.vspdData2.Text = arrRet(1)
			End If
			call vspdData2_Change( C_CtrlVal ,  .vspdData2.Row)
			For lngRows = 1 To .vspdData2.MaxRows
				.vspdData2.Row = lngRows
				.vspdData2.Col = C_CtrlVal
 				iTempCtrlVal = Trim(.vspdData2.Text)
				.vspdData2.Col = C_Colid
 				iTempColid = Trim(.vspdData2.Text)
		
				IF iTempColid = "BANK_CD" and iTempCtrlVal = "" Then
					.vspdData2.Col = C_CtrlVal
					.vspdData2.Text = arrRet(2)
					' query bank_nm 				
					.vspdData2.Col = C_ColNm  
					itempColNm = Trim(.vspdData2.Text) 

					strSelect	=	  itempColNm     		
					strFrom		=	  " B_BANK "		
					strWhere	=	  " BANK_CD = " & FilterVar(arrRet(2), "''", "S") & ""

					If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) Then 	
						arrVal1 = Split(lgF2By2, Chr(11))			
						.vspdData2.Col = C_CtrlValNm  '은행명 
						.vspdData2.Text = Trim(arrVal1(1))
					End if 
					Call CopyToHSheet2(frm1.vspdData.ActiveRow,.vspdData2.Row)
					Exit For
				END IF
				
			Next
		Else	
  		.vspdData2.Row =  .vspdData2.ActiveRow
  		.vspdData2.Col =  C_CtrlVal
  		.vspdData2.Text = arrRet(0)
  
			If Len(Trim(pstrFld2)) > 0 Then
				.vspdData2.Col =  C_CtrlValNm
				.vspdData2.Text = arrRet(1)
			End If
		Call vspdData2_Change( C_CtrlVal ,  .vspdData2.Row)
		End If

	End With

End Function




'======================================================================================================
' Function Name : FindCtrlNM
' Function Desc : 관리항목값 명을 찿아 setting한다.
'=======================================================================================================
Function FindCtrlNM(ByVal Row)

    Dim iFld1
	Dim iFld2
	Dim iTable
	Dim istrCode
	Dim sWhere
	Dim IntRetCD

	If Row < 1 Then Exit Function

	'---------- Coding part -------------------------------------------------------------
	ggoSpread.Source = frm1.vspdData2

	With frm1.vspdData2

		.Row		= Row
		.Col		= C_CtrlVal
		istrCode	= Trim(.Text)

		.Col		= C_Tableid
		iTable		= .Text

		If iTable <> "" AND istrCode <> "" Then 
			.Col	= C_Colid
			iFld1	= .Text

			.Col	= C_ColNm
			iFld2	= .Text

			sWhere	= iFld1 & " =  " & FilterVar(istrCode , "''", "S") & ""

			.Col	= C_MajorCD

			If  .Text <> "" Then
				
				sWhere = sWhere & " and  Major_CD =  " & FilterVar(.Text , "''", "S") & ""
				
			End If

			frm1.vspdData2.Col = C_CtrlValNm

			If CommonQueryRs(iFld2, iTable, sWhere, lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) Then
    			arrVal = Split(lgF0, Chr(11))
				frm1.vspdData2.Text = arrVal(0)
			ELSE
				IntRetCD = DisplayMsgBox("110330", "X", "X", "X")									'필수입력 check!!
				' 관리항목값이 바르지 않습니다.
				frm1.vspdData2.Text = ""
				frm1.vspdData2.Col = C_CtrlVal
				frm1.vspdData2.Text = ""
				Exit Function
			END IF
		End if

	End With

End Function

'=======================================================================================================
'   Function Name : DeleteHSheet
'   Function Desc : 입력받은 Item번호와 관계된 관리항목 Hidden 그리드 데이타 삭제 
'=======================================================================================================
Function DeleteHSheet(ByVal strItemSeq)

	Dim boolExist
	Dim lngRows, lngRow2, lngRow3, lngCol3
	Dim StrData
	Dim strCtrlItemSeq

	DeleteHSheet = False
	boolExist = False

	With frm1

		Call SortHSheet()

		'------------------------------------
		' Find First Row
		'------------------------------------
        For lngRows = 1 To .vspdData3.MaxRows
			.vspdData3.Row = lngRows
			.vspdData3.Col = 1

			If strItemSeq = .vspdData3.Text Then
				boolExist = True
				Exit For
			End If
		Next

		lngRow2 = 1
		'------------------------------------
        ' Data Delete
        '------------------------------------
        If boolExist = True Then
			While lngRows <= .vspdData3.MaxRows

				.vspdData3.Row = lngRows
				.vspdData3.Col = 1

				If strItemSeq <> .vspdData3.Text Then
					lngRows = .vspdData3.MaxRows + 1
				Else
					If frm1.vspdData2.MaxRows > 0 Then
						.vspdData2.Col = C_HItemSeq
						.vspdData2.Row = .vspdData2.MaxRows
						strCtrlItemSeq = .vspdData2.Text
						'msgbox "strItemSeq" & strItemSeq & " :: " & "strCtrlItemSeq=" & strCtrlItemSeq & " :: " & ".vspdData2.Row=" & .vspdData2.Row
						If strCtrlItemSeq = strItemSeq Then
							.vspdData2.Action = 5
							.vspdData2.MaxRows = .vspdData2.MaxRows - 1
						Else
							lngRow2 =  lngRow2 + 1
						End If
					End If
					
					.vspdData3.Action = 5
					.vspdData3.MaxRows = .vspdData3.MaxRows - 1
				End If

			Wend

        End If

	End With

	DeleteHSheet = True

End Function


'======================================================================================================
' Function Name : SortHSheet
' Function Desc : 관리항목 Hidden Grid 정렬 
'=======================================================================================================
Function SortHSheet()

    With frm1
    
        .vspdData3.BlockMode	= True
        .vspdData3.Col			= 0
        .vspdData3.Col2			= .vspdData3.MaxCols
        .vspdData3.Row			= 1
        .vspdData3.Row2			= .vspdData3.MaxRows
        .vspdData3.SortBy		= 0											'SS_SORT_BY_ROW

        .vspdData3.SortKey(1)	= 1
        .vspdData3.SortKey(2)	= 2

        .vspdData3.SortKeyOrder(1) = 1										'SS_SORT_ORDER_ASCENDING
        .vspdData3.SortKeyOrder(2) = 1										'SS_SORT_ORDER_ASCENDING

        .vspdData3.Col			= 0
        .vspdData3.Col2			= .vspdData3.MaxCols
        .vspdData3.Row			= 0
        .vspdData3.Row2			= .vspdData3.MaxRows
        .vspdData3.Action		= 25										'SS_ACTION_SORT
        .vspdData3.BlockMode	= False
        
    End With

End Function

'=======================================================================================================
' Function Name : ShowHidden
' Function Desc : 관리항목 Hidden Grid를 표시 
'=======================================================================================================
Sub ShowHidden()

	Dim strHidden
	Dim lngRows
	Dim lngCols

    With frm1.vspdData3

        For lngRows = 1 To .MaxRows
            .Row = lngRows
            For lngCols = 1 To .MaxCols
	            .Col = lngCols
			    strHidden = strHidden & " | " & .Text
			Next
            strHidden = strHidden & Chr(12) & vbcrlf
        Next

    End With

End Sub

'=======================================================================================================
'   Function Name : CopyToHSheet
'   Function Desc : 관리항목그리드의 Value변경시 Hidden Grid에 Data 반영, Item 그리드에 변경여부 표시 
'=======================================================================================================
Sub CopyToHSheet(ByVal Row)

	Dim lRow
	Dim iCols

	With frm1

	    lRow = FindData

	    If lRow > 0 Then
	    
            .vspdData3.Row = lRow
            .vspdData2.Row = Row
            .vspdData3.Col = 0
            .vspdData2.Col = 0
            .vspdData3.Text = .vspdData2.Text
            
            .vspdData2.Col = C_DtlSeq
            .vspdData3.Col = 2
            .vspdData3.Text = .vspdData2.Text
                
            .vspdData2.Col = C_CtrlCd
            .vspdData3.Col = .vspdData3.Col + 1
            .vspdData3.Text = .vspdData2.Text
                
            .vspdData2.Col = C_CtrlNm
            .vspdData3.Col = .vspdData3.Col + 1
            .vspdData3.Text = .vspdData2.Text
                
            .vspdData2.Col = C_CtrlVal
            .vspdData3.Col = .vspdData3.Col + 1
            .vspdData3.Text = .vspdData2.Text
                
            .vspdData2.Col = C_CtrlPB
            .vspdData3.Col = .vspdData3.Col + 1
            .vspdData3.Text = .vspdData2.Text
                
            .vspdData2.Col = C_CtrlValNm
            .vspdData3.Col = .vspdData3.Col + 1
            .vspdData3.Text = .vspdData2.Text
                
            .vspdData2.Col = C_Seq
            .vspdData3.Col = .vspdData3.Col + 1
            .vspdData3.Text = .vspdData2.Text
                
            .vspdData2.Col = C_Tableid
            .vspdData3.Col = .vspdData3.Col + 1
            .vspdData3.Text = .vspdData2.Text
                
            .vspdData2.Col = C_Colid
            .vspdData3.Col = .vspdData3.Col + 1
            .vspdData3.Text = .vspdData2.Text
                
            .vspdData2.Col = C_ColNm
            .vspdData3.Col = .vspdData3.Col + 1
            .vspdData3.Text = .vspdData2.Text
                
            .vspdData2.Col = C_Datatype
            .vspdData3.Col = .vspdData3.Col + 1
            .vspdData3.Text = .vspdData2.Text
                
            .vspdData2.Col = C_DataLen
            .vspdData3.Col = .vspdData3.Col + 1
            .vspdData3.Text = .vspdData2.Text
                
            .vspdData2.Col = C_DRFg
            .vspdData3.Col = .vspdData3.Col + 1
            .vspdData3.Text = .vspdData2.Text

            .vspdData2.Col = C_MajorCd
            .vspdData3.Col = .vspdData3.Col + 1
            .vspdData3.Text = .vspdData2.Text
            
            .vspdData2.Col = C_MajorCd + 1 
            .vspdData3.Col = .vspdData3.Col + 1
            .vspdData3.Text = .vspdData2.Text
            
        End If

	End With

	frm1.vspdData.Row = frm1.vspdData.ActiveRow
	frm1.vspdData.Col = 0

	If frm1.vspdData.Text <> ggoSpread.InsertFlag And frm1.vspdData.Text <> ggoSpread.DeleteFlag Then
   	    frm1.vspdData.Text = ggoSpread.UpdateFlag
	End if

End Sub

'=======================================================================================================
'   Function Name : FindData
'   Function Desc : 현재의 Item, Dtl에 해당하는 Hidden Grid의 Index를 Return
'=======================================================================================================
Function FindData()

	Dim strApNo
	Dim strItemSeq
	Dim strDtlSeq
	Dim lRows

    FindData = 0

    With frm1

        For lRows = 1 To .vspdData3.MaxRows

			.vspdData3.Row	= lRows
			.vspdData3.Col	= 1
            strItemSeq		= .vspdData3.Text
            .vspdData3.Col	= 2
            strDtlSeq		= .vspdData3.Text

            .vspdData.Row	= frm1.vspdData.ActiveRow
            .vspdData2.Row	= frm1.vspdData2.ActiveRow

            .vspdData.Col	= C_ItemSeq
            
            If strItemSeq = .vspdData.Text Then
                .vspdData2.Col = C_DtlSeq
                If strDtlSeq = .vspdData2.Text Then
                    FindData = lRows
                    Exit Function
                End If
            End If

        Next

    End With

End Function


'=======================================================================================================
'   Function Name : CopyFromData
'   Function Desc : 관리항목 Hidden 그리드에서 입력받은 Item번호에 
'                   해당하는 관리항목 값을 표시, 해당 관리항목이 없으면 False 값 Return
'=======================================================================================================
Function CopyFromData(ByVal strItemSeq)

	Dim lngRows , indx, indx1
	Dim boolExist
	Dim iCols
	Dim tmpDrCrFG
	Dim iStrData, iStrFlag
	Dim arrFlag
	Dim strHItemSeq

    boolExist = False

	ggoSpread.Source = frm1.vspdData2
	Call ggoSpread.ClearSpreadData()
	
    CopyFromData			= boolExist

    With frm1

        Call SortHSheet()
      '------------------------------------
      ' Find First Row
      '------------------------------------
        For lngRows = 1 To .vspdData3.MaxRows
            .vspdData3.Row = lngRows
            .vspdData3.Col = 1
            If strItemSeq = .vspdData3.Text Then
                boolExist = True
                Exit For
            End If
        Next

      '------------------------------------
      ' Show Data
      '------------------------------------
		.vspdData3.Row = lngRows
		
        If boolExist = True Then

			ggoSpread.Source = .vspdData2
			Call ggoSpread.ClearSpreadData()
            .vspdData2.Redraw = False

            For indx = lngRows to .vspdData3.MaxRows

                .vspdData3.Row = indx
                .vspdData3.Col = 1

                If strItemSeq = .vspdData3.Text Then
					
                    .vspdData3.Col= 0
                    iStrFlag = iStrFlag & .vspdData3.Text & Chr(12)
                    
                    .vspdData3.Col = 1
                    strHItemSeq = .vspdData3.Text
                    
                    .vspdData3.Col = 2
                    iStrData = iStrData & Chr(11) & .vspdData3.Text
                    .vspdData3.Col = .vspdData3.Col + 1
                    iStrData = iStrData & Chr(11) & .vspdData3.Text 
                    .vspdData3.Col = .vspdData3.Col + 1
                    iStrData = iStrData & Chr(11) & .vspdData3.Text
                    .vspdData3.Col = .vspdData3.Col + 1
                    iStrData = iStrData & Chr(11) & .vspdData3.Text
                    .vspdData3.Col = .vspdData3.Col + 1
                    iStrData = iStrData & Chr(11) & .vspdData3.Text
                    .vspdData3.Col = .vspdData3.Col + 1
                    iStrData = iStrData & Chr(11) & .vspdData3.Text
                    .vspdData3.Col = .vspdData3.Col + 1
                    iStrData = iStrData & Chr(11) & .vspdData3.Text
                    .vspdData3.Col = .vspdData3.Col + 1
                    iStrData = iStrData & Chr(11) & .vspdData3.Text
                    .vspdData3.Col = .vspdData3.Col + 1
                    iStrData = iStrData & Chr(11) & .vspdData3.Text
                    .vspdData3.Col = .vspdData3.Col + 1
                    iStrData = iStrData & Chr(11) & .vspdData3.Text
                    .vspdData3.Col = .vspdData3.Col + 1
                    iStrData = iStrData & Chr(11) & .vspdData3.Text
                    .vspdData3.Col = .vspdData3.Col + 1
                    iStrData = iStrData & Chr(11) & .vspdData3.Text
                    .vspdData3.Col = .vspdData3.Col + 1
                    iStrData = iStrData & Chr(11) & .vspdData3.Text
                    iStrData = iStrData & Chr(11) & strHItemSeq
                    .vspdData3.Col = .vspdData3.Col + 1
                    iStrData = iStrData & Chr(11) & .vspdData3.Text
                    .vspdData3.Col = .vspdData3.Col + 1
                    iStrData = iStrData & Chr(11) & .vspdData3.Text
                    iStrData = iStrData & Chr(11) & Chr(12)

                End If

            Next

			ggoSpread.SSShowData iStrData 
			
			If iStrFlag <> "" Then
				arrFlag = Split(iStrFlag, Chr(12))			
				For indx1 = 0 to Ubound(arrFlag) - 1

				    .vspdData2.Row = indx1 + 1
					.vspdData2.Col = 0
				    .vspdData2.Text = arrFlag(indx1)

				Next
			End If

            frm1.vspdData2.Redraw = True

        End If

    End With

    CopyFromData = boolExist

End Function

'=======================================================================================================
' Function Name : CheckSpread3
' Function Desc : 저장시에  관리항목 필수여부 check 하기위해 호출되는 Function
'=======================================================================================================
Function CheckSpread3()
	Dim indx
	Dim tmpDrCrFG,tmpItemSeq

	CheckSpread3 = False

	With frm1
		For jj = 1 To .vspdData.MaxRows
			.vspdData.row = jj
			.vspdData.col = C_DrCRFG
			tmpDrCrFG = Left(.vspddata.Text,1)
			.vspdData.col = C_ItemSeq
			tmpItemSeq = .vspddata.Text

	 		For indx = 1 to .vspdData3.MaxRows
			    .vspdData3.Row = indx
	 			.vspdData3.Col = 8

	 			If tmpItemSeq = .vspddata3.Text Then
					.vspdData3.Col = 14

					If (tmpDrCrFG = .vspddata3.Text) Or .vspddata3.Text = "DC" Then
  						.vspdData3.Col = 5
						If Trim(.vspdData3.Text) = "" Then
							Exit Function
			  			End If
					End If
				End If	
			Next
		Next	
	End With

	CheckSpread3 = True
End Function

'=======================================================================================================
' Function Name : DbQuery3
' Function Desc : Item 그리드 변경시 관리항목 조회 
'=======================================================================================================
Function DbQuery3(ByVal Row)

	Dim	ICurItemSeq
	Dim	IDtlRow
	Dim strVal
	Dim lngRows
	Dim strSelect
	Dim strSelect1
	Dim strWhere
	Dim IntRetCD
	Dim arrVal
	Dim arrTemp, arrTemp1
	Dim indx, indx1
	Dim tmpDrCrFg

	On Error Resume Next
	Err.Clear
	
	DbQuery3 = False

	Call DisableToolBar(parent.TBC_QUERY)	
	Call LayerShowHide(1)

	With frm1
	    If CopyFromData(.hItemSeq.Value) = True Then
'			Call DeleteHSheet(.hItemSeq.Value)
		 	Call LayerShowHide(0)
		 	Call RestoreToolBar()
		 	Call SetSpread2Color()
		 	Exit Function
		End If	
		ggoSpread.Source = frm1.vspdData2
		Call ggoSpread.ClearSpreadData()
	    .vspdData.Row = Row
	    .vspdData.Col = C_ItemSeq
	    ICurItemSeq	 = .vspdData.Text
	    .vspdData.Col = C_DrCrFg
		frm1.vspdData.Col = C_DrCrFg
		tmpDrCrFG = frm1.vspdData.text
	    .vspdData.Col = C_AcctCd
	End With

	If CommonQueryRs("ACCT_NM", " A_ACCT (NOLOCK)" , "ACCT_CD =  " & FilterVar(Frm1.vspdData.Text , "''", "S") & "", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) Then
    	frm1.vspdData.Col	= C_AcctNm
    	arrVal				= Split(lgF0, Chr(11))
		frm1.vspdData.Text	= arrVal(0)
	Else
'		frm1.vspdData.Text	= ""
		frm1.vspdData.Col	= C_AcctNm
		frm1.vspdData.Text	= ""
'		IntRetCD			= DisplayMsgBox("110100", "X", "X", "X")
		Call LayerShowHide(0)
		Call RestoreToolBar()
		Exit Function
	End If

	frm1.vspdData.Col = C_AcctCd

    strSelect =	            " B.CTRL_ITEM_SEQ,  A.CTRL_CD, A.CTRL_NM , '', '',"
    strSelect = strSelect & " CASE  WHEN A.COLM_DATA_TYPE = " & FilterVar("D", "''", "S") & "  THEN " & FilterVar("(Format : YYYY-MM-DD)", "''", "S") & "  END , " & ICurItemSeq  & ", LTrim(ISNULL(A.TBL_ID,'')),LTrim(ISNULL(A.DATA_COLM_ID,'')), "
    strSelect = strSelect & " LTrim(ISNULL(A.DATA_COLM_NM,'')),  LTrim(ISNULL(A.COLM_DATA_TYPE,'')), LTrim(ISNULL(A.DATA_LEN,'')), "
    strSelect = strSelect & " CASE WHEN B.DR_FG = " & FilterVar("Y", "''", "S") & "  AND  B.CR_FG = " & FilterVar("Y", "''", "S") & "  THEN " & FilterVar("DC", "''", "S") & "  "
'    strSelect = strSelect & " WHEN B.DR_FG = 'Y' AND 'DR'='" & Trim(tmpDrCrFG) & "' THEN 'Y' "  
'    strSelect = strSelect & " WHEN B.CR_FG = 'Y' AND 'CR'='" & Trim(tmpDrCrFG) & "' THEN 'Y' "  
    strSelect = strSelect & " WHEN B.DR_FG = " & FilterVar("Y", "''", "S") & "  AND  B.CR_FG = " & FilterVar("N", "''", "S") & "  THEN " & FilterVar("D", "''", "S") & "  "
    strSelect = strSelect & " WHEN B.DR_FG = " & FilterVar("N", "''", "S") & "  AND  B.CR_FG = " & FilterVar("Y", "''", "S") & "  THEN " & FilterVar("C", "''", "S") & "  "
    strSelect = strSelect & " END	, " & ICurItemSeq & " , "
    strSelect = strSelect & " LTrim(ISNULL(A.MAJOR_CD, '')), CHAR(8) "

    strSelect1 =ICurItemSeq & " ,B.CTRL_ITEM_SEQ,  A.CTRL_CD, A.CTRL_NM , '', '',"
    strSelect1 = strSelect1 & " CASE  WHEN A.COLM_DATA_TYPE = " & FilterVar("D", "''", "S") & "  THEN " & FilterVar("(Format : YYYY-MM-DD)", "''", "S") & "  END , " & ICurItemSeq  & ", LTrim(ISNULL(A.TBL_ID,'')),LTrim(ISNULL(A.DATA_COLM_ID,'')), "
    strSelect1 = strSelect1 & " LTrim(ISNULL(A.DATA_COLM_NM,'')),  LTrim(ISNULL(A.COLM_DATA_TYPE,'')), LTrim(ISNULL(A.DATA_LEN,'')), "
'    strSelect1 = strSelect1 & " CASE WHEN B.DR_FG = 'Y' AND  B.CR_FG = 'Y' THEN 'DC' "
'    strSelect1 = strSelect1 & " WHEN B.DR_FG = 'Y' AND 'DR'='" & Trim(tmpDrCrFG) & "' THEN 'Y' "  
'    strSelect1 = strSelect1 & " WHEN B.CR_FG = 'Y' AND 'CR'='" & Trim(tmpDrCrFG) & "' THEN 'Y' "  

    strSelect1 = strSelect1 & " CASE WHEN B.DR_FG = " & FilterVar("Y", "''", "S") & "  AND  B.CR_FG = " & FilterVar("Y", "''", "S") & "  THEN " & FilterVar("DC", "''", "S") & "  "
    strSelect1 = strSelect1 & " WHEN B.DR_FG = " & FilterVar("Y", "''", "S") & "  AND  B.CR_FG = " & FilterVar("N", "''", "S") & "  THEN " & FilterVar("D", "''", "S") & "  "  'tmpDrCrFG
    strSelect1 = strSelect1 & " WHEN B.DR_FG = " & FilterVar("N", "''", "S") & "  AND  B.CR_FG = " & FilterVar("Y", "''", "S") & "  THEN " & FilterVar("C", "''", "S") & "  "  
    strSelect1 = strSelect1 & " END	, "
    strSelect1 = strSelect1 & " LTrim(ISNULL(A.MAJOR_CD, '')), CHAR(8) "

	strWhere =  "A.CTRL_CD = B.CTRL_CD  AND B.ACCT_CD =  " & FilterVar(Frm1.vspdData.Text, "''", "S") & ""
	strWhere =  strWhere & " Order By B.CTRL_ITEM_SEQ "

	frm1.vspdData2.ReDraw = False
	
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	If CommonQueryRs2by2(strSelect, " A_CTRL_ITEM  A (NOLOCK), A_ACCT_CTRL_ASSN B (NOLOCK)" , strWhere , lgF2By2) Then
	
		ggoSpread.Source = frm1.vspdData2
		arrTemp =  Split(lgF2By2,Chr(12))
		For Indx = 0 To Ubound(arrTemp) - 1
			arrTemp(indx) = Replace(arrTemp(indx), Chr(8), indx + 1)
		Next
		lgF2By2 = Join(arrTemp,Chr(12))
		ggoSpread.SSShowData lgF2By2

		For lngRows = 1 to frm1.vspdData2.MaxRows
			frm1.vspddata2.Row	= lngRows
			frm1.vspddata2.Col	= 0
			frm1.vspddata2.Text	= ggoSpread.InsertFlag
		Next

		If CommonQueryRs2by2(strSelect1, " A_CTRL_ITEM  A (NOLOCK), A_ACCT_CTRL_ASSN  B (NOLOCK)" , strWhere , lgF2By2) Then
			ggoSpread.Source = frm1.vspdData3
			IDtlRow = frm1.vspdData3.MaxRows
			arrTemp1 =  Split(lgF2By2,Chr(12))
			For Indx1 = 0 To Ubound(arrTemp1) - 1
				arrTemp1(indx1) = Replace(arrTemp1(indx1), Chr(8), indx1 + 1)
			Next
			lgF2By2 = Join(arrTemp1,Chr(12))

			ggoSpread.SSShowData lgF2By2
			For lngRows = IDtlRow + 1 To frm1.vspdData3.MaxRows
				frm1.vspddata3.Row	= lngRows
				frm1.vspddata3.Col	= 0
				frm1.vspddata3.Text	= ggoSpread.InsertFlag
			Next
		End If

		Call SetSpread2Color()

    End If

    frm1.vspdData2.ReDraw = True

    Call LayerShowHide(0)
    Call RestoreToolBar()

	If Err.number = 0 Then
		DbQuery3 = True
	End If
	
	Set gActiveElement = document.ActiveElement

End Function

'=======================================================================================================
' Function Name : DbQueryOk3
' Function Desc : DbQuery3가 성공적일 경우 MyBizASP 에서 호출되는 Function
'=======================================================================================================
Sub DbQueryOk3()

	Call SetSpread2Color()

End Sub

'=======================================================================================================
'   Event Name : vspdData2_ButtonClicked
'   Event Desc : 관리항목 팝업버튼 클릭시 관리항목 팝업호출 
'=======================================================================================================
Sub vspdData2_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)

	Dim iFld1 
	Dim iFld2
	Dim iTable
	Dim iTempTable
	Dim iTempCtrlVal
	Dim istrCode
	Dim FldNm
	Dim strSelect, strFrom, strWhere
	Dim arrTemp
	Dim strVatType
	Dim strVatRate
	Dim lRows
	Dim lngRows
	Dim ii
	Dim Test
		
	On Error Resume Next
	
	'---------- Coding part -------------------------------------------------------------
	ggoSpread.Source = frm1.vspdData2

	With frm1.vspdData2
		If Row > 0 And Col = C_CtrlPB Then
			.Row = Row
			.Col = C_ctrlNm
			FldNm = .Text

			.Col = C_CtrlVal
			istrCode = .Text 

			.Col = C_Tableid
			iTable = Trim(.Text)

			.Col = C_Colid
			iFld1 = Trim(.Text)
		
			.Col = C_ColNm
			iFld2 = Trim(.Text)

			.Col = C_MajorCD

			IF  .Text <> "" Then
				strWhere = " Major_CD =  " & FilterVar(.Text , "''", "S") & ""
			ElseIF iTable = "B_ACCT_DEPT" Then
				If lgOrgChangeId = "" Then
			 		lgOrgChangeId = frm1.hOrgChangeId.Value	
			 	 End If		

				If  lgOrgChangeId <> "" then
					strWhere = " Org_Change_Id =  " & FilterVar(lgOrgChangeId, "''", "S") & ""
				else
					strWhere = " Org_Change_Id =  " & FilterVar(parent.gChangeOrgId , "''", "S") & ""
				end if

			ELse
				strWhere = ""	
			END IF	
	
			IF iFld1 = "BANK_ACCT_NO" Then
				For lngRows = 1 To .MaxRows
					.Row = lngRows
					.Col = C_Tableid 
					iTempTable = Trim(.Text)
					.Col = C_CtrlVal
 					iTempCtrlVal = Trim(.Text)

					If iTempTable = "B_BANK" and iTempCtrlVal <> "" Then
						strWhere = " BANK_CD LIKE  " & FilterVar(iTempCtrlVal, "''", "S") & " AND ISNULL(BANK_ACCT_PRNT, 'N') <> 'Y' "
						Exit For
					End If
					
					If iTempTable = "B_BANK_ACCT" Then
						strWhere = " ISNULL(BANK_ACCT_PRNT, 'N') <> 'Y' "
					End If					
				Next
			END IF						
			If iTable <> "" Then 
 				Call OpenCtrlPB(iTable, iFld1, iFld2, istrCode, FldNm, strWhere)
			End if

			frm1.vspddata2.Col = C_CtrlCd
			If Trim(frm1.vspddata2.Text) = "V4" Then

				frm1.vspdData2.Col = C_CtrlVal
				strVatType = Trim(frm1.vspdData2.text)
				If Trim(strVatType) <> "" Then			
				
					strSelect	= "reference"
					strFrom		= "b_configuration"
					strWhere	= "major_cd = " & FilterVar("B9001", "''", "S") & "  and seq_no = 1 and minor_cd =  " & FilterVar(strVatType , "''", "S") & ""
					
					If CommonQueryRs(strSelect, strFrom, strWhere, lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) Then
						arrTemp = Split(lgF0, chr(11))
						strVatRate = arrTemp(0)				
					End If
					
					frm1.vspddata2.Col = C_CtrlCd				
					For ii = i To frm1.vspdData2.MaxRows
						frm1.vspdData2.Row = ii
						If Trim(frm1.vspddata2.Text) = "V7" Then
							frm1.vspddata2.Col = C_CtrlVal
							frm1.vspddata2.Text = strVatRate
						End If
					Next			
				End If
			End If 

		End If
	End With
	
	'매출부가세에서 세금계산서를 setting하면 자동으로 vat_rate가 hidden으로 복사되도록 한다.
	With frm1
	.vspdData2.Col = C_CtrlCd        
        	For lRows = 1 To .vspdData2.MaxRows        
            	.vspdData2.Row = lRows								'ActiveRow설정 
				If Trim(.vspdData2.Text) = "V7" Then
					.vspddata2.action = 0
					CopyToHSheet lRows

				End If      
        	Next
	End With

End Sub

'=======================================================================================================
'   Event Name : vspdData2_Change
'   Event Desc : 관리항목 그리드 데이타 변경시 입력값에대한 유효성 Check
'=======================================================================================================
Sub vspdData2_Change(ByVal Col, ByVal Row)

	Dim iLen
	Dim sPreCtrlVal
	Dim IntRetCD

   	ggoSpread.Source = frm1.vspdData2
	ggoSpread.UpdateRow Row

	frm1.vspdData2.Row = Row
	frm1.vspdData2.Col = 0

	Select Case Col
		Case   C_CtrlVal
	    '----------------------------------
		' 입력된 관리항목의 DataType Check yyyy-mm-dd
		'----------------------------------
		    frm1.vspdData2.Col = C_DataType

	        If Trim(frm1.vspdData2.Text) = "D" Then
				frm1.vspdData2.Col = C_CtrlVal
				sPreCtrlVal = frm1.vspdData2.Text
				If IsDate(frm1.vspdData2.Text) = False or IsNumeric(Mid(frm1.vspdData2.Text,1,4)) = False or _
					IsNumeric(Mid(frm1.vspdData2.Text,6,2)) = False or _
					IsNumeric(Mid(frm1.vspdData2.Text,9,2)) = False or _
					Mid(frm1.vspdData2.Text,5,1) <> "-" or _
					Mid(frm1.vspdData2.Text,8,1) <> "-" or _
					Mid(frm1.vspdData2.Text,1,4) < "1900" Then
						frm1.vspdData2.Text = sPreCtrlVal
						IntRetCD = DisplayMsgBox("174223", "X", "X", "X")							'필수입력 check!!
						' 입력하신 날짜는 부적합합니다.
						frm1.vspdData2.Text = ""
						Exit Sub
				End If
			ElseIf Trim(frm1.vspdData2.Text) = "N" Then
				frm1.vspdData2.Col = C_CtrlVal
				sPreCtrlVal = frm1.vspdData2.Text
				If IsNumeric(frm1.vspdData2.Text) = False Then
					frm1.vspdData2.Text = sPreCtrlVal
					IntRetCD = DisplayMsgBox("229924", "X", "X", "X")								'필수입력 check!!
					' 숫자를 입력하십시오 
					frm1.vspdData2.Text = ""
					Exit Sub
				Else
				frm1.vspdData2.Text = replace(formatnumber(frm1.vspdData2.Text,2), parent.gComNumDec & "00", "")
				End If
	        End If

	        '------------------------------------
	        ' 입력된 관리항목의 길이Check
	        '------------------------------------
	        frm1.vspdData2.Col = C_CtrlVal

	        iLen = Len(frm1.vspdData2.Text)
			sPreCtrlVal = frm1.vspdData2.Text
	        frm1.vspdData2.Col = C_DataLen

	        If iLen > Int(frm1.vspdData2.Text) Then
				frm1.vspdData2.Text = sPreCtrlVal
				IntRetCD = DisplayMsgBox("110320", "X", "X", "X")									'필수입력 check!!
			'  관리항목값의 길이를 확인하십시오.
				frm1.vspdData2.Col = C_CtrlVal
				frm1.vspdData2.Text = ""
				Exit Sub
	        End If

	        frm1.vspdData2.Col = C_DataType

	        If Trim(frm1.vspdData2.Text) <> "D" And Trim(frm1.vspdData2.Text) <> "N" Then
				FindCtrlNM   Row																			'관리항목값을 check하고 관리항목명을 찾아준다.
			End If
    End Select
'	CopyToHSheet Row
	Call CopyToHSheet2(frm1.vspdData.ActiveRow,Row)

    lgBlnFlgChgValue = True

End Sub

'==========================================================================================
'   Event Name : vspdData2_Click
'   Event Desc : This event is spread sheet data changed
'==========================================================================================
Sub vspdData2_Click(ByVal Col, ByVal Row)

	Call SetPopupMenuItemInf("0000111111")

	gMouseClickStatus = "SP2C"	'Split 상태코드 
	   
	Set gActiveSpdSheet = frm1.vspdData2

	If Row <= 0 Then                                                    'If there is no data.
		ggoSpread.Source = frm1.vspdData2
		If lgSortKey = 1 Then
			ggoSpread.SSSort Col
			lgSortKey = 2
		Else
			ggoSpread.SSSort Col, lgSortKey
			lgSortKey = 1
		End If
		Exit Sub
	End If

End Sub
'========================================================================================================
'   Event Name : vspdData2_ColWidthChange
'   Event Desc :
'========================================================================================================
Sub vspdData2_ColWidthChange(ByVal pvCol1, ByVal pvCol2)

    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

End Sub
'========================================================================================================
'   Event Name : vspdData2_DblClick
'   Event Desc : 
'========================================================================================================
Sub vspdData2_DblClick(ByVal Col, ByVal Row)

    Dim iColumnName

    If Row <= 0 Then
		Exit Sub
    End If
    If frm1.vspdData2.MaxRows = 0 Then
		Exit Sub
	End If
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End   ) --------------------------------------------------------------

End Sub
'==========================================================================================
'   Event Desc : Spread Split 상태코드 
'==========================================================================================
Sub vspdData2_MouseDown(Button, Shift, X, Y)
	If Button = 2 And gMouseClickStatus = "SP2C" Then
		gMouseClickStatus = "SP2CR"
	End If
End Sub
'========================================================================================================
'   Event Name : vspdData2_GotFocus
'   Event Desc : This event is spread sheet data changed
'========================================================================================================
Sub vspdData2_GotFocus()
    ggoSpread.Source = Frm1.vspdData2
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End   ) --------------------------------------------------------------
End Sub
'=======================================================================================================
'   Event Name : vspdData2_onfocus
'   Event Desc :
'=======================================================================================================
Sub vspdData2_onfocus()

'    If lgIntFlgMode <> parent.OPMD_UMODE Then
'        Call SetToolbar("1110100000011111")                                     '버튼 툴바 제어 
'    Else
'        Call SetToolbar("1111100000011111")                                     '버튼 툴바 제어 
'    End If

End Sub
'========================================================================================================
'   Event Name : vspdData2_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData2_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )

    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )
    Call GetCtrlSpreadColumnPos("A")

End Sub

'=======================================================================================================
'   Event Name : SetGridFocus2
'   Event Desc :
'=======================================================================================================
Sub SetGridFocus2()	

	With frm1
		.vspdData2.Row		= 1
		.vspdData2.Col		= C_DtlSeq
		.vspdData2.Action	= 1
	End With

End Sub

'==========================================================================================
'   Event Desc : Grid의 Max Count 를 찾는다.
'==========================================================================================
Function MaxSpreadVal(ByVal objSpread, ByVal intCol, byval Row)

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

		If tmpval > MaxValue   Then
		   MaxValue = cdbl(tmpVal)
		End If
	Next

	MaxValue = MaxValue + 1

	objSpread.row	= row
	objSpread.col	= intCol
	objSpread.Text	= MaxValue

end Function

'==========================================================================================
'   Event Desc : 현금계정을 가지고 온다.
'==========================================================================================
Function GetCheckAcct()

	If CommonQueryRs( "ACCT_CD", "A_ACCT" , " ACCT_TYPE = " & FilterVar("A0", "''", "S") & "  AND DEL_FG <> " & FilterVar("Y", "''", "S") & " " , lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) Then
		arrVal		= Split(lgF0, Chr(11))
		lgCashAcct	= arrVal(0)
	End If

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
			   If IsNumeric(UNICDbl(pObject.Text)) Then
			      iSum = iSum + UNICDbl(pObject.Text)
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


' 권한관리 추가 ==============================================================
'==========================================================================================
'	Name : SetAuthorityFlag
'	Description :
'==========================================================================================
Sub SetAuthorityFlag()
	If CommonQueryRs("TOP 1 USR_ID", "Z_USR_AUTHORITY_VALUE", "USR_ID =  " & FilterVar(parent.gUsrId , "''", "S") & " AND MODULE_CD = " & FilterVar("A", "''", "S") & "  ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) Then

	    If UCase(parent.gUsrId) = UCase(Replace(lgF0,Chr(11),"")) Then
		
	      lgAuthorityFlag = "Y"
	  	Else
	    	lgAuthorityFlag = "N"
	    End If
	Else
	  	lgAuthorityFlag = "N"
	End If
End Sub

' 권한관리 추가 END ==========================================================

 '========================================== Open???()  =============================================
'	Name : Open???()
'	Description : 중복되어 있는 PopUp을 재정의, 재정의가 필요한 경우는 반드시 CommonPopUp.vbs 와 
'				  ManufactPopUp.vbs 에서 Copy하여 재정의한다.
'=====================================================================================================
Function AutorityMakeSql(Byval iWhere, Byval strCode1,Byval strCode2,Byval strCode3,Byval strCode4,Byval strCode5)

	Dim arrstrRet(2)

	Select Case iWhere
		Case "DEPT"

			If lgAuthorityFlag = "Y" Then		                                    '권한관리 추가 
				arrstrRet(0) = "B_ACCT_DEPT A, Z_USR_AUTHORITY_VALUE B "     		'권한관리 추가 
			Else																	'권한관리 추가 
				arrstrRet(0) = "B_ACCT_DEPT A "    									' TABLE 명칭 
			End If

			If lgAuthorityFlag = "Y" Then		                                    '권한관리 추가 
				arrstrRet(1) = "ORG_CHANGE_ID =  " & FilterVar(strCode1, "''", "S") & " AND A.DEPT_CD = B.CODE_VALUE AND B.USR_ID =  " & FilterVar(parent.gUsrId , "''", "S") & " AND B.MODULE_CD = " & FilterVar("A", "''", "S") & " "		 '권한관리 추가 
			Else                                                              '권한관리 추가 
				arrstrRet(1) = "ORG_CHANGE_ID =  " & FilterVar(strCode1, "''", "S") & ""
			End If

		Case "DEPT_ITEM"

			If lgAuthorityFlag = "Y" Then																		'권한관리 추가 
				arrstrRet(0) = " B_ACCT_DEPT A, B_COST_CENTER B, B_BIZ_AREA C, Z_USR_AUTHORITY_VALUE D "		'권한관리 추가 
			Else																								'권한관리 추가 
				arrstrRet(0) = " B_ACCT_DEPT A, B_COST_CENTER B, B_BIZ_AREA C "									' TABLE 명칭 
			End If

			If lgAuthorityFlag = "Y" Then																		'권한관리 추가 
  				arrstrRet(1) = " B.COST_CD = A.COST_CD AND B.BIZ_AREA_CD = C.BIZ_AREA_CD " & _
  							   " AND C.BIZ_AREA_CD = (SELECT F.BIZ_AREA_CD" & _
  												   " FROM B_ACCT_DEPT D, B_COST_CENTER E, B_BIZ_AREA F" & _
  												   " WHERE D.DEPT_CD =  " & FilterVar(strCode2, "''", "S") & "" & _
  												   " AND D.ORG_CHANGE_ID =  " & FilterVar(strCode1, "''", "S") & "" & _
  												   " AND E.COST_CD = D.COST_CD AND E.BIZ_AREA_CD = F.BIZ_AREA_CD) AND A.DEPT_CD = D.CODE_VALUE AND D.USR_ID =  " & FilterVar(parent.gUsrId , "''", "S") & " AND D.MODULE_CD = " & FilterVar("A", "''", "S") & " " & _
  							   " AND A.ORG_CHANGE_ID =  " & FilterVar(strCode1, "''", "S") & ""	                  '권한관리 추가 
			else                                                                                        '권한관리 추가 
				arrstrRet(1) = " B.COST_CD = A.COST_CD AND B.BIZ_AREA_CD = C.BIZ_AREA_CD " & _
							   " AND C.BIZ_AREA_CD = (SELECT F.BIZ_AREA_CD" & _
													   " FROM B_ACCT_DEPT D, B_COST_CENTER E, BIZ_AREA_CD F" & _
													   " WHERE D.DEPT_CD =  " & FilterVar(strCode2, "''", "S") & "" & _
													   " AND D.ORG_CHANGE_ID =  " & FilterVar(strCode1, "''", "S") & "" & _
													   " AND E.COST_CD = D.COST_CD AND E.BIZ_AREA_CD = F.BIZ_AREA_CD)" & _
								" AND A.ORG_CHANGE_ID =  " & FilterVar(strCode1, "''", "S") & ""
			End If

	End Select

	AutorityMakeSql = arrstrRet

End Function


'=========================================================================================================
' 매입부가세. 매출부가세의 경우 자동입력을 위해서 만든 Function
' 최초작성자 : 박심서 
'=========================================================================================================
Function AutoInputDetail(ByVal strAcctCd, ByVal strDeptCd, ByVal strDate, ByVal Row)

	Dim strIoFg, strIoFgNm
	Dim strVatType, strVatTypeNm
	Dim strVatRate
	Dim strYYYY, strMM , strDD

'	Dim lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6
	Dim strSelect
	Dim strFrom
	Dim strWhere
	Dim arrTemp
	Dim strInVat, strOutVat, strAcctType, strBizAreaCd, strBizAreaNm
	Dim strOrgChangeID
	Dim indx

	Dim lngRow, lngCol
	Dim str
	Dim strCtrlCd
	Dim strDtlSeq

	lgBlnFlgChgValue = True

	strSelect	= "acct_type"
	strFrom		= "a_acct"
	strWhere	= "acct_cd =  " & FilterVar(strAcctCd , "''", "S") & ""

	If CommonQueryRs(strSelect, strFrom, strWhere, lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) Then
		arrTemp		= Split(lgF0, chr(11))
		strAcctType =  arrTemp(0)
	End If

	If Trim(strDeptCd) <> "" And Trim(strDate) <> ""  Then
		strSelect	=			 " dept_cd, org_change_id, internal_cd "    		
		strFrom		=			 " b_acct_dept(NOLOCK) "		
		strWhere	=			 " dept_Cd =   " & FilterVar(strDeptCd, "''", "S") & "  "
		strWhere	= strWhere & " and org_change_id = (select distinct org_change_id "			
		strWhere	= strWhere & " from b_acct_dept where org_change_dt = ( select max(org_change_dt)"
		strWhere	= strWhere & " from b_acct_dept where org_change_dt <= " & FilterVar(UniConvDateToYYYYMMDD(strDate, parent.gDateFormat,""), "''", "S") & "))"			

		If CommonQueryRs(strSelect, strFrom, strWhere, lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) Then
			arrTemp		= Split(lgF1, chr(11))
			strOrgChangeID =  arrTemp(0)
		End If
	End If

	If Trim(strDeptCd) <> "" And Trim(strOrgChangeID) <> "" Then
		strSelect	= "tax_biz_area_cd, tax_biz_area_nm "
		strFrom		= "b_tax_biz_area "
		strWhere	= "tax_biz_area_cd in  (SELECT	C.REPORT_BIZ_AREA_CD "
		strWhere	= strWhere & "			FROM	B_ACCT_DEPT A, B_COST_CENTER B, B_BIZ_AREA C "', b_tax_biz_area D "
		strWhere	= strWhere & "			WHERE	A.DEPT_CD =  " & FilterVar(strDeptCd, "''", "S") & " "
		strWhere	= strWhere & "			AND		A.Org_change_id =  " & FilterVar(strOrgChangeID, "''", "S") & " "
		strWhere	= strWhere & "			AND A.COST_CD = B.COST_CD "
		strWhere	= strWhere & "			AND B.BIZ_AREA_CD = C.BIZ_AREA_CD "
		'strWhere	= strWhere & "			AND C.REPORT_BIZ_AREA_CD = D.tax_biz_area_cd "
		strWhere	= strWhere & "		    ) "

		If CommonQueryRs(strSelect, strFrom, strWhere, lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) Then
				arrTemp			= Split(lgF0, chr(11))
				strBizAreaCd	= arrTemp(0)
				arrTemp			= Split(lgF1, chr(11))
				strBizAreaNm	= arrTemp(0)
		End If
	End If

	If UCase(Trim(strAcctType)) = "VP" Or UCase(Trim(strAcctType)) = "VR" Then
		'매입매출부가세인경우 계산서유형 가져오기 
		If strAcctType = "VP" Then
			strIoFg		= "I"
			strSelect	= "MINOR_NM"
			strFrom		= "B_MINOR"
			strWhere	= "MAJOR_CD = " & FilterVar("A1003", "''", "S") & "  AND MINOR_CD = " & FilterVar("I", "''", "S") & " "
			If CommonQueryRs(strSelect, strFrom, strWhere, lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) Then
				arrTemp		= Split(lgF0, chr(11))
				strIoFgNm	= arrTemp(0)
			End If

		ElseIf strAcctType = "VR" Then
			strIoFg		= "O"
			strSelect	= "MINOR_NM "
			strFrom		= "B_MINOR "
			strWhere	= "MAJOR_CD = " & FilterVar("A1003", "''", "S") & "  AND MINOR_CD = " & FilterVar("O", "''", "S") & " "
			If CommonQueryRs(strSelect, strFrom, strWhere, lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) Then
				arrTemp		= Split(lgF0, chr(11))
				strIoFgNm	= arrTemp(0)
			End If
		End If

		'매입매출부가세인 경우만 경우만 
		If strAcctType = "VP" Or strAcctType = "VR" Then
			'계산서일 
			Call ExtractDateFrom(strDate, parent.gDateFormat, parent.gComDateType, strYYYY, strMM, strDD)
			strDate = strYYYY & "-" & strMM & "-" & strDD

			'계산서타입 
			frm1.vspdData.row = row
			If C_VATTYPE = "" Then
				strVatType		= ""
				strVatTypeNm	= ""
			Else
				frm1.vspdData.Col	= C_VATTYPE
				strVatType			= Trim(frm1.vspdData.Text)
				frm1.vspdData.Col	= C_VATNM
				strVatTypeNm		= Trim(frm1.vspdData.Text)
			End If

			'부가세율 
			If Trim(strVatType) <> "" Then

				strSelect	= "REFERENCE"
				strFrom		= "B_CONFIGURATION"
				strWhere	= "MAJOR_CD = " & FilterVar("B9001", "''", "S") & "  AND SEQ_NO = 1 AND MINOR_CD =  " & FilterVar(strVatType , "''", "S") & ""

				If CommonQueryRs(strSelect, strFrom, strWhere, lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) Then
					arrTemp		= Split(lgF0, chr(11))
					strVatRate	= arrTemp(0)
				End If

			End If

		End If

		lngRow = frm1.vspdData2.MaxRows
		For indx = 1 to lngRow
			frm1.vspddata2.Col	= C_CtrlCd
			frm1.vspddata2.Row	= indx
			strCtrlCd			= UCase(Trim(frm1.vspddata2.Text))

			frm1.vspddata2.Col = C_CtrlVal

			Select Case strCtrlCd

				'Case "V1"
				'	If Trim(frm1.vspddata2.Text) = "" Then
				'		frm1.vspddata2.Text		= ""
				'	End If

				Case "V2"
					If Trim(frm1.vspddata2.Text) = "" Then
						frm1.vspddata2.Text		= strDate
					End If

				Case "V3"
					If Trim(frm1.vspddata2.Text) = "" Then
						frm1.vspddata2.Text		= strIoFg
						frm1.vspddata2.Col		= C_CtrlValNm
						frm1.vspddata2.Text		= strIoFgNm
					End If

				Case "V4"
					frm1.vspddata2.Text		= strVatType
					frm1.vspddata2.Col		= C_CtrlValNm
					frm1.vspddata2.Text		= strVatTypeNm

				Case "V5"
					If Trim(frm1.vspddata2.Text) = "" Then
						frm1.vspddata2.Text		= strBizAreaCd
						frm1.vspddata2.Col		= C_CtrlValNm
						frm1.vspddata2.Text		= strBizAreaNm
					End If

				'Case "V6"
				'	If Trim(frm1.vspddata2.Text) = "" Then
				'		frm1.vspddata2.Text		= ""
				'	End If

				Case "V7"
					frm1.vspddata2.Text		= strVatRate

			End Select

		Next
	End If

End Function

'=======================================================================================================
'   Function Name : CopyToHSheet2
'   Function Desc : 관리항목그리드의 Value를 자동settgin할때 Hidden Grid로 복사하기(CopyToHSheet의 ActiveRow의 맹점보완)
'   최초작성자      : 박심서 
'=======================================================================================================
Sub CopyToHSheet2(ByVal MasterRow, ByVal DetailRow)

	Dim lRow
	Dim iCols

	With frm1

	    lRow = FindData2(MasterRow,DetailRow)
	    If lRow > 0 Then

            .vspdData3.Row = lRow
            .vspdData2.Row = DetailRow
            .vspdData3.Col = 0
            .vspdData2.Col = 0
            .vspdData3.Text = .vspdData2.Text

			.vspdData2.Col = C_DtlSeq
			.vspdData3.Col = 2
			.vspdData3.Text = .vspdData2.Text
					    
			.vspdData2.Col = C_CtrlCd
			.vspdData3.Col = .vspdData3.Col + 1
			.vspdData3.Text = .vspdData2.Text
					    
			.vspdData2.Col = C_CtrlNm
			.vspdData3.Col = .vspdData3.Col + 1
			.vspdData3.Text = .vspdData2.Text
					    
			.vspdData2.Col = C_CtrlVal
			.vspdData3.Col = .vspdData3.Col + 1
			.vspdData3.Text = .vspdData2.Text
					    
			.vspdData2.Col = C_CtrlPB
			.vspdData3.Col = .vspdData3.Col + 1
			.vspdData3.Text = .vspdData2.Text
					    
			.vspdData2.Col = C_CtrlValNm
			.vspdData3.Col = .vspdData3.Col + 1
			.vspdData3.Text = .vspdData2.Text
					    
			.vspdData2.Col = C_Seq
			.vspdData3.Col = .vspdData3.Col + 1
			.vspdData3.Text = .vspdData2.Text
					    
			.vspdData2.Col = C_Tableid
			.vspdData3.Col = .vspdData3.Col + 1
			.vspdData3.Text = .vspdData2.Text
					    
			.vspdData2.Col = C_Colid
			.vspdData3.Col = .vspdData3.Col + 1
			.vspdData3.Text = .vspdData2.Text
					    
			.vspdData2.Col = C_ColNm
			.vspdData3.Col = .vspdData3.Col + 1
			.vspdData3.Text = .vspdData2.Text
					    
			.vspdData2.Col = C_Datatype
			.vspdData3.Col = .vspdData3.Col + 1
			.vspdData3.Text = .vspdData2.Text
					    
			.vspdData2.Col = C_DataLen
			.vspdData3.Col = .vspdData3.Col + 1
			.vspdData3.Text = .vspdData2.Text
					    
			.vspdData2.Col = C_DRFg
			.vspdData3.Col = .vspdData3.Col + 1
			.vspdData3.Text = .vspdData2.Text
					    
			.vspdData2.Col = C_MajorCd
			.vspdData3.Col = .vspdData3.Col + 1
			.vspdData3.Text = .vspdData2.Text
			
			.vspdData2.Col = C_MajorCd + 1
			.vspdData3.Col = .vspdData3.Col + 1
			.vspdData3.Text = .vspdData2.Text

        End If

	End With

	frm1.vspdData.Row = MasterRow												'frm1.vspdData.ActiveRow
	frm1.vspdData.Col = 0

	If frm1.vspdData.Text <> ggoSpread.InsertFlag and frm1.vspdData.Text <> ggoSpread.DeleteFlag Then
   	    frm1.vspdData.Text = ggoSpread.UpdateFlag
	End if

End Sub

'=======================================================================================================
'   Function Name : FindData2
'   Function Desc : 현재의 Item, Dtl에 해당하는 Hidden Grid의 Index를 Return
'                   관리항목그리드의 Value를 자동settgin할때 Hidden Grid로 복사하기(CopyToHSheet의 ActiveRow의 맹점보완)
'=======================================================================================================
Function FindData2(MasterRow,DetailRow)

	Dim strApNo
	Dim strItemSeq
	Dim strDtlSeq
	Dim lRows

    FindData2 = 0

    With frm1

        For lRows = 1 To .vspdData3.MaxRows

            .vspdData3.Row = lRows
            .vspdData3.Col = 1
            strItemSeq = .vspdData3.Text
            .vspdData3.Col = 2
            strDtlSeq = .vspdData3.Text

            .vspdData.Row = MasterRow'frm1.vspdData.ActiveRow
            .vspdData2.Row = DetailRow

            .vspdData.Col = C_ItemSeq
            If strItemSeq = .vspdData.Text Then

                .vspdData2.Col = C_DtlSeq
                If strDtlSeq = .vspdData2.Text Then

                    FindData2 = lRows
                    Exit Function

                End If

            End If
        Next

    End With

End Function
'=======================================================================================================
'   Function Name : FindExchRate
'   Function Desc : 1.날짜, Row를 입력받아 날짜에 해당하는 환율정보를 읽어온다.
'=======================================================================================================

Function FindExchRate(Byval strDate, Byval FromCurrency,Byval Row )
	Dim strSelect, strFrom, strWhere
	Dim arrTemp
	Dim strExchFg
	Dim strExchRate
	Dim lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6	

	strSelect	= "b.minor_cd"
	strFrom		= "b_company a, b_minor b"
	strWhere	= "b.major_cd = " & FilterVar("a1004", "''", "S") & "  and	a.xch_rate_fg = b.minor_cd"
	If CommonQueryRs(strSelect, strFrom, strWhere, lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) Then				
		arrTemp = Split(lgF0, chr(11))
		strExchFg =  arrTemp(0)
	End If

	If UCase(strExchFg) <> "D" Then 	' Fixed Exchange Rate
		strDate = Mid(strDate, 1, 6)
		strSelect	= "std_rate"
		strFrom		= "b_monthly_exchange_rate (noLock) "
		strWhere	= "from_currency =  " & FilterVar(FromCurrency , "''", "S") & ""
		strWhere	= strWhere & " And to_currency   =  " & FilterVar(parent.gCurrency , "''", "S") & ""
		strWhere	= strWhere & " And apprl_yrmnth  =  " & FilterVar(strDate , "''", "S") & ""

		If CommonQueryRs(strSelect, strFrom, strWhere, lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) Then				
			arrTemp = Split(lgF0, chr(11))
			strExchRate =  arrTemp(0)
			frm1.vspdData.row  = Row
			frm1.vspdData.Col  = C_ExchRate
			frm1.vspdData.Text = UNIConvNumPCToCompanyByCurrency(strExchRate, parent.gCurrency, parent.ggExchRateNo, parent.gLocRndPolicyNo, "X")
		Else
			IntRetCD = DisplayMsgBox("121600", "X", "X", "X")
		End If
	Else					' Floating Exchange Rate

		strSelect	= "top 1 std_rate"
		strFrom		= "b_daily_exchange_rate (noLock) "
		strWhere	= "from_currency =  " & FilterVar(FromCurrency , "''", "S") & ""
		strWhere	= strWhere & " And to_currency   =  " & FilterVar(parent.gCurrency , "''", "S") & ""
		strWhere	= strWhere & " And apprl_dt  <= convert(char(21), " & FilterVar(strDate, "''", "S") & ", 20) order by apprl_dt desc"

		If CommonQueryRs(strSelect, strFrom, strWhere, lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) Then				
			arrTemp = Split(lgF0, chr(11))
			strExchRate =  arrTemp(0)
			frm1.vspdData.row  = Row
			frm1.vspdData.Col  = C_ExchRate
			frm1.vspdData.Text = UNIConvNumPCToCompanyByCurrency(strExchRate, parent.gCurrency, parent.ggExchRateNo, parent.gLocRndPolicyNo, "X")
		Else
			IntRetCD = DisplayMsgBox("121500", "X", "X", "X")
		End If
	End If
	
End Function    
'=======================================================================================================
'   Function Name : AcctCheck
'   Function Desc : 1.입출금 전표일때 계정이 현금계정으로 입력되는지 check
'                   2.미결관리를 하는 회사일 경우 미결반제계정이 들어오는지 확인.
'=======================================================================================================
Function AcctCheck(Byval Acctcd, Byval Inputtype, Byval DrCrFg )
	Dim arrTemp
	Dim strOpenAcctFg
	
	Dim arrTemp1
	Dim strDrCrFg
	Dim strMgntFg

	AcctCheck = False
	
	IF (Inputtype = "01" Or Inputtype = "02" ) and Acctcd = lgCashAcct then
		IntRetCD = DisplayMsgBox("113106", "X", "X", "X")		
		frm1.vspdData.Text = ""
		frm1.vspdData.Col = C_AcctNm
		frm1.vspdData.Text = ""		
		Exit Function		
	END IF						
	
	If CommonQueryRs(" ISNULL(OPEN_ACCT_FG," & FilterVar("N", "''", "S") & " ) ", " B_COMPANY " , "1 = 1 ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) Then
		arrTemp = Split(lgF0, chr(11))		
	End If	

	IF arrTemp(0) = "Y" AND Trim(Acctcd) <> "" AND Trim(DrCrFg) <> "" Then					
		If CommonQueryRs(" BAL_FG, ISNULL(MGNT_FG," & FilterVar("N", "''", "S") & " ) ", " A_ACCT(NOLOCK) " , " Acct_cd =  " & FilterVar(Acctcd , "''", "S") & "", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) Then
			arrTemp1 = Split(lgF0, chr(11))		
			strDrCrFg = arrTemp1(0)		
			arrTemp1 = Split(lgF1, chr(11))		
			strMgntFg = arrTemp1(0)		
		End If	
		
		If strMgntFg = "Y" AND DrCrFg <> strDrCrFg Then
			IntRetCD = DisplayMsgBox("119306", "X", "X", "X")		
			frm1.vspdData.Text = ""
			frm1.vspdData.Col = C_AcctNm
			frm1.vspdData.Text = ""		
			Exit Function		
		END If	
	End IF	
	
	AcctCheck = True
End Function

