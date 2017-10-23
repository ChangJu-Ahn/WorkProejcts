'=======================================================================================================
'	관리항목 그리드 상수선언 
'=======================================================================================================
Const C_HMaxCols_2 = 16
Const C_NoteSep_2  = ","							'비고 seperate.

Dim C_DtlSeq_2
Dim C_CtrlCd_2
Dim C_CtrlNm_2
Dim C_CtrlVal_2
Dim C_CtrlPB_2
Dim C_CtrlValNm_2
Dim C_Seq_2
Dim C_Tableid_2
Dim C_Colid_2
Dim C_ColNm_2
Dim C_Datatype_2
Dim C_DataLen_2
Dim C_DRFg_2
Dim C_HItemSeq_2
Dim C_MajorCd_2

Dim lgCashAcct_2									'현금계정을 미리 저장한다.
Dim lgAuthorityFlag_2								'권한관리 추가 

'========================================================================================================
' Name : initSpreadPosVariables()
' Desc : Initialize Column Const value
'========================================================================================================
Sub initCtrlSpreadPosVariables2()

	C_DtlSeq_2    = 1
	C_CtrlCd_2    = 2
	C_CtrlNm_2    = 3
	C_CtrlVal_2   = 4
	C_CtrlPB_2    = 5
	C_CtrlValNm_2 = 6
	C_Seq_2       = 7
	C_Tableid_2   = 8
	C_Colid_2     = 9
	C_ColNm_2     = 10
	C_Datatype_2  = 11
	C_DataLen_2   = 12
	C_DRFg_2      = 13
	C_HItemSeq_2  = 14
	C_MajorCd_2   = 15

End Sub
'=======================================================================================================
'   Event Name : InitCtrlSpread()
'   Event Desc : 관리항목 그리드 초기화 
'=======================================================================================================
Sub InitCtrlSpread2()

	Call initCtrlSpreadPosVariables2()
	
    With frm1

		ggoSpread.Source			= .vspddata5
		ggoSpread.Spreadinit "V20021217",,parent.gAllowDragDropSpread

		.vspddata5.ReDraw			= False
		
'		.vspddata5.AutoClipboard	= False
		.vspddata5.MaxCols			= C_MajorCd_2 + 1

		Call ggoSpread.ClearSpreadData()

		Call AppendNumberPlace("6","3","0")
		Call GetCtrlSpreadColumnPos2("A")

		ggoSpread.SSSetFloat	C_DtlSeq_2,		"NO" ,				     6,"6",	ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gComNum1000,	parent.gComNumDec,	2,	,	,	"0",	"999"
		ggoSpread.SSSetEdit		C_CtrlCd_2,		"관리항목",			10,	2
		ggoSpread.SSSetEdit		C_CtrlNm_2,		"관리항목명",		30,	3
		ggoSpread.SSSetEdit		C_CtrlVal_2,	"관리항목 VALUE",	32,	,		,							30,							2
		ggoSpread.SSSetButton	C_CtrlPB_2
		ggoSpread.SSSetEdit		C_CtrlValNm_2,	"관리항목 VALUE명",	45
		ggoSpread.SSSetEdit		C_Seq_2,		"A",				     8,	,		,							3
		ggoSpread.SSSetEdit		C_Tableid_2,	"B",					32
		ggoSpread.SSSetEdit		C_Colid_2,		"C",					32
		ggoSpread.SSSetEdit		C_ColNm_2,		"D",					32
		ggoSpread.SSSetEdit		C_Datatype_2,	"E",					 2
		ggoSpread.SSSetFloat	C_DataLen_2,	"F",					 3,"6",	ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gComNum1000,	parent.gComNumDec,	,	,	,	"0",	"999"
		ggoSpread.SSSetEdit		C_DRFg_2,		"G",					 1
		ggoSpread.SSSetEdit		C_HItemSeq_2,   "H",					 1
		ggoSpread.SSSetEdit		C_MajorCd_2,	"I",					 1

		Call ggoSpread.MakePairsColumn(C_CtrlVal_2,C_CtrlPB_2)

		Call ggoSpread.SSSetColHidden(C_CtrlCd_2,C_CtrlCd_2,True)
		Call ggoSpread.SSSetColHidden(C_Seq_2,C_Seq_2,True)
		Call ggoSpread.SSSetColHidden(C_Tableid_2,C_Tableid_2,True)
		Call ggoSpread.SSSetColHidden(C_Colid_2,C_Colid_2,True)
		Call ggoSpread.SSSetColHidden(C_ColNm_2,C_ColNm_2,True)
		Call ggoSpread.SSSetColHidden(C_Datatype_2,C_Datatype_2,True)
		Call ggoSpread.SSSetColHidden(C_DataLen_2,C_DataLen_2,True)
		Call ggoSpread.SSSetColHidden(C_DRFg_2,C_DRFg_2,True)
		Call ggoSpread.SSSetColHidden(C_HItemSeq_2,C_HItemSeq_2,True)
		Call ggoSpread.SSSetColHidden(C_MajorCd_2,C_MajorCd_2,True)
		Call ggoSpread.SSSetColHidden(.vspddata5.MaxCols,.vspddata5.MaxCols,True)

		.vspddata5.ReDraw = True

    End With
    
	Call CtrlSpreadLock2("X","X", -1, -1)

End Sub

'=======================================================================================================
'   Event Name : InitCtrlHSpread2()
'   Event Desc : 관리항목 그리드 초기화 
'=======================================================================================================
Sub InitCtrlHSpread2()
	
	frm1.vspddata6.ReDraw = False
	ggoSpread.Source = frm1.vspddata6
	Call ggoSpread.ClearSpreadData()
	frm1.vspddata6.MaxCols = C_HMaxCols_2
	frm1.vspddata6.ReDraw = True

End Sub

'=======================================================================================================
' Function Name : CtrlSpreadLock2
' Function Desc : 관리항목 그리드 Lock
'=======================================================================================================
Sub CtrlSpreadLock2(Byval stsFg, Byval Index, ByVal lRow  , ByVal lRow2 )

	Dim objSpread

	With frm1

		ggoSpread.Source	= .vspddata5

		.vspddata5.Redraw	= False

		ggoSpread.SpreadLock  C_DtlSeq_2,    lRow,  C_DtlSeq_2,    lRow2
		ggoSpread.SpreadLock  C_CtrlCd_2,    lRow,  C_CtrlCd_2,    lRow2
		ggoSpread.SpreadLock  C_CtrlNm_2,    lRow,  C_CtrlNm_2,    lRow2
		ggoSpread.SpreadLock  C_CtrlValNm_2, lRow,  C_CtrlValNm_2, lRow2
		    		
		.vspddata5.Redraw = True

	End With

End Sub

'=======================================================================================================
'   Event Name : SetSpread4Color()
'   Event Desc : 관리항목 그리드 색상설정, Protect, Require
'=======================================================================================================
Sub SetSpread4Color()

Dim indx
Dim tmpDrCrFG
Dim strStartRow, strEndRow

    With frm1
		strStartRow = 1
		strEndRow	= .vspddata5.MaxRows

		ggoSpread.Source	= .vspddata5
		.vspddata5.ReDraw	= False		
		.vspddata4.Col		= C_DrCRFG
		
		If lgCurrRow = "" then
			lgCurrRow = 1
		End If
		If (lgCurrRow >= 1 And .vspddata4.Row > lgCurrRow) Or .vspddata4.Row < 1 Then
			.vspddata4.Row	= .vspddata4.ActiveRow
		End If

		tmpDrCrFG = LEFT(.vspddata4.Text,1)

		ggoSpread.SSSetProtected C_DtlSeq_2,		strStartRow,	strEndRow
		ggoSpread.SSSetProtected C_CtrlCd_2,		strStartRow,	strEndRow
		ggoSpread.SSSetProtected C_CtrlNm_2,		strStartRow,	strEndRow
		ggoSpread.SSSetProtected C_CtrlValNm_2,	    strStartRow,	strEndRow

		For indx = 1 to .vspddata5.MaxRows
			.vspddata5.Row = indx
			.vspddata5.Col = C_DRFg_2

'			msgbox "lgCurrRow=" & lgCurrRow & "tmpDrCrFG=" & tmpDrCrFG & ".vspddata5.Text=" & .vspddata5.Text
			If (.vspddata5.Text = tmpDrCrFG And .vspddata5.Text <> "") Or .vspddata5.Text = "Y" Or .vspddata5.Text = "DC" Then
				ggoSpread.SSSetRequired C_CtrlVal_2, indx, indx
			Else
				ggoSpread.SpreadUnLock  C_CtrlVal_2, indx, C_CtrlVal_2, indx
			End If
		Next
		
		.vspddata5.ReDraw = True

    End With

End Sub
'========================================================================================
' Function Name : GetCtrlSpreadColumnPos2
' Description   : 
'========================================================================================
Sub GetCtrlSpreadColumnPos2(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspddata5
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
            C_DtlSeq_2    = iCurColumnPos(1)
			C_CtrlCd_2    = iCurColumnPos(2)
			C_CtrlNm_2    = iCurColumnPos(3)
			C_CtrlVal_2   = iCurColumnPos(4)
			C_CtrlPB_2    = iCurColumnPos(5)
			C_CtrlValNm_2 = iCurColumnPos(6)
			C_Seq_2       = iCurColumnPos(7)
			C_Tableid_2   = iCurColumnPos(8)
			C_Colid_2     = iCurColumnPos(9)
			C_ColNm_2     = iCurColumnPos(10)
			C_Datatype_2  = iCurColumnPos(11)
			C_DataLen_2   = iCurColumnPos(12)
			C_DRFg_2      = iCurColumnPos(13)
			C_HItemSeq_2  = iCurColumnPos(14)
			C_MajorCd_2   = iCurColumnPos(15)

    End Select    
End Sub
'=======================================================================================================
'   Event Name : OpenCtrlPB2
'   Event Desc : 관리항목 PopUp
'=======================================================================================================
Function OpenCtrlPB2(Byval strTable, Byval strFld1 , Byval strFld2 , Byval strCode , Byval FldNm, ByVal sWhere )

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
'20080519 수정		
'	ElseIf strFld1 = "COST_CD" Then	
'	  	arrParam(0) = "관리항목 VALUE 팝업"						' 팝업 명칭 
'	  	arrParam(1) = "B_ACCT_DEPT A, B_COST_CENTER B "	    	' TABLE 명칭 
'	  	arrParam(2) = strCode						' Code Condition
'	  	arrParam(3) = ""							' Name Cindition
'	  	'arrParam(4) = sWhere						' Where Condition
'		arrParam(4) = "A.ORG_CHANGE_ID =  " & FilterVar(frm1.hOrgChangeId.value , "''", "S") & ""
'		arrParam(4) = arrParam(4) & " And A.COST_CD = B.COST_CD And B.BIZ_AREA_CD = ( Select B.BIZ_AREA_CD"
'		arrParam(4) = arrParam(4) & " FROM B_ACCT_DEPT A, B_COST_CENTER B WHERE A.DEPT_CD =  " & FilterVar(frm1.txtDeptCd.value , "''", "S") & ""
'		arrParam(4) = arrParam(4) & " And A.COST_CD = B.COST_CD And A.ORG_CHANGE_ID =  " & FilterVar(frm1.hOrgChangeId.value , "''", "S") & ")"	  	
'	  	arrParam(5) = "코스트코드"   				' 조건필드의 라벨 명칭 
'	  
'	  	arrField(0) = "A.COST_CD"    				' Field명(0)
'	  	arrField(1) = "B.COST_NM"	    		    ' Field명(1)
'	  
'	  	arrHeader(0) = "관리항목 VALUE"				' Header명(0)
'	  	arrHeader(1) = "관리항목 VALUE명"		
	ElseIf strFld1 = "COST_CD" Then	
	  	arrParam(0) = "관리항목 VALUE 팝업"						' 팝업 명칭 
	  	arrParam(1) = "B_ACCT_DEPT A, B_COST_CENTER B "	    	' TABLE 명칭 
	  	arrParam(2) = strCode						' Code Condition
	  	arrParam(3) = ""							' Name Cindition
		arrParam(4) = "A.ORG_CHANGE_ID =  " & FilterVar(frm1.hOrgChangeId.value , "''", "S") & ""
		arrParam(4) = arrParam(4) & " And A.COST_CD =* B.COST_CD And B.BIZ_UNIT_CD = ( Select B.BIZ_UNIT_CD"
		arrParam(4) = arrParam(4) & " FROM B_ACCT_DEPT A, B_COST_CENTER B WHERE A.DEPT_CD =  " & FilterVar(frm1.txtDeptCd.value , "''", "S") & ""
		arrParam(4) = arrParam(4) & " And A.DEPT_CD = B.DEPT_CD And A.ORG_CHANGE_ID =  " & FilterVar(frm1.hOrgChangeId.value , "''", "S") & ")"	  	
	  	arrParam(5) = "코스트코드"   				' 조건필드의 라벨 명칭 

	  	arrField(0) = "B.COST_CD"    				' Field명(0)
	  	arrField(1) = "B.COST_NM"	    		    ' Field명(1)
	  
	  	arrHeader(0) = "관리항목 VALUE"				' Header명(0)
	  	arrHeader(1) = "관리항목 VALUE명"			
	ElseIf strFld1 = "EMP_NO" Then	
	  	arrParam(0) = "관리항목 VALUE 팝업"						' 팝업 명칭 
	  	arrParam(1) = "HAA010T A,B_MINOR B  "	    	' TABLE 명칭 
	  	arrParam(2) = strCode						' Code Condition
	  	arrParam(3) = ""							' Name Cindition
		arrParam(4) = "A.ROLL_PSTN =  B.MINOR_CD "
		arrParam(4) = arrParam(4) & " And B.MAJOR_CD = 'H0002' "
	  	arrParam(5) = "사원코드"   				' 조건필드의 라벨 명칭 

	  	arrField(0) = "A.EMP_NO"    				' Field명(0)
	  	arrField(1) = "A.NAME"	    		    ' Field명(1)
	  	arrField(2) = "A.DEPT_NM"	    		    ' Field명(1)
	  	arrField(3) = "B.MINOR_NM"	    		    ' Field명(1)
	  
	  	arrHeader(0) = "사번"				' Header명(0)
	  	arrHeader(1) = "사원명"
	  	arrHeader(2) = "부서명"				' Header명(0)
	  	arrHeader(3) = "직급"			  				  	
	ElseIf strFld1 = "CREDIT_NO" Then	
	  	arrParam(0) = "관리항목 VALUE 팝업"			' 팝업 명칭 
	  	arrParam(1) = strTable	    				' TABLE 명칭 
	  	arrParam(2) = strCode						' Code Condition
	  	arrParam(3) = ""							' Name Cindition
	  	arrParam(4) = sWhere						' Where Condition
	  	arrParam(5) = FldNm      					' 조건필드의 라벨 명칭 
	  
	  	arrField(0) = strFld1	    				' Field명(0)
	  	arrField(1) = strFld2	    		        ' Field명(1)
	  	arrField(2) = "credit_eng_nm"	    		    ' Field명(2)
	  
	  	arrHeader(0) = "관리항목 VALUE"				' Header명(0)
	  	arrHeader(1) = "관리항목 VALUE명"
	  	arrHeader(2) = "관리자명"				' Header명(0)

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
	If strFld1 = "BP_CD" then
	
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
	
	
	Else
			arrRet = window.showModalDialog("../../comasp/adoAcctctrl_ko441_1_popup.asp", Array(arrParam, arrField, arrHeader), _
				"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	End If
	
	'####### 관리항목이 거래처일때, 거래처정보가 상세히 보이도록 수정 #######
	'200803181431
	'If strFld1 <> "BP_CD" or strFld1 <> "EMP_NO" Then
	'	arrRet = window.showModalDialog("../../comasp/adoAcctctrl_ko441_1_popup.asp", Array(arrParam, arrField, arrHeader), _
	'			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	'Else
	'	'관리항목이 거래처일때.
	'	Dim iCalledAspName
	'	Dim IntRetCD

		
	'	iCalledAspName = AskPRAspName("b1261pa1")	
	'	If Trim(iCalledAspName) = "" then
	'		IntRetCD = DisplaayMsgBox("900040",parent.VB_INFORMATION, "b1261pa1", "x")
	'		IsOpenPop = False
	'		Exit Function
	'	End if
	
	'	arrRet = window.showModalDialog(iCalledAspName,Array(window.parent, strCode),"dialogWidth=800px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	'End If
	'200803181431
	
	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetCtrlPB2(arrRet, strFld2, strFld1)
	End If

End Function

'=======================================================================================================
'   Event Name : SetCtrlPB2
'   Event Desc : 관리항목 PopUp Data Setting
'=======================================================================================================
Function SetCtrlPB2(Byval arrRet, Byval pstrFld2, Byval strFld1)
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
			.vspddata5.Row =  .vspddata5.ActiveRow
			.vspddata5.Col =  C_CtrlVal_2
			.vspddata5.Text = arrRet(0)

			If Len(Trim(pstrFld2)) > 0 Then
				.vspddata5.Col =  C_CtrlValNm_2
				.vspddata5.Text = arrRet(1)
			End If
			call vspdData5_Change( C_CtrlVal_2 ,  .vspddata5.Row)
			For lngRows = 1 To .vspddata5.MaxRows
				.vspddata5.Row = lngRows
				.vspddata5.Col = C_CtrlVal_2
 				iTempCtrlVal = Trim(.vspddata5.Text)
				.vspddata5.Col = C_Colid_2
 				iTempColid = Trim(.vspddata5.Text)
		
				IF iTempColid = "BANK_CD" and iTempCtrlVal = "" Then
					.vspddata5.Col = C_CtrlVal_2
					.vspddata5.Text = arrRet(2)
					' query bank_nm 				
					.vspddata5.Col = C_ColNm_2  
					itempColNm = Trim(.vspddata5.Text) 

					strSelect	=	  itempColNm     		
					strFrom		=	  " B_BANK "		
					strWhere	=	  " BANK_CD = " & FilterVar(arrRet(2), "''", "S") & ""

					If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) Then 	
						arrVal1 = Split(lgF2By2, Chr(11))			
						.vspddata5.Col = C_CtrlValNm_2  '은행명 
						.vspddata5.Text = Trim(arrVal1(1))
					End if 
					Call CopyToHSheet22(frm1.vspddata4.ActiveRow,.vspddata5.Row)
					Exit For
				END IF
				
			Next
		Else	
  		.vspddata5.Row =  .vspddata5.ActiveRow
  		.vspddata5.Col =  C_CtrlVal_2
  		.vspddata5.Text = arrRet(0)
  
			If Len(Trim(pstrFld2)) > 0 Then
				.vspddata5.Col =  C_CtrlValNm_2
				.vspddata5.Text = arrRet(1)
			End If
		Call vspdData5_Change( C_CtrlVal_2 ,  .vspddata5.Row)
		End If

	End With

End Function


'======================================================================================================
' Function Name : FindCtrlNM2
' Function Desc : 관리항목값 명을 찿아 setting한다.
'=======================================================================================================
Function FindCtrlNM2(ByVal Row)

    Dim iFld1
	Dim iFld2
	Dim iTable
	Dim istrCode
	Dim sWhere
	Dim IntRetCD

	If Row < 1 Then Exit Function

	'---------- Coding part -------------------------------------------------------------
	ggoSpread.Source = frm1.vspddata5

	With frm1.vspddata5

		.Row		= Row
		.Col		= C_CtrlVal_2
		istrCode	= Trim(.Text)

		.Col		= C_Tableid_2
		iTable		= .Text

		If iTable <> "" AND istrCode <> "" Then 
			.Col	= C_Colid_2
			iFld1	= .Text

			.Col	= C_ColNm_2
			iFld2	= .Text

			sWhere	= iFld1 & " =  " & FilterVar(istrCode , "''", "S") & ""

			.Col	= C_MajorCd_2

			If  .Text <> "" Then
				sWhere = sWhere & " and  Major_CD =  " & FilterVar(.Text , "''", "S") & ""
			End If

			frm1.vspddata5.Col = C_CtrlValNm_2

			If CommonQueryRs(iFld2, iTable, sWhere, lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) Then
    			arrVal = Split(lgF0, Chr(11))
				frm1.vspddata5.Text = arrVal(0)
			ELSE
				IntRetCD = DisplayMsgBox("110330", "X", "X", "X")									'필수입력 check!!
				' 관리항목값이 바르지 않습니다.
				frm1.vspddata5.Text = ""
				frm1.vspddata5.Col = C_CtrlVal_2
				frm1.vspddata5.Text = ""
				Exit Function
			END IF
		End if

	End With

End Function

'=======================================================================================================
'   Function Name : DeleteHSheet2
'   Function Desc : 입력받은 Item번호와 관계된 관리항목 Hidden 그리드 데이타 삭제 
'=======================================================================================================
Function DeleteHSheet2(ByVal strItemSeq)

	Dim boolExist
	Dim lngRows, lngRow2, lngRow3, lngCol3
	Dim StrData
	Dim strCtrlItemSeq

	DeleteHSheet2 = False
	boolExist = False

	With frm1

		Call SortHSheet2()

		'------------------------------------
		' Find First Row
		'------------------------------------
        For lngRows = 1 To .vspddata6.MaxRows
			.vspddata6.Row = lngRows
			.vspddata6.Col = 1

			If strItemSeq = .vspddata6.Text Then
				boolExist = True
				Exit For
			End If
		Next

		lngRow2 = 1
		'------------------------------------
        ' Data Delete
        '------------------------------------
        If boolExist = True Then
			While lngRows <= .vspddata6.MaxRows

				.vspddata6.Row = lngRows
				.vspddata6.Col = 1

				If strItemSeq <> .vspddata6.Text Then
					lngRows = .vspddata6.MaxRows + 1
				Else
					If frm1.vspddata5.MaxRows > 0 Then
						.vspddata5.Col = C_HItemSeq_2
						.vspddata5.Row = .vspddata5.MaxRows
						strCtrlItemSeq = .vspddata5.Text
						'msgbox "strItemSeq" & strItemSeq & " :: " & "strCtrlItemSeq=" & strCtrlItemSeq & " :: " & ".vspddata5.Row=" & .vspddata5.Row
						If strCtrlItemSeq = strItemSeq Then
							.vspddata5.Action = 5
							.vspddata5.MaxRows = .vspddata5.MaxRows - 1
						Else
							lngRow2 =  lngRow2 + 1
						End If
					End If
					
					.vspddata6.Action = 5
					.vspddata6.MaxRows = .vspddata6.MaxRows - 1
				End If

			Wend

        End If

	End With

	DeleteHSheet2 = True

End Function


'======================================================================================================
' Function Name : SortHSheet2
' Function Desc : 관리항목 Hidden Grid 정렬 
'=======================================================================================================
Function SortHSheet2()

    With frm1
    
        .vspddata6.BlockMode	= True
        .vspddata6.Col			= 0
        .vspddata6.Col2			= .vspddata6.MaxCols
        .vspddata6.Row			= 1
        .vspddata6.Row2			= .vspddata6.MaxRows
        .vspddata6.SortBy		= 0											'SS_SORT_BY_ROW

        .vspddata6.SortKey(1)	= 1
        .vspddata6.SortKey(2)	= 2

        .vspddata6.SortKeyOrder(1) = 1										'SS_SORT_ORDER_ASCENDING
        .vspddata6.SortKeyOrder(2) = 1										'SS_SORT_ORDER_ASCENDING

        .vspddata6.Col			= 0
        .vspddata6.Col2			= .vspddata6.MaxCols
        .vspddata6.Row			= 0
        .vspddata6.Row2			= .vspddata6.MaxRows
        .vspddata6.Action		= 25										'SS_ACTION_SORT
        .vspddata6.BlockMode	= False
        
    End With

End Function

'=======================================================================================================
' Function Name : ShowHidden2
' Function Desc : 관리항목 Hidden Grid를 표시 
'=======================================================================================================
Sub ShowHidden2()

	Dim strHidden
	Dim lngRows
	Dim lngCols

    With frm1.vspddata6

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
'   Function Name : CopyToHSheet21
'   Function Desc : 관리항목그리드의 Value변경시 Hidden Grid에 Data 반영, Item 그리드에 변경여부 표시 
'=======================================================================================================
Sub CopyToHSheet21(ByVal Row)

	Dim lRow
	Dim iCols

	With frm1

	    lRow = FindData21

	    If lRow > 0 Then
	    
            .vspddata6.Row = lRow
            .vspddata5.Row = Row
            .vspddata6.Col = 0
            .vspddata5.Col = 0
            .vspddata6.Text = .vspddata5.Text
            
            .vspddata5.Col = C_DtlSeq_2
            .vspddata6.Col = 2
            .vspddata6.Text = .vspddata5.Text
                
            .vspddata5.Col = C_CtrlCd_2
            .vspddata6.Col = .vspddata6.Col + 1
            .vspddata6.Text = .vspddata5.Text
                
            .vspddata5.Col = C_CtrlNm_2
            .vspddata6.Col = .vspddata6.Col + 1
            .vspddata6.Text = .vspddata5.Text
                
            .vspddata5.Col = C_CtrlVal_2
            .vspddata6.Col = .vspddata6.Col + 1
            .vspddata6.Text = .vspddata5.Text
                
            .vspddata5.Col = C_CtrlPB_2
            .vspddata6.Col = .vspddata6.Col + 1
            .vspddata6.Text = .vspddata5.Text
                
            .vspddata5.Col = C_CtrlValNm_2
            .vspddata6.Col = .vspddata6.Col + 1
            .vspddata6.Text = .vspddata5.Text
                
            .vspddata5.Col = C_Seq_2
            .vspddata6.Col = .vspddata6.Col + 1
            .vspddata6.Text = .vspddata5.Text
                
            .vspddata5.Col = C_Tableid_2
            .vspddata6.Col = .vspddata6.Col + 1
            .vspddata6.Text = .vspddata5.Text
                
            .vspddata5.Col = C_Colid_2
            .vspddata6.Col = .vspddata6.Col + 1
            .vspddata6.Text = .vspddata5.Text
                
            .vspddata5.Col = C_ColNm_2
            .vspddata6.Col = .vspddata6.Col + 1
            .vspddata6.Text = .vspddata5.Text
                
            .vspddata5.Col = C_Datatype_2
            .vspddata6.Col = .vspddata6.Col + 1
            .vspddata6.Text = .vspddata5.Text
                
            .vspddata5.Col = C_DataLen_2
            .vspddata6.Col = .vspddata6.Col + 1
            .vspddata6.Text = .vspddata5.Text
                
            .vspddata5.Col = C_DRFg_2
            .vspddata6.Col = .vspddata6.Col + 1
            .vspddata6.Text = .vspddata5.Text

            .vspddata5.Col = C_MajorCd_2
            .vspddata6.Col = .vspddata6.Col + 1
            .vspddata6.Text = .vspddata5.Text
            
            .vspddata5.Col = C_MajorCd_2 + 1 
            .vspddata6.Col = .vspddata6.Col + 1
            .vspddata6.Text = .vspddata5.Text
            
        End If

	End With

	frm1.vspddata4.Row = frm1.vspddata4.ActiveRow
	frm1.vspddata4.Col = 0

	If frm1.vspddata4.Text <> ggoSpread.InsertFlag And frm1.vspddata4.Text <> ggoSpread.DeleteFlag Then
   	    frm1.vspddata4.Text = ggoSpread.UpdateFlag
	End if

End Sub

'=======================================================================================================
'   Function Name : FindData21
'   Function Desc : 현재의 Item, Dtl에 해당하는 Hidden Grid의 Index를 Return
'=======================================================================================================
Function FindData21()

	Dim strApNo
	Dim strItemSeq
	Dim strDtlSeq
	Dim lRows

    FindData21 = 0

    With frm1

        For lRows = 1 To .vspddata6.MaxRows

			.vspddata6.Row	= lRows
			.vspddata6.Col	= 1
            strItemSeq		= .vspddata6.Text
            .vspddata6.Col	= 2
            strDtlSeq		= .vspddata6.Text

            .vspddata4.Row	= frm1.vspddata4.ActiveRow
            .vspddata5.Row	= frm1.vspddata5.ActiveRow

            .vspddata4.Col	= C_ItemSeq
            
            If strItemSeq = .vspddata4.Text Then
                .vspddata5.Col = C_DtlSeq_2
                If strDtlSeq = .vspddata5.Text Then
                    FindData21 = lRows
                    Exit Function
                End If
            End If

        Next

    End With

End Function


'=======================================================================================================
'   Function Name : CopyFromData2
'   Function Desc : 관리항목 Hidden 그리드에서 입력받은 Item번호에 
'                   해당하는 관리항목 값을 표시, 해당 관리항목이 없으면 False 값 Return
'=======================================================================================================
Function CopyFromData2(ByVal strItemSeq)

	Dim lngRows , indx, indx1
	Dim boolExist
	Dim iCols
	Dim tmpDrCrFG
	Dim iStrData, iStrFlag
	Dim arrFlag
	Dim strHItemSeq

    boolExist = False

	ggoSpread.Source = frm1.vspddata5
	Call ggoSpread.ClearSpreadData()
	
    CopyFromData2			= boolExist

    With frm1

        Call SortHSheet2()

      '------------------------------------
      ' Find First Row
      '------------------------------------
        For lngRows = 1 To .vspddata6.MaxRows
            .vspddata6.Row = lngRows
            .vspddata6.Col = 1

            If strItemSeq = .vspddata6.Text Then
                boolExist = True
                Exit For
            End If
        Next

      '------------------------------------
      ' Show Data
      '------------------------------------
		.vspddata6.Row = lngRows
		
        If boolExist = True Then

			ggoSpread.Source = .vspddata5
			Call ggoSpread.ClearSpreadData()
            .vspddata5.Redraw = False

            For indx = lngRows to .vspddata6.MaxRows

                .vspddata6.Row = indx
                .vspddata6.Col = 1

                If strItemSeq = .vspddata6.Text Then
					
                    .vspddata6.Col= 0
                    iStrFlag = iStrFlag & .vspddata6.Text & Chr(12)
                    
                    .vspddata6.Col = 1
                    strHItemSeq = .vspddata6.Text
                    
                    .vspddata6.Col = 2
                    iStrData = iStrData & Chr(11) & .vspddata6.Text
                    .vspddata6.Col = .vspddata6.Col + 1
                    iStrData = iStrData & Chr(11) & .vspddata6.Text 
                    .vspddata6.Col = .vspddata6.Col + 1
                    iStrData = iStrData & Chr(11) & .vspddata6.Text
                    .vspddata6.Col = .vspddata6.Col + 1
                    iStrData = iStrData & Chr(11) & .vspddata6.Text
                    .vspddata6.Col = .vspddata6.Col + 1
                    iStrData = iStrData & Chr(11) & .vspddata6.Text
                    .vspddata6.Col = .vspddata6.Col + 1
                    iStrData = iStrData & Chr(11) & .vspddata6.Text
                    .vspddata6.Col = .vspddata6.Col + 1
                    iStrData = iStrData & Chr(11) & .vspddata6.Text
                    .vspddata6.Col = .vspddata6.Col + 1
                    iStrData = iStrData & Chr(11) & .vspddata6.Text
                    .vspddata6.Col = .vspddata6.Col + 1
                    iStrData = iStrData & Chr(11) & .vspddata6.Text
                    .vspddata6.Col = .vspddata6.Col + 1
                    iStrData = iStrData & Chr(11) & .vspddata6.Text
                    .vspddata6.Col = .vspddata6.Col + 1
                    iStrData = iStrData & Chr(11) & .vspddata6.Text
                    .vspddata6.Col = .vspddata6.Col + 1
                    iStrData = iStrData & Chr(11) & .vspddata6.Text
                    .vspddata6.Col = .vspddata6.Col + 1
                    iStrData = iStrData & Chr(11) & .vspddata6.Text
                    iStrData = iStrData & Chr(11) & strHItemSeq
                    .vspddata6.Col = .vspddata6.Col + 1
                    iStrData = iStrData & Chr(11) & .vspddata6.Text
                    .vspddata6.Col = .vspddata6.Col + 1
                    iStrData = iStrData & Chr(11) & .vspddata6.Text
                    iStrData = iStrData & Chr(11) & Chr(12)

                End If

            Next

			ggoSpread.SSShowData iStrData 
			
			If iStrFlag <> "" Then
				arrFlag = Split(iStrFlag, Chr(12))			
				For indx1 = 0 to Ubound(arrFlag) - 1

				    .vspddata5.Row = indx1 + 1
					.vspddata5.Col = 0
				    .vspddata5.Text = arrFlag(indx1)

				Next
			End If

            frm1.vspddata5.Redraw = True

        End If

    End With

    CopyFromData2 = boolExist

End Function

'=======================================================================================================
' Function Name : CheckSpread4
' Function Desc : 저장시에  관리항목 필수여부 check 하기위해 호출되는 Function
'=======================================================================================================

Function CheckSpread6()

	Dim indx
	Dim tmpDrCrFG

	CheckSpread6 = False

	With frm1
	
	 	For indx = 1 to .vspddata6.MaxRows
		    .vspddata6.Row = indx
		    .vspddata6.Col = 14
		    If .vspddata6.Text = "Y" Or .vspddata6.Text = "DC" Then
  			  .vspddata6.Col = 5
			  If Trim(.vspddata6.Text) = "" Then
				Exit Function
		  	  End If
		    End If
		Next

        End With

	CheckSpread6 = True

End Function

'=======================================================================================================
' Function Name : DbQuery4
' Function Desc : Item 그리드 변경시 관리항목 조회 
'=======================================================================================================
Function DbQuery4(ByVal Row)

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
	
	on Error Resume Next
	Err.Clear
	
	DbQuery4 = False

	Call DisableToolBar(parent.TBC_QUERY)	
	Call LayerShowHide(1)
	
	With frm1
       If CopyFromData2(.hItemSeq.Value) = True Then
			Call LayerShowHide(0)
			Call RestoreToolBar()
			Call SetSpread2Color2()
			Exit Function
	   End If

		ggoSpread.Source = frm1.vspddata5
		Call ggoSpread.ClearSpreadData()
       .vspddata4.Row = Row
       .vspddata4.Col = C_ItemSeq
       ICurItemSeq	 = .vspddata4.Text
       .vspddata4.Col = C_AcctCd
	End With

	If CommonQueryRs("ACCT_NM", " A_ACCT (NOLOCK)" , "ACCT_CD =  " & FilterVar(Frm1.vspddata4.Text , "''", "S") & "", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) Then
    	frm1.vspddata4.Col	= C_AcctNm
    	arrVal				= Split(lgF0, Chr(11))
		frm1.vspddata4.Text	= arrVal(0)
	Else
		frm1.vspddata4.Text	= ""
		frm1.vspddata4.Col	= C_AcctNm
		frm1.vspddata4.Text	= ""
		IntRetCD			= DisplayMsgBox("110100", "X", "X", "X")
		Call LayerShowHide(0)
		Call RestoreToolBar()
		Exit Function
	End If

	frm1.vspddata4.Col = C_AcctCd

    strSelect =	            " B.CTRL_ITEM_SEQ,  A.CTRL_CD, A.CTRL_NM , '', '',"
    strSelect = strSelect & " CASE  WHEN A.COLM_DATA_TYPE = " & FilterVar("D", "''", "S") & "  THEN " & FilterVar("(Format : YYYYMMDD)", "''", "S") & "  END , " & ICurItemSeq  & ", LTrim(ISNULL(A.TBL_ID,'')),LTrim(ISNULL(A.DATA_COLM_ID,'')), "
    strSelect = strSelect & " LTrim(ISNULL(A.DATA_COLM_NM,'')),  LTrim(ISNULL(A.COLM_DATA_TYPE,'')), LTrim(ISNULL(A.DATA_LEN,'')), "
    strSelect = strSelect & " CASE WHEN B.DR_FG = " & FilterVar("Y", "''", "S") & "  AND  B.CR_FG = " & FilterVar("Y", "''", "S") & "  THEN " & FilterVar("DC", "''", "S") & "  "
    strSelect = strSelect & " WHEN B.DR_FG = " & FilterVar("Y", "''", "S") & "  AND  B.CR_FG = " & FilterVar("N", "''", "S") & "  THEN " & FilterVar("D", "''", "S") & "  "
    strSelect = strSelect & " WHEN B.DR_FG = " & FilterVar("N", "''", "S") & "  AND  B.CR_FG = " & FilterVar("Y", "''", "S") & "  THEN " & FilterVar("C", "''", "S") & "  "
    strSelect = strSelect & " END	, " & ICurItemSeq & " , "
    strSelect = strSelect & " LTrim(ISNULL(A.MAJOR_CD, '')), CHAR(8) "

    strSelect1 =ICurItemSeq & " ,B.CTRL_ITEM_SEQ,  A.CTRL_CD, A.CTRL_NM , '', '',"
    strSelect1 = strSelect1 & " CASE  WHEN A.COLM_DATA_TYPE = " & FilterVar("D", "''", "S") & "  THEN " & FilterVar("(Format : YYYYMMDD)", "''", "S") & "  END , " & ICurItemSeq  & ", LTrim(ISNULL(A.TBL_ID,'')),LTrim(ISNULL(A.DATA_COLM_ID,'')), "
    strSelect1 = strSelect1 & " LTrim(ISNULL(A.DATA_COLM_NM,'')),  LTrim(ISNULL(A.COLM_DATA_TYPE,'')), LTrim(ISNULL(A.DATA_LEN,'')), "
    strSelect1 = strSelect1 & " CASE WHEN B.DR_FG = " & FilterVar("Y", "''", "S") & "  AND  B.CR_FG = " & FilterVar("Y", "''", "S") & "  THEN " & FilterVar("DC", "''", "S") & "  "
    strSelect1 = strSelect1 & " WHEN B.DR_FG = " & FilterVar("Y", "''", "S") & "  AND  B.CR_FG = " & FilterVar("N", "''", "S") & "  THEN " & FilterVar("D", "''", "S") & "  "
    strSelect1 = strSelect1 & " WHEN B.DR_FG = " & FilterVar("N", "''", "S") & "  AND  B.CR_FG = " & FilterVar("Y", "''", "S") & "  THEN " & FilterVar("C", "''", "S") & "  "
    strSelect1 = strSelect1 & " END	, "
    strSelect1 = strSelect1 & " LTrim(ISNULL(A.MAJOR_CD, '')), CHAR(8) "

	strWhere =  "A.CTRL_CD = B.CTRL_CD  AND B.ACCT_CD =  " & FilterVar(Frm1.vspddata4.Text, "''", "S") & ""
	strWhere =  strWhere & " Order By B.CTRL_ITEM_SEQ "

	frm1.vspddata5.ReDraw = False
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	If CommonQueryRs2by2(strSelect, " A_CTRL_ITEM   A (NOLOCK), A_ACCT_CTRL_ASSN B (NOLOCK)" , strWhere , lgF2By2) Then
		ggoSpread.Source = frm1.vspddata5
		arrTemp =  Split(lgF2By2,Chr(12))
		For Indx = 0 To Ubound(arrTemp) - 1
			arrTemp(indx) = Replace(arrTemp(indx), Chr(8), indx + 1)
		Next
		lgF2By2 = Join(arrTemp,Chr(12))
		ggoSpread.SSShowData lgF2By2

		For lngRows = 1 to frm1.vspddata5.MaxRows
			frm1.vspddata5.Row	= lngRows
			frm1.vspddata5.Col	= 0
			frm1.vspddata5.Text	= ggoSpread.InsertFlag
		Next

		If CommonQueryRs2by2(strSelect1, " A_CTRL_ITEM  A (NOLOCK), A_ACCT_CTRL_ASSN B (NOLOCK)" , strWhere , lgF2By2) Then
			ggoSpread.Source = frm1.vspddata6
			IDtlRow = frm1.vspddata6.MaxRows
			arrTemp1 =  Split(lgF2By2,Chr(12))
			For Indx1 = 0 To Ubound(arrTemp1) - 1
				arrTemp1(indx1) = Replace(arrTemp1(indx1), Chr(8), indx1 + 1)
			Next
			lgF2By2 = Join(arrTemp1,Chr(12))

			ggoSpread.SSShowData lgF2By2
			For lngRows = IDtlRow + 1 To frm1.vspddata6.MaxRows
				frm1.vspddata6.Row	= lngRows
				frm1.vspddata6.Col	= 0
				frm1.vspddata6.Text	= ggoSpread.InsertFlag
			Next
		End If

		Call SetSpread2Color2()

    End If

    frm1.vspddata5.ReDraw = True

    Call LayerShowHide(0)
    Call RestoreToolBar()

	If Err.number = 0 Then
		DbQuery4 = True
	End If
	
	Set gActiveElement = document.ActiveElement

End Function

'=======================================================================================================
' Function Name : DbQueryOk4
' Function Desc : DbQuery4가 성공적일 경우 MyBizASP 에서 호출되는 Function
'=======================================================================================================
Sub DbQueryOk4()

	Call SetSpread2Color2()

End Sub

'=======================================================================================================
'   Event Name : vspdData5_ButtonClicked
'   Event Desc : 관리항목 팝업버튼 클릭시 관리항목 팝업호출 
'=======================================================================================================
Sub vspdData5_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)

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

	'---------- Coding part -------------------------------------------------------------
	ggoSpread.Source = frm1.vspddata5

	With frm1.vspddata5
		If Row > 0 And Col = C_CtrlPB_2 Then
			.Row = Row
			.Col = C_CtrlNm_2
			FldNm = .Text

			.Col = C_CtrlVal_2
			istrCode = .Text 

			.Col = C_Tableid_2
			iTable = Trim(.Text)

			.Col = C_Colid_2
			iFld1 = Trim(.Text)
			
			.Col = C_ColNm_2
			iFld2 = Trim(.Text)

			.Col = C_MajorCd_2

			IF  .Text <> "" Then
				strWhere = " Major_CD =  " & FilterVar(.Text , "''", "S") & ""
			ElseIF iTable = "B_ACCT_DEPT" Then
				'If Trim(frm1.hOrgChangeId.value) <> "" Then
				'	strWhere = " Org_Change_Id = '" & Trim(frm1.hOrgChangeId.value) & "'"
				'else
					strWhere = " Org_Change_Id =  " & FilterVar(gChangeOrgId , "''", "S") & ""
				'end if	
			ELse
				strWhere = ""	
			END IF	
			IF iFld1 = "BANK_ACCT_NO" Then
				For lngRows = 1 To .MaxRows
					.Row = lngRows
					.Col = C_Tableid_2 
					iTempTable = Trim(.Text)
					.Col = C_CtrlVal_2
 					iTempCtrlVal = Trim(.Text)
					IF iTempTable = "B_BANK" and iTempCtrlVal <> "" Then
						strWhere = " BANK_CD LIKE  " & FilterVar(iTempCtrlVal, "''", "S") & ""
						Exit For
					END IF
				Next
			END IF						
			If iTable <> "" Then 
 				Call OpenCtrlPB2(iTable, iFld1, iFld2, istrCode, FldNm, strWhere)
			End if

'			.Row = Row
			frm1.vspddata5.Col = C_CtrlCd_2
			If Trim(frm1.vspddata5.Text) = "V4" Then

				frm1.vspddata5.Col = C_CtrlVal_2
				strVatType = Trim(frm1.vspddata5.text)
				If Trim(strVatType) <> "" Then			
				
					strSelect	= "reference"
					strFrom		= "b_configuration"
					strWhere	= "major_cd = " & FilterVar("B9001", "''", "S") & "  and seq_no = 1 and minor_cd =  " & FilterVar(strVatType , "''", "S") & ""
					
					If CommonQueryRs(strSelect, strFrom, strWhere, lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) Then
						arrTemp = Split(lgF0, chr(11))
						strVatRate = arrTemp(0)				
					End If
					
					frm1.vspddata5.Col = C_CtrlCd_2				
					For ii = i To frm1.vspddata5.MaxRows
						frm1.vspddata5.Row = ii
						If Trim(frm1.vspddata5.Text) = "V7" Then
							frm1.vspddata5.Col = C_CtrlVal_2
							frm1.vspddata5.Text = strVatRate
						End If
					Next			
				End If
			End If 

		End If
	End With
	
	'매출부가세에서 세금계산서를 setting하면 자동으로 vat_rate가 hidden으로 복사되도록 한다.
	With frm1
	.vspddata5.Col = C_CtrlCd_2        
        	For lRows = 1 To .vspddata5.MaxRows        
            	.vspddata5.Row = lRows								'ActiveRow설정 
				If Trim(.vspddata5.Text) = "V7" Then
					.vspddata5.action = 0
					CopyToHSheet21 lRows
				End If      
        	Next
	End With

End Sub

'=======================================================================================================
'   Event Name : vspdData5_Change
'   Event Desc : 관리항목 그리드 데이타 변경시 입력값에대한 유효성 Check
'=======================================================================================================
Sub vspdData5_Change(ByVal Col, ByVal Row)

	Dim iLen
	Dim sPreCtrlVal
	Dim IntRetCD

   	ggoSpread.Source = frm1.vspddata5
	ggoSpread.UpdateRow Row

	frm1.vspddata5.Row = Row
	frm1.vspddata5.Col = 0

	Select Case Col
		Case   C_CtrlVal_2
	    '----------------------------------
		' 입력된 관리항목의 DataType Check yyyy-mm-dd
		'----------------------------------
		    frm1.vspddata5.Col = C_Datatype_2

	        If Trim(frm1.vspddata5.Text) = "D" Then
				frm1.vspddata5.Col = C_CtrlVal_2
				sPreCtrlVal = frm1.vspddata5.Text
				
				if LEN(frm1.vspdData5.Text) = 8 then
					frm1.vspdData5.Text = Mid(frm1.vspdData5.Text,1,4) & "-" & Mid(frm1.vspdData5.Text,5,2) + "-" + Mid(frm1.vspdData5.Text,7,2) 
				End if 
				
				If IsDate(frm1.vspddata5.Text) = False or IsNumeric(Mid(frm1.vspddata5.Text,1,4)) = False or _
					IsNumeric(Mid(frm1.vspddata5.Text,6,2)) = False or _
					IsNumeric(Mid(frm1.vspddata5.Text,9,2)) = False or _
					Mid(frm1.vspddata5.Text,5,1) <> "-" or _
					Mid(frm1.vspddata5.Text,8,1) <> "-" or _
					Mid(frm1.vspddata5.Text,1,4) < "1900" Then
						frm1.vspddata5.Text = sPreCtrlVal
						IntRetCD = DisplayMsgBox("174223", "X", "X", "X")							'필수입력 check!!
						' 입력하신 날짜는 부적합합니다.
						frm1.vspddata5.Text = ""
						Exit Sub
				End If
			ElseIf Trim(frm1.vspddata5.Text) = "N" Then
				frm1.vspddata5.Col = C_CtrlVal_2
				sPreCtrlVal = frm1.vspddata5.Text
				If IsNumeric(frm1.vspddata5.Text) = False Then
					frm1.vspddata5.Text = sPreCtrlVal
					IntRetCD = DisplayMsgBox("229924", "X", "X", "X")								'필수입력 check!!
					' 숫자를 입력하십시오 
					frm1.vspddata5.Text = ""
					Exit Sub
				Else
					frm1.vspddata5.Text = replace(formatnumber(frm1.vspddata5.Text,2), parent.gComNumDec & "00", "")
				End If
	        End If

	        '------------------------------------
	        ' 입력된 관리항목의 길이Check
	        '------------------------------------
	        frm1.vspddata5.Col = C_CtrlVal_2

	        iLen = Len(frm1.vspddata5.Text)
			sPreCtrlVal = frm1.vspddata5.Text
	        frm1.vspddata5.Col = C_DataLen_2

	        If iLen > Int(frm1.vspddata5.Text) Then
				frm1.vspddata5.Text = sPreCtrlVal
				IntRetCD = DisplayMsgBox("110320", "X", "X", "X")									'필수입력 check!!
			'  관리항목값의 길이를 확인하십시오.
				frm1.vspddata5.Col = C_CtrlVal_2
				frm1.vspddata5.Text = ""
				Exit Sub
	        End If

	        frm1.vspddata5.Col = C_Datatype_2

	        If Trim(frm1.vspddata5.Text) <> "D" And Trim(frm1.vspddata5.Text) <> "N" Then
				FindCtrlNM2   Row																			'관리항목값을 check하고 관리항목명을 찾아준다.
			End If
    End Select

'	CopyToHSheet21 Row
	CopyToHSheet22 frm1.vspdData4.ActiveRow,Row

    lgBlnFlgChgValue = True

End Sub

'==========================================================================================
'   Event Name : vspdData5_Click
'   Event Desc : This event is spread sheet data changed
'==========================================================================================
Sub vspdData5_Click(ByVal Col, ByVal Row)

	Call SetPopupMenuItemInf("0000111111")

	gMouseClickStatus = "SP2C"	'Split 상태코드 
	   
	Set gActiveSpdSheet = frm1.vspddata5

	If Row <= 0 Then                                                    'If there is no data.
		ggoSpread.Source = frm1.vspddata5
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
'   Event Name : vspdData5_ColWidthChange
'   Event Desc :
'========================================================================================================
Sub vspdData5_ColWidthChange(ByVal pvCol1, ByVal pvCol2)

    ggoSpread.Source = frm1.vspddata5
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

End Sub
'========================================================================================================
'   Event Name : vspdData5_DblClick
'   Event Desc : 
'========================================================================================================
Sub vspdData5_DblClick(ByVal Col, ByVal Row)

    Dim iColumnName

    If Row <= 0 Then
		Exit Sub
    End If
    If frm1.vspddata5.MaxRows = 0 Then
		Exit Sub
	End If
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End   ) --------------------------------------------------------------

End Sub
'==========================================================================================
'   Event Desc : Spread Split 상태코드 
'==========================================================================================
Sub vspdData5_MouseDown(Button, Shift, X, Y)
	If Button = 2 And gMouseClickStatus = "SP2C" Then
		gMouseClickStatus = "SP2CR"
	End If
End Sub
'========================================================================================================
'   Event Name : vspdData5_GotFocus
'   Event Desc : This event is spread sheet data changed
'========================================================================================================
Sub vspdData5_GotFocus()
    ggoSpread.Source = Frm1.vspddata5
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End   ) --------------------------------------------------------------
End Sub
'=======================================================================================================
'   Event Name : vspdData5_onfocus
'   Event Desc :
'=======================================================================================================
Sub vspdData5_onfocus()

'    If lgIntFlgMode <> parent.OPMD_UMODE Then
'        Call SetToolbar("1110100000011111")                                     '버튼 툴바 제어 
'    Else
'        Call SetToolbar("1111100000011111")                                     '버튼 툴바 제어 
'    End If

End Sub
'========================================================================================================
'   Event Name : vspdData5_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData5_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )

    ggoSpread.Source = frm1.vspddata5
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )
    Call GetCtrlSpreadColumnPos2("A")

End Sub

'=======================================================================================================
'   Event Name : SetGridFocus5
'   Event Desc :
'=======================================================================================================
Sub SetGridFocus5()	

	With frm1
		.vspddata5.Row		= 1
		.vspddata5.Col		= C_DtlSeq_2
		.vspddata5.Action	= 1
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
		lgCashAcct_2	= arrVal(0)
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
'	Name : SetAuthorityFlag2
'	Description :
'==========================================================================================
Sub SetAuthorityFlag2()
	If CommonQueryRs("TOP 1 USR_ID", "Z_USR_AUTHORITY_VALUE", "USR_ID =  " & FilterVar(gUsrId , "''", "S") & " AND MODULE_CD = " & FilterVar("A", "''", "S") & "  ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) Then

	    If UCase(gUsrId) = UCase(Replace(lgF0,Chr(11),"")) Then
		
	      lgAuthorityFlag_2 = "Y"
	  	Else
	    	lgAuthorityFlag_2 = "N"
	    End If
	Else
	  	lgAuthorityFlag_2 = "N"
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

			If lgAuthorityFlag_2 = "Y" Then		                                    '권한관리 추가 
				arrstrRet(0) = "B_ACCT_DEPT A, Z_USR_AUTHORITY_VALUE B "     		'권한관리 추가 
			Else																	'권한관리 추가 
				arrstrRet(0) = "B_ACCT_DEPT A "    									' TABLE 명칭 
			End If

			If lgAuthorityFlag_2 = "Y" Then		                                    '권한관리 추가 
				arrstrRet(1) = "ORG_CHANGE_ID =  " & FilterVar(strCode1, "''", "S") & " AND A.DEPT_CD = B.CODE_VALUE AND B.USR_ID =  " & FilterVar(gUsrId , "''", "S") & " AND B.MODULE_CD = " & FilterVar("A", "''", "S") & " "		 '권한관리 추가 
			Else                                                              '권한관리 추가 
				arrstrRet(1) = "ORG_CHANGE_ID =  " & FilterVar(strCode1, "''", "S") & ""
			End If

		Case "DEPT_ITEM"

			If lgAuthorityFlag_2 = "Y" Then																		'권한관리 추가 
				arrstrRet(0) = " B_ACCT_DEPT A, B_COST_CENTER B, B_BIZ_AREA C, Z_USR_AUTHORITY_VALUE D "		'권한관리 추가 
			Else																								'권한관리 추가 
				arrstrRet(0) = " B_ACCT_DEPT A, B_COST_CENTER B, B_BIZ_AREA C "									' TABLE 명칭 
			End If

			If lgAuthorityFlag_2 = "Y" Then																		'권한관리 추가 
  				arrstrRet(1) = " B.COST_CD = A.COST_CD AND B.BIZ_AREA_CD = C.BIZ_AREA_CD " & _
  							   " AND C.BIZ_AREA_CD = (SELECT F.BIZ_AREA_CD" & _
  												   " FROM B_ACCT_DEPT D, B_COST_CENTER E, B_BIZ_AREA F" & _
  												   " WHERE D.DEPT_CD =  " & FilterVar(strCode2, "''", "S") & "" & _
  												   " AND D.ORG_CHANGE_ID =  " & FilterVar(strCode1, "''", "S") & "" & _
  												   " AND E.COST_CD = D.COST_CD AND E.BIZ_AREA_CD = F.BIZ_AREA_CD) AND A.DEPT_CD = D.CODE_VALUE AND D.USR_ID =  " & FilterVar(gUsrId , "''", "S") & " AND D.MODULE_CD = " & FilterVar("A", "''", "S") & " " & _
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
Function AutoInputDetail2(ByVal strAcctCd, ByVal strDeptCd, ByVal strDate, ByVal Row)

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
		strWhere	= strWhere & "			FROM	B_ACCT_DEPT A, B_COST_CENTER B, B_BIZ_AREA C, B_BIZ_AREA D "
		strWhere	= strWhere & "			WHERE	A.DEPT_CD =  " & FilterVar(strDeptCd, "''", "S") & " "
		strWhere	= strWhere & "			AND		A.Org_change_id =  " & FilterVar(strOrgChangeID, "''", "S") & " "
		strWhere	= strWhere & "			AND A.COST_CD = B.COST_CD "
		strWhere	= strWhere & "			AND B.BIZ_AREA_CD = C.BIZ_AREA_CD "
		strWhere	= strWhere & "			AND C.REPORT_BIZ_AREA_CD = D.BIZ_AREA_CD "
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
			frm1.vspddata4.row = row
			If C_VATTYPE = "" Then
				strVatType		= ""
				strVatTypeNm	= ""
			Else
				frm1.vspddata4.Col	= C_VATTYPE
				strVatType			= Trim(frm1.vspddata4.Text)
				frm1.vspddata4.Col	= C_VATNM
				strVatTypeNm		= Trim(frm1.vspddata4.Text)
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

		lngRow = frm1.vspddata5.MaxRows
		For indx = 1 to lngRow
			frm1.vspddata5.Col	= C_CtrlCd_2
			frm1.vspddata5.Row	= indx
			strCtrlCd			= UCase(Trim(frm1.vspddata5.Text))

			frm1.vspddata5.Col = C_CtrlVal_2

			Select Case strCtrlCd

				'Case "V1"
				'	If Trim(frm1.vspddata5.Text) = "" Then
				'		frm1.vspddata5.Text		= ""
				'	End If

				Case "V2"
					If Trim(frm1.vspddata5.Text) = "" Then
						frm1.vspddata5.Text		= strDate
					End If

				Case "V3"
					If Trim(frm1.vspddata5.Text) = "" Then
						frm1.vspddata5.Text		= strIoFg
						frm1.vspddata5.Col		= C_CtrlValNm_2
						frm1.vspddata5.Text		= strIoFgNm
					End If

				Case "V4"
					frm1.vspddata5.Text		= strVatType
					frm1.vspddata5.Col		= C_CtrlValNm_2
					frm1.vspddata5.Text		= strVatTypeNm

				Case "V5"
					If Trim(frm1.vspddata5.Text) = "" Then
						frm1.vspddata5.Text		= strBizAreaCd
						frm1.vspddata5.Col		= C_CtrlValNm_2
						frm1.vspddata5.Text		= strBizAreaNm
					End If

				'Case "V6"
				'	If Trim(frm1.vspddata5.Text) = "" Then
				'		frm1.vspddata5.Text		= ""
				'	End If

				Case "V7"
					frm1.vspddata5.Text		= strVatRate

			End Select

		Next
	End If

End Function

'=======================================================================================================
'   Function Name : CopyToHSheet22
'   Function Desc : 관리항목그리드의 Value를 자동settgin할때 Hidden Grid로 복사하기(CopyToHSheet21의 ActiveRow의 맹점보완)
'   최초작성자      : 박심서 
'=======================================================================================================
Sub CopyToHSheet22(ByVal MasterRow, ByVal DetailRow)

	Dim lRow
	Dim iCols

	With frm1

	    lRow = FindData22(MasterRow,DetailRow)
	    If lRow > 0 Then

            .vspddata6.Row = lRow
            .vspddata5.Row = DetailRow
            .vspddata6.Col = 0
            .vspddata5.Col = 0
            .vspddata6.Text = .vspddata5.Text

			.vspddata5.Col = C_DtlSeq_2
			.vspddata6.Col = 2
			.vspddata6.Text = .vspddata5.Text
					    
			.vspddata5.Col = C_CtrlCd_2
			.vspddata6.Col = .vspddata6.Col + 1
			.vspddata6.Text = .vspddata5.Text
					    
			.vspddata5.Col = C_CtrlNm_2
			.vspddata6.Col = .vspddata6.Col + 1
			.vspddata6.Text = .vspddata5.Text
					    
			.vspddata5.Col = C_CtrlVal_2
			.vspddata6.Col = .vspddata6.Col + 1
			.vspddata6.Text = .vspddata5.Text
					    
			.vspddata5.Col = C_CtrlPB_2
			.vspddata6.Col = .vspddata6.Col + 1
			.vspddata6.Text = .vspddata5.Text
					    
			.vspddata5.Col = C_CtrlValNm_2
			.vspddata6.Col = .vspddata6.Col + 1
			.vspddata6.Text = .vspddata5.Text
					    
			.vspddata5.Col = C_Seq_2
			.vspddata6.Col = .vspddata6.Col + 1
			.vspddata6.Text = .vspddata5.Text
					    
			.vspddata5.Col = C_Tableid_2
			.vspddata6.Col = .vspddata6.Col + 1
			.vspddata6.Text = .vspddata5.Text
					    
			.vspddata5.Col = C_Colid_2
			.vspddata6.Col = .vspddata6.Col + 1
			.vspddata6.Text = .vspddata5.Text
					    
			.vspddata5.Col = C_ColNm_2
			.vspddata6.Col = .vspddata6.Col + 1
			.vspddata6.Text = .vspddata5.Text
					    
			.vspddata5.Col = C_Datatype_2
			.vspddata6.Col = .vspddata6.Col + 1
			.vspddata6.Text = .vspddata5.Text
					    
			.vspddata5.Col = C_DataLen_2
			.vspddata6.Col = .vspddata6.Col + 1
			.vspddata6.Text = .vspddata5.Text
					    
			.vspddata5.Col = C_DRFg_2
			.vspddata6.Col = .vspddata6.Col + 1
			.vspddata6.Text = .vspddata5.Text
					    
			.vspddata5.Col = C_MajorCd_2
			.vspddata6.Col = .vspddata6.Col + 1
			.vspddata6.Text = .vspddata5.Text
			
			.vspddata5.Col = C_MajorCd_2 + 1
			.vspddata6.Col = .vspddata6.Col + 1
			.vspddata6.Text = .vspddata5.Text

        End If

	End With

	frm1.vspddata4.Row = MasterRow												'frm1.vspddata4.ActiveRow
	frm1.vspddata4.Col = 0

	If frm1.vspddata4.Text <> ggoSpread.InsertFlag and frm1.vspddata4.Text <> ggoSpread.DeleteFlag Then
   	    frm1.vspddata4.Text = ggoSpread.UpdateFlag
	End if

End Sub

'=======================================================================================================
'   Function Name : FindData22
'   Function Desc : 현재의 Item, Dtl에 해당하는 Hidden Grid의 Index를 Return
'                   관리항목그리드의 Value를 자동settgin할때 Hidden Grid로 복사하기(CopyToHSheet21의 ActiveRow의 맹점보완)
'=======================================================================================================
Function FindData22(MasterRow,DetailRow)

	Dim strApNo
	Dim strItemSeq
	Dim strDtlSeq
	Dim lRows

    FindData22 = 0

    With frm1

        For lRows = 1 To .vspddata6.MaxRows

            .vspddata6.Row = lRows
            .vspddata6.Col = 1
            strItemSeq = .vspddata6.Text
            .vspddata6.Col = 2
            strDtlSeq = .vspddata6.Text

            .vspddata4.Row = MasterRow'frm1.vspddata4.ActiveRow
            .vspddata5.Row = DetailRow

            .vspddata4.Col = C_ItemSeq
            If strItemSeq = .vspddata4.Text Then

                .vspddata5.Col = C_DtlSeq_2
                If strDtlSeq = .vspddata5.Text Then

                    FindData22 = lRows
                    Exit Function

                End If

            End If
        Next

    End With

End Function
Sub CopyToHSheet4(ByVal MasterRow, ByVal DetailRow)
Dim lRow
'Dim iCols
	With frm1 
        
	    lRow = FindData4(MasterRow,DetailRow)
'		msgbox "CopyToHSheet4 FindData4=" & lRow
	    If lRow > 0 Then
			.vspdData6.Row = lRow
            .vspdData5.Row = DetailRow
            .vspdData6.Col = 0
            .vspdData5.Col = 0
            .vspdData6.Text = .vspdData5.Text

			.vspdData5.Col = C_DtlSeq_2
			.vspdData6.Col = 2
			.vspdData6.Text = .vspdData5.Text
					    
			.vspdData5.Col = C_CtrlCd_2
			.vspdData6.Col = .vspdData6.Col + 1
			.vspdData6.Text = .vspdData5.Text
					    
			.vspdData5.Col = C_CtrlNm_2
			.vspdData6.Col = .vspdData6.Col + 1
			.vspdData6.Text = .vspdData5.Text
					    
			.vspdData5.Col = C_CtrlVal_2
			.vspdData6.Col = .vspdData6.Col + 1
			.vspdData6.Text = .vspdData5.Text
					    
			.vspdData5.Col = C_CtrlPB_2
			.vspdData6.Col = .vspdData6.Col + 1
			.vspdData6.Text = .vspdData5.Text
					    
			.vspdData5.Col = C_CtrlValNm_2
			.vspdData6.Col = .vspdData6.Col + 1
			.vspdData6.Text = .vspdData5.Text
					    
			.vspdData5.Col = C_Seq_2
			.vspdData6.Col = .vspdData6.Col + 1
			.vspdData6.Text = .vspdData5.Text
					    
			.vspdData5.Col = C_Tableid_2
			.vspdData6.Col = .vspdData6.Col + 1
			.vspdData6.Text = .vspdData5.Text
					    
			.vspdData5.Col = C_Colid_2
			.vspdData6.Col = .vspdData6.Col + 1
			.vspdData6.Text = .vspdData5.Text
					    
			.vspdData5.Col = C_ColNm_2
			.vspdData6.Col = .vspdData6.Col + 1
			.vspdData6.Text = .vspdData5.Text
					    
			.vspdData5.Col = C_Datatype_2
			.vspdData6.Col = .vspdData6.Col + 1
			.vspdData6.Text = .vspdData5.Text
					    
			.vspdData5.Col = C_DataLen_2
			.vspdData6.Col = .vspdData6.Col + 1
			.vspdData6.Text = .vspdData5.Text
					    
			.vspdData5.Col = C_DRFg_2
			.vspdData6.Col = .vspdData6.Col + 1
			.vspdData6.Text = .vspdData5.Text
					    
			.vspdData5.Col = C_MajorCd_2
			.vspdData6.Col = .vspdData6.Col + 1
			.vspdData6.Text = .vspdData5.Text
			
			.vspdData5.Col = C_MajorCd_2 + 1
			.vspdData6.Col = .vspdData6.Col + 1
			.vspdData6.Text = .vspdData5.Text

            
            
        End If

	End With
	
	frm1.vspddata4.Row = MasterRow'frm1.vspddata4.ActiveRow
	frm1.vspddata4.Col = 0
	if frm1.vspddata4.Text <> ggoSpread.InsertFlag and frm1.vspddata4.Text <> ggoSpread.DeleteFlag then
   	    frm1.vspddata4.Text = ggoSpread.UpdateFlag
	End if
	
End Sub

'=======================================================================================================
'   Function Name : FindData4
'   Function Desc : 현재의 Item, Dtl에 해당하는 Hidden Grid의 Index를 Return
'                   관리항목그리드의 Value를 자동settgin할때 Hidden Grid로 복사하기(CopyToHSheet3의 ActiveRow의 맹점보완)
'=======================================================================================================
Function FindData4(MasterRow,DetailRow)
Dim strApNo
Dim strItemSeq
Dim strDtlSeq
Dim lRows

    FindData4 = 0

    With frm1
        
        For lRows = 1 To .vspddata6.MaxRows
        
            .vspddata6.Row = lRows
            .vspddata6.Col = 1
            strItemSeq = .vspddata6.Text
            .vspddata6.Col = 2
            strDtlSeq = .vspddata6.Text
            
            .vspddata4.Row = MasterRow'frm1.vspddata4.ActiveRow
            .vspddata5.Row = DetailRow
            
            .vspddata4.Col = C_ItemSeq
            If strItemSeq = .vspddata4.Text Then
                
                
                .vspddata5.Col = C_DtlSeq
                
                If strDtlSeq = .vspddata5.Text Then
                    
                    FindData4 = lRows
                    Exit Function
                    
                End If
                
            End If    
        Next
        
    End With        
    
End Function


'=======================================================================================================
'   Function Name : AcctCheck
'   Function Desc : 1.입출금 전표일때 계정이 현금계정으로 입력되는지 check
'                   2.미결관리를 하는 회사일 경우 미결반제계정이 들어오는지 확인.
'=======================================================================================================
Function AcctCheck2(Byval Acctcd, Byval Inputtype, Byval DrCrFg )
	Dim arrTemp
	Dim strOpenAcctFg
	
	Dim arrTemp1
	Dim strDrCrFg
	Dim strMgntFg

	AcctCheck2 = False

	IF Inputtype <> "03" and Acct_cd = lgCashAcct then
		IntRetCD = DisplayMsgBox("113106", "X", "X", "X")		
		frm1.vspdData4.Text = ""
		frm1.vspdData4.Col = C_AcctNm
		frm1.vspdData4.Text = ""		
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
			frm1.vspdData4.Text = ""
			frm1.vspdData4.Col = C_AcctNm
			frm1.vspdData4.Text = ""		
			Exit Function		
		END If	
	End IF	
	
	AcctCheck2 = True	

End Function    



