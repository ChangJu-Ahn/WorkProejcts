'=======================================================================================================
'	관리항목 그리드 상수선언 
'=======================================================================================================
Const C_NoteSep = ","							'비고 seperate.

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
Dim C_MajorCd

Dim lgCashAcct									'현금계정을 미리 저장한다.
Dim lgAuthorityFlag                             '권한관리 추가 
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
	C_MajorCd   = 14

End Sub
'=======================================================================================================
'   Event Name : InitCtrlSpread()
'   Event Desc : 관리항목 그리드 초기화 
'=======================================================================================================
Sub InitCtrlSpread()
	
	Call initCtrlSpreadPosVariables()
	
    With frm1

		ggoSpread.Source			= .vspdData2
		ggoSpread.Spreadinit "V20021103",,parent.gAllowDragDropSpread

		.vspdData2.ReDraw			= False
		
		.vspdData2.AutoClipboard	= False
		.vspdData2.MaxCols			= C_MajorCd + 1
		.vspdData2.Col				= .vspdData2.MaxCols
		.vspdData2.ColHidden		= True

		.vspdData2.MaxRows			= 0

		Call parent.AppendNumberPlace("6","3","0")

		Call GetCtrlSpreadColumnPos("A")
		ggoSpread.SSSetFloat		C_DtlSeq,		"NO" ,				6,	"6",	parent.ggStrIntegeralPart,	parent.ggStrDeciPointPart,	parent.gComNum1000,	parent.gComNumDec,	2,	,	,	"0",	"999"
		ggoSpread.SSSetEdit		C_CtrlCd,		"관리항목",			10,	2
		ggoSpread.SSSetEdit		C_CtrlNm,		"관리항목명",		30,	3
		ggoSpread.SSSetEdit		C_CtrlVal,		"관리항목 VALUE",	32,	,		,							30,							2
		ggoSpread.SSSetButton	C_CtrlPB
		ggoSpread.SSSetEdit		C_CtrlValNm,	"관리항목 VALUE명",	45
		ggoSpread.SSSetEdit		C_Seq,			"A",					8,	,		,							3
		ggoSpread.SSSetEdit		C_Tableid,		"B",					32
		ggoSpread.SSSetEdit		C_Colid,		"C",					32
		ggoSpread.SSSetEdit		C_ColNm,		"D",					32
		ggoSpread.SSSetEdit		C_DataType,		"E",					2
		ggoSpread.SSSetFloat		C_DataLen,		"F",					3,	"6",	parent.ggStrIntegeralPart,	parent.ggStrDeciPointPart,	parent.gComNum1000,	parent.gComNumDec,	,	,	,	"0",	"999"
		ggoSpread.SSSetEdit		C_DRFg,			"G",					1
		ggoSpread.SSSetEdit		C_MajorCd,		"H",					1

		Call ggoSpread.MakePairsColumn(C_CtrlVal,C_CtrlPB)

		Call ggoSpread.SSSetColHidden(C_CtrlCd,C_CtrlCd,True)
		Call ggoSpread.SSSetColHidden(C_Seq,C_Seq,True)
		Call ggoSpread.SSSetColHidden(C_Tableid,C_Tableid,True)
		Call ggoSpread.SSSetColHidden(C_Colid,C_Colid,True)
		Call ggoSpread.SSSetColHidden(C_ColNm,C_ColNm,True)
		Call ggoSpread.SSSetColHidden(C_DataType,C_DataType,True)
		Call ggoSpread.SSSetColHidden(C_DataLen,C_DataLen,True)
		Call ggoSpread.SSSetColHidden(C_MajorCd,C_MajorCd,True)
		Call ggoSpread.SSSetColHidden(C_DRFg,C_DRFg,True)

		.vspdData2.ReDraw = True

    End With
    
End Sub

'=======================================================================================================
' Function Name : CtrlSpreadLock
' Function Desc : 관리항목 그리드 Lock
'=======================================================================================================
Sub CtrlSpreadLock(Byval stsFg, Byval Index, ByVal lRow  , ByVal lRow2 )

	Dim objSpread
    
	With frm1

		ggoSpread.Source	= .vspdData2
		Set objSpread		= .vspdData2
		lRow2				= objSpread.MaxRows
		objSpread.Redraw	= False

		ggoSpread.SpreadLock  1, lRow,  2, lRow2
		ggoSpread.SpreadLock  3, lRow,  3, lRow2
		ggoSpread.SpreadLock  4, lRow,  4, lRow2
		ggoSpread.SpreadLock  5, lRow,  5, lRow2
		ggoSpread.SpreadLock  6, lRow,  6, lRow2
		    		
		objSpread.Redraw = True
		Set objSpread = Nothing

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
		strEndRow	= .vspddata2.MaxRows

		ggoSpread.Source	= .vspdData2
		.vspdData2.ReDraw	= False
	'	.vspddata.Col		= C_DrCRFG
	'	tmpDrCrFG			= LEFT(.vspddata.Text,1)
		
		ggoSpread.SSSetProtected C_DtlSeq,		strStartRow,	strEndRow
		ggoSpread.SSSetProtected C_CtrlCd,		strStartRow,	strEndRow
		ggoSpread.SSSetProtected C_CtrlNm,		strStartRow,	strEndRow
		ggoSpread.SSSetProtected C_CtrlValNm,	strStartRow,	strEndRow

		For indx = 1 to .vspddata2.MaxRows

			.vspddata2.Row = indx
			.vspddata2.Col = C_DrFg

			'If (.vspddata2.text = tmpDrCrFG And .vspddata2.text <> "") Or .vspddata2.text = "Y" Or .vspddata2.text = "DC" Then
			If (.vspddata2.text <> "") Or .vspddata2.text = "Y" Or .vspddata2.text = "DC" Then	
				ggoSpread.SSSetRequired C_CtrlVal, indx, indx
			Else
				ggoSpread.SpreadUnLock C_CtrlVal, indx, C_CtrlVal, indx
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
			C_MajorCd   = iCurColumnPos(14)
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
	ENd if
	
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

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
					'Call CopyToHSheet2(frm1.vspdData.ActiveRow,.vspdData2.Row)
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

			call vspdData2_Change( C_CtrlVal ,  .vspdData2.Row)
		End if
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
				frm1.vspdData2.text = arrVal(0)
			ELSE
				IntRetCD = DisplayMsgBox("110330", "X", "X", "X")									'필수입력 check!!
				' 관리항목값이 바르지 않습니다.
				frm1.vspdData2.text = ""
				frm1.vspdData2.Col = C_CtrlVal
				frm1.vspdData2.text = ""
				Exit Function
			END IF
		End if

	End With

End Function


'=======================================================================================================
' Function Name : CheckSpread4
' Function Desc : 저장시에  관리항목 필수여부 check 하기위해 호출되는 Function
'=======================================================================================================

Function CheckSpread4()

	Dim indx
	Dim tmpDrCrFG

	CheckSpread4 = False

	With frm1
		For indx = 1 to .vspddata2.MaxRows

			.vspddata2.Row = indx
			.vspddata2.Col = C_DrFg
			If (.vspddata2.text <> "") Or .vspddata2.text = "Y" Or .vspddata2.text = "DC" Then	
			  .vspdData2.Col = C_CtrlVal
			  If Trim(.vspdData2.text) = "" Then
				Exit Function
		  	  End If
		    End If
		Next

    End With

	CheckSpread4 = True

End Function

'=======================================================================================================
' Function Name : DbQuery4
' Function Desc : Item 그리드 변경시 관리항목 조회 
'=======================================================================================================
Function DbQuery4()


	Dim	IDtlRow
	Dim strVal
	Dim lngRows
	Dim strSelect
	Dim strSelect1
	Dim strWhere
	Dim IntRetCD
	Dim arrVal

	frm1.vspdData2.MaxRows = 0
	
    Call LayerShowHide(1)
	If CommonQueryRs("ACCT_NM" , "A_ACCT (NOLOCK)" ,  "ACCT_CD = " & FilterVar(frm1.txtAcctCd.value, "''", "S"), lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then   			        								
    	arrVal				= Split(lgF0, Chr(11))
		frm1.txtAcctNm.value	= arrVal(0)
	Else
		frm1.txtAcctNm.value = ""
		IntRetCD			= DisplayMsgBox("110100", "X", "X", "X")
		Call LayerShowHide(0)
		Exit Function
	END IF

	

    strSelect =	            " B.CTRL_ITEM_SEQ,  A.CTRL_CD, A.CTRL_NM , '', '',"
    strSelect = strSelect & " CASE  WHEN A.COLM_DATA_TYPE = " & FilterVar("D", "''", "S") & "  THEN " & FilterVar("(Format : YYYY-MM-DD)", "''", "S") & "  END , 1, LTrim(ISNULL(A.TBL_ID,'')),LTrim(ISNULL(A.DATA_COLM_ID,'')), "
    strSelect = strSelect & " LTrim(ISNULL(A.DATA_COLM_NM,'')),  LTrim(ISNULL(A.COLM_DATA_TYPE,'')), LTrim(ISNULL(A.DATA_LEN,'')), "
    strSelect = strSelect & " CASE WHEN B.DR_FG = " & FilterVar("Y", "''", "S") & "  AND  B.CR_FG = " & FilterVar("Y", "''", "S") & "  THEN " & FilterVar("DC", "''", "S") & "  "
    strSelect = strSelect & " WHEN B.DR_FG = " & FilterVar("Y", "''", "S") & "  AND  B.CR_FG = " & FilterVar("N", "''", "S") & "  THEN " & FilterVar("D", "''", "S") & "  "
    strSelect = strSelect & " WHEN B.DR_FG = " & FilterVar("N", "''", "S") & "  AND  B.CR_FG = " & FilterVar("Y", "''", "S") & "  THEN " & FilterVar("C", "''", "S") & "  "
    strSelect = strSelect & " END	, "
    strSelect = strSelect & " LTrim(ISNULL(A.MAJOR_CD, '')) "

    strSelect1 =  "1 ,B.CTRL_ITEM_SEQ,  A.CTRL_CD, A.CTRL_NM , '', '',"
    strSelect1 = strSelect1 & " CASE  WHEN A.COLM_DATA_TYPE = " & FilterVar("D", "''", "S") & "  THEN " & FilterVar("(Format : YYYY-MM-DD)", "''", "S") & "  END , 1, LTrim(ISNULL(A.TBL_ID,'')),LTrim(ISNULL(A.DATA_COLM_ID,'')), "
    strSelect1 = strSelect1 & " LTrim(ISNULL(A.DATA_COLM_NM,'')),  LTrim(ISNULL(A.COLM_DATA_TYPE,'')), LTrim(ISNULL(A.DATA_LEN,'')), "
    strSelect1 = strSelect1 & " CASE WHEN B.DR_FG = " & FilterVar("Y", "''", "S") & "  AND  B.CR_FG = " & FilterVar("Y", "''", "S") & "  THEN " & FilterVar("DC", "''", "S") & "  "
    strSelect1 = strSelect1 & " WHEN B.DR_FG = " & FilterVar("Y", "''", "S") & "  AND  B.CR_FG = " & FilterVar("N", "''", "S") & "  THEN " & FilterVar("D", "''", "S") & "  "
    strSelect1 = strSelect1 & " WHEN B.DR_FG = " & FilterVar("N", "''", "S") & "  AND  B.CR_FG = " & FilterVar("Y", "''", "S") & "  THEN " & FilterVar("C", "''", "S") & "  "
    strSelect1 = strSelect1 & " END	, "
    strSelect1 = strSelect1 & " LTrim(ISNULL(A.MAJOR_CD, '')) "

	strWhere =  "A.CTRL_CD = B.CTRL_CD  AND B.ACCT_CD =  " & FilterVar(Frm1.txtAcctCd.value, "''", "S") & ""
	strWhere =  strWhere & " Order By B.CTRL_ITEM_SEQ "

	frm1.vspdData2.ReDraw = False
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	If CommonQueryRs2by2(strSelect, " A_CTRL_ITEM	A (NOLOCK), A_ACCT_CTRL_ASSN B (NOLOCK)" , strWhere , lgF2By2) Then
		ggoSpread.Source = frm1.vspdData2
		
		ggoSpread.SSShowData lgF2By2

		For lngRows = 1 to frm1.vspdData2.Maxrows
			frm1.vspddata2.row	= lngRows
			frm1.vspddata2.Col	= 0
			frm1.vspddata2.Text	= ggoSpread.InsertFlag
		Next

		
		Call SetSpread4Color()

    End If

    frm1.vspdData2.ReDraw = True
    Call LayerShowHide(0)

End Function

'=======================================================================================================
' Function Name : DbQueryOk3
' Function Desc : DbQuery3가 성공적일 경우 MyBizASP 에서 호출되는 Function
'=======================================================================================================
Function DbQueryOk4()

	Call SetSpread4Color()

End Function

'=======================================================================================================
'   Event Name : vspdData2_ButtonClicked
'   Event Desc : 관리항목 팝업버튼 클릭시 관리항목 팝업호출 
'=======================================================================================================
Sub vspdData2_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)

	Dim iFld1
	Dim iFld2
	Dim iTable
	Dim istrCode
	Dim FldNm
	Dim strSelect, strFrom, strWhere
	Dim arrTemp
	Dim strVatType
	Dim strVatRate
	Dim lRows
	Dim indx

	'---------- Coding part -------------------------------------------------------------
	ggoSpread.Source = frm1.vspdData2

	With frm1.vspdData2

		If Row > 0 And Col = C_CtrlPB Then
			.Row		= Row
			.Col		= C_ctrlNm
			FldNm		= .Text

			.Col		= C_CtrlVal
			istrCode	= .Text

			.Col		= C_Tableid
			iTable		= .Text

			.Col		= C_Colid
			iFld1		= .Text

			.Col		= C_ColNm
			iFld2		= .Text

			.Col		= C_MajorCD

			If .Text <> "" Then
				strWhere = " Major_CD =  " & FilterVar(.Text , "''", "S") & ""
			Else
				strWhere = ""
			End If

			If iTable <> "" Then
 				Call OpenCtrlPB(iTable, iFld1, iFld2, istrCode, FldNm, strWhere)
			End if

			frm1.vspddata2.Col = C_CtrlCd
			
			If Trim(frm1.vspddata2.Text) = "V4" Then
				frm1.vspdData2.Col	= C_CtrlVal
				strVatType			= Trim(frm1.vspdData2.text)
				
				If Trim(strVatType) <> "" Then
					strSelect	= "reference"
					strFrom		= "b_configuration"
					strWhere	= "major_cd = " & FilterVar("B9001", "''", "S") & "  and seq_no = 1 and minor_cd =  " & FilterVar(strVatType , "''", "S") & ""

					If CommonQueryRs(strSelect, strFrom, strWhere, lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) Then
						arrTemp = Split(lgF0, chr(11))
						strVatRate = arrTemp(0)
					End If

					frm1.vspddata2.Col = C_CtrlCd
					
					For indx = 1 To frm1.vspdData2.MaxRows
						frm1.vspdData2.Row = indx
						If Trim(frm1.vspddata2.Text) = "V7" Then
							frm1.vspddata2.Col = C_CtrlVal
							frm1.vspddata2.Text = strVatRate
						End If
					Next
				End If
			End If
		End If

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
				sPreCtrlVal = frm1.vspdData2.text
				If IsDate(frm1.vspdData2.text) = False or IsNumeric(Mid(frm1.vspdData2.text,1,4)) = False or _
					IsNumeric(Mid(frm1.vspdData2.text,6,2)) = False or _
					IsNumeric(Mid(frm1.vspdData2.text,9,2)) = False or _
					Mid(frm1.vspdData2.text,5,1) <> "-" or _
					Mid(frm1.vspdData2.text,8,1) <> "-" or _
					Mid(frm1.vspdData2.text,1,4) < "1900" Then
						frm1.vspdData2.text = sPreCtrlVal
						IntRetCD = DisplayMsgBox("174223", "X", "X", "X")							'필수입력 check!!
						' 입력하신 날짜는 부적합합니다.
						frm1.vspdData2.text = ""
						Exit Sub
				End If
			ElseIf Trim(frm1.vspdData2.Text) = "N" Then
				frm1.vspdData2.Col = C_CtrlVal
				sPreCtrlVal = frm1.vspdData2.text
				If IsNumeric(frm1.vspdData2.text) = False Then
					frm1.vspdData2.text = sPreCtrlVal
					IntRetCD = DisplayMsgBox("229924", "X", "X", "X")								'필수입력 check!!
					' 숫자를 입력하십시오 
					frm1.vspdData2.text = ""
					Exit Sub
				End If
	        End If

	        '------------------------------------
	        ' 입력된 관리항목의 길이Check
	        '------------------------------------
	        frm1.vspdData2.Col = C_CtrlVal

	        iLen = Len(frm1.vspdData2.text)
			sPreCtrlVal = frm1.vspdData2.text
	        frm1.vspdData2.Col = C_DataLen

	        If iLen > Int(frm1.vspdData2.text) Then
				frm1.vspdData2.text = sPreCtrlVal
				IntRetCD = DisplayMsgBox("110320", "X", "X", "X")									'필수입력 check!!
			'  관리항목값의 길이를 확인하십시오.
				frm1.vspdData2.Col = C_CtrlVal
				frm1.vspdData2.text = ""
				Exit Sub
	        End If

	        frm1.vspdData2.Col = C_DataType

	        If Trim(frm1.vspdData2.Text) <> "D" And Trim(frm1.vspdData2.Text) <> "N" Then
				FindCtrlNM   Row																			'관리항목값을 check하고 관리항목명을 찾아준다.
			End If
    End Select

	

    lgBlnFlgChgValue = True

End Sub

'==========================================================================================
'   Event Name : vspdData2_Click
'   Event Desc : This event is spread sheet data changed
'==========================================================================================
Sub vspdData2_Click(ByVal Col, ByVal Row)

	gMouseClickStatus = "SP2C"	'Split 상태코드 
	   
	Set gActiveSpdSheet = frm1.vspdData2

	If frm1.vspdData2.MaxRows = 0 Then                                                    'If there is no data.
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

	If Row <= 0 Then
	   ggoSpread.Source = frm1.vspdData2
	   Exit Sub
	End If
    Call SetPopupMenuItemInf("0000111111")

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
'        Call parent.MASetToolbar("1110100000011111")                                     '버튼 툴바 제어 
'    Else
'        Call parent.MASetToolbar("1111100000011111")                                     '버튼 툴바 제어 
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
	objSpread.text	= MaxValue

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


' 권한관리 추가 ==============================================================
'==========================================================================================
'	Name : SetAuthorityFlag
'	Description :
'==========================================================================================
Sub SetAuthorityFlag()

	If CommonQueryRs("TOP 1 USR_ID", "Z_USR_AUTHORITY_VALUE", "USR_ID =  " & FilterVar(gUsrId , "''", "S") & " AND MODULE_CD = " & FilterVar("A", "''", "S") & "  ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) Then
	    If UCase(gUsrId) = UCase(Replace(lgF0,Chr(11),"")) Then
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
				arrstrRet(1) = "ORG_CHANGE_ID =  " & FilterVar(strCode1, "''", "S") & " AND A.DEPT_CD = B.CODE_VALUE AND B.USR_ID =  " & FilterVar(gUsrId , "''", "S") & " AND B.MODULE_CD = " & FilterVar("A", "''", "S") & " "		 '권한관리 추가 
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

