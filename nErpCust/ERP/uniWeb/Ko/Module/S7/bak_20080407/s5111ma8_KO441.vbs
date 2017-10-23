  ' External ASP File
'========================================
Const BIZ_PGM_ID 		= "s5111MB8_KO441.asp"					                       '☆: Biz Logic ASP Name
Const BIZ_PGM_JUMP_ID1	= "s5111ma1"												'☆: JUMP시 비지니스 로직 ASP명 
Const BIZ_PGM_JUMP_ID2	= "s5111ma2"												'☆: JUMP시 비지니스 로직 ASP명 

' Constant variables 
'========================================
Const C_MaxKey          = 13                                           '☆: key count of SpreadSheet

Const C_PopSoldToParty	= 1
Const C_PopSalesGrp		= 2
Const C_PopBillType		= 3
Const C_PopSoNo			= 4


' User-defind Variables
'========================================
Dim IsOpenPop  

Dim lgBlnOpenedFlag
Dim	lgBlnSoldToPartyChg
Dim lgBlnSalesGrpChg
Dim	lgBlnBillTypeChg
Dim lgIntStartRow

'========================================
Function InitVariables()
	lgPageNo         = ""
    lgIntFlgMode     = Parent.OPMD_CMODE                          'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False	                           'Indicates that no value changed
    lgSortKey        = 1   

    lgStrPrevKey = ""										'initializes Previous Key
	    
	lgBlnSoldToPartyChg = False		' 주문처 변경여부 
	lgBlnSalesGrpChg	= False		' 영업그룹 변경여부 
	lgBlnBillTypeChg	= False		' 매출채권 변경여부 
End Function

'=======================================================
Sub SetDefaultVal()
	With frm1
		.txtFromDt.Text = UNIGetFirstDay(EndDate, Parent.gDateFormat)
		.txtToDt.Text = EndDate	
		.rdoPostfiFlagAll.checked = True
		.txtPostfiFlag.value = frm1.rdoPostfiFlagAll.value   
		If Parent.gSalesGrp <> "" Then
			.txtSalesGrp.value = Parent.gSalesGrp
			Call GetSalesGrpNm()
		End If

		Call SetFocusToDocument("M")
		.txtFromDt.Focus
	End With
	lgBlnFlgChgValue = False
End Sub

'========================================
Sub InitSpreadSheet()
	
	Call SetZAdoSpreadSheet("S5111MA8","S","A","V20030714", Parent.C_SORT_DBAGENT, frm1.vspdData, C_MaxKey, "X", "X" )
	
	Call SetSpreadLock 
	    
End Sub

'========================================
Sub SetSpreadLock()
	ggoSpread.SpreadLockWithOddEvenRowColor()
'	frm1.vspdData.OperationMode = 3
End Sub	

'========================================
Function CookiePage()

	On Error Resume Next

	Const CookieSplit = 4877						
	
	If frm1.vspdData.ActiveRow > 0 Then
		frm1.vspdData.Row = frm1.vspdData.ActiveRow
		frm1.vspdData.Col = GetKeyPos("A",1)
		WriteCookie CookieSplit , frm1.vspdData.Text
	Else
		WriteCookie CookieSplit , ""
	End If

End Function

'========================================
Function JumpChgCheck(ByVal Choice)

	Const CookieSplit = 4877

	ggoSpread.Source = frm1.vspdData

	frm1.vspdData.Row = frm1.vspdData.ActiveRow
	frm1.vspdData.Col = GetKeyPos("A",2)						'매출구분 

	Select Case Choice
	Case 1
		If Trim(frm1.vspdData.Text) = "N" OR frm1.vspdData.Row = 0 Then		
			PgmJump(BIZ_PGM_JUMP_ID1)
		Else
			MsgBox "해당 매출채권번호는 정상매출채권이 아닙니다.", vbInformation, parent.gLogoName
			WriteCookie CookieSplit , ""
			Exit Function
		End If
	Case 2
		If Trim(frm1.vspdData.Text) = "Y" OR frm1.vspdData.Row = 0 Then		
			PgmJump(BIZ_PGM_JUMP_ID2)
		Else
			MsgBox "해당 매출채권번호는 예외매출채권이 아닙니다.", vbInformation, parent.gLogoName
			WriteCookie CookieSplit , ""
			Exit Function
		End If
	End Select

End Function

'========================================
Function OpenBillDtl()
	Dim iCalledAspName
	Dim iArrParam(3)
	
	On Error Resume Next

	If lgIntFlgMode <> Parent.OPMD_UMODE Then
		Call DisplayMsgBox("900002", "X", "X", "X")
		Exit Function
	End IF

	If IsOpenPop = True Then Exit Function

	With frm1.vspdData
		.Row = .activerow
		.Col = GetKeyPos("A",1)		:	iArrParam(0) = .Text	' 매출채권번호 
		.Col = GetKeyPos("A",3)		:	iArrParam(1) = .Text	' 주문처 
		.Col = GetKeyPos("A",4)		:	iArrParam(2) = .Text	' 주문처명 
		.Col = GetKeyPos("A",13)	:	iArrParam(3) = .Text	' 화폐 
	End With
	
	IsOpenPop = True
   
	iCalledAspName = AskPRAspName("s5112ra7")	
	if Trim(iCalledAspName) = "" then
		Call DisplayMsgBox("900040",parent.VB_INFORMATION, "s5112ra7", "X")
		IsOpenPop = False
		exit Function
	end if

	Call window.showModalDialog(iCalledAspName,Array(window.parent,iArrParam),"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

End Function

'========================================
Function OpenSoNo()
	Dim iCalledAspName
	Dim iStrRet

	If IsOpenPop = True Then Exit Function
			
	IsOpenPop = True

	frm1.txtSoNo.focus
			
	iCalledAspName = AskPRAspName("s3111pa1")	
	if Trim(iCalledAspName) = "" then
		Call DisplayMsgBox("900040",parent.VB_INFORMATION, "s3111pa1", "X")
		IsOpenPop = False
		Exit Function
	end if

	iStrRet = window.showModalDialog(iCalledAspName, Array(window.parent,""), _
		"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If iStrRet <> "" Then
		frm1.txtSoNo.value = iStrRet 
	End If
End Function

'========================================
Function OpenConPopup(ByVal pvIntWhere)

	Dim iArrRet
	Dim iArrParam(5), iArrField(6), iArrHeader(6)

	OpenConPopup = False
	
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	Select Case pvIntWhere
	Case C_PopBillType												
		iArrParam(1) = "s_bill_type_config"				' TABLE 명칭 
		iArrParam(2) = Trim(frm1.txtBillType.value)		' Code Condition
		iArrParam(3) = ""								' Name Cindition
		iArrParam(4) = "USAGE_FLAG=" & FilterVar("Y", "''", "S") & " "					' Where Condition
		iArrParam(5) = "매출채권형태"				' TextBox 명칭 

		iArrField(0) = "ED15" & Parent.gColSep & "bill_type"
		iArrField(1) = "ED30" & Parent.gColSep & "bill_type_nm"

		iArrHeader(0) = "매출채권형태"				' Header명(0)
		iArrHeader(1) = "매출채권형태명"			' Header명(1)
		
		frm1.txtBillType.focus

	Case C_PopSalesGrp												
		iArrParam(1) = "B_SALES_GRP"
		iArrParam(2) = Trim(frm1.txtSalesGrp.value)
		iArrParam(3) = ""
		iArrParam(4) = "USAGE_FLAG=" & FilterVar("Y", "''", "S") & " "
		iArrParam(5) = "영업그룹"
		
		iArrField(0) = "ED15" & Parent.gColSep & "SALES_GRP"
		iArrField(1) = "ED30" & Parent.gColSep & "SALES_GRP_NM"
    
	    iArrHeader(0) = "영업그룹"
	    iArrHeader(1) = "영업그룹명"
	    
	    frm1.txtSalesGrp.focus

	Case C_PopSoldToParty												
		iArrParam(1) = "B_BIZ_PARTNER"
		iArrParam(2) = Trim(frm1.txtSoldToParty.value)
		iArrParam(3) = ""
'		iArrParam(4) = "USAGE_FLAG=" & FilterVar("Y", "''", "S") & "  AND BP_TYPE IN (" & FilterVar("CS", "''", "S") & ", " & FilterVar("C", "''", "S") & " )"
		iArrParam(4) = "BP_TYPE IN (" & FilterVar("CS", "''", "S") & ", " & FilterVar("C", "''", "S") & " )"
		iArrParam(5) = "주문처"
			
		iArrField(0) = "ED15" & Parent.gColSep & "BP_CD"
		iArrField(1) = "ED30" & Parent.gColSep & "BP_NM"
		    
		iArrHeader(0) = "주문처"
		iArrHeader(1) = "주문처명"

		frm1.txtSoldToParty.focus
		
	End Select
 
	iArrParam(0) = iArrParam(5)							

	iArrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(iArrParam, iArrField, iArrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If iArrRet(0) <> "" Then
		Call SetConPopup(iArrRet,pvIntWhere)
		OpenConPopup = True
	End If	
End Function

'========================================
Sub PopZAdoConfigGrid()
	If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
		Exit Sub
	End If
	
	Call OpenOrderByPopup("A")
End Sub

'========================================
Function OpenOrderByPopup(ByVal pSpdNo)
	Dim arrRet
	
	On Error Resume Next
	
	If IsOpenPop = True Then Exit Function
	IsOpenPop = True

	arrRet = window.showModalDialog("../../ComAsp/ZADOSortPopup.asp",Array(ggoSpread.GetXMLData("A"),gMethodText),"dialogWidth=" & parent.SORTW_WIDTH & "px; dialogHeight=" & parent.SORTW_HEIGHT & "px;; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	
	If arrRet(0) = "X" Then
	   Exit Function
	Else
	   Call ggoSpread.SaveXMLData(pSpdNo,arrRet(0),arrRet(1))
       Call InitVariables
       Call InitSpreadSheet()       
   End If
End Function

'=======================================================
Function SetConPopup(Byval pvArrRet,ByVal pvIntWhere)

	SetConPopup = False

	Select Case pvIntWhere
	Case C_PopBillType
		frm1.txtBillType.value = pvArrRet(0) 
		frm1.txtBillTypeNm.value = pvArrRet(1)   
	Case C_PopSoldToParty
		frm1.txtSoldToParty.value = pvArrRet(0) 
		frm1.txtSoldToPartyNm.value = pvArrRet(1)   
	Case C_PopSalesGrp
		frm1.txtSalesGrp.value = pvArrRet(0) 
		frm1.txtSalesGrpNm.value = pvArrRet(1)   
	End Select

	SetConPopup = True

End Function

Sub FormatField()
    With frm1
        ' 날짜 OCX Foramt 설정 
        Call FormatDATEField(.txtFromDt)
        Call FormatDATEField(.txtToDt)
    End With
End Sub

Sub LockFieldInit()
    With frm1
        ' 날짜 OCX
        Call LockObjectField(.txtFromDt, "O")
        Call LockObjectField(.txtToDt, "O")
    End With

End Sub

'========================================
Sub Form_Load()
    Call LoadInfTB19029											  '⊙: Load table , B_numeric_format
   
    'Html에서 tag 숫자가 1과 2로 시작하는 부분 각각Format
'    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
	
'	Call ggoOper.LockField(Document, "N")                         '⊙: Lock  Suitable  Field
	Call FormatField    
	Call LockFieldInit
    Call InitVariables
        Call GetValue_ko441()											  '⊙: Initializes local global variables
	Call SetDefaultVal	
	Call InitSpreadSheet()
	lgBlnOpenedFlag = True
    Call SetToolbar("11000000000011")										'⊙: 버튼 툴바 제어 
End Sub

'========================================
Sub Form_QueryUnload(Cancel, UnloadMode)
End Sub

'========================================
Function GetSalesGrpNm()
	Dim iStrCode
	
	iStrCode = Trim(frm1.txtSalesGrp.value)
	If iStrCode <> "" Then
		iStrCode = " " & FilterVar(iStrCode, "''", "S") & ""
		If Not GetCodeName(iStrCode, "default", "default", "default", "default", "" & FilterVar("SG", "''", "S") & "", C_PopSalesGrp) Then
			frm1.txtSalesGrp.value = ""
			frm1.txtSalesGrpNm.value = ""
		End If
	Else
		frm1.txtSalesGrpNm.value = ""
	End If
End Function

'========================================
Function txtSoldToParty_OnKeyDown()
	lgBlnSoldToPartyChg = True
	lgBlnFlgChgValue = True
End Function

'==========================================
Function txtSalesGrp_OnKeyDown()
	lgBlnSalesGrpChg = True
	lgBlnFlgChgValue = True
End Function

'==========================================
Function txtBillType_OnKeyDown()
	lgBlnBillTypeChg = True
	lgBlnFlgChgValue = True
End Function

'	Description : 조회조건의 유효성을 Check한다.
'   주의사항 : 화면의 tab order 별로 기술한다. 
'==========================================
Function ChkValidityQueryCon()
	Dim iStrCode
	Dim ChkVal
	ChkValidityQueryCon = True
	ChkVal = 0
	If lgBlnSoldToPartyChg Then
		iStrCode = Trim(frm1.txtSoldToParty.value)
		If iStrCode <> "" Then
			iStrCode = " " & FilterVar(iStrCode, "''", "S") & ""
			If Not GetCodeName(iStrCode, "" & FilterVar("C%", "''", "S") & "", "default", "default", "default", "" & FilterVar("BP", "''", "S") & "", C_PopSoldToParty) Then
				Call DisplayMsgBox("970000", "X", frm1.txtSoldtoparty.alt, "X")
				frm1.txtSoldToPartyNm.value = ""
				ChkValidityQueryCon = False
				ChkVal = 1
			End If
		Else
			frm1.txtSoldToPartyNm.value = ""
		End If
		lgBlnSoldToPartyChg	= False
	End If

	If lgBlnSalesGrpChg Then
		iStrCode = Trim(frm1.txtSalesGrp.value)
		If iStrCode <> "" Then
			iStrCode = " " & FilterVar(iStrCode, "''", "S") & ""
			If Not GetCodeName(iStrCode, "default", "default", "default", "default", "" & FilterVar("SG", "''", "S") & "", C_PopSalesGrp) Then
				If ChkValidityQueryCon = True Then
					Call DisplayMsgBox("970000", "X", frm1.txtSalesGrp.alt, "X")
					ChkValidityQueryCon = False
					ChkVal = 2
				End If
					frm1.txtSalesGrpNm.value = ""
			End If
		Else
			frm1.txtSalesGrpNm.value = ""
		End If
		lgBlnSalesGrpChg = False
	End If
			
	If lgBlnBillTypeChg Then
		iStrCode = Trim(frm1.txtBillType.value)
		If iStrCode <> "" Then
			iStrCode = " " & FilterVar(iStrCode, "''", "S") & ""
			If Not GetCodeName(iStrCode, "default", "default", "default", "default", "" & FilterVar("BT", "''", "S") & "", C_PopBillType) Then
				If ChkValidityQueryCon = True Then	
					Call DisplayMsgBox("970000", "X", frm1.txtBillType.alt, "X")
					ChkValidityQueryCon = False
					ChkVal = 3
				End If
				frm1.txtBillTypeNm.value = ""
			'	Exit Function
			End If
		Else
			frm1.txtBillTypeNm.value = ""
		End If
		lgBlnBillTypeChg = False
	End If
	If ChkValidityQueryCon = False Then
	Select Case ChkVal 
	
		Case 1
			frm1.txtSoldtoparty.focus
		Case 2
			frm1.txtSalesGrp.focus
		Case 3
			frm1.txtBillType.focus
		End Select 
	End If
End Function

'	Description : 코드값에 해당하는 명을 Display한다.
'====================================================
Function GetCodeName(ByVal pvStrArg1, ByVal pvStrArg2, ByVal pvStrArg3, ByVal pvStrArg4, ByVal pvIntArg5, ByVal pvStrFlag, ByVal pvIntWhere)

	Dim iStrSelectList, iStrFromList, iStrWhereList
	Dim iStrRs
	Dim iArrRs(2), iArrTemp
	
	GetCodeName = False

	iStrSelectList = " * "
	iStrFromList = " dbo.ufn_s_GetCodeName (" & pvStrArg1 & ", " & pvStrArg2 & ", " & pvStrArg3 & ", " & pvStrArg4 & ", " & pvIntArg5 & ", " & pvStrFlag & ") "
	iStrWhereList = ""
	
	Err.Clear
	
	If CommonQueryRs2by2(iStrSelectList,iStrFromList,iStrWhereList,iStrRs) Then
		iArrTemp = Split(iStrRs, Chr(11))
		iArrRs(0) = iArrTemp(1)
		iArrRs(1) = iArrTemp(2)
		iArrRs(2) = iArrTemp(3)
		GetCodeName = SetConPopup(iArrRs, pvIntWhere)
	Else
		' 관련 Popup Display
		'If lgBlnOpenedFlag Then GetCodeName = OpenConPopup(pvIntWhere)
	End if
End Function

'========================================
Sub rdoPostfiFlagAll_OnClick()
	frm1.txtPostfiFlag.value = frm1.rdoPostfiFlagAll.value 
End Sub

'========================================
Sub rdoPostfiFlagNo_OnClick()
	frm1.txtPostfiFlag.value = frm1.rdoPostfiFlagNo.value 
End Sub

'========================================
Sub rdoPostfiFlagYes_OnClick()
	frm1.txtPostfiFlag.value = frm1.rdoPostfiFlagYes.value 
End Sub

'========================================
Function vspdData_DblClick(ByVal Col, ByVal Row)
	If frm1.vspdData.ActiveRow > 0 Then	Call OpenBillDtl
End Function

'==========================================
Sub vspdData_Click(ByVal Col , ByVal Row)

    Call SetPopupMenuItemInf("00000000001")

	gMouseClickStatus = "SPC"

    Set gActiveSpdSheet = frm1.vspdData
        
	frm1.vspdData.ReDraw = False
    If Row = 0 Then
'		frm1.vspdData.OperationMode = 0

        If lgSortKey = 1 Then
            ggoSpread.SSSort Col, lgSortKey
            lgSortKey = 2
        Else
            ggoSpread.SSSort ,lgSortKey
            lgSortKey = 1
        End If    
	Else
'		frm1.vspdData.OperationMode = 3
    End If
End Sub

'==========================================
Sub vspdData_MouseDown(Button , Shift , x , y)
    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If
End Sub    

'========================================
Sub vspdData_TopLeftChange(ByVal OldLeft, ByVal OldTop, ByVal NewLeft, ByVal NewTop)
	
	If OldLeft <> NewLeft Then Exit Sub

	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	    '☜: 재쿼리 체크	
		If lgPageNo <> "" Then		                                                    '다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
			Call DisableToolBar(Parent.TBC_QUERY)
			Call DbQuery
		End If
	End If

End Sub

'========================================
Sub vspdData_Keypress(KeyAscii)
	If KeyAscii = 13 Then Call MainQuery()
End Sub

'========================================
Sub txtFromDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtFromDt.Action = 7		
		Call SetFocusToDocument("M")
		frm1.txtFromDt.Focus
	End If
End Sub

'========================================
Sub txtToDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtToDt.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtToDt.Focus
	End If
End Sub

'========================================
Sub txtFromDt_Keypress(KeyAscii)
	If KeyAscii = 13	Then Call MainQuery()
End Sub

'========================================
Sub txtToDt_Keypress(KeyAscii)
	If KeyAscii = 13	Then Call MainQuery()
End Sub

'========================================
Function FncQuery() 
    FncQuery = False                                                        '⊙: Processing is NG
    
    Err.Clear                                                               '☜: Protect system from crashing
	If ValidDateCheck(frm1.txtFromDt, frm1.txtToDt) = False Then Exit Function
   
    Call ggoOper.ClearField(Document, "2")	         						'⊙: Clear Contents  Field
    
	' 조회조건 유효값 check
	If 	lgBlnFlgChgValue Then
		If Not ChkValidityQueryCon Then	Exit Function
	End If
	
    Call InitVariables 														'⊙: Initializes local global variables
    
	If DbQuery = False Then Exit Function									

    FncQuery = True		
    
End Function

'========================================
Function FncPrint() 
    Call parent.FncPrint()
End Function

'========================================
Function FncExcel() 
	Call parent.FncExport(Parent.C_MULTI)
End Function

'========================================
Function FncFind() 
    Call parent.FncFind(Parent.C_MULTI , False)                                     
End Function

'========================================
Function FncSplitColumn()
    
    Dim ACol
    Dim ARow
    Dim iRet
    Dim iColumnLimit
    
    iColumnLimit  = C_SoldToPartyNm
   
    ACol = Frm1.vspdData.ActiveCol
    ARow = Frm1.vspdData.ActiveRow

    If ACol > iColumnLimit Then
		'◎ Frm1없으면 frm1삭제 
		Frm1.vspdData.Col = iColumnLimit	:	Frm1.vspdData.Row = 0
		iRet = DisplayMsgBox("900030", "X", Trim(frm1.vspdData.Text), "X")
		Exit Function
    End If   

    Frm1.vspdData.ScrollBars = Parent.SS_SCROLLBAR_NONE
    
    ggoSpread.Source = Frm1.vspdData
    
    ggoSpread.SSSetSplit(ACol)    
    
    Frm1.vspdData.Col = ACol
    Frm1.vspdData.Row = ARow
    
    Frm1.vspdData.Action = 0    
    
    Frm1.vspdData.ScrollBars = Parent.SS_SCROLLBAR_BOTH
    
End Function

'========================================
Function FncExit()
    FncExit = True
End Function

'========================================
Function DbQuery() 

	Err.Clear														'☜: Protect system from crashing
	DbQuery = False													'⊙: Processing is NG
	
	If LayerShowHide(1) = False Then
		Exit Function
	End If
	
	Dim strVal
	
    With frm1
		strVal = BIZ_PGM_ID & "?txtHMode=" & Parent.UID_M0001					
		If lgIntFlgMode = Parent.OPMD_UMODE Then
			' Scroll시 
			strVal = strVal & "&txtBillType=" & Trim(.txtHBillType.value)
			strVal = strVal & "&txtSalesGrp=" & Trim(.txtHSalesGrp.value)
			strVal = strVal & "&txtSoldToParty=" & Trim(.txtHSoldToParty.value)
			strVal = strVal & "&txtFromDt=" & Trim(.txtHFromDt.value)
			strVal = strVal & "&txtToDt=" & Trim(.txtHToDt.value)
			strVal = strVal & "&txtSoNo=" & Trim(.txtHSoNo.value)
			strVal = strVal & "&txtPostfiFlag=" & Trim(.txtHPostfiFlag.value)
		Else
			' 처음 조회시 
			strVal = strVal & "&txtBillType=" & Trim(.txtBillType.value)
			strVal = strVal & "&txtSalesGrp=" & Trim(.txtSalesGrp.value)
			strVal = strVal & "&txtSoldToParty=" & Trim(.txtSoldToParty.value)
			strVal = strVal & "&txtFromDt=" & Trim(.txtFromDt.Text)
			strVal = strVal & "&txtToDt=" & Trim(.txtToDt.Text)
			strVal = strVal & "&txtSoNo=" & Trim(.txtSoNo.value)
			strVal = strVal & "&txtPostfiFlag=" & Trim(.txtPostfiFlag.value)
		End If

        strVal = strVal & "&lgPageNo="		 & lgPageNo					'☜: Next key tag 
		strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")
        strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
		strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))
		lgIntStartRow = .vspdData.MaxRows + 1
                strVal = strVal & "&gBizArea=" & lgBACd 
                strVal = strVal & "&gPlant=" & lgPLCd 
                strVal = strVal & "&gSalesGrp=" & lgSGCd 
                strVal = strVal & "&gSalesOrg=" & lgSOCd 
	End With 
    
	Call RunMyBizASP(MyBizASP, strVal)									
    DbQuery = True    

End Function

'=========================================
Function DbQueryOk()	    												'☆: 조회 성공후 실행로직 
	Call SetToolbar("11000000000111")
	If frm1.vspdData.MaxRows > 0 Then
		frm1.vspdData.Focus
		frm1.vspdData.SelModeSelected = True
		If lgIntFlgMode <> Parent.OPMD_UMODE Then
			frm1.vspdData.Row = 1
			lgIntFlgMode = Parent.OPMD_UMODE
		End If
		Call FormatSpreadCellByCurrency()
	Else
		Call SetFocusToDocument("M")
		frm1.txtFromDt.focus
	End If

End Function

' 화폐별로 Cell Formating을 재설정한다.
Sub FormatSpreadCellByCurrency()
	With frm1
		Call ReFormatSpreadCellByCellByCurrency(.vspdData,lgIntStartRow,.vspdData.MaxRows,GetKeyPos("A",13),GetKeyPos("A",5),"A","Q","X","X")
		Call ReFormatSpreadCellByCellByCurrency(.vspdData,lgIntStartRow,.vspdData.MaxRows,GetKeyPos("A",13),GetKeyPos("A",6),"A","Q","X","X")
		Call ReFormatSpreadCellByCellByCurrency(.vspdData,lgIntStartRow,.vspdData.MaxRows,GetKeyPos("A",13),GetKeyPos("A",7),"A","Q","X","X")
		Call ReFormatSpreadCellByCellByCurrency(.vspdData,lgIntStartRow,.vspdData.MaxRows,GetKeyPos("A",13),GetKeyPos("A",8),"A","Q","X","X")
		Call ReFormatSpreadCellByCellByCurrency2(.vspdData,lgIntStartRow,.vspdData.MaxRows,parent.gCurrency,GetKeyPos("A",9),"A","Q","X","X")
		Call ReFormatSpreadCellByCellByCurrency2(.vspdData,lgIntStartRow,.vspdData.MaxRows,parent.gCurrency,GetKeyPos("A",10),"A","Q","X","X")
		Call ReFormatSpreadCellByCellByCurrency2(.vspdData,lgIntStartRow,.vspdData.MaxRows,parent.gCurrency,GetKeyPos("A",11),"A","Q","X","X")
		Call ReFormatSpreadCellByCellByCurrency2(.vspdData,lgIntStartRow,.vspdData.MaxRows,parent.gCurrency,GetKeyPos("A",12),"A","Q","X","X")
	End With
End Sub

