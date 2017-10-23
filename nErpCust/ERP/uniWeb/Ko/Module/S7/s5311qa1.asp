<%@ LANGUAGE="VBSCRIPT" %>
<%
'**********************************************************************************************
'*  1. Module Name          : 영업 
'*  2. Function Name        : 매출채권관리 
'*  3. Program ID           : S5311QA1
'*  4. Program Name         : 세금계산서현황조회 
'*  5. Program Desc         : 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2002/05/30
'*  8. Modified date(Last)  : 2003/05/27
'*  9. Modifier (First)     : Byun Jee Hyun
'* 10. Modifier (Last)      : Kim Hyungsuk
'*							: Hwang Seongbae
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************
%>
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/IncSvrCcm.inc" --> 
<!-- #Include file="../../inc/incSvrHTML.inc" -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentA.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs">   </SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit                              '☜: indicates that All variables must be declared in advance

' External ASP File
Const BIZ_PGM_ID        = "s5311qb1.asp"
Const BIZ_PGM_JUMP_ID	= "s5311ma1"

' Constant variables 
Const C_MaxKey          = 12                                           '☆: key count of SpreadSheet

Const C_PopBillToParty	= 1
Const C_PopSalesGrp		= 2
Const C_PopTaxBizArea	= 3

' Common variables 
<!-- #Include file="../../inc/lgvariables.inc" -->	

' User-defind Variables
Dim IsOpenPop  

Dim	lgBlnBillToPartyChg
Dim lgBlnSalesGrpChg
Dim	lgBlnTaxBizAreaChg

Dim EndDate

' 시스템 날짜 
EndDate = UniConvDateAToB("<%=GetSvrDate%>", parent.gServerDateFormat, parent.gDateFormat)

'========================================
Function InitVariables()
	lgPageNo         = ""
    lgIntFlgMode     = Parent.OPMD_CMODE                          'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False	                           'Indicates that no value changed
    lgSortKey        = 1   
	
    '---- Coding part--------------------------------------------------------------------
    lgStrPrevKey = ""										'initializes Previous Key

	lgBlnBillToPartyChg = False
	lgBlnSalesGrpChg	= False
	lgBlnTaxBizAreaChg	= False
End Function

'========================================
Sub SetDefaultVal()
	With frm1
		.txtFromDt.Text = UNIGetFirstDay(EndDate, Parent.gDateFormat)
		.txtToDt.Text = EndDate	
		If Parent.gSalesGrp <> "" Then
			.txtSalesGrp.value = Parent.gSalesGrp
			Call GetSalesGrpNm()
		End If

		.txtBillToParty.Focus
	End With
End Sub

'==========================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A( "Q", "S", "NOCOOKIE", "QA") %>
	<% Call LoadBNumericFormatA("Q", "S", "NOCOOKIE", "QA") %>
End Sub

'========================================
Sub InitSpreadSheet()
	Call SetZAdoSpreadSheet("S5311QA1","S","A","V20030301", Parent.C_SORT_DBAGENT, frm1.vspdData, C_MaxKey, "X", "X" )
	Call SetSpreadLock 
'	frm1.vspdData.OperationMode = 3
End Sub

'========================================
Sub SetSpreadLock()
	frm1.vspdData.ReDraw = False
	ggoSpread.SpreadLockWithOddEvenRowColor()
	frm1.vspdData.ReDraw = True
End Sub

'========================================
Function CookiePage()

	On Error Resume Next

	Const CookieSplit = 4877						<%'Cookie Split String : CookiePage Function Use%>
	
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
	frm1.vspdData.Col = lgKeyPos(2)						'매출구분 

	Select Case Choice
	Case 1
		If Trim(frm1.vspdData.Text) = "N" OR frm1.vspdData.Row = 0 Then		<% '매출채권인경우 : N %>
			PgmJump(BIZ_PGM_JUMP_ID1)
		Else
			MsgBox "해당 매출채권번호는 정상매출채권이 아닙니다.", vbInformation, gLogoName
			WriteCookie CookieSplit , ""
			Exit Function
		End If
	Case 2
		If Trim(frm1.vspdData.Text) = "Y" OR frm1.vspdData.Row = 0 Then		<% '예외매출채권인경우 : Y %>
			PgmJump(BIZ_PGM_JUMP_ID2)
		Else
			MsgBox "해당 매출채권번호는 예외매출채권이 아닙니다.", vbInformation, gLogoName
			WriteCookie CookieSplit , ""
			Exit Function
		End If
	End Select

End Function

'========================================
Function OpenTaxBillDtl()
	Dim iArrRet
	Dim iArrParam(7)
	Dim iCalledAspName
	Dim IntRetCD
	
	On Error Resume Next

	If lgIntFlgMode <> Parent.OPMD_UMODE Then
		Call DisplayMsgBox("900002", "X", "X", "X")
		Exit Function
	End IF

	If IsOpenPop = True Then Exit Function

	With frm1.vspdData
		.row = .activerow
		.Col = GetKeyPos("A",1) : iArrParam(0) = .Text		' 세금계산서 관리번호 
		.Col = GetKeyPos("A",2) : iArrParam(1) = .Text		' 세금계산서 번호 
		.Col = GetKeyPos("A",3) : iArrParam(2) = .Text		' 발행처 
		.Col = GetKeyPos("A",4) : iArrParam(3) = .Text		' 발행처명 
		.Col = GetKeyPos("A",5) : iArrParam(4) = .Text		' 발행여부 
		.Col = GetKeyPos("A",6) : iArrParam(5) = .Text		' 공급가액 
		.Col = GetKeyPos("A",7) : iArrParam(6) = .Text		' 화폐 
		.Col = GetKeyPos("A",8) : iArrParam(7) = .Text		' 부가세액 
	End With
	
	IsOpenPop = True
	
	iCalledAspName = AskPRAspName("s5312ra1")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "s5312ra1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	iArrRet = window.showModalDialog(iCalledAspName,Array(window.parent,iArrParam),"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

End Function

'========================================
Function OpenConPopup(ByVal pvIntWhere)

	Dim iArrRet
	Dim iArrParam(5), iArrField(6), iArrHeader(6)

	OpenConPopup = False
	
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	Select Case pvIntWhere
	Case C_PopBillToParty												
		iArrParam(1) = "dbo.b_biz_partner BP"			' TABLE 명칭 
		iArrParam(2) = Trim(frm1.txtBillToParty.value)	' Code Condition
		iArrParam(3) = ""								' Name Cindition
'		iArrParam(4) = "EXISTS (SELECT * FROM dbo.b_biz_partner_ftn BF WHERE BP.bp_cd = BF.partner_bp_cd AND BF.partner_ftn = " & FilterVar("SBI", "''", "S") & ") " & _
'					   "AND BP.usage_flag = " & FilterVar("Y", "''", "S") & " "		' Where Condition
		iArrParam(4) = "EXISTS (SELECT * FROM dbo.b_biz_partner_ftn BF WHERE BP.bp_cd = BF.partner_bp_cd AND BF.partner_ftn = " & FilterVar("SBI", "''", "S") & ") " & _
				   ""		' Where Condition
		iArrParam(5) = "발행처"						' TextBox 명칭 
			
		iArrField(0) = "ED15" & Parent.gColSep & "BP.bp_cd"	' Field명(0)
		iArrField(1) = "ED30" & Parent.gColSep & "BP.bp_nm"	' Field명(1)
		    
		iArrHeader(0) = "발행처"					' Header명(0)
		iArrHeader(1) = "발행처명"					' Header명(1)
		
		frm1.txtBillToParty.focus

	Case C_PopSalesGrp												
		iArrParam(1) = "dbo.B_SALES_GRP"
		iArrParam(2) = Trim(frm1.txtSalesGrp.value)
		iArrParam(3) = ""
		iArrParam(4) = "USAGE_FLAG=" & FilterVar("Y", "''", "S") & " "
		iArrParam(5) = "영업그룹"
		
		iArrField(0) = "ED15" & Parent.gColSep & "SALES_GRP"
		iArrField(1) = "ED30" & Parent.gColSep & "SALES_GRP_NM"
    
	    iArrHeader(0) = "영업그룹"
	    iArrHeader(1) = "영업그룹명"
	    
	    frm1.txtSalesGrp.focus

	Case C_PopTaxBizArea
		iArrParam(0) = "세금신고사업장"					
		iArrParam(1) = "dbo.B_TAX_BIZ_AREA"
		iArrParam(2) = Trim(frm1.txtTaxBizArea.value)
		iArrParam(3) = ""
		iArrParam(4) = ""
		iArrParam(5) = "세금신고사업장"							

		iArrField(0) = "ED15" & Parent.gColSep & "TAX_BIZ_AREA_CD"
		iArrField(1) = "ED30" & Parent.gColSep & "TAX_BIZ_AREA_NM"

		iArrHeader(0) = "세금신고사업장"							
		iArrHeader(1) = "세금신고사업장명"							

		frm1.txtTaxBizArea.focus
	End Select
 
	iArrParam(0) = iArrParam(5)							' 팝업 명칭 

	iArrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(iArrParam, iArrField, iArrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	
	If iArrRet(0) <> "" Then OpenConPopup = SetConPopup(iArrRet,pvIntWhere)
	
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

'========================================
Function SetConPopup(Byval pvArrRet,ByVal pvIntWhere)

	SetConPopup = False

	Select Case pvIntWhere
	Case C_PopTaxbizArea
		frm1.txtTaxBizArea.value = pvArrRet(0) 
		frm1.txtTaxBizAreaNm.value = pvArrRet(1)   
	Case C_PopBillToParty
		frm1.txtBillToParty.value = pvArrRet(0) 
		frm1.txtBillToPartyNm.value = pvArrRet(1)   
	Case C_PopSalesGrp
		frm1.txtSalesGrp.value = pvArrRet(0) 
		frm1.txtSalesGrpNm.value = pvArrRet(1)   
	End Select

	SetConPopup = True

End Function

'========================================
Sub Form_Load()
    Call LoadInfTB19029											  '⊙: Load table , B_numeric_format
  
    'Html에서 tag 숫자가 1과 2로 시작하는 부분 각각Format
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
	
	Call ggoOper.LockField(Document, "N")                         '⊙: Lock  Suitable  Field
  
  	Call InitVariables											  '⊙: Initializes local global variables
	Call SetDefaultVal	
	Call InitSpreadSheet()
	
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
Function txtBillToParty_OnKeyDown()
	lgBlnBillToPartyChg = True
	lgBlnFlgChgValue = True
End Function

'========================================
Function txtSalesGrp_OnKeyDown()
	lgBlnSalesGrpChg = True
	lgBlnFlgChgValue = True
End Function

'========================================
Function txtTaxBizArea_OnKeyDown()
	lgBlnTaxBizAreaChg = True
	lgBlnFlgChgValue = True
End Function

'	Description : 조회조건의 유효성을 Check한다.
'   주의사항 : 화면의 tab order 별로 기술한다. 
'========================================
Function ChkValidityQueryCon()
	Dim iStrCode

	ChkValidityQueryCon = True

	If lgBlnBillToPartyChg Then

		iStrCode = Trim(frm1.txtBillToParty.value)
		If iStrCode <> "" Then
			iStrCode = " " & FilterVar(iStrCode, "''", "S") & ""
			If Not GetCodeName(iStrCode, "" & FilterVar("SBI", "''", "S") & "", "default", "default", "default", "" & FilterVar("BF", "''", "S") & "", C_PopBillToParty) Then
				Call DisplayMsgBox("970000", "X", frm1.txtBilltoparty.alt, "X")
				frm1.txtBillToPartyNm.value = ""
				ChkValidityQueryCon = False
				frm1.txtBillToParty.focus
			End If
		Else
			frm1.txtBillToPartyNm.value = ""
		End If
		lgBlnBillToPartyChg	= False
	End If

	If lgBlnTaxBizAreaChg Then
		iStrCode = Trim(frm1.txtTaxBizArea.value)
		If iStrCode <> "" Then
			iStrCode = " " & FilterVar(iStrCode, "''", "S") & ""
			If Not GetTaxBizArea("NM") Then
				If ChkValidityQueryCon = True Then
					Call DisplayMsgBox("970000", "X", frm1.txtTaxBizArea.alt, "X")
					ChkValidityQueryCon = False
					frm1.txtTaxBizArea.focus
				End If
				frm1.txtTaxBizAreaNm.value = ""
			End If
		Else
			frm1.txtTaxBizAreaNm.value = ""
		End If
		lgBlnTaxBizAreaChg = False
	End If

	If lgBlnSalesGrpChg Then
		
		iStrCode = Trim(frm1.txtSalesGrp.value)
		If iStrCode <> "" Then
			iStrCode = " " & FilterVar(iStrCode, "''", "S") & ""
			If Not GetCodeName(iStrCode, "default", "default", "default", "default", "" & FilterVar("SG", "''", "S") & "", C_PopSalesGrp) Then
				If ChkValidityQueryCon = True Then
					Call DisplayMsgBox("970000", "X", frm1.txtSalesGrp.alt, "X")
					ChkValidityQueryCon = False
					frm1.txtSalesGrp.focus
				End If
				frm1.txtSalesGrpNm.value = ""
			End If
		Else
			frm1.txtSalesGrpNm.value = ""
		End If
		lgBlnSalesGrpChg = False
	End If
	 
End Function

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
	End if
End Function

'====================================================
Function GetTaxBizArea(Byval pvStrFlag)

	Dim iStrSelectList, iStrFromList, iStrWhereList
	Dim iStrRs
	Dim iArrTaxBizArea(1), iArrTemp
	
	GetTaxBizArea = False
	
	iStrSelectList = " * "
	iStrFromList = " dbo.ufn_s_GetTaxBizArea ('', '',  " & FilterVar(frm1.txtTaxBizArea.value, "''", "S") & ",  " & FilterVar(pvStrFlag, "''", "S") & ") "
	iStrWhereList = ""
	
	Err.Clear
    
	If CommonQueryRs2by2(iStrSelectList,iStrFromList,iStrWhereList,iStrRs) Then
		iArrTemp = Split(iStrRs, Chr(11))
		iArrTaxBizArea(0) = iArrTemp(1)
		iArrTaxBizArea(1) = iArrTemp(2)
		GetTaxBizArea = SetConPopup(iArrTaxBizArea, C_PopTaxBizArea)
	Else
		If Err.number <> 0 Then	Err.Clear 

		 '세금 신고 사업장을 Editing한 경우 
		'GetTaxBizArea = OpenConPopup(C_PopTaxBizArea)
	End if
End Function

'==========================================
Sub rdoTaxTypeFlg1_OnClick()
	frm1.txtTaxRadio.value = frm1.rdoTaxTypeFlg1.value
End Sub

'==========================================
Sub rdoTaxTypeFlg2_OnClick()
	frm1.txtTaxRadio.value = frm1.rdoTaxTypeFlg2.value
End Sub

'==========================================
Sub rdoTaxTypeFlg3_OnClick()
	frm1.txtTaxRadio.value = frm1.rdoTaxTypeFlg3.value
End Sub
	
'==========================================
Sub rdoTexIssueFlg1_OnClick()
	frm1.txtIssueRadio.value = frm1.rdoTexIssueFlg1.value
End Sub

'==========================================
Sub rdoTexIssueFlg2_OnClick()
	frm1.txtIssueRadio.value = frm1.rdoTexIssueFlg2.value
End Sub

'==========================================
Sub rdoTexIssueFlg3_OnClick()
	frm1.txtIssueRadio.value = frm1.rdoTexIssueFlg3.value
End Sub

'========================================
Function vspdData_DblClick(ByVal Col, ByVal Row)
	If frm1.vspdData.ActiveRow > 0 Then Call OpenTaxBillDtl
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
    
	frm1.vspdData.ReDraw = True
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
		If lgPageNo <> "" Then Call DbQuery
	End If

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
Sub vspdData_Keypress(KeyAscii)
	If KeyAscii = 13	Then Call MainQuery()
End Sub

'========================================
Function FncQuery() 
    
    FncQuery = False                                                        '⊙: Processing is NG
    
    Err.Clear                                                               '☜: Protect system from crashing
    
	If Not ValidDateCheck(frm1.txtFromDt, frm1.txtToDt) Then Exit Function
   
	' 조회조건 유효값 check
	If 	lgBlnFlgChgValue Then
		If Not ChkValidityQueryCon Then	Exit Function
	End If
	
    Call ggoOper.ClearField(Document, "2")	         						'⊙: Clear Contents  Field

    Call InitVariables 														'⊙: Initializes local global variables
    
	If Not DbQuery Then Exit Function									

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
    Call parent.FncFind(Parent.C_MULTI , False)                                     <%'☜:화면 유형, Tab 유무 %>
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
	
	If LayerShowHide(1) = False Then Exit Function
	
	Dim strVal
	
    With frm1
		strVal = BIZ_PGM_ID & "?txtHMode=" & Parent.UID_M0001					<%'☜: 비지니스 처리 ASP의 상태 %>
		If lgIntFlgMode = Parent.OPMD_UMODE Then
			' Scroll시 
			strVal = strVal & "&txtBillToParty=" & Trim(.txtHBillToParty.value)
			strVal = strVal & "&txtSalesGrp=" & Trim(.txtHSalesGrp.value)
			strVal = strVal & "&txtTaxBizArea=" & Trim(.txtHTaxBizArea.value)
			strVal = strVal & "&txtFromDt=" & Trim(.txtHFromDt.value)
			strVal = strVal & "&txtToDt=" & Trim(.txtHToDt.value)
			strVal = strVal & "&txtIssueRadio=" & Trim(.txtHIssueRadio.value)
			strVal = strVal & "&txtTaxRadio=" & Trim(.txtHTaxRadio.value)	
			
		Else
			' 처음 조회시 
			strVal = strVal & "&txtBillToParty=" & Trim(.txtBillToParty.value)
			strVal = strVal & "&txtTaxBizArea=" & Trim(.txtTaxBizArea.value)
			strVal = strVal & "&txtSalesGrp=" & Trim(.txtSalesGrp.value)
			strVal = strVal & "&txtTaxRadio=" & Trim(.txtTaxRadio.value)
			strVal = strVal & "&txtIssueRadio=" & Trim(.txtIssueRadio.value)
			strVal = strVal & "&txtFromDt=" & Trim(.txtFromDt.Text)
			strVal = strVal & "&txtToDt=" & Trim(.txtToDt.Text)
		End If

        strVal = strVal & "&lgPageNo="		 & lgPageNo					'☜: Next key tag 
		strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")
        strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
		strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))
	End With    
    
	Call RunMyBizASP(MyBizASP, strVal)									<%'☜: 비지니스 ASP 를 가동 %>
    DbQuery = True    

End Function

'=========================================
Function DbQueryOk()	    												'☆: 조회 성공후 실행로직 

	If frm1.vspdData.MaxRows > 0 Then
		frm1.vspdData.Focus
		frm1.vspdData.SelModeSelected = True
		If lgIntFlgMode <> Parent.OPMD_UMODE Then
			frm1.vspdData.Row = 1
'			frm1.vspdData.OperationMode = 3
		End If
		lgIntFlgMode = Parent.OPMD_UMODE
		Call SetToolbar("11000000000111") '2005/09/29 박정순 수정 
	Else
		frm1.txtBillToParty.focus
	End If

End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>
<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">

<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR >
		<TD <%=HEIGHT_TYPE_00%>></TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>세금계산서현황</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* align=right><A href="vbscript:OpenTaxBillDtl()">세금계산서내역</A></TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=*>
		<TD WIDTH=100% CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS=TD5>발행처</TD>
									<TD CLASS=TD6>
										<INPUT TYPE=TEXT NAME="txtBillToParty" SIZE=10 MAXLENGTH=10 TAG="11XXXU" ALT="발행처"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBillToParty" align=top TYPE="BUTTON" ONCLICK ="vbscript:OpenConPopup C_PopBillToParty">&nbsp;
										<INPUT TYPE=TEXT NAME="txtBillToPartyNm" SIZE=20 TAG="14">
									</TD>
									<TD CLASS=TD5>발행일</TD>
									<TD CLASS=TD6>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 NAME="txtFromDt" CLASS="FPDTYYYYMMDD" tag="11X1" Title="FPDATETIME" ALT="시작일"></OBJECT>');</SCRIPT>&nbsp;~&nbsp;
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime2 NAME="txtToDt" CLASS="FPDTYYYYMMDD" tag="11X1" Title="FPDATETIME" ALT="종료일"></OBJECT>');</SCRIPT>
									</TD>	
								</TR>
								<TR>
									<TD CLASS=TD5>세금신고사업장</TD>
									<TD CLASS=TD6>
										<INPUT TYPE=TEXT NAME="txtTaxBizArea" SIZE=10 MAXLENGTH=10 TAG="11XXXU" ALT="세금신고사업장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnTaxBizArea" align=top TYPE="BUTTON" ONCLICK ="vbscript:OpenConPopup C_PopTaxBizArea">&nbsp;
										<INPUT TYPE=TEXT NAME="txtTaxBizAreaNm" SIZE=20 TAG="14">
									</TD>
									<TD CLASS=TD5 NOWRAP>영업그룹</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtSalesGrp" SIZE=10 MAXLENGTH=4 TAG="11XXXU" ALT="영업그룹"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSalesGrp" align=top TYPE="BUTTON" ONCLICK ="vbscript:OpenConPopup C_PopSalesGrp">&nbsp;<INPUT TYPE=TEXT NAME="txtSalesGrpNm" SIZE=20 TAG="14"></TD>
								</TR>	
								<TR>
									<TD CLASS=TD5 NOWRAP>계산서형태</TD> 
									<TD CLASS=TD6 NOWRAP>
										<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoTaxTypeFlg" TAG="11X" VALUE="" CHECKED ID="rdoTaxTypeFlg1"><LABEL FOR="rdoTaxTypeFlg1">전체</LABEL>&nbsp;&nbsp;&nbsp;
										<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoTaxTypeFlg" TAG="11X" VALUE="R" ID="rdoTaxTypeFlg2"><LABEL FOR="rdoTaxTypeFlg2">영수</LABEL>&nbsp;&nbsp;&nbsp;
										<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoTaxTypeFlg" TAG="11X" VALUE="D" ID="rdoTaxTypeFlg3"><LABEL FOR="rdoTaxTypeFlg3">청구</LABEL>			
									</TD>
									<TD CLASS=TD5 NOWRAP>발행여부</TD> 
									<TD CLASS=TD6 NOWRAP>
										<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoTexIssueFlg" TAG="11X" VALUE="" CHECKED ID="rdoTexIssueFlg1"><LABEL FOR="rdoTexIssueFlg1">전체</LABEL>&nbsp;&nbsp;&nbsp;
										<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoTexIssueFlg" TAG="11X" VALUE="Y" ID="rdoTexIssueFlg2"><LABEL FOR="rdoTexIssueFlg2">발행</LABEL>&nbsp;&nbsp;&nbsp;
										<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoTexIssueFlg" TAG="11X" VALUE="N" ID="rdoTexIssueFlg3"><LABEL FOR="rdoTexIssueFlg3">미발행</LABEL>			
									</TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=100% valign=top>
							<TABLE <%=LR_SPACE_TYPE_20%>>
								<TR>
									<TD HEIGHT="100%">
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=vaSpread1> <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
									</TD>
								</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_01%> WIDTH=100%></TD>
	</TR>  
	<TR HEIGHT=20>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_30%>>
			    <TR>
					<TD WIDTH="*" ALIGN=RIGHT><a href = "vbscript:PgmJump(BIZ_PGM_JUMP_ID)" ONCLICK="VBSCRIPT:CookiePage">세금계산서등록</a></TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>

<INPUT TYPE=HIDDEN NAME="txtTaxRadio" tag="14" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtIssueRadio" tag="14" TABINDEX="-1">

<INPUT TYPE=HIDDEN NAME="txtHBillToParty" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHTaxBizArea" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHSalesGrp" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHTaxRadio" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHIssueRadio" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHFromDt" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHToDt" tag="24" TABINDEX="-1">
</FORM>

<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm" TABINDEX="-1"></iframe>
</DIV>

</BODY>
</HTML>
