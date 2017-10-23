<%@ LANGUAGE="VBSCRIPT" %>

<%
'**********************************************************************************************
'*  1. Module Name          : 영업 
'*  2. Function Name        : 매출채권관리 
'*  3. Program ID           : S5311PA1
'*  4. Program Name         : 세금계산서관리번호 팝업 
'*  5. Program Desc         : 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2000/12/09
'*  8. Modified date(Last)  : 2003/05/27
'*  9. Modifier (First)     : Byun Jee Hyun
'* 10. Modifier (Last)      : Kim Hyungsuk
'*							: Hwang Seongbae
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                            2000/12/09
'*                            2001/12/18	Date표준적용 
'*                            2002/06/04	ADO 표준변경 및 Default 기간 변경 
'*                            2003/05/27	표준반영 
'**********************************************************************************************
%>
<HTML>
<HEAD>
<TITLE>세금계산서관리번호</TITLE>

<!-- #Include file="../../inc/IncSvrCcm.inc" --> 
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentA.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript" SRC="../../inc/incImage.js"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMA_KO441.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit                              '☜: indicates that All variables must be declared in advance
<!-- #Include file="../../inc/lgvariables.inc" -->

' External ASP File
Const BIZ_PGM_ID 		= "s5311pb1_ko441.asp"                              '☆: Biz Logic ASP Name

' 1.2 Constant variables 
Const C_MaxKey          = 4                                           '☆: key count of SpreadSheet
Const C_PopBillToParty	= 1
Const C_PopSalesGrp		= 2
Const C_PopVatType		= 3
Const C_PopTaxBizArea	= 4

' User-defind Variables
Dim IsOpenPop  

Dim lgBlnOpenedFlag
Dim	lgBlnBillToPartyChg
Dim lgBlnSalesGrpChg
Dim	lgBlnVatTypeChg
Dim	lgBlnTaxBizAreaChg

Dim arrParent
Dim PopupParent

arrParent = window.dialogArguments
Set PopupParent  = arrParent(0)
top.document.title = PopupParent.gActivePRAspName

Dim EndDate

' 시스템 날짜 
EndDate = UniConvDateAToB("<%=GetSvrDate%>", PopupParent.gServerDateFormat, PopupParent.gDateFormat)

'========================================
Function InitVariables()
	lgPageNo         = ""
    lgBlnFlgChgValue = False											'Indicates that no value changed
    lgIntFlgMode     = PopupParent.OPMD_CMODE                           'Indicates that current mode is Create mode
    lgSortKey        = 1   
        
	lgBlnBillToPartyChg = False
	lgBlnSalesGrpChg	= False
	lgBlnVatTypeChg		= False
	lgBlnTaxBizAreaChg	= False
End Function

'========================================
Sub SetDefaultVal()
	With frm1
		.txtFromDt.Text = UNIGetFirstDay(EndDate, PopupParent.gDateFormat)
		.txtToDt.Text = EndDate

		If PopupParent.gSalesGrp <> "" Then
			.txtSalesGrp.value = PopupParent.gSalesGrp
			Call GetSalesGrpNm()
		End If
	End With
	Self.Returnvalue = ""
	If lgSGCd <> "" then 
		Call ggoOper.SetReqAttr(frm1.txtSalesGrp, "Q") 
        	frm1.txtSalesGrp.value = lgSGCd
	End If	

End Sub

'========================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A( "Q", "S", "NOCOOKIE", "PA") %>
	<% Call LoadBNumericFormatA("Q", "S", "NOCOOKIE", "PA") %>
End Sub

'========================================
Sub InitSpreadSheet()
	Call SetZAdoSpreadSheet("s5311pa1","S","A","V20030320", PopupParent.C_SORT_DBAGENT, frm1.vspdData, C_MaxKey, "X", "X" )
	Call SetSpreadLock 	
	frm1.vspddata.OperationMode = 3
End Sub

'========================================
Sub SetSpreadLock()
'	ggoSpread.SpreadLock 1 , -1
    ggoSpread.Source = frm1.vspdData
    ggoSpread.SpreadLockWithOddEvenRowColor()
End Sub

'========================================
Function OKClick()
	Dim iStrRtnVal
	
	With frm1.vspdData
		If .ActiveRow > 0 Then
			Call .Gettext(GetKeyPos("A",1), .ActiveRow, iStrRtnVal)
			Self.Returnvalue = iStrRtnVal
		End If
	End With
	Self.Close()
End Function

'========================================
Function CancelClick()
	Self.Close()
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
		iArrParam(4) = "EXISTS (SELECT * FROM dbo.b_biz_partner_ftn BF WHERE BP.bp_cd = BF.partner_bp_cd AND BF.partner_ftn = " & FilterVar("SBI", "''", "S") & ") " & _
					   "AND BP.usage_flag = " & FilterVar("Y", "''", "S") & " "		' Where Condition
		iArrParam(5) = "발행처"						' TextBox 명칭 
			
		iArrField(0) = "ED15" & PopupParent.gColSep & "BP.bp_cd"	' Field명(0)
		iArrField(1) = "ED30" & PopupParent.gColSep & "BP.bp_nm"	' Field명(1)
		    
		iArrHeader(0) = "발행처"					' Header명(0)
		iArrHeader(1) = "발행처명"					' Header명(1)
		
		frm1.txtBillToParty.focus

	Case C_PopSalesGrp
                If frm1.txtSalesGrp.className = "protected" Then
                	IsOpenPop = False
                        Exit Function
                End If 													
		iArrParam(1) = "dbo.B_SALES_GRP"
		iArrParam(2) = Trim(frm1.txtSalesGrp.value)
		iArrParam(3) = ""
		iArrParam(4) = "USAGE_FLAG=" & FilterVar("Y", "''", "S") & " "
		iArrParam(5) = "영업그룹"
		
		iArrField(0) = "ED15" & PopupParent.gColSep & "SALES_GRP"
		iArrField(1) = "ED30" & PopupParent.gColSep & "SALES_GRP_NM"
    
	    iArrHeader(0) = "영업그룹"
	    iArrHeader(1) = "영업그룹명"
	    
	    frm1.txtSalesGrp.focus

	Case C_PopVatType
		iArrParam(1) = "dbo.b_minor A"
		iArrParam(2) = Trim(frm1.txtVatType.value)
		iArrParam(3) = ""
		iArrParam(4) = "A.major_cd = " & FilterVar("B9001", "''", "S") & ""
		iArrParam(5) = "VAT유형"
	
		iArrField(0) = "ED15" & PopupParent.gColSep & "A.minor_cd"
		iArrField(1) = "ED30" & PopupParent.gColSep & "A.minor_nm"
		    
		iArrHeader(0) = "VAT유형"
		iArrHeader(1) = "VAT유형명"
		
		frm1.txtVATType.focus

	Case C_PopTaxBizArea
		iArrParam(0) = "세금신고사업장"					
		iArrParam(1) = "dbo.b_tax_biz_area"
		iArrParam(2) = Trim(frm1.txtTaxBizArea.value)
		iArrParam(3) = ""
		iArrParam(4) = ""
		iArrParam(5) = "세금신고사업장"							

		iArrField(0) = "ED15" & PopupParent.gColSep & "TAX_BIZ_AREA_CD"
		iArrField(1) = "ED30" & PopupParent.gColSep & "TAX_BIZ_AREA_NM"

		iArrHeader(0) = "세금신고사업장"							
		iArrHeader(1) = "세금신고사업장명"							

		frm1.txtTaxBizArea.focus
	End Select
 
	iArrParam(0) = iArrParam(5)							<%' 팝업 명칭 %> 

	iArrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(iArrParam, iArrField, iArrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If iArrRet(0) <> "" Then OpenConPopup = SetConPopup(iArrRet,pvIntWhere)
	
End Function

'========================================
Function OpenSortPopup()
	Dim arrRet
	
	On Error Resume Next 
	
	If IsOpenPop = True Then Exit Function
	IsOpenPop = True

	arrRet = window.showModalDialog("../../ComAsp/ZADOSortPopup.asp",Array(ggoSpread.GetXMLData("A"),gMethodText),"dialogWidth=" & PopupParent.SORTW_WIDTH & "px; dialogHeight=" & PopupParent.SORTW_HEIGHT & "px;; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	
	If arrRet(0) = "X" Then
	   Exit Function
	Else
	   Call ggoSpread.SaveXMLData("A",arrRet(0),arrRet(1))
       Call InitVariables
       Call InitSpreadSheet()       
   End If
End Function

'=======================================================
Function SetConPopup(Byval pvArrRet,ByVal pvIntWhere)

	SetConPopup = False

	Select Case pvIntWhere
	Case C_PopBillToParty
		frm1.txtBillToParty.value = pvArrRet(0) 
		frm1.txtBillToPartyNm.value = pvArrRet(1)   
	Case C_PopSalesGrp
		frm1.txtSalesGrp.value = pvArrRet(0) 
		frm1.txtSalesGrpNm.value = pvArrRet(1)   
	Case C_PopVatType
		frm1.txtVatType.value = pvArrRet(0) 
		frm1.txtVatTypeNm.value = pvArrRet(1)   
	Case C_PopTaxBizArea
		frm1.txtTaxBizArea.value = pvArrRet(0) 
		frm1.txtTaxBizAreaNm.value = pvArrRet(1)   
	End Select

	SetConPopup = True

End Function

'========================================
Sub Form_Load()
    Call LoadInfTB19029											  '⊙: Load table , B_numeric_format
    
    'Html에서 tag 숫자가 1과 2로 시작하는 부분 각각Format
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
	
	Call ggoOper.LockField(Document, "N")                         '⊙: Lock  Suitable  Field
    
	Call InitVariables
        Call GetValue_ko441()											  '⊙: Initializes local global variables
	Call SetDefaultVal	
	Call InitSpreadSheet()
	Call MM_preloadImages("../../../CShared/image/Query.gif","../../../CShared/image/OK.gif","../../../CShared/image/Cancel.gif")
	lgBlnOpenedflag = True
	DbQuery()

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
Function txtVatType_OnKeyDown()
	lgBlnVATTypeChg = True
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
				frm1.txtBilltoparty.focus
				ChkValidityQueryCon = False
				Exit Function
			End If
		Else
			frm1.txtBillToPartyNm.value = ""
		End If
		lgBlnBillToPartyChg	= False
	End If

	If lgBlnSalesGrpChg Then
		iStrCode = Trim(frm1.txtSalesGrp.value)
		If iStrCode <> "" Then
			iStrCode = " " & FilterVar(iStrCode, "''", "S") & ""
			If Not GetCodeName(iStrCode, "default", "default", "default", "default", "" & FilterVar("SG", "''", "S") & "", C_PopSalesGrp) Then
				Call DisplayMsgBox("970000", "X", frm1.txtSalesGrp.alt, "X")
				frm1.txtSalesGrp.focus
				ChkValidityQueryCon = False
				Exit Function
			End If
		Else
			frm1.txtSalesGrpNm.value = ""
		End If
		lgBlnSalesGrpChg = False
	End If
			
	If lgBlnVatTypeChg Then
		iStrCode = Trim(frm1.txtVatType.value)
		If iStrCode <> "" Then
			iStrCode = " " & FilterVar(iStrCode, "''", "S") & ""
			If Not GetCodeName(iStrCode, "" & FilterVar("B9001", "''", "S") & "", "default", "default", "default", "" & FilterVar("MJ", "''", "S") & "", C_PopVatType) Then
				Call DisplayMsgBox("970000", "X", frm1.txtVatType.alt, "X")
				frm1.txtVatType.focus
				ChkValidityQueryCon = False
				Exit Function
			End If
		Else
			frm1.txtVatTypeNm.value = ""
		End If
		lgBlnVatTypeChg = False
	End If

	If lgBlnTaxBizAreaChg Then
		iStrCode = Trim(frm1.txtTaxBizArea.value)
		If iStrCode <> "" Then
			iStrCode = " " & FilterVar(iStrCode, "''", "S") & ""
			If Not GetTaxBizArea("NM") Then
				Call DisplayMsgBox("970000", "X", frm1.txtTaxBizArea.alt, "X")
				frm1.txtTaxBizArea.focus
				ChkValidityQueryCon = False
				Exit Function
			End If
		Else
			frm1.txtTaxBizAreaNm.value = ""
		End If
		lgBlnTaxBizAreaChg = False
	End If
End Function

'	Description : 코드값에 해당하는 명을 Display한다.
'========================================
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
		'If lgBlnOpenedFlag Then	GetCodeName = OpenConPopup(pvIntWhere)
	End if
End Function

'	Description : 세금신고 사업장 정보를 Fetch한다.
'========================================
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

		' 세금 신고 사업장을 Editing한 경우 
		'GetTaxBizArea = OpenConPopup(C_PopTaxBizArea)
	End if
End Function

'========================================
Function vspdData_DblClick(ByVal Col, ByVal Row)

	If Row = 0 Or Frm1.vspdData.MaxRows = 0 Then 
          Exit Function
    End If

	If frm1.vspdData.MaxRows > 0 Then
		If frm1.vspdData.ActiveRow = Row Or frm1.vspdData.ActiveRow > 0 Then
			Call OKClick
		End If
	End If
End Function

'========================================
Function vspdData_KeyPress(KeyAscii)
	On Error Resume Next
	If KeyAscii = 13 And frm1.vspdData.ActiveRow > 0 Then
		Call OKClick()
	ElseIf KeyAscii = 27 Then
		Call CancelClick()
	End If
End Function

'========================================
Sub vspdData_TopLeftChange(ByVal OldLeft, ByVal OldTop, ByVal NewLeft, ByVal NewTop)
	If OldLeft <> NewLeft Then Exit Sub

	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	    '☜: 재쿼리 체크	
		If lgPageNo <> "" Then	Call DbQuery
	End If
End Sub

'========================================
Sub txtFromDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtFromDt.Action = 7		
		Call SetFocusToDocument("P")
		frm1.txtFromDt.Focus
	End If
End Sub

'========================================
Sub txtToDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtToDt.Action = 7
		Call SetFocusToDocument("P")
		frm1.txtToDt.Focus
	End If
End Sub

'========================================
Sub txtFromDt_Keypress(KeyAscii)
	On Error Resume Next
	If KeyAscii = 27 Then
		Call CancelClick()
	Elseif KeyAscii = 13 Then
		Call FncQuery()
	End if
End Sub

'========================================
Sub txtToDt_Keypress(KeyAscii)
	On Error Resume Next
	If KeyAscii = 27 Then
		Call CancelClick()
	Elseif KeyAscii = 13 Then
		Call FncQuery()
	End if
End Sub

'========================================
Function FncQuery() 
    
    FncQuery = False                                                        '⊙: Processing is NG
    
    Err.Clear                                                               '☜: Protect system from crashing
	With frm1
		If ValidDateCheck(.txtFromDt, .txtToDt) = False Then Exit Function

		If UniConvDateToYYYYMMDD(.txtFromDt.text , PopupParent.gDateFormat , "") > UniConvDateToYYYYMMDD(EndDate, PopupParent.gDateFormat , "") Then		
			Call DisplayMsgBox("970025", "X", .txtFromDt.ALT, "현재일" & "(" & EndDate & ")")	
			.txtToDt.Focus()
			Exit Function
		End If

'		If UniConvDateToYYYYMMDD(.txtToDt.text , PopupParent.gDateFormat , "") > UniConvDateToYYYYMMDD(EndDate, PopupParent.gDateFormat , "") Then		
'			Call DisplayMsgBox("970025", "X", .txtToDt.ALT, "현재일" & "(" & EndDate & ")")	
'			.txtToDt.Focus()
'			Exit Function
'		End If
	End With
   
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
Function DbQuery() 

	Err.Clear														'☜: Protect system from crashing
	DbQuery = False													'⊙: Processing is NG
	
	If LayerShowHide(1) = False Then Exit Function
	
	Dim strVal
	
    With frm1
		strVal = BIZ_PGM_ID & "?txtHMode=" & PopupParent.UID_M0001					<%'☜: 비지니스 처리 ASP의 상태 %>
		If lgIntFlgMode = PopupParent.OPMD_UMODE Then
			' Scroll시 
			strVal = strVal & "&txtBillToParty=" & .txtHBillToParty.value
			strVal = strVal & "&txtFromDt=" & .txtHFromDt.value
			strVal = strVal & "&txtToDt=" & .txtHToDt.value
			strVal = strVal & "&txtSalesGrp=" & .txtHSalesGrp.value
			strVal = strVal & "&txtVatType=" & .txtHVatType.value
			strVal = strVal & "&txtTaxBizArea=" & .txtHTaxBizArea.value
			strVal = strVal & "&txtPostFlag=" & .txtHPostFlag.value
		Else
			strVal = strVal & "&txtBillToParty=" & Trim(.txtBillToParty.value)
			' 처음 조회시 
			strVal = strVal & "&txtFromDt=" & Trim(.txtFromDt.Text)				<%'☆: 조회 조건 데이타 %>
			If Len(Trim(.txtToDt.text)) Then
				strVal = strVal & "&txtToDt=" & Trim(.txtToDt.Text)
'			Else
'				strVal = strVal & "&txtToDt=" & EndDate
			End if
			strVal = strVal & "&txtSalesGrp=" & Trim(.txtSalesGrp.value)
			strVal = strVal & "&txtVatType=" & Trim(.txtVatType.value)
			strVal = strVal & "&txtTaxBizArea=" & Trim(.txtTaxBizArea.value)
			
			If .rdoTaxIssueFlg2.checked = True Then
				strVal = strVal & "&txtPostFlag=Y"
			ElseIf frm1.rdoTaxIssueFlg3.checked = True Then
				strVal = strVal & "&txtPostFlag=N"
			Else
				strVal = strVal & "&txtPostFlag="
			End If			   	
		End If
		
        strVal = strVal & "&lgPageNo="		 & lgPageNo						'☜: Next key tag 
		strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")
		
        strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
		strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))
	End With 

                strVal = strVal & "&gBizArea=" & lgBACd 
                strVal = strVal & "&gPlant=" & lgPLCd 
                strVal = strVal & "&gSalesGrp=" & lgSGCd 
                strVal = strVal & "&gSalesOrg=" & lgSOCd  
   
	Call RunMyBizASP(MyBizASP, strVal)									<%'☜: 비지니스 ASP 를 가동 %>
    DbQuery = True    

End Function

'=========================================
Function DbQueryOk()	    												'☆: 조회 성공후 실행로직 

	If frm1.vspdData.MaxRows > 0 Then
		lgIntFlgMode = PopupParent.OPMD_UMODE
		frm1.vspdData.Focus
		frm1.vspdData.Row = 1	
		frm1.vspdData.SelModeSelected = True		
	Else
		frm1.txtBillToParty.focus
	End If

End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>
<%
'#########################################################################################################
'       					6. Tag부 
'#########################################################################################################
 %>
<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
	<TABLE <%=LR_SPACE_TYPE_20%>>
		<TR>
			<TD <%=HEIGHT_TYPE_02%> WIDTH=100%>
				<FIELDSET CLASS="CLSFLD">
					<TABLE <%=LR_SPACE_TYPE_40%>>
						<TR>
							<TD CLASS=TD5>발행처</TD>
							<TD CLASS=TD6>
								<INPUT TYPE=TEXT NAME="txtBillToParty" SIZE=10 MAXLENGTH=10 TAG="11XXXU" ALT="발행처"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBillToParty" align=top TYPE="BUTTON" ONCLICK ="vbscript:OpenConPopUp C_PopBillToParty">&nbsp;
								<INPUT TYPE=TEXT NAME="txtBillToPartyNm" SIZE=20 TAG="14">
							</TD>
							<TD CLASS=TD5>발행일</TD>
							<TD CLASS=TD6>
								<script language =javascript src='./js/s5311pa1_fpDateTime1_txtFromDt.js'></script>&nbsp;~&nbsp;
								<script language =javascript src='./js/s5311pa1_fpDateTime2_txtToDt.js'></script>
							</TD>	
						</TR>	
						<TR>
							<TD CLASS=TD5 NOWRAP>영업그룹</TD>
							<TD CLASS=TD6 NOWRAP>
								<INPUT TYPE=TEXT NAME="txtSalesGrp" SIZE=10 MAXLENGTH=4 TAG="11XXXU" ALT="영업그룹"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSalesGrp" align=top TYPE="BUTTON" ONCLICK ="vbscript:OpenConPopUp C_PopSalesGrp">&nbsp;
								<INPUT TYPE=TEXT NAME="txtSalesGrpNm" SIZE=20 TAG="14">
							</TD>
							<TD CLASS=TD5>VAT형태</TD>
							<TD CLASS=TD6>
								<INPUT TYPE=TEXT NAME="txtVATType" SIZE=10 MAXLENGTH=4 TAG="11XXXU" ALT="VAT유형"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnVATType" align=top TYPE="BUTTON" ONCLICK ="vbscript:OpenConPopUp C_PopVatType">&nbsp;
								<INPUT TYPE=TEXT NAME="txtVATTypeNm" SIZE=20 TAG="14">
							</TD>
						</TR>
						<TR>
							<TD CLASS=TD5>세금신고사업장</TD>
							<TD CLASS=TD6>
								<INPUT TYPE=TEXT NAME="txtTaxBizArea" SIZE=10 MAXLENGTH=10 TAG="11XXXU" ALT="세금신고사업장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnTaxBizArea" align=top TYPE="BUTTON" ONCLICK ="vbscript:OpenConPopUp C_PopTaxBizArea">&nbsp;
								<INPUT TYPE=TEXT NAME="txtTaxBizAreaNm" SIZE=20 TAG="14">
							</TD>
							<TD CLASS=TD5 NOWRAP>발행여부</TD> 
							<TD CLASS=TD6 NOWRAP>
								<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoTaxIssueFlg" TAG="11X" VALUE="A" CHECKED ID="rdoTaxIssueFlg1"><LABEL FOR="rdoTaxIssueFlg1">전체</LABEL>&nbsp;&nbsp;&nbsp;
								<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoTaxIssueFlg" TAG="11X" VALUE="Y" ID="rdoTaxIssueFlg2"><LABEL FOR="rdoTaxIssueFlg2">발행</LABEL>&nbsp;&nbsp;&nbsp;
								<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoTaxIssueFlg" TAG="11X" VALUE="N" ID="rdoTaxIssueFlg3"><LABEL FOR="rdoTaxIssueFlg3">미발행</LABEL>			
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
			<TD WIDTH=100% HEIGHT=* valign=top>
				<TABLE <%=LR_SPACE_TYPE_20%>>
					<TR>
						<TD HEIGHT="100%" NOWRAP>
							<script language =javascript src='./js/s5311pa1_vaSpread_vspdData.js'></script>
						</TD>
					</TR>
				</TABLE>
			</TD>
		</TR>
		<TR>
			<TD <%=HEIGHT_TYPE_01%>></TD>
		</TR>
		<TR HEIGHT=20>
			<TD WIDTH=100%>
				<TABLE <%=LR_SPACE_TYPE_30%>>
					<TR>
						<TD WIDTH=10>&nbsp;</TD>
						<TD WIDTH=70% NOWRAP>     <IMG SRC="../../../CShared/image/query_d.gif"  Style="CURSOR: hand" ALT="Search" NAME="Search" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Query.gif',1)"  ONCLICK="FncQuery()"   ></IMG>
												  <IMG SRC="../../../CShared/image/zpConfig_d.gif"  Style="CURSOR: hand" ALT="Config" NAME="Config" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/zpConfig.gif',1)"  OnClick="OpenSortPopup()"></IMG>			</TD>
						<TD WIDTH=30% ALIGN=RIGHT><IMG SRC="../../../CShared/image/ok_d.gif"     Style="CURSOR: hand" ALT="OK"     NAME="pop1"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/OK.gif',1)"     ONCLICK="OkClick()"    ></IMG>
								                  <IMG SRC="../../../CShared/image/cancel_d.gif" Style="CURSOR: hand" ALT="CANCEL" NAME="pop2"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)" ONCLICK="CancelClick()"></IMG>					</TD>
						<TD WIDTH=10>&nbsp;</TD>
					</TR>
				</TABLE>
			</TD>
		</TR>
		<TR>
			<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC= "../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME></TD>
		</TR>
	</TABLE>
<INPUT TYPE=HIDDEN NAME="txtHBillToParty" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHFromDt" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHToDt" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHSalesGrp" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHVatType" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHTaxBizArea" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHPostFlag" TAG="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>

</BODY>
</HTML>

