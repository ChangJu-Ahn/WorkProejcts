<%@ LANGUAGE="VBSCRIPT" %>
<%
'**********************************************************************************************
'*  1. Module Name          : 영업 
'*  2. Function Name        : 매출채권관리 
'*  3. Program ID           : S5111PA1
'*  4. Program Name         : 매출채권번호 팝업 
'*  5. Program Desc         : 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2002/08/01
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Hwangseongbae
'* 10. Modifier (Last)      : 
'*							: 
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************
%>
<HTML>
<HEAD>
<TITLE></TITLE>
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentA.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript" SRC="../../inc/incImage.js"></SCRIPT>
<SCRIPT LANGUAGE="VBScript">

Option Explicit

' External ASP File
'========================================
Const BIZ_PGM_ID 		= "s5111pb1.asp"                              '☆: Biz Logic ASP Name

' Constant variables 
'========================================
Const C_MaxKey          = 3                                           '☆: key count of SpreadSheet

Const C_PopSoldToParty	= 1
Const C_PopBillType		= 2
Const C_PopSalesGrp		= 3
Const C_PopSoNo			= 4
Const C_PopDnNo			= 5

' Common variables 
'========================================
<!-- #Include file="../../inc/lgvariables.inc" -->	

' User-defind Variables
'========================================
Dim IsOpenPop  

Dim lgBlnOpenedFlag
Dim	lgBlnSoldToPartyChg
Dim lgBlnBillTypeChg
Dim lgBlnSalesGrpChg
Dim lgBlnSoNoChg
Dim lgBlnDnNoChg

Dim lgIntStartRow
Dim lgExceptFlag

Dim arrParent

arrParent = window.dialogArguments
Set PopupParent = arrParent(0)

top.document.title = PopupParent.gActivePRAspName

Dim EndDate

' 시스템 날짜 
EndDate = UniConvDateAToB("<%=GetSvrDate%>", PopupParent.gServerDateFormat, PopupParent.gDateFormat)

'========================================
Function InitVariables()
	lgPageNo         = ""
    lgBlnFlgChgValue = False	                           'Indicates that no value changed
    lgIntFlgMode     = PopupParent.OPMD_CMODE                          'Indicates that current mode is Create mode
    lgSortKey        = 1   
        
	lgBlnSoldToPartyChg = False
	lgBlnSalesGrpChg	= False
	lgBlnBillTypeChg	= False
	lgBlnSoNoChg		= False
	lgBlnDnNoChg		= False
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
			
		<%If Request("txtExceptFlag") = "A" Then%>
		lgExceptFlag = "NULL"
		<%Else%>
		lgExceptFlag = "<%=Request("txtExceptFlag")%>"
		<%End If%>
	End With
	Self.Returnvalue = ""
End Sub

'==========================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A( "Q", "S", "NOCOOKIE", "PA") %>
	<% Call LoadBNumericFormatA("Q", "S", "NOCOOKIE", "PA") %>
End Sub

'========================================
Sub InitSpreadSheet()

	Call SetZAdoSpreadSheet("S5111PA1","S","A","V20030320",PopupParent.C_SORT_DBAGENT,frm1.vspdData, _
								C_MaxKey, "X","X")		
	Call SetSpreadLock 		
	    
End Sub

'========================================
Sub SetSpreadLock()
	ggoSpread.SpreadLockWithOddEvenRowColor()
	frm1.vspddata.OperationMode = 3
End Sub	

'========================================
Function OKClick()
	Dim intColCnt
		
	If frm1.vspdData.ActiveRow > 0 Then	
		frm1.vspdData.row = frm1.vspdData.ActiveRow
		frm1.vspdData.Col = GetKeyPos("A",1)
		Self.Returnvalue = frm1.vspdData.Text
	End If
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
	Case C_PopSoldToParty												
		iArrParam(1) = "dbo.b_biz_partner BP"			<%' TABLE 명칭 %>
		iArrParam(2) = Trim(frm1.txtSoldToParty.value)	<%' Code Condition%>
		iArrParam(3) = ""								<%' Name Cindition%>
		iArrParam(4) = "BP.bp_type IN (" & FilterVar("C", "''", "S") & " , " & FilterVar("CS", "''", "S") & ") AND BP.usage_flag = " & FilterVar("Y", "''", "S") & " "		<%' Where Condition%>
		iArrParam(5) = "주문처"						<%' TextBox 명칭 %>
			
		iArrField(0) = "ED15" & PopupParent.gColSep & "BP.bp_cd"	<%' Field명(0)%>
		iArrField(1) = "ED30" & PopupParent.gColSep & "BP.bp_nm"	<%' Field명(1)%>
		    
		iArrHeader(0) = "주문처"					<%' Header명(0)%>
		iArrHeader(1) = "주문처명"					<%' Header명(1)%>
	
		frm1.txtSoldToParty.focus
		
	Case C_PopBillType												
		iArrParam(1) = "s_bill_type_config"				<%' TABLE 명칭 %>
		iArrParam(2) = Trim(frm1.txtBillType.value)		<%' Code Condition%>
		iArrParam(3) = ""								<%' Name Cindition%>

		If lgExceptFlag = "NULL" Then
			iArrParam(4) = " NOT (export_flag = " & FilterVar("Y", "''", "S") & "  and except_flag = " & FilterVar("N", "''", "S") & " ) and USAGE_FLAG=" & FilterVar("Y", "''", "S") & "  AND as_flag = " & FilterVar("N", "''", "S") & " "
		ElseIf lgExceptFlag = "Y" Then
			iArrParam(4) = "USAGE_FLAG=" & FilterVar("Y", "''", "S") & "  AND except_flag = IsNull(" & FilterVar(lgExceptFlag, "''", "S") & ", except_flag) AND as_flag = " & FilterVar("N", "''", "S") & " "	<%' Where Condition%>
		ElseIf lgExceptFlag = "N"Then
			iArrParam(4) = "USAGE_FLAG=" & FilterVar("Y", "''", "S") & "  AND export_flag = " & FilterVar("N", "''", "S") & "  AND except_flag = IsNull(" & FilterVar(lgExceptFlag, "''", "S") & ", except_flag) AND as_flag = " & FilterVar("N", "''", "S") & " "	<%' Where Condition%>
		End If

		iArrParam(5) = Trim(frm1.txtBillType.alt)		<%' TextBox 명칭 %>

		iArrField(0) = "ED15" & PopupParent.gColSep & "bill_type"	<%' Field명(0)%>
		iArrField(1) = "ED30" & PopupParent.gColSep & "bill_type_nm"<%' Field명(1)%>

		iArrHeader(0) = "매출채권형태"				<%' Header명(0)%>
		iArrHeader(1) = "매출채권형태명"				<%' Header명(1)%>
		
		frm1.txtBillType.focus

	Case C_PopSalesGrp												
		iArrParam(1) = "dbo.B_SALES_GRP"				<%' TABLE 명칭 %>
		iArrParam(2) = Trim(frm1.txtSalesGrp.value)	<%' Code Condition%>
		iArrParam(3) = ""								<%' Name Cindition%>
		iArrParam(4) = "USAGE_FLAG=" & FilterVar("Y", "''", "S") & " "					<%' Where Condition%>
		iArrParam(5) = "영업그룹"					<%' TextBox 명칭 %>
		
		iArrField(0) = "ED15" & PopupParent.gColSep & "SALES_GRP"		<%' Field명(0)%>
		iArrField(1) = "ED30" & PopupParent.gColSep & "SALES_GRP_NM"	<%' Field명(1)%>
    
	    iArrHeader(0) = "영업그룹"					<%' Header명(0)%>
	    iArrHeader(1) = "영업그룹명"				<%' Header명(1)%>
	    
	    frm1.txtSalesGrp.focus

	Case C_PopSoNo
		iArrParam(1) = "S_SO_HDR SH, B_BIZ_PARTNER SP, B_SALES_GRP SG"
		iArrParam(2) = Trim(frm1.txtSONo.value)
		iArrParam(3) = ""
		iArrParam(4) = "SH.SOLD_TO_PARTY = SP.BP_CD AND SH.SALES_GRP = SG.SALES_GRP AND SH.cfm_flag = " & FilterVar("Y", "''", "S") & "  AND EXISTS (SELECT * FROM S_SO_DTL SD WHERE SD.SO_NO = SH.SO_NO AND SD.BILL_QTY > 0) "
		iArrParam(5) = "수주번호"

		iArrField(0) = "ED12" & PopupParent.gColSep & "SH.SO_NO"
		iArrField(1) = "ED10" & PopupParent.gColSep & "SH.SOLD_TO_PARTY"
		iArrField(2) = "ED15" & PopupParent.gColSep & "SP.BP_NM"
		iArrField(3) = "DD10" & PopupParent.gColSep & "SH.SO_DT"
		iArrField(4) = "ED15" & PopupParent.gColSep & "SG.SALES_GRP_NM"
		iArrField(5) = "ED10" & PopupParent.gColSep & "SH.PAY_METH"
		
		iArrHeader(0) = "수주번호"
		iArrHeader(1) = "주문처"
		iArrHeader(2) = "주문처명"
		iArrHeader(3) = "수주일"
		iArrHeader(4) = "영업그룹명"
		iArrHeader(5) = "결제방법"
		
		frm1.txtSONo.focus

	Case C_PopDNNo
		iArrParam(1) = "S_DN_HDR DH, B_BIZ_PARTNER SH, B_MINOR MT, B_SALES_GRP SG"
		iArrParam(2) = Trim(frm1.txtDNNo.value)
		iArrParam(3) = ""
		iArrParam(4) = "DH.SHIP_TO_PARTY = SH.BP_CD AND DH.MOV_TYPE = MT.MINOR_CD AND MT.MAJOR_CD = " & FilterVar("I0001", "''", "S") & " AND DH.SALES_GRP = SG.SALES_GRP AND DH.post_flag = " & FilterVar("Y", "''", "S") & "  AND EXISTS (SELECT * FROM s_dn_dtl DD WHERE DD.dn_no = DH.dn_no AND DD.bill_qty > 0) "						<%' Where Condition%>
		iArrParam(5) = "출하번호"

		iArrField(0) = "ED10" & PopupParent.gColSep & "DH.DN_NO"
		iArrField(1) = "ED10" & PopupParent.gColSep & "DH.SHIP_TO_PARTY"
		iArrField(2) = "ED15" & PopupParent.gColSep & "SH.BP_NM"
		iArrField(3) = "DD10" & PopupParent.gColSep & "DH.DLVY_DT"
		iArrField(4) = "DD10" & PopupParent.gColSep & "DH.ACTUAL_GI_DT"
		iArrField(5) = "ED15" & PopupParent.gColSep & "MT.MINOR_NM"
		iArrField(6) = "ED15" & PopupParent.gColSep & "SG.SALES_GRP_NM"

		iArrHeader(0) = "출하번호"
		iArrHeader(1) = "납품처"
		iArrHeader(2) = "납품처명"
		iArrHeader(3) = "납품일"
		iArrHeader(4) = "실제출고일"
		iArrHeader(5) = "출하형태명"
		iArrHeader(6) = "영업그룹명"
		
		frm1.txtDNNo.focus

	End Select
 
	iArrParam(0) = iArrParam(5)							<%' 팝업 명칭 %> 

	iArrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(iArrParam, iArrField, iArrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If iArrRet(0) <> "" Then
		OpenConPopup = SetConPopup(iArrRet,pvIntWhere)
	End If	
	
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

'========================================
Function SetConPopup(Byval pvArrRet,ByVal pvIntWhere)

	SetConPopup = False

	Select Case pvIntWhere
	Case C_PopSoldToParty
		frm1.txtSoldToParty.value = pvArrRet(0) 
		frm1.txtSoldToPartyNm.value = pvArrRet(1)   
		frm1.txtSoldToParty.focus
	Case C_PopBillType
		frm1.txtBillType.value = pvArrRet(0) 
		frm1.txtBillTypeNm.value = pvArrRet(1)
		frm1.txtBillType.focus   
	Case C_PopSalesGrp
		frm1.txtSalesGrp.value = pvArrRet(0) 
		frm1.txtSalesGrpNm.value = pvArrRet(1)
		frm1.txtSalesGrp.focus   
	Case C_PopSoNO
		frm1.txtSoNo.value = pvArrRet(0) 
	Case C_PopDnNo
		frm1.txtDnNo.value = pvArrRet(0) 
	End Select

	SetConPopup = True

End Function

'========================================
Sub Form_Load()
   
    Call LoadInfTB19029											  '⊙: Load table , B_numeric_format
       'Html에서 tag 숫자가 1과 2로 시작하는 부분 각각Format
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
	
	Call ggoOper.LockField(Document, "N")                         '⊙: Lock  Suitable  Field
    
	Call InitVariables											  '⊙: Initializes local global variables
	Call SetDefaultVal	
	Call InitSpreadSheet()
	
	Call MM_preloadImages("../../../CShared/image/Query.gif","../../../CShared/image/OK.gif","../../../CShared/image/Cancel.gif")
	lgBlnOpenedflag = True
	DbQuery()
	
End Sub

'========================================
Sub Form_QueryUnload(Cancel, UnloadMode)
End Sub

'==========================================
Function GetSalesGrpNm()
	Dim iStrCode
	
	With frm1
		iStrCode = Trim(.txtSalesGrp.value)
		If iStrCode <> "" Then
			iStrCode = " " & FilterVar(iStrCode, "''", "S") & ""
			If Not GetCodeName(iStrCode, "default", "default", "default", "default", "" & FilterVar("SG", "''", "S") & "", C_PopSalesGrp) Then
				.txtSalesGrp.value = ""
				.txtSalesGrpNm.value = ""
				.txtSalesGrp.focus
			Else
				<%If Request("txtExceptFlag") = "Y" Then %>
				.rdoReleaseFlg1.focus
				<%Else%>
				.txtSoNo.focus
				<%End If%>
			End If
		Else
			.txtSalesGrpNm.value = ""
		End If
	End With
End Function

'==========================================
Function txtSoldToParty_OnKeyDown()
	lgBlnSoldToPartyChg = True
	lgBlnFlgChgValue = True
End Function

'==========================================
Function txtBillType_OnKeyDown()
	lgBlnBillTypeChg = True
	lgBlnFlgChgValue = True
End Function

'==========================================
Function txtSalesGrp_OnKeyDown()
	lgBlnSalesGrpChg = True
	lgBlnFlgChgValue = True
End Function

'==========================================
Function txtSoNo_OnKeyDown()
	lgBlnSoNoChg = True
	lgBlnFlgChgValue = True
End Function

'==========================================
Function txtDnNo_OnKeyDown()
	lgBlnDnNoChg = True
	lgBlnFlgChgValue = True
End Function

'	Description : 조회조건의 유효성을 Check한다.
'   주의사항 : 화면의 tab order 별로 기술한다. 
'==========================================
Function ChkValidityQueryCon()
	Dim iStrCode

	ChkValidityQueryCon = True

	If lgBlnSoldToPartyChg Then
		iStrCode = Trim(frm1.txtSoldToParty.value)
		If iStrCode <> "" Then
			iStrCode = " " & FilterVar(iStrCode, "''", "S") & ""
			If Not GetCodeName(iStrCode, "" & FilterVar("C%", "''", "S") & "", "default", "default", "default", "" & FilterVar("BP", "''", "S") & "", C_PopSoldToParty) Then
				Call DisplayMsgBox("970000", "X", frm1.txtSoldtoparty.alt, "X")
				frm1.txtSoldtoparty.focus
				ChkValidityQueryCon = False
				Exit Function
			End If
		Else
			frm1.txtSoldToPartyNm.value = ""
		End If
		lgBlnSoldToPartyChg	= False
	End If

	If lgBlnBillTypeChg Then
		iStrCode = Trim(frm1.txtBillType.value)
		If iStrCode <> "" Then
			iStrCode = " " & FilterVar(iStrCode, "''", "S") & ""
			If Not GetCodeName(iStrCode, FilterVar(lgExceptFlag, "''", "S"), "" & FilterVar("N", "''", "S") & " ", "NULL", "default", "" & FilterVar("BT", "''", "S") & "", C_PopBillType) Then
				Call DisplayMsgBox("970000", "X", frm1.txtBillType.alt, "X")
				frm1.txtBillType.focus
				ChkValidityQueryCon = False
				Exit Function
			End If
		Else
			frm1.txtBillTypeNm.value = ""
		End If
		lgBlnBillTypeChg = False
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
		'If lgBlnOpenedFlag Then	GetCodeName = OpenConPopup(pvIntWhere)
	End if
End Function

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
	If KeyAscii = 27 Then
		Call CancelClick()
	Elseif KeyAscii = 13 Then
		Call FncQuery()
	End if
End Sub

'========================================
Sub txtToDt_Keypress(KeyAscii)
	If KeyAscii = 27 Then
		Call CancelClick()
	Elseif KeyAscii = 13 Then
		Call FncQuery()
	End if
End Sub

'========================================
Function vspdData_DblClick(ByVal Col, ByVal Row)
	If Frm1.vspdData.ActiveRow > 0 Then	Call OKClick
End Function

'========================================
Function vspdData_KeyPress(KeyAscii)
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
		If lgPageNo <> "" Then
			If CheckRunningBizProcess Then Exit Sub
			Call DbQuery
		End If
	End If
End Sub

'========================================
Function FncQuery() 
    
    FncQuery = False
    
    Err.Clear

	If ValidDateCheck(frm1.txtFromDt, frm1.txtToDt) = False Then Exit Function

	' 조회조건 유효값 check
	If 	lgBlnFlgChgValue Then
		If Not ChkValidityQueryCon Then	Exit Function
	End If

    Call ggoOper.ClearField(Document, "2")
	
    Call InitVariables
    
	If DbQuery = False Then Exit Function									

    FncQuery = True		
End Function

'========================================
Function DbQuery() 

	Err.Clear
	DbQuery = False
	
	If LayerShowHide(1) = False Then Exit Function
	
	Dim strVal
	
    With frm1
		strVal = BIZ_PGM_ID & "?txtHMode=" & PopupParent.UID_M0001					<%'☜: 비지니스 처리 ASP의 상태 %>
		If lgIntFlgMode = PopupParent.OPMD_UMODE Then
			' Scroll시 
			strVal = strVal & "&txtSoldToParty=" & .txtHSoldToParty.value
			strVal = strVal & "&txtFromDt=" & .txtHFromDt.value
			strVal = strVal & "&txtToDt=" & .txtHToDt.value
			strVal = strVal & "&txtBillType=" & .txtHBillType.value
			strVal = strVal & "&txtSalesGrp=" & .txtHSalesGrp.value
			strVal = strVal & "&txtSoNo=" & .txtHSoNo.value
			strVal = strVal & "&txtDnNo=" & .txtHDnNo.value
			strVal = strVal & "&txtPostFlag=" & .txtHPostFlag.value
			strVal = strVal & "&txtExceptFlag=" & .txtHExceptFlag.value
		Else
			' 처음 조회시 
			strVal = strVal & "&txtSoldToParty=" & Trim(.txtSoldToParty.value)
			strVal = strVal & "&txtFromDt=" & Trim(.txtFromDt.Text)				<%'☆: 조회 조건 데이타 %>
			strVal = strVal & "&txtToDt=" & Trim(.txtToDt.Text)
			strVal = strVal & "&txtBillType=" & Trim(.txtBillType.value)
			strVal = strVal & "&txtSalesGrp=" & Trim(.txtSalesGrp.value)

			If .rdoReleaseFlg2.checked = True Then
				strVal = strVal & "&txtPostFlag=Y"
			ElseIf frm1.rdoReleaseFlg3.checked = True Then
				strVal = strVal & "&txtPostFlag=N"
			Else
				strVal = strVal & "&txtPostFlag="
			End If
			
			<% If Request("txtExceptflag") = "A" Then%>
				strVal = strVal & "&txtSoNo=" & Trim(.txtSoNo.value)
				strVal = strVal & "&txtDnNo=" & Trim(.txtDnNo.value)

				If .rdoExceptFlg2.checked = True Then
					strVal = strVal & "&txtExceptFlag=" & FilterVar("Y", "''", "S") & " "
				ElseIf frm1.rdoExceptFlg3.checked = True Then
					strVal = strVal & "&txtExceptFlag=" & FilterVar("N", "''", "S") & " "
				Else
					strVal = strVal & "&txtExceptFlag=NULL"
				End If
			' 정상매출 
			<% Elseif Request("txtExceptflag") = "N" Then %>
				strVal = strVal & "&txtSoNo=" & Trim(.txtSoNo.value)
				strVal = strVal & "&txtDnNo=" & Trim(.txtDnNo.value)
				strVal = strVal & "&txtExceptFlag=" & FilterVar("N", "''", "S") & " "
			' 예외매출 
			<% Else %>
				strVal = strVal & "&txtSoNo="
				strVal = strVal & "&txtDnNo="
				strVal = strVal & "&txtExceptFlag=" & FilterVar("Y", "''", "S") & " "
			<% End If%>
		End If

        strVal = strVal & "&lgPageNo="		 & lgPageNo
		strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")
		
        strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
		strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))
		
		lgIntStartRow = .vspdData.MaxRows + 1
	End With    
 
	Call RunMyBizASP(MyBizASP, strVal)
    DbQuery = True    

End Function

'=========================================
Function DbQueryOk()
	If frm1.vspdData.MaxRows > 0 Then
		frm1.vspdData.Focus
		frm1.vspdData.SelModeSelected = True
		If lgIntFlgMode <> PopupParent.OPMD_UMODE Then
			frm1.vspdData.Row = 1
			lgIntFlgMode = PopupParent.OPMD_UMODE
		End If
		Call FormatSpreadCellByCurrency()
	Else
		frm1.txtSoldToParty.focus
	End If

End Function

' 화폐별로 Cell Formating을 재설정한다.
Sub FormatSpreadCellByCurrency()
	With frm1
		Call ReFormatSpreadCellByCellByCurrency(.vspdData,lgIntStartRow, .vspdData.MaxRows,GetKeyPos("A",2),GetKeyPos("A",3),"A","Q","X","X") 
	End With
End Sub

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
							<TD CLASS=TD5>주문처</TD>
							<TD CLASS=TD6><INPUT TYPE=TEXT NAME="txtSoldToParty" SIZE=10 MAXLENGTH=10 TAG="11XXXU" ALT="주문처"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSoldToParty" align=top TYPE="BUTTON" ONCLICK ="vbscript:OpenConPopUp C_PopSoldToParty">&nbsp;<INPUT TYPE=TEXT NAME="txtSoldToPartyNm" SIZE=20 TAG="14" Alt="주문처명"></TD>
							<TD CLASS=TD5>매출채권일</TD>
							<TD CLASS=TD6>
								<script language =javascript src='./js/s5111pa1_fpDateTime1_txtFromDt.js'></script>&nbsp;~&nbsp;
								<script language =javascript src='./js/s5111pa1_fpDateTime2_txtToDt.js'></script>
							</TD>	
						</TR>	
						<TR>
							<TD CLASS=TD5>매출채권형태</TD>
							<TD CLASS=TD6><INPUT TYPE=TEXT NAME="txtBillType" ALT="매출채권형태" SIZE=10 MAXLENGTH=20 TAG="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBillType" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConPopUp C_PopBillType">&nbsp;<INPUT TYPE=TEXT NAME="txtBillTypeNm" SIZE=20 TAG="14"></TD>
							<TD CLASS=TD5 NOWRAP>영업그룹</TD>
							<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtSalesGrp" SIZE=10 MAXLENGTH=4 TAG="11XXXU" ALT="영업그룹"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSalesGrp" align=top TYPE="BUTTON" ONCLICK ="vbscript:OpenConPopUp C_PopSalesGrp">&nbsp;<INPUT TYPE=TEXT NAME="txtSalesGrpNm" SIZE=20 TAG="14"></TD>
						</TR>
						<% If Request("txtExceptFlag") <> "Y" Then %>
						<TR>
							<TD CLASS=TD5 NOWRAP>수주번호</TD>
							<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtSONo" SIZE=30 MAXLENGTH=18 TAG="11XXXU" ALT="수주번호"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSONo" align=top TYPE="BUTTON" OnClick="vbscript:OpenConPopUp C_PopSoNO"></TD>
							<TD CLASS=TD5 NOWRAP>출하번호</TD>
							<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtDNNo" SIZE=30 MAXLENGTH=18 TAG="11XXXU" ALT="출하번호"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDNNo" align=top TYPE="BUTTON" OnClick="vbscript:OpenConPopUp C_PopDnNO"></TD>
						</TR>
						<% End If%>
						<TR>
							<TD CLASS=TD5 NOWRAP>확정여부</TD> 
							<TD CLASS=TD6 NOWRAP>
								<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoReleaseFlg" TAG="11X" VALUE="A" ID="rdoReleaseFlg1"><LABEL FOR="rdoReleaseFlg1">전체</LABEL>&nbsp;&nbsp;&nbsp;
								<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoReleaseFlg" TAG="11X" VALUE="Y" ID="rdoReleaseFlg2"><LABEL FOR="rdoReleaseFlg2">확정</LABEL>&nbsp;&nbsp;&nbsp;
								<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoReleaseFlg" TAG="11X" VALUE="N" CHECKED ID="rdoReleaseFlg3"><LABEL FOR="rdoReleaseFlg3">미확정</LABEL>			
							</TD>
							<% If Request("txtExceptFlag") = "A" Then %>
							<TD CLASS=TD5 NOWRAP>예외매출채권여부</TD> 
							<TD CLASS=TD6 NOWRAP>
								<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoExceptFlg" TAG="11X" VALUE="A" CHECKED ID="rdoExceptFlg1"><LABEL FOR="rdoExceptFlg1">전체</LABEL>&nbsp;&nbsp;&nbsp;
								<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoExceptFlg" TAG="11X" VALUE="Y" ID="rdoExceptFlg2"><LABEL FOR="rdoExceptFlg2">예</LABEL>&nbsp;&nbsp;&nbsp;
								<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoExceptFlg" TAG="11X" VALUE="N" ID="rdoExceptFlg3"><LABEL FOR="rdoExceptFlg3">아니오</LABEL>			
							</TD>
							<% Else %>
							<TD CLASS=TD5 NOWRAP></TD> 
							<TD CLASS=TD6 NOWRAP></TD>
							<% End If%>
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
							<script language =javascript src='./js/s5111pa1_vaSpread_vspdData.js'></script>
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
			<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC= "../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX ="-1"></IFRAME></TD>
		</TR>
	</TABLE>
<INPUT TYPE=HIDDEN NAME="txtHSoldToParty" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHFromDt" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHToDt" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHBilltype" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHSalesGrp" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHSoNo" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHDnNo" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHPostFlag" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtHExceptFlag" TAG="24">

</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>

