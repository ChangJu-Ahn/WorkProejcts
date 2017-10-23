<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% session.CodePage=949 %>

<%
'**********************************************************************************************
'*  1. Module Name          : 영업
'*  2. Function Name        : 매출채권관리
'*  3. Program ID           : S3111BA3
'*  4. Program Name         : 출하정보참조(매출채권등록)
'*  5. Program Desc         : 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2000/12/09
'*  8. Modified date(Last)  : 2002/06/04
'*  9. Modifier (First)     : Byun Jee Hyun
'* 10. Modifier (Last)      : Kim Hyungsuk
'*							: Hwangseongbae
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                            2000/12/09
'*                            2001/12/18	Date표준적용
'*                            2002/06/04	ADO 표준변경 및 Default 기간 변경
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
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMA_KO441.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript">

Option Explicit

' External ASP File
'========================================
Const BIZ_PGM_ID 		= "s3111bb3.asp"                              '☆: Biz Logic ASP Name

' Constant variables 
'========================================
Const C_MaxKey          = 4                                            '☆: key count of SpreadSheet

Const C_PopSoldToParty	= 1
Const C_PopSalesGrp		= 2
Const C_PopBillType		= 3
Const C_PopCurrency		= 4
Const C_PopDNNo			= 5

' Common variables 
'========================================
<!-- #Include file="../../inc/lgvariables.inc" -->	

' User-defind Variables
'========================================
Dim gblnWinEvent											'☜: ShowModal Dialog(PopUp) 
														    'Window가 여러 개 뜨는 것을 방지하기 위해
														    'PopUp Window가 사용중인지 여부를 나타냄
Dim lgBlnOpenedFlag
Dim	lgBlnSoldToPartyChg
Dim lgBlnSalesGrpChg
Dim	lgBlnBillTypeChg
Dim	lgBlnCurrencyChg

Dim arrParent

arrParent = window.dialogArguments
Set PopupParent = arrParent(0)

top.document.title = PopupParent.gActivePRAspName

Dim iDBSYSDate
Dim EndDate

iDBSYSDate = "<%=GetSvrDate%>"
EndDate = UniConvDateAToB(iDBSYSDate, PopupParent.gServerDateFormat, PopupParent.gDateFormat)

'========================================
Function InitVariables()
	lgPageNo         = ""
    lgBlnFlgChgValue = False	                           'Indicates that no value changed
    lgIntFlgMode     = PopupParent.OPMD_CMODE              'Indicates that current mode is Create mode
    lgSortKey        = 1   
        
    gblnWinEvent = False
        
	lgBlnSoldToPartyChg = False		' 주문처 변경여부
	lgBlnSalesGrpChg	= False		' 영업그룹 변경여부
	lgBlnBillTypeChg	= False		' 매출채권 변경여부
	lgBlnCurrencyChg	= False		' 화폐 변경여부
End Function

'=======================================================
Sub SetDefaultVal()
	Dim iArrReturn
		
	With frm1
		.txtFromDt.Text = UNIDateClientFormat(UniConvDateAToB(UniConvDateToYYYYMM(EndDate, PopupParent.gDateFormat, "-") & "-01", PopupParent.gServerDateFormat ,PopupParent.gAPDateFormat))
		.txtToDt.Text = EndDate

		If PopupParent.gSalesGrp <> "" Then
			.txtSalesGrp.value = PopupParent.gSalesGrp
			Call GetSalesGrpNm()
		End If
	End With
	Redim iArrReturn(0)
	Self.Returnvalue = iArrReturn
	
	If lgSGCd <> "" then 
		Call ggoOper.SetReqAttr(frm1.txtSalesGrp, "Q") 
        	frm1.txtSalesGrp.value = lgSGCd
	End If
	
End Sub

'==========================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"-->
	<% Call LoadInfTB19029A("I", "*", "NOCOOKIE", "RA") %>	
End Sub

'========================================
Sub InitSpreadSheet()
	
	Call SetZAdoSpreadSheet("S3111BA3","S","A","V20021202",PopupParent.C_SORT_DBAGENT,frm1.vspdData, _
								C_MaxKey, "X","X")		
	Call SetSpreadLock 	
	    
End Sub

'========================================
	Sub SetSpreadLock()
	    With frm1
	    .vspdData.ReDraw = False
'		ggoSpread.SpreadLock 1 , -1
		ggoSpread.SpreadLockWithOddEvenRowColor()
		.vspddata.OperationMode = 3
	    .vspdData.ReDraw = True

	    End With
	End Sub	


'========================================
Function OKClick()
	Dim iArrReturn
	' 공통 eventpopup.vbs에서 에러가 발생.
	' 해당 error를 clear함. (공통 수정시 삭제)
	err.Clear
	With frm1
		If .vspdData.ActiveRow > 0 Then	
			Redim iArrReturn(3)
			.vspdData.Row = .vspdData.ActiveRow
			.vspdData.Col = GetKeyPos("A",1)			' 수주번호
			iArrReturn(0) = .vspdData.Text
			.vspdData.Col = GetKeyPos("A",2)			' 매출채권형태
			iArrReturn(1) = .vspdData.Text
			.vspdData.Col = GetKeyPos("A",3)			' 매출채권형태명
			iArrReturn(2) = .vspdData.Text
			.vspdData.Col = GetKeyPos("A",4)			' 출고번호
			iArrReturn(3) = .vspdData.Text
			
			Self.Returnvalue = iArrReturn
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
	
	If gblnWinEvent = True Then Exit Function

	gblnWinEvent = True

	Select Case pvIntWhere

	Case C_PopSoldToParty												
		iArrParam(1) = "dbo.b_biz_partner BP"			<%' TABLE 명칭%>
		iArrParam(2) = Trim(frm1.txtSoldToParty.value)	<%' Code Condition%>
		iArrParam(3) = ""								<%' Name Cindition%>
		iArrParam(4) = "BP.bp_type IN (" & FilterVar("C", "''", "S") & " , " & FilterVar("CS", "''", "S") & ") AND BP.usage_flag = " & FilterVar("Y", "''", "S") & " "		<%' Where Condition%>
		iArrParam(5) = frm1.txtSoldtoparty.alt '"주문처"						<%' TextBox 명칭%>
			
		iArrField(0) = "ED15" & PopupParent.gColSep & "BP.bp_cd"	<%' Field명(0)%>
		iArrField(1) = "ED30" & PopupParent.gColSep & "BP.bp_nm"	<%' Field명(1)%>
		    
		iArrHeader(0) = "주문처"					<%' Header명(0)%>
		iArrHeader(1) = "주문처명"					<%' Header명(1)%>
		
		frm1.txtSoldtoparty.focus

	Case C_PopSalesGrp	
	
		If frm1.txtSalesGrp.className = "protected" Then
	     OpenConPopup = False
	     Exit Function
	    End If
											
		iArrParam(1) = "dbo.B_SALES_GRP"
		iArrParam(2) = Trim(frm1.txtSalesGrp.value)
		iArrParam(3) = ""
		iArrParam(4) = "USAGE_FLAG=" & FilterVar("Y", "''", "S") & " "
		iArrParam(5) = Trim(frm1.txtSalesGrp.alt)
		
		iArrField(0) = "ED15" & PopupParent.gColSep & "SALES_GRP"
		iArrField(1) = "ED30" & PopupParent.gColSep & "SALES_GRP_NM"
    
	    iArrHeader(0) = "영업그룹"
	    iArrHeader(1) = "영업그룹명"

		frm1.txtSalesGrp.focus
		
	Case C_PopBillType												
		iArrParam(1) = "s_bill_type_config"
		iArrParam(2) = Trim(frm1.txtBillType.value)
		iArrParam(3) = ""
		iArrParam(4) = "USAGE_FLAG=" & FilterVar("Y", "''", "S") & "  AND except_flag = " & FilterVar("N", "''", "S") & "  AND export_flag = " & FilterVar("N", "''", "S") & "  AND ref_dn_flag = " & FilterVar("Y", "''", "S") & "  "
		iArrParam(5) = Trim(frm1.txtBillType.alt)

		iArrField(0) = "ED15" & PopupParent.gColSep & "bill_type"
		iArrField(1) = "ED30" & PopupParent.gColSep & "bill_type_nm"

		iArrHeader(0) = "매출채권형태"
		iArrHeader(1) = "매출채권형태명"
		
		frm1.txtBillType.focus

	Case C_PopCurrency
		iArrParam(1) = "B_CURRENCY"
		iArrParam(2) = Trim(frm1.txtCurrency.value)
		iArrParam(3) = ""
		iArrParam(4) = ""
		iArrParam(5) = Trim(frm1.txtCurrency.alt)
			
		iArrField(0) = "ED15" & PopupParent.gColSep & "CURRENCY"
		iArrField(1) = "ED30" & PopupParent.gColSep & "CURRENCY_DESC"
		    
		iArrHeader(0) = "화폐"
		iArrHeader(1) = "화폐명"

	Case C_PopDNNo
		iArrParam(1) = "S_DN_HDR DH, B_BIZ_PARTNER SH, B_MINOR MT, B_SALES_GRP SG"
		iArrParam(2) = Trim(frm1.txtDNNo.value)
		iArrParam(3) = ""
		iArrParam(4) = "DH.SHIP_TO_PARTY = SH.BP_CD AND DH.MOV_TYPE = MT.MINOR_CD AND MT.MAJOR_CD = " & FilterVar("I0001", "''", "S") & " AND DH.SALES_GRP = SG.SALES_GRP AND DH.post_flag = " & FilterVar("Y", "''", "S") & "  AND EXISTS (SELECT * FROM s_dn_dtl DD WHERE DD.dn_no = DH.dn_no AND DD.bill_qty < DD.gi_qty) "						<%' Where Condition%>
		iArrParam(5) = "출하번호"

		iArrField(0) = "ED15" & PopupParent.gColSep & "DH.DN_NO"
		iArrField(1) = "ED15" & PopupParent.gColSep & "DH.SHIP_TO_PARTY"
		iArrField(2) = "ED20" & PopupParent.gColSep & "SH.BP_NM"
		iArrField(3) = "DD15" & PopupParent.gColSep & "DH.DLVY_DT"
		iArrField(4) = "DD15" & PopupParent.gColSep & "DH.ACTUAL_GI_DT"
		iArrField(5) = "ED20" & PopupParent.gColSep & "MT.MINOR_NM"
		iArrField(6) = "ED20" & PopupParent.gColSep & "SG.SALES_GRP_NM"


		iArrHeader(0) = "출하번호"
		iArrHeader(1) = "납품처"
		iArrHeader(2) = "납품처명"
		iArrHeader(3) = "납품일"
		iArrHeader(4) = "실제출고일"
		iArrHeader(5) = "출하형태명"
		iArrHeader(6) = "영업그룹명"
		
		frm1.txtDnNo.focus

	End Select
 
	iArrParam(0) = iArrParam(5)							<%' 팝업 명칭%> 

	iArrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(iArrParam, iArrField, iArrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	gblnWinEvent = False

	If iArrRet(0) = "" Then
		Exit Function
	Else
		Call SetConPopup(iArrRet,pvIntWhere)
		OpenConPopup = True
	End If	
	
End Function

'========================================
Function OpenSortPopup()
	Dim arrRet
	
	On Error Resume Next 
	
	If gblnWinEvent = True Then Exit Function
	gblnWinEvent = True

	arrRet = window.showModalDialog("../../ComAsp/ZADOSortPopup.asp",Array(ggoSpread.GetXMLData("A"),gMethodText),"dialogWidth=" & PopupParent.SORTW_WIDTH & "px; dialogHeight=" & PopupParent.SORTW_HEIGHT & "px;; center: Yes; help: No; resizable: No; status: No;")

	gblnWinEvent = False
	
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
	Case C_PopSalesGrp
		frm1.txtSalesGrp.value = pvArrRet(0) 
		frm1.txtSalesGrpNm.value = pvArrRet(1)   
	Case C_PopBillType
		frm1.txtBillType.value = pvArrRet(0) 
		frm1.txtBillTypeNm.value = pvArrRet(1)   
	Case C_PopCurrency
		frm1.txtCurrency.value = pvArrRet(0) 
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

' 영업그룹명을 Fetch한다.
'==========================================
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

'==========================================
Function txtSoldToParty_OnKeyDown()
	lgBlnFlgChgValue = True
	lgBlnSoldToPartyChg = True
End Function

'==========================================
Function txtSalesGrp_OnKeyDown()
	lgBlnFlgChgValue = True
	lgBlnSalesGrpChg = True
End Function

'==========================================
Function txtBillType_OnKeyDown()
	lgBlnFlgChgValue = True
	lgBlnBillTypeChg = True
End Function

'==========================================
Function txtCurrency_OnKeyDown()
	lgBlnFlgChgValue = True
	lgBlnCurrencyChg = True
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
			
	If lgBlnBillTypeChg Then
		iStrCode = Trim(frm1.txtBillType.value)
		If iStrCode <> "" Then
			iStrCode = " " & FilterVar(iStrCode, "''", "S") & ""
			If Not GetCodeName(iStrCode, "" & FilterVar("N", "''", "S") & " ", "" & FilterVar("N", "''", "S") & " ", "" & FilterVar("Y", "''", "S") & " ", "default", "" & FilterVar("BT", "''", "S") & "", C_PopBillType) Then
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

	If lgBlnCurrencyChg Then
		iStrCode = Trim(frm1.txtCurrency.value)
		If iStrCode <> "" Then
			iStrCode = " " & FilterVar(iStrCode, "''", "S") & ""
			If Not GetCodeName(iStrCode, "default", "default", "default", "default", "" & FilterVar("CR", "''", "S") & "", C_PopCurrency) Then
				Call DisplayMsgBox("970000", "X", frm1.txtCurrency.alt, "X")
				frm1.txtCurrency.focus
				ChkValidityQueryCon = False
				Exit Function
			End If
		Else
			frm1.txtCurrency.value = ""
		End If
		lgBlnCurrencyChg = False
	End If

End Function

'==========================================
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
		'If lgBlnOpenedFlag Then GetCodeName = OpenConPopup(pvIntWhere)
	End if
End Function

'========================================
Function vspdData_DblClick(ByVal Col, ByVal Row)
	If Frm1.vspdData.ActiveRow > 0 Then	Call OKClick
End Function

'========================================
    Function vspdData_KeyPress(KeyAscii)
		On Error Resume Next
		If KeyAscii = 13 And vspdData.ActiveRow > 0 Then
			Call OKClick()
		ElseIf KeyAscii = 27 Then
			Call CancelClick()
		End If
    End Function

'========================================
Sub vspdData_TopLeftChange(ByVal OldLeft, ByVal OldTop, ByVal NewLeft, ByVal NewTop)
	
	If OldLeft <> NewLeft Then
	    Exit Sub
	End If

	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then
		If lgPageNo <> "" Then
			If DbQuery = False Then
				Exit Sub
			End if
		End If
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

'==========================================
Function FncQuery() 
    
    FncQuery = False                                                        '⊙: Processing is NG
    
    Err.Clear                                                               '☜: Protect system from crashing
	'** ValidDateCheck(pObjFromDt, pObjToDt) : 'pObjToDt'이 'pObjFromDt'보다 크거나 같아야 할때 **
	With frm1
		If ValidDateCheck(.txtFromDt, .txtToDt) = False Then Exit Function

		If UniConvDateToYYYYMMDD(.txtFromDt.text , PopupParent.gDateFormat , "") > UniConvDateToYYYYMMDD(EndDate, PopupParent.gDateFormat , "") Then		
			Call DisplayMsgBox("970025", "X", .txtFromDt.ALT, "현재일" & "(" & EndDate & ")")
			.txtFromDt.focus	
			Exit Function
		End If

		If UniConvDateToYYYYMMDD(.txtToDt.text , PopupParent.gDateFormat , "") > UniConvDateToYYYYMMDD(EndDate, PopupParent.gDateFormat , "") Then		
			Call DisplayMsgBox("970025", "X", .txtToDt.ALT, "현재일" & "(" & EndDate & ")")	
			.txtToDt.Focus
			Exit Function
		End If
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
	
	If LayerShowHide(1) = False Then
		Exit Function
	End If
	
	Dim strVal
	
    With frm1
		strVal = BIZ_PGM_ID & "?txtHMode=" & PopupParent.UID_M0001					<%'☜: 비지니스 처리 ASP의 상태 %>
		If lgIntFlgMode = PopupParent.OPMD_UMODE Then
			' Scroll시
			strVal = strVal & "&txtSoldToParty=" & .txtHSoldToParty.value
			strVal = strVal & "&txtFromDt=" & .txtHFromDt.value
			strVal = strVal & "&txtToDt=" & .txtHToDt.value
			strVal = strVal & "&txtSalesGrp=" & .txtHSalesGrp.value
			strVal = strVal & "&txtBillType=" & .txtHBillType.value
			strVal = strVal & "&txtCurrency=" & .txtHCurrency.value
			strVal = strVal & "&txtDnNo=" & .txtHDnNo.value
		Else
			strVal = strVal & "&txtSoldToParty=" & Trim(.txtSoldToParty.value)
			' 처음 조회시
			strVal = strVal & "&txtFromDt=" & Trim(.txtFromDt.Text)				<%'☆: 조회 조건 데이타%>
			If Len(Trim(.txtToDt.text)) Then
				strVal = strVal & "&txtToDt=" & Trim(.txtToDt.Text)
			Else
				strVal = strVal & "&txtToDt=" & EndDate
			End if
			strVal = strVal & "&txtSalesGrp=" & Trim(.txtSalesGrp.value)
			strVal = strVal & "&txtBillType=" & Trim(.txtBillType.value)		
			strVal = strVal & "&txtDnNo=" & Trim(.txtDnNo.value)
			strVal = strVal & "&txtCurrency=" & Trim(.txtCurrency.value)
		End If
		
        strVal = strVal & "&lgPageNo="		 & lgPageNo						'☜: Next key tag 
		strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")
		
        strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
		strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))
	End With    
    
	Call RunMyBizASP(MyBizASP, strVal)
    DbQuery = True    
End Function

'=========================================
Function DbQueryOk()

	If frm1.vspdData.MaxRows > 0 Then
		lgIntFlgMode = PopupParent.OPMD_UMODE
		frm1.vspdData.Focus
	Else
		frm1.txtSoldToParty.focus
	End If

End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>
<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<BODY SCROLL=NO TABINDEX="-1">
<TABLE <%=LR_SPACE_TYPE_20%>>
	<TR>
		<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
	</TR>
	<TR>
		<TD HEIGHT=20 WIDTH=100%>
			<FIELDSET CLASS="CLSFLD">
				<TABLE <%=LR_SPACE_TYPE_40%>>
					<TR>
						<TD CLASS=TD5>주문처</TD>
						<TD CLASS=TD6><INPUT TYPE=TEXT NAME="txtSoldtoparty" ALT="주문처" SIZE=10 MAXLENGTH=10 TAG="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSoldToParty" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConPopUp C_PopSoldToParty">&nbsp;<INPUT TYPE=TEXT NAME="txtSoldtoPartyNm" SIZE=20 TAG="14"></TD>
						<TD CLASS="TD5" NOWRAP>출하일</TD>
						<TD CLASS="TD6" NOWRAP>
							<TABLE CELLSPACING=0 CELLPADDING=0>
								<TR>
									<TD>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 NAME="txtFromDt" CLASS="FPDTYYYYMMDD" tag="11X1" Alt="시작일" Title="FPDATETIME"></OBJECT>');</SCRIPT>
									</TD>
									<TD>
										&nbsp;~&nbsp;
									</TD>
									<TD>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime2 NAME="txtToDt" CLASS="FPDTYYYYMMDD" tag="11X1" Alt="종료일" Title="FPDATETIME"></OBJECT>');</SCRIPT>
									</TD>
								</TR>
							</TABLE>
						</TD>
					</TR>
					<TR>
						<TD CLASS=TD5>영업그룹</TD>
						<TD CLASS=TD6><INPUT TYPE=TEXT NAME="txtSalesGrp" ALT="영업그룹" SIZE=10 MAXLENGTH=4 TAG="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSalesGrp" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConPopUp C_PopSalesGrp">&nbsp;<INPUT TYPE=TEXT NAME="txtSalesGrpNm" SIZE=20 TAG="14"></TD>
						<TD CLASS=TD5>매출채권형태</TD>
						<TD CLASS=TD6><INPUT TYPE=TEXT NAME="txtBillType" ALT="매출채권형태" SIZE=10 MAXLENGTH=20 TAG="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBillType" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConPopUp C_PopBillType">&nbsp;<INPUT TYPE=TEXT NAME="txtBillTypeNm" SIZE=20 TAG="14"></TD>
					</TR>
					<TR>
						<TD CLASS=TD5>화폐</TD>
						<TD CLASS=TD6><INPUT TYPE=TEXT NAME="txtCurrency" ALT="화폐" SIZE=10 MAXLENGTH=3 TAG="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCurrency" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConPopUp C_PopCurrency"></TD>
						<TD CLASS="TD5" NOWRAP>출하번호</TD>
						<TD CLASS="TD6"><INPUT NAME="txtDnNo" ALT="출고번호" TYPE="Text" MAXLENGTH=18 SiZE=20 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDnNo" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConPopUp C_PopDnNo"></TD>
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
					<TD HEIGHT="100%">
						<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% TAG="23" TITLE="SPREAD" id=OBJECT1> <PARAM NAME="MaxRows" Value=0> <PARAM NAME="MaxCols" Value=0> </OBJECT>');</SCRIPT>
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
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO NORESIZE framespacing=0 TABINDEX ="-1"></IFRAME></TD>
	</TR>
</TABLE>

<INPUT TYPE=HIDDEN NAME="txtHSoldToParty" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHFromDt" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHToDt" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHSalesGrp" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHBillType" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHCurrency" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHDnNo" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>

</BODY>
</HTML>
