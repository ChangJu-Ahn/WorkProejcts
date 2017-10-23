<%@ LANGUAGE="VBSCRIPT" %>
<%
'**********************************************************************************************
'*  1. Module Name          : 영업 
'*  2. Function Name        : 매출채권관리 
'*  3. Program ID           : S5111RA2
'*  4. Program Name         : 이전매출채권참조 
'*  5. Program Desc         : ADO Query
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2000/12/09
'*  8. Modified date(Last)  : 2001/12/18
'*  9. Modifier (First)     : Byun Jee Hyun
'* 10. Modifier (Last)      : Kim Hyungsuk
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                            2000/12/09
'*                            2001/12/18	Date표준적용 
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
Const BIZ_PGM_ID        = "s5111rb2.asp"

' Constant variables 
'========================================
Const C_MaxKey          = 5                                    '☆☆☆☆: Max key value

Const C_PopSoldToParty	= 1
Const C_PopBillToParty	= 2
Const C_PopSalesGrp		= 3

' Common variables 
'========================================
<!-- #Include file="../../inc/lgvariables.inc" -->

' User-defind Variables
'========================================
Dim lgBlnOpenPop			' Popup의 Open여부 
Dim lgIntStartRow
Dim strReturn					<% '--- Return Parameter Group %>
Dim lgBlnOpenedFlag

Dim arrParent

arrParent = window.dialogArguments
Set PopupParent = arrParent(0)

top.document.title = PopupParent.gActivePRAspName

Dim EndDate

' 시스템 날짜 
EndDate = UniConvDateAToB("<%=GetSvrDate%>", PopupParent.gServerDateFormat, PopupParent.gDateFormat)

'========================================
Sub InitVariables()
    lgBlnFlgChgValue = False                               'Indicates that no value changed
    lgPageNo     = ""                                  'initializes Previous Key
    lgSortKey        = 1
	lgIntFlgMode     = PopupParent.OPMD_CMODE                          'Indicates that current mode is Create mode

	Redim strReturn(1)
	strReturn(0) = ""
	Self.Returnvalue = strReturn
End Sub

'=========================================
Sub SetDefaultVal()
	frm1.txtBillFrDt.text = UNIGetFirstDay(EndDate, PopupParent.gDateFormat)
	frm1.txtBillToDt.text = EndDate
	If PopupParent.gSalesGrp <> "" Then
		frm1.txtSalesGrp.value = PopupParent.gSalesGrp
		Call GetSalesGrpNm()
	End If
End Sub

'==========================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"-->
	<% Call LoadInfTB19029A("I", "*", "NOCOOKIE", "RA") %>	
	<% Call LoadBNumericFormatA("I", "*", "NOCOOKIE", "RA") %>
End Sub

'==========================================
Sub InitSpreadSheet()

	Call SetZAdoSpreadSheet("S5111RA2","S","A","V20030301",PopupParent.C_SORT_DBAGENT,frm1.vspdData, _
								C_MaxKey, "X","X")		
	Call SetSpreadLock 	
   
End Sub

'=========================================
Sub SetSpreadLock()
'	ggoSpread.SpreadLock 1 , -1
	ggoSpread.SpreadLockWithOddEvenRowColor()
	frm1.vspddata.OperationMode = 3
End Sub

'========================================
Function OpenConPopup(ByVal pvIntWhere)

	Dim iArrRet
	Dim iArrParam(5), iArrField(6), iArrHeader(6)

	OpenConPopup = False
	
	If lgBlnOpenPop = True Then Exit Function

	lgBlnOpenPop = True

	With frm1
		Select Case pvIntWhere
			' 주문처 
			Case C_PopSoldToParty												
				iArrParam(1) = "dbo.b_biz_partner BP"				' TABLE 명칭 
				iArrParam(2) = Trim(frm1.txtSoldToParty.value)		' Code Condition
				iArrParam(3) = ""									' Name Cindition
				iArrParam(4) = "BP.bp_type IN (" & FilterVar("C", "''", "S") & " , " & FilterVar("CS", "''", "S") & ") AND BP.usage_flag = " & FilterVar("Y", "''", "S") & " "		' Where Condition
					
				iArrField(0) = "ED15" & PopupParent.gColSep & "BP.bp_cd"	' Select Column
				iArrField(1) = "ED30" & PopupParent.gColSep & "BP.bp_nm"
				    
				iArrHeader(0) = .txtSoldtoParty.Alt							' Spread Title명 
				iArrHeader(1) = .txtSoldtoPartyNm.Alt
	
				frm1.txtSoldToParty.focus
			
			' 발행처 
			Case C_PopBillToParty												
				iArrParam(1) = "dbo.b_biz_partner BP"			<%' TABLE 명칭 %>
				iArrParam(2) = Trim(.txtBillToParty.value)	<%' Code Condition%>
				iArrParam(3) = ""								<%' Name Cindition%>
				iArrParam(4) = "EXISTS (SELECT * FROM dbo.b_biz_partner_ftn BF WHERE BP.bp_cd = BF.partner_bp_cd AND BF.partner_ftn = " & FilterVar("SBI", "''", "S") & ") " & _
							   "AND BP.usage_flag = " & FilterVar("Y", "''", "S") & " "		<%' Where Condition%>
					
				iArrField(0) = "ED15" & PopupParent.gColSep & "BP.bp_cd"	<%' Field명(0)%>
				iArrField(1) = "ED30" & PopupParent.gColSep & "BP.bp_nm"	<%' Field명(1)%>
				    
				iArrHeader(0) = .txtBillToParty.Alt
				iArrHeader(1) = .txtBillToPartyNm.Alt

				frm1.txtBillToParty.focus
				
			' 영업그룹 
			Case C_PopSalesGrp												
				iArrParam(1) = "dbo.B_SALES_GRP"
				iArrParam(2) = Trim(frm1.txtSalesGrp.value)
				iArrParam(3) = ""
				iArrParam(4) = "USAGE_FLAG=" & FilterVar("Y", "''", "S") & " "
				
				iArrField(0) = "ED15" & PopupParent.gColSep & "SALES_GRP"
				iArrField(1) = "ED30" & PopupParent.gColSep & "SALES_GRP_NM"
    
			    iArrHeader(0) = .txtSalesGrp.Alt
			    iArrHeader(1) = .txtSalesGrpNm.Alt
			    
			    frm1.txtSalesGrp.focus
		End Select
	End With
 
	iArrParam(0) = iArrHeader(0)							' 팝업 Title
	iArrParam(5) = iArrHeader(0)							' 조회조건 명칭 

	iArrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(iArrParam, iArrField, iArrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgBlnOpenPop = False

	If iArrRet(0) <> "" Then
		OpenConPopup = SetConPopup(iArrRet,pvIntWhere)
	End If	
	
End Function

'========================================
Function OpenSortPopup()
	Dim arrRet
	
	On Error Resume Next 
	
	If lgBlnOpenPop = True Then Exit Function
	lgBlnOpenPop = True

	arrRet = window.showModalDialog("../../ComAsp/ZADOSortPopup.asp",Array(ggoSpread.GetXMLData("A"),gMethodText),"dialogWidth=" & PopupParent.SORTW_WIDTH & "px; dialogHeight=" & PopupParent.SORTW_HEIGHT & "px;; center: Yes; help: No; resizable: No; status: No;")

	lgBlnOpenPop = False
	
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

	With frm1
		Select Case pvIntWhere
			Case C_PopSoldToParty
				.txtSoldToParty.value = pvArrRet(0) 
				.txtSoldToPartyNm.value = pvArrRet(1)   

			Case C_PopBillToParty
				.txtBillToParty.value = pvArrRet(0) 
				.txtBillToPartyNm.value = pvArrRet(1)   

			Case C_PopSalesGrp
				.txtSalesGrp.value = pvArrRet(0) 
				.txtSalesGrpNm.value = pvArrRet(1)
		End Select
	End With
	
	SetConPopup = True

End Function

'========================================
Sub Form_Load()
	
    Call LoadInfTB19029														'⊙: Load table , B_numeric_format
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec)
    Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field
    
	Call InitVariables														'⊙: Initializes local global variables
	Call SetDefaultVal	
	Call InitSpreadSheet()
	Call MM_preloadImages("../../../CShared/image/Query.gif","../../../CShared/image/OK.gif","../../../CShared/image/Cancel.gif")
	lgBlnOpenedFlag = True
	Call FncQuery()
End Sub

'==========================================
Sub Form_QueryUnload(Cancel , UnloadMode )
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

'	Description : 코드값에 해당하는 명을 Display한다.
'====================================================
Function GetCodeName(ByVal pvStrArg1, ByVal pvStrArg2, ByVal pvStrArg3, ByVal pvStrArg4, ByVal pvIntArg5, ByVal pvStrFlag, ByVal pvIntWhere)

	Dim iStrSelectList, iStrFromList, iStrWhereList
	Dim iStrRs
	Dim iArrTemp
	
	GetCodeName = False
	
	iStrSelectList = " * "
	iStrFromList = " dbo.ufn_s_GetCodeName (" & pvStrArg1 & ", " & pvStrArg2 & ", " & pvStrArg3 & ", " & pvStrArg4 & ", " & pvIntArg5 & ", " & pvStrFlag & ") "
	iStrWhereList = ""
	
	Err.Clear
	
	If CommonQueryRs2by2(iStrSelectList,iStrFromList,iStrWhereList,iStrRs) Then
		iArrTemp = Split(iStrRs, Chr(11))
		frm1.txtSalesGrp.value = iArrTemp(1)
		frm1.txtSalesGrpNm.value = iArrTemp(2)
		GetCodeName = True
	Else
		'Item Change시 명을 Fetch하는 것으로 표준 변경시 Enable 시킨다.
		'If lgBlnOpenedFlag Then GetCodeName = OpenConPopup(pvIntWhere)
	End if
End Function

'=======================================================
Sub txtBillFrDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtBillFrDt.Action = 7 
		Call SetFocusToDocument("P")
		frm1.txtBillFrDt.Focus
    End If
End Sub

'=======================================================
Sub txtBillToDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtBillToDt.Action = 7 
		Call SetFocusToDocument("P")
		frm1.txtBillToDt.Focus
    End If
End Sub

'==========================================
Sub txtBillFrDt_Keypress(KeyAscii)
	If KeyAscii = 27 Then
		Call CancelClick()
	Elseif KeyAscii = 13 Then
		Call FncQuery()
	End if
End Sub

'==========================================
Sub txtBillToDt_Keypress(KeyAscii)
	If KeyAscii = 27 Then
		Call CancelClick()
	Elseif KeyAscii = 13 Then
		Call FncQuery()
	End if
End Sub

'========================================
Function vspdData_KeyPress(KeyAscii)
	If KeyAscii = 13 And frm1.vspdData.ActiveRow > 0 Then
		Call OKClick()
	ElseIf KeyAscii = 27 Then
		Call CancelClick()
	End If
End Function

'========================================
Function vspdData_DblClick(ByVal Col, ByVal Row)
	If Frm1.vspdData.ActiveRow > 0 Then	Call OKClick
End Function
	
'========================================
Sub vspdData_TopLeftChange(ByVal OldLeft, ByVal OldTop, ByVal NewLeft, ByVal NewTop)
	If OldLeft <> NewLeft Then	Exit Sub

	if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	'☜: 재쿼리 체크'
		If lgPageNo <> "" Then								<% '⊙: 다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 %>
			If CheckRunningBizProcess = True Then Exit Sub
			Call DBQuery
		End if	    
	End if	    

End Sub

'========================================
Function OKClick()
	Dim iStrReturn

	Redim iStrReturn(2)
	
	With frm1.vspdData
		If .ActiveRow > 0 Then	
			.Row = .ActiveRow
			.Col = GetKeyPos("A",1)		:	iStrReturn(0) = Trim(.Text)
			.Col = GetKeyPos("A",2)		:	iStrReturn(1) = Trim(.Text)
			.Col = GetKeyPos("A",3)		:	iStrReturn(2) = Trim(.Text)
		End If
	End With
	
	Self.Returnvalue = iStrReturn
	Self.Close()
End Function

'========================================
Function CancelClick()
	Self.Close()
End Function

'========================================
Function FncQuery() 

    FncQuery = False                                                        '⊙: Processing is NG
    
    Err.Clear                                                               '☜: Protect system from crashing

	With frm1
		If ValidDateCheck(frm1.txtBillFrDt, frm1.txtBillToDt) = False Then Exit Function

		If UniConvDateToYYYYMMDD(.txtBillFrDt.text , PopupParent.gDateFormat , "") > UniConvDateToYYYYMMDD(EndDate, PopupParent.gDateFormat , "") Then		
			Call DisplayMsgBox("970025", "X", .txtBillFrDt.ALT, "현재일" & "(" & EndDate & ")")
			.txtBillFrDt.focus	
			Exit Function
		End If

		If UniConvDateToYYYYMMDD(.txtBillToDt.text , PopupParent.gDateFormat , "") > UniConvDateToYYYYMMDD(EndDate, PopupParent.gDateFormat , "") Then		
			Call DisplayMsgBox("970025", "X", .txtBillToDt.ALT, "현재일" & "(" & EndDate & ")")	
			.txtBillToDt.Focus
			Exit Function
		End If
	End With

    Call ggoOper.ClearField(Document, "2")									'⊙: Clear Contents  Field
    Call InitVariables 														'⊙: Initializes local global variables
    
    Call DbQuery															'☜: Query db data

    FncQuery = True		
End Function

'========================================
Function DbQuery() 
	Dim iStrVal

    DbQuery = False
    
    Err.Clear                                                               '☜: Protect system from crashing
			
	If LayerShowHide(1) = False Then Exit Function 
    
    With frm1

		iStrVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001
		If lgIntFlgMode = PopupParent.OPMD_UMODE Then
			' Scroll시 
			iStrVal = iStrVal & "&txtSoldtoParty=" & Trim(.txtHSoldtoParty.value)
			iStrVal = iStrVal & "&txtBillToParty=" & Trim(.txtHBillToParty.value)
			iStrVal = iStrVal & "&txtSalesGrp=" & Trim(.txtHSalesGrp.value)
			iStrVal = iStrVal & "&txtBillFrDt=" & Trim(.txtHBillFrDt.value)
			iStrVal = iStrVal & "&txtBillToDt=" & Trim(.txtHBillToDt.value)
		Else
			' 처음 조회시 
			iStrVal = iStrVal & "&txtSoldtoParty=" & Trim(.txtSoldtoParty.value)
			iStrVal = iStrVal & "&txtBillToParty=" & Trim(.txtBillToParty.value)
			iStrVal = iStrVal & "&txtSalesGrp=" & Trim(.txtSalesGrp.value)
			iStrVal = iStrVal & "&txtBillFrDt=" & Trim(.txtBillFrDt.text)
			If Len(Trim(.txtBillToDt.text)) Then
				iStrVal = iStrVal & "&txtBillToDt=" & Trim(.txtBillToDt.text)
			Else
				iStrVal = iStrVal & "&txtBillToDt=" & EndDate
			End if
		End If
		
        iStrVal = iStrVal & "&lgPageNo="   & lgPageNo                      '☜: Next key tag
		iStrVal = iStrVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")
		
        iStrVal = iStrVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
		iStrVal = iStrVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))

		lgIntStartRow = .vspdData.MaxRows + 1
   
        Call RunMyBizASP(MyBizASP, iStrVal)										'☜: 비지니스 ASP 를 가동 
    End With
    
    DbQuery = True
End Function

'========================================
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
		Call ReFormatSpreadCellByCellByCurrency(.vspdData,lgIntStartRow, .vspdData.MaxRows,GetKeyPos("A",4),GetKeyPos("A",5),"A","I","X","X") 
	End With
End Sub

'========================================
</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>
<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_20%>>
	<TR>
		<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
	</TR>
	<TR>
		<TD HEIGHT=20 WIDTH=100%>
			<FIELDSET CLASS="CLSFLD">
				<TABLE <%=LR_SPACE_TYPE_40%>>
					<TR>
						<TD CLASS=TD5 NOWRAP>주문처</TD>
						<TD CLASS=TD6 NOWRAP>
							<INPUT TYPE=TEXT NAME="txtSoldtoParty" SIZE=10 MAXLENGTH=10 TAG="11XXXU" ALT="주문처"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSoldtoParty" align=top TYPE="BUTTON" ONCLICK ="vbscript:OpenConPopUp C_PopSoldToParty ">&nbsp;
							<INPUT TYPE=TEXT NAME="txtSoldtoPartyNm" SIZE=20 TAG="14" ALT="주문처명">
						</TD>
						<TD CLASS=TD5 NOWRAP>발행처</TD>
						<TD CLASS=TD6 NOWRAP>
							<INPUT TYPE=TEXT NAME="txtBillToParty" SIZE=10 MAXLENGTH=10 TAG="11XXXU" ALT="발행처"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBillToParty" align=top TYPE="BUTTON" ONCLICK ="vbscript:OpenConPopUp C_PopBillToParty">&nbsp;
							<INPUT TYPE=TEXT NAME="txtBillToPartyNm" SIZE=20 TAG="14" ALT="발행처명">
						</TD>
					</TR>
					<TR>
						<TD CLASS=TD5 NOWRAP>매출채권일</TD>
						<TD CLASS=TD6 NOWRAP>
							<script language =javascript src='./js/s5111ra2_fpDateTime1_txtBillFrDt.js'></script>&nbsp;~&nbsp;
							<script language =javascript src='./js/s5111ra2_fpDateTime2_txtBillToDt.js'></script>
						</TD>
						<TD CLASS=TD5 NOWRAP>영업그룹</TD>
						<TD CLASS=TD6 NOWRAP>
							<INPUT TYPE=TEXT NAME="txtSalesGrp" SIZE=10 MAXLENGTH=4 TAG="11XXXU" ALT="영업그룹"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSalesGrp" align=top TYPE="BUTTON" ONCLICK ="OpenConPopUp C_PopSalesGrp">&nbsp;
							<INPUT TYPE=TEXT NAME="txtSalesGrpNm" SIZE=20 TAG="14" ALT="영업그룹명">
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
					<TD HEIGHT="100%">
						<script language =javascript src='./js/s5111ra2_vaSpread_vspdData.js'></script>
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

<INPUT TYPE=HIDDEN NAME="txtHBillFrDt" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHBillToDt" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHSoldToParty" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHBillToParty" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHSalesGrp" tag="24">

</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
