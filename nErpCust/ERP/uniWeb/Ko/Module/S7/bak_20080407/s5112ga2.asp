<%@ LANGUAGE="VBSCRIPT" %>
<%
'**********************************************************************************************
'*  1. Module Name          : 영업 
'*  2. Function Name        : 매출채권관리 
'*  3. Program ID           : S5112GA2
'*  4. Program Name         : 매출채권순위조회 
'*  5. Program Desc         : ADO Query
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2000/12/09
'*  8. Modified date(Last)  : 2001/12/18
'*							  2003/05/23
'*  9. Modifier (First)     : Byun Jee Hyun
'* 10. Modifier (Last)      : Kim Hyungsuk
'*							  Hwang Seongbae
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                            2000/12/09
'*                            2001/12/18	Date표준적용 
'*							  2003/05/23	표준적용 
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
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit

' External ASP File
'========================================
Const BIZ_PGM_ID        = "s5112gb2.asp"

' Constant variables 
'========================================
Const C_MaxKey          = 10

Const C_PopSoldToParty	= 1
Const C_PopBillType		= 2
Const C_PopSalesGrp		= 3

' Common variables 
'========================================
<!-- #Include file="../../inc/lgvariables.inc" -->	

' User-defind Variables
'========================================
Dim lgBlnOpenPop

Dim EndDate

' 시스템 날짜 
EndDate = UniConvDateAToB("<%=GetSvrDate%>", parent.gServerDateFormat, parent.gDateFormat)

'=========================================
Sub InitVariables()
    lgBlnFlgChgValue = False                               'Indicates that no value changed
    lgStrPrevKey     = ""                                  'initializes Previous Key
    lgSortKey        = 1

End Sub

'=========================================
Sub SetDefaultVal()
	frm1.txtBillFrDt.Text = UNIGetFirstDay(EndDate, Parent.gDateFormat)
	frm1.txtBillToDt.text = EndDate
	If Parent.gSalesGrp <> "" Then
		frm1.txtSalesGrp.value = Parent.gSalesGrp
		Call GetSalesGrpNm()
	End If
	frm1.txtSoldToParty.focus
End Sub

'==========================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A( "Q", "S", "NOCOOKIE", "QA") %>
End Sub

'==========================================
Sub InitSpreadSheet()
	Call SetZAdoSpreadSheet("S5112GA2","G","A","V20030711", Parent.C_GROUP_DBAGENT, frm1.vspdData, C_MaxKey, "X", "X" )
	Call SetSpreadLock 
End Sub

'=========================================
Sub SetSpreadLock()
	ggoSpread.SpreadLockWithOddEvenRowColor()
	frm1.vspdData.OperationMode = 3
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
'				iArrParam(4) = "BP.bp_type IN (" & FilterVar("C", "''", "S") & " , " & FilterVar("CS", "''", "S") & ") AND BP.usage_flag = " & FilterVar("Y", "''", "S") & " "		' Where Condition
				iArrParam(4) = "BP.bp_type IN (" & FilterVar("C", "''", "S") & " , " & FilterVar("CS", "''", "S") & ")  "		' Where Condition
					
				iArrField(0) = "ED15" & Parent.gColSep & "BP.bp_cd"	' Select Column
				iArrField(1) = "ED30" & Parent.gColSep & "BP.bp_nm"
				    
				iArrHeader(0) = .txtSoldtoParty.Alt							' Spread Title명 
				iArrHeader(1) = .txtSoldtoPartyNm.Alt
	
				frm1.txtSoldToParty.focus
				
			' 매출채권형태 
			Case C_PopBillType												
				iArrParam(1) = "s_bill_type_config"
				iArrParam(2) = Trim(.txtBillType.value)
				iArrParam(3) = ""
				iArrParam(4) = "USAGE_FLAG=" & FilterVar("Y", "''", "S") & " "

				iArrField(0) = "ED15" & Parent.gColSep & "bill_type"
				iArrField(1) = "ED30" & Parent.gColSep & "bill_type_nm"

				iArrHeader(0) = .txtBillType.Alt
				iArrHeader(1) = .txtBillTypeNm.Alt
				
				frm1.txtBillType.focus

			' 영업그룹 
			Case C_PopSalesGrp												
				iArrParam(1) = "dbo.B_SALES_GRP"
				iArrParam(2) = Trim(frm1.txtSalesGrp.value)
				iArrParam(3) = ""
				iArrParam(4) = "USAGE_FLAG=" & FilterVar("Y", "''", "S") & " "
				
				iArrField(0) = "ED15" & Parent.gColSep & "SALES_GRP"
				iArrField(1) = "ED30" & Parent.gColSep & "SALES_GRP_NM"
    
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
Sub PopZAdoConfigGrid()
	If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
		Exit Sub
	End If
	
	Call OpenGroupPopup("A")
End Sub

'========================================
Function OpenGroupPopup(ByVal pSpdNo)
	Dim arrRet

	If lgBlnOpenPop = True Then Exit Function
	lgBlnOpenPop = True
	
	arrRet = window.showModalDialog("../../ComAsp/ZADOGroupPopup.asp",Array(ggoSpread.GetXMLData("A"),gMethodText),"dialogWidth=" & parent.GROUPW_WIDTH & "px; dialogHeight=" & parent.GROUPW_HEIGHT & "px;; center: Yes; help: No; resizable: No; status: No;")

	lgBlnOpenPop = False
	
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

	With frm1
		Select Case pvIntWhere
			Case C_PopSoldToParty
				.txtSoldToParty.value = pvArrRet(0) 
				.txtSoldToPartyNm.value = pvArrRet(1)   

			Case C_PopBillType
				.txtBillType.value = pvArrRet(0) 
				.txtBillTypeNm.value = pvArrRet(1)   

			Case C_PopSalesGrp
				.txtSalesGrp.value = pvArrRet(0) 
				.txtSalesGrpNm.value = pvArrRet(1)
		End Select
	End With
	
	SetConPopup = True

End Function

'=========================================
Sub Form_Load()
    Call LoadInfTB19029														'⊙: Load table , B_numeric_format
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field
  
	Call InitVariables														'⊙: Initializes local global variables
	Call SetDefaultVal	
	Call InitSpreadSheet()
    Call SetToolbar("11000000000011")							'⊙: 버튼 툴바 제어 
End Sub

'==========================================
Sub Form_QueryUnload(Cancel , UnloadMode )
End Sub

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
		'Item Change시 명을 Fetch하는 것으로 표준 변경시 Enable 시킨다.
		'If lgBlnOpenedFlag Then GetCodeName = OpenConPopup(pvIntWhere)
	End if
End Function

'=======================================================
Sub vspdData_Click(ByVal Col, ByVal Row)

    Call SetPopupMenuItemInf("00000000001")

    gMouseClickStatus = "SPC"
    
     Set gActiveSpdSheet = frm1.vspdData

    If Row = 0 Then
		frm1.vspdData.OperationMode = 0
		ggoSpread.Source = frm1.vspdData

        If lgSortKey = 1 Then
            ggoSpread.SSSort, lgSortKey
            lgSortKey = 2
        Else
            ggoSpread.SSSort, lgSortKey
            lgSortKey = 1
        End If    
	Else
		frm1.vspdData.OperationMode = 3
    End If
    
    'Call SetSpreadColumnValue("A",frm1.vspdData,Col,Row)

End Sub

'==========================================
Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If

End Sub

'==========================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
	If OldLeft <> NewLeft Then Exit Sub

	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	    '☜: 재쿼리 체크	
		If lgStrPrevKey <> "" Then		                                                    '다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
			Call DisableToolBar(Parent.TBC_QUERY)
			Call DbQuery
		End If
	End If
End Sub

'==========================================
Sub rdoTexIssueFlg1_OnClick()
	frm1.txtRadio1.value = frm1.rdoTexIssueFlg1.value
End Sub

'==========================================
Sub rdoTexIssueFlg2_OnClick()
	frm1.txtRadio1.value = frm1.rdoTexIssueFlg2.value
End Sub

'==========================================
Sub rdoTexIssueFlg3_OnClick()
	frm1.txtRadio1.value = frm1.rdoTexIssueFlg3.value
End Sub
	
'==========================================
Sub rdoOrderFlg1_OnClick()
	frm1.txtRadio2.value = frm1.rdoOrderFlg1.value
End Sub

'==========================================
Sub rdoOrderFlg2_OnClick()
	frm1.txtRadio2.value = frm1.rdoOrderFlg2.value
End Sub
	
'==========================================
Sub txtBillFrDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtBillFrDt.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtBillFrDt.Focus
	End If
End Sub

'==========================================
Sub txtBillToDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtBillToDt.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtBillToDt.Focus
	End If
End Sub

'==========================================
Sub txtBillFrDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub

'==========================================
Sub txtBillToDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub

'==========================================
Function FncQuery() 

    FncQuery = False                                                        '⊙: Processing is NG
    
    Err.Clear                                                               '☜: Protect system from crashing

	If ValidDateCheck(frm1.txtBillFrDt, frm1.txtBillToDt) = False Then Exit Function
    
    lgIntFlgMode     = Parent.OPMD_CMODE 

    Call ggoOper.ClearField(Document, "2")									'⊙: Clear Contents  Field
    Call InitVariables 														'⊙: Initializes local global variables
    
    Call DbQuery															'☜: Query db data

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
    
    iColumnLimit  = frm1.vspdData.MaxCols
    
    ACol = Frm1.vspdData.ActiveCol
    ARow = Frm1.vspdData.ActiveRow

    If ACol > iColumnLimit Then
       iRet = DisplayMsgBox("900030", "X", iColumnLimit , "X")
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
	Dim strVal
	Dim strGroupOrderByList
    Dim arrGroupOrderByList
    DbQuery = False
    
    Err.Clear                                                               '☜: Protect system from crashing
			
		If   LayerShowHide(1) = False Then
             Exit Function 
        End If

    
    With frm1

		strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001			
		If lgIntFlgMode = Parent.OPMD_UMODE Then
			strVal = strVal & "&txtSoldToParty=" & Trim(.txtHSoldToParty.value)
			strVal = strVal & "&txtBillType=" & Trim(.txtHBillType.value)
			strVal = strVal & "&txtSalesGrp=" & Trim(.txtHSalesGrp.value)
			strVal = strVal & "&txtBillFrDt=" & Trim(.txtHBillFrDt.value)
			strVal = strVal & "&txtBillToDt=" & Trim(.txtHBillToDt.value)
			strVal = strVal & "&txtRadio1=" & Trim(.txtHRadio1.value)
		Else
			strVal = strVal & "&txtSoldToParty=" & Trim(.txtSoldToParty.value)
			strVal = strVal & "&txtBillType=" & Trim(.txtBillType.value)
			strVal = strVal & "&txtSalesGrp=" & Trim(.txtSalesGrp.value)
			strVal = strVal & "&txtBillFrDt=" & Trim(.txtBillFrDt.text)
			strVal = strVal & "&txtBillToDt=" & Trim(.txtBillToDt.text)
			strVal = strVal & "&txtRadio1=" & Trim(.txtRadio1.value)		
		End If
        strVal = strVal & "&lgStrPrevKey="   & lgStrPrevKey                      '☜: Next key tag
		strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")
        
        arrGroupOrderByList = split(MakeSQLGroupOrderByList("A"),"Order")
        strGroupOrderByList = arrGroupOrderByList(0)
     
        If frm1.rdoOrderFlg1.checked Then
			strGroupOrderByList = strGroupOrderByList & " ORDER BY SUM(CASE WHEN B.BILL_AMT_LOC >= 0 THEN B.BILL_QTY ELSE  - B.BILL_QTY END) DESC  "
        ElseIf frm1.rdoOrderFlg2.checked Then
        	strGroupOrderByList = strGroupOrderByList & " ORDER BY SUM(B.BILL_AMT_LOC) DESC "
		End If		
	    
        strVal = strVal & "&lgTailList="     & strGroupOrderByList
		strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))
       
        Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 
    End With
    
    DbQuery = True
End Function

'========================================
Function DbQueryOk()														'☆: 조회 성공후 실행로직 

	Call SetToolbar("11000000000111")							'⊙: 버튼 툴바 제어 
	If frm1.vspdData.MaxRows > 0 Then
		frm1.vspdData.Focus
		frm1.vspdData.SelModeSelected = True
		If lgIntFlgMode <> Parent.OPMD_UMODE Then
			frm1.vspdData.Row = 1
			lgIntFlgMode = Parent.OPMD_UMODE		
		End If
    Else
       frm1.txtSoldToParty.focus
    End If   
End Function

'========================================
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>매출채권순위</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* >&nbsp;</TD>
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
									<TD CLASS="TD5" NOWRAP>주문처</TD>
									<TD CLASS="TD6"><INPUT NAME="txtSoldToParty" ALT="주문처" TYPE="Text" MAXLENGTH=10 SiZE=10 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSItemDC" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConPopup C_PopSoldToParty ">&nbsp;<INPUT NAME="txtSoldToPartyNm" TYPE="Text" SIZE=20 tag="14" Alt="주문처명"></TD>
									<TD CLASS=TD5>매출채권형태</TD>
									<TD CLASS=TD6><INPUT TYPE=TEXT NAME="txtBillType" SIZE=10 MAXLENGTH=20 TAG="11XXXU" ALT="매출채권형태"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSORef" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConPopup C_PopBillType">&nbsp;<INPUT TYPE=TEXT NAME="txtBillTypeNm" SIZE=20 TAG="14" Alt="매출채권형태명"></TD>
								</TR>	
								<TR>	
									<TD CLASS=TD5 NOWRAP>영업그룹</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtSalesGrp" SIZE=10 MAXLENGTH=4 TAG="11XXXU" ALT="영업그룹"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSalesGrp" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConPopup C_PopSalesGrp">&nbsp;<INPUT TYPE=TEXT NAME="txtSalesGrpNm" SIZE=20 TAG="14" Alt="영업그룹명"></TD>
									<TD CLASS=TD5 NOWRAP>매출채권일</TD>
									<TD CLASS=TD6 NOWRAP>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 NAME="txtBillFrDt" CLASS="FPDTYYYYMMDD" tag="11X1" Title="FPDATETIME" ALT="매출채권시작일"></OBJECT>');</SCRIPT>&nbsp;~&nbsp;
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime2 NAME="txtBillToDt" CLASS="FPDTYYYYMMDD" tag="11X1" Title="FPDATETIME" ALT="매출채권종료일"></OBJECT>');</SCRIPT>
									</TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>확정여부</TD> 
									<TD CLASS=TD6 NOWRAP>
										<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoTexIssueFlg" TAG="11X" VALUE="A" CHECKED ID="rdoTexIssueFlg1"><LABEL FOR="rdoTexIssueFlg1">전체</LABEL>&nbsp;&nbsp;&nbsp;
										<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoTexIssueFlg" TAG="11X" VALUE="Y" ID="rdoTexIssueFlg2"><LABEL FOR="rdoTexIssueFlg2">확정</LABEL>&nbsp;&nbsp;&nbsp;
										<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoTexIssueFlg" TAG="11X" VALUE="N" ID="rdoTexIssueFlg3"><LABEL FOR="rdoTexIssueFlg3">미확정</LABEL>			
									<TD CLASS=TD5 NOWRAP>순위결정기준</TD> 
									<TD CLASS=TD6 NOWRAP>
										<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoOrderFlg" TAG="11X" VALUE="Q" CHECKED ID="rdoOrderFlg1"><LABEL FOR="rdoOrderFlg1">매출채권량</LABEL>&nbsp;&nbsp;&nbsp;
										<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoOrderFlg" TAG="11X" VALUE="M" ID="rdoOrderFlg2"><LABEL FOR="rdoOrderFlg2">매출채권자국금액</LABEL>&nbsp;&nbsp;&nbsp;
									</TD>
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
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>

<INPUT TYPE=HIDDEN NAME="txtRadio1" tag="14" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtRadio2" tag="14" TABINDEX="-1">

<INPUT TYPE=HIDDEN NAME="txtHSoldToParty" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHBillType" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHSalesGrp" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHBillFrDt" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHBillToDt" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHRadio1" tag="14" TABINDEX="-1">

</FORM>

<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm" TABINDEX="-1"></iframe>
</DIV>

</BODY>
</HTML>
