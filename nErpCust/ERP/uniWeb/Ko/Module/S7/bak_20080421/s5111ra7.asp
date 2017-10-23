<%@ LANGUAGE="VBSCRIPT" %>
<%
'**********************************************************************************************
'*  1. Module Name          : 영업 
'*  2. Function Name        : 매출채권관리 
'*  3. Program ID           : s5111ra4
'*  4. Program Name         : 선수금현황 
'*  5. Program Desc         : (예외)매출채권등록, 매출채권등록에서 선수금 현황 Popup
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2002/05/07
'*  8. Modified date(Last)  : 2003/05/23
'*  9. Modifier (First)     : Hwang Seongbae
'* 10. Modifier (Last)      : Hwang Seongbae
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                            2002/12/27 Include 성능향상 강준구 
'**********************************************************************************************
%>
<HTML>
<HEAD>
<TITLE>선수금현황</TITLE>

<!-- #Include file="../../inc/IncSvrCcm.inc" --> 
<!-- #Include file="../../inc/IncSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentA.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentVariables.vbs"> </SCRIPT>

<SCRIPT LANGUAGE="JavaScript" SRC="../../inc/incImage.js"></SCRIPT>
<SCRIPT LANGUAGE="VBScript">

Option Explicit

' External ASP File
'========================================
Const BIZ_PGM_ID 		= "s5111rb7.asp"                              '☆: Biz Logic ASP Name

' Constant variables 
'========================================
Const C_MaxKey          = 9                                           '☆: key count of SpreadSheet
Const C_PopPreRcptType	= 0
Const C_PopPreRcptNo	= 1

' Common variables 
'========================================
<!-- #Include file="../../inc/lgvariables.inc" -->	

' User-defind Variables
'========================================
Dim IsOpenPop  

Dim lgArrReturn												'☜: Return Parameter Group
Dim lgBlnPrRcptTypeChg
Dim lgIntStartRow

Dim arrParent
ArrParent = window.dialogArguments
Set PopupParent  = ArrParent(0)
'20021227 kangjungu dynamic popup
top.document.title = PopupParent.gActivePRAspName

Dim EndDate

' 시스템 날짜 
EndDate = UniConvDateAToB("<%=GetSvrDate%>", PopupParent.gServerDateFormat, PopupParent.gDateFormat)

'========================================
Function InitVariables()
	lgStrPrevKey     = ""								   'initializes Previous Key
	lgPageNo         = ""
    lgBlnFlgChgValue = False	                           'Indicates that no value changed
    lgIntFlgMode     = PopupParent.OPMD_CMODE                          'Indicates that current mode is Create mode
    lgSortKey        = 1   
        
    lgBlnPrRcptTypeChg = False
End Function

'========================================
Sub SetDefaultVal()
	Dim  iArrParam

	iArrParam = ArrParent(1)
	With frm1
		.txtToPrRctpDt.Text 	= iArrParam(0)				' 매출채권일 
		.txtHBillDt.Value		= iArrParam(0)				' 매출채권일 
		.txtSoldToParty.value 	= iArrParam(1)				' 주문처 
		.txtSoldToPartyNm.value	= iArrParam(2)				' 주문처 명 
		.txtDocCur.value 		= iArrParam(3)				' 화폐단위 
		.txtPrrcptNo.value		= iArrParam(4)				' 선수금번호 
	End With
		
    Redim lgArrReturn(0)        
    Self.Returnvalue = lgArrReturn     
End Sub

'========================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A( "I", "*", "NOCOOKIE", "PA") %>
	<% Call LoadBNumericFormatA("I", "*", "NOCOOKIE", "RA") %>
	
End Sub

'========================================
Sub InitSpreadSheet()
		Call SetZAdoSpreadSheet("S5111RA7","S","A","V20030523", PopupParent.C_SORT_DBAGENT, frm1.vspdData, C_MaxKey, "X", "X" )
	    Call SetSpreadLock 	  	  
End Sub

'========================================
Sub SetSpreadLock()
	ggoSpread.SpreadLockWithOddEvenRowColor()
'	ggoSpread.SpreadLock 1 , -1
	frm1.vspddata.OperationMode = 3
End Sub	

'========================================
Function OKClick()

	Dim intColCnt
		
	If frm1.vspdData.ActiveRow > 0 Then	
		Redim lgArrReturn(frm1.vspdData.MaxCols - 1)
		frm1.vspdData.Row = frm1.vspdData.ActiveRow
			
		For intColCnt = 0 To frm1.vspdData.MaxCols - 2
			frm1.vspdData.Col = GetKeyPos("A",intColCnt + 1)
			lgArrReturn(intColCnt) = frm1.vspdData.Text
		Next	
					
	End If
		
	Self.Returnvalue = lgArrReturn
	Self.Close()
	
End Function

'========================================
Function CancelClick()
	Redim lgArrReturn(0)
	lgArrReturn(0) = ""
	Self.Returnvalue = lgArrReturn
	Self.Close()
End Function

'========================================
Function OpenConPopUp(ByVal pvIntWhere)
	Dim iArrRet
	Dim iArrParam(5), iArrField(6), iArrHeader(6)

	If IsOpenPop = True Then Exit Function
	
	OpenConPopUp = False

	With frm1
		Select Case pvIntWhere
			Case C_PopPreRcptType
				iArrParam(0) = .txtPrrcptType.Alt							' 팝업 명칭 
				iArrParam(1) = "a_jnl_item"	 									' TABLE 명칭 
				iArrParam(2) = Trim(.txtPrrcptType.Value)					' Code Condition
				iArrParam(3) = ""												' Name Cindition
				iArrParam(4) = "jnl_type = " & FilterVar("PR", "''", "S") & ""								' Where Condition
				iArrParam(5) = .txtPrrcptType.Alt							' 조건필드의 라벨 명칭 

			    iArrField(0) = "JNL_CD"											' Field명(0)
			    iArrField(1) = "JNL_NM"											' Field명(1)
    
			    iArrHeader(0) = .txtPrrcptType.Alt							' Header명(0)
				iArrHeader(1) = .txtPrrcptTypeNm.Alt						' Header명(1)
				
				.txtPrrcptType.focus

			Case C_PopPreRcptNo
				iArrParam(0) = .txtPrrcptNo.Alt								' 팝업 명칭 
				iArrParam(1) = "f_prrcpt FP INNER JOIN a_jnl_item AJ ON (FP.prrcpt_type = AJ.jnl_cd)"	' TABLE 명칭 
				iArrParam(2) = Trim(.txtPrrcptNo.Value)						' Code Condition
				iArrParam(3) = ""											' Name Cindition
				' Where Condition
				iArrParam(4) = "FP.bp_cd =  " & FilterVar(.txtSoldToParty.value , "''", "S") & "" & _
							  " AND FP.doc_cur =  " & FilterVar(.txtDocCur.value , "''", "S") & "" & _
							  " AND FP.bal_amt > 0 AND FP.conf_fg = " & FilterVar("C", "''", "S") & "  AND AJ.jnl_type = " & FilterVar("PR", "''", "S") & " "
				If Len(Trim(.txtPrrcptType.value)) Then
					iArrParam(4) = iArrParam(4) & " AND FP.prrcpt_type =  " & FilterVar(.txtPrrcptType.value , "''", "S") & ""
				End If
					
				If Len(Trim(.txtFrPrRctpDt.Text)) Then
					iArrParam(4) = iArrParam(4) & " AND FP.prrcpt_dt >=  " & FilterVar(UNIConvDate(.txtFrPrRctpDt.Text), "''", "S") & ""
				End If
					
				If Len(Trim(.txtToPrRctpDt.Text)) Then
					iArrParam(4) = iArrParam(4) & " AND FP.prrcpt_dt <=  " & FilterVar(UNIConvDate(.txtToPrRctpDt.Text), "''", "S") & ""
				Else
					iArrParam(4) = iArrParam(4) & " AND FP.prrcpt_dt <=  " & FilterVar(UNIConvDate(.txtHBillDt.value), "''", "S") & ""
				End If

				iArrParam(5) = .txtPrrcptNo.Alt									' 조건필드의 라벨 명칭 

				iArrField(0) = "ED30" & PopupParent.gColSep & "FP.prrcpt_no"	' 선수금번호 
				iArrField(1) = "DD" & PopupParent.gColSep & "FP.prrcpt_dt"		' 선수금일자 
    
				iArrHeader(0) = .txtPrrcptNo.Alt								' Header명(0)
				iArrHeader(1) = "발생일자"									' Header명(1)
				
				.txtPrrcptNo.focus
		End Select
	End With

	IsOpenPop = True

	iArrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(iArrParam, iArrField, iArrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False

	If iArrRet(0) <> "" Then OpenConPopUp = SetPopUp(iArrRet, pvIntWhere)
	
End Function

'========================================
Function OpenSortPopup()
	Dim lgIsOpenPop
	Dim arrRet
	
	On Error Resume Next
	
	If lgIsOpenPop = True Then Exit Function
	lgIsOpenPop = True

	arrRet = window.showModalDialog("../../ComAsp/ZADOSortPopup.asp",Array(ggoSpread.GetXMLData("A"),gMethodText),"dialogWidth=" & PopupParent.SORTW_WIDTH & "px; dialogHeight=" & PopupParent.SORTW_HEIGHT & "px;; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False
	
	If arrRet(0) = "X" Then
	   Exit Function
	Else
	   Call ggoSpread.SaveXMLData("A",arrRet(0),arrRet(1))
       Call InitVariables
       Call InitSpreadSheet()       
   End If
End Function

'	Description : 조회조건의 유효성을 Check한다.
'   주의사항 : 화면의 tab order 별로 기술한다. 
'==========================================
Function ChkValidityQueryCon()
	Dim iStrCode

	ChkValidityQueryCon = True

	If lgBlnPrRcptTypeChg Then
		iStrCode = Trim(frm1.txtPrRcptType.value)
		If iStrCode <> "" Then
			iStrCode = " " & FilterVar(iStrCode, "''", "S") & ""
			If Not GetCodeName(iStrCode, "" & FilterVar("PR", "''", "S") & "", "default", "default", "default", "" & FilterVar("TI", "''", "S") & "", C_PopPreRcptType)  Then
				Call DisplayMsgBox("970000", "X", frm1.txtPrRcptType.alt, "X")
				frm1.txtPrRcptType.focus
				ChkValidityQueryCon = False
				Exit Function
			End If
		Else
			frm1.txtPrRcptTypeNm.value = ""
		End If
		lgBlnPrRcptTypeChg	= False
	End If

End Function

'	Description : 코드값에 해당하는 명을 Display한다.
'===========================================
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
		GetCodeName = SetPopup(iArrRs, pvIntWhere)
	Else
		' 관련 Popup Display
		'GetCodeName = OpenConPopUp(pvIntWhere)
	End if
End Function

'===========================================
Function SetPopUp(Byval pvArrRet, Byval pvIntWhere)
	SetPopup = False
	With frm1
		Select Case pvIntWhere
			Case C_PopPreRcptType
				.txtPrrcptType.value = pvArrRet(0)
				.txtPrrcptTypeNm.value = pvArrRet(1)
				.txtPrrcptType.focus
			Case C_PopPreRcptNo
				.txtPrrcptNo.value = pvArrRet(0)
				.txtPrrcptNo.focus
		End Select
	End With
	SetPopup = True
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

	DbQuery()
End Sub

'========================================
Sub Form_QueryUnload(Cancel, UnloadMode)
End Sub

'==========================================
Function txtPrRcptType_OnKeyDown()
	lgBlnPrRcptTypeChg = True
	lgBlnFlgChgValue = True
End Function

'========================================
Sub txtFrPrRctpDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtFrPrRctpDt.Action = 7		
		Call SetFocusToDocument("P")
		frm1.txtFrPrRctpDt.Focus
	End If
End Sub

'========================================
Sub txtToPrRctpDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtToPrRctpDt.Action = 7
		Call SetFocusToDocument("P")
		frm1.txtToPrRctpDt.Focus
	End If
End Sub

'========================================
Sub txtFrPrRctpDt_KeyPress(KeyAscii)
    If KeyAscii = 27 Then
       Call CancelClick()
    Elseif KeyAscii = 13 Then
       Call FncQuery()
    End if
End Sub

'========================================
Sub txtToPrRctpDt_KeyPress(KeyAscii)
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
     If KeyAscii = 13 Then
     <%If Request("txtFlag") = "B" Then%>
		Call FncQuery()
     <%Else %>
		If frm1.vspdData.ActiveRow > 0 Then
			Call OKClick()
		Else
			Call FncQuery()
		End If
	 <%End If%>
     ElseIf KeyAscii = 27 Then
        Call CancelClick()
     End If
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
Function FncQuery() 
    
    FncQuery = False                                                        '⊙: Processing is NG
    
    Err.Clear                                                               '☜: Protect system from crashing
	
	With frm1
		If ValidDateCheck(.txtFrPrRctpDt, .txtToPrRctpDt) = False Then Exit Function

		If UniConvDateToYYYYMMDD(.txtFrPrRctpDt.text , PopupParent.gDateFormat , "") > UniConvDateToYYYYMMDD(.txtHBillDt.value, PopupParent.gDateFormat , "") Then		
			Call DisplayMsgBox("970025", "X", .txtFrPrRctpDt.ALT, .txtHBillDt.alt & "(" & .txtHBillDt.value & ")")
			.txtFrPrRctpDt.Focus	
			Exit Function
		End If

		If UniConvDateToYYYYMMDD(.txtToPrRctpDt.text , PopupParent.gDateFormat , "") > UniConvDateToYYYYMMDD(.txtHBillDt.value, PopupParent.gDateFormat , "") Then		
			Call DisplayMsgBox("970025", "X", .txtToPrRctpDt.ALT, .txtHBillDt.alt & "(" & .txtHBillDt.value & ")")
			.txtToPrRctpDt.Focus	
			Exit Function
		End If
	End With
   
	' 조회조건 유효값 check
	If 	lgBlnFlgChgValue Then
		If Not ChkValidityQueryCon Then	Exit Function
	End If
	
    Call ggoOper.ClearField(Document, "2")	         						'⊙: Clear Contents  Field

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
	
	Dim iStrVal
    With frm1
		iStrVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001						'☜: 비지니스 처리 ASP의 상태	
		iStrVal = iStrVal & "&txtFrPrRctpDt=" & Trim(.txtFrPrRctpDt.Text)		'☆: 조회 조건 데이타 
		If Len(Trim(.txtToPrRctpDt.Text)) Then
			iStrVal = iStrVal & "&txtToPrRctpDt=" & Trim(.txtToPrRctpDt.Text)
		Else
			iStrVal = iStrVal & "&txtToPrRctpDt=" & Trim(.txtHBillDt.value)
		End If
		iStrVal = iStrVal & "&txtSoldToParty=" & Trim(.txtSoldToParty.value)
		iStrVal = iStrVal & "&txtDocCur=" & Trim(.txtDocCur.value)
		iStrVal = iStrVal & "&txtPrrcptType=" & Trim(.txtPrrcptType.value)
		iStrVal = iStrVal & "&txtPrrcptNo=" & Trim(.txtPrrcptNo.value)
		
        iStrVal = iStrVal & "&lgPageNo="		 & lgPageNo						'☜: Next key tag 
		iStrVal = iStrVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")			 
		iStrVal = iStrVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
		iStrVal = iStrVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))
		
		lgIntStartRow = .vspdData.MaxRows + 1
		
        Call RunMyBizASP(MyBizASP, iStrVal)		    						'☜: 비지니스 ASP 를 가동 
        
    End With
    
    
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
		Call SetFocusToDocument("P")
		frm1.txtFrPrRctpDt.focus
	End If

End Function

' 화폐별로 Cell Formating을 재설정한다.
Sub FormatSpreadCellByCurrency()
	With frm1
		Call ReFormatSpreadCellByCellByCurrency2(.vspdData,lgIntStartRow, .vspdData.MaxRows,.txtDocCur.value,GetKeyPos("A",3),"A","I","X","X") 
		Call ReFormatSpreadCellByCellByCurrency2(.vspdData,lgIntStartRow, .vspdData.MaxRows,.txtDocCur.value,GetKeyPos("A",4),"A","I","X","X") 
		Call ReFormatSpreadCellByCellByCurrency2(.vspdData,lgIntStartRow, .vspdData.MaxRows,.txtDocCur.value,GetKeyPos("A",5),"A","I","X","X") 
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
		<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
	</TR>
	<TR>
		<TD HEIGHT=20 WIDTH=100%>
			<FIELDSET CLASS="CLSFLD">
				<TABLE <%=LR_SPACE_TYPE_40%>>
					<TR>
						<TD CLASS=TD5 NOWRAP>발생기간</TD>
						<TD CLASS=TD6>
							<TABLE CELLSPACING=0 CELLPADDING=0>
								<TR>
									<TD>
										<script language =javascript src='./js/s5111ra7_fpDateTime1_txtFrPrRctpDt.js'></script>
									</TD>
									<TD>
										&nbsp;~&nbsp;
									</TD>
									<TD>
										<script language =javascript src='./js/s5111ra7_fpDateTime2_txtToPrRctpDt.js'></script>
									</TD>
								</TR>
							</TABLE>
						</TD>					
						<TD CLASS=TD5 NOWRAP>선수금번호</TD>
						<TD CLASS=TD6><INPUT TYPE=TEXT NAME="txtPrrcptNo" SIZE=27 MAXLENGTH=18 tag="11XXXU" ALT="선수금번호" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPrrcptNo" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenConPopUp C_PopPreRcptNo"></TD>
					</TR>
					<TR>
						<TD CLASS=TD5 NOWRAP>선수금유형</TD>
						<TD CLASS=TD6 NOWRAP>
							<INPUT TYPE=TEXT NAME="txtPrrcptType" SIZE=10 MAXLENGTH=10 TAG="11XXXU" ALT="선수금유형"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="txtPrrcptType" align=top TYPE="BUTTON" OnClick="vbscript:OpenConPopUp C_PopPreRcptType">
							<INPUT TYPE=TEXT NAME="txtPrrcptTypeNm" SIZE=25 TAG="14" ALT="선수금유형명">
						</TD>
						<TD CLASS="TD5" NOWRAP></TD>
						<TD CLASS="TD6" NOWRAP></TD>
					</TR>					
					<TR>
						<TD CLASS=TD5 NOWRAP><% If Right(Request("txtFlag"),1) = "H" Then%> 주문처 <%Else%> 수입자 <%End If%></TD>
						<TD CLASS=TD6><INPUT TYPE=TEXT NAME="txtSoldToParty" SIZE=10 MAXLENGTH=10 TAG="14XXXU" ALT="주문처"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSoldToParty" align=top TYPE="BUTTON">&nbsp;<INPUT TYPE=TEXT NAME="txtSoldToPartyNm" SIZE=25 TAG="14"></TD>
						<TD CLASS="TD5" NOWRAP>화폐</TD>
						<TD CLASS="TD6" NOWRAP><INPUT NAME="txtDocCur" ALT="화폐" TYPE="Text" SIZE=10 MAXLENGTH=3 tag="14XXXU" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDocCur" align=top TYPE="BUTTON"></TD>
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
						<script language =javascript src='./js/s5111ra7_vaSpread_vspdData.js'></script>
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
							                  <IMG SRC="../../../CShared/image/zpConfig_d.gif" Style="CURSOR: hand" ALT="Config" NAME="Config" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/zpConfig.gif',1)" OnClick="OpenSortPopup()" ></IMG>
					</TD>
					<TD WIDTH=30% ALIGN=RIGHT>
					<%if Left(Trim(Request("txtFlag")),1) <> "B" Then %>
							<IMG SRC="../../../CShared/image/ok_d.gif"     Style="CURSOR: hand" ALT="OK"     NAME="pop1"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/OK.gif',1)"     ONCLICK="OkClick()"    ></IMG>
					<%End If %>
							<IMG SRC="../../../CShared/image/cancel_d.gif" Style="CURSOR: hand" <%if Left(Request("txtFlag"),1) = "B" Then %> ALT="CLOSE" <%Else%> ALT="CANCEL" <%End If%> NAME="pop2"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)" ONCLICK="CancelClick()"></IMG>					</TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME></TD>
	</TR>
</TABLE>

<INPUT TYPE=HIDDEN NAME="txtHBillDt" tag="14" <%if Right(Request("txtFlag"),1) = "H" Then %> ALT="매출채권일" <%Else%> ALT="발행일" <%End If%> >
</FORM>

<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>

</BODY>
</HTML>
