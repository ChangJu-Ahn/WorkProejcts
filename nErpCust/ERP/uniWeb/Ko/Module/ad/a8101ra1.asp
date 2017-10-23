<%@ LANGUAGE="VBSCRIPT" %>
<%Response.Expires = -1%>
<!--'**********************************************************************************************
'*  1. Module Name          : Accounting
'*  2. Function Name        : 
'*  3. Program ID           : a8101ma1
'*  4. Program Name         : 본지점전표번호PopUp
'*  5. Program Desc         : 본지점전표등록에서 전표번호를 PopUp하는 ASP
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2000/12/09
'*  8. Modified date(Last)  : 2001/02/14
'*  9. Modifier (First)     : 안혜진 
'* 10. Modifier (Last)      : Hersheys
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                            
'********************************************************************************************** -->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/incSvrCcm.inc"  -->
<!-- #Include file="../../inc/incSvrHTML.inc"  -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">


<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAMain.vbs">			 </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAEvent.vbs">		 </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAOperation.vbs">	 </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs">		 </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentA.vbs">		 </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript" SRC="../../inc/incImage.js">				 </SCRIPT>
<SCRIPT LANGUAGE=vbscript>
Option Explicit	

'########################################################################################################
'#                       4.  Data Declaration Part
'########################################################################################################
'========================================================================================================
'=                       4.1 External ASP File
Const BIZ_PGM_ID        = "a8101rb1.asp"
'========================================================================================================
'=                       4.2 Constant variables For spreadsheet
Const C_MaxKey			= 1
'========================================================================================================
'=                       4.3 Common variables 
<!-- #Include file="../../inc/lgvariables.inc" -->
'========================================================================================================
'=                       4.4 User-defind Variables
Dim lgIsOpenPop											'☜: Popup status
Dim lgMark
Dim IsOpenPop											'☜: 마크 
Dim lsPoNo												'☆: Jump시 Cookie로 보낼 Grid value

Dim arrReturn
Dim arrParent
Dim arrParam

<%	
Dim lsSvrDate
lsSvrDate = GetSvrDate
%>

arrParent		= window.dialogArguments
Set PopupParent = arrParent(0)
arrParam		= arrParent(1)

top.document.title = PopupParent.gActivePRAspName

'========================================================================================================
Sub InitVariables()

    Redim arrReturn(0)

    lgBlnFlgChgValue = False 
    lgStrPrevKey     = ""   
    lgSortKey        = 1

	Self.Returnvalue = arrReturn

End Sub

'========================================================================================================
Sub SetDefaultVal()

	Dim StartDate

	StartDate = UNIDateAdd("M", -1, "<%=lsSvrDate%>", PopupParent.gServerDateFormat)
	frm1.txtfrtempgldt.Text	= UniConvDateAToB(StartDate ,PopupParent.gServerDateFormat,PopupParent.gDateFormat)
	frm1.txttotempgldt.Text	= UniConvDateAToB("<%=lsSvrDate%>" ,PopupParent.gServerDateFormat,PopupParent.gDateFormat)

End Sub

'========================================================================================================
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->

	<% Call LoadInfTB19029A("Q", "*", "NOCOOKIE", "RA") %>
	<% Call LoadBNumericFormatA("Q", "*", "NOCOOKIE", "RA") %>

End Sub

'========================================================================================================
Sub InitSpreadSheet()
	frm1.vspdData.OperationMode = 3
	Call SetZAdoSpreadSheet("A8101RA1", "S", "A", "V20021220", PopupParent.C_SORT_DBAGENT, frm1.vspdData, C_MaxKey, "X", "X")
	Call SetSpreadLock()
End Sub

'========================================================================================================
Sub SetSpreadLock()
    With frm1
		.vspdData.ReDraw = False
		ggoSpread.SpreadLockWithOddEvenRowColor()
		.vspdData.ReDraw = True
    End With
End Sub

'========================================================================================================
Sub Form_Load()
    Call LoadInfTB19029														'⊙: Load table , B_numeric_format
    Call ggoOper.FormatField(Document, "1", ggStrIntegeralPart, ggStrDeciPointPart, PopupParent.gDateFormat, PopupParent.gComNum1000, PopupParent.gComNumDec)

    Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field
    
	Call InitVariables														'⊙: Initializes local global variables
	Call SetDefaultVal
	Call InitSpreadSheet()
End Sub

'========================================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )

End Sub

'========================================================================================================
Function OKClick()
	Dim intColCnt, arrReturn

	If frm1.vspdData.MaxRows < 1 Then
	   Call CancelClick()
	   Exit Function
	End If

	If frm1.vspdData.ActiveRow > 0 Then
		Redim arrReturn(C_MaxKey - 1)

		For intColCnt = 0 To C_MaxKey - 1
			arrReturn(intColCnt) = GetKeyPosVal("A",intColCnt + 1)
		Next

		Self.Returnvalue = arrReturn
	End If

	Self.Close()
End Function

'========================================================================================================
Function CancelClick()

	Self.Close()

End Function

'========================================================================================================
Function FncQuery() 
	Dim IntRetCD

    If Not chkField(Document, "1") Then						'⊙: This function check indispensable field
		Exit Function
    End If

	If CompareDateByFormat(frm1.txtFrTempGlDt.text,frm1.txtToTempGlDt.text,frm1.txtFrTempGlDt.Alt,frm1.txtToTempGlDt.Alt, _
                        "970025",frm1.txtFrTempGlDt.UserDefinedFormat,PopupParent.gComDateType,True) = False Then
    	Exit Function
    End If

    ggoSpread.Source = frm1.vspdData
    ggoSpread.ClearSpreadData
	
	Call InitVariables										'⊙: Initializes local global variables
    Call DbQuery()											'☜: Query db data
End Function

'========================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)

    If Row = 0 Then
        ggoSpread.Source = frm1.vspdData
        If lgSortKey = 1 Then
            ggoSpread.SSSort, lgSortKey
            lgSortKey = 2
        Else
            ggoSpread.SSSort, lgSortKey
            lgSortKey = 1
        End If
    End If

	If Row < 1 Then Exit Sub

	Call SetSpreadColumnValue("A", frm1.vspdData, Col, Row)

End Sub

'========================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)

	If Frm1.vspdData.MaxRows > 0 Then
		If Frm1.vspdData.ActiveRow = Row Or Frm1.vspdData.ActiveRow > 0 Then
			Call OKClick()
		End If
	End If

End Sub

'========================================================================================================
Function vspdData_KeyPress(KeyAscii)

    If KeyAscii = 13 And frm1.vspdData.ActiveRow > 0 Then
        Call OKClick()
    ElseIf KeyAscii = 27 Then
        Call CancelClick()
    End If

End Function

'========================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then
		If lgStrPrevKey <> "" Then
			Call DbQuery
		End If
	End if

End Sub

'========================================================================================================
Function DbQuery() 

	Dim strVal

	Call LayerShowHide(1)
    
    With frm1

		strVal = BIZ_PGM_ID & "?txtfrtempgldt=" & Trim(.txtfrtempgldt.Text)
		strVal = strVal & "&txttotempgldt="     & Trim(.txttotempgldt.Text)
		strVal = strVal & "&txtfrtempglno="     & Trim(.txtfrtempglNo.value)
		strVal = strVal & "&txttotempglno="     & Trim(.txttotempglNo.value)
		strVal = strVal & "&txtdeptcd="         & Trim(.txtdeptcd.value)
        strVal = strVal & "&lgStrPrevKey="      & lgStrPrevKey
		strVal = strVal & "&lgSelectListDT="    & GetSQLSelectListDataType("A")
        strVal = strVal & "&lgTailList="        & MakeSQLGroupOrderByList("A")
		strVal = strVal & "&lgSelectList="      & EnCoding(GetSQLSelectList("A"))

    End With
    
	Call RunMyBizASP(MyBizASP, strVal)

End Function

'========================================================================================================
Function DbQueryOk()

    lgBlnFlgChgValue = True 

	If frm1.vspdData.MaxRows > 0 Then
		frm1.vspdData.Focus
	End If

End Function

'========================================================================================================
Sub txtfrtempgldt_DblClick(Button)
	if Button = 1 then
		frm1.txtfrtempgldt.Action = 7
        Call SetFocusToDocument("P")
        frm1.txtfrtempgldt.focus

	End if
End Sub

'========================================================================================================
Sub txttotempgldt_DblClick(Button)
	if Button = 1 then
		frm1.txttotempgldt.Action = 7
        Call SetFocusToDocument("P")
        frm1.txttotempgldt.focus
	End if
End Sub

'========================================================================================================
Sub txtfrtempgldt_Keypress(KeyAscii)
    On Error Resume Next
    If KeyAscii = 27 Then
       Call CancelClick()
    Elseif KeyAscii = 13 Then
       Call FncQuery()
    End if
End Sub

'========================================================================================================
Sub txttotempgldt_Keypress(KeyAscii)
    On Error Resume Next
    If KeyAscii = 27 Then
       Call CancelClick()
    Elseif KeyAscii = 13 Then
       Call FncQuery()
    End if
End Sub


 '========================================================================================================
'                        5.4 User-defined Method
'========================================================================================================
'========================================================================================================
Function OpenOrderByPopup()

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

'========================================================================================================
Function OpenPopUp(Byval strCode)

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

			arrParam(0) = "부서 팝업"				' 팝업 명칭 
			arrParam(1) = "B_ACCT_DEPT"    			' TABLE 명칭 
			arrParam(2) = strCode						' Code Condition
			arrParam(3) = ""							' Name Cindition
			arrParam(4) = "ORG_CHANGE_ID = " & FilterVar(PopupParent.gChangeOrgId, "''", "S")		' Where Condition
			arrParam(5) = "부서코드"					' 조건필드의 라벨 명칭 

			arrField(0) = "DEPT_CD"	     				' Field명(0)
			arrField(1) = "DEPT_NM"			    		' Field명(1)

			arrHeader(0) = "부서코드"					' Header명(0)
			arrHeader(1) = "부서명"				' Header명(1)
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetPopUp(arrRet)
	End If
	frm1.txtDeptCd.focus 
	
End Function

'========================================================================================================
Function SetPopUp(ByRef arrRet)

	With frm1
		.txtDeptCd.value = arrRet(0)
		.txtDeptNm.value = arrRet(1)

	End With

End Function

'========================================================================================================
Function MousePointer(pstr1)
      Select case UCase(pstr1)
            case "PON"
				window.document.search.style.cursor = "wait"
            case "POFF"
				window.document.search.style.cursor = ""
      End Select
End Function
</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->	
</HEAD>
<BODY SCROLL=NO TABINDEX="-1">
<FORM NAME=frm1 TARGET=MyBizASP METHOD=POST>
<TABLE <%=LR_SPACE_TYPE_20%>>
	<TR>
		<TD <%=HEIGHT_TYPE_02%> WIDTH="100%"></TD>
	</TR>
	<TR>
		<TD HEIGHT=20>
			<FIELDSET CLASS=CLSFLD>
				<TABLE <%=LR_SPACE_TYPE_40%>>
					<TR>				
						<TD CLASS=TD5 NOWRAP>결의일자</TD>
						<TD CLASS=TD6 NOWRAP>
							<script language =javascript src='./js/a8101ra1_I628234457_txtfrtempgldt.js'></script>&nbsp;~&nbsp;
							<script language =javascript src='./js/a8101ra1_I688155597_txttotempgldt.js'></script></TD>
						<TD CLASS=TD5 NOWRAP>결의번호</TD>
						<TD CLASS=TD6 NOWRAP>
						<INPUT TYPE=TEXT NAME=txtfrtempglNo SIZE=10 MAXLENGTH=18 STYLE="TEXT-ALIGN: left" tag="1" ALT="전표번호">&nbsp;~&nbsp;
						<INPUT TYPE=TEXT NAME=txttotempglNo SIZE=10 MAXLENGTH=18 STYLE="TEXT-ALIGN: left" tag="1" ALT="전표번호"></TD>
					</TR>
					<TR>
						<TD CLASS=TD5 NOWRAP>부서코드</TD>
						<TD CLASS=TD6 NOWRAP><INPUT NAME=txtDeptCd ALT="부서코드" MAXLENGTH=10 SIZE=10 STYLE="TEXT-ALIGN: left" tag ="11"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME=btnCostCd ALIGN=TOP TYPE=BUTTON ONCLICK="vbscript:Call OpenPopup(frm1.txtDeptCd.Value)">&nbsp;
											 <INPUT NAME=txtDeptNm ALT="부서명"   MAXLENGTH=20 SIZE=20 STYLE="TEXT-ALIGN: left" tag ="14X"></TD>
						<TD CLASS=TD5 NOWRAP></TD>
						<TD CLASS=TD6 NOWRAP></TD>
					</TR>
				</TABLE>
			</FIELDSET>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_03%> WIDTH="100%"></TD>
	</TR>
	<TR HEIGHT="100%">
		<TD WIDTH="100%" HEIGHT=* VALIGN=TOP>
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR HEIGHT="100%">
					<TD WIDTH="100%">
						<script language =javascript src='./js/a8101ra1_I540197598_vspdData.js'></script>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</TR>
	<TR HEIGHT=20>
		<TD WIDTH="100%">
			<TABLE <%=LR_SPACE_TYPE_30%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD WIDTH="70%" NOWRAP>
						<IMG SRC="../../../CShared/image/query_d.gif"    Style="CURSOR: hand" ALT="Search" NAME=Search ONMOUSEOUT="javascript:MM_swapImgRestore()" ONMOUSEOVER="javascript:MM_swapImage(this.name,'','../../../CShared/image/Query.gif',1)"    ONCLICK="FncQuery()">        </IMG>
						<IMG SRC="../../../CShared/image/zpConfig_d.gif" Style="CURSOR: hand" ALT="Config" NAME=Config ONMOUSEOUT="javascript:MM_swapImgRestore()" ONMOUSEOVER="javascript:MM_swapImage(this.name,'','../../../CShared/image/zpConfig.gif',1)" ONCLICK="OpenOrderByPopup()"></IMG></TD>
					<TD WIDTH="30%" ALIGN=RIGHT>
						<IMG SRC="../../../CShared/image/ok_d.gif"       Style="CURSOR: hand" ALT="OK"     NAME=pop1   ONMOUSEOUT="javascript:MM_swapImgRestore()" ONMOUSEOVER="javascript:MM_swapImage(this.name,'','../../../CShared/image/OK.gif',1)"       ONCLICK="OkClick()">         </IMG>
						<IMG SRC="../../../CShared/image/cancel_d.gif"   Style="CURSOR: hand" ALT="CANCEL" NAME=pop2   ONMOUSEOUT="javascript:MM_swapImgRestore()" ONMOUSEOVER="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)"   ONCLICK="CancelClick()">     </IMG></TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH="100%" HEIGHT=<%=BizSize%>>
			<IFRAME NAME=MyBizASP WIDTH="100%" HEIGHT=<%=BizSize%> SRC="../../blank.htm" FRAMEBORDER=0 SCROLLING=NO NORESIZE framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
</FORM>
<DIV ID=MousePT NAME=MousePT>
<IFRAME NAME=MouseWindow FRAMEBORDER=0 SCROLLING=NO NORESIZE FRAMESPACING=0 width=220 HEIGHT=41 SRC="../../inc/cursor.htm"></IFRAME>
