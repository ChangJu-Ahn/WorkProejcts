
<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name       : 경영손익 
'*  2. Function Name     : 
'*  3. Program ID        : GH0080a1
'*  4. Program Name      : 품목그룹 손익추이표출력 
'*  5. Program Desc      : 품목그룹 손익추이표출력 
'*  6. Comproxy 리스트   : 
'*  7. 최초 작성년월일   : 2001/12/18
'*  8. 최종 수정년월일   : 2001/12/18
'*  9. 최초 작성자       : 이 강 영 
'* 10. 최종 작성자       : 이 강 영 
'* 11. 전체 comment      :
'**********************************************************************************************-->

<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!--
########################################################################################################
#						   3.    External File Include Part
########################################################################################################-->

<!--
========================================================================================================
=                          3.1 Server Side Script
========================================================================================================-->
<!-- #Include file="../../inc/incSvrCcm.inc"  -->
<!-- #Include file="../../inc/IncSvrHTML.inc"  -->

<!--
========================================================================================================
=                          3.2 Style Sheet
======================================================================================================== -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		

<!--
========================================================================================================
=                          3.3 Client Side Script
======================================================================================================== -->
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAMain.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAEvent.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliVariables.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAOperation.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliRdsQuery.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incEB.vbs"></SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>
Option Explicit                                                        '☜: indicates that All variables must be declared in advance

'########################################################################################################
'#                       4.  Data Declaration Part
'########################################################################################################

'========================================================================================================
'=                       4.1 External ASP File
'========================================================================================================
'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim IsOpenPop
Dim lgOldRow

'########################################################################################################
'#                       5.Method Declaration Part
'########################################################################################################

'========================================================================================================
'========================================================================================================
'                        5.1 Common Method-1
'========================================================================================================
'========================================================================================================

'========================================================================================================
' Name : InitVariables()	
' Desc : Initialize value
'========================================================================================================
Sub InitVariables()

    lgIntFlgMode = Parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
         
End Sub

'========================================================================================================
' Name : SetDefaultVal()	
' Desc : Set default value
'========================================================================================================
Sub SetDefaultVal()
	Dim StartDate
	StartDate = "<%=GetSvrDate%>"
	
	frm1.txtYyyymm.focus()    
	frm1.txtYyyymm.Text	= UniConvDateAToB(StartDate ,parent.gServerDateFormat,parent.gDateFormat) 
    Call ggoOper.FormatDate(frm1.txtYyyymm, Parent.gDateFormat, 3) 
End Sub

'========================================================================================================
' Name : LoadInfTB19029()	
' Desc : Set System Number format
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 

	<% Call loadInfTB19029A("P", "G", "NOCOOKIE", "PA") %>

	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub
'========================================================================================================
' Name : InitComboBox()
' Desc : Set ComboBox
'========================================================================================================
'Sub InitComboBox()

'End Sub
'========================================================================================================
'========================================================================================================
'                        5.2 Common Method-2
'========================================================================================================
'========================================================================================================

'========================================================================================================
' Name : Form_Load
' Desc : developer describe this line Called by Window_OnLoad() evnt
'========================================================================================================
Sub Form_Load()

    Call LoadInfTB19029                                                     '⊙: Load table , B_numeric_format
    Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
    Call InitVariables 
    
'    Call ggoOper.FormatDate(frm1.txtpay_yymm, Parent.gDateFormat, 2)                    '싱글에서 년월말 입력하고 싶은경우 다음 함수를 콜한다.
    
    Call SetDefaultVal
    Call SetToolbar("1000000000000111")

End Sub

'========================================================================================================
' Name : Form_QueryUnload
' Desc : developer describe this line Called by Window_OnUnLoad() evnt
'========================================================================================================
Sub Form_QueryUnload(Cancel, UnloadMode)

End Sub
'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================
Function FncQuery()
    
    FncQuery = False															 '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
    
    FncQuery = True                                                              '☜: Processing is OK

End Function
'========================================================================================
' Function Name : txtGrade_onKeyPress
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================
Function txtGrade_onKeyPress(Key)    
    
    frm1.action = "../../blank.htm"       
    
End Function
	
'======================================================================================================
' Function Name : FncBtnPrint
' Function Desc : This function is related to Print Button
'=======================================================================================================

Function FncBtnPrint() 
	Dim condvar
	Dim lngPos
	Dim intCnt
    Dim StrEbrFile
    Dim var1,var2
    Dim strYear, strMonth, strDay
    	
    If Not chkField(Document, "1") Then									<%'⊙: This function check indispensable field%>
       Exit Function
    End If
	
	StrEbrFile = "ge008oa1"
	ObjName = AskEBDocumentName(StrEBrFile, "ebr")
		
'	Call ExtractDateFrom(frm1.txtYyyymm.Text,frm1.txtYyyymm.UserDefinedFormat,Parent.gComDateType,strYear,strMonth,strDay)
	
	var1 = Trim(frm1.txtYyyymm.Text)
	var2 = Trim(UCase(frm1.txtFr_dept_cd.value))	
	
    if var2 = "" then
		var2 = "%"
		frm1.txtFr_dept_nm.value = ""
	else
		Call CommonQueryRs(" ITEM_GROUP_NM "," B_ITEM_GROUP "," ITEM_GROUP_CD =   " & FilterVar(frm1.txtFr_dept_cd.value, "''", "S") & "" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
		if Trim(Replace(lgF0,Chr(11),"")) = "X" then
		  frm1.txtFr_dept_nm.value = ""
		else   
		  frm1.txtFr_dept_nm.value = Trim(Replace(lgF0,Chr(11),""))
		end if    	    		
	End if	
	
    <%'--출력조건을 지정하는 부분 수정 %>
	
	condvar = "YYYY|" & var1
	condvar = condvar & "|ITEM_GROUP_CD|" & var2	
	
	Call FncEBRPrint(EBAction,ObjName,condvar)				

End Function


'========================================================================================
' Function Name : FncBtnPreview()
' Function Desc : This function is related to Preview Button
'========================================================================================
Function FncBtnPreview()
'On Error Resume Next                                                    '☜: Protect system from crashing
    
    If Not chkField(Document, "1") Then									<%'⊙: This function check indispensable field%>
       Exit Function
    End If
	
	dim condvar
	dim arrParam, arrField, arrHeader
    Dim StrEbrFile
	Dim var1, var2	
    Dim strYear, strMonth, strDay

	StrEbrFile = "ge008oa1"
	ObjName = AskEBDocumentName(StrEBrFile, "ebr")
		
'	Call ExtractDateFrom(frm1.txtYyyymm.Text,frm1.txtYyyymm.UserDefinedFormat,Parent.gComDateType,strYear,strMonth,strDay)
	
	var1 = Trim(frm1.txtYyyymm.Text)
	var2 = Trim(UCase(frm1.txtFr_dept_cd.value))	

    if var2 = "" then
		var2 = "%"
		frm1.txtFr_dept_nm.value = ""
	else
		Call CommonQueryRs(" ITEM_GROUP_NM "," B_ITEM_GROUP "," ITEM_GROUP_CD =   " & FilterVar(frm1.txtFr_dept_cd.value, "''", "S") & "" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
		if Trim(Replace(lgF0,Chr(11),"")) = "X" then
		  frm1.txtFr_dept_nm.value = ""
		else   
		  frm1.txtFr_dept_nm.value = Trim(Replace(lgF0,Chr(11),""))
		end if    	    		
	End if	

	condvar = "YYYY|" & var1
	condvar = condvar & "|ITEM_GROUP_CD|" & var2

	Call FncEBRPreview(ObjName,condvar)

End Function

'========================================================================================================
' Name : FncPrint
' Desc : developer describe this line Called by MainDeleteRow in Common.vbs
'========================================================================================================
Function FncPrint()
	Call FncPrint()                                                      '☜: Protect system from crashing
End Function

'========================================================================================================
' Name : FncFind
' Desc : developer describe this line Called by MainFind in Common.vbs
'========================================================================================================
Function FncFind() 
	Call FncFind(Parent.C_SINGLE, False)
End Function

'========================================================================================================
' Name : FncExit
' Desc : developer describe this line Called by MainExit in Common.vbs
'========================================================================================================
Function FncExit()
	FncExit = True
End Function

'========================================================================================================
'========================================================================================================
'                        5.4 User-defined Method
'========================================================================================================
Function OpenPopUp()
Dim arrRet
Dim arrParam(5), arrField(6), arrHeader(6)
dim strgChangeOrgId

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True



	arrParam(0) = "품목그룹"		    	<%' 팝업 명칭 %>
	arrParam(1) = "(select distinct item_group_cd as item_group_cd from g_item_group_profit where substring(yyyymm,1,4) = " & FilterVar(frm1.txtYyyymm.Text, "''", "S") & ") a left outer join b_item_group b on a.item_group_cd = b.item_group_cd " <%' TABLE 명칭 %>
	arrParam(2) = UCase(Trim(frm1.txtFr_dept_cd.Value))                        <%' Code Condition%>
	arrParam(3) = "" 		            	<%' Name Cindition%>
	arrParam(4) = ""                        <%' Where Condition%>
	arrParam(5) = "품목그룹"

    arrField(0) = "a.ITEM_GROUP_CD"	     			<%' Field명(1)%>
	arrField(1) = "case when b.item_group_nm is null then '' else b.ITEM_GROUP_NM end"					<%' Field명(0)%>

    arrHeader(0) = "품목그룹코드"			    	<%' Header명(0)%>
    arrHeader(1) = "품목그룹명"				<%' Header명(1)%>

    
    
	arrRet = window.showModalDialog("../../comasp/AdoCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
	   Frm1.txtFr_dept_cd.value = arrRet(0)
	   frm1.txtFr_dept_nm.value = arrRet(1)
	End If	

End Function




'========================================================================================================
' Name : txtPay_yymm_DblClick
' Desc : 달력 Popup을 호출 
'========================================================================================================
Sub txtYyyymm_DblClick(Button)
	If Button = 1 Then
		frm1.txtYyyymm.Action = 7
	End If
End Sub
'========================================================================================================
'========================================================================================================
'                        5.5 Tag Event
'========================================================================================================
'========================================================================================================
'☜: 아래 OBJECT Tag는 InterDev 사용자를 위한것으로 프로그램이 완성되면 아래 Include 코드로 대체되어야 한다 
</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>

<BODY TABINDEX="-1" SCROLL="NO">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE CLASS="BatchTB1" CELLSPACING=0 CELLPADDING=0>
	<TR>
		<TD <%=HEIGHT_TYPE_00%>></TD>
	</TR>
	<TR HEIGHT=23>
		<TD>
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>품목그룹별 손익추이표</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* HEIGHT="right">&nbsp;</TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=*>
		<TD CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD HEIGHT=20>
						<TABLE <%=LR_SPACE_TYPE_60%>>
						    <TR>
								<TD CLASS=TD5  NOWRAP>대상년월</TD>
								<TD CLASS=TD6  NOWRAP><script language =javascript src='./js/gh008oa1_fpDateTime1_txtYyyymm.js'></script></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>품목그룹</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtFr_dept_cd" ALT="품목그룹코드" TYPE="Text" SiZE=10 MAXLENGTH=18 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSItemDC" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp()">
			                                         <INPUT NAME="txtFr_dept_nm" ALT="품목그룹명" TYPE="Text" SiZE=20 MAXLENGTH=40 tag="14XXXU"></TD>
							</TR>
    					</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	
	<TR>
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</TR>
	<TR HEIGHT=20>
		<TD>
		    <TABLE <%=LR_SPACE_TYPE_30%>>
		        <TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD>
		                <BUTTON NAME="btnPreview" CLASS="CLSSBTN" onclick="VBScript:FncBtnPreview()" Flag=1>미리보기</BUTTON>&nbsp;
		                <BUTTON NAME="btnPrint"   CLASS="CLSSBTN" OnClick="VBScript:FncBtnPrint()" Flag=1>인쇄</BUTTON>

		            </TD>
					<TD WIDTH=* ALIGN="right">&nbsp;</TD>
					<TD WIDTH=10>&nbsp;</TD>
		        </TR>
		    </TABLE>
		</TD>
	</TR>
	<TR>
		<TD HEIGHT=20><IFRAME type=hidden NAME="MyBizASP" SRC = "../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=no noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
<FORM NAME=EBAction TARGET="MyBizASP" METHOD="POST">
	<INPUT TYPE="HIDDEN" NAME="uname">
	<INPUT TYPE="HIDDEN" NAME="dbname">
	<INPUT TYPE="HIDDEN" NAME="filename">
	<INPUT TYPE="HIDDEN" NAME="condvar">
	<INPUT TYPE="HIDDEN" NAME="date">	
</FORM>
</BODY>
</HTML>

