<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name       : 인사/급여관리 
'*  2. Function Name     : 급/상여공제관리 
'*  3. Program ID        : h5402oa1
'*  4. Program Name      : 월국민연금출력 
'*  5. Program Desc      : 월국민연금출력 
'*  6. Comproxy 리스트   : 
'*  7. 최초 작성년월일   : 2001/06/01
'*  8. 최종 수정년월일   : 2003/06/11
'*  9. 최초 작성자       : TGS 최용철 
'* 10. 최종 작성자       : Lee SiNa
'* 11. 전체 comment      :
'**********************************************************************************************-->

<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/IncServer.asp" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incEB.vbs"></SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>
Option Explicit
'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim IsOpenPop
Dim lgOldRow

'========================================================================================================
' Name : InitVariables()	
' Desc : Initialize value
'========================================================================================================
Sub InitVariables()

    lgIntFlgMode = Parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
End Sub

'========================================================================================================
' Name : SetDefaultVal()	
' Desc : Set default value
'========================================================================================================
Sub SetDefaultVal()
	Dim strYear
	Dim strMonth
	Dim strDay

	Call ExtractDateFrom("<%=GetsvrDate%>",Parent.gServerDateFormat , Parent.gServerDateType ,strYear,strMonth,strDay)
	
	frm1.txtpay_yymm.Focus
		
	frm1.txtpay_yymm.Year = strYear 		'년월 default value setting
	frm1.txtpay_yymm.Month = strMonth 
End Sub

'========================================================================================================
' Name : LoadInfTB19029()	
' Desc : Set System Number format
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "H", "NOCOOKIE", "OA") %>
End Sub
'========================================================================================================
' Name : Form_Load
' Desc : developer describe this line Called by Window_OnLoad() evnt
'========================================================================================================
Sub Form_Load()

    Call LoadInfTB19029                                                     '⊙: Load table , B_numeric_format
    Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
    Call InitVariables 
    
    Call ggoOper.FormatDate(frm1.txtpay_yymm, Parent.gDateFormat, 2)                    '싱글에서 년월만 입력하고 싶은경우 다음 함수를 콜한다.
    
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
	
	Dim strUrl
	Dim lngPos
	Dim intCnt
    Dim StrEbrFile, ObjName

	dim pay_yymm, grade, insur_type, insur_type1 

    If Not chkField(Document, "1") Then									<%'⊙: This function check indispensable field%>
		Exit Function
    End If
	
	with frm1
	    If .txtPrnt_type(0).Checked Then
	        StrEbrFile = "h5402oa1_1"
	    Else
	        StrEbrFile = "h5402oa1_2"
	    End If
	End with
	
	pay_yymm = frm1.txtPay_yymm.year & frm1.txtPay_yymm.month
	
	grade = frm1.txtGrade.value
	
	if grade = "" then
		grade = "%"
	End if		
	
	strUrl = "pay_yymm|" & pay_yymm 
	strUrl = strUrl & "|grade|" & grade

    ObjName = AskEBDocumentName(StrEbrFile, "ebr")
	
    call FncEBRPrint(EBAction , ObjName , strUrl)
	
End Function


'========================================================================================
' Function Name : FncBtnPreview()
' Function Desc : This function is related to Preview Button
'========================================================================================
Function FncBtnPreview()
	
	dim strUrl
	dim arrParam, arrField, arrHeader
    Dim StrEbrFile, ObjName
		
	dim pay_yymm, grade, insur_type, insur_type1 
	    
    If Not chkField(Document, "1") Then									<%'⊙: This function check indispensable field%>
       Exit Function
    End If

	with frm1
	    If .txtPrnt_type(0).Checked Then
	        StrEbrFile = "h5402oa1_1"
	    Else
	        StrEbrFile = "h5402oa1_2"
	    End If
	End with
	
    pay_yymm = frm1.txtPay_yymm.year & right("0" & frm1.txtPay_yymm.Month,2)
	grade = frm1.txtGrade.value
		
	if grade = "" then
		grade = "%"
	End if	
	
	strUrl = "pay_yymm|" & pay_yymm 
	strUrl = strUrl & "|grade|" & grade

    ObjName = AskEBDocumentName(StrEbrFile, "ebr")
	
	call FncEBRPreview(ObjName , strUrl)
	
End Function

'========================================================================================================
' Name : FncPrint
' Desc : developer describe this line Called by MainDeleteRow in Common.vbs
'========================================================================================================
Function FncPrint()
	Call Parent.FncPrint()                                                      '☜: Protect system from crashing
End Function

'========================================================================================================
' Name : FncFind
' Desc : developer describe this line Called by MainFind in Common.vbs
'========================================================================================================
Function FncFind() 
	Call Parent.FncFind(Parent.C_SINGLE, False)
End Function

'========================================================================================================
' Name : FncExit
' Desc : developer describe this line Called by MainExit in Common.vbs
'========================================================================================================
Function FncExit()
	FncExit = True
End Function
'========================================================================================================
' Name : txtPay_yymm_DblClick
' Desc : 달력 Popup을 호출 
'========================================================================================================
Sub txtPay_yymm_DblClick(Button)
	If Button = 1 Then
		Call SetFocusToDocument("M")	
		frm1.txtPay_yymm.Action = 7
		frm1.txtPay_yymm.focus
	End If
End Sub
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>월국민연금출력</font></td>
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
								<TD CLASS=TD5  NOWRAP>해당년월</TD>
								<TD CLASS=TD6  NOWRAP><script language =javascript src='./js/h5402oa1_txtPay_yymm_txtPay_yymm.js'></script></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>출력등급</TD>
								<TD CLASS=TD6 NOWRAP><INPUT ID=txtGrade NAME="txtGrade"  MAXLENGTH="2" SIZE=10 ALT ="출력등급" tag="11XXXU"></TD>	
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>출력선택</TD>
				        	    <TD CLASS="TD6"><INPUT TYPE="RADIO" CLASS="RADIO" NAME="txtPrnt_type" VALUE = "1" ID="med_entr_flag1" TAG="11XXXU" VALUE="현황출력" CHECKED><LABEL FOR="txtPrnt_type">현황출력</LABEL></TD>
				        	</TR>
				        	<TR>	
				        		<TD CLASS="TD5" NOWRAP></TD>
				        	    <TD CLASS="TD6"><INPUT TYPE="RADIO" CLASS="RADIO" NAME="txtPrnt_type" VALUE = "2" ID="med_entr_flag2" TAG="11XXXU" VALUE="집계표출력"><LABEL FOR="txtPrnt_type">집계표출력</LABEL></TD>
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
		                <BUTTON NAME="btnPreview" CLASS="CLSSBTN" onclick="VBScript:FncBtnPreview()">미리보기</BUTTON>&nbsp;
		                <BUTTON NAME="btnPrint"   CLASS="CLSSBTN" OnClick="VBScript:FncBtnPrint()">인쇄</BUTTON>

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
<FORM NAME="EBAction" TARGET = "MyBizASP" METHOD="POST">
	<INPUT TYPE="HIDDEN" NAME="uname">
	<INPUT TYPE="HIDDEN" NAME="dbname">
	<INPUT TYPE="HIDDEN" NAME="filename">
	<INPUT TYPE="HIDDEN" NAME="condvar">
	<INPUT TYPE="HIDDEN" NAME="date">
</FORM>
</BODY>
</HTML>

