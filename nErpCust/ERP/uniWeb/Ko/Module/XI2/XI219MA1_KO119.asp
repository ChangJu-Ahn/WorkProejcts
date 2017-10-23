<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : Interface 관리
*  2. Function Name        : MES전송관리
*  3. Program ID           : XI219MA1_KO119
*  4. Program Name         : 자재LOT정보등록
*  5. Program Desc         : 
*  6. Comproxy List        :
*  7. Modified date(First) : 2006/05/17
*  8. Modified date(Last)  : 2006/05/17
*  9. Modifier (First)     :
* 10. Modifier (Last)      :
* 11. Comment              :
=======================================================================================================-->
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
<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->
<!--
========================================================================================================
=                          3.2 Style Sheet
======================================================================================================== -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		

<!--
========================================================================================================
=                          3.3 Client Side Script
======================================================================================================== -->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMaMain.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMaOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="XI219MA1_KO119.vbs"> </SCRIPT>
<Script Language="VBScript">
Option Explicit                                                        '☜: Turn on the Option Explicit option.
'########################################################################################################
'#                       4.  Data Declaration Part
'########################################################################################################
'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim IsOpenPop          
Dim iDBSYSDate
Dim EndDate, StartDate

iDBSYSDate = "<%=GetSvrDate%>"
'------ ☆: 초기화면에 뿌려지는 마지막 날짜 ------
EndDate = UNIConvDateAtoB(iDBSYSDate, Parent.gServerDateFormat, Parent.gDateFormat)
'------ ☆: 초기화면에 뿌려지는 시작 날짜 ------
StartDate = UnIDateAdd("d", -1, EndDate, Parent.gDateFormat)
            'Convert DB date type to Company

'========================================================================================================
' Name : LoadInfTB19029()	
' Desc : Set System Number format
' Para : 1. Currency
'        2. I(Input),Q(Query),P(Print),B(Bacth)
'        3. "*" is for Common module
'           "A" is for Accounting
'           "I" is for Inventory
'           ...
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	<% Call loadInfTB19029A("I", "*","NOCOOKIE","MA") %>
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub

'========================================================================================================
' Name : Form_Load
' Desc : This sub is called from window_OnLoad in Common.vbs automatically
'========================================================================================================
Sub Form_Load()

	On Error Resume Next
    Err.Clear                                                                        '☜: Clear err status

	Call LoadInfTB19029()                                                              '☜: Load table , B_numeric_format
		
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
      
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
	Call ggoOper.LockField(Document, "N")											 '⊙: Lock Field
    
    Call InitSpreadSheet()                                                             'Setup the Spread sheet

	Call InitVariables()
    Call SetDefaultVal()

	frm1.txtPrintFrDt.focus
	Call SetToolBar("11101101001111")                                              '☆: Developer must customize

	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Call InitSpreadComboBox()
	Call CookiePage(0)       
	                                                       
End Sub

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	

</HEAD>

<!--
'########################################################################################################
'#						6. TAG 부																		#
'######################################################################################################## 
-->
<BODY TABINDEX=-1 SCROLL=no>
<FORM NAME=frm1 TARGET=MyBizASP METHOD=POST>
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR >
		<TD <%=HEIGHT_TYPE_00%>></TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH="100%">
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS=CLSMTABP>
						<TABLE ID=MyTab CELLSPACING=0 CELLPADDING=0>
							<TR>
								<TD BACKGROUND="../../../CShared/image/table/seltab_up_bg.gif"><IMG SRC="../../../CShared/image/table/seltab_up_left.gif" WIDTH=9 HEIGHT=23></TD>
								<TD BACKGROUND="../../../CShared/image/table/seltab_up_bg.gif" ALIGN=center CLASS=CLSMTAB><FONT COLOR=white>자재LOT정보등록</FONT></TD>
								<TD BACKGROUND="../../../CShared/image/table/seltab_up_bg.gif" ALIGN=right><IMG SRC="../../../CShared/image/table/seltab_up_right.gif" WIDTH=10 HEIGHT=23></TD>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* ALIGN=right>&nbsp;</A></TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=*>
		<TD WIDTH="100%" CLASS=Tab11>
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD <%=HEIGHT_TYPE_02%> WIDTH="100%"></TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH="100%">
						<FIELDSET CLASS=CLSFLD>
							<TABLE <%=LR_SPACE_TYPE_40%>>
									<TD CLASS=TD5 NOWRAP>발행기간</TD>
                                    <TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=javaScript SRC="./js/XI219MA1_KO119_txtPrintFrDt_txtPrintFrDt.js"></SCRIPT>
                                    &nbsp;~&nbsp;<SCRIPT LANGUAGE=javaScript SRC="./js/XI219MA1_KO119_txtPrintToDt_txtPrintToDt.js"></SCRIPT></TD>
									<TD CLASS=TD5 NOWRAP>품목</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=text NAME=txtItemCd SIZE=20 MAXLENGTH=18 tag="11XXXU" ALT="품목"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME=btnItemCd ALIGN=top TYPE=button ONCLICK="vbScript:Call OpenPopUp(frm1.txtItemCD.value, 0)">&nbsp;<INPUT TYPE=text NAME=txtItemNm SIZE=20 tag="14" ALT="품목명"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>납품처</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=text NAME=txtBpCd SIZE=15 MAXLENGTH=10 tag="11XXXU" ALT="납품처"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME=btnBpCd ALIGN=top TYPE=button ONCLICK="vbScript:Call OpenPopUp(frm1.txtBpCd.value, 1)">&nbsp;<INPUT TYPE=text NAME=txtBpNm SIZE=25 tag="14" ALT="납품처명"></TD>
                                    <TD CLASS=TD5 NOWRAP>LOT번호</TD>
									<TD CLASS=TD6 NOWRAP><input TYPE=text NAME=txtLotNo SIZE=25 MAXLENGTH=25 tag="11XXXU" ALT="LOT NO"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>삭제여부</TD>
									<TD CLASS=TD6><INPUT TYPE=radio CLASS=Radio NAME=rdoDelFlag ID=rdoDelFlagAll tag = "12"><LABEL FOR=rdoDelFlagAll>전체</LABEL>
									&nbsp;<INPUT TYPE=radio CLASS=Radio NAME=rdoDelFlag ID=rdoDelFlagNomal tag = "12" CHECKED><LABEL FOR=rdoDelFlagNomal>정상</LABEL>
									&nbsp;<INPUT TYPE=radio CLASS=Radio NAME=rdoDelFlag ID=rdoDelFlagDel tag = "12"><LABEL FOR=rdoDelFlagDel>삭제</LABEL></TD>
									<TD CLASS=TD5 NOWRAP>MES수신여부</TD>
									<TD CLASS=TD6><INPUT TYPE=radio CLASS=Radio NAME=rdoMesRcvFlag ID=rdoMesRcvFlagAll tag = "12" CHECKED><LABEL FOR=rdoMesRcvFlagAll>전체</LABEL>
									&nbsp;<INPUT TYPE=radio CLASS=Radio NAME=rdoMesRcvFlag ID=rdoMesRcvFlagNomal tag = "12"><LABEL FOR=rdoMesRcvFlagNomal>성공</LABEL>
									&nbsp;<INPUT TYPE=radio CLASS=Radio NAME=rdoMesRcvFlag ID=rdoMesRcvFlagFail tag = "12"><LABEL FOR=rdoMesRcvFlagFail>실패</LABEL></TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH="100%"></TD>
				</TR>
				<TR>
					<TD WIDTH="100%" HEIGHT="100%" VALIGN=top>
							<TABLE <%=LR_SPACE_TYPE_30%>>
							<TR>
								<TD HEIGHT="100%">
									<SCRIPT LANGUAGE=javaScript SRC="./js/XI219MA1_KO119_vspdData_vspdData.js"></SCRIPT>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR >
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</TR>
	<TR HEIGHT=20>
	    <TD>
	        <TABLE <%=LR_SPACE_TYPE_30%>>
	            <TR>
	                <TD WIDTH=10>&nbsp;</TD>
					<TD CLASS=TD5 NOWRAP>파일경로</TD>
					<TD WIDTH=210><INPUT TYPE=text ID=txtFileName NAME=txtFileName SIZE=30 MAXLENGTH=100 STYLE="TEXT-ALIGN: left" ALT="화일명" tag="14X" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnOpenFilePath" ALIGN=top TYPE=button ONCLICK="vbScript:Call GetOpenFilePath()"></TD>
					<TD WIDTH=10 ALIGN=left><BUTTON NAME=btnExe CLASS=CLSSBTN ONCLICK="ExeReflect()" Flag=1>Import</BUTTON></TD>
	                <TD WIDTH=*><SCRIPT LANGUAGE=javaScript SRC="./js/XI219MA1_KO119_cFLkUpExcel.js"></SCRIPT></TD>
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
<P ID=divTextArea></P>
<TEXTAREA CLASS=hidden NAME=txtSpread   tag=24 TABINDEX=-1></TEXTAREA>
<INPUT TYPE=hidden NAME=txtMaxRows      tag=24 TABINDEX=-1>
<INPUT TYPE=hidden NAME=txtMode         tag=24 TABINDEX=-1>
<INPUT TYPE=hidden NAME=txtKeyStream    tag=24 TABINDEX=-1>

<INPUT TYPE=hidden NAME=hFilePath       tag=14 TABINDEX=-1>
<INPUT TYPE=hidden NAME=hPrintFrDt      tag=14 TABINDEX=-1>
<INPUT TYPE=hidden NAME=hPrintToDt      tag=14 TABINDEX=-1>
<INPUT TYPE=hidden NAME=hItemCd         tag=14 TABINDEX=-1>
<INPUT TYPE=hidden NAME=hBpCd           tag=14 TABINDEX=-1>
<INPUT TYPE=hidden NAME=hLotNo          tag=14 TABINDEX=-1>
<INPUT TYPE=hidden NAME=hDelFlag        tag=14 TABINDEX=-1>
<INPUT TYPE=hidden NAME=hMesRcvFlag     tag=14 TABINDEX=-1>

      
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

