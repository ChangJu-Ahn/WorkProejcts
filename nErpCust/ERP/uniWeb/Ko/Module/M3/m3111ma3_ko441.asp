<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Procurement
'*  2. Function Name        : 
'*  3. Program ID           : m3111ma3
'*  4. Program Name         : 발주일괄확정/확정취소 
'*  5. Program Desc         : 발주일괄확정/확정취소 
'*  6. Component List       : 
'*  7. Modified date(First) : 2000/04/14
'*  8. Modified date(Last)  : 2003/05/21
'*  9. Modifier (First)     : Shin jin hyun
'* 10. Modifier (Last)      : KANG SU HWAN
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change" 
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************
-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- '******************************************  1.1 Inc 선언   **********************************************
'	기능: Inc. Include
'********************************************************************************************************* !-->
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->
<!--'==========================================  1.1.1 Style Sheet  ======================================-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'==========================================  1.1.2 공통 Include   =====================================-->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMA_KO441.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="m3111ma3_ko441.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>
Option Explicit				


<!-- #Include file="../../inc/lgvariables.inc" -->

         
Dim StartDate,EndDate
EndDate = "<%=GetSvrDate%>"
StartDate = UNIDateAdd("m", -1, EndDate, Parent.gServerDateFormat)
EndDate   = UniConvDateAToB(EndDate, Parent.gServerDateFormat, Parent.gDateFormat)
StartDate = UniConvDateAToB(StartDate, Parent.gServerDateFormat, Parent.gDateFormat)  


'========================================================================================
' Function Name : LoadInfTB19029
'========================================================================================
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029A("I", "*","NOCOOKIE","MA") %>
<% Call LoadBNumericFormatA("I", "*","NOCOOKIE","MA") %>
End Sub

'==========================================  3.1.1 Form_Load()  ======================================
Sub Form_Load()
    Call LoadInfTB19029                             
    'Call ggoOper.LockField(Document, "N")           
    'Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
    Call FormatDATEField(frm1.txtFrDt)
    Call FormatDATEField(frm1.txtToDt)
        
    Call LockObjectField(frm1.txtFrDt, "O")
    Call LockObjectField(frm1.txtToDt, "O")
    
    Call InitSpreadSheet                            
    Call GetValue_ko441()
    Call SetDefaultVal
    Call InitVariables                                  
End Sub

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>
<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR>
		<TD <%=HEIGHT_TYPE_00%>></TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSLTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSLTAB"><font color=white>발주일괄확정/확정취소</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=*>&nbsp;</TD>
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
					<FIELDSET CLASS="CLSFLD"><TABLE <%=LR_SPACE_TYPE_40%>>
					<TR>
						<TD CLASS=TD5 NOWRAP>구매그룹</TD>
						<TD CLASS=TD6 NOWRAP>
							<INPUT TYPE=TEXT ALT="구매그룹"  NAME="txtPur_Grp" SIZE=10 MAXLENGTH=4 CLASS=required STYLE="text-transform:uppercase" tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPur_Grp" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPurGrp()">
               			 	<INPUT TYPE=TEXT ALT="구매그룹"  NAME="txtPur_Grp_Nm" SIZE=20 MAXLENGH=20 CLASS=protected readonly=true tag="14X"></TD>
							<TD CLASS=TD5 NOWRAP>확정여부</TD>
							<TD CLASS=TD6 NOWRAP>
								<input type=radio CLASS="RADIO" name="rdoConfirmFlg" id="rdoConfirmFlg_Yes" value="Y" tag = "11"><label for="rdoConfirmFlg_Yes">확정</label>&nbsp;&nbsp;
								<input type=radio CLASS = "RADIO" name="rdoConfirmFlg" id="rdoConfirmFlg_No" value="N" tag = "11"><label for="rdoConfirmFlg_No">미확정</label>
							</TD>
						</TR>
						<TR>
							<TD CLASS="TD5" NOWRAP>공급처</TD>
							<TD CLASS="TD6" NOWRAP>
								<INPUT TYPE=TEXT AlT="공급처"  NAME="txtSupplierCd" SIZE=10 MAXLENGTH=10 STYLE="text-transform:uppercase" tag="11NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnGroup1Cd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenSupplier()">
               					<INPUT TYPE=TEXT AlT="공급처" ID="txtSupplierNm" NAME="arrCond" CLASS=protected readonly=true tag="14X">
               				</TD>
							<TD CLASS="TD5" NOWRAP>발주일</TD>
							<TD CLASS="TD6" NOWRAP>
								<table cellspacing=0 cellpadding=0>
									<tr>
										<td NOWRAP>
											<script language =javascript src='./js/m3111ma3_fpDateTime1_txtFrDt.js'></script>
										</td>
										<td NOWRAP>~</td>
										<td NOWRAP>
											<script language =javascript src='./js/m3111ma3_fpDateTime1_txtToDt.js'></script>
										</td>
									<tr>
								</table>
							</TD>
						</TR>
					</TABLE></FIELDSET></TD>
			</TR>
			<TR>
				<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
			</TR>
			<TR>
				<TD WIDTH=100% valign=top><TABLE <%=LR_SPACE_TYPE_20%>>
					<TR>
						<TD HEIGHT="100%">
						    <script language =javascript src='./js/m3111ma3_I475566178_vspdData.js'></script>
						</TD>
					</TR></TABLE>
				</TD>
			</TR>
		</TABLE></TD>
	</TR>
    <tr>
		<TD <%=HEIGHT_TYPE_01%>></TD>
    </tr>
    <tr HEIGHT="20">
	    <td WIDTH="100%">
	    	<table <%=LR_SPACE_TYPE_30%>>
				<tr> 
					<TD WIDTH=10>&nbsp;</TD>
					<td WIDTH="*" align="left">
					<button name="btnSelect" class="clsmbtn" >일괄선택</button>&nbsp;
					<BUTTON NAME="btnDisSelect" CLASS="CLSMBTN">일괄선택취소</BUTTON>
					</td>
					<TD WIDTH=10>&nbsp;</TD>
				</tr>
	    	</table>
	    </td>
    </tr>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 tabindex="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<P ID="divTextArea"></P>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnFrDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnToDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnSupplier" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnGroup" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnCfmflg" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnrdoflg" tag="24">
</FORM>

    <DIV ID="MousePT" NAME="MousePT">
    <iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
    </DIV>

</BODY>
</HTML>
