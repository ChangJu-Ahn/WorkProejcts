<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Procurement
'*  2. Function Name        : 
'*  3. Program ID           : M2111MA1
'*  4. Program Name         : 구매요청등록 
'*  5. Program Desc         : 구매요청등록 
'*  6. Component List       : 
'*  7. Modified date(First) : 
'*  8. Modified date(Last)  : 2003/05/21
'*  9. Modifier (First)     : MINHJ
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
' 기능: Inc. Include
'********************************************************************************************************* !-->
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->
<!--'==========================================  1.1.1 Style Sheet  ======================================
'==========================================================================================================!-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'==========================================  1.1.2 공통 Include   ======================================
'==========================================================================================================!-->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="m2111ma1.vbs"></SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>
Option Explicit    


<!--'=============================== 2.1.2 LoadInfTB19029() =================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== !-->
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "*", "NOCOOKIE", "MA") %>
End Sub

<!-- '==========================================  2.2.1 SetDefaultVal()  ========================================
' Name : SetDefaultVal()
' Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
'========================================================================================================= !-->
Sub SetDefaultVal()
	frm1.txtPlantCd.Value = Parent.gPlant
	frm1.txtPlantNm.Value = Parent.gPlantNm
	frm1.txtReqDt.text =UNIDateClientFormat("<%=GetSvrDate%>")
	frm1.txtDeptCd.Value = Parent.gDepart
	frm1.hdnTrackingflg.Value = "N"
	Call ggoOper.SetReqAttr(frm1.txtReqNo2, "D")
	Call changeTagTracking()
	Call SetToolBar("1110100000001111")
	Set gActiveElement = document.activeElement
	frm1.txtReqNo.focus 	
End Sub


<!-- '==========================================  3.1.1 Form_Load()  ======================================
' Name : Form_Load()
' Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분 
'========================================================================================================= !-->
Sub Form_Load()
	Call LoadInfTB19029 
	'Call ggoOper.FormatField(Document, "2", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec,,, ggStrMinPart, ggStrMaxPart)
	Call FormatDATEField(frm1.txtReqDt)
    Call FormatDoubleSingleField(frm1.txtReqQty)
    Call FormatDATEField(frm1.txtDlvyDt)
    Call FormatDoubleSingleField(frm1.txtPoQty)
    Call FormatDoubleSingleField(frm1.txtGmQty)
    
	Call SetDefaultVal
	Call InitVariables
	  
    Call LockObjectField(frm1.txtReqDt, "R")
    Call LockObjectField(frm1.txtReqQty, "R")
    Call LockObjectField(frm1.txtDlvyDt, "R")
    Call LockObjectField(frm1.txtPoQty, "P")
    Call LockObjectField(frm1.txtGmQty, "P")
    
	Call InitSpreadSheet
	Call ReadCookiePage()
End Sub



</SCRIPT>
<!-- #Include file="../../inc/uni2kCM.inc" --> 
</HEAD>
<BODY TABINDEX="-1" SCROLL="NO">

<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_00%>>
 
	<TR>
		<TD <%=HEIGHT_TYPE_00%>></TD>
	</TR>
	 
	<TR HEIGHT="23">
		<TD WIDTH="100%">
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10></TD>
					<TD>
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>구매요청</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
							</TR>
						</TABLE>
					</TD>
					<TD WIDTH=500>&nbsp;</TD>
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
					<TD HEIGHT=10 WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS="TD5">요청번호</TD>
									<TD CLASS="TD6"><INPUT TYPE=TEXT ALT="요청번호" NAME="txtReqNo"  SIZE=32 MAXLENGTH=18 CLASS=required STYLE="text-transform:uppercase" tag="12NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnORGCd1" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenReqNo()"></TD>
									<TD CLASS="TD6">&nbsp;</TD>
									<TD CLASS="TD6">&nbsp;</TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
					</TR>
				    
				<TR>
				<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
    
				 <TR>
				  <TD WIDTH=100% valign=top>

				    <TABLE <%=LR_SPACE_TYPE_60%>>
				     <TR>
				      <TD CLASS="TD5">요청번호</TD>
				      <TD CLASS="TD6"><INPUT TYPE=TEXT ALT="요청번호" NAME="txtReqNo2"  SIZE=34 MAXLENGTH=18 STYLE="text-transform:uppercase" tag="21XXXU"></TD>
				      <TD CLASS="TD5" NOWRAP></TD>
				      <TD CLASS="TD6" NOWRAP></TD>
				     </TR>
				     <TR>
				      <TD CLASS="TD5" NOWRAP>공장</TD>
				      <TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="공장" NAME="txtPlantCd" SIZE=10 MAXLENGTH=4 CLASS=required STYLE="text-transform:uppercase" tag="23NXXU" ONCHANGE="vbscript:changeItemPlant()"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnORG2Cd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenPlant()" OnMouseOver="vbscript:PopUpMouseOver()" OnMouseOut="vbscript:PopUpMouseOut()">
				              <INPUT TYPE=TEXT ALT="공장" NAME="txtPlantNm" SIZE=20 CLASS=protected readonly=true tag="24x" tabindex = -1></TD>
				      <TD CLASS="TD5" NOWRAP>품목</TD>
				      <TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="품목" NAME="txtItemCd"  SIZE=15 MAXLENGTH=18 CLASS=required STYLE="text-transform:uppercase" tag="23NXXU" ONCHANGE="vbscript:changeItemPlant()"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnORG2Cd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenItem()" OnMouseOver="vbscript:PopUpMouseOver()" OnMouseOut="vbscript:PopUpMouseOut()">
				              <INPUT TYPE=TEXT ALT="품목" NAME="txtItemNm" SIZE=20 CLASS=protected readonly=true tag="24x" tabindex = -1></TD>
				     </TR>
				     <TR>
				      <TD CLASS="TD5" NOWRAP>요청일</TD>
				      <TD CLASS="TD6" NOWRAP><script language =javascript src='./js/m2111ma1_fpDateTime1_txtReqDt.js'></script></TD>
				      <TD CLASS="TD5" NOWRAP>규격</TD>
				      <TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="규격" NAME="txtSpec" SIZE=30 CLASS=protected readonly=true tag="24" tabindex = -1></TD>
				     </TR>
				     <TR>
				      <TD CLASS="TD5" NOWRAP>요청량</TD>
				      <TD CLASS="TD6" NOWRAP><script language =javascript src='./js/m2111ma1_fpDoubleSingle1_txtReqQty.js'></script></td>
				      <TD CLASS="TD5" NOWRAP>필요일</TD>
				      <TD CLASS="TD6" NOWRAP><script language =javascript src='./js/m2111ma1_fpDateTime2_txtDlvyDt.js'></script></TD>
				     </TR>
				     <TR>
				      <TD CLASS="TD5" NOWRAP>요청부서</TD>
				      <TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT Alt="요청부서" NAME="txtDeptCd" SIZE=10 MAXLENGTH=10  STYLE="text-transform:uppercase" tag="2XNXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnORG2Cd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenDept()">
				              <INPUT TYPE=TEXT Alt="요청부서" NAME="txtDeptNm" SIZE=20 CLASS=protected readonly=true tag="24x" tabindex = -1></TD>
				      <TD CLASS="TD5" NOWRAP>요청단위</TD>
				      <TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="요청단위"  NAME="txtReqUnitCd" SIZE=10 MAXLENGTH=3 CLASS=required STYLE="text-transform:uppercase" tag="22XNXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnGroupCd2" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenUnit()">
				     </TR>
				     <TR>
				      <TD CLASS="TD5" NOWRAP>입고창고</TD>
				      <TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="입고창고" NAME="txtStorageCd"  SIZE=10 MAXLENGTH=7 STYLE="text-transform:uppercase" tag="2XNXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnORG2Cd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenStorage()">
				              <INPUT TYPE=TEXT ALT="입고창고" NAME="txtstorageNm" SIZE=20 CLASS=protected readonly=true tag="24X" tabindex = -1></TD>
				      <TD CLASS="TD5" NOWRAP>요청자</TD>
				      <TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT Alt="요청자"  NAME="txtEmpCd" MAXLENGTH=50 SIZE=34 tag="2XN"></TD>
				     </TR>
				     <TR>
				      <TD CLASS="TD5" NOWRAP>구매조직</TD>
				      <TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT Alt="구매조직" MAXLENGTH=4 NAME="txtOrgCd" SIZE=10 MAXLENGTH=4 CLASS=required STYLE="text-transform:uppercase" tag="22NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnORG2Cd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenOrg()">
				              <INPUT TYPE=TEXT Alt="구매조직" NAME="txtOrgNm" SIZE=20 CLASS=protected readonly=true tag="24X" tabindex = -1></TD>
				      <TD CLASS="TD5" NOWRAP>Tracking No.</TD>
				      <TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="Tracking No." NAME="txtTrackingNo" SIZE=32 MAXLENGTH=25 CLASS=protected readonly=true tag="24" tabindex = -1><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnORG2Cd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenTrackingNo()"></TD>
				     </TR>
				     <TR>
				      <TD CLASS="TD5" NOWRAP>요청진행상태</TD>
				      <TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="요청진행상태" NAME="txtReqStateCd" SIZE=10 CLASS=protected readonly=true tag="24">&nbsp;&nbsp;&nbsp;&nbsp;
				              <INPUT TYPE=TEXT ALT="요청진행상태" NAME="txtReqStateNm" SIZE=20 CLASS=protected readonly=true tag="24" tabindex = -1></TD>
				      <TD CLASS="TD5" NOWRAP>요청구분</TD>
				      <TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="요청구분" NAME="txtReqTypeCd" SIZE=10 CLASS=protected readonly=true tag="24">&nbsp;&nbsp;&nbsp;&nbsp;
				              <INPUT TYPE=TEXT ALT="요청구분" NAME="txtReqTypeNm" SIZE=20 CLASS=protected readonly=true tag="24" tabindex = -1></TD>
				     </TR>
				     <TR>
				      <TD CLASS="TD5" NOWRAP>발주량</TD>
				      <TD CLASS="TD6" NOWRAP><script language =javascript src='./js/m2111ma1_fpDoubleSingle2_txtPoQty.js'></script></td>
				      <TD CLASS="TD5" NOWRAP>입고량</TD>
				      <TD CLASS="TD6" NOWRAP><script language =javascript src='./js/m2111ma1_fpDoubleSingle3_txtGmQty.js'></script></td>
				     </TR>
				     
				     <%Call SubFillRemBodyTD5656(5)%>
						
						<%'e-Ware 결재 상태 보여주기 
						If gEWare <> "" Then
						%>
						<TR>
							<TD CLASS="TD5" NOWRAP>I/F Status</TD>
							<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="I/F Status" NAME="txtIFStatusNm" SIZE=11 tag="24"></TD>
							<TD CLASS="TD5" NOWRAP></TD>
							<TD CLASS="TD6" NOWRAP></TD>
						</TR>
						<%	
						End If
						%>
						
					  <TR>
						<TD HEIGHT="100%" WIDTH=100% COLSPAN=4>
							<script language =javascript src='./js/m2111ma1_OBJECT_vspdData2.js'></script>
						</TD>
					  </TR>
			            
					</TABLE>
				 </TD> 
				</TR>
			</Table>
		</TD>
	</TR>
   
	<tr>
	  <td <%=HEIGHT_TYPE_01%>></TD>
	</tr>
	    
	<tr HEIGHT="20">
		<td WIDTH="100%">
			<table <%=LR_SPACE_TYPE_30%>>
				<tr>	 
					<td WIDTH="*" ALIGN="RIGHT"><a href="VBSCRIPT:PgmJump(BIZ_PGM_JUMP_ID)" ONCLICK="VBSCRIPT:WriteCookiePage()">구매요청조회</a></td>
					<td WIDTH="20">&nbsp;</td>
				</tr>
			</table>
		</td>
	</tr>
	    
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 tabindex="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>

<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnProcurType" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnTrackingflg" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnMrpNo" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnOrg" tag="24">
</FORM>

    <DIV ID="MousePT" NAME="MousePT">
    <iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
    </DIV>

</BODY>
</HTML>
    