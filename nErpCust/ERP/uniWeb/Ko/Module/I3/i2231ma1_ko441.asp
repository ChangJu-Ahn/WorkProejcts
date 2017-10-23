<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Inventory
'*  2. Function Name        : 
'*  3. Program ID           : I2231ma1.asp
'*  4. Program Name         : 재고마감 및 취소작업 
'*  5. Program Desc         : 
'*  6. Comproxy List        : 
'                             +B25019LookUpPlant
'*  7. Modified date(First) : 1999/09/10
'*  8. Modified date(Last)  : 2000/11/25
'*  9. Modifier (First)        :  Mr  Kim Nam Hoon
'* 10. Modifier (Last)      : HAN
'* 11. Comment              :
'**********************************************************************************************
-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- '#########################################################################################################
'            1. 선 언 부 
'##########################################################################################################-->
<!-- '******************************************  1.1 Inc 선언   **********************************************
' 기능: Inc. Include
'********************************************************************************************************* -->
<!-- #Include file="../../inc/incSvrCcm.inc" -->      

<!--'==========================================  1.1.1 Style Sheet  ======================================
'==========================================================================================================-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css"> 

<!--'==========================================  1.1.2 공통 Include   ======================================
'==========================================================================================================-->

<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>

<SCRIPT LANGUAGE="VBScript"   SRC="i2231ma1_ko441.vbs"></SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit                                                       


'==========================================  1.2.2 Global 변수 선언  =====================================
' 1. 변수 표준에 따름. prefix로 g를 사용함.
' 2.Array인 경우는 ()를 반드시 사용하여 일반 변수와 구별해 됨 
'========================================================================================================= 
<!-- #Include file="../../inc/lgvariables.inc" -->
Dim IsOpenPop          

'=============================== 2.1.2 LoadInfTB19029() ========================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== 
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "I","NOCOOKIE","MA") %>
End Sub


'==========================================  3.1.1 Form_Load()  ======================================
' Name : Form_Load()
' Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분 
'========================================================================================================= 
Sub Form_Load()
	Call InitVariables         
	Call LoadInfTB19029
	
	Call LockObjectField(frm1.txtInvClsDt,"R")
    Call FormatDATEField(frm1.txtInvClsDt)
                 
	Call ggoOper.FormatDate(frm1.txtInvClsDt,Parent.gDateFormat,"2")

	Call SetToolbar("10000000000011")
	Call SetDefaultVal

	frm1.txtPlantCd.focus
End Sub


</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" --> 
</HEAD>
<BODY TABINDEX="-1" SCROLL="NO">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST"> 
<TABLE CLASS="BatchTB1" CELLSPACING=0 CELLPADDING=0>
	<TR>
		<TD <%=HEIGHT_TYPE_00%></TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_10%> WIDTH=100% border=0>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD>
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>재고마감작업</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
							</TR>
						</TABLE>
					</TD>
					<TD WIDTH=* align=right><A href="vbscript:OpenCancelListRef()">취소대상수불정보</A></TD>
					<TD WIDTH=*>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=*>
		<TD WIDTH=100% CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD>
						<TABLE <%=LR_SPACE_TYPE_60%> >    
							<TR>
								<TD CLASS="TD5" NOWRAP>공장</TD>
								<TD CLASS="TD6"><INPUT TYPE=TEXT NAME="txtPlantCd" CLASS=required STYLE="Text-Transform: uppercase" SIZE=10 MAXLENGTH=4 tag="22XXXU" ALT="공장" onBlur="vbscript:txtPlantCd_LostFocus()"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlant" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenPlant()">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" CLASS=protected readonly=true TABINDEX="-1" SIZE=20 MAXLENGTH=20 tag="24"></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>최종마감년월</TD>
								<TD CLASS="TD6">
								<script language =javascript src='./js/i2231ma1_fpDateTime2_txtInvClsDt.js'></script></TD>
							</TR>

						    <!-- 2009.10.21...kbs...재고마감시 수불오류내역 체크 및 보정 부분 추가    -->
							<TR>
								<TD CLASS="TD5" NOWRAP>오류/보정건수</TD>
								<TD CLASS="TD6"><INPUT TYPE=TEXT NAME="txtCheckCnt1" CLASS=protected readonly=true TABINDEX="-1" SIZE=10 tag="24" ALT="오류건수1" >&nbsp;&nbsp;/&nbsp;&nbsp;
										<INPUT TYPE=TEXT NAME="txtUpdateCnt1" CLASS=protected readonly=true TABINDEX="-1" SIZE=10 tag="24" ALT="보정건수1" >
								</TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP></TD>
								<TD CLASS="TD6"><INPUT TYPE=TEXT NAME="txtCheckCnt2" CLASS=protected readonly=true TABINDEX="-1" SIZE=10 tag="24" ALT="오류건수2" >&nbsp;&nbsp;/&nbsp;&nbsp;
										<INPUT TYPE=TEXT NAME="txtUpdateCnt2" CLASS=protected readonly=true TABINDEX="-1" SIZE=10 tag="24" ALT="보정건수2" >
								</TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP></TD>
								<TD CLASS="TD6"><INPUT TYPE=TEXT NAME="txtCheckCnt3" CLASS=protected readonly=true TABINDEX="-1" SIZE=10 tag="24" ALT="오류건수3" >&nbsp;&nbsp;/&nbsp;&nbsp;
										<INPUT TYPE=TEXT NAME="txtUpdateCnt3" CLASS=protected readonly=true TABINDEX="-1" SIZE=10 tag="24" ALT="보정건수3" >
								</TD>
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
			<TABLE>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
				<!-- 2009.10.21...kbs...재고마감시 수불오류내역 체크 및 보정 부분 추가    -->
				   <!--	<TD><BUTTON NAME="btnRun" ONCLICK="vbscript:Fncsave1()" CLASS="CLSMBTN">재고마감 Simulation</BUTTON>&nbsp;<BUTTON NAME="btnConfirm" ONCLICK="vbscript:Fncsave2()" CLASS="CLSMBTN">재고마감 확정</BUTTON>&nbsp;<BUTTON NAME="btnCancel" ONCLICK="vbscript:Fncsave3()" CLASS="CLSMBTN">재고마감 취소</BUTTON>	-->
					<TD><BUTTON NAME="btnIoChk" ONCLICK="vbscript:Fncsave11()" CLASS="CLSMBTN">재고수불 체크</BUTTON>&nbsp;
					    <BUTTON NAME="btnIoUpd" ONCLICK="vbscript:Fncsave12()" CLASS="CLSMBTN">재고수불 보정</BUTTON>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;

					    <BUTTON NAME="btnRun" ONCLICK="vbscript:Fncsave1()" CLASS="CLSMBTN">재고마감 Simulation</BUTTON>&nbsp;
					    <BUTTON NAME="btnConfirm" ONCLICK="vbscript:Fncsave2()" CLASS="CLSMBTN">재고마감 확정</BUTTON>&nbsp;
					    <BUTTON NAME="btnCancel" ONCLICK="vbscript:Fncsave3()" CLASS="CLSMBTN">재고마감 취소</BUTTON>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=20 FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24" TABINDEX="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

