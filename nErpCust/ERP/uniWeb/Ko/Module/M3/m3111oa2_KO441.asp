<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Procurement
'*  2. Function Name        : 
'*  3. Program ID           : m3111oa2_KO441
'*  4. Program Name         : ���ּ����� 
'*  5. Program Desc         : ���ּ����� 
'*  6. Component List       : 
'*  7. Modified date(First) : 2000/06/01
'*  8. Modified date(Last)  : 2003/05/21
'*  9. Modifier (First)     : Shin jin hyun
'* 10. Modifier (Last)      : KANG SU HWAN
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change" 
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'**********************************************************************************************
-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- '#########################################################################################################
'												1. �� �� �� 
'##########################################################################################################!-->
<!-- '******************************************  1.1 Inc ����   **********************************************
'	���: Inc. Include
'********************************************************************************************************* !-->
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->
<!--'==========================================  1.1.1 Style Sheet  ======================================
'==========================================================================================================!-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'==========================================  1.1.2 ���� Include   ======================================
'==========================================================================================================!-->

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incEB.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs"></SCRIPT> 
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMA_KO441.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="m3111oa2_KO441.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Dim StartDate
Dim EndDate

EndDate = "<%=GetSvrDate%>"
StartDate = UNIDateAdd("m", -1, EndDate, Parent.gServerDateFormat)
EndDate   = UniConvDateAToB(EndDate, Parent.gServerDateFormat, Parent.gDateFormat)
StartDate = UniConvDateAToB(StartDate, Parent.gServerDateFormat, Parent.gDateFormat)  

'==========================================  1.2.2 Global ���� ����  =====================================
'	1. ���� ǥ�ؿ� ����. prefix�� g�� �����.
'	2.Array�� ���� ()�� �ݵ�� ����Ͽ� �Ϲ� ������ ������ �� 
'=========================================================================================================
'==========================================  1.2.2 Global ���� ����  =====================================
'	1. ���� ǥ�ؿ� ����. prefix�� g�� �����.
'	2.Array�� ���� ()�� �ݵ�� ����Ͽ� �Ϲ� ������ ������ �� 
'=========================================================================================================

Dim lgBlnFlgChgValue         
Dim lgIntFlgMode             
Dim lgIntGrpCount            

<% '==========================================  1.2.3 Global Variable�� ����  ===============================
'========================================================================================================= %>
       
Dim lblnWinEvent
Dim IsOpenPop, lgIsOpenPop  

'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'========================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "M","NOCOOKIE","OA") %>
	<% Call LoadBNumericFormatA("Q", "M", "NOCOOKIE", "OA") %>
End Sub


</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	

<SCRIPT LANGUAGE="JavaScript">
<!-- Hide script from old browsers
function setCookie(name, value, expire)
{
    //alert(value)
    //alert(escape(value))
	document.cookie = name + "=" + escape(value)
		+ ((expire == null) ? "" : ("; expires=" + expire.toGMTString()))
	document.cookie = name + "=" + escape(value)
		+ "; path=/EasyBaseWeb/bin"
		+ ((expire == null) ? "" : ("; expires=" + expire.toGMTString()))
	document.cookie = name + "=" + escape(value)
		+ "; path=/EasyBaseWeb/lib"
		+ ((expire == null) ? "" : ("; expires=" + expire.toGMTString()))
}

setCookie("client", "-1", null)
setCookie("owner", "admin", null)
setCookie("identity", "admin", null)
-->
</SCRIPT>

</HEAD>
<BODY TABINDEX="-1" SCROLL="NO">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST"> 
<TABLE CLASS="BatchTB1" CELLSPACING=0 CELLPADDING=0>
	<TR>
		<TD <%=HEIGHT_TYPE_00%>></TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD>
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>���ּ�����(S)</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=*>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
	    		<TR>
					<TD WIDTH=100%>
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD CLASS="TD5" NOWRAP>�ܰ�ǥ��</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=radio Class="Radio" ALT="�ܰ�ǥ�ÿ���" NAME="rdoPoflg" id = "rdoPoflg1" Value="Y"  tag="12"><label for="rdoPoflg1">&nbsp;ǥ��&nbsp;</label>
													   <INPUT TYPE=radio Class="Radio" ALT="�ܰ�ǥ�ÿ���" NAME="rdoPoflg" id = "rdoPoflg2" Value="N" checked tag="12"><label for="rdoPoflg2">&nbsp;��ǥ��&nbsp;</label>
								</TD>									
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>����ó</TD>
								<TD CLASS="TD6" NOWRAP><INPUT STYLE = "text-transform:uppercase" TYPE=TEXT NAME="txtBpCd" SIZE=10 MAXLENGTH=10 ALT="����ó" tag="11NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBpCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenBpCd()">
													   <INPUT TabIndex = -1 CLASS = Protected TYPE=TEXT NAME="txtBpNm" SIZE=20 MAXLENGTH=18 ALT="����ó" tag="14X"></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>���ű׷�</TD>
								<TD CLASS="TD6" NOWRAP><INPUT STYLE = "text-transform:uppercase" TYPE=TEXT NAME="txtPurGrpCd" SIZE=10 MAXLENGTH=10 ALT="���ű׷�" tag="11NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPurGrpCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPurGrpCd()">
													   <INPUT TabIndex = -1 CLASS = Protected TYPE=TEXT NAME="txtPurGrpNm" SIZE=20 MAXLENGTH=18 ALT="���ű׷�" tag="14X"></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>������</TD>
								<TD CLASS="TD6" NOWRAP>
									<table cellspacing=0 cellpadding=0>
										<tr>
											<td>
												<OBJECT ALT=������ NAME="txtFrDt" classid=<%=gCLSIDFPDT%> id=fpDateTime1 style="HEIGHT: 20px; WIDTH: 100px" tag="11N1" Title="FPDATETIME"></OBJECT>
											</td>
											<td>~</td>
											<td>
												<OBJECT ALT=������ NAME="txtToDt" classid=<%=gCLSIDFPDT%> id=fpDateTime2 style="HEIGHT: 20px; WIDTH: 100px" tag="11N1" Title="FPDATETIME"></OBJECT>
											</td>
										<tr>
									</table>
								</TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>���ֹ�ȣ</TD>
								<TD CLASS="TD6" NOWRAP><INPUT STYLE = "text-transform:uppercase" TYPE=TEXT NAME="txtPoNo"  SIZE=32 MAXLENGTH=18 ALT="���ֹ�ȣ" tag="11XXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnORGCd1" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPoNo()">
													   <INPUT type="hidden" NAME="txtSoNo1" STYLE="BORDER-RIGHT: 0px solid;BORDER-TOP: 0px solid;BORDER-LEFT: 0px solid;BORDER-BOTTOM: 0px solid" TYPE="Text" SIZE=1 DISABLED=TRUE Tag="11">����</TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP></TD>
								<TD CLASS="TD6" NOWRAP><INPUT STYLE = "text-transform:uppercase" TYPE=TEXT NAME="txtPoNo1"  SIZE=32 MAXLENGTH=18 ALT="���ֹ�ȣ" tag="11XXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnORGCd2" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPoNo1()">
													   <INPUT type="hidden" NAME="txtSoNo2" STYLE="BORDER-RIGHT: 0px solid;BORDER-TOP: 0px solid;BORDER-LEFT: 0px solid;BORDER-BOTTOM: 0px solid" TYPE="Text" SIZE=1 DISABLED=TRUE Tag="11">����</TD>
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
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_30%>>
			    <TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD><BUTTON NAME="btnRun" CLASS="CLSSBTN" ONCLICK="vbscript:btnPreview()" Flag=1>�̸�����</BUTTON>&nbsp;<BUTTON NAME="btnPrint" CLASS="CLSSBTN" ONCLICK="vbscript:FncBtnPrint()" Flag=1>�μ�</BUTTON></TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 tabindex=-1></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="hdnImportflg" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
<FORM NAME=EBAction TARGET="MyBizASP" METHOD="POST" >
    <input type="hidden" name="uname">
    <input type="hidden" name="dbname">
    <input type="hidden" name="filename">
    <input type="hidden" name="condvar">
	<input type="hidden" name="date">
</FORM>
</BODY>
</HTML>