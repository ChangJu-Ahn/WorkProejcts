<%@ LANGUAGE="VBSCRIPT" %>
<!--'**********************************************************************************************
'*  1. Module Name            : Inventory List onhand stock
'*  2. Function Name          : 
'*  3. Program ID             : I2221ma1.asp
'*  4. Program Name           : 
'*  5. Program Desc           : ǰ�������Ȳ��ȸ 
'*  6. Comproxy List          :      
'*  7. Modified date(First)   : 2002/09/03
'*  8. Modified date(Last)    : 2005/02/17
'*  9. Modifier (First)       : Lee Seung Wook
'* 10. Modifier (Last)        : 
'* 11. Comment                :
'**********************************************************************************************-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!--########################################################################################################
						1. �� �� ��																		#
########################################################################################################-->
<!--********************************************  1.1 Inc ����  ***************************************-->
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->
<!--'============================================  1.1.1 Style Sheet  ===================================-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">
<!--'============================================  1.1.2 ���� Include  ==================================-->
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAMain.vbs"> </SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAEvent.vbs"> </SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE=  "VBScript" SRC = "../../inc/incCliRdsQuery.vbs"></SCRIPT>

<SCRIPT LANGUAGE="VBScript"   SRC="i2221ma1.vbs"></SCRIPT>

<SCRIPT LANGUAGE = "VBScript">
Option Explicit

'==========================================  1.2.2 Global ���� ����  ==================================
'	1. ���� ǥ�ؿ� ����. prefix�� g�� �����.
'	2.Array�� ���� ()�� �ݵ�� ����Ͽ� �Ϲ� ������ ������ �� 
'========================================================================================================= 
<!-- #Include file="../../inc/lgvariables.inc" -->

Dim lgStrPrevKeyIndex2
Dim IsOpenPop 

Dim gblnWinEvent
Dim strReturn
Dim lgOldRow


'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== 
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call LoadInfTB19029A("Q", "I", "NOCOOKIE", "MA") %>
End Sub
 
'=========================================  3.1.1 Form_Load()  ==========================================
'=	Name : Form_Load()																					=
'=	Description : Window Load�� �����ʱ�ȭ �� ȭ���ʱ�ȭ�� �ϱ� ���� �Լ��� Call�ϴ� �κ�				=
'========================================================================================================
Sub Form_Load()
	Call LoadInfTB19029
	Call InitSpreadSheet("")    
    Call InitVariables
    Call SetDefaultVal
    
    Call SetToolbar("11000000000011")

End Sub



</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>
<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR>
		<TD <%=HEIGHT_TYPE_00%> >
		</TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH=100%>
			<TABLE  <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<TD background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></TD>
								<TD background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>ǰ�� �����ȸ</font></TD>
								<TD background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></TD>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* align=right><A href="vbscript:OpenOnhandDtlRef()">������ȸ</A></TD>					
					<TD WIDTH=10></TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=*>
		<TD WIDTH=100% CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD <%=HEIGHT_TYPE_02%> >
					</TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH=100%>
					<FIELDSET CLASS="CLSFLD">
					<TABLE WIDTH=100% <%=LR_SPACE_TYPE_40%> >
						<TR>
							<TD CLASS="TD5">ǰ��</TD>
							<TD CLASS="TD6">
								<INPUT NAME="txtItemCd" CLASS=required STYLE="Text-Transform: uppercase" TYPE="TEXT" MAXLENGTH=18 tag="12XXXU" ALT="ǰ��" SIZE=15 ><img SRC="../../../CShared/image/btnPopup.gif" NAME="ImgSaleOrgCode"  align="top" TYPE="BUTTON" WIDTH="16" HEIGHT="20" ONCLICK="vbscript:OpenItemCd()">&nbsp;<INPUT NAME="txtItemNm" CLASS=protected readonly=true TABINDEX="-1" TYPE="TEXT" MAXLENGTH="40" SIZE=35 tag="14N"></TD>
							<TD CLASS="TD5">����</TD>
							<TD CLASS="TD6">
								<INPUT NAME="txtBasicUnit" CLASS=protected readonly=true TABINDEX="-1" TYPE="TEXT" MAXLENGTH=3 tag="14N" ALT="����" SIZE=10 ></TD>
						</TR>
						<TR>
							<TD CLASS="TD5">�԰�</TD>
							<TD CLASS="TD6" >
								<INPUT NAME="txtSpec" CLASS=protected readonly=true TABINDEX="-1" TYPE="TEXT" MAXLENGTH=40 tag="14N" ALT="�԰�" SIZE=54 ></TD>
							<TD CLASS="TD5"></TD>
							<TD CLASS="TD6"></TD>							
						</TR>
						<TR>
							<TD <%=HEIGHT_TYPE_03%> >
							</TD>
						</TR>
					</TABLE>
					</TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=* valign=top>
						<TABLE WIDTH="100%" <%=LR_SPACE_TYPE_20%>>
							<TR HEIGHT="40%">
								<TD WIDTH="100%" colspan=4>
									<script language =javascript src='./js/i2221ma1_vaSpread1_vspdData1.js'></script>	
								</TD>
							</TR>
							<TR HEIGHT="60%">
								<TD WIDTH="100%" colspan=4>
									<script language =javascript src='./js/i2221ma1_vaSpread2_vspdData2.js'></script>
								</TD>	
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_01%> >
		</TD>	
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>>
		<IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0>
		</IFRAME>
		</TD>	
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA>	
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
<INPUT TYPE=HIDDEN NAME="hPlantCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hItemCd" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>	
	
	
	
	
	
