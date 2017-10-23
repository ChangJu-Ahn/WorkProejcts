<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name            : Inventory List onhand stock
'*  2. Function Name          : 
'*  3. Program ID             : I2211ma1.asp
'*  4. Program Name           : 
'*  5. Program Desc           : �����Ȳ��ȸ 
'*  6. Comproxy List          :      
'*  7. Modified date(First)   : 2000/04/01
'*  8. Modified date(Last)    : 2003/10/17
'*  9. Modifier (First)       :  Nam hoon kim
'* 10. Modifier (Last)        :	 Lee Seung Wook	
'* 11. Comment                :
'**********************************************************************************************
-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- '#########################################################################################################
'            1. �� �� �� 
'##########################################################################################################-->
<!-- '******************************************  1.1 Inc ����   **********************************************
' ���: Inc. Include
'********************************************************************************************************* -->
<!-- #Include file="../../inc/incSvrCcm.inc" -->

<!--'==========================================  1.1.1 Style Sheet  ======================================
'==========================================================================================================-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'==========================================  1.1.2 ���� Include   ======================================
'==========================================================================================================-->

<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>

<SCRIPT LANGUAGE="VBScript"   SRC="i2211ma1.vbs"></SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit        


'==========================================  1.2.2 Global ���� ����  =====================================
' 1. ���� ǥ�ؿ� ����. prefix�� g�� �����.
' 2.Array�� ���� ()�� �ݵ�� ����Ͽ� �Ϲ� ������ ������ �� 
'========================================================================================================= 
<!-- #Include file="../../inc/lgvariables.inc" -->
Dim lgStrPrevKey1
Dim lgStrPrevKey2

'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== 
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029A("Q", "I","NOCOOKIE","MA") %>
End Sub


'==========================================  3.1.1 Form_Load()  ======================================
Sub Form_Load()

    Call LoadInfTB19029             

    '----------  Coding part  -------------------------------------------------------------
 
    Call InitSpreadSheet    
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
						<TABLE <%=LR_SPACE_TYPE_10%>>
							<TR>
								<TD WIDTH=10>&nbsp;</TD>
								<TD CLASS="CLSMTABP">
									<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
										<TR>
											<TD background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></TD>
											<TD background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>�����Ȳ��ȸ</font></TD>
											<TD background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></TD>
										</TR>
									</TABLE>
								</TD>
								<TD WIDTH=* align=right><A href="vbscript:OpenOnhandDtlRef()">��������</A></TD>     
								<TD WIDTH=10>
								</TD>
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
										<TABLE WIDTH=100% <%=LR_SPACE_TYPE_40%>>
											<TR>
												<TD CLASS="TD5" NOWRAP>����</TD>      
												<TD CLASS="TD6" NOWRAP >
													<input NAME="txtPlant_Cd" TYPE="Text" CLASS=required STYLE="Text-Transform: uppercase" MAXLENGTH="4" tag="12XXXU" ALT = "����" size="8"><img SRC="../../../CShared/image/btnPopup.gif" NAME="ImgSaleOrgCode"  align="top" TYPE="BUTTON" WIDTH="16" HEIGHT="20" ONCLICK="vbscript:OpenPlantCode()">&nbsp;<input NAME="txtPlant_Nm" TYPE="Text" CLASS=protected readonly=true TABINDEX="-1" MAXLENGTH="40" SIZE=25 tag="14N"></TD>    
												<TD CLASS="TD5" NOWRAP>â��</TD>
												<TD CLASS="TD6" NOWRAP >
													<input NAME="txtSL_Cd" TYPE="Text" CLASS=required STYLE="Text-Transform: uppercase" MAXLENGTH="7" tag="12XXXU" ALT = "â��" size="8"><img SRC="../../../CShared/image/btnPopup.gif" NAME="ImgSaleOrgCode" align="top" TYPE="BUTTON" WIDTH="16" HEIGHT="20" ONCLICK="vbscript:OpenSLCode()">&nbsp;<input NAME="txtSL_Nm" TYPE="Text" CLASS=protected readonly=true TABINDEX="-1" MAXLENGTH="40" SIZE=25 tag="14N"></TD>    
											</TR>
											<TR>
												<TD CLASS="TD5" NOWRAP>ǰ��</TD>      
												<TD CLASS="TD6" NOWRAP >
													<input NAME="txtItem_Cd" TYPE="Text" STYLE="Text-Transform: uppercase" MAXLENGTH="18" tag="11NXXU" ALT = "ǰ��" size="15"><img SRC="../../../CShared/image/btnPopup.gif" NAME="ImgSaleOrgCode" align="top" TYPE="BUTTON" WIDTH="16" HEIGHT="20" ONCLICK="vbscript:OpenItemCode()">&nbsp;<input NAME="txtItem_Nm" TYPE="Text" CLASS=protected readonly=true TABINDEX="-1" MAXLENGTH="40" tag="14N"></TD>     
												<TD CLASS="TD5" NOWRAP>������</TD>      
												<TD CLASS="TD6" NOWRAP >
													<input NAME="txtinvunit" TYPE="Text" STYLE="Text-Transform: uppercase" MAXLENGTH="3" tag="11NXXU" ALT = "������" size="8"><img SRC="../../../CShared/image/btnPopup.gif" NAME="ImgSaleOrgCode" align="top" TYPE="BUTTON" WIDTH="16" HEIGHT="20" ONCLICK="vbscript:OpenEntryUnit()">&nbsp;<input NAME="txtunit_Nm" TYPE="Text" CLASS=protected readonly=true TABINDEX="-1" MAXLENGTH="40" SIZE=20 tag="14N"></TD>     
											</TR>
											<TR>
												<TD CLASS="TD5" NOWRAP>ǰ����ȿ��üũ</TD>
												<TD CLASS="TD6">
												<INPUT TYPE="RADIO" CLASS="RADIO" NAME="RadioOutputType" ID="rdoCase1" TAG="1X"><LABEL FOR="rdoCase1">��</LABEL>
												<INPUT TYPE="RADIO" CLASS="RADIO" NAME="RadioOutputType" ID="rdoCase2" TAG="1X" checked><LABEL FOR="rdoCase2">�ƴϿ�</LABEL>
												</TD>
												<TD CLASS="TD5" NOWRAP>��ǰ��������</TD>      
												<TD CLASS="TD6" NOWRAP >
													<INPUT TYPE="RADIO" CLASS="RADIO" NAME="RadioOutputType2" ID="rdoCase1" TAG="1X"><LABEL FOR="rdoCase1">��������</LABEL>
													<INPUT TYPE="RADIO" CLASS="RADIO" NAME="RadioOutputType2" ID="rdoCase2" TAG="1X" checked><LABEL FOR="rdoCase2">��ǰ��</LABEL>
												</TD>
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
											<TR>
												<TD HEIGHT="100%">
													<script language =javascript src='./js/i2211ma1_OBJECT1_vspdData.js'></script></TD>
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
					<TR HEIGHT=20 >
						<TD>
							<TABLE <%=LR_SPACE_TYPE_30%> >
							</TABLE>
						</TD>
					</TR>
					<TR>
						<TD WIDTH=100% HEIGHT=<%=BizSize%>>
							<IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1">
							</IFRAME>
						</TD>
					</TR>
				</TABLE>

<TEXTAREA CLASS="HIDDEN" NAME="txtSpread" tag="24" TABINDEX="-1"></TEXTAREA>
	<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" TABINDEX="-1">
	<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24" TABINDEX="-1">
	<INPUT TYPE=HIDDEN NAME="hPlant_Cd" tag="24" TABINDEX="-1">
	<INPUT TYPE=HIDDEN NAME="hSL_Cd" tag="24" TABINDEX="-1">
	<INPUT TYPE=HIDDEN NAME="hItem_Cd" tag="24" TABINDEX="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>           