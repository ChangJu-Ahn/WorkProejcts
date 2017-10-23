<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Prucurement
'*  2. Function Name        : 
'*  3. Program ID           : M3112MA1
'*  4. Program Name         : ���ֳ������ 
'*  5. Program Desc         : ���ֳ������ 
'*  6. Component List       : 
'*  7. Modified date(First) : 1999/09/10
'*  8. Modified date(Last)  : 2003/05/23
'*  9. Modifier (First)     : Shin Jin Hyun
'* 10. Modifier (Last)      : Kang Su Hwan
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
<!-- '******************************************  1.1 Inc ����   **********************************************
'	���: Inc. Include
'********************************************************************************************************* !-->
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!--'==========================================  1.1.1 Style Sheet  ======================================
'==========================================================================================================!-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'==========================================  1.1.2 ���� Include   ======================================
'==========================================================================================================!-->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs"></SCRIPT> 
<SCRIPT LANGUAGE="VBScript"   SRC="m3112ma1.vbs"></SCRIPT> 

<SCRIPT LANGUAGE=VBSCRIPT>
<!-- #Include file="../../inc/lgvariables.inc" -->

'===================  LoadInfTB19029()  ===========================================================
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029A("I", "*","NOCOOKIE","MA") %>
<% Call LoadBNumericFormatA("I", "*","NOCOOKIE","MA") %>
End Sub
'==========================================  2.1.1 InitVariables()  ======================================



Sub InitVariables()

    lgIntFlgMode = Parent.OPMD_CMODE              
    lgBlnFlgChgValue = False           
    lgIntGrpCount = 0                  
    lgStrPrevKey = ""                  
    lgLngCurRows = 0                   
    frm1.vspdData.MaxRows = 0
    
End Sub
'==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(���� Include ���Ͽ� ����)�� �����ʱ�ȭ �� ȭ���ʱ�ȭ�� �ϱ� ���� �Լ��� Call�ϴ� �κ� 
'========================================================================================================= 
Sub Form_Load()

	Call LoadInfTB19029    	
    Call initFormatField() 
    Call InitSpreadSheet                    
    Call SetDefaultVal
    Call InitVariables                      
    Call CookiePage(0)
	' === 2005.07.15 �ܰ� �ϰ��ҷ����� ���� ���� =======
	Call SetPriceType
	' === 2005.07.15 �ܰ� �ϰ��ҷ����� ���� ���� =======    
    
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
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>���ֳ���</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=*>&nbsp;</TD>
					<TD WIDTH=* align=right><A href="vbscript:OpenReqRef">���ſ�û����</A> </TD>
					<TD WIDTH=10>&nbsp;</TD>
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
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS="TD5" NOWRAP>���ֹ�ȣ</TD>
									<TD CLASS="TD6" NOWRAP><INPUT CLASS = REQUIRED STYLE = "text-transform:uppercase" TYPE=TEXT NAME="txtPoNo"  SIZE=29 MAXLENGTH=18 ALT="���ֹ�ȣ" tag="12NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnORGCd1" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPoNo()"></TD>
									<TD CLASS="TD6" NOWRAP>&nbsp;</TD>
									<TD CLASS="TD6" NOWRAP>&nbsp;</TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=* VALIGN=TOP>
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD CLASS="TD5" NOWRAP>��������</TD>
								<TD CLASS="TD6" NOWRAP><INPUT  CLASS = protected readonly = True TabIndex = -1 TYPE=TEXT ALT="��������" NAME="txtPoTypeCd" SIZE=10 tag="24X">
													   <INPUT  CLASS = protected readonly = True TabIndex = -1 TYPE=TEXT NAME="txtPoTypeNm" SIZE=20 ALT ="��������" tag="24X" ></TD>
								<TD CLASS="TD5" NOWRAP>������</TD>
								<TD CLASS="TD6" NOWRAP><script language =javascript src='./js/m3112ma1_fpDateTime2_txtPoDt.js'></script></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>����ó</TD>
								<TD CLASS="TD6" NOWRAP><INPUT  CLASS = protected readonly = True TabIndex = -1 TYPE=TEXT ALT="����ó" NAME="txtSupplierCd" SIZE=10 tag="24X">
													   <INPUT  CLASS = protected readonly = True TabIndex = -1 TYPE=TEXT NAME="txtSupplierNm" SIZE=20 ALT ="����ó" tag="24X" ></TD>
								<TD CLASS="TD5" NOWRAP>���ű׷�</TD>
								<TD CLASS="TD6" NOWRAP><INPUT  CLASS = protected readonly = True TabIndex = -1 TYPE=TEXT ALT="���ű׷�" NAME="txtGroupCd" SIZE=10 tag="24X">
													   <INPUT  CLASS = protected readonly = True TabIndex = -1 TYPE=TEXT NAME="txtGroupNm" SIZE=20 ALT ="���ű׷�" tag="24X" ></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>���ּ��ݾ�</TD>
								<TD CLASS="TD6" NOWRAP><script language =javascript src='./js/m3112ma1_fpDoubleSingle1_txtGrossAmt.js'></script></td>
								<TD CLASS="TD5" NOWRAP>ȭ��</TD>
								<TD CLASS="TD6" NOWRAP><INPUT  CLASS = protected readonly = True TabIndex = -1 TYPE=TEXT ALT="ȭ��" NAME="txtCurr" SIZE=10 tag="24X">
													   <INPUT  CLASS = protected readonly = True TabIndex = -1 TYPE=TEXT NAME="txtCurrNm" SIZE=20 ALT ="ȭ��" tag="24X" ></TD>
							</TR>
							<TR>
								<TD HEIGHT=100% WIDTH=100% COLSPAN=4>
									<script language =javascript src='./js/m3112ma1_I838412088_vspdData.js'></script>
								</TD>
							</TR>
						</TABLE>
		
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
    <tr>
		<TD <%=HEIGHT_TYPE_01%>></TD>
    </tr>
    <tr HEIGHT="20">
		<td WIDTH="100%">
			<table <%=LR_SPACE_TYPE_30%>>
				<tr>
					<TD WIDTH=10>&nbsp;</TD>
					<td align="Left"><a><button name="btnCfmSel" id="btnCfm" class="clsmbtn" ONCLICK="vbscript:Cfm()">Ȯ��</button></a>&nbsp;
					<BUTTON NAME="btnCallPrice" CLASS="CLSMBTN">�ܰ��ҷ�����</BUTTON>&nbsp</td>					
					<td WIDTH="*" align=right><a href="VBSCRIPT:CookiePage(1)">���ֵ��</a> | <a href="VBSCRIPT:CookiePage(2)">�����</a> | <A href="vbscript:JumpOrderRun()">���ְ�������</A></td>
					<TD WIDTH=10>&nbsp;</TD>
				</tr>
			</table>
		</td>
    </tr>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> SRC="m3112mb1.asp" FRAMEBORDER=0 SCROLLING=No noresize framespacing=0 tabindex="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<P ID="divTextArea"></P>
<TEXTAREA CLASS="hidden"  NAME="txtSpread" tag="24" TABINDEX="-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="txtRelease" tag="14">
<INPUT TYPE=HIDDEN NAME="txthdnPoNo" tag="14">
<INPUT TYPE=HIDDEN NAME="txtQuerytype" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnDlvyDt" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnImportFlg" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnSubContraFlg" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnClsflg" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnReleaseflg" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnRcptflg" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnRetflg" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnMvmtType" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnXch" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnRcptType" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnMode" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnTrackingflg" tag="14">
<INPUT TYPE=HIDDEN NAME="txtHMaintNo" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnVATType" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnVATRate" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnVATINCFLG" tag="1">
<INPUT TYPE=HIDDEN NAME="hdnXchRateOp"  tag="14">
<INPUT TYPE=HIDDEN NAME="hdnIVFlg"  tag="14">
<INPUT TYPE=HIDDEN NAME="hdnmaxrow"  tag="14">
<INPUT TYPE=HIDDEN NAME="hdnreference"  tag="14">

<!-- 2005.07.15 -->
<INPUT TYPE=HIDDEN NAEM="hdnPriceType" tag="24">
</FORM>

    <DIV ID="MousePT" NAME="MousePT">
    <iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
    </DIV>

</BODY>
</HTML>
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                        