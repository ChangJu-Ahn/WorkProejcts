
<%@ LANGUAGE="VBSCRIPT" %>
<!--'**********************************************************************************************
'*  1. Module Name		: Production
'*  2. Function Name	: 
'*  3. Program ID		: b1b01ma2.asp
'*  4. Program Name		: ǰ��������ȸ 
'*  5. Program Desc		:
'*  6. Component List	: 
'*  7. Modified date(First)	: 2000/12/18
'*  8. Modified date(Last)	: 2002/11/14
'*  9. Modifier (First)		: Jung Yu Kyung
'* 10. Modifier (Last)		: Hong Chang Ho
'* 11. Comment		:
'**********************************************************************************************-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/Cookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "b1b01ma2.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript">

Option Explicit

<!-- #Include file="../../inc/lgVariables.inc" -->

Dim StartDate

StartDate = UniConvDateAToB("<%=GetSvrDate%>", parent.gServerDateFormat, parent.gDateFormat)

Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<%Call loadInfTB19029A("Q", "*", "NOCOOKIE", "MA")%>
End Sub

Sub Form_Load()

    Call LoadInfTB19029
	Call ggoOper.FormatField(Document, "A", ggStrIntegeralPart, ggStrDeciPointPart, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec)
	
    Call ggoOper.LockField(Document, "N")											'��: Lock  Suitable  Field
   
	Call SetDefaultVal
   	Call InitComboBox
    Call InitVariables		

    Call InitSpreadSheet	
	Call SetToolbar("11000000000011")
	
	frm1.txtItemCd.focus
	Set gActiveElement = document.activeElement 
End Sub

Sub InitComboBox()
    On Error Resume Next
    Err.Clear
    
    Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("P1001", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboItemAcct, lgF0, lgF1, Chr(11))

	Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("P1002", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboItemClass, lgF0, lgF1, Chr(11))

End Sub

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>
<BODY TABINDEX="-1" SCROLL="NO">
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>ǰ��������ȸ</font></td>
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
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS=TD5 NOWRAP>ǰ��</TD>
									<TD CLASS=TD6 NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtItemCd" SIZE=20 MAXLENGTH=18 tag="11XXXU"  ALT="ǰ��"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenItemCd()">&nbsp;<INPUT TYPE=TEXT NAME="txtItemNm" SIZE=25 tag="14"></TD>
									<TD CLASS=TD5 NOWRAP>����� ǰ��Ŭ����</TD>
									<TD CLASS=TD6 NOWRAP><SELECT NAME="cboItemClass" ALT="����� ǰ��Ŭ����" STYLE="Width: 168px;" tag="11"><OPTION VALUE=""></OPTION></SELECT></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>ǰ��׷�</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtHighItemGroupCd" SIZE=15 MAXLENGTH=10 tag="11XXXU"  ALT="ǰ��׷�"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btHighItemGroupCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenItemGroup()" >&nbsp;<INPUT TYPE=TEXT NAME="txtHighItemGroupNm" SIZE=30 tag="14"></TD>
									<TD CLASS=TD5 NOWRAP>ǰ�����</TD>
									<TD CLASS=TD6 NOWRAP><SELECT NAME="cboItemAcct" ALT="ǰ�����" STYLE="Width: 168px;" tag="11"><OPTION VALUE=""></OPTION></SELECT></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>������</TD>
									<TD CLASS=TD6 NOWRAP>
										<script language =javascript src='./js/b1b01ma2_txtFinishStartDt_txtFinishStartDt.js'></script>
										&nbsp;~&nbsp;
										<script language =javascript src='./js/b1b01ma2_txtFinishEndDt_txtFinishEndDt.js'></script>
	
									</TD>
									<TD CLASS=TD5 NOWRAP>��ȿ����</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE="RADIO" NAME="rdoDefaultFlg" ID="rdoDefaultFlg1" CLASS="RADIO" tag="1X" Value="ALL" CHECKED><LABEL FOR="rdoDefaultFlg1">��ü</LABEL>
													     <INPUT TYPE="RADIO" NAME="rdoDefaultFlg" ID="rdoDefaultFlg2" CLASS="RADIO" tag="1X" Value="Y"><LABEL FOR="rdoDefaultFlg2">��</LABEL>
													     <INPUT TYPE="RADIO" NAME="rdoDefaultFlg" ID="rdoDefaultFlg3" CLASS="RADIO" tag="1X" Value="N"><LABEL FOR="rdoDefaultFlg3">�ƴϿ�</LABEL>
									</TD>				   
								</TR>		
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=* valign=top>
						<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD HEIGHT=* WIDTH=40%>
									<script language =javascript src='./js/b1b01ma2_vspdData_vspdData.js'></script>
								</TD>
								
								<TD HEIGHT=* WIDTH=60%>
									
										<TABLE <%=LR_SPACE_TYPE_60%>>
											<TR>
												<TD CLASS=TD5 NOWRAP>ǰ��</TD>
												<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd1" SIZE=25 MAXLENGTH=18 tag="24" ALT="ǰ��">&nbsp;<INPUT TYPE=TEXT NAME="txtItemNm1" SIZE=35 MAXLENGTH=40 tag="24" ALT="ǰ���"></TD>
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>ǰ�����ĸ�</TD>
												<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemDetailNm" SIZE=50 MAXLENGTH=50  tag="24" ALT="ǰ�����ĸ�"></TD>
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>ǰ�����</TD>
												<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemAcct" SIZE=25 MAXLENGTH=40  tag="24" ALT="ǰ�����"></TD>
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>���ش���</TD>
												<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtUnit" SIZE=5 MAXLENGTH=3  tag="24" ALT="���ش���"></TD>
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>ǰ��׷�</TD>
												<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemGroup" SIZE=15 MAXLENGTH=10  tag="24" ALT="ǰ��׷�">&nbsp;<INPUT TYPE=TEXT NAME="txtItemGroupNm" SIZE=40  tag="24" ALT="ǰ��׷��"></TD>
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>Phantom ����</TD>
												<TD CLASS=TD6 NOWRAP><INPUT TYPE="RADIO" NAME="rdoPhantomFlg" ID="rdoPhantomFlg1" CLASS="RADIO" tag="24X" Value="Y" CHECKED><LABEL FOR="rdoPhantomFlg1">��</LABEL>
												     				 <INPUT TYPE="RADIO" NAME="rdoPhantomFlg" ID="rdoPhantomFlg2" CLASS="RADIO" tag="24X" Value="N"><LABEL FOR="rdoPhantomFlg2">�ƴϿ�</LABEL></TD>
											<TR>
												<TD CLASS=TD5 NOWRAP>���ձ��ű���</TD>
												<TD CLASS=TD6 NOWRAP><INPUT TYPE="RADIO" NAME="rdoBlanketPurFlg" ID="rdoBlanketPurFlg1" CLASS="RADIO" tag="24X" Value="Y" CHECKED><LABEL FOR="rdoBlanketPurFlg1">��</LABEL>
												     				 <INPUT TYPE="RADIO" NAME="rdoBlanketPurFlg" ID="rdoBlanketPurFlg2" CLASS="RADIO" tag="24X" Value="N"><LABEL FOR="rdoBlanketPurFlg2">�ƴϿ�</LABEL></TD>
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>����ǰ��</TD>
												<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtBaseItem" SIZE=25  tag="24" ALT="����ǰ��">&nbsp;<INPUT TYPE=HIDDEN NAME="txtBaseItemNm" SIZE=40 tag="24" ALT="����ǰ���"></TD>
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>�����ǰ��Ŭ����</TD>
												<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtSumItemClass" SIZE=25 tag="24" ALT="�����ǰ��Ŭ����">
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>��ȿ����</TD>
												<TD CLASS=TD6 NOWRAP><INPUT TYPE="RADIO" NAME="rdoValidFlg" ID="rdoValidFlg1" CLASS="RADIO" tag="24X" Value="Y" CHECKED><LABEL FOR="rdoValidFlg1">��</LABEL>
												     				 <INPUT TYPE="RADIO" NAME="rdoValidFlg" ID="rdoValidFlg2" CLASS="RADIO" tag="24X" Value="N"><LABEL FOR="rdoValidFlg2">�ƴϿ�</LABEL></TD>
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>��������</TD>    				 
												<TD CLASS=TD6 NOWRAP><INPUT TYPE="RADIO" NAME="rdoPicFlg" ID="rdoPicFlg1" CLASS="RADIO" tag="24X" Value="Y" CHECKED><LABEL FOR="rdoPicFlg1">��</LABEL>
												     				 <INPUT TYPE="RADIO" NAME="rdoPicFlg" ID="rdoPicFlg2" CLASS="RADIO" tag="24X" Value="N"><LABEL FOR="rdoPicFlg2">�ƴϿ�</LABEL></TD>     				 
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>ǰ��԰�</TD>
												<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemSpec" SIZE=50 tag="24" ALT="ǰ��԰�">
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>Net�߷�</TD>
												<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtUnitWeight" SIZE=15  tag="24X3" ALT="Net�߷�" STYLE="TEXT-ALIGN: right">&nbsp;<INPUT TYPE=TEXT NAME="txtWeightUnit" align=top SIZE=5 MAXLENGTH=3  tag="24" ALT="����"></TD>
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>Gross�߷�</TD>
												<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtGrossWeight" SIZE=15  tag="24X3" ALT="Gross�߷�" STYLE="TEXT-ALIGN: right">&nbsp;<INPUT TYPE=TEXT NAME="txtGrossWeightUnit" align=top SIZE=5 MAXLENGTH=3  tag="24" ALT="����"></TD>
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>CBM(����)</TD>
												<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtCBM" SIZE=15  tag="24X3" ALT="CBM(����)" STYLE="TEXT-ALIGN: right">&nbsp;<INPUT TYPE=TEXT NAME="txtCBMInfo" align=top SIZE=40 MAXLENGTH=50  tag="24" ALT="CBM����"></TD>
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>�����ȣ</TD>
												<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtDrawingNm" SIZE=30  tag="24" ALT="�����ȣ"></TD>
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>HS�ڵ�</TD>
												<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtHsCode" SIZE=20  tag="24" ALT="HS�ڵ�">&nbsp;<INPUT TYPE=TEXT NAME="txtHsCodeUnit" align=top SIZE=5 MAXLENGTH=3  tag="24"  ALT="����"></TD>
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>��ȿ�Ⱓ</TD>
												<TD CLASS=TD6 NOWRAP>
													<script language =javascript src='./js/b1b01ma2_I361859291_txtValidFromDt.js'></script> &nbsp;~&nbsp;
													<script language =javascript src='./js/b1b01ma2_I817060148_txtValidToDt.js'></script>
												</TD>
											</TR>
										</TABLE>
									
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
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_30%>>
				<TR>
					<TD WIDTH=* ALIGN=RIGHT><A href="vbscript:JumpItem">ǰ����</A>&nbsp;|&nbsp;<A href="vbscript:JumpItemImage">ǰ��������</A></TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME></TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24"><INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24"><INPUT TYPE=HIDDEN NAME="hSumItemClass" tag="24">
<INPUT TYPE=HIDDEN NAME="hItemCd" tag="24"><INPUT TYPE=HIDDEN NAME="hItemAcct" tag="24">
<INPUT TYPE=HIDDEN NAME="hItemGroup" tag="24"><INPUT TYPE=HIDDEN NAME="hAvailableItem" tag="24">
<INPUT TYPE=HIDDEN NAME="hStartDt" tag="24"><INPUT TYPE=HIDDEN NAME="hEndDt" tag="24">
<INPUT TYPE=HIDDEN NAME="txtUserId" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
