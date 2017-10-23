
<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name			: Production
'*  2. Function Name		: 
'*  3. Program ID			: p4611ma1.asp
'*  4. Program Name			: Order Closing
'*  5. Program Desc			:
'*  6. Component List		: 
'*  7. Modified date(First) : 2000/04/04
'*  8. Modified date(Last)  : 2003/03/23
'*  9. Modifier (First)     : Kim, Gyoung-Don
'* 10. Modifier (Last)      : Chen, Jae Hyun
'* 11. Comment				:
'* 12. History              : Tracking No 9�ڸ����� 25�ڸ��� ����(2003.03.03)
'*                          : Order Number���� �ڸ��� ����(2003.04.14) Park Kye Jin
'**********************************************************************************************
-->

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
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliRDSQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "p4611ma1_KO441.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit                                                        '��: indicates that All variables must be declared in advance

<!-- #Include file="../../inc/lgvariables.inc" -->

Dim LocSvrDate
Dim StartDate
Dim EndDate

	LocSvrDate = UniConvDateAToB("<%=GetSvrDate%>",parent.gServerDateFormat,parent.gDateFormat)													
	StartDate = UNIDateAdd("D",-10,LocSvrDate, parent.gDateFormat)	'��: �ʱ�ȭ�鿡 �ѷ����� ó�� ��¥ 
	EndDate = UNIDateAdd("D", 20,LocSvrDate, parent.gDateFormat)	'��: �ʱ�ȭ�鿡 �ѷ����� ������ ��¥ 

'=========================================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'=========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "*", "NOCOOKIE", "MA") %>
End Sub

'==========================================================================================================
'	Name : InitComboBox()
'	Description : Combo Display
'========================================================================================================= 
Sub InitComboBox()
	Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("P1015", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboProdMgr, lgF0, lgF1, Chr(11))
    Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("P3211", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboOrderType, lgF0, lgF1, Chr(11))
    frm1.cboProdMgr.value = ""
    frm1.cboOrderType.value = ""
End Sub

'==========================================================================================================
'	Name : Form_Load()
'	Description : Window On Load(���� Include ���Ͽ� ����)�� �����ʱ�ȭ �� ȭ���ʱ�ȭ�� �ϱ� ���� �Լ��� Call�ϴ� �κ� 
'==========================================================================================================
Sub Form_Load()

    Call LoadInfTB19029                                                     				'��: Load table , B_numeric_format
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)    	
	Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)    
    Call ggoOper.LockField(Document, "N")                                          			'��: Lock  Suitable  Field
    Call InitSpreadSheet                                                    				'��: Setup the Spread sheet

    Call InitVariables																		'��: Initializes local global variables

    Call SetToolbar("11000000000011")														'��: ��ư ���� ���� 

    Call SetDefaultVal

    Call InitComboBox
    
    If parent.gPlant <> "" Then
		frm1.txtPlantCd.value = UCase(parent.gPlant)
		frm1.txtPlantNm.value = parent.gPlantNm
		frm1.txtItemCd.focus 
		Set gActiveElement = document.activeElement
	Else
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement
	End If
    
End Sub

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>
<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE  <%=LR_SPACE_TYPE_00%>>
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>������������</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* align=right><A href="vbscript:OpenPartRef()">��ǰ����</A> | <A href="vbscript:OpenOprRef()">��������</A> | <A href="vbscript:OpenProdRef()">��������</A> | <A href="vbscript:OpenRcptRef()">�԰���</A> | <A href="vbscript:OpenConsumRef()">��ǰ�Һ񳻿�</A></TD>
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
									<TD CLASS=TD5 NOWRAP>����</TD>
									<TD CLASS=TD6 NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtPlantCd" SIZE=6 MAXLENGTH=4 tag="12xxxU" ALT="����"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenConPlant()">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=25 tag="14"></TD>
									<TD CLASS=TD5 NOWRAP>�۾���ȹ����</TD>
									<TD CLASS=TD6 NOWRAP>												
										<script language =javascript src='./js/p4611ma1_OBJECT1_txtProdFromDt.js'></script>&nbsp;~&nbsp;
										<script language =javascript src='./js/p4611ma1_OBJECT2_txtProdtODt.js'></script>
									</TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>ǰ��</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd" SIZE=18 MAXLENGTH=18 tag="11xxxU" ALT="ǰ��"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenItemInfo frm1.txtItemCd.value">&nbsp;<INPUT TYPE=TEXT NAME="txtItemNm" SIZE=25 tag="14"></TD>
									<TD CLASS=TD5 NOWRAP>����������ȣ</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtProdOrderNo" SIZE=18 MAXLENGTH=18 tag="11xxxU" ALT="����������ȣ"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnProdOrder" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenProdOrderNo()"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>ǰ��׷�</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemGroupCd" SIZE=15 MAXLENGTH=10 tag="11XXXU" ALT="ǰ��׷�"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemGroupCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenItemGroup()">&nbsp;<INPUT TYPE=TEXT NAME="txtItemGroupNm" SIZE=25 MAXLENGTH=40 tag="14" ALT="ǰ��׷��"></TD>
									<TD CLASS=TD5 NOWRAP>Tracking No.</TD>
									<TD CLASS=TD6 NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtTrackingNo" SIZE=25 MAXLENGTH=25 tag="11xxxU" ALT="Tracking No."><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnTrackingNo" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenTrackingInfo frm1.txtTrackingNo.value,0"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>��������</TD>
									<TD CLASS=TD6 NOWRAP><SELECT NAME="cboProdMgr" ALT="��������" STYLE="Width: 98px;" tag="11"><OPTION VALUE=""></OPTION></SELECT></TD>	
									<TD CLASS=TD5 NOWRAP>���ñ���</TD>
									<TD CLASS=TD6 NOWRAP><SELECT NAME="cboOrderType" ALT="���ñ���" STYLE="Width: 98px;" tag="11"><OPTION VALUE=""></OPTION></SELECT></TD>
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
					<TABLE <%=LR_SPACE_TYPE_60%>>
						<TR>
							<TD HEIGHT="100%" colspan=4>
								<script language =javascript src='./js/p4611ma1_I837250530_vspdData.js'></script>
							</TD>
						</TR>
						<TR>
							<TD HEIGHT=5 WIDTH=100% colspan=4></TD>
						</TR>
						<TR>
							<TD WIDTH=40% colspan=1>
							<FIELDSET valign=top>
								<LEGEND>����</LEGEND>
								<TABLE CLASS="TB2" CELLSPACING=0>
									<TR>
										<TD CLASS=TD5 NOWRAP>����������</TD>
										<TD CLASS=TD6 NOWRAP>												
											<script language =javascript src='./js/p4611ma1_OBJECT1_txtPlanStratDt.js'></script>
										</TD>
									</TR>	
									<TR>
										<TD CLASS=TD5 NOWRAP>�ϷΌ����</TD>
										<TD CLASS=TD6 NOWRAP>												
											<script language =javascript src='./js/p4611ma1_OBJECT2_txtPlanEndDt.js'></script>
										</TD>
									</TR>	
									<TR>
										<TD CLASS=TD5 NOWRAP>������ȹ����</TD>
										<TD CLASS=TD6 NOWRAP>											
											<script language =javascript src='./js/p4611ma1_OBJECT1_txtPlannedStratDt.js'></script>
										</TD>
									</TR>	
									<TR>
										<TD CLASS=TD5 NOWRAP>�Ϸ��ȹ����</TD>
										<TD CLASS=TD6 NOWRAP>												
											<script language =javascript src='./js/p4611ma1_OBJECT2_txtPlannedEndDt.js'></script>
										</TD>
									</TR>	
									<TR>
										<TD CLASS=TD5 NOWRAP>�۾�������</TD>
										<TD CLASS=TD6 NOWRAP>												
											<script language =javascript src='./js/p4611ma1_I618068617_txtReleaseDt.js'></script>
										</TD>
									</TR>	
									<TR>
										<TD CLASS=TD5 NOWRAP>��������</TD>
										<TD CLASS=TD6 NOWRAP>												
											<script language =javascript src='./js/p4611ma1_OBJECT1_txtRealStratDt.js'></script>
										</TD>
									</TR>
									<TR>
										<TD CLASS=TD5 NOWRAP>���û���</TD>
										<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtOrderStatus" SIZE=8 MAXLENGTH=3 tag="24" ALT="���û���"></TD>
									</TR>										
								</TABLE>	
							</FIELDSET>			
							</TD>						
							<TD WIDTH=30% colspan=1>
							<FIELDSET valign=top>
								<LEGEND>������������</LEGEND>
								<TABLE CLASS="TB2" CELLSPACING=0>
									<TR>
										<TD CLASS=TD5 NOWRAP>��������</TD>
										<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtOrderUnit" SIZE=8 MAXLENGTH=3 tag="24" ALT="��������"></TD>
									</TR>
									<TR>
										<TD CLASS=TD5 NOWRAP>��������</TD>
										<TD CLASS=TD6 NOWRAP>
											<script language =javascript src='./js/p4611ma1_fpDoubleSingle4_txtOrderQty.js'></script>
										</TD>
									</TR>	
									<TR>
										<TD CLASS=TD5 NOWRAP>��������</TD>
										<TD CLASS=TD6 NOWRAP>
											<script language =javascript src='./js/p4611ma1_fpDoubleSingle5_txtProdQty.js'></script> 
										</TD>
									</TR>
									<TR>
										<TD CLASS=TD5 NOWRAP>��ǰ����</TD>
										<TD CLASS=TD6 NOWRAP>
											<script language =javascript src='./js/p4611ma1_fpDoubleSingle6_txtGoodQty.js'></script>
										</TD>
									</TR>
									<TR>
										<TD CLASS=TD5 NOWRAP>�ҷ�����</TD>
										<TD CLASS=TD6 NOWRAP>
											<script language =javascript src='./js/p4611ma1_fpDoubleSingle7_txtBadQty.js'></script>
										</TD>
									</TR>
									<TR>
										<TD CLASS=TD5 NOWRAP>�԰����</TD>
										<TD CLASS=TD6 NOWRAP>
											<script language =javascript src='./js/p4611ma1_fpDoubleSingle8_txtRcptQty.js'></script>
										</TD>
									</TR>
									<TR>
										<TD CLASS=TD5 NOWRAP>�԰������</TD>
										<TD CLASS=TD6 NOWRAP>
											<script language =javascript src='./js/p4611ma1_fpDoubleSingle9_txtUnRcptQty.js'></script>
										</TD>
									</TR>
								</TABLE>	
							</FIELDSET>
							</TD>
							<TD WIDTH=30% colspan=1>
							<FIELDSET valign=top>
								<LEGEND>���ش�������</LEGEND>
								<TABLE CLASS="TB2" CELLSPACING=0>
									<TR>
										<TD CLASS=TD5 NOWRAP>���ش���</TD>
										<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtBaseUnit" SIZE=8 MAXLENGTH=3 tag="24" ALT="���ش���"></TD>
									</TR>
									<TR>
										<TD CLASS=TD5 NOWRAP>��������</TD>
										<TD CLASS=TD6 NOWRAP>
											<script language =javascript src='./js/p4611ma1_fpDoubleSingle1_txtOrderQty1.js'></script>
										</TD>
									</TR>	
									<TR>
										<TD CLASS=TD5 NOWRAP>��������</TD>
										<TD CLASS=TD6 NOWRAP>
											<script language =javascript src='./js/p4611ma1_fpDoubleSingle2_txtProdQty1.js'></script>
										</TD>
									</TR>
									<TR>
										<TD CLASS=TD5 NOWRAP>��ǰ����</TD>
										<TD CLASS=TD6 NOWRAP>
											<script language =javascript src='./js/p4611ma1_fpDoubleSingle3_txtGoodQty1.js'></script>
										</TD>
									</TR>
									<TR>
										<TD CLASS=TD5 NOWRAP>�ҷ�����</TD>
										<TD CLASS=TD6 NOWRAP>
											<script language =javascript src='./js/p4611ma1_fpDoubleSingle4_txtBadQty1.js'></script>
										</TD>
									</TR>
									<TR>
										<TD CLASS=TD5 NOWRAP>�԰����</TD>
										<TD CLASS=TD6 NOWRAP>
											<script language =javascript src='./js/p4611ma1_fpDoubleSingle5_txtRcptQty1.js'></script>
										</TD>
									</TR>
									<TR>
										<TD CLASS=TD5 NOWRAP>�԰������</TD>
										<TD CLASS=TD6 NOWRAP>
											<script language =javascript src='./js/p4611ma1_fpDoubleSingle6_txtUnRcptQty1.js'></script>
										</TD>
									</TR>
								</TABLE>	
							</FIELDSET>
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
    <TR HEIGHT="20">
      <TD WIDTH="100%">
		<TABLE <%=LR_SPACE_TYPE_30%>>
			<TR>
			  <TD WIDTH=10>&nbsp;</TD>
			  <TD WIDTH="*" align="left"><a><button name="btnAutoSel" class="clsmbtn">��ü����</button></a></TD>
			  <TD WIDTH=10><a><button name="btnAutoTest" class="clsmbtn">�׽�Ʈ</button></a>&nbsp;</TD>
			</TR>
		</TABLE>
      </TD>
    </TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC= "../../blank.htm " WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX = "-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" TABINDEX = "-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24"><INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24"><INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMajorFlg" tag="24"><INPUT TYPE=HIDDEN NAME="hPlantCd" tag="24"><INPUT TYPE=HIDDEN NAME="hItemCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hProdOrderNo" tag="24"><INPUT TYPE=HIDDEN NAME="hProdFromDt" tag="24"><INPUT TYPE=HIDDEN NAME="hProdToDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hOrderType" tag="24"><INPUT TYPE=HIDDEN NAME="hProdMgr" tag="24">
<INPUT TYPE=HIDDEN NAME="hTrackingNo" tag="24"><INPUT TYPE=HIDDEN NAME="hItemGroupCd" tag="24">
<P ID="divTextArea"></P>
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
