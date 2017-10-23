<%@ LANGUAGE="VBSCRIPT" %>

<!--
'**********************************************************************************************
'*  1. Module Name			: Production
'*  2. Function Name		: Production Order Management
'*  3. Program ID			: p4113ma1.asp
'*  4. Program Name			: Production Order Status
'*  5. Program Desc			: List Production Order
'*  6. Comproxy List		: +B19029LookupNumericFormat
'*	   Biz ASP  List		: +p3211mba.asp		List Production Order Header
'*  7. Modified date(First)	: 2000/04/12
'*  8. Modified date(Last)	: 2003/05/20
'*  9. Modifier (First)		: Kim, GyoungDon
'* 10. Modifier (Last)		: Chen, JaeHyun
'* 11. Comment
'*                          : Order Number���� �ڸ��� ����(2003.04.14) Park Kye Jin				:
'**********************************************************************************************
-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- '#########################################################################################################
'												1. �� �� �� 
'#########################################################################################################-->
<!-- '******************************************  1.1 Inc ����   **********************************************
'	���: Inc. Include
'********************************************************************************************************* -->
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->
<!--'==========================================  1.1.1 Style Sheet  ======================================
'==========================================================================================================-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'==========================================  1.1.2 ���� Include   ======================================
'==========================================================================================================-->
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliVariables.vbs"></SCRIPT> 
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "p4113ma1.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit                                                             '��: indicates that All variables must be declared in advance

Dim LocSvrDate
	LocSvrDate = "<%=GetSvrDate%>"
	
'==========================================  1.2.0 Common variables =====================================
'	1. Common variables
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================


'==========================================  2.2.6 InitComboBox()  =======================================
'	Name : InitComboBox()
'	Description : Combo Display
'=========================================================================================================
Sub InitComboBox()

	Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("P3211", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboOrderType, lgF0, lgF1, Chr(11))
    
    Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("P1017", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboOrderStatus, lgF0, lgF1, Chr(11))

	frm1.cboOrderType.value = ""
	frm1.cboOrderStatus.value = ""
    
End Sub

'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'========================================================================================
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029A("Q", "P", "NOCOOKIE", "QA") %>
End Sub

'==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(���� Include ���Ͽ� ����)�� �����ʱ�ȭ �� ȭ���ʱ�ȭ�� �ϱ� ���� �Լ��� Call�ϴ� �κ� 
'=========================================================================================================
Sub Form_Load()

    Call LoadInfTB19029                                                     				'��: Load table , B_numeric_format
    
    Call ggoOper.FormatField(Document, "A",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)

    Call ggoOper.LockField(Document, "N")                                          			'��: Lock  Suitable  Field
    Call InitSpreadSheet                                                    				'��: Setup the Spread sheet

	Call SetDefaultVal
    Call InitVariables		'��: Initializes local global variables
    Call InitComboBox
    Call InitSpreadComboBox

    Call SetToolBar("11000000000011")										'��: ��ư ���� ���� 
	
	If Parent.gPlant <> "" Then
		frm1.txtPlantCd.Value = Parent.gPlant
		frm1.txtPlantNm.Value = Parent.gPlantNm
		frm1.txtItemCd.focus
		Set gActiveElement = document.activeElement 
	Else
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement 
	End If

End Sub

</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>����������Ȳ��ȸ</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* align=right><A href="vbscript:OpenReworkRef()">���۾�����</A> | <A href="vbscript:OpenPartRef()">��ǰ����</A> | <A href="vbscript:OpenOprRef()">��������</A> | <A href="vbscript:OpenProdRef()">��������</A> | <A href="vbscript:OpenRcptRef()">�԰���</A> | <A href="vbscript:OpenConsumRef()">��ǰ�Һ񳻿�</A></TD>
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
									<TD CLASS="TD6">
										<script language =javascript src='./js/p4113ma1_OBJECT1_txtProdFromDt.js'></script>
										&nbsp;~&nbsp;
										<script language =javascript src='./js/p4113ma1_OBJECT2_txtProdtODt.js'></script>
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
									<TD CLASS=TD5 NOWRAP>���û���</TD>
									<TD CLASS=TD6 NOWRAP><SELECT NAME="cboOrderStatus" ALT="���û���" STYLE="Width: 98px;" tag="11"><OPTION VALUE=""></OPTION></SELECT></TD>
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
					<TD><!-- ù��° �� ���� -->
					<TABLE <%=LR_SPACE_TYPE_20%>>
						<TR>
							<TD HEIGHT="100%" colspan=4>
								<script language =javascript src='./js/p4113ma1_A_vspdData.js'></script>
							</TD>
						</TR>
						<TR>
							<TD HEIGHT=5 WIDTH=100% colspan=4></TD>
						</TR>
						<TR>
							<TD WIDTH=66% colspan=2>
							<FIELDSET valign=top>
								<LEGEND>���ش�������</LEGEND>
								<TABLE CLASS="TB2" CELLSPACING=0>
									<TR>
										<TD CLASS=TD5 NOWRAP>���ش���</TD>
										<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtBaseUnit" SIZE=5 MAXLENGTH=3 tag="24" ALT="������"></TD>
										<TD CLASS=TD5 NOWRAP></TD>
										<TD CLASS=TD6 NOWRAP></TD>
									</TR>	
									<TR>
										<TD CLASS=TD5 NOWRAP>��������</TD>
										<TD CLASS=TD6 NOWRAP>
											<script language =javascript src='./js/p4113ma1_I431471735_txtOrderQty.js'></script>
										</TD>
										<TD CLASS=TD5 NOWRAP>��������</TD>
										<TD CLASS=TD6 NOWRAP>
											<script language =javascript src='./js/p4113ma1_I958357095_txtProdQty.js'></script>
										</TD>
									</TR>	
									<TR>
										<TD CLASS=TD5 NOWRAP>��ǰ����</TD>
										<TD CLASS=TD6 NOWRAP>
											<script language =javascript src='./js/p4113ma1_I521103113_txtGoodQty.js'></script>
										</TD>
										<TD CLASS=TD5 NOWRAP>�ҷ�����</TD>
										<TD CLASS=TD6 NOWRAP>
											<script language =javascript src='./js/p4113ma1_I146408760_txtBadQty.js'></script>
										</TD>
									</TR>
									<TR>
										<TD CLASS=TD5 NOWRAP>ǰ����ǰ</TD>
										<TD CLASS=TD6 NOWRAP>
											<script language =javascript src='./js/p4113ma1_I198610717_txtInspGoodQty.js'></script>
										</TD>
										<TD CLASS=TD5 NOWRAP>ǰ���ҷ�</TD>
										<TD CLASS=TD6 NOWRAP>
											<script language =javascript src='./js/p4113ma1_I951398158_txtInspBadQty.js'></script>
										</TD>
									</TR>									
									<TR>
										<TD CLASS=TD5 NOWRAP>�԰����</TD>
										<TD CLASS=TD6 NOWRAP>
											<script language =javascript src='./js/p4113ma1_I183605903_txtRcptQty.js'></script>
										</TD>
										<TD CLASS=TD5 NOWRAP></TD>
										<TD CLASS=TD6 NOWRAP></TD>
									</TR>
								</TABLE>	
							</FIELDSET>			
							</TD>
							<TD WIDTH=34% colspan=2>
							<FIELDSET valign=top>
								<LEGEND>��������</LEGEND>
								<TABLE CLASS="TB2" CELLSPACING=0>
									<TR>
										<TD CLASS=TD5 NOWRAP>����������</TD> 
										<TD CLASS=TD6 NOWRAP>
											<script language =javascript src='./js/p4113ma1_I706934762_txtPlanStratDt.js'></script>
										</TD>
									</TR>
									<TR>
										<TD CLASS=TD5 NOWRAP>�ϷΌ����</TD>
										<TD CLASS=TD6 NOWRAP>
											<script language =javascript src='./js/p4113ma1_OBJECT3_txtPlanEndDt.js'></script>
										</TD>
									</TR>		
									<TR>
										<TD CLASS=TD5 NOWRAP>�۾�������</TD>
										<TD CLASS=TD6 NOWRAP>
											<script language =javascript src='./js/p4113ma1_I357572263_txtReleaseDt.js'></script>
										</TD>
									</TR>
									<TR>
										<TD CLASS=TD5 NOWRAP>��������</TD>
										<TD CLASS=TD6 NOWRAP>
											<script language =javascript src='./js/p4113ma1_OBJECT4_txtRealStratDt.js'></script>
										</TD>
									</TR>	
									<TR>
										<TD CLASS=TD5 NOWRAP>�ǿϷ���</TD>
										<TD CLASS=TD6 NOWRAP>
											<script language =javascript src='./js/p4113ma1_I465394866_txtRealEndDt.js'></script>
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
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC= "../../blank.htm " WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX = "-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" TABINDEX = "-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24"><INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24"><INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMajorFlg" tag="24"><INPUT TYPE=HIDDEN NAME="hPlantCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hItemCd" tag="24"><INPUT TYPE=HIDDEN NAME="hProdOrderNo" tag="24"><INPUT TYPE=HIDDEN NAME="hOrderType" tag="24">
<INPUT TYPE=HIDDEN NAME="hProdFromDt" tag="24"><INPUT TYPE=HIDDEN NAME="hProdToDt" tag="24"><INPUT TYPE=HIDDEN NAME="hOrderStatus" tag="24">
<INPUT TYPE=HIDDEN NAME="hTrackingNo" tag="24"><INPUT TYPE=HIDDEN NAME="hItemGroupCd" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
