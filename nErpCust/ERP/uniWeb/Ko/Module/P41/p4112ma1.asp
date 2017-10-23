
<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name			: Production
'*  2. Function Name		: Production Order Management
'*  3. Program ID			: p4112ma1.asp
'*  4. Program Name			: Manage Production Order (Multi)
'*  5. Program Desc			: Create, Update, Delete Production Order
'*  6. Comproxy List		: +B19029LookupNumericFormat
'*	   Biz ASP  List		: +p4112mb1.asp		List Production Order Header
'*							  +p4112mb2.asp		Manage Production Order
'*							  +p4112mb0.asp	    LookUp Item By Plant
'*							  +p4112mb3.asp	    Release Production Order
'*							  +p4110ma1.asp		Order Explosion	
'*  7. Modified date(First)	: 2000/04/12
'*  8. Modified date(Last)	: 2005.09/29
'*  9. Modifier (First)		: Park, BumSoo
'* 10. Modifier (Last)		: Chen, Jaehyun
'* 11. Comment	
'*                          : Order Number���� �ڸ��� ����(2003.04.14) Park Kye Jin			:
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
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/Cookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "p4112ma1.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit									'��: indicates that All variables must be declared in advance

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

	'****************************
	'List Minor code(Order Type)
	'****************************
	Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("P3211", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboOrderType, lgF0, lgF1, Chr(11))
    
    frm1.cboOrderType.value = ""

End Sub

'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'========================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "*","NOCOOKIE","MA") %>
End Sub

'==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(���� Include ���Ͽ� ����)�� �����ʱ�ȭ �� ȭ���ʱ�ȭ�� �ϱ� ���� �Լ��� Call�ϴ� �κ� 
'========================================================================================================= 
Sub Form_Load()

    Call LoadInfTB19029                                                     				'��: Load table , B_numeric_format
    
	Call ggoOper.FormatField(Document, "A",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
	
    Call ggoOper.LockField(Document, "N")                                          			'��: Lock  Suitable  Field
    Call InitSpreadSheet                                                    				'��: Setup the Spread sheet

	Call SetDefaultVal
    Call InitVariables																		'��: Initializes local global variables
    Call InitComboBox()
    Call InitSpreadComboBox()
    Call SetToolBar("11001101001111")														'��: ��ư ���� ���� 
	
	If parent.ReadCookie("txtPlantCd") <> "" Then
		Call SetCookieVal
	End If
	
	If parent.gPlant <> "" and frm1.txtPlantCd.Value = "" Then
		frm1.txtPlantCd.Value = parent.gPlant
		frm1.txtPlantNm.Value = parent.gPlantNm
		
		Call LookUpInvClsDt
		Call txtPlantCd_OnChange
		frm1.txtItemCd.focus
		Set gActiveElement = document.activeElement
	Else
		If frm1.txtPlantCd.Value <> "" Then
			Call LookUpInvClsDt
			Call txtPlantCd_OnChange
		End If
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>������������(Multi)</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* align=right><A href="vbscript:OpenStockRef()">�����Ȳ</A> | <A href="vbscript:OpenPartRef()">��ǰ����</A> | <A href="vbscript:OpenOprRef()">��������</A></TD>
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
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name=txtProdFromDt CLASS=FPDTYYYYMMDD title=FPDATETIME tag="12" ALT="������"></OBJECT>');</SCRIPT>
										&nbsp;~&nbsp;
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name=txtProdToDt CLASS=FPDTYYYYMMDD title=FPDATETIME tag="12" ALT="������"></OBJECT>');</SCRIPT>
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
									<TD CLASS=TD6 NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtTrackingNo" SIZE=25 MAXLENGTH=25 tag="11xxxU" ALT="Tracking No."><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnTrackingNo" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenTrackingInfo()"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>MRP Run��ȣ</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtMRPRunNo" SIZE=20 MAXLENGTH=18 tag="11xxxU" ALT="MRP Run��ȣ"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnMRPRunNo" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenMRPRunNo()"></TD>
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
					<TABLE <%=LR_SPACE_TYPE_20%>>
						<TR>
							<TD HEIGHT="100%" colspan=4>
								<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData ID = "A" WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD"> <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
							</TD>
						</TR>
						<TR>
							<TD HEIGHT=5 WIDTH=100% colspan=4></TD>
						</TR>
						<TR>
							<TD WIDTH=35% colspan=2>
							<FIELDSET valign=top>
								<LEGEND>������������</LEGEND>
								<TABLE CLASS="TB2" CELLSPACING=0>
									<TR>
										<TD CLASS=TD5 NOWRAP>������������</TD>
										<TD CLASS=TD6 NOWRAP>
											<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name=txtFixedMRPQty CLASS = FPDS115 title=FPDOUBLESINGLE SIZE="20" MAXLENGTH="20" ALT="������������" tag="24X3"></OBJECT>');</SCRIPT>
										</TD>
									</TR>	
									<TR>
										<TD CLASS=TD5 NOWRAP>�ּҿ�������</TD>
										<TD CLASS=TD6 NOWRAP>
											<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name=txtMinMRPQty CLASS = FPDS115 title=FPDOUBLESINGLE SIZE="20" MAXLENGTH="20" ALT="�ּҿ�������" tag="24X3"></OBJECT>');</SCRIPT>
										</TD>
									</TR>	
									<TR>
										<TD CLASS=TD5 NOWRAP>�ִ��������</TD>
										<TD CLASS=TD6 NOWRAP>
											<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name=txtMaxMRPQty CLASS = FPDS115 title=FPDOUBLESINGLE SIZE="20" MAXLENGTH="20" ALT="�ִ��������" tag="24X3"> </OBJECT>');</SCRIPT>
										</TD>
									</TR>
									<TR>
										<TD CLASS=TD5 NOWRAP>�ø���</TD>
										<TD CLASS=TD6 NOWRAP>
											<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name=txtRoundQty CLASS = FPDS115 title=FPDOUBLESINGLE SIZE="20" MAXLENGTH="20" ALT="�ø���" tag="24X3"></OBJECT>');</SCRIPT>
										</TD>
									</TR>
								</TABLE>	
							</FIELDSET>			
							</TD>
							<TD WIDTH=65% colspan=2>
							<FIELDSET valign=top>
								<LEGEND>�Ϲ�����</LEGEND>
								<TABLE CLASS="TB2" CELLSPACING=0>
									<TR>
										<TD CLASS=TD5 NOWRAP>ǰ����ȿ��</TD> 
										<TD CLASS=TD6 NOWRAP>
											<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name=txtValidFromDT CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="ǰ����ȿ��" tag="24"> </OBJECT>');</SCRIPT>
										</TD>
										<TD CLASS=TD5 NOWRAP>ǰ���ȿ��</TD>
										<TD CLASS=TD6 NOWRAP>
											<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name=txtValidToDT CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="ǰ���ȿ��" tag="24"> </OBJECT>');</SCRIPT>
										</TD>
									</TR>	
									<TR>
										<TD CLASS=TD5 NOWRAP>���� L/T</TD>
										<TD CLASS=TD6 NOWRAP>
											<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name=txtOrderLtMFG CLASS = FPDS115 title=FPDOUBLESINGLE SIZE="20" MAXLENGTH="20" ALT="���� L/T" tag="24X3"></OBJECT>');</SCRIPT>
										</TD>
										<TD CLASS=TD5 NOWRAP>��������</TD>
										<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtOrderUnitMFG" SIZE=10 tag="24" ALT="��������" ></TD>
									</TR>	
									<TR>
										<TD CLASS=TD5 NOWRAP>����ǰ��ҷ���</TD>
										<TD CLASS=TD6 NOWRAP>
											<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name=txtScrapRateMFG CLASS = FPDS115 title=FPDOUBLESINGLE SIZE="20" MAXLENGTH="20" ALT="����ǰ��ҷ���" tag="24X3"></OBJECT>');</SCRIPT>
										</TD>
										<TD CLASS=TD5 NOWRAP>�����˻�����</TD>
										<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtMPSMgr" SIZE=10 tag="24" ALT="�����˻�����" ></TD>
									</TR>	
									<TR>
										<TD CLASS=TD5 NOWRAP>MRP�����</TD>
										<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtMRPMgr" SIZE=10 tag="24" ALT="MRP�����" ></TD>
										<TD CLASS=TD5 NOWRAP>��������</TD>
										<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtProdMgr" SIZE=10 tag="24" ALT="��������" ></TD>
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
	<TR HEIGHT=20>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_30%>>
				<TR>				
					<TD WIDTH=10>&nbsp;</TD>
					<TD Align=left><BUTTON NAME="btnRelease" ONCLICK="vbscript:ReleaseOrder()" CLASS="CLSMBTN">��������Ȯ��</BUTTON></TD>
					<TD WIDTH=* Align=right><A href="vbscript:JumpOrderRun">������������</A></TD>
					<TD WIDTH=10>&nbsp;</TD>
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
<INPUT TYPE=HIDDEN NAME="hTrackingNo" tag="24"><INPUT TYPE=HIDDEN NAME="hMRPRunNo" tag="24"><INPUT TYPE=HIDDEN NAME="hItemGroupCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hOprCostFlag" tag="24">
<P ID="divTextArea"></P>
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
