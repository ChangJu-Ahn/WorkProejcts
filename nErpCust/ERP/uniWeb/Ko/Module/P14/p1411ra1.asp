
<%@ LANGUAGE="VBSCRIPT" %>
<%'******************************************************************************************************
'*  1. Module Name          : Production																*
'*  2. Function Name        :																			*
'*  3. Program ID           : p1411ra1.asp																*
'*  4. Program Name         :																			*
'*  5. Program Desc         : Change History Detail																*
'*  6. Comproxy List        :																			*
'*  7. Modified date(First) : 2002/04/09																*
'*  8. Modified date(Last)  :																			*
'*  9. Modifier (First)     : Park Kye Jin																*
'* 10. Modifier (Last)      : 																			*
'* 11. Comment              :																			*
'******************************************************************************************************%>

<HTML>
<HEAD>
<!--####################################################################################################
'#						1. �� �� ��																		#
'#####################################################################################################-->

<!--********************************************  1.1 Inc ����  *****************************************
'*	Description : Inc. Include																			*
'*****************************************************************************************************-->
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->
<!--============================================  1.1.1 Style Sheet  ====================================
'=====================================================================================================-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">
<!--============================================  1.1.2 ���� Include  ===================================
'=====================================================================================================-->
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliPAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliPAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliPAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "JavaScript" SRC = "../../inc/incImage.js"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript">

Option Explicit

'********************************************  1.2 Global ����/��� ����  *******************************
'*	Description : 1. Constant�� �ݵ�� �빮�� ǥ��														*
'********************************************************************************************************
			'��: �����Ͻ� ���� ASP�� 
'============================================  1.2.1 Global ��� ����  ==================================
'========================================================================================================
	
'==========================================  1.2.0 Common variables =====================================
'	1. Common variables
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
	
'============================================  1.2.2 Global ���� ����  ==================================
'========================================================================================================
Dim arrReturn
Dim IsOpenPop
Dim arrParent

'------ Set Parameters from Parent ASP ------
ArrParent			= window.dialogArguments
Set PopupParent		= ArrParent(0)

top.document.title = PopupParent.gActivePRAspName
'============================================  1.2.3 Global Variable�� ����  ============================
'========================================================================================================
'----------------  ���� Global ������ ����  -------------------------------------------------------

'++++++++++++++++  Insert Your Code for Global Variables Assign  ++++++++++++++++++++++++++++++++++

'########################################################################################################
'#						2. Function ��																	#
'#																										#
'#	���� : �����ڰ� ������ �Լ�, �� Event���� �Լ��� ������ ��� ����� ���� �Լ� ���					#
'#	�������� ���� ���� : 1. Sub �Ǵ� Function�� ȣ���� �� �ݵ�� Call�� ����.							#
'#						 2. Sub, Function �̸��� _�� ���� �ʵ��� �Ѵ�. (Event�� �����ϱ� ����)			#
'########################################################################################################
'*******************************************  2.1 ���� �ʱ�ȭ �Լ�  *************************************
'*	���: �����ʱ�ȭ																					*
'*	Description : Global���� ó��, �����ʱ�ȭ ���� �۾��� �Ѵ�.											*
'********************************************************************************************************

'==========================================  2.1.1 InitVariables()  =====================================
'=	Name : InitVariables()																				=
'=	Description : ���� �ʱ�ȭ(Global ����, �ʱ�ȭ�� �ʿ��� ���� �Ǵ� Flag���� Setting�Ѵ�.)				=
'========================================================================================================
Function InitVariables()

End Function

'=======================================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'=======================================================================================================%>
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "P", "NOCOOKIE","RA") %>
	<% Call loadBNumericFormatA("I", "P", "NOCOOKIE","RA") %>
End Sub

'==========================================   2.1.2 InitSetting()   =====================================
'=	Name : InitSetting()																				=
'=	Description : Passed Parameter�� Variable�� Setting�Ѵ�.											=
'========================================================================================================
Function InitSetting()

	Dim ArgArray						<%'Arguments�� �Ѱܹ��� Array%>
	
	ArgArray  = ArrParent(1)
	
	frm1.txtChangedField.value = ArgArray(0)
	frm1.txtPrntItemCd.value = ArgArray(1)	'item cd
	frm1.txtPrntItemNm.value = ArgArray(2)	
	frm1.txtChildItemCd.value = ArgArray(3)	
	frm1.txtChildItemNm.value = ArgArray(4)	
	frm1.txtBomSeq.value = ArgArray(5)	
	frm1.txtChangeCode.value = ArgArray(6)	
	frm1.txtInsertDt.value = ArgArray(25)	
	frm1.txtInsertUserId.value = ArgArray(26)
	
	If 	ArgArray(6) = "Add" then
		frm1.txtChildItemQty.value = ArgArray(7)	
		frm1.txtChildUnit.value = ArgArray(8)	
		frm1.txtPrntItemQty.value = ArgArray(9)	
		frm1.txtPrntUnit.value = ArgArray(10)	
		frm1.txtSafetyLt.value = ArgArray(11)	
		frm1.txtLossRate.value = ArgArray(12)
		frm1.txtSupplyFlg.value = ArgArray(13)	
		frm1.txtValidFromDt.Text = ArgArray(14)
		frm1.txtValidToDt.Text = ArgArray(15)	
		frm1.txtChildItemQty1.value = ""	
		frm1.txtChildUnit1.value = ""
		frm1.txtPrntItemQty1.value = ""	
		frm1.txtPrntUnit1.value = ""
		frm1.txtSafetyLt1.value = ""
		frm1.txtLossRate1.value = ""
		frm1.txtSupplyFlg1.value = ""
		frm1.txtValidFromDt1.Text = ""
		frm1.txtValidToDt1.Text = ""
	End If	
	
	If 	ArgArray(6) = "Delete" then
		frm1.txtChildItemQty.value = ""
		frm1.txtChildUnit.value = ""	
		frm1.txtPrntItemQty.value = ""
		frm1.txtPrntUnit.value = ""
		frm1.txtSafetyLt.value = ""
		frm1.txtLossRate.value = ""
		frm1.txtSupplyFlg.value = ""
		frm1.txtValidFromDt.Text = ""
		frm1.txtValidToDt.Text = ""
		frm1.txtChildItemQty1.value = ArgArray(16)	
		frm1.txtChildUnit1.value = ArgArray(17)	
		frm1.txtPrntItemQty1.value = ArgArray(18)	
		frm1.txtPrntUnit1.value = ArgArray(19)	
		frm1.txtSafetyLt1.value = ArgArray(20)	
		frm1.txtLossRate1.value = ArgArray(21)
		frm1.txtSupplyFlg1.value = ArgArray(22)	
		frm1.txtValidFromDt1.Text = ArgArray(23)
		frm1.txtValidToDt1.Text = ArgArray(24)
	End If		
	
	If 	ArgArray(6) = "Change" then
		frm1.txtChildItemQty.value = ArgArray(7)	
		frm1.txtChildUnit.value = ArgArray(8)	
		frm1.txtPrntItemQty.value = ArgArray(9)	
		frm1.txtPrntUnit.value = ArgArray(10)	
		frm1.txtSafetyLt.value = ArgArray(11)	
		frm1.txtLossRate.value = ArgArray(12)
		frm1.txtSupplyFlg.value = ArgArray(13)	
		frm1.txtValidFromDt.Text = ArgArray(14)
		frm1.txtValidToDt.Text = ArgArray(15)	
		frm1.txtChildItemQty1.value = ArgArray(16)	
		frm1.txtChildUnit1.value = ArgArray(17)	
		frm1.txtPrntItemQty1.value = ArgArray(18)	
		frm1.txtPrntUnit1.value = ArgArray(19)	
		frm1.txtSafetyLt1.value = ArgArray(20)	
		frm1.txtLossRate1.value = ArgArray(21)
		frm1.txtSupplyFlg1.value = ArgArray(22)	
		frm1.txtValidFromDt1.Text = ArgArray(23)
		frm1.txtValidToDt1.Text = ArgArray(24)
	End If
	
End Function

'*******************************************  2.2 ȭ�� �ʱ�ȭ �Լ�  *************************************
'*	���: ȭ���ʱ�ȭ																					*
'*	Description : ȭ���ʱ�ȭ, Combo Display, ȭ�� Clear �� ȭ�� �ʱ�ȭ �۾��� �Ѵ�.						*
'********************************************************************************************************

'++++++++++++++++++++++++++++++++++++++++++  2.3 ������ ���� �Լ�  ++++++++++++++++++++++++++++++++++++++
'+	������ ���� Function, Procedure																		+
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

'===========================================  2.3.1 OkClick()  ==========================================
'=	Name : OkClick()																					=
'=	Description : Return Array to Opener Window when OK button click									=
'========================================================================================================

'=========================================  2.3.2 CancelClick()  ========================================
'=	Name : CancelClick()																				=
'=	Description : Return Array to Opener Window for Cancel button click 								=
'========================================================================================================
Function CancelClick()
	self.close()
End Function
'=========================================  2.3.3 Mouse Pointer ó�� �Լ� ===============================
'========================================================================================================
Function MousePointer(pstr1)
      Select case UCase(pstr1)
            case "PON"
				window.document.search.style.cursor = "wait"
            case "POFF"
				window.document.search.style.cursor = ""
      End Select
End Function

'*******************************************  2.4 POP-UP ó���Լ�  **************************************
'*	���: POP-UP																						*
'*	Description : POP-UP Call�ϴ� �Լ� �� Return Value setting ó��										*
'********************************************************************************************************

'===========================================  2.4.1 POP-UP Open �Լ�()  =================================
'=	Name : Open???()																					=
'=	Description : POP-UP Open																			=
'========================================================================================================



'=======================================  2.4.2 POP-UP Return�� ���� �Լ�  ==============================
'=	Name : Set???()																						=
'=	Description : Reference �� POP-UP�� Return���� �޴� �κ�											=
'========================================================================================================



'++++++++++++++++++++++++++++++++++++++++++  2.5 ������ ���� �Լ�  ++++++++++++++++++++++++++++++++++++++
'+	���� ���α׷����� �ʿ��� ������ ���� Procedure(Sub, Function, Validation & Calulation ���� �Լ�)	+
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

'########################################################################################################
'#						3. Event ��																		#
'#	���: Event �Լ��� ���� ó��																		#
'#	����: Windowó��, Singleó��, Gridó�� �۾�.														#
'#		  ���⼭ Validation Check, Calcuration �۾��� ������ Event�� �߻�.								#
'#		  �� Object������ Grouping�Ѵ�.																	#
'########################################################################################################
'********************************************  3.1 Windowó��  ******************************************
'*	Window�� �߻� �ϴ� ��� Even ó��																	*
'********************************************************************************************************
'=========================================  3.1.1 Form_Load()  ==========================================
'=	Name : Form_Load()																					=
'=	Description : Window Load�� �����ʱ�ȭ �� ȭ���ʱ�ȭ�� �ϱ� ���� �Լ��� Call�ϴ� �κ�				=
'========================================================================================================
Sub Form_Load()
	Call AppendNumberPlace("6", "3", "0")
	Call AppendNumberPlace("7", "2", "2")
	Call AppendNumberPlace("8", "11", "6")
	Call ggoOper.FormatField(Document, "A", ggStrIntegeralPart, ggStrDeciPointPart, PopupParent.gDateFormat, PopupParent.gComNum1000, PopupParent.gComNumDec)
	
	Call MM_preloadImages("../../../CShared/image/Query.gif","../../../CShared/image/OK.gif","../../../CShared/image/Cancel.gif")
	Call LoadInfTB19029											'��: Load table , B_numeric_format		
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec)    		
	Call ggoOper.LockField(Document, "N")						<% '��: Lock  Suitable  Field %>
	Call InitVariables											'��: Initializes local global variables
	Call InitSetting()
End Sub

'*********************************************  3.2 Tag ó��  *******************************************
'*	Document�� TAG���� �߻� �ϴ� Event ó��																*
'*	Event�� ��� �Ʒ��� ����� Event�̿��� ����� �����ϸ� �ʿ�� �߰� �����ϳ�							*
'*	Event�� �浹�� ����Ͽ� �ۼ��Ѵ�.																	*
'********************************************************************************************************

'########################################################################################################
'#					     4. Common Function��															#
'########################################################################################################
'########################################################################################################
'#						5. Interface ��																	#
'########################################################################################################
'********************************************  5.1 DbQuery()  *******************************************
' Function Name : DbQuery																				*
' Function Desc : This function is data query and display												*
'********************************************************************************************************

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>
<%
'########################################################################################################
'#						6. TAG ��																		#
'########################################################################################################
%>
<BODY SCROLL=NO TABINDEX="-1">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE CELLSPACING=0 CLASS="basicTB">
	<TR>
		<TD HEIGHT=40>
			<FIELDSET CLASS="CLSFLD">
				<TABLE WIDTH=100% CELLSPACING=0>					
					<TR>
						<TD CLASS=TD5 NOWRAP>��ǰ��</TD>
						<TD CLASS=TD6 NOWRAP colspan=3><INPUT TYPE=TEXT NAME="txtPrntItemCd" SIZE=18 MAXLENGTH=18 tag="14xxxU" ALT="��ǰ��">&nbsp;<INPUT TYPE=TEXT NAME="txtPrntItemNm" SIZE=40 MAXLENGTH=50 tag="14" ALT="��ǰ��"></TD>
					</TR>
					<TR>
						<TD CLASS=TD5 NOWRAP>��ǰ��</TD>
						<TD CLASS=TD6 NOWRAP colspan=3><INPUT TYPE=TEXT NAME="txtChildItemCd" SIZE=18 MAXLENGTH=18 tag="14xxxU" ALT="��ǰ��">&nbsp;<INPUT TYPE=TEXT NAME="txtChildItemNm" SIZE=40 MAXLENGTH=50 tag="14" ALT="��ǰ��"></TD>
					</TR>
					<TR>
						<TD CLASS=TD5 NOWRAP>��ǰ�����</TD>
						<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT title=FPDOUBLESINGLE CLASS=FPDS140 name=txtBomSeq CLASSID=<%=gCLSIDFPDS%> SIZE="6" MAXLENGTH="6" ALT="��ǰ�����" tag="24X6Z"> </OBJECT>');</SCRIPT></TD>
						<TD CLASS=TD5 NOWRAP>���躯�汸��</TD>
						<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtChangeCode" SIZE=10 MAXLENGTH=10 tag="14" ALT="���躯�汸��"></TD>
					</TR>
					<TR>
						<TD CLASS=TD5 NOWRAP>���躯����</TD>
						<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtInsertDt" SIZE=25 MAXLENGTH=25 tag="14xxxU" ALT="���躯����"></TD>
						<TD CLASS=TD5 NOWRAP>���躯����</TD>
						<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtInsertUserId" SIZE=13 MAXLENGTH=13 tag="14xxxU" ALT="���躯����"></TD>
					</TR>
					<TR>
						<TD CLASS=TD5 NOWRAP>������ʵ�</TD>
						<TD CLASS=TD6 NOWRAP colspan=3><INPUT TYPE=TEXT NAME="txtChangedField" SIZE=90 MAXLENGTH=120 tag="14xxxU" ALT="������ʵ�"></TD>
					</TR>											
				</TABLE>
			</FIELDSET>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=* valign=top>
			<TABLE CLASS="TB2" CELLSPACING=0>
				<TR>
					<TD WIDTH=50%  valign=top>
						<FIELDSET>
							<LEGEND>������</LEGEND>
							<TABLE CLASS="TB2" CELLSPACING=0>
								<TR> 
									<TD CLASS=TD5 NOWRAP>��ǰ����ؼ�</TD>
									<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT title=FPDOUBLESINGLE CLASS=FPDS140 name=txtChildItemQty CLASSID=<%=gCLSIDFPDS%> SIZE="15" MAXLENGTH="15" ALT="��ǰ����ؼ�" tag="20X8Z"> </OBJECT>');</SCRIPT></TD>
								<TR>
									<TD CLASS=TD5 NOWRAP>����</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtChildUnit" SIZE=8 MAXLENGTH=3 tag="24"  ALT="��ǰ����ش���"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>��ǰ����ؼ�</TD>
									<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT title=FPDOUBLESINGLE CLASS=FPDS140 name=txtPrntItemQty CLASSID=<%=gCLSIDFPDS%> SIZE="15" MAXLENGTH="15" ALT="��ǰ����ؼ�" tag="20X8Z"> </OBJECT>');</SCRIPT></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>����</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtPrntUnit" SIZE=8 MAXLENGTH=3 tag="24"  ALT="��ǰ����ش���"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>����L/T</TD>
									<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT title=FPDOUBLESINGLE CLASS=FPDS140 name=txtSafetyLt CLASSID=<%=gCLSIDFPDS%> SIZE="15" MAXLENGTH="15" ALT="����L/T" tag="20X6Z"> </OBJECT>');</SCRIPT></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>Loss��</TD>
									<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT title=FPDOUBLESINGLE CLASS=FPDS140 name=txtLossRate CLASSID=<%=gCLSIDFPDS%> SIZE="15" MAXLENGTH="15" ALT="Loss��" tag="20X7Z"> </OBJECT>');</SCRIPT></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>�����󱸺�</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtSupplyFlg" ALT="�����󱸺�" SIZE=8 MAXLENGTH=3 tag="24"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>������</TD>
									<TD CLASS=TD6 NOWRAP>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name=txtValidFromDt CLASS=FPDTYYYYMMDD title=FPDATETIME tag="24" ALT="������"></OBJECT>');</SCRIPT></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>������</TD>
									<TD CLASS=TD6 NOWRAP>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name=txtValidToDt CLASS=FPDTYYYYMMDD title=FPDATETIME tag="24" ALT="������"></OBJECT>');</SCRIPT></TD>
								</TR>											
							</TABLE>
						</FIELDSET>
					</TD>
					<TD WIDTH=50% valign=top>
						<FIELDSET>
							<LEGEND>������</LEGEND>
							<TABLE CLASS="TB2" CELLSPACING=0>
								<TR> 
									<TD CLASS=TD5 NOWRAP>��ǰ����ؼ�</TD>
									<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT title=FPDOUBLESINGLE CLASS=FPDS140 name=txtChildItemQty1 CLASSID=<%=gCLSIDFPDS%> SIZE="15" MAXLENGTH="15" ALT="��ǰ����ؼ�" tag="20X8Z"> </OBJECT>');</SCRIPT></TD>														
								<TR>
									<TD CLASS=TD5 NOWRAP>����</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtChildUnit1" SIZE=8 MAXLENGTH=3 tag="24"  ALT="��ǰ����ش���"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>��ǰ����ؼ�</TD>
									<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT title=FPDOUBLESINGLE CLASS=FPDS140 name=txtPrntItemQty1 CLASSID=<%=gCLSIDFPDS%> SIZE="15" MAXLENGTH="15" ALT="��ǰ����ؼ�" tag="20X8Z" > </OBJECT>');</SCRIPT></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>����</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtPrntUnit1" SIZE=8 MAXLENGTH=3 tag="24"  ALT="��ǰ����ش���"></TD>
								</TR>

								<TR>
									<TD CLASS=TD5 NOWRAP>����L/T</TD>
									<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT title=FPDOUBLESINGLE CLASS=FPDS140 name=txtSafetyLt1 CLASSID=<%=gCLSIDFPDS%> SIZE="15" MAXLENGTH="15" ALT="����L/T" tag="20X6Z"> </OBJECT>');</SCRIPT></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>Loss��</TD>
									<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT title=FPDOUBLESINGLE CLASS=FPDS140 name=txtLossRate1 CLASSID=<%=gCLSIDFPDS%> SIZE="15" MAXLENGTH="15" ALT="Loss��" tag="20X7Z"> </OBJECT>');</SCRIPT></TD>					
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>�����󱸺�</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtSupplyFlg1" ALT="�����󱸺�" SIZE=8 MAXLENGTH=3 tag="24"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>������</TD>
									<TD CLASS=TD6 NOWRAP>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name=txtValidFromDt1 CLASS=FPDTYYYYMMDD title=FPDATETIME tag="24" ALT="������"></OBJECT>');</SCRIPT></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>������</TD>
									<TD CLASS=TD6 NOWRAP>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name=txtValidToDt1 CLASS=FPDTYYYYMMDD title=FPDATETIME tag="24" ALT="������"></OBJECT>');</SCRIPT></TD>
								</TR>											
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>	
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD HEIGHT=*>
			<TABLE CLASS="basicTB" CELLSPACING=0>
				<TR>
					<TD WIDTH=70% NOWRAP>&nbsp;&nbsp;
					<TD WIDTH=30% ALIGN=RIGHT>
					<IMG SRC="../../../CShared/image/cancel_d.gif" Style="CURSOR: hand" ALT="CANCEL" NAME="pop2" ONCLICK="CancelClick()" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)"></IMG>&nbsp;&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC= "../../blank.htm " WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX = "-1"></IFRAME>
		</TD>
	</TR>
</TABLE>

</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
