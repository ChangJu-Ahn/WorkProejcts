
<%@ LANGUAGE="VBSCRIPT" %>
<!--'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p1101ma1.asp
'*  4. Program Name         : �Ⱓ ���� 
'*  5. Program Desc         :
'*  6. Component List		:
'*  7. Modified date(First) : 2000/04/10
'*  8. Modified date(Last)  : 2000/04/18
'*  9. Modifier (First)     : Mr  Kim Gyoung-Don
'* 10. Modifier (Last)      : Lee Hwa Jung
'* 11. Comment              :
'**********************************************************************************************-->

<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!--'#########################################################################################################
'												1. �� �� �� 
'##########################################################################################################-->
<!--'******************************************  1.1 Inc ����   **********************************************
'	���: Inc. Include
'*********************************************************************************************************-->
<!--
========================================================================================================
=                          1.1.1 Server Side Script
========================================================================================================-->

<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<!--'==========================================  1.1.1 Style Sheet  ======================================
'==========================================================================================================-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		

<!--==========================================  1.1.2 ���� Include   ======================================
==========================================================================================================-->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit                                                             

'******************************************  1.2 Global ����/��� ����  ***********************************
'	1. Constant�� �ݵ�� �빮�� ǥ��.
'**********************************************************************************************************
'==========================================  1.2.1 Global ��� ����  ======================================
'==========================================================================================================
Dim BaseDate
Dim strYear
Dim strMonth
DIm strDay
Dim lgMaxYear
DIm lgMinYear

'========================================================================================================
'=                       1.2.1 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	

BaseDate = "<%=GetSvrDate%>"
Call ExtractDateFrom(BaseDate, parent.gServerDateFormat, parent.gServerDateType, strYear, strMonth, strDay)

lgMaxYear = StrYear + 20
lgMinYear = StrYear - 10

Const BIZ_PGM_BATCH_ID = "p1101mb2.asp"												
Const BIZ_PGM_LOOKUP_ID = "p1101mb4.asp"											

'==========================================  1.2.2 Global ���� ����  =====================================
'	1. ���� ǥ�ؿ� ����. prefix�� g�� �����.
'	2.Array�� ���� ()�� �ݵ�� ����Ͽ� �Ϲ� ������ ������ �� 
'=========================================================================================================

Dim lgNextNo						'��: ȭ���� Single/SingleMulti �ΰ�츸 �ش� 
Dim lgPrevNo						' ""
'==========================================  1.2.3 Global Variable�� ����  ===============================
'=========================================================================================================
'----------------  ���� Global ������ ����  -----------------------------------------------------------
Dim IsOpenPop

'++++++++++++++++  Insert Your Code for Global Variables Assign  ++++++++++++++++++++++++++++++++++++++

'#########################################################################################################
'												2. Function�� 
'
'	���� : �����ڰ� ������ �Լ�, �� Event���� �Լ��� ������ ��� ����� ���� �Լ� �⽽ 
'	�������� ���� ���� : 1. Sub �Ǵ� Function�� ȣ���� �� �ݵ�� Call�� ����.
'		     	     	 2. Sub, Function �̸��� _�� ���� �ʵ��� �Ѵ�. (Event�� �����ϱ� ����) 
'#########################################################################################################
'==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : ���� �ʱ�ȭ(Global ����, �ʱ�ȭ�� �ʿ��� ���� �Ǵ� Flag���� Setting�Ѵ�.)
'=========================================================================================================

Sub InitVariables()

    lgIntFlgMode = Parent.OPMD_CMODE                                               
    lgBlnFlgChgValue = False                                                
    '----------  Coding part  -------------------------------------------------------------
    IsOpenPop = False														

End Sub

'=============================== 2.1.2 LoadInfTB19029() =================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'========================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call LoadInfTB19029A("I", "*", "NOCOOKIE", "MA") %>
End Sub


'******************************************  2.2 ȭ�� �ʱ�ȭ �Լ�  ***************************************
'	���: ȭ���ʱ�ȭ 
'	����: ȭ���ʱ�ȭ, Combo Display, ȭ�� Clear �� ȭ�� �ʱ�ȭ �۾��� �Ѵ�. 
'*********************************************************************************************************
'==========================================  2.2.1 SetDefaultVal()  ======================================
'	Name : SetDefaultVal()
'	Description : ȭ�� �ʱ�ȭ(���� Field�� �� �� ȭ���� �� �� Default���� ������� �ϴ� Field�� Setting)
'=========================================================================================================
Sub SetDefaultVal()
	
	Call ggoOper.SetReqAttr(frm1.txtDayCnt,"Q")
	Call ggoOper.SetReqAttr(frm1.cboWeekDay,"Q")
	
	frm1.txtDayCnt.text = ""
	frm1.txtYear.text = StrYear
	
End Sub

Sub InitComboBox()
	Call SetCombo(frm1.cboPeriodType, "01", "��")								'��: InitCombo ���� �ؾ� �Ǵµ� �ӽ÷� ���� ���� 
    Call SetCombo(frm1.cboPeriodType, "02", "��")
    Call SetCombo(frm1.cboPeriodType, "03", "��")
    Call SetCombo(frm1.cboPeriodType, "04", "��")
    
    Call SetCombo(frm1.cboWeekDay, "2", "��")								'��: InitCombo ���� �ؾ� �Ǵµ� �ӽ÷� ���� ���� 
    Call SetCombo(frm1.cboWeekDay, "3", "ȭ")
    Call SetCombo(frm1.cboWeekDay, "4", "��")
    Call SetCombo(frm1.cboWeekDay, "5", "��")
    Call SetCombo(frm1.cboWeekDay, "6", "��")
    Call SetCombo(frm1.cboWeekDay, "7", "��")
    Call SetCombo(frm1.cboWeekDay, "1", "��")
End Sub

'******************************************  2.3 Operation ó���Լ�  *************************************
'	���: Operation ó���κ� 
'	����: Tabó��, Reference���� ���Ѵ�. 
'*********************************************************************************************************

'******************************************  2.4 POP-UP ó���Լ�  ****************************************
'	���: ���� POP-UP
'   ����: ���� POP-UP�� ���� Open�� Include�Ѵ�. 
'	      �ϳ��� ASP���� Popup�� �ߺ��Ǹ� �ϳ��� ��ũ���� ����ϰ� �������� �������Ͽ� ����Ѵ�.
'*********************************************************************************************************

'========================================== 2.4.2 Open???()  =============================================
'	Name : Open???()
'	Description : �ߺ��Ǿ� �ִ� PopUp�� ������, �����ǰ� �ʿ��� ���� �ݵ�� CommonPopUp.vbs �� 
'				  ManufactPopUp.vbs ���� Copy�Ͽ� �������Ѵ�.
'=========================================================================================================
'++++++++++++++++  Insert Your Code for PopUp(Open)  +++++++++++++++++++++++++++++++++++++++++++++++++++++
'------------------------------------------  OpenCalType()  ----------------------------------------------
'	Name : OpenCalType()
'	Description : Calendar Type Popup
'---------------------------------------------------------------------------------------------------------

Function OpenCalType()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "Į���� Ÿ�� �˾�"			' �˾� ��Ī 
	arrParam(1) = "P_MFG_CALENDAR_TYPE"				' TABLE ��Ī 
	arrParam(2) = Trim(frm1.txtClnrType.Value)		' Code Condition
	arrParam(3) = ""								' Name Cindition
	arrParam(4) = ""								' Where Condition
	arrParam(5) = "Į���� Ÿ��"					' TextBox ��Ī 
	
    arrField(0) = "CAL_TYPE"						' Field��(0)
    arrField(1) = "CAL_TYPE_NM"						' Field��(1)
    
    arrHeader(0) = "Į���� Ÿ��"				' Header��(0)
    arrHeader(1) = "Į���� Ÿ�Ը�"				' Header��(1)

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetCalType(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtClnrType.focus
    
End Function

'==========================================  2.4.3 Set???()  =============================================
'	Name : Set???()
'	Description : PopUp���� ���۵� ���� Ư�� Tag Object�� ���� 
'=========================================================================================================
'++++++++++++++++  Insert Your Code for PopUp(Open)  +++++++++++++++++++++++++++++++++++++++++++++++++++++
'------------------------------------------  SetCalType()  -----------------------------------------------
'	Name : SetCalType()
'	Description : Plant Popup���� Return�Ǵ� �� setting
'---------------------------------------------------------------------------------------------------------
Function SetCalType(byval arrRet)
	frm1.txtClnrType.value = arrRet(0)
	frm1.txtClnrTypeNm.value = arrRet(1) 
End Function

'++++++++++++++++++++++++++++++++++++++++++  2.5 ������ ���� �Լ�  +++++++++++++++++++++++++++++++++++++++
'    ������ ���α׷� ���� �ʿ��� ������ ���� Procedure (Sub, Function, Validation & Calulation ���� �Լ�)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

'#########################################################################################################
'												3. Event�� 
'	���: Event �Լ��� ���� ó�� 
'	����: Windowó��, Singleó��, Gridó�� �۾�.
'         ���⼭ Validation Check, Calcuration �۾��� ������ Event�� �߻�.
'         �� Object������ Grouping�Ѵ�.
'#########################################################################################################
'******************************************  3.1 Window ó��  ********************************************
'	Window�� �߻� �ϴ� ��� Even ó��	
'*********************************************************************************************************
'==========================================  3.1.1 Form_Load()  ==========================================
'	Name : Form_Load()
'	Description : Window On Load(���� Include ���Ͽ� ����)�� �����ʱ�ȭ �� ȭ���ʱ�ȭ�� �ϱ� ���� �Լ��� Call�ϴ� �κ� 
'=========================================================================================================
Sub Form_Load()
    
    Call LoadInfTB19029																'��: Load table , B_numeric_format
	
	Call AppendNumberPlace("6","4","0")
	Call AppendNumberRange("6",lgMinYear,lgMaxYear)
	Call AppendNumberPlace("7","3","0")
	Call AppendNumberRange("7","1","365")
	Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec,FALSE,,ggStrMinPart,ggStrMaxPart)
    Call ggoOper.LockField(Document, "N")											'��: Lock  Suitable  Field
    
    '----------  Coding part  -------------------------------------------------------------
    Call SetToolbar("10000000000011")
    Call InitComboBox
    Call SetDefaultVal
    Call InitVariables
	frm1.txtClnrType.focus 
	Set gActiveElement = document.activeElement 
End Sub

'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================

Sub Form_QueryUnload(Cancel , UnloadMode )
    
End Sub

'==========================================================================================
'   Event Name : cboPeriodType_OnChange()
'   Event Desc :
'==========================================================================================
Sub cboPeriodType_onChange()
	If frm1.cboPeriodType.value = "03" Then
		Call ggoOper.SetReqAttr(frm1.cboWeekDay,"N")
		frm1.cboWeekDay.value = "1"
		
		Call ggoOper.SetReqAttr(frm1.txtDayCnt,"Q")		
		frm1.txtDayCnt.text = ""

	ElseIf frm1.cboPeriodType.value = "04" Then
		Call ggoOper.SetReqAttr(frm1.txtDayCnt,"N")		
		frm1.txtDayCnt.text = "1"
		Call ggoOper.SetReqAttr(frm1.cboWeekDay,"Q")
		frm1.cboWeekDay.value = ""	
	Else
		Call ggoOper.SetReqAttr(frm1.cboWeekDay,"Q")
		frm1.cboWeekDay.value = ""	
		Call ggoOper.SetReqAttr(frm1.txtDayCnt,"Q")		
		frm1.txtDayCnt.text = "" 		
	End If	
End Sub

'#########################################################################################################
'												4. Common Function�� 
'	���: Common Function
'	����: ȯ��ó���Լ�, VAT ó�� �Լ� 
'#########################################################################################################

'#########################################################################################################
'												5. Interface�� 
'	���: Interface
'	����: ������ Toolbar�� ���� ó���� ���Ѵ�. 
'	      Toolbar�� ��ġ������� ����ϴ� ������ �Ѵ�. 
'	<< ���뺯�� ���� �κ� >>
' 	���뺯�� : Global Variables�� �ƴ����� ������ Sub�� Function���� ���� ����ϴ� ������ �������� 
'				�����ϵ��� �Ѵ�.
' 	1. ������Ʈ���� Call�ϴ� ���� 
'    	   ADF (ADS, ADC, ADF�� �״�� ���)
'    	   - ADF�� Set�ϰ� ����� �� �ٷ� Nothing �ϵ��� �Ѵ�.
' 	2. ������Ʈ�ѿ��� Return�� ���� �޴� ���� 
'    		strRetMsg
'#########################################################################################################
'*******************************  5.1 Toolbar(Main)���� ȣ��Ǵ� Function *******************************
'	���� : Fnc�Լ��� ���� �����ϴ� ��� Function
'*********************************************************************************************************

'========================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================

Function FncSave() 
	Dim IntRetCD
		
	lgBlnFlgChgValue = False 
	
	If frm1.txtClnrType.value = "" Then
		frm1.txtClnrTypeNm.value = ""
	End If
	 
	If Not chkField(Document, "2") Then
		Exit Function
	End If
	
	If CInt(frm1.txtYear.Value) < CInt(lgMinYear) Then
		Call DisplayMsgBox("970023","X","�⵵",lgMinYear)
		frm1.txtYear.focus
		Set gActiveElement = document.activeElement 
		Exit Function
	End If
	
	If CInt(frm1.txtYear.Value) > CInt(lgMaxYear) Then
		Call DisplayMsgBox("972004","X","�⵵",lgMaxYear)
		frm1.txtYear.focus
		Set gActiveElement = document.activeElement 
		Exit Function
	End If		
	
	If frm1.cboPeriodType.value = "04" Then
		If frm1.txtDayCnt.Value < 1 Then
			Call DisplayMsgBox("970023","X","�Ⱓ���ϼ�","1")
			frm1.txtDayCnt.focus 
			Set gActiveElement = document.activeElement 
			Exit Function
		End If
	End If
	
	IntRetCD = DisplayMsgBox("900018",Parent.VB_YES_NO,"X","X")
	
	If IntRetCD = vbNo Then
		Exit Function
	End If
		
	Call LookUpLotPeriod	
	
End Function

'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================

Function FncPrint() 
    Call parent.FncPrint()
End Function

'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================

Function FncExcel() 
    Call parent.FncExport(Parent.C_SINGLE)												'��: ȭ�� ���� 
End Function

'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================

Function FncFind() 
    Call parent.FncFind(Parent.C_SINGLE, False)                                         '��:ȭ�� ����, Tab ���� 
End Function

'========================================================================================
' Function Name : FncCancel
' Function Desc : 
'========================================================================================

Function FncCancel() 
    On Error Resume Next
End Function

'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================

Function FncExit()
    FncExit = True
End Function

'========================================================================================
' Function Name : LookUpLotPeriod
' Function Desc : �Ⱓ ���� ��ư�� ������ ������ �Ⱓ�� �ִ� �� ��ȸ�Ѵ�.
'========================================================================================

Function LookUpLotPeriod()

	Dim strVal
	
    LayerShowHide(1)
		
    With frm1
		.txtUpdtUserId.value = Parent.gUsrID
		
		strVal = BIZ_PGM_LOOKUP_ID & "?txtMode=" & Parent.UID_M0001							'��: 
		strVal = strVal & "&txtYear=" & Trim(.txtYear.text)				'��: ��ȸ ���� ����Ÿ 
		strVal = strVal & "&txtClnrType=" & Trim(.txtClnrType.value)				'��: ��ȸ ���� ����Ÿ 
		strVal = strVal & "&txtUpdtUserId=" & Trim(.txtUpdtUserId.value)
	
	End With
	
	Call RunMyBizASP(MyBizASP, strVal)										'��: �����Ͻ� ASP �� ���� 

End Function

Function LotPerdLookUpOk()
	Dim rtnVal
	
	rtnVal = DisplayMsgBox("188100",Parent.VB_YES_NO,"X","X")
	
	If rtnVal = vbYes Then
		Call DbExecute
	Else
		Call LayerShowHide(0)
		Call BtnDisabled(0)
	End If

End Function
'========================================================================================
' Function Name : DbExecute
' Function Desc : ���� �� ���������� ����Ǿ��� ��쿡 
'========================================================================================
Function DbExecute()
    
    With frm1
		.txtMode.value = Parent.UID_M0002											'��: ���� ���� 
		'.txtFlgMode.value = lgIntFlgMode									'��: �ű��Է�/���� ���� 
		.txtInsrtUserId.value  = Parent.gUsrID
		.txtUpdtUserId.value = Parent.gUsrID
	End With
       
    Call ExecMyBizASP(frm1, BIZ_PGM_BATCH_ID)										'��: �����Ͻ� ASP �� ���� 
	
End Function

Function DbExecOk()
	Call DisplayMsgBox("183114","X","X","X")	
End Function


Function LotPerdNo()
	
		Call LayerShowHide(0)
		Call BtnDisabled(0)
	
	
End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>
<BODY TABINDEX="-1" SCROLL="NO">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE CLASS="BatchTB1" CELLSPACING=0 CELLPADDING=0>
	<TR>
		<TD HEIGHT=5>&nbsp;</TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH=100%>
			<TABLE CELLSPACING=0 CELLPADDING=0 WIDTH=100% border=0>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD>
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>�Ⱓ����</font></td>
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
			<TABLE WIDTH="100%" HEIGHT="100%" CELLSPACING=0 CELLPADDING=0>				
				<TR>
					<TD WIDTH=100% HEIGHT=* valign=top>
						<TABLE WIDTH=100% CELLSPACING=0 CELLPADDING=0>
							<TR>
								<TD CLASS=TD5 NOWRAP>&nbsp;</TD>
								<TD CLASS=TD6 NOWRAP>&nbsp;</TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>Į���� Ÿ��</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtClnrType" SIZE=5 MAXLENGTH=2 tag="22XXXU" ALT="Į���� Ÿ��"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCalType" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenCalType()">&nbsp;<INPUT TYPE=TEXT NAME="txtClnrTypeNm" SIZE=30 MAXLENGTH=30 tag="24" ALT="Į���� Ÿ�Ը�"></TD>							</TR>	
							<TR>
								<TD CLASS=TD5 NOWRAP>&nbsp;</TD>
								<TD CLASS=TD6 NOWRAP>&nbsp;</TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>�⵵</TD>
								<TD CLASS=TD6 NOWRAP>	
									<script language =javascript src='./js/p1101ma1_I169267725_txtYear.js'></script>								
								</TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>&nbsp;</TD>
								<TD CLASS=TD6 NOWRAP>&nbsp;</TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>�Ⱓ����</TD>
								<TD CLASS=TD6 NOWRAP><SELECT NAME="cboPeriodType" ALT="�Ⱓ����" STYLE="Width: 98px;" tag="22"></SELECT></TD>	
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>&nbsp;</TD>
								<TD CLASS=TD6 NOWRAP>&nbsp;</TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>�Ⱓ ���ۿ���</TD>
								<TD CLASS=TD6 NOWRAP><SELECT NAME="cboWeekDay" ALT="�Ⱓ ���ۿ���" STYLE="Width: 98px;" tag="24"><OPTION VALUE=""></OPTION></SELECT></TD>	
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>&nbsp;</TD>
								<TD CLASS=TD6 NOWRAP>&nbsp;</TD>
							</TR>
    						<TR>
								<TD CLASS=TD5 NOWRAP>�Ⱓ���ϼ�</TD>
								<TD CLASS=TD6 NOWRAP>
									<script language =javascript src='./js/p1101ma1_I603434699_txtDayCnt.js'></script>								
								</TD>
							</TR> 							
							<TR>
								<TD CLASS=TD5 NOWRAP>&nbsp;</TD>
								<TD CLASS=TD6 NOWRAP>&nbsp;</TD>
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
					<TD><BUTTON NAME="btnExec" CLASS="CLSMBTN" Flag=1 ONCLICK=FncSave>����</BUTTON></TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC= "../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 tabindex="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24"><INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
