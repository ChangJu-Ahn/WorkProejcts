
<%@ LANGUAGE="VBSCRIPT" %>

<!--'**********************************************************************************************
'*  1. Module Name          : Finance
'*  2. Function Name        : Finance Management
'*  3. Program ID           : f5103ba1.asp
'*  4. Program Name         : ������ǥ��ȣ�ڵ�ä�� 
'*  5. Program Desc         :
'*  6. Comproxy List        : +B19029LookupNumericFormat
'*  7. Modified date(First) : 2000/09/25
'*  8. Modified date(Last)  : 2002/08/19
'*  9. Modifier (First)     : hersheys
'* 10. Modifier (Last)      : Shin Myoung_Ha
'* 11. Comment              :
'* 12. History				: 1. ���翵���� ���̴¹��� ���� - 2002/08/09
'*							  2. ��¥ ���� OCX TEXT �� VALUE �߸��� ��� ���� - 2002/08/09 
'*							  3. ������ȣ �Է¶��� "\"���ڿ� ���� �Է½� ���ڷ� �ν�("\"�� �Է��ϸ� ���ڷ� �ν���) - 2002/08/19
'*                            
'********************************************************************************************** -->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!--
'##########################################################################################################
'												1. �� �� �� 
'##########################################################################################################

'******************************************  1.1 Inc ����   **********************************************
'	���: Inc. Include
'*********************************************************************************************************  -->
<!-- #Include file="../../inc/IncSvrCcm.inc"  -->				<!--: ȭ��ó��ASP���� �����۾��� �ʿ��� ���  -->
<!-- #Include file="../../inc/IncSvrHTML.inc"  -->

<!--
'===============================================  1.1.1 Style Sheet  ======================================
'========================================================================================================== -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		<!-- '��: �ش� ��ġ�� ���� �޶���, ��� ���  -->

<!--
'=============================================  1.1.2 ���� Include   ======================================
'========================================================================================================== -->
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>

<SCRIPT LANGUAGE="VBScript">

Option Explicit                                                              '��: indicates that All variables must be declared in advance 

'##########################################################################################################
'												1. �� �� �� 
'##########################################################################################################


'******************************************  1.1 Inc ����   **********************************************
'	���: Inc. Include
'********************************************************************************************************* 
<!-- #Include file="../../inc/lgvariables.inc" -->

'******************************************  1.2 Global ����/��� ����  ***********************************
'	1. Constant�� �ݵ�� �빮�� ǥ��.
'**********************************************************************************************************

'==========================================  1.2.1 Global ��� ����  ======================================
'==========================================================================================================

Const BIZ_PGM_ID = "f5103bb1.asp"  

'==========================================  1.2.3 Global Variable�� ����  ===============================
'========================================================================================================= 


'-------------------  ���� Global ������ ����  ----------------------------------------------------------- 


'+++++++++++++++++++  Insert Your Code for Global Variables Assign  ++++++++++++++++++++++++++++++++++++++ 
Dim IsOpenPop          

'#########################################################################################################
'												2. Function�� 
'
'	���� : �����ڰ� ������ �Լ�, �� Event���� �Լ��� ������ ��� ����� ���� �Լ� �⽽ 
'	�������� ���� ���� : 1. Sub �Ǵ� Function�� ȣ���� �� �ݵ�� Call�� ����.
'		     	     	 2. Sub, Function �̸��� _�� ���� �ʵ��� �Ѵ�. (Event�� �����ϱ� ����) 
'######################################################################################################### 

'***************************************  2.1 Pop-Up �Լ�   **********************************************
'	���: Pop-Up 
'********************************************************************************************************* 


'========================================  2.2.1 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : ȭ�� �ʱ�ȭ(���� Field�� �� �� ȭ���� �� �� Default���� ������� �ϴ� Field�� Setting)
'========================================================================================================= 
Sub SetDefaultVal()
'	frm1.txtIssueDt.text = UNIDateClientFormat("<%=GetSvrDate%>")
	frm1.txtIssueDt.text = UniConvDateAToB("<%=GetSvrDate%>",Parent.gServerDateFormat,Parent.gDateFormat) 
End Sub

 
'------------------------------------------  OpenCalType()  ----------------------------------------------
'	Name :InitComboBox()
'	Description : 
'--------------------------------------------------------------------------------------------------------- 
Sub InitComboBox()
	Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("F1001", "''", "S") & "  ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboNoteKind ,lgF0  ,lgF1  ,Chr(11))
End Sub


'=======================================================================================================
'	Name : OpenBankCd()
'	Description : Bank Code PopUp
'=======================================================================================================
Function OpenBankCd(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "���� �˾�"				'�˾� ��Ī 
	arrParam(1) = "B_BANK"						'TABLE ��Ī 
	arrParam(2) = strCode						'Code Condition
	arrParam(3) = ""							'Name Cindition
	arrParam(4) = ""							'Where Condition
	arrParam(5) = "�����ڵ�"			
	
    arrField(0) = "BANK_CD"					    'Field��(0)
    arrField(1) = "BANK_NM"			    'Field��(1)
    
    arrHeader(0) = "�����ڵ�"				'Header��(0)
    arrHeader(1) = "�����"					'Header��(1)
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=430px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetBankCd(arrRet, iWhere)
	End If	

End Function

'======================================================================================================
'	Name : fncnew()
'	Description : BankCd Popup���� Return�Ǵ� �� setting
'=======================================================================================================
Sub fncnew()
	
	Call ggoOper.ClearField(Document, "1")                                      '��: Clear Condition Field
    Call ggoOper.ClearField(Document, "2")                                      '��: Clear Contents  Field
    Call ggoOper.LockField(Document, "N")                                       '��: Lock  Suitable  Field    
    
    Call SetDefaultVal()    

End SUb 

'======================================================================================================
'	Name : cboNoteKind_Onchange()
'	Description : BankCd Popup���� Return�Ǵ� �� setting
'=======================================================================================================
Sub cboNoteKind_Onchange()
	'Call ggoOper.ClearField(Document, "1")                                      '��: Clear Condition Field
    'Call ggoOper.ClearField(Document, "2")                                      '��: Clear Contents  Field
    'Call ggoOper.LockField(Document, "N")                                       '��: Lock  Suitable  Field    
End SUb 
'======================================================================================================
'	Name : SetBankCd()
'	Description : BankCd Popup���� Return�Ǵ� �� setting
'=======================================================================================================
Function SetBankCd(Byval arrRet, Byval iWhere)
	
	With frm1
	   	If iWhere = 0 Then
    		.txtBankCd.value = arrRet(0)
    		.txtBankNm.value = arrRet(1)
    	End If
	
	End With
	
End Function

'===========================================================
'�۾����� 
'===========================================================
Function FnButtonExec()
	Dim strVal
	Dim intRetCD
	Dim intRtn

	'-----------------------
	'Check content area
	'-----------------------
	If Not ChkField(Document, "1") Then                             '��: Check contents area
		Exit Function
	End If
		
	With frm1		
		.txtFromNo.value = Trim(.txtFromNo.value)		
		.txtToNo.value = .txtToNo.value		
		
		intRtn = instr(1,.txtFromNo.value ,"\")
				
		if intrtn > 0 then
			Call DisplayMsgBox("700119", "X", .txtFromNo.Alt, "X")	'���ڸ� �Է��ϼ���.
			.txtFromNo.focus
			Set gActiveElement = document.activeElement
			
			Exit Function
		end if
		
		intRtn = instr(1,.txtToNo.value ,"\")
				
		if intrtn > 0 then
			Call DisplayMsgBox("700119", "X", .txtToNo.Alt, "X")	'���ڸ� �Է��ϼ���.
			.txtToNo.focus
			Set gActiveElement = document.activeElement
			
			Exit Function
		end if		
		
		If Not IsNumeric(.txtFromNo.value) Then		
			Call DisplayMsgBox("700119", "X", .txtFromNo.Alt, "X")	'���ڸ� �Է��ϼ���.
			.txtFromNo.focus
			Set gActiveElement = document.activeElement
			
			Exit Function
		End If
	
		If Not IsNumeric(.txtToNo.value) Then
			Call DisplayMsgBox("700119", "X", .txtToNo.Alt, "X")	'���ڸ� �Է��ϼ���.
			.txtToNo.focus
			Set gActiveElement = document.activeElement			
			Exit Function
		End If

		If Int(.txtFromNo.value) > Int(.txtToNo.value) Then
			Call DisplayMsgBox("970025", "X", .txtFromNo.Alt, .txtToNo.Alt)	'From�� To���� �۰ų� ���ƾ� �մϴ�.
			.txtFromNo.focus
			Set gActiveElement = document.activeElement			
			Exit Function
		End If
	End With
	
	IntRetCD = DisplayMsgBox("900018",Parent.VB_YES_NO, "X", "X")	'�۾��� �����Ͻðڽ��ϱ�?
	If IntRetCD = vbNo Then		
		Exit Function
	End If
	
	Call LayerShowHide(1) 
    
	strVal = BIZ_PGM_ID & "?txtMode		=" & Parent.UID_M0002							'��: �����Ͻ� ó�� ASP�� ���� 
	strVal = strVal & "&cboNoteKind		=" & UCase(Trim(frm1.cboNoteKind.value))
	strVal = strVal & "&txtBankCd		=" & UCase(Trim(frm1.txtBankCd.value))
	strVal = strVal & "&txtNoteNo		=" & UCase(Trim(frm1.txtNoteNo.value))
	strVal = strVal & "&txtFromNo		=" & Trim(frm1.txtFromNo.value)
	strVal = strVal & "&txtToNo			=" & Trim(frm1.txtToNo.value)
	strVal = strVal & "&txtIssueDt		=" & UNIConvDate(Trim(frm1.txtIssueDt.Text))
	strVal = strVal & "&intLenToNO		=" & Len(Trim(frm1.txtToNo.value))
	
	Call RunMyBizASP(MyBizASP, strVal)			'��: �����Ͻ� ASP �� ���� 
		    
End Function


'##########################################################################################################
'												3. Event�� 
'	���: Event �Լ��� ���� ó�� 
'	����: Windowó��, Singleó��, Gridó�� �۾�.
'         ���⼭ Validation Check, Calcuration �۾��� ������ Event�� �߻�.
'         �� Object������ Grouping�Ѵ�.
'##########################################################################################################

'******************************************  3.1 Window ó��  *********************************************
'	Window�� �߻� �ϴ� ��� Even ó��	
'**********************************************************************************************************

'==========================================  3.1.1 Form_Load()  ===========================================
'	Name : Form_Load()
'	Description : Window On Load(���� Include ���Ͽ� ����)�� �����ʱ�ȭ �� ȭ���ʱ�ȭ�� �ϱ� ���� �Լ��� Call�ϴ� �κ� 
'========================================================================================================= 
Sub Form_Load()

   ' Call LoadInfTB19029

    Call ggoOper.FormatField(Document, "1", ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")											'��: Lock  Suitable  Field
    
    '----------  Coding part  -------------------------------------------------------------
    Call SetToolBar("10100000000011")
    Call SetDefaultVal
    Call InitComboBox
    
	frm1.cboNoteKind.Focus
End Sub

'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================

Sub LoadInfTB19029()

   ' Call parent.LoadBAInf("I","*","X","X","YMD",aGetSvrDate)
End Sub


Sub Form_QueryUnload(Cancel , UnloadMode )

End Sub


'=======================================================================================================
'   Event Name : txtIssueDt_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtIssueDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtIssueDt.Action = 7
    End If
End Sub


'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================

Function FncPrint() 
    On Error Resume Next                                                    '��: Protect system from crashing
    
    parent.FncPrint()
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


'======================================================================================================
' Function Name : FncExit
' Function Desc : 
'=======================================================================================================
Function FncExit()
	FncExit = True
End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->	
</HEAD>

<BODY SCROLL="NO">
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
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>������ǥ��ȣ�ڵ�ä��</font></td>
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
					<TD WIDTH=100% HEIGHT=100% valign=top>
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD CLASS="TD5">������ǥ����</TD>
								<TD CLASS="TD6" COLSPAN = 3><SELECT NAME="cboNoteKind" tag="22" STYLE="WIDTH: 105px;" ALT="������ǥ����"></TD>
							</TR>
							<TR>
								<TD CLASS="TD5">��������</TD>
								<TD CLASS="TD6"><INPUT TYPE=TEXT NAME="txtBankCd" SIZE=11 MAXLENGTH=10 tag="12XXXU" ALT="��������" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBankCd" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenBankCd(frm1.txtBankCd.value,0)">&nbsp;<INPUT TYPE=TEXT NAME="txtBankNm" SIZE=25 tag="14"></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>������ǥ��ȣ</TD>
								<TD CLASS="TD6"><INPUT CLASS="txtNoteNo" TYPE=TEXT NAME="txtNoteNo" SIZE=20 MAXLENGTH=20 tag="12XXXU" ALT="������ǥ��ȣ">&nbsp;<INPUT CLASS="txtFromNo" TYPE=TEXT NAME="txtFromNo" SIZE=10 MAXLENGTH=9 tag="12" ALT="���۹�ȣ"> ~ <INPUT CLASS="txtToNo" TYPE=TEXT NAME="txtToNo" SIZE=10 MAXLENGTH=9 tag="12" ALT="�����ȣ"></TD>
							</TR>
							<TR>
								<TD CLASS="TD5">������</TD>
								<TD CLASS="TD6"><script language =javascript src='./js/f5103ba1_fpDateTime1_txtIssueDt.js'></script></TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_01%> WIDTH=100%></TD>
	</TR>  
	<TR HEIGHT=20>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_30%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD><BUTTON NAME="btnExec" CLASS="CLSMBTN" OnClick="VBScript:Call FnButtonExec()" Flag=1>�� ��</BUTTON></TD>
					<TD WIDTH=*>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" WIDTH=100% HEIGHT= <%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24" TABINDEX="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

