<%@ LANGUAGE="VBSCRIPT" %>
<!--'**********************************************************************************************
'*  1. Module Name          : Quality
'*  2. Function Name        : 
'*  3. Program ID           : q3611ma1
'*  4. Program Name         : ������ó�� 
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2003-09-04
'*  8. Modified date(Last)  : 2003-09-04
'*  9. Modifier (First)     : Jaewoo Koh
'* 10. Modifier (Last)      : Jaewoo Koh
'* 11. Comment              :
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
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit                                                           

'==========================================================================================================
Const BIZ_PGM_SUMMARY_ID = "q3611mb1.asp"
Const BIZ_PGM_CONFIRM_ID = "q3611mb2.asp"
Const BIZ_PGM_CANCEL_CONFIRM_ID = "q3611mb3.asp"

'=========================================================================================================
'	1. ���� ǥ�ؿ� ����. prefix�� g�� �����.
'	2.Array�� ���� ()�� �ݵ�� ����Ͽ� �Ϲ� ������ ������ �� 
'========================================================================================================= 
<!-- #Include file="../../inc/lgVariables.inc" -->

'========================================================================================================= 
'----------------  ���� Global ������ ����  ----------------------------------------------------------- 

'++++++++++++++++  Insert Your Code for Global Variables Assign  ++++++++++++++++++++++++++++++++++++++ 
Dim IsOpenPop       

Dim CompanyYM

CompanyYM = UNIMonthClientFormat(UniConvDateAToB("<%=GetSvrDate%>", Parent.gServerDateFormat, Parent.gAPDateFormat))

'=========================================================================================================
'	Name : InitVariables()
'	Description : ���� �ʱ�ȭ(Global ����, �ʱ�ȭ�� �ʿ��� ���� �Ǵ� Flag���� Setting�Ѵ�.)
'========================================================================================================= 
Sub InitVariables()
    IsOpenPop = False
End Sub

'==========================================  2.2.1 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : ȭ�� �ʱ�ȭ(���� Field�� �� �� ȭ���� �� �� Default���� ������� �ϴ� Field�� Setting)
'========================================================================================================= 
Sub SetDefaultVal()
	frm1.txtSummaryDt.Text = CompanyYM
End Sub

'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== 
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "Q", "NOCOOKIE","MA") %>
End Sub

'------------------------------------------  OpenPlant()  -------------------------------------------------
'	Name : OpenPlant()
'	Description : Plant PopUp
'--------------------------------------------------------------------------------------------------------- 
Sub OpenPlant() 
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Sub

	IsOpenPop = True

	arrParam(0) = "�����˾�"	
	arrParam(1) = "B_PLANT"				
	arrParam(2) = Trim(frm1.txtPlantCd.Value)
	arrParam(3) = ""
	arrParam(4) = ""			
	arrParam(5) = "����"			
	
	arrField(0) = "PLANT_CD"	
	arrField(1) = "PLANT_NM"	
	
	arrHeader(0) = "�����ڵ�"		
	arrHeader(1) = "�����"		

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		frm1.txtPlantCd.Value = arrRet(0)		
		frm1.txtPlantNm.Value = arrRet(1)
	End If	

	frm1.txtPlantCd.Focus	
	Set gActiveElement = document.activeElement
End Sub


'=========================================================================================================
' Name : SummaryOperation()    
' Description : MPS ���� Main Function          
'========================================================================================================= 
Function SummaryOperation()
    
    SummaryOperation = False
	
	Err.Clear           
	On Error Resume Next
	
	'-----------------------
	'Check condition area
	'----------------------- 
	If Not chkField(Document, "1") Then
		Exit Function
	End If
	
	If Not CheckInspClassSelected() Then
		Call DisplayMsgBox("224705","X","X","X") 		'���õ� �˻�з��� �����ϴ�. 
		Exit Function
	End If
	'-----------------------
	'Save function call area
	'-----------------------
	If DbSummaryOperation("S") = False then Exit Function	
	
    SummaryOperation = True 
End Function

'=========================================================================================================
' Name : Confirm()    
' Description : 
'========================================================================================================= 
Function Confirm()

    Confirm = False

	Err.Clear           
	On Error Resume Next
	
	'-----------------------
	'Check condition area
	'----------------------- 
	If Not chkField(Document, "1") Then
		Exit Function
	End If
	
	If Not CheckInspClassSelected() Then
		Call DisplayMsgBox("224705","X","X","X") 		'���õ� �˻�з��� �����ϴ�. 
		Exit Function
	End If
	'-----------------------
	'Save function call area
	'-----------------------
	If DbSummaryOperation("C") = False then Exit Function	
	
    Confirm = True  

End Function

'=========================================================================================================
' Name : CancelConfirm()    
' Description : 
'========================================================================================================= 
Function CancelConfirm()

    CancelConfirm = False

	Err.Clear           
	On Error Resume Next
	
	'-----------------------
	'Check condition area
	'----------------------- 
	If Not chkField(Document, "1") Then
		Exit Function
	End If
	
	If Not CheckInspClassSelected() Then
		Call DisplayMsgBox("224705","X","X","X") 		'���õ� �˻�з��� �����ϴ�. 
		Exit Function
	End If
	'-----------------------
	'Save function call area
	'-----------------------
	If DbSummaryOperation("R") = False then Exit Function	    
    CancelConfirm = True  

End Function

'========================================================================================
' Function Name : DbSummaryOperation
' Function Desc : 
'========================================================================================
Function DbSummaryOperation(Byval pvStrAction) 
	DbSummaryOperation = False 
	
	err.Clear 
	On Error Resume Next    
	
	Call LayerShowHide(1)
	
	Dim strVal
	Dim strR_YesOrNo
	Dim strP_YesOrNo
	Dim strF_YesOrNo
	Dim strS_YesOrNo
	
	With frm1
		If .chkInspClass_R.checked = true then
			strR_YesOrNo = "R"
		End If
	
		If .chkInspClass_P.checked = true then
			strP_YesOrNo = "P"
		End If
	
		If .chkInspClass_F.checked = true then
			strF_YesOrNo = "F"
		End If
	
		If .chkInspClass_S.checked = true then
			strS_YesOrNo = "S"
		End If
		
		strVal = BIZ_PGM_SUMMARY_ID & "?txtAction=" & pvStrAction _
										& "&txtPlantCd=" & UCase(.txtPlantCd.value) _
										& "&txtYr=" & Left(.txtSummaryDt.DateValue,4) _
										& "&txtMnth=" & Mid(.txtSummaryDt.DateValue,5, 2) _
										& "&txtRYesorNo=" & strR_YesOrNo _
										& "&txtPYesorNo=" & strP_YesOrNo _
										& "&txtFYesorNo=" & strF_YesOrNo _
										& "&txtSYesorNo=" & strS_YesOrNo
			
	End With
	
	Call RunMyBizASP(MyBizASP, strVal)
	
    DbSummaryOperation = True
End Function

'========================================================================================
' Function Name : CheckInspClassSelected
' Function Desc : �˻�з��� ��� �ϳ��� ���õǾ��� �� üũ 
'========================================================================================
Function CheckInspClassSelected()
	CheckInspClassSelected = True
	With frm1
		If .chkInspClass_R.checked = false And _
			.chkInspClass_P.checked = false And _
			.chkInspClass_F.checked = false And _
			.chkInspClass_S.checked = false then

			CheckInspClassSelected = False

		End If
	End With
End Function
'=========================================================================================================
'	Name : Form_Load()
'	Description : Window On Load(���� Include ���Ͽ� ����)�� �����ʱ�ȭ �� ȭ���ʱ�ȭ�� �ϱ� ���� �Լ��� Call�ϴ� �κ� 
'========================================================================================================= 
Sub Form_Load()
	Call LoadInfTB19029
    Call ggoOper.FormatField(Document, "1", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")  
    
    Call ggoOper.FormatDate(frm1.txtSummaryDt, Parent.gDateFormat, 2)
    Call SetDefaultVal
    
    Call SetToolBar("10000000000011")
    Call InitVariables    
    
    If Parent.gPlant <> "" Then
		frm1.txtPlantCd.value = UCase(Parent.gPlant)
		frm1.txtPlantNm.value = Parent.gPlantNm
	Else
		frm1.txtPlantCd.focus 
    End If
    
End Sub

'=========================================================================================================
'   Event Name : txtSummaryDt_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=========================================================================================================
Sub txtSummaryDt_DblClick(Button) 
	If Button = 1 Then 
		frm1.txtSummaryDt.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtSummaryDt.Focus
	End If 
End Sub

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
    Call parent.FncExport(parent.C_SINGLE)
End Function

'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================
Function FncFind() 
    Call parent.FncFind(parent.C_SINGLE, False)
End Function

'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================
Function FncExit()
	FncExit = True
End Function

</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
</HEAD>
<BODY TABINDEX="-1" SCROLL="NO">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE CLASS="BatchTB4" CELLSPACING=0 CELLPADDING=0>
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>������ó��</font></td>
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
						<TABLE CLASS="BasicTB" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<TD CLASS="TD5" NOWRAP>����</TD>
									<TD CLASS="TD6" NOWRAP>
										<INPUT TYPE=TEXT NAME="txtPlantCd" SIZE="10" MAXLENGTH="4" ALT="����" TAG="12XXXU"><IMG ALIGN=top HEIGHT=20 NAME=btnPlantPopup ONCLICK=vbscript:OpenPlant() SRC="../../../CShared/image/btnPopup.gif" WIDTH=16  TYPE="BUTTON">&nbsp;<INPUT NAME="txtPlantNm" TAG="14X">
								</TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>������</TD>
								<TD CLASS="TD6" NOWRAP>
									<script language =javascript src='./js/q3611ma1_txtSummaryDt_txtSummaryDt.js'></script>
								</TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>�˻�з�</TD>
								<TD CLASS=TD6 NOWRAP>
									<INPUT TYPE="CHECKBOX" CLASS="Check" NAME=chkInspClass ID=chkInspClass_R VALUE="R" TAG="11"><LABEL FOR="chkInspClass_R">���԰˻�</LABEL>&nbsp;
									<INPUT TYPE="CHECKBOX" CLASS="Check" NAME=chkInspClass ID=chkInspClass_P VALUE="P" TAG="11"><LABEL FOR="chkInspClass_P">�����˻�</LABEL>&nbsp;
									<INPUT TYPE="CHECKBOX" CLASS="Check" NAME=chkInspClass ID=chkInspClass_F VALUE="F" TAG="11"><LABEL FOR="chkInspClass_F">�����˻�</LABEL>&nbsp;
									<INPUT TYPE="CHECKBOX" CLASS="Check" NAME=chkInspClass ID=chkInspClass_S VALUE="S" TAG="11"><LABEL FOR="chkInspClass_S">���ϰ˻�</LABEL>
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
					<TD><BUTTON NAME="btnSummary" CLASS="CLSMBTN" Flag=1 onclick="SummaryOperation()">����</BUTTON>&nbsp;<BUTTON NAME="btnConfirm" CLASS="CLSMBTN" Flag=1 onclick="Confirm()">Ȯ��</BUTTON>&nbsp;<BUTTON NAME="btnCancelConfirm" CLASS="CLSMBTN" Flag=1 onclick="CancelConfirm()">Ȯ�����</BUTTON></TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME></TD>
	</TR>
</TABLE>
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

