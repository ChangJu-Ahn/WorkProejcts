<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : PRODUCTION
'*  2. Function Name        :
'*  3. Program ID           : p4920ma1
'*  4. Program Name         : �ڿ��Һ��� batch
'*  5. Program Desc         :
'*  6. Comproxy List        :
'*  7. Modified date(First) : 2005-02-04
'*  8. Modified date(Last)  :
'*  9. Modifier (First)     : Yoon, Jeong Woo
'* 10. Modifier (Last)      :
'* 11. Comment              :
'**********************************************************************************************
-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!--'#########################################################################################################
'												1. �� �� �� 
'##########################################################################################################-->
<!-- '******************************************  1.1 Inc ����   **********************************************
'	���: Inc. Include
'********************************************************************************************************* -->
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->
<!--=======================================  1.1.1 Style Sheet  ======================================
'==========================================================================================================-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--=======================================  1.1.2 ���� Include   ======================================
'==========================================================================================================-->
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incEB.vbs"></SCRIPT>              <!--��:Print Program needs this vbs file-->
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit																	'��: indicates that All variables must be declared in advance

'******************************************  1.2 Global ����/��� ����  ***********************************
'	1. Constant�� �ݵ�� �빮�� ǥ��.
'**********************************************************************************************************


'==========================================  1.2.1 Global ��� ����  ======================================
'==========================================================================================================

'==========================================  1.2.2 Global ���� ����  =====================================
'	1. ���� ǥ�ؿ� ����. prefix�� g�� �����.
'	2.Array�� ���� ()�� �ݵ�� ����Ͽ� �Ϲ� ������ ������ �� 
'=========================================================================================================
'==========================================  1.2.2 Global ���� ����  =====================================
'	1. ���� ǥ�ؿ� ����. prefix�� g�� �����.
'	2.Array�� ���� ()�� �ݵ�� ����Ͽ� �Ϲ� ������ ������ �� 
'=========================================================================================================
Const BIZ_PGM_EXEC_ID = "p4920mb1.asp"

Dim  lgBlnFlgChgValue           ' Variable is for Dirty flag
Dim  lgIntFlgMode               ' Variable is for Operation Status
Dim  lgIntGrpCount              ' initializes Group View Size
Dim  IsOpenPop
'==========================================  1.2.3 Global Variable�� ����  ===============================
'=========================================================================================================
'----------------  ���� Global ������ ����  -----------------------------------------------------------
'++++++++++++++++  Insert Your Code for Global Variables Assign  ++++++++++++++++++++++++++++++++++++++
Dim  LocSvrDate
Dim  EndDate
Dim  ToDate

LocSvrDate = "<%=GetSvrDate%>"
EndDate = UniConvDateAtoB(LocSvrDate,parent.gServerDateFormat,parent.gDateFormat)     	'��: �ʱ�ȭ�鿡 �ѷ����� ���� ��¥ 
ToDate = UNIDateAdd("D",7,EndDate,parent.gDateFormat)							    '��: �ʱ�ȭ�鿡 �ѷ����� ������ ��¥ 

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

    lgIntFlgMode = parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
    IsOpenPop = False
End Sub

'******************************************  2.2 ȭ�� �ʱ�ȭ �Լ�  ***************************************
'	���: ȭ���ʱ�ȭ 
'	����: ȭ���ʱ�ȭ, Combo Display, ȭ�� Clear �� ȭ�� �ʱ�ȭ �۾��� �Ѵ�.
'*********************************************************************************************************
'==========================================  2.2.1 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : ȭ�� �ʱ�ȭ(���� Field�� �� �� ȭ���� �� �� Default���� ������� �ϴ� Field�� Setting)
'=========================================================================================================
Sub SetDefaultVal()
	frm1.txtWorkDt.Text = EndDate
'	frm1.txtEndDt.Text = ToDate
End Sub

'=======================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'========================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "P", "NOCOOKIE", "OA") %>
End Sub

'#########################################################################################################
'												3. Event�� 
'	���: Event �Լ��� ���� ó�� 
'	����: Windowó��, Singleó��, Gridó�� �۾�.
'         ���⼭ Validation Check, Calcuration �۾��� ������ Event�� �߻�.
'         �� Object������ Grouping�Ѵ�.
'##########################################################################################################
'******************************************  3.1 Window ó��  *********************************************
'	Window�� �߻� �ϴ� ��� Even ó�� 
'*********************************************************************************************************
'==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(���� Include ���Ͽ� ����)�� �����ʱ�ȭ �� ȭ���ʱ�ȭ�� �ϱ� ���� �Լ��� Call�ϴ� �κ� 
'=========================================================================================================
Sub Form_Load()

    Call LoadInfTB19029                                                     '��: Load table , B_numeric_format

	Call ggoOper.FormatField(Document, "x",ggStrIntegeralPart, ggStrDeciPointPart, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec)

    Call ggoOper.LockField(Document, "N")
                                       '��: Lock  Suitable  Field
	Call SetDefaultVal
    Call InitVariables		'��: Initializes local global variables

    Call SetToolbar("10000000000011")										'��: ��ư ���� ���� 

    If parent.gPlant <> "" Then
		frm1.txtPlantCd.value = Ucase(parent.gPlant)
		frm1.txtPlantNm.value = parent.gPlantNm
		Set gActiveElement = document.activeElement
	Else
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement
	End If

End Sub

'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================
Sub Form_QueryUnload(Cancel, UnloadMode)

End Sub

'**************************  3.2 HTML Form Element & Object Eventó��  **********************************
'	Document�� TAG���� �߻� �ϴ� Event ó�� 
'	Event�� ��� �Ʒ��� ����� Event�̿��� ����� �����ϸ� �ʿ�� �߰� �����ϳ� 
'	Event�� �浹�� ����Ͽ� �ۼ��Ѵ�.
'*********************************************************************************************************

'******************************  3.2.1 Object Tag ó��  *********************************************
'	Window�� �߻� �ϴ� ��� Even ó�� 
'*********************************************************************************************************
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
Function FncQuery()
	On Error Resume Next                                                    '��: Protect system from crashing
End Function

Function FncSave()
	On Error Resume Next                                                    '��: Protect system from crashing
End Function

Function FncNew()
	On Error Resume Next                                                    '��: Protect system from crashing
End Function

Function FncDelete()
	On Error Resume Next                                                    '��: Protect system from crashing
End Function

Function FncInsertRow()
	On Error Resume Next                                                    '��: Protect system from crashing
End Function

Function FncDeleteRow()
	On Error Resume Next                                                    '��: Protect system from crashing
End Function

Function FncCopy()
	On Error Resume Next                                                    '��: Protect system from crashing
End Function

Function FncCancel()
	On Error Resume Next                                                    '��: Protect system from crashing
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
	Call parent.FncExport(parent.C_SINGLE)
End Function

'========================================================================================
' Function Name : FncFind
' Function Desc :
'========================================================================================
Function FncFind()
    Call parent.FncFind(parent.C_SINGLE , False)                                     '��:ȭ�� ����, Tab ���� 
End Function

'========================================================================================
' Function Name : BtnPrint
' Function Desc : This function is related to Print Button
'========================================================================================
'Function BtnPrint()
'	Dim  strEbrFile
'    Dim  objName
'
'	Dim  var1
'	Dim  var2
'	Dim  var3
'	Dim  var4
'	Dim  var5
'
'	Dim  strUrl
'	Dim  arrParam, arrField, arrHeader
'
'	If Not chkfield(Document, "x") Then									'��: This function check indispensable field
'		Call BtnDisabled(0)
'       Exit Function
'    End If
'
''	If ValidDateCheck(frm1.txtWorkDt, frm1.txtEndDt) = False Then
''		Call BtnDisabled(0)
''		Exit Function
''	End IF
'
'	If frm1.txtPlantCd.value = "" Then
'		frm1.txtPlantNm.value = ""
'	End If
'
'	If frm1.txtFromWcCd.value = "" Then
'		frm1.txtFromWcNm.value = ""
'	End If
'
'	If frm1.txtToWcCd.value = "" Then
'		frm1.txtToWcNm.value = ""
'	End If
'
'	var1 = Trim(frm1.txtPlantCd.value)
'
'	If frm1.txtFromWcCd.value = "" Then
'		var2 = "0"
'	Else
'		var2 = Trim(frm1.txtFromWcCd.value)
'	End If
'
'	If frm1.txtToWcCd.value = "" Then
'		var3 = "zzzzzzz"
'	Else
'		var3 = Trim(frm1.txtToWcCd.value)
'	End If
'
'	var4 = UniConvDateAtoB(frm1.txtWorkDt.Text,parent.gDateFormat,parent.gServerDateFormat)
'	var5 = UniConvDateAtoB(frm1.txtEndDt.Text,parent.gDateFormat,parent.gServerDateFormat)
'
'	strUrl = strUrl & "plant_cd|" & var1
'	strUrl = strUrl & "|from_wc_cd|" & var2
'	strUrl = strUrl & "|to_wc_cd|" & var3
'	strUrl = strUrl & "|start_date|" & var4
'	strUrl = strUrl & "|end_date|" & var5
'
'	strEbrFile = "p4920ma1"
'	objName = AskEBDocumentName(strEbrFile,"ebr")
'
''----------------------------------------------------------------
'' Print �Լ����� �߰��Ǵ� �κ� 
''----------------------------------------------------------------
'	call FncEBRprint(EBAction, objName, strUrl)
''----------------------------------------------------------------
'
'	Call BtnDisabled(0)
'
'	frm1.btnRun(1).focus
'	Set gActiveElement = document.activeElement
'
'End Function

'========================================================================================
' Function Name : BtnExecute
' Function Desc : This function is related to Preview Button
'========================================================================================
Function BtnExecute()                                                    '��: Protect system from crashing
    Dim  objName
	Dim  arrParam, arrField, arrHeader
	Dim IntRetCD

	Call BtnDisabled(1)

	If Not chkfield(Document, "x") Then									'��: This function check indispensable field
		Call BtnDisabled(0)
       Exit Function
    End If

	IntRetCD = DisplayMsgBox("900018", parent.VB_YES_NO, "x", "x")	'��: Display Message
	If IntRetCD = vbNo Then
		Exit Function
	Else
	
		If DbSave = False Then
	'		Call BtnDisabled(0)
			Exit Function
		End If

	End If

	Call BtnDisabled(0)

	frm1.btnRun.focus
	Set gActiveElement = document.activeElement

End Function

Function DbSaveOk()
	Call DisplayMsgBox("183114", "X", "x", "x")
End Function

Function DbSaveFail()
	Call DisplayMsgBox("800506", "X", "x", "x")
End Function

'========================================================================================
' Function Name : DbSave()
' Function Desc : 
'========================================================================================
Function DbSave()
	Dim strVal
	Dim strPlantCd
	Dim strWcCd
	Dim strItemCd
	Dim strFrDt

    DbSave = False														'��: Processing is NG

    strVal = BIZ_PGM_EXEC_ID & "?txtMode=" & parent.UID_M0003
    strVal = strVal & "&txtPlantCd=" & Trim(frm1.txtPlantCd.value)				'��: 
    strVal = strVal & "&txtWorkDt=" & Trim(frm1.txtWorkDt.Text)				'��: 
    strVal = strVal & "&txtResourceCd=" & Trim(frm1.txtResourceCd.value)		'��: 

	Call RunMyBizASP(MyBizASP, strVal)										'��: �����Ͻ� ASP �� ���� 

    DbSave = True     

    LayerShowHide(1)
End Function


'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================
Function FncQuery()
	On Error Resume Next
End Function


'========================================================================================
' Function Name : FncExit
' Function Desc :
'========================================================================================
Function FncExit()
	Dim IntRetCD

	FncExit = False

	FncExit = True
End Function


'########################################################################################
'########################################################################################
'# Area Name   : User-defined Method Part
'# Description : This part declares user-defined method
'########################################################################################
'########################################################################################
    '----------  Coding part  -------------------------------------------------------------

Function OpenPlantCd()

	Dim  arrRet
	Dim  arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "�����˾�"		' �˾� ��Ī 
	arrParam(1) = "B_PLANT"				' TABLE ��Ī 
	arrParam(2) = Trim(frm1.txtPlantCd.Value)	' Code Condition
	arrParam(3) = ""					' Name Cindition
	arrParam(4) = ""					' Where Condition
	arrParam(5) = "����"			' TextBox ��Ī 

    arrField(0) = "PLANT_CD"			' Field��(0)
    arrField(1) = "PLANT_NM"			' Field��(1)

    arrHeader(0) = "����"			' Header��(0)
    arrHeader(1) = "�����"			' Header��(1)

	arrRet = window.showModalDialog("../../comasp/commonpopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetPlantCd(arrRet)
	End If

	Call SetFocusToDocument("M")
	frm1.txtPlantCd.focus

End Function

'------------------------------------------  OpenResource()  -------------------------------------------------
'	Name : OpenResource()
'	Description : Resource PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenResource()

	Dim arrRet
	Dim arrParam(5), arrField(6),arrHeader(6)

	If frm1.txtPlantCd.value = "" Then
		Call DisplayMsgBox("971012","X", "����","X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement
		Exit Function
	End If

	IsOpenPop = True
	arrParam(0) = "�ڿ��˾�"
	arrParam(1) = "P_RESOURCE"
	arrParam(2) = Trim(frm1.txtResourceCd.Value)
	arrParam(3) = ""
	arrParam(4) = "PLANT_CD = " & FilterVar(UCase(frm1.txtPlantCd.value), "''", "S")
	arrParam(5) = "�ڿ�"

    arrField(0) = "RESOURCE_CD"
    arrField(1) = "DESCRIPTION"

    arrHeader(0) = "�ڿ�"
    arrHeader(1) = "�ڿ���"

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetResource(arrRet)
	End If

	Call SetFocusToDocument("M")
	frm1.txtResourceCd.focus

End Function

'------------------------------------------  SetResource()  ----------------------------------------------
'	Name : SetResource()
'	Description : Resource Popup���� Return�Ǵ� �� setting
'---------------------------------------------------------------------------------------------------------
Function SetResource(byval arrRet)
	frm1.txtResourceCd.Value    = arrRet(0)
	frm1.txtResourceNm.Value    = arrRet(1)
End Function

'------------------------------------------  SetPlantCd()  -----------------------------------------------
'	Name : SetPlantCd()
'	Description : Resource Popup���� Return�Ǵ� �� setting
'---------------------------------------------------------------------------------------------------------
Function SetPlantCd(ByVal arrRet)
	frm1.txtPlantCd.value = arrRet(0)
	frm1.txtPlantNm.value = arrRet(1)
End Function
'=======================================================================================================
'   Event Name : txtValidFromDt_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtWorkDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtWorkDt.Action = 7
        Call SetFocusToDocument("M")
		Frm1.txtWorkDt.Focus
    End If
End Sub

'��: �Ʒ� OBJECT Tag�� InterDev ����ڸ� ���Ѱ����� ���α׷��� �ϼ��Ǹ� �Ʒ� Include �ڵ�� ��ü�Ǿ�� �Ѵ� 
</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->
</HEAD>
<BODY TABINDEX="-1" SCROLL="NO">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE CLASS="BatchTB1" CELLSPACING=0 CELLPADDING=0>
	<TR>
		<TD HEIGHT=5>&nbsp;<% ' ���� ���� %></TD>
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>�ڿ��Һ��� BATCH</font></td>
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
			<TABLE CLASS="BasicTB" CELLSPACING=0 >
	    		<TR>
					<TD HEIGHT=20>
							<TABLE CLASS="TB3" WIDTH=100% CELLSPACING=0 CELLPADDING=0>
								<TR>
								    <TD CLASS="TD5" NOWRAP>����</TD>
									<TD CLASS="TD6" NOWRAP>
									<INPUT TYPE=TEXT NAME="txtPlantCd" SIZE=10 MAXLENGTH=4 tag="x2xxxU" ALT="����"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenPlantCd()">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=25 MAXLENGTH=40 tag="x4" ALT="�����">
									</TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>�۾���</TD>
									<TD CLASS="TD6" NOWRAP>
										<script language =javascript src='./js/p4920ma1_I872908192_txtWorkDt.js'></script>
										<!--OBJECT classid=<%=gCLSIDFPDT%> name=txtEndDt CLASS=FPDTYYYYMMDD title=FPDATETIME tag="x2X1"
											ALT="������" MAXLENGTH="10" SIZE="10" VIEWASTEXT >
										</OBJECT-->
									</TD>
								</TR>
								<TR>
								<TD CLASS=TD5 NOWRAP>�ڿ�</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtResourceCd" SIZE=10 MAXLENGTH=10 tag="x2xxxU" ALT="�ڿ�"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnResourceCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenResource()">&nbsp;<INPUT TYPE=TEXT NAME="txtResourceNm" SIZE=25 tag="14"></TD>
								</TR>
								<!--TR>
								    <TD CLASS="TD5" NOWRAP>�۾���</TD>
									<TD CLASS="TD6" NOWRAP>
									<INPUT TYPE=TEXT NAME="txtFromWcCd" SIZE=10 MAXLENGTH=7 tag="x1xxxU" ALT="�۾���"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnFromWcCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenFromWcCd()">&nbsp;<INPUT TYPE=TEXT NAME="txtFromWcNm" SIZE=40 MAXLENGTH=40 tag="x4" ALT="�۾����">&nbsp;~&nbsp;
									</TD>
								</TR>
								<TR>
								    <TD CLASS="TD5" NOWRAP>&nbsp;</TD>
									<TD CLASS="TD6" NOWRAP>
									<INPUT TYPE=TEXT NAME="txtToWcCd" SIZE=10 MAXLENGTH=7 tag="x1xxxU" ALT="�۾���"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnToWcCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenToWcCd()">&nbsp;<INPUT TYPE=TEXT NAME="txtToWcNm" SIZE=40 MAXLENGTH=40 tag="x4" ALT="�۾����">&nbsp;
									</TD>
								</TR-->
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
				     <TD WIDTH = 10 > &nbsp; </TD>
				     <TD>
		               <BUTTON NAME="btnRun" CLASS="CLSSBTN" ONCLICK="vbscript:BtnExecute()" Flag=1>����</BUTTON>&nbsp;<!--BUTTON NAME="btnRun" CLASS="CLSSBTN" ONCLICK="vbscript:FncExit()" Flag=1>����</BUTTON-->
                     </TD>
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
<!-- Print Program must contain this HTML Code -->
<FORM NAME=EBAction TARGET="MyBizASP" METHOD="POST" >
    <input type="hidden" name="uname">
    <input type="hidden" name="dbname">
    <input type="hidden" name="filename">
    <input type="hidden" name="condvar">
	<input type="hidden" name="date">
</FORM>
<!-- End of Print HTML Code -->
</BODY>
</HTML>