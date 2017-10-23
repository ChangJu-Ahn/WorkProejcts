
<%@ LANGUAGE="VBSCRIPT" %>
<% Option Explicit %>
<!--'**********************************************************************************************
'*
'*  1. Module Name          : Account
'*  2. Function Name        : 
'*  3. Program ID           : a5401ma1
'*  4. Program Name         : �̰�������ص�� 
'*  5. Program Desc         : �̰�������ص�� 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2002/11/7
'*  8. Modified date(Last)  : 2002/11/7
'*  9. Modifier (First)     : Jung Sung Ki
'* 10. Modifier (Last)      : Jung Sung Ki
'* 11. Comment              :
'*
'***********************************************************************k*********************** -->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!--
'##########################################################################################################
'												1. �� �� �� 
'##########################################################################################################

'********************************************  1.1 Inc ����   ********************************************
'	���: Inc. Include
'********************************************************************************************************* -->
<!-- #Include file="../../inc/IncSvrCcm.inc"  -->

<!--

'============================================  1.1.1 Style Sheet  =======================================
'======================================================================================================== -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		<!-- '��: �ش� ��ġ�� ���� �޶���, ��� ���  -->
<!--
'============================================  1.1.2 ���� Include  ======================================
'======================================================================================================== -->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"	  SRC="../../inc/incCliRdsQuery.vbs">	</SCRIPT>

<SCRIPT LANGUAGE="VBScript">

Option Explicit                                                '��: indicates that All variables must be declared in advance 


'********************************************  1.2 Global ����/��� ����  *********************************
'	1. Constant�� �ݵ�� �빮�� ǥ��.
'**********************************************************************************************************
<!-- #Include file="../../inc/lgvariables.inc" -->

'============================================  1.2.1 Global ��� ����  ====================================
'==========================================================================================================

Const BIZ_PGM_ID = "a5401mb1.asp"											 '��: �����Ͻ� ���� ASP�� 

'============================================  1.2.2 Global ���� ����  ===================================
'	1. ���� ǥ�ؿ� ����. prefix�� g�� �����.
'	2. Array�� ���� ()�� �ݵ�� ����Ͽ� �Ϲ� ������ ������ �� 
'========================================================================================================= 

'==========================================  1.2.3 Global Variable�� ����  ===============================
'========================================================================================================= 
 '----------------  ���� Global ������ ����  ----------------------------------------------------------- 

 '++++++++++++++++  Insert Your Code for Global Variables Assign  ++++++++++++++++++++++++++++++++++++++ 
Dim lgMpsFirmDate, lgLlcGivenDt											 '��: �����Ͻ� ���� ASP���� �����ϹǷ� Dim 

Dim lgCurName()															'�� : ���� ȭ��� �ʿ��� ��Į ���� ���� 
Dim cboOldVal          
Dim IsOpenPop          
Dim lgCboKeyPress      
Dim lgOldIndex								
Dim lgOldIndex2        



'========================================================================================================= 
Sub InitVariables()

    lgIntFlgMode = Parent.OPMD_CMODE                                               '��: Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                                                '��: Indicates that no value changed
    lgIntGrpCount = 0                                                       '��: Initializes Group View Size
    '-----------------------  Coding part  ------------------------------------------------------------- 
    IsOpenPop = False														'��: ����� ���� �ʱ�ȭ 
    lgCboKeyPress = False
    lgOldIndex = -1
    lgOldIndex2 = -1
    lgMpsFirmDate=""
    lgLlcGivenDt=""
	if Trim(frm1.txtAcctBaseNo.value)="" then
		frm1.txtAcctBaseNo.value = frm1.hAcctBaseNo.value
	end if

	frm1.txtAcctBaseNo.focus
	Set gActiveElement = document.activeElement
  
End Sub


'========================================================================================================= 
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
	<%Call LoadInfTB19029A("I","*", "NOCOOKIE", "MA") %>
End Sub


'******************************************  2.2 ȭ�� �ʱ�ȭ �Լ�  ***************************************
'	���: ȭ���ʱ�ȭ 
'	����: ȭ���ʱ�ȭ, Combo Display, ȭ�� Clear �� ȭ�� �ʱ�ȭ �۾��� �Ѵ�. 
'********************************************************************************************************* 

'--------------------------------------------------------------------------------------------------------- 

Sub InitComboBox_One()
	Dim IntRetCD1
	Dim IntValMM
	Dim IntValDD
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	                   'Select                 From        Where                Return value list  
    Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("F5004", "''", "S") & "  order by minor_cd ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboCardMM ,lgF0  ,lgF1  ,Chr(11))

	Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("F5005", "''", "S") & "  order by minor_cd ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboCardDD ,lgF0  ,lgF1  ,Chr(11))

    '------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub





Function OpenAcctBaseNo(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "�̰�����˾�"					' �˾� ��Ī 
	arrParam(1) = "A_OPEN_ACCT_BASE"						' TABLE ��Ī 
	arrParam(2) = strCode							' Code Condition
	arrParam(3) = ""								' Name Cindition
	arrParam(4) = ""								' Where Condition
	arrParam(5) = "�̰�����ڵ�"
	
    arrField(0) = "ACCT_BASE_NO"							' Field��(0)
    arrField(1) = "ACCT_BASE_NM"						' Field��(1)
    
    arrHeader(0) = "�̰����"					' Header��(0)
    arrHeader(1) = "�̰������"					' Header��(1)
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
	frm1.txtAcctBaseNo.focus
	    Exit Function
	Else
		frm1.txtAcctBaseNo.focus
		frm1.txtAcctBaseNo.value = arrRet(0)
		frm1.txtAcctBaseNm.value = arrRet(1)
	End If	

End Function



'++++++++++++++++++++++++++++++++++++++++++  2.5 ������ ���� �Լ�  +++++++++++++++++++++++++++++++++++++++
'    ������ ���α׷� ���� �ʿ��� ������ ���� Procedure (Sub, Function, Validation & Calulation ���� �Լ�)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ 
'Sub Combo_Change(Index As Integer)
'	lgBlnFlgChgValue = True
'End Sub


'###########################################################################################################
'												3. Event�� 
'	���: Event �Լ��� ���� ó�� 
'	����: Windowó��, Singleó��, Gridó�� �۾�.
'         ���⼭ Validation Check, Calcuration �۾��� ������ Event�� �߻�.
'         �� Object������ Grouping�Ѵ�.
'##########################################################################################################


'==========================================================================================================
Sub Form_Load()

    Call InitVariables																'��: Initializes local global variables
    Call LoadInfTB19029																'��: Load table , B_numeric_format
	Call AppendNumberPlace("6","4","0")
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")											'��: Lock  Suitable  Field
    
    '----------  Coding part  -------------------------------------------------------------
    Call SetToolbar("1100100000001111")
    Call InitComboBox_One

	frm1.txtAcctBaseNo.focus 
	frm1.txtAcctBaseNo.value="1"

    lgBlnFlgChgValue = False                                                '��: Indicates that no value changed

	FncQuery 

	Set gActiveElement = document.activeElement


End Sub


'==========================================================================================

Sub Form_QueryUnload(Cancel , UnloadMode )

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


'========================================================================================

Function FncQuery() 
    Dim IntRetCD 
    
    FncQuery = False                                                        '��: Processing is NG

    Err.Clear                                                               '��: Protect system from crashing
  '-----------------------
    'Check previous data area
    '----------------------- 
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO, "X", "X")				'��: "Will you destory previous data"
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

  '-----------------------
    'Erase contents area
    '----------------------- 
    Call InitVariables															'��: Initializes local global variables
    Call ggoOper.ClearField(Document, "2")										'��: Clear Contents  Field

  '-----------------------
    'Check condition area
    '----------------------- 
    If Not chkField(Document, "1") Then									'��: This function check indispensable field
       Exit Function
    End If

  '-----------------------
    'Query function call area
    '----------------------- 
    Call DbQuery																'��: Query db data
    FncQuery = True																'��: Processing is OK
        
End Function



'========================================================================================

Function FncNew() 
    Dim IntRetCD 
    
    FncNew = False                                                     '��: Processing is NG
    
  '-----------------------
    'Check previous data area
    '-----------------------
    If lgBlnFlgChgValue = True Then
       IntRetCD = DisplayMsgBox("900015", Parent.VB_YES_NO, "X", "X")           '��: "Will you destory previous data"
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
  '-----------------------
    'Erase condition area
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "1")                                      '��: Clear Condition Field
    Call ggoOper.ClearField(Document, "2")                                      '��: Clear Contents  Field
    Call ggoOper.LockField(Document, "N")                                       '��: Lock  Suitable  Field
    Call InitVariables															'��: Initializes local global variables
    
    Call SetToolbar("1100100000001111")

	frm1.txtAcctBaseNo.focus

    FncNew = True																'��: Processing is OK
    Set gActiveElement = document.activeElement
    
End Function


'========================================================================================

Function FncDelete() 
    Dim IntRetCD
    
    FncDelete = False														'��: Processing is NG
    
  '-----------------------
    'Precheck area
    '-----------------------
    If lgIntFlgMode <> Parent.OPMD_UMODE Then                                      'Check if there is retrived data
        Call DisplayMsgBox("900002", "X", "X", "X")                                
        Exit Function
    End If
    
  '-----------------------
    'Delete function call area
    '-----------------------
    IntRetCD = DisplayMsgBox("900003",Parent.VB_YES_NO, "X", "X")
    IF IntRetCD = vbNO Then
		Exit Function
	End IF	
    
    Call DbDelete															'��: Delete db data
    
    FncDelete = True                                                        '��: Processing is OK
    
End Function


'========================================================================================
Function FncSave() 
    Dim IntRetCD 
    

	FncSave = False                                                         '��: Processing is NG

	Err.Clear                                                               '��: Protect system from crashing
	    
	'-----------------------
	'Precheck area
	'-----------------------
	If lgBlnFlgChgValue = False Then
	    IntRetCD = DisplayMsgBox("900001", "X", "X", "X")                          '��: No data changed!!
	    Exit Function
	End If
	'-----------------------
	'Check content area
	'-----------------------
	If Not chkField(Document, "2") Then                             '��: Check contents area
	   Exit Function
	End If

   

	'-----------------------
	'Save function call area
	'-----------------------
	IF  DbSave	= False then 		                                     '��: Save db data 
		Exit Function
	End If    
	    
	FncSave = True                                                          '��: Processing is OK
    
End Function



'========================================================================================

Function FncCopy() 
    On Error Resume Next                                                    '��: Protect system from crashing
   
End Function



'========================================================================================

Function FncCancel() 
    On Error Resume Next                                                    '��: Protect system from crashing
End Function



'========================================================================================

Function FncInsertRow() 
     On Error Resume Next                                                   '��: Protect system from crashing
End Function



'========================================================================================

Function FncDeleteRow() 
    On Error Resume Next                                                    '��: Protect system from crashing
End Function



'========================================================================================

Function FncPrint() 
    On Error Resume Next                                                    '��: Protect system from crashing
    
    parent.FncPrint()
End Function


'========================================================================================

Function FncPrev() 
    On Error Resume Next                                                    '��: Protect system from crashing

End Function


'========================================================================================

Function FncNext() 
    On Error Resume Next                                                    '��: Protect system from crashing

End Function


'========================================================================================

Function FncExcel() 
    Call parent.FncExport(Parent.C_SINGLE)												'��: ȭ�� ���� 
End Function


'========================================================================================

Function FncFind() 
    Call parent.FncFind(Parent.C_SINGLE, False)                                         '��:ȭ�� ����, Tab ���� 
End Function


'========================================================================================

Function FncExit()
	Dim IntRetCD
	FncExit = False
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO, "X", "X")			'��: "Will you destory previous data"
		
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    FncExit = True
End Function


'========================================================================================

Function DbDelete() 
    On Error Resume Next                                                    '��: Protect system from crashing

End Function


'========================================================================================

Function DbDeleteOk()														'��: ���� ������ ���� ���� 
	Call FncNew()
End Function




'========================================================================================
' Function Name : cboXCH_RATE_FG_OnChange
' Function Desc : 
'========================================================================================

Sub txtCashAmt_Change() 
	lgBlnFlgChgValue = True
End Sub

Sub cboCardMM_OnChange() 
	lgBlnFlgChgValue = True
End Sub


Sub cboCardDD_OnChange()
	lgBlnFlgChgValue = True
End Sub


'========================================================================================

Function DbQuery() 
    
    Err.Clear                                                               '��: Protect system from crashing
    DbQuery = False                                                         '��: Processing is NG
    Call LayerShowHide(1)                                                   '��: Protect system from crashing
    

    Dim strVal
    strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001							'��: �����Ͻ� ó�� ASP�� ���� 
    strVal = strVal & "&txtAcctBaseNo=" & Trim(frm1.txtAcctBaseNo.value)				'��: ��ȸ ���� ����Ÿ 

	Call RunMyBizASP(MyBizASP, strVal)										'��: �����Ͻ� ASP �� ���� 

    DbQuery = True                                                          '��: Processing is NG

End Function


'========================================================================================

Function DbQueryOk()														'��: ��ȸ ������ ������� 
	
    '-----------------------
    'Reset variables area
    '-----------------------
    Call SetToolbar("1100100000011111")
   
    lgBlnFlgChgValue = False
    
    Call ggoOper.LockField(Document, "Q")									'��: This function lock the suitable field

    lgIntFlgMode = Parent.OPMD_UMODE												'��: Indicates that current mode is Update mode
End Function



'========================================================================================

Function DbSave() 

    Err.Clear																'��: Protect system from crashing

	DbSave = False															'��: Processing is NG

    Dim strVal
    
    Call LayerShowHide(1)                                                   '��: Protect system from crashing

	With frm1
	
		.txtMode.value = Parent.UID_M0002											'��: �����Ͻ� ó�� ASP �� ���� 
		.txtFlgMode.value     = lgIntFlgMode
	
		Call ExecMyBizASP(frm1, BIZ_PGM_ID)										
	End With

    DbSave = True                                                           '��: Processing is NG
    
End Function


'========================================================================================

Function DbSaveOk()															'��: ���� ������ ���� ���� 
    
    lgBlnFlgChgValue = False
    
    FncQuery

End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->	
</HEAD>

<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="post">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR >
		<TD <%=HEIGHT_TYPE_00%>></TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH="100%">
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><IMG src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23" ></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>�̰�������</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><IMG src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23" ></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=*>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=*>
		<TD WIDTH="100%" CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH="100%">
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS="TD5" NOWRAP>�̰��������</TD>
									<TD CLASS="TD656" NOWRAP><INPUT NAME="txtAcctBaseNo" MAXLENGTH="2" SIZE=10 ALT ="�̰��������" tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCalType" align=top TYPE="BUTTON" onclick="vbscript:Call OpenAcctBaseNo(frm1.txtAcctBaseNo.value,0)">&nbsp;
													<INPUT NAME="txtAcctBaseNm" MAXLENGTH="30" SIZE=30 ALT ="�̰�������ظ�" tag="14X"></TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=100% valign=top>

						<TABLE  <%=LR_SPACE_TYPE_60%>>
							<TR >
								<TD CLASS="TD5" NOWRAP HEIGHT=30>������������</TD>
							    <TD CLASS="TD6" NOWRAP COLSPAN="3"><script language =javascript src='./js/a5401ma1_I492656892_txtCashAmt.js'></script>&nbsp; �� ���� </TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP HEIGHT=30>�ſ�ī����������</TD>
								<TD CLASS=TD6 NOWRAP><SELECT NAME="cboCardMM" ALT="�ſ�ī����������" STYLE="WIDTH: 100px" tag="22"></SELECT> ����<SELECT NAME="cboCardDD" ALT="�ſ�ī����������" STYLE="WIDTH: 100px" tag="22"></SELECT> �ϱ���</TD>
								<TD CLASS=TD6 NOWRAP></TD>
								<TD CLASS=TD6 NOWRAP></TD>
							</TR>
								<TR HEIGHT="*">
									<TD CLASS=TD5></TD>
									<TD CLASS=TD6 COLSPAN="3">&nbsp;</TD>
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
		<TD WIDTH="100%" HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH="100%" HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 tabindex="-1"></IFRAME>	
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=hidden NAME="txtMode" tag="24" tabindex="-1">
<INPUT TYPE=hidden NAME="txtFlgMode" tag="24" tabindex="-1">
<INPUT TYPE=hidden NAME="hAcctBaseNo" tag="24" tabindex="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>

</BODY>
</HTML>

