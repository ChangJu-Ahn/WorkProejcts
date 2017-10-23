<%@ LANGUAGE="VBSCRIPT" %>
<!--'======================================================================================================
'*  1. Module Name          : Accounting
'*  2. Function Name        : Treasury - Loan
'*  3. Program ID           : f4220ba1
'*  4. Program Name         : ���Աݻ�ȯ����(posam)
'*  5. Program Desc         : Create of Loan Schedule
'*  6. Comproxy List        : PAFG400
'*  7. Modified date(First) : 2002.04.12
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Oh, Soo min
'* 10. Modifier (Last)      : 2003.05.19
'* 11. Comment              :
'======================================================================================================= -->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- '#########################################################################################################
'												1. �� �� �� 
'##############################################################################################################
'******************************************  1.1 Inc ����   **************************************************
'	���: Inc. Include
'*********************************************************************************************************  -->
<!-- #Include file="../../inc/incSvrCcm.inc"  -->
<!-- #Include file="../../inc/incSvrHTML.inc"  -->		<!-- '��: ȭ��ó��ASP���� �����۾��� �ʿ��� ���  -->

<!--'==========================================  1.1.1 Style Sheet  ======================================
'========================================================================================================== -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliMAMain.vbs">  </SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliMAEvent.vbs">  </SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliVariables.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliMAOperation.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit  
                                                            '��: indicates that All variables must be declared in advance 
 '#########################################################################################################
'												1. �� �� �� 
'##########################################################################################################

 '******************************************  1.1 Inc ����   **********************************************
'	���: Inc. Include
'********************************************************************************************************* 

'******************************************  1.2 Global ����/��� ����  ***********************************
'	1. Constant�� �ݵ�� �빮�� ǥ��.
'<!-- #Include file="../../inc/lgvariables.inc" -->
'**********************************************************************************************************
'==========================================  1.2.1 Global ��� ����  ======================================
'==========================================================================================================

Const BIZ_PGM_ID = "f4220bb1_ko441.asp"  
 '==========================================  1.2.3 Global Variable�� ����  ===============================
'========================================================================================================= 

 '----------------  ���� Global ������ ����  ----------------------------------------------------------- 

 '++++++++++++++++  Insert Your Code for Global Variables Assign  ++++++++++++++++++++++++++++++++++++++ 
Dim lgMpsFirmDate, lgLlcGivenDt
Dim lgIsOpenPop
Dim lgBlnFlgChgValue														   '��: Variable is for Dirty flag
Dim lgCurName()															'�� : ���� ȭ��� �ʿ��� ��Į ���� ���� 
Dim IsOpenPop          
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
'    lgIntFlgMode = OPMD_CMODE                                               '��: Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                                                '��: Indicates that no value changed
'    lgIntGrpCount = 0                                                       '��: Initializes Group View Size
    '----------  Coding part  -------------------------------------------------------------
    IsOpenPop = False														'��: ����� ���� �ʱ�ȭ 
'    lgCboKeyPress = False
End Sub

'=============================== 2.1.2 LoadInfTB19029() ========================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== 
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% Call loadInfTB19029A("I", "*", "NOCOOKIE", "BA") %>
'<% Call LoadInfTB19029(gCurrency, "I", "*") %>
End Sub

 '******************************************  2.1 Pop-Up �Լ�   **********************************************
'	���: Pop-Up 
'********************************************************************************************************* 

 '------------------------------------------  OpenPopUp()  -------------------------------------------------
'	Name : OpenPopUp()
'	Description : Data Code PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenPopupLoan()

	Dim iCalledAspName
	Dim arrRet
	Dim arrParam(3)	

	If IsOpenPop = True Then Exit Function
	iCalledAspName = AskPRAspName("f4232ra1")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "f4232ra1", "X")
		IsOpenPop = False
		Exit Function
	End If
	

	IsOpenPop = True
    
	arrRet = window.showModalDialog(iCalledAspName & "?PGM=" & gStrRequestMenuID , Array(window.parent,arrParam), _
		     "dialogWidth=850px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = ""  Then			
		frm1.txtLoanNo.focus
		Exit Function
	Else		
		frm1.txtLoanNo.value = arrRet(0)
		frm1.txtLoanNm.value = arrRet(1)
	End If
	
	frm1.txtLoanNo.focus

End Function

 '++++++++++++++++  Insert Your Code for PopUp(Open)  ++++++++++++++++++++++++++++++++++++++++++++++++++ 

Sub SetDefaultVal()
	Dim strSvrDate
	DIm strYear, strMonth, strDay
	Dim frDt, toDt
	
	strSvrDate = "<%=GetSvrDate%>"
	Call ExtractDateFrom(strSvrDate, parent.gServerDateFormat, parent.gServerDateType, strYear,strMonth,strDay)
		
	frDt = UniConvYYYYMMDDToDate(parent.gDateFormat, strYear, strMonth, "01")
	toDt = UniConvYYYYMMDDToDate(parent.gDateFormat, strYear, strMonth, strDay)
	
	frm1.txtDateFr.Text = frDt    
	frm1.txtDateTo.Text = toDt
End Sub

 '------------------------------------------  fnButtonExec()  --------------------------------------------------
'	Name : fnButtonExec()
'	Description : ���� ��ư Ŭ�� �� ����. 
'--------------------------------------------------------------------------------------------------------- 

Function fnButtonExec()
    Dim strVal           
    Dim strFrdt
    Dim strTodt
    Dim strTarget
	Dim IntRetCD
        
    If Not chkField(Document, "1") Then									'��: This function check indispensable field
		Exit Function
    End If
	
	'--------------------------------------------------------------
	'������ ���� ���� Check
	'--------------------------------------------------------------

	If CompareDateByFormat(frm1.txtDateFr.Text, frm1.txtDateTo.Text, frm1.txtDateFr.Alt, frm1.txtDateTo.Alt, _
						"970025", frm1.txtDateFr.UserDefinedFormat, parent.gComDateType, true) = False Then
			frm1.txtDateFr.focus											'��: GL Date Compare Common Function
			Exit Function
	End if
	
	IntRetCD = DisplayMsgBox("900018",parent.VB_YES_NO,"X","X")
	
	If IntRetCD = vbNo Then
		Exit Function
	End If
	
    Call LayerShowHide(1)
    
    strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0002							'��: �����Ͻ� ó�� ASP�� ����    
       
    strVal = strVal & "&txtDateFr=" & Trim(frm1.txtDateFr.text)  
    strVal = strVal & "&txtDateTo=" & Trim(frm1.txtDateTo.text)
    strVal = strVal & "&txtLoanNo=" & Trim(frm1.txtLoanNo.value)
    strVal = strVal & "&txtBizAreaCd=" & Trim(frm1.txtBizAreaCd.value)
    strVal = strVal & "&txtBizAreaCd1=" & Trim(frm1.txtBizAreaCd1.value)
	Call RunMyBizASP(MyBizASP, strVal)										'��: �����Ͻ� ASP �� ���� 
	    
End Function

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

    Call InitVariables																'��: Initializes local global variables
    Call LoadInfTB19029                           '��: Load table , B_numeric_format

	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)

    Call ggoOper.FormatDate(frm1.txtDateFr, parent.gDateFormat, 1)		'����� 
    Call ggoOper.FormatDate(frm1.txtDateTo, parent.gDateFormat, 1)		'����� 

    Call SetDefaultVal

    Call ggoOper.LockField(Document, "N")									'��: Lock  Suitable  Field
    Call SetToolbar("1000000000001111")
    
    frm1.fpDateFr.focus 
End Sub

'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================

Sub Form_QueryUnload(Cancel , UnloadMode )
End Sub

'=======================================================================================================
'   Event Name : txt1StPrRdpDt_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtDateFr_DblClick(Button)
    If Button = 1 Then
        frm1.txtDateFr.Action = 7
        Call SetFocusToDocument("M")
        frm1.txtDateFr.Focus
    End If
    lgBlnFlgChgValue = True
End Sub

'=======================================================================================================
'   Event Name : txtDateTo_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtDateTo_DblClick(Button)
    If Button = 1 Then
        frm1.txtDateTo.Action = 7
        Call SetFocusToDocument("M")
        frm1.txtDateTo.Focus
    End If
    lgBlnFlgChgValue = True
End Sub

'----------------------------------------  OpenBizAreaCd()  -------------------------------------------------
'	Name : BizAreaCd()
'	Description : Business PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenBizAreaCd()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	arrParam(0) = "����� �˾�"				' �˾� ��Ī 
	arrParam(1) = "B_BIZ_AREA"					' TABLE ��Ī 
	arrParam(2) = Trim(frm1.txtBizAreaCd.Value)	' Code Condition
	arrParam(3) = ""							' Name Cindition
	arrParam(4) = ""							' Where Condition
	arrParam(5) = "����� �ڵ�"			

    arrField(0) = "BIZ_AREA_CD"					' Field��(0)
    arrField(1) = "BIZ_AREA_NM"					' Field��(1)

    arrHeader(0) = "������ڵ�"				' Header��(0)
	arrHeader(1) = "������"				' Header��(1)

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If arrRet(0) = "" Then
'		frm1.txtBizAreaCd.focus
		Exit Function
	Else
		frm1.txtBizAreaCd.Value	= arrRet(0)
		frm1.txtBizAreaNm.Value	= arrRet(1)
	End If
End Function


'----------------------------------------  OpenBizAreaCd()  -------------------------------------------------
'	Name : BizAreaCd()
'	Description : Business PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenBizAreaCd1()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	arrParam(0) = "����� �˾�"				' �˾� ��Ī 
	arrParam(1) = "B_BIZ_AREA"					' TABLE ��Ī 
	arrParam(2) = Trim(frm1.txtBizAreaCd1.Value)	' Code Condition
	arrParam(3) = ""							' Name Cindition
	arrParam(4) = ""							' Where Condition
	arrParam(5) = "����� �ڵ�"			

    arrField(0) = "BIZ_AREA_CD"					' Field��(0)
    arrField(1) = "BIZ_AREA_NM"					' Field��(1)

    arrHeader(0) = "������ڵ�"				' Header��(0)
	arrHeader(1) = "������"				' Header��(1)

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If arrRet(0) = "" Then
'		frm1.txtBizAreaCd.focus
		Exit Function
	Else
		frm1.txtBizAreaCd1.Value = arrRet(0)
		frm1.txtBizAreaNm1.Value = arrRet(1)
	End If
End Function


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
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================

Function FncPrint() 
    Call Parent.FncPrint()
End Function

'======================================================================================================
' Function Name : FncFind
' Function Desc : 
'=======================================================================================================
Function FncFind() 
	Call parent.FncFind(parent.C_SINGLE, False)                                         '��:ȭ�� ����, Tab ���� 
End Function

'========================================================================================
' Function Name : FncQuery
' Function Desc : 
'========================================================================================
Function FncQuery()
    
End Function

'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================
Function FncExit()
    FncExit = True
End Function

</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc"  -->	
</HEAD>
<!-- '#########################################################################################################
'       					6. Tag�� 
'#########################################################################################################  -->

<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE CLASS="BatchTB2" CELLSPACING=0 CELLPADDING=0>
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" ><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>���Աݻ�ȯ����</font></td>
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
					<TD>
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD CLASS="TD5" NOWRAP>��������</TD>               
								<TD CLASS="TD6" NOWRAP><script language =javascript src='./js/f4220ba1_fpDateFr_txtDateFr.js'></script>&nbsp;~&nbsp;
													   <script language =javascript src='./js/f4220ba1_fpDateTo_txtDateTo.js'></script>
								</TD>
							</TR>					
							<TR>
								<TD CLASS="TD5" NOWRAP>���Աݹ�ȣ</TD>
								<TD CLASS="TD6" NOWRAP><INPUT NAME="txtLoanNo" ALT="���Աݹ�ȣ" SIZE=15 MAXLENGTH="18"  tag="1XXXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnLoanNo" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPopUpLoan()">&nbsp;<INPUT NAME="txtLoanNm" ALT="���Աݳ���" tag="14"></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>�����</TD>
								<TD CLASS="TD6" NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtBizAreaCd" SIZE=13 MAXLENGTH=10 STYLE="TEXT-ALIGN: Left" tag="11XXXU" ALT="�����"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBizAreaCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenBizAreaCd()">
								                       <INPUT TYPE=TEXT NAME="txtBizAreaNm" SIZE=25 tag="14">&nbsp;~</TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP></TD>
								<TD CLASS="TD6" NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtBizAreaCd1" SIZE=13 MAXLENGTH=10 STYLE="TEXT-ALIGN: Left" tag="11XXXU" ALT="�����"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBizAreaCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenBizAreaCd1()">
									                      <INPUT TYPE=TEXT NAME="txtBizAreaNm1" SIZE=25 tag="14"></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>ó���Ǽ�</TD>
								<TD CLASS="TD6" NOWRAP><INPUT NAME="txtCount" ALT="ó���Ǽ�" SIZE=15 STYLE="TEXT-ALIGN:right" tag="14X"></TD>
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
					<TD><BUTTON NAME="btnExec" CLASS="CLSMBTN" OnClick="VBScript:Call fnButtonExec()" Flag=1>����</BUTTON></TD>
					<TD WIDTH=*>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% FRAMEBORDER=0 SCROLLING=no noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</FORM>
</BODY>
</HTML>
