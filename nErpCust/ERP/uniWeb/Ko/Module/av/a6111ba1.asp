<%@ LANGUAGE="VBSCRIPT" %>
<!--'======================================================================================================
'*  1. Module Name          : ȸ����� 
'*  2. Function Name        : �ΰ������� 
'*  3. Program ID		    : A6111MA1
'*  4. Program Name         : �ΰ��������е��ϻ��� 
'*  5. Program Desc         : �ΰ��������е��ϻ��� 
'*  6. Component List       : +
'*  7. Modified date(First) : 2000/04/20
'*  8. Modified date(Last)  : 2002/09/11
'*  9. Modifier (First)     : Jong Hwan, Kim
'* 10. Modifier (Last)      : Hye young ,Lee
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'======================================================================================================= -->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/IncSvrCcm.inc"  -->
<!-- #Include file="../../inc/incSvrHTML.inc"  -->

<!--'==========================================  1.1.1 Style Sheet  ======================================
'========================================================================================================== -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'==========================================  1.1.2 ���� Include   ======================================
'========================================================================================================== -->

<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAMain.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAEvent.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliVariables.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAOperation.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/Cookie.vbs">	</SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit                                                              '��: indicates that All variables must be declared in advance 

'==========================================================================================================
Dim strYear, strMonth, strDay, dtToday, EndDate, StartDate
EndDate     =   "<%=GetSvrDate%>"

Call ExtractDateFrom(EndDate, parent.gServerDateFormat, parent.gServerDateType, strYear, strMonth, strDay)
StartDate   = UniConvYYYYMMDDToDate(parent.gDateFormat, strYear, strMonth, "01")
EndDate     = UniConvYYYYMMDDToDate(parent.gDateFormat, strYear, strMonth, strDay)


Const BIZ_PGM_ID = "a6111bb1.asp"											 '��: �����Ͻ� ���� ASP�� 
 '==========================================  1.2.2 Global ���� ����  =====================================
'	1. ���� ǥ�ؿ� ����. prefix�� g�� �����.
'	2.Array�� ���� ()�� �ݵ�� ����Ͽ� �Ϲ� ������ ������ �� 
'========================================================================================================= 
Dim lgBlnFlgConChg				'��: Condition ���� Flag
Dim lgBlnFlgChgValue				'��: Variable is for Dirty flag
Dim lgIntGrpCount				'��: Group View Size�� ������ ���� 
Dim lgIntFlgMode					'��: Variable is for Operation Status

Dim lgNextNo						'��: ȭ���� Single/SingleMulti �ΰ�츸 �ش� 
Dim lgPrevNo						' ""

Dim lgBlnStartFlag				' �޼��� �����Ͽ� ���α׷� ���۽��� Check Flag

'========================================================================================================= 
Dim lgMpsFirmDate, lgLlcGivenDt	 '��: �����Ͻ� ���� ASP���� �����ϹǷ� 

Dim  lgCurName()					'�� : ���� ȭ��� �ʿ��� ��Į ���� ���� 
Dim  cboOldVal          
Dim  IsOpenPop          



'========================================================================================================= 
Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE   '��: Indicates that current mode is Create mode
    lgBlnFlgChgValue = False    '��: Indicates that no value changed
    lgIntGrpCount = 0           '��: Initializes Group View Size
    '----------  Coding part  -------------------------------------------------------------
    IsOpenPop = False			'��: ����� ���� �ʱ�ȭ 
    lgMpsFirmDate=""
    lgLlcGivenDt=""
End Sub

'=============================== 2.1.2 LoadInfTB19029() ========================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== 
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% Call loadInfTB19029A("I", "A","NOCOOKIE","MA") %>
End Sub

'========================================================================================================= 
Sub SetDefaultVal()
	
	frm1.txtIssueDT1.Text = StartDate
	frm1.txtIssueDT2.Text = EndDate
	frm1.txtReportDt.Text = EndDate
	frm1.txtBizAreaCD.focus 
	
    'frm1.txtIssueDt1.focus
    'frm1.btnExecute.disabled = True
    
    'frm1.txtBizAreaCD.value	= parent.gBizArea

	lgBlnStartFlag = False
End Sub

 '==========================================  2.2.6 InitComboBox()  =======================================
'	Name : InitComboBox()
'	Description : Combo Display
'========================================================================================================= 
Sub InitComboBox()
End Sub


'------------------------------------------  OpenPopUp()  -------------------------------------------------
'	Name : OpenPopUp()
'	Description : PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenPopUp(Byval strCode, Byval iWhere)
Dim arrRet
Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	Select Case iWhere
		Case 0
			arrParam(0) = "���ݽŰ����� �˾�"					' �˾� ��Ī 
			arrParam(1) = "B_TAX_BIZ_AREA"	 			' TABLE ��Ī 
			arrParam(2) = strCode						' Code Condition
			arrParam(3) = ""							' Name Cindition
			arrParam(4) = ""							' Where Condition
			arrParam(5) = "���ݽŰ������ڵ�"					' �����ʵ��� �� ��Ī 

			arrField(0) = "TAX_BIZ_AREA_CD"				' Field��(0)
			arrField(1) = "TAX_BIZ_AREA_NM"				' Field��(0)
    
			arrHeader(0) = "���ݽŰ������ڵ�"					' Header��(0)
			arrHeader(1) = "���ݽŰ������"					' Header��(0)
	End Select
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtBizAreaCD.focus
		Exit Function
	Else
		Call SetPopUp(arrRet, iWhere)
	End If	

End Function

'------------------------------------------  SetPopUp()  --------------------------------------------------
'	Name : SetPopUp()
'	Description : CtrlItem Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 

Function SetPopUp(Byval arrRet, Byval iWhere)
	With frm1
		Select Case iWhere
			Case 0		' ����� 
				.txtBizAreaCD.focus
				.txtBizAreaCD.value = UCase(Trim(arrRet(0)))
				.txtBizAreaNM.value = arrRet(1)
		End Select
	End With
End Function

'========================================================================================================= 
Sub Form_Load()

    Call InitVariables							'��: Initializes local global variables
    Call LoadInfTB19029							'��: Load table , B_numeric_format

    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")		'��: Lock  Suitable  Field
    '----------  Coding part  -------------------------------------------------------------
    Call SetDefaultVal

    Call SetToolbar("1000000000000001")										'��: ��ư ���� ���� 
	frm1.txtBizAreaCD.focus 
End Sub

'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================

Sub Form_QueryUnload(Cancel , UnloadMode )

End Sub

'=======================================================================================================
'   Event Name : txtIssueDt1_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtIssueDt1_DblClick(Button)
    If Button = 1 Then
        frm1.txtIssueDt1.Action = 7
  		Call SetFocusToDocument("M")
		frm1.txtIssueDt1.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtStartDt1_Change()
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtIssueDt1_Change()
    'lgBlnFlgChgValue = True
End Sub

'=======================================================================================================
'   Event Name : txtIssueDt2_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtIssueDt2_DblClick(Button)
    If Button = 1 Then
        frm1.txtIssueDt2.Action = 7
  		Call SetFocusToDocument("M")
		frm1.txtIssueDt2.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtStartDt2_Change()
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtIssueDt2_Change()
    'lgBlnFlgChgValue = True
End Sub

'=======================================================================================================
'   Event Name : txtReportDt_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtReportDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtReportDt.Action = 7
  		Call SetFocusToDocument("M")
		frm1.txtReportDt.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtReportDt_Change()
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtReportDt_Change()
    'lgBlnFlgChgValue = True
End Sub

 '#########################################################################################################
'												4. Common Function�� 
'	���: Common Function
'	����: ȯ��ó���Լ�, VAT ó�� �Լ� 
'######################################################################################################### 
Function subVatDisk() 
Dim RetFlag
Dim strVal
Dim IntRetCD
Dim intI, strFileName, intChrChk	'Ư������ Check

	'-----------------------
    'Check content area
    '-----------------------
    
    'ȭ�ϸ����� ����� �� ���� Ư������ \/:*?"<>|&. ���Կ��� Ȯ�� 
	strFileName = frm1.txtFileName.value
	
	For intI = 1 To Len(strFileName)
		intChrChk = ASC(Mid(strFileName, intI, 1))
		If intChrChk = ASC("\") Or intChrChk = ASC("/") Or intChrChk = ASC(":") Or intChrChk = ASC("*") Or _
			intChrChk = ASC("?") Or intChrChk = 34 Or intChrChk = ASC("<") Or intChrChk = ASC(">") Or _
			intChrChk = ASC("|") OR intChrChk = ASC("&") OR intChrChk = ASC(".") Then
				intRetCD =  DisplayMsgBox("970029","X" , frm1.txtFileName.Alt, frm1.txtIssueDt2.Alt)
				Exit Function
		End If
	Next
	
	' Required�� ǥ�õ� Element���� �Է� [��/��]�� Check �Ѵ�.
	' ChkField(pDoc, pStrGrp) As Boolean
    If Not chkField(Document, "1") Then        '��: Check contents area
       Exit Function
    End If

    If CompareDateByFormat(frm1.txtIssueDt1.text,frm1.txtIssueDt2.text,frm1.txtIssueDt1.Alt,frm1.txtIssueDt2.Alt, _
        	               "970025",frm1.txtIssueDt1.UserDefinedFormat,parent.gComDateType, true) = False Then
	   frm1.txtIssueDt1.focus
	   Exit Function
	End If
    

	RetFlag = DisplayMsgBox("900018", parent.VB_YES_NO,"x","x")   '�� �ٲ�κ� 
	'RetFlag = Msgbox("�۾��� ���� �Ͻðڽ��ϱ�?", vbOKOnly + vbInformation, "����")
	If RetFlag = VBNO Then
		Exit Function
	End IF

    Err.Clear                                                               '��: Protect system from crashing

    With frm1

		Call LayerShowHide(1)
	
	    strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001							'��: �����Ͻ� ó�� ASP�� ���� 
		strVal = strVal & "&txtIssueDt1=" & Trim(.txtIssueDt1.text)		'��: ��ȸ ���� ����Ÿ 
		strVal = strVal & "&txtIssueDt2=" & Trim(.txtIssueDt2.text)		'��: ��ȸ ���� ����Ÿ 
		strVal = strVal & "&txtBizAreaCD=" & UCase(Trim(.txtBizAreaCD.value))	'��: ��ȸ ���� ����Ÿ 
		strVal = strVal & "&txtReportDt=" & Trim(.txtReportDt.text)		'��: ��ȸ ���� ����Ÿ 
		strVal = strVal & "&txtFileName=" & Trim(.txtFileName.value)			'��: ��ȸ ���� ����Ÿ 

		Call RunMyBizASP(MyBizASP, strVal)										'��: �����Ͻ� ASP �� ���� 
	
    End With
    
End Function

Function subVatDiskOK(ByVal pFileName) 
Dim strVal
    Err.Clear                                                               '��: Protect system from crashing

	strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0002							'��: �����Ͻ� ó�� ASP�� ���� 
	strVal = strVal & "&txtFileName=" & pFileName							'��: ��ȸ ���� ����Ÿ 

	Call RunMyBizASP(MyBizASP, strVal)										'��: �����Ͻ� ASP �� ���� 
End Function



'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================

Function FncQuery() 
End Function


'========================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'========================================================================================

Function FncNew() 
End Function


'========================================================================================
' Function Name : FncDelete
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================

Function FncDelete() 
End Function


'========================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================

Function FncSave() 
End Function


'========================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================

Function FncCopy() 
    On Error Resume Next                                                    '��: Protect system from crashing
End Function


'========================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================

Function FncCancel() 
    On Error Resume Next                                                    '��: Protect system from crashing
End Function


'========================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================

Function FncInsertRow() 
     On Error Resume Next                                                   '��: Protect system from crashing
End Function


'========================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================

Function FncDeleteRow() 
    On Error Resume Next                                                    '��: Protect system from crashing
End Function


'========================================================================================
' Function Name : FncPrev
' Function Desc : This function is related to Previous Button
'========================================================================================

Function FncPrev() 
End Function


'========================================================================================
' Function Name : FncNext
' Function Desc : This function is related to Next Button
'========================================================================================

Function FncNext() 
End Function


'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================

Function FncPrint() 
End Function


'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================

Function FncExcel() 
End Function


'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================

Function FncFind() 
End Function


'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================

Function FncExit()
    FncExit = True
End Function

 '*******************************  5.2 Fnc�Լ����� ȣ��Ǵ� ���� Function  *******************************
'	���� : 
'********************************************************************************************************* 


'========================================================================================
' Function Name : DbDelete
' Function Desc : This function delete data
'========================================================================================

Function DbDelete() 
End Function


'========================================================================================
' Function Name : DbDeleteOk
' Function Desc : DbDelete�� �������϶� ���� 
'========================================================================================

Function DbDeleteOk()														'��: ���� ������ ���� ���� 
End Function


'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================

Function DbQuery() 
End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================

Function DbQueryOk()							'��: ��ȸ ������ ������� 
End Function


'========================================================================================
' Function Name : DBSave
' Function Desc : ���� ���� ������ ���� , �������̸� DBSaveOk ȣ��� 
'========================================================================================

Function DbSave() 
End Function

'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncSave�� �ִ°��� �ű� 
'========================================================================================

Function DbSaveOk()			'��: ���� ������ ���� ���� 
End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->
</HEAD>



<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE CLASS="BatchTB1"  CELLSPACING=0 CELLPADDING=0>
	<TR>
		<TD <%=HEIGHT_TYPE_00%>><!-- ' ���� ���� --></TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSLTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSLTABP"><font color=white>�ΰ��������е��ϻ���</font></td>
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
								<TD CLASS=TD5>&nbsp;</TD>
								<TD CLASS=TD6>&nbsp;</TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>��꼭������</TD>
								<TD CLASS=TD6><script language =javascript src='./js/a6111ba1_fpDateTime2_txtIssueDt1.js'></script>
											  &nbsp; ~ &nbsp;
											  <script language =javascript src='./js/a6111ba1_fpDateTime2_txtIssueDt2.js'></script></TD>
							</TR>
							<TR>
								<TD CLASS=TD5>&nbsp;</TD>
								<TD CLASS=TD6>&nbsp;</TD>
							</TR>
								<TR>
								<TD CLASS=TD5 NOWRAP>���ݽŰ�����</TD>
								<TD CLASS=TD6><INPUT TYPE=TEXT ID="txtBizAreaCD" NAME="txtBizAreaCD" SIZE=10 MAXLENGTH=10 STYLE="TEXT-ALIGN: left" tag="12XXXU" ALT="���ݽŰ�����"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btn" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtBizAreaCD.Value, 0)">&nbsp;
												<INPUT TYPE=TEXT ID="txtBizAreaNM" NAME="txtBizAreaNM" SIZE=30 MAXLENGTH=50 STYLE="TEXT-ALIGN: left" tag="14X" ALT="���ݽŰ�����"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5>&nbsp;</TD>
								<TD CLASS=TD6>&nbsp;</TD>
							</TR>		
							<TR>
								<TD CLASS=TD5 NOWRAP>�Ű�����</TD>
								<TD CLASS=TD6><script language =javascript src='./js/a6111ba1_fpDateTime2_txtReportDt.js'></script></TD>
							</TR>
							<TR>
								<TD CLASS=TD5>&nbsp;</TD>
								<TD CLASS=TD6>&nbsp;</TD>
							</TR>		
							<TR>
								<TD CLASS=TD5 NOWRAP>ȭ�ϸ�</TD>
								<TD CLASS=TD6><INPUT TYPE=TEXT ID="txtFileName" NAME="txtFileName" SIZE=30 MAXLENGTH=30 STYLE="TEXT-ALIGN: left" tag="12X" ALT="ȭ�ϸ�"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5>&nbsp;</TD>
								<TD CLASS=TD6>&nbsp;</TD>
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
					<TD><BUTTON NAME="btnExecute" CLASS="CLSMBTN" OnClick="VBScript:Call subVatDisk()" Flag=1>�� ��</BUTTON></TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>>
			<IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm" TABINDEX="-1"></iframe>
</DIV>
</FORM>
</BODY>
</HTML>

