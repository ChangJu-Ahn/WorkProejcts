<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% session.CodePage=949 %>


<!--'======================================================================================================
'*  1. Module Name          : ACCOUNT
'*  2. Function Name        : 
'*  3. Program ID		    : A6116BA1
'*  4. Program Name         : ������������� ����
'*  5. Program Desc         : ������������� ��ġ
'*  6. Component List       : +
'*  7. Modified date(First) : 2002/10/17
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : namyo, lee
'* 10. Modifier (Last)      : 
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
<!-- #Include file="../../inc/IncSvrHTML.inc"  -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAMain.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAEvent.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliVariables.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAOperation.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliRdsQuery.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incEB.vbs"></SCRIPT>
<SCRIPT LANGUAGE="javascript"	SRC="../../inc/TabScript.js"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit                                                              '��: indicates that All variables must be declared in advance 

'==========================================================================================================

	Dim strYear, strMonth, strDay, dtToday, EndDate 
	EndDate = "<%=GetSvrDate%>"

	Call ExtractDateFrom(EndDate, parent.gServerDateFormat, parent.gServerDateType, strYear, strMonth, strDay)
	EndDate = UniConvYYYYMMDDToDate(parent.gDateFormat, strYear, strMonth, strDay)



Const BIZ_PGM_ID = "a6116bb1.asp"											 '��: �����Ͻ� ���� ASP��

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


 '==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : ���� �ʱ�ȭ(Global ����, �ʱ�ȭ�� �ʿ��� ���� �Ǵ� Flag���� Setting�Ѵ�.)
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

'======================================================================================== 
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% Call loadInfTB19029A("I", "*","NOCOOKIE","BA") %>
End Sub

'========================================================================================================= 
Sub SetDefaultVal()
		Call ExtractDateFrom(parent.gFiscStart, parent.gServerDateFormat, parent.gServerDateType, strYear, strMonth, strDay)
		frm1.txtIssueDT1.year =	strYear
		frm1.txtIssueDT1.Month = strMonth
		Call ExtractDateFrom("<%=GetSvrDate%>", parent.gServerDateFormat, parent.gServerDateType, strYear, strMonth, strDay)
		
		frm1.txtIssueDT2.year = strYear
		frm1.txtIssueDT2.Month = strMonth
		frm1.txtBizAreaCD.value	= ""
		frm1.txtBizAreaNM.value	= ""
		
	lgBlnStartFlag = False

End Sub

'========================================================================================================= 
Function OpenPopUp(Byval strCode, Byval iWhere)
Dim arrRet
Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	Select Case iWhere
		Case 0,1
			arrParam(0) = "���ݽŰ����� �˾�"					' �˾� ��Ī
			arrParam(1) = "B_Tax_BIZ_AREA"	 			' TABLE ��Ī
			arrParam(2) = strCode						' Code Condition
			arrParam(3) = ""							' Name Cindition
			arrParam(4) = ""							' Where Condition
			arrParam(5) = "���ݽŰ������ڵ�"					' �����ʵ��� �� ��Ī

			arrField(0) = "Tax_BIZ_AREA_CD"				' Field��(0)
			arrField(1) = "Tax_BIZ_AREA_NM"				' Field��(0)
    
			arrHeader(0) = "���ݽŰ������ڵ�"					' Header��(0)
			arrHeader(1) = "���ݽŰ������"					' Header��(0)
	End Select
    
	arrRet = window.showModalDialog("../../comasp/commonpopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=480px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Select Case iWhere
			Case 0		' �����
				frm1.txtBizAreaCD.focus
			Case 1		' �����
				frm1.txtBizAreaCD2.focus
		End Select
		Exit Function
	Else
		Call SetPopUp(arrRet, iWhere)
	End If	

End Function

'========================================================================================================= 
Function SetPopUp(Byval arrRet, Byval iWhere)
	With frm1
		Select Case iWhere
			Case 0		' �����
				.txtBizAreaCD.focus
				.txtBizAreaCD.value = UCase(Trim(arrRet(0)))
				.txtBizAreaNM.value = arrRet(1)
			Case 1		' �����
				.txtBizAreaCD2.focus
				.txtBizAreaCD2.value = UCase(Trim(arrRet(0)))
				.txtBizAreaNM2.value = arrRet(1)	
				
		End Select
	End With
End Function


'========================================================================================================= 
Sub Form_Load()
    Call InitVariables							'��: Initializes local global variables
    Call LoadInfTB19029							'��: Load table , B_numeric_format
	  Call AppendNumberPlace("6","16","2")
	  Call AppendNumberPlace("7","16","0")
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")		'��: Lock  Suitable  Field
 	Call ggoOper.FormatNumber(frm1.txtCnt, "9999999", "0", False)	   
	Call ggoOper.FormatDate(frm1.txtIssueDt1, parent.gDateFormat, 2)
	Call ggoOper.FormatDate(frm1.txtIssueDt2, parent.gDateFormat, 2)
	
	'----------  Coding part  -------------------------------------------------------------
	Call SetDefaultVal
	Call SetToolbar("1000000000001111")
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

'   Event Name : txtIssueDt1_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtFromIssueDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtFromIssueDt.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtFromIssueDt.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtStartDt1_Change()
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtFromIssueDt_Change()
    'lgBlnFlgChgValue = True
End Sub

'=======================================================================================================
'   Event Name : txtIssueDt2_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtToIssueDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtToIssueDt.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtToIssueDt.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtStartDt2_Change()
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtToIssueDt_Change()
    'lgBlnFlgChgValue = True
End Sub



 '#########################################################################################################
'												4. Common Function��
'	���: Common Function
'	����: ȯ��ó���Լ�, VAT ó�� �Լ�
'######################################################################################################### 
Function subExportDisk() 
Dim RetFlag
Dim strVal
Dim intRetCD
Dim intI, strFileName, intChrChk	'Ư������ Check
Dim strYear1,strMonth1, strDay1, strDate1
Dim strYear2,strMonth2, strDay2	, strDate2
Dim strMsg
Dim varFromDt
Dim varToDt, varToDt2
Dim varYearMonth
Dim varSingoGubun
Dim varMonthDiff
dim chkYn 
	
	'-----------------------
	'Check content area
	'-----------------------
	
'	If Not chkField(Document, "1") Then        '��: Check contents area
'	  Exit Function
'	End If

	'*************************************************************************
	'//�ʼ��׸� üũ : �ǿ����� üũ�ؾ��� �׸��� �ٸ��⶧���� ����
	'*************************************************************************
	If Trim(frm1.txtIssueDt1.text) = "" Then
		RetFlag = DisplayMsgBox("970029","X" , frm1.txtIssueDt1.Alt, "X") 	
		Exit Function
	End If
	If Trim(frm1.txtIssueDt2.text) = "" Then
		RetFlag = DisplayMsgBox("970029","X" , frm1.txtIssueDt2.Alt, "X") 	
		Exit Function
	End If
		
	If Trim(frm1.txtBizAreaCD.value) = "" Then
		RetFlag = DisplayMsgBox("970029","X" , frm1.txtBizAreaCD.Alt, "X") 	
		Exit Function
	End If
	

			
	frm1.txtFileName.value = ""
		
	If (frm1.txtIssueDt1.year & right(("0" & frm1.txtIssueDt1.Month),2)  > frm1.txtIssueDt2.Year & right(("0" & frm1.txtIssueDt2.Month), 2)) Then
		IntRetCD = DisplayMsgBox("113118", "X", "X", "X")			'��: "Will you destory previous data"
		frm1.txtIssueDt1.focus
		Exit Function
    End If

	
	varFromDt = frm1.txtIssueDt1.Year & "-" & Right(("0" & frm1.txtIssueDt1.Month),2) & "-" & "01"
	VarToDt = FilterVar(frm1.txtIssueDt2.year,"2999","SNM") & "-" & Right("0" & FilterVar(frm1.txtIssueDt2.Month,"12","SNM"),2) & "-" & "01"
	varToDt2 = VarToDt
	VarToDt = DateAdd("D",-1, DateAdd("M",1,cdate(VarToDt)))
	
	'////�Ű����� ���� �ŷ��Ⱓ üũ ///////////////////////
	varMonthDiff = DateDiff("m",Cdate(varFromDt),Cdate(varToDt2)) + 1
	
	Select Case varMonthDiff
			Case 1 
					varSingoGubun = "1"
			Case 2
					varSingoGubun = "2"
			Case 3
					varSingoGubun = "3"
			Case 6
					varSingoGubun = "4"
			Case Else
				intRetCD =  DisplayMsgBox("115115","X" ,"X", "X")
				Exit Function
	End Select
	
	varYearMonth = FilterVar(frm1.txtIssueDt2.year,"2999","SNM") & Right("0" & FilterVar(frm1.txtIssueDt2.Month,"12","SNM"),2) 
	
	varFromDt = UNIDateClientFormat(varFromDt)
	VarToDt = UNIDateClientFormat(VarToDt)
	

	 
	RetFlag = DisplayMsgBox("900018", parent.VB_YES_NO,"x","x")   '�� �ٲ�κ�
	If RetFlag = VBNO Then
		Exit Function
	End IF
	
	if frm1.chkYN(0).checked then 
		chkYn="N"
    else
		chkYn="Y"
    end if
    
	Err.Clear                                                               '��: Protect system from crashing

	With frm1
		Call LayerShowHide(1)
		strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001							'��: �����Ͻ� ó�� ASP�� ���� 
		strVal = strVal & "&txtIssueDt1=" & varFromDt							'��: ��ȸ ���� ����Ÿ
		strVal = strVal & "&txtIssueDt2=" & VarToDt								'��: ��ȸ ���� ����Ÿ
		strVal = strVal & "&txtBizAreaCD=" & UCase(Trim(.txtBizAreaCD.value))	'��: ��ȸ ���� ����Ÿ
		strVal = strVal & "&txtReportDt=" & VarToDt								'��: ��ȸ ���� ����Ÿ
		strVal = strVal & "&txtFileName=" & Trim(.txtFileName.value)			'��: ��ȸ ���� ����Ÿ
		strVal = strVal & "&txtYearMonth=" & varYearMonth						'��: ��ȸ ���� ����Ÿ
		strVal = strVal & "&txtSingoGubun=" & varSingoGubun						'��: ��ȸ ���� ����Ÿ	
		strVal = strVal & "&txtCnt=" & frm1.txtCnt.text						'��: ��ȸ ���� ����Ÿ
		strVal = strVal & "&txtDocAmt=" & frm1.txtDocAmt.text						'��: ��ȸ ���� ����Ÿ
		strVal = strVal & "&txtLocAmt=" & frm1.txtLocAmt.text						'��: ��ȸ ���� ����Ÿ	
	    strVal = strVal & "&chkYn=" & chkYn
	
		Call RunMyBizASP(MyBizASP, strVal)	
												'��: �����Ͻ� ASP �� ����
	End With

End Function

Function subExportDiskOK(ByVal pFileName) 
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
	Call parent.FncPrint()
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
    Call parent.FncFind(parent.C_SINGLE, True)                                         '��:ȭ�� ����, Tab ���� 
End Function


'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================

Function FncExit()
    FncExit = True
End Function


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
<!-- '#########################################################################################################
'       					6. Tag��
'#########################################################################################################  -->

<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE CLASS="BatchTB1" CELLSPACING=0 CELLPADDING=0>
	<TR>
		<TD <%=HEIGHT_TYPE_00%>><!-- ' ���� ���� --></TD>
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>��������������ϻ���</font></td>
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
								<TD CLASS=TD5>&nbsp;</TD>
								<TD CLASS=TD6>&nbsp;</TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>�ŷ��Ⱓ</TD>
								<TD CLASS=TD6><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime2 name=txtIssueDt1 CLASS=FPDTYYYYMM title=FPDATETIME ALT="�ŷ�������" tag="12X1" VIEWASTEXT></OBJECT>');</SCRIPT>
											  &nbsp; ~ &nbsp;
											  <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime2 name=txtIssueDt2 CLASS=FPDTYYYYMM title=FPDATETIME ALT="�ŷ�������" tag="12X1" VIEWASTEXT></OBJECT>');</SCRIPT></TD>
							</TR>
							<TR>
									<TD CLASS=TD5 NOWRAP>���հ�������</TD>
									<TD CLASS=TD6>
									<INPUT TYPE="RADIO" CLASS="RADIO" NAME="chkYN" TAG="1X" VALUE="N" CHECKED ID="chkYN0"><LABEL FOR="chkYN0">����庰</LABEL>&nbsp;
				        	        <INPUT TYPE="RADIO" CLASS="RADIO" NAME="chkYN" TAG="1X" VALUE="Y"  ID="chkYN1"><LABEL FOR="chkYN1">����</LABEL>
				
									 </TD>
								</TR>
								
							<TR>
								<TD CLASS=TD5 NOWRAP>���ݽŰ�����</TD>
								<TD CLASS=TD6><INPUT TYPE=TEXT ID="txtBizAreaCD" NAME="txtBizAreaCD" SIZE=10 MAXLENGTH=10 STYLE="TEXT-ALIGN: left" tag="12XXXU" ALT="�Ű�����"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btn" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtBizAreaCD.Value, 0)">&nbsp;<INPUT TYPE=TEXT ID="txtBizAreaNM" NAME="txtBizAreaNM" SIZE=30 MAXLENGTH=50 STYLE="TEXT-ALIGN: left" tag="14X" ALT="���ݽŰ�����"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>ȭ�ϸ�</TD>
								<TD CLASS=TD6><INPUT TYPE=TEXT ID="txtFileName" NAME="txtFileName" SIZE=20 MAXLENGTH=20 STYLE="TEXT-ALIGN: left" tag="14X" ALT="ȭ�ϸ�"></TD>
							</TR>
							<TR>
							 	<TD CLASS="TD5">�������Ǽ�</TD>
								<TD CLASS="TD6"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle1 NAME="txtCnt" style="HEIGHT: 20px; WIDTH: 100px" tag="11X" ALT="�������Ǽ�" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT></TD>
							</TR>
							<TR>
							 	<TD CLASS="TD5">��������ȭ�ݾ�</TD>
								<TD CLASS="TD6"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle2 name=txtDocAmt style="HEIGHT: 20px; WIDTH: 150px"CLASS=FPDS115 title=FPDOUBLESINGLE tag="11X6X" ALT="��������ȭ�ݾ�"></OBJECT>');</SCRIPT></TD>
							</TR>
							<TR>
							 	<TD CLASS="TD5">��������ȭ�ݾ�</TD>
								<TD CLASS="TD6"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle3 name=txtLocAmt style="HEIGHT: 20px; WIDTH: 150px" CLASS=FPDS115 title=FPDOUBLESINGLE tag="11X7X" ALT="��������ȭ�ݾ�" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT></TD>
							</TR>
							<!--
							<TR>
								<TD CLASS=TD5 NOWRAP>�ϰ��븮����</TD>
								<TD CLASS="TD6"><input type="checkbox" class = "check" name="chkDari" value="Y"></TD>
							</TR>	-->
							<TR>
								<TD CLASS=TD5></TD>
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
					<TD>
					<BUTTON NAME="btnExecute" CLASS="CLSMBTN" OnClick="VBScript:Call subExportDisk()" Flag=1>�� ��</BUTTON>&nbsp;
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>>
			<IFRAME  NAME="MyBizASP" WIDTH=100% HEIGHT=0 FRAMEBORDER=0 SCROLLING=no noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm" TABINDEX="-1"></iframe>
</DIV>
</FORM>
</BODY>
</HTML>

