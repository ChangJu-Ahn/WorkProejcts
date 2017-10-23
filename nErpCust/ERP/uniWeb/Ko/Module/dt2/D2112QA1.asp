<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Purchase
'*  2. Function Name        : 
'*  3. Program ID           : m5134ma1
'*  4. Program Name         : ���ڰ�꼭 ����(����) 
'*  5. Program Desc         : ���ڰ�꼭�� ���Ͽ� ���� �Ǵ� ��������ϴ� ��� 
'*  6. Component List       : PAGG015.dll
'*  7. Modified date(First) : 2000/10/14
'*  8. Modified date(Last)  : 2003/10/31
'*  9. Modifier (First)     : Lee MIn Hyung
'* 10. Modifier (Last)      : Lee Min HYung
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change" 
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'**********************************************************************************************
-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/IncServer.asp" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentA.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentVariables.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit  

<!-- #Include file="../../inc/lgvariables.inc" -->	

Dim iDBSYSDate

iDBSYSDate = "<%=GetSvrDate%>"

'======================================================================================== 
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% Call loadInfTB19029A("Q", "A","NOCOOKIE","MA") %>
<% Call LoadBNumericFormatA("Q", "A", "NOCOOKIE", "MA") %>
End Sub

'########################################################################################################
'#                       4.  Data Declaration Part
'########################################################################################################

'=                       4.1 External ASP File
'========================================================================================================
Const BIZ_PGM_ID  = "D2112QB1.asp"
'==========================================  1.2.1 Global ��� ����  ======================================
'=                       4.2 Constant variables 
'========================================================================================================

Dim IsOpenPop          
Dim lgOldRow, lgRow
Dim lgSortKey1
Dim lgSortKey2

'add header datatable column
Dim C_inv_type
Dim C_inv_no
Dim C_dt_inv_no
Dim C_process_date
Dim C_success_flag
Dim	C_re_flag
Dim	C_change_reason
Dim	C_change_remark
Dim	C_change_remark2
Dim	C_change_remark3
Dim C_error_desc
Dim C_insert_user_id
Dim	C_user_name

'########################################################################################################
'#                       5.Method Declaration Part
'########################################################################################################

'                        5.1 Common Method-1
'========================================================================================================= 
'========================================================================================================= 
Sub Form_Load()
    Call LoadInfTB19029	
    
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")                                   
    
    Call InitVariables														
    Call SetDefaultVal	
    Call InitComboBox
    Call InitSpreadSheet()		
    
    Call SetToolbar("110000000000111")										'��: ��ư ���� ����    	
End Sub

'========================================================================================================= 
Sub InitComboBox()
    Call SetCombo(frm1.cboJobType, "SD", "����������(SD)")
    Call SetCombo(frm1.cboJobType, "MM", "���Կ�����(MM)")
    Call SetCombo(frm1.cboJobType, "PMS", "����������(PMS)")
End Sub

Sub InitSpreadPosVariables()
	'add tab1 header datatable column
	C_inv_type			= 1
	C_inv_no				= 2
	C_dt_inv_no			= 3
	C_process_date		= 4
	C_success_flag		= 5
	C_re_flag			= 6
	C_error_desc		= 7
	C_change_reason	= 8
	C_change_remark	= 9
	C_change_remark2	= 10
	C_change_remark3	= 11
	C_insert_user_id	= 12
	C_user_name			= 13
End Sub


'========================================================================================================= 
Sub InitVariables()
    lgIntFlgMode = parent.OPMD_CMODE				'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False							'Indicates that no value changed
End Sub

'========================================================================================================= 
Sub SetDefaultVal()
	'������ ���ڴ� ������ ���ڸ� ��ȸ�Ѵ�.
    Dim EndDate
	EndDate = UniConvDateAToB(iDBSYSDate, parent.gServerDateFormat, parent.gDateFormat)

	'������ ���ڴ� ���� ~ ���� �̴�.
	frm1.txtIssuedFromDt.text  = EndDate
	frm1.txtIssuedToDt.text    = EndDate
End Sub

'========================================================================================
Sub InitSpreadSheet()

	Call initSpreadPosVariables()

	With frm1.vspdData	
		.MaxCols = C_user_name + 1								'��: �ִ� Columns�� �׻� 1�� ������Ŵ 
		.Col = .MaxCols												'��: ����� �� Hidden Column
		.ColHidden = True
		.MaxRows = 0
		ggoSpread.Source = frm1.vspdData
		.ReDraw = False
		ggoSpread.Spreadinit "V20090707",, parent.gAllowDragDropSpread
		.ReDraw = False

		Call GetSpreadColumnPos("A")

		' uniGrid1 setting
		ggoSpread.SSSetEdit  	C_inv_type,			"������", 			10, ,,18
		ggoSpread.SSSetEdit  	C_inv_no,			"���ݰ�꼭��ȣ", 	15, ,,18
		ggoSpread.SSSetEdit		C_dt_inv_no,		"���۰�����ȣ",		30, 2,,50
		ggoSpread.SSSetEdit  	C_process_date,   "�۾���",				20, ,,35
		ggoSpread.SSSetEdit  	C_success_flag,  	"��������",			10, ,,10
		ggoSpread.SSSetEdit  	C_re_flag,			"����࿩��",		10, ,,10
		ggoSpread.SSSetEdit  	C_error_desc, 		"��������", 			40, ,,300
		ggoSpread.SSSetEdit  	C_change_reason,	"��������",			20, ,,150
		ggoSpread.SSSetEdit  	C_change_remark,	"���1",				15, ,,15
		ggoSpread.SSSetEdit  	C_change_remark2,	"���2",				15, ,,15
		ggoSpread.SSSetEdit  	C_change_remark3,	"���3",				15, ,,15
		ggoSpread.SSSetEdit  	C_insert_user_id,	"�۾���",   			15, ,,18
		ggoSpread.SSSetEdit  	C_user_name,		"�۾��ڸ�",   		15, ,,18

		.ReDraw = True
	End With

	Call SetSpreadLock()
End Sub

'========================================================================================
Sub SetSpreadLock()
	With frm1
		ggoSpread.Source = .vspdData
		.vspdData.ReDraw = False

		ggoSpread.SpreadLockWithOddEvenRowColor()

		frm1.vspddata.col = C_inv_type
		frm1.vspddata.row = 0
		frm1.vspddata.ColHeadersShow = True

		.vspdData.ReDraw = True
	End With
End Sub


'========================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
	Dim iCurColumnPos

	Select Case UCase(pvSpdNo)
		Case "A"
			ggoSpread.Source = frm1.vspdData
			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			
			C_inv_type			= iCurColumnPos(1)
			C_inv_no				= iCurColumnPos(2)
			C_dt_inv_no			= iCurColumnPos(3)
			C_process_date		= iCurColumnPos(4)
			C_success_flag  	= iCurColumnPos(5)
			C_re_flag			= iCurColumnPos(6)
			C_error_desc 		= iCurColumnPos(7)
			C_change_reason 	= iCurColumnPos(8)
			C_change_remark 	= iCurColumnPos(9)
			C_change_remark2 	= iCurColumnPos(10)
			C_change_remark3 	= iCurColumnPos(11)
			C_insert_user_id	= iCurColumnPos(12)
			C_user_name			= iCurColumnPos(13)
	End Select    
End Sub

'========================================================================================================= 
Sub txtIssuedFromDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13 Then 
		frm1.txtIssuedToDt.focus
		Call FncQuery
	End If
End Sub

'========================================================================================================= 
Sub txtIssuedToDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13 Then 
		frm1.txtIssuedFromDt.focus
		Call FncQuery
	End If
End Sub


'#########################################################################################################
'												4. Common Function�� 
'=========================================================================================================
Function FncQuery() 

    Dim IntRetCD 

    FncQuery = False																		'��: Processing is NG

    Err.Clear																				'��: Protect system from crashing

    '-----------------------
    'Check previous data area
    '-----------------------
    With frm1
	    ggoSpread.Source = .vspdData
	    If ggoSpread.SSCheckChange = True Then
			IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "X", "X")					'����Ÿ�� ����Ǿ����ϴ�. ��ȸ�Ͻðڽ��ϱ�?
	    	If IntRetCD = vbNo Then
		      	Exit Function
	    	End If
	    End If

		'-----------------------
	    'Check condition area
	    '-----------------------
		If Not chkFieldByCell(.txtIssuedFromDt, "A", "1") Then Exit Function
		If Not chkFieldByCell(.txtIssuedToDt, "A", "1") Then Exit Function

	   If CompareDateByFormat( .txtIssuedFromDt.text, _
										.txtIssuedToDt.text, _
										.txtIssuedFromDt.Alt, _
										.txtIssuedToDt.Alt, _
										"970025", _
										.txtIssuedFromDt.UserDefinedFormat, _
										parent.gComDateType, _
										True) = False Then		
			Exit Function
		End If

		'-----------------------
		'Erase contents area
		'-----------------------
		'	    Call ggoOper.ClearField(Document, "2")												'��: Clear Contents  Field
		ggoSpread.Source = frm1.vspdData
		ggoSpread.ClearSpreadData

		Call InitVariables 																	'��: Initializes local global variables

		FncQuery = True	
	End With
	
	Call DBquery()
End Function

'========================================================================================
Function FncCancel() 
	If frm1.vspdData.MaxRows < 1 Then Exit Function
	
    ggoSpread.Source = frm1.vspdData	
    ggoSpread.EditUndo																	'��: Protect system from crashing    
End Function

'=======================================================================================================
Function FncPrint()
    Call parent.FncPrint()																'��: Protect system from crashing
End Function

'=======================================================================================================
Function FncExcel() 
    Call parent.FncExport(parent.C_MULTI)												'��: ȭ�� ���� 
End Function


'=======================================================================================================
Function FncFind() 
    Call parent.FncFind(parent.C_MULTI, False)											'��:ȭ�� ����, Tab ���� 
End Function

'========================================================================================
Sub FncSplitColumn()
    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   
End Sub

'=======================================================================================================
Function FncExit()
	Dim IntRetCD
	
	FncExit = False
	
	ggoSpread.Source = frm1.vspdData	
	If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO, "X", "X")								'����Ÿ�� ����Ǿ����ϴ�. ���� �Ͻðڽ��ϱ�?
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If

	FncExit = True
End Function

'========================================================================================
Sub PopSaveSpreadColumnInf()
	ggoSpread.Source = gActiveSpdSheet
	Call ggoSpread.SaveSpreadColumnInf()
End Sub

'========================================================================================
Sub PopRestoreSpreadColumnInf()
	ggoSpread.Source = gActiveSpdSheet
	Call ggoSpread.RestoreSpreadInf()
	Call InitSpreadSheet()      
	Call ggoSpread.ReOrderingSpreadData()
End Sub

'========================================================================================
Function DbQuery() 
	Dim strVal
	Dim txtJobType
	Dim txtFromDate
	Dim txtToDate
	Dim txtERPUser
	Dim txtINVNo
	
	DbQuery = False
	
	With frm1
		If .cboJobType.value = "" Then
			txtJobType = "%"
		Else
			txtJobType = .cboJobType.value
		End If

		If .txtIssuedFromDt.text = "" Then
			txtFromDate = "1900-01-01"
		Else
			txtFromDate = .txtIssuedFromDt.text
		End If
		
		If .txtIssuedToDt.text = "" Then
			txtToDate = "9999-12-31"
		Else
			txtToDate = .txtIssuedToDt.text
		End If

		If .txtuserId.value = "" Then
			txtERPUser = "%"
            .txtuserNm.value = ""
		Else
			txtERPUser = .txtuserId.value
		End If

		If .txtBillNo.value = "" Then
			txtINVNo = "%"
		Else
			txtINVNo = .txtBillNo.value
		End If

		strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001 & _
		                      "&txtJobType=" & Trim(txtJobType) & _
		                      "&txtFromDate=" & Trim(txtFromDate) & _
		                      "&txtToDate=" & Trim(txtToDate) & _
		                      "&txtERPUser=" & Trim(txtERPUser) & _
		                      "&txtINVNo=" & Trim(txtINVNo)
	End With
	
	Call LayerShowHide(1)
	Call RunMyBizASP(MyBizASP, strVal)																'��: �����Ͻ� ASP �� ���� 
	
	DbQuery = True
End Function

'========================================================================================
Function DbQueryOk()																		'��: ��ȸ ������ ������� 
	'-----------------------
	'Reset variables area
	'-----------------------
	lgIntFlgMode = parent.OPMD_UMODE																'��: Indicates that current mode is Update mode

	With frm1
		Call LayerShowHide(0)
	End With
End Function

Function Open_User1()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim OriginCol, TempCd
	Dim IntRetCD

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "usr_id, usr_nm"				<%' �˾� ��Ī %>
	arrParam(1) = "z_usr_mast_rec"				<%' TABLE ��Ī %>
	arrParam(2) = frm1.txtuserId.value			<%' Code Condition%>
	arrParam(4) = ""							<%' Name Cindition%>
	arrParam(5) = "�����"						<%' �����ʵ��� �� ��Ī %>
		
    arrField(0) = "usr_id"						<%' Field��(0)%>
    arrField(1) = "usr_nm"						<%' Field��(1)%>
    
    arrHeader(0) = "�����"						<%' Header��(0)%>
    arrHeader(1) = "����ڸ�"					<%' Header��(1)%>
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", _
												Array(arrParam, arrField, arrHeader), _
												"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	 
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		frm1.txtuserId.value = arrRet(0)
		frm1.txtuserNm.value = arrRet(1)
	End If	
	
End Function

'=======================================================================================================
'   Event Name : txtYr1_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtIssuedFromDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtIssuedFromDt.Action = 7
    End If
End Sub

'=======================================================================================================
'   Event Name : txtYr1_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtIssuedToDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtIssuedToDt.Action = 7
    End If
End Sub

Sub vspdData_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    
    Call GetSpreadColumnPos("A")
End Sub


 Sub vspdData_Click(ByVal Col , ByVal Row )

	Call SetPopupMenuItemInf("0000111111")
	gMouseClickStatus = "SPC"
	Set gActiveSpdSheet = frm1.vspdData

	If frm1.vspdData.MaxRows = 0 Then                                                    'If there is no data.
		Exit Sub
	End If

	If Row <= 0 Then
		ggoSpread.Source = frm1.vspdData
		If lgSortKey = 1 Then
			ggoSpread.SSSort Col			'Sort In Assending
			lgSortKey = 2
		Else
			ggoSpread.SSSort Col, lgSortKey	'Sort In Desending
			lgSortKey = 1
		End If
		Exit Sub
	End If
End Sub

Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If

End Sub   
</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->	
</HEAD>
<% '#########################################################################################################
'       					6. Tag�� 
'######################################################################################################### %>
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
							<TD CLASS="CLSSTABP">
								<TABLE ID="MyTab1" CELLSPACING=0 CELLPADDING=0>
									<TR>
										<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
										<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>LOG��ȸ</font></td>
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
							<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
						</TR>
						<TR>
							<TD HEIGHT=20 WIDTH=100%>
								<FIELDSET CLASS="CLSFLD">
									<TABLE <%=LR_SPACE_TYPE_40%>>
										<TR>
 											<TD CLASS="TD5" NOWRAP>������</TD>
											<TD CLASS="TD6" NOWRAP><SELECT NAME="cboJobType" ALT="������" STYLE="Width: 120px;" tag="11"><OPTION VALUE=""></OPTION></SELECT></TD>
											<TD CLASS="TD5"NOWRAP>�۾���</TD>
											<TD CLASS="TD6"NOWRAP>
												<SCRIPT LANGUAGE=JavaScript>ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 name=txtIssuedFromDt CLASS=FPDTYYYYMMDD title=FPDATETIME tag="11X1" ALT="��������"></OBJECT>');</SCRIPT> ~
 												<SCRIPT LANGUAGE=JavaScript>ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime2 name=txtIssuedToDt CLASS=FPDTYYYYMMDD title=FPDATETIME tag="11X1" ALT="��������"></OBJECT>');</SCRIPT>
 											</TD>
										</TR>
										<TR>
											<TD CLASS="TD5">ERP�����</TD>
											<TD CLASS="TD6">
												<INPUT TYPE=TEXT NAME="txtuserId" SIZE=10  MAXLENGTH=13 tag="11XXXU" ALT="ERP���"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnUser" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Open_User1()">
												<INPUT TYPE=TEXT NAME="txtuserNm" tag="14X">
											</TD>
 											<TD CLASS="TD5" NOWRAP>���ݰ�꼭��ȣ</TD>
											<TD CLASS="TD6" NOWRAP>
												<INPUT CLASS="clstxt" TYPE=TEXT NAME="txtBillNo" SIZE=30 MAXLENGTH=40 STYLE="TEXT-ALIGN: Left" tag="11XXXU" ALT="���ݰ�꼭��ȣ">
											</TD>
										</TR>
									</TABLE>
								</FIELDSET>
							</TD>
						</TR>
						<TR>
							<TD <%=HEIGHT_TYPE_03%> WIDTH="100%"></TD>
						</TR>
						<TR>
							<TD WIDTH=100% HEIGHT=* valign=top>
								<TABLE <%=LR_SPACE_TYPE_20%>>
									<TR HEIGHT="60%">
										<TD  WIDTH="100%" colspan=4><SCRIPT LANGUAGE=JavaScript>ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=vspdData><PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"></OBJECT>');</SCRIPT></TD>
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
				<TD WIDTH="100%" HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH="100%" HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME></TD>
			</TR>
		</TABLE>
	</FORM>
	<DIV ID="MousePT" NAME="MousePT">
		<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=280 height=41 src="../../inc/cursor.htm"></iframe>
	</DIV>
	<FORM NAME=EBAction TARGET="MyBizASP"   METHOD="POST">
		<INPUT TYPE="HIDDEN" NAME="uname"       TABINDEX="-1">
		<INPUT TYPE="HIDDEN" NAME="dbname"      TABINDEX="-1">
		<INPUT TYPE="HIDDEN" NAME="filename"    TABINDEX="-1">
		<INPUT TYPE="HIDDEN" NAME="condvar"     TABINDEX="-1">
		<INPUT TYPE="HIDDEN" NAME="date">	
	</Form>
</BODY>
</HTML>
