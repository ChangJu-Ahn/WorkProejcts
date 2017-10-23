<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Quality Management
'*  2. Function Name        : 
'*  3. Program ID           : Q1441MA1
'*  4. Program Name         : T������ �ý��� ��� �׷���
'*  5. Program Desc         : Quality Configuration
'*  6. Component List       : 
'*  7. Modified date(First) : 2002/05/14
'*  8. Modified date(Last)  : 2003/05/15
'*  9. Modifier (First)     : Koh Jae Woo
'* 10. Modifier (Last)      : Park Hyun Soo
'* 11. Comment
'* 12. Common Coding Guide  : this mark(��) means that "Do not change" 
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'**********************************************************************************************
-->
<HTML>
<HEAD>
<TITLE>T������ �ý��� ��� �׷���</TITLE>

<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/Cookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "JavaScript" SRC = "../../inc/TabScript.js"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript">

Option Explicit

Const BIZ_PGM_QRY_ID = "q1441mb1.asp"							'��: Query �����Ͻ� ���� ASP�� 
Const PGM_JUMP_ID1 = "q1411ma1"
'/* Issue: �˻��� �������� Return - START */
Const PGM_JUMP_ID2_1 = "q1413ma1.asp"
Const PGM_JUMP_ID2_2 = "q1413ma2.asp"
Const PGM_JUMP_ID2_3 = "q1413ma3.asp"
Const PGM_JUMP_ID2_4 = "q1413ma4.asp"

Dim lgReturnPage
'/* Issue: �˻��� �������� Return - END */

Dim IsOpenPop        

<!-- #Include file="../../inc/lgvariables.inc" -->

'==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : ���� �ʱ�ȭ(Global ����, �ʱ�ȭ�� �ʿ��� ���� �Ǵ� Flag���� Setting�Ѵ�.)
'========================================================================================================= 
Sub InitVariables()

    lgIntFlgMode = Parent.OPMD_CMODE                   	'Indicates that current mode is Create mode
    lgIntGrpCount = 0        	              	'initializes Group View Size
    lgStrPrevKey = ""                           		'initializes Previous Key
    lgLngCurRows = 0                         		'initializes Deleted Rows Count
End Sub

'==========================================  2.2.1 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : ȭ�� �ʱ�ȭ(���� Field�� �� �� ȭ���� �� �� Default���� ������� �ϴ� Field�� Setting)
'========================================================================================================= 
Sub SetDefaultVal()
	frm1.txtLotsize.Text = "<%=Request("txtLotSize")%>"
	frm1.txtSamplesize.Text = "<%=Request("txtSamplesize")%>"	
	frm1.txtAcceptCount.Text = "<%=Request("txtAcceptSize")%>"	
	frm1.txtProcessDefectRatio.Text = "<%=Request("txtDefectRate")%>"	
	frm1.txtReplaceMode.value = "<%=Request("txtReplaceMode")%>"
	
	'/* Issue: �˻��� �������� Return - START */
	lgReturnPage = "<%=Request("txtPageCode")%>"
	'/* Issue: �˻��� �������� Return - END */
End Sub

'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== 
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "*", "NOCOOKIE","MA") %>
End Sub

'=============================================  2.3.3()  ======================================
'=	Event Name : ReturnClick
'=	Event Desc :
'========================================================================================================
Function ReturnClick()	
	PgmJump(PGM_JUMP_ID1)
End Function

'/* Issue: �˻��� �������� Return - START */
'=============================================  2.3.4()  ======================================
'=	Event Name : Return2Click
'=	Event Desc :
'========================================================================================================
Function Return2Click()
	Select Case lgReturnPage
		Case "OA"		'OC
			Location.href = PGM_JUMP_ID2_1
		Case "SA"		'Screen
			Location.href = PGM_JUMP_ID2_2
		Case "AA"		'Adjust
			Location.href = PGM_JUMP_ID2_3
		Case "OA2"		'OC 2ȸ 
			Location.href = PGM_JUMP_ID2_4
	End Select 
End Function
'/* Issue: �˻��� �������� Return - END */

'==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(���� Include ���Ͽ� ����)�� �����ʱ�ȭ �� ȭ���ʱ�ȭ�� �ϱ� ���� �Լ��� Call�ϴ� �κ� 
'========================================================================================================= 
Sub Form_Load()
	Call LoadInfTB19029                                                     	'��: Load table , B_numeric_format
	Call AppendNumberPlace("6", "3", "2")
	Call ggoOper.LockField(Document, "N")                                   	'��: Lock  Suitable  Field
	Call ggoOper.FormatField(Document, "1", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec)
	Call ggoOper.FormatField(Document, "2", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec)
		
	Call InitVariables																'��: Initializes local global variables          
	Call SetDefaultVal
	Call SetToolbar("10000000000111")
	Call fncQuery					'�������� ���� 
	
End Sub

'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )
    
End Sub

'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================
Function FncQuery() 

	Dim IntRetCD 
	Dim Replace
	
	FncQuery = False                                                        '��: Processing is NG
	
	Err.Clear     	                                                     '��: Protect system from crashing

    '-----------------------
	'Erase contents area
	'----------------------- 

	Call InitVariables						'��: Initializes local global variables
	'-----------------------
	'Check condition area
	'----------------------- 
	'��: This function check indispensable field
	With frm1
		.ChartFX1.ToolBar = 0					'���� �����ϱ� 
		.ChartFX1.CloseData 1 Or &H800				'��Ʈ FX���� ������ ä�� �ʱ�ȭ 
		.ChartFX2.ToolBar = 0					'���� �����ϱ� 
		.ChartFX2.CloseData 1 Or &H800				'��Ʈ FX���� ������ ä�� �ʱ�ȭ 
		.ChartFX3.ToolBar = 0					'���� �����ϱ� 
		.ChartFX3.CloseData 1 Or &H800				'��Ʈ FX���� ������ ä�� �ʱ�ȭ 
		.ChartFX4.ToolBar = 0					'���� �����ϱ� 
		.ChartFX4.CloseData 1 Or &H800				'��Ʈ FX���� ������ ä�� �ʱ�ȭ 
	End With
	
	'-----------------------
	'Query function call area
	'----------------------- 
	If frm1.txtReplaceMode.value <> "" Then
		Replace = frm1.txtReplaceMode.value
	ElseIf frm1.txtReplaceMode.value = "" Then
		Replace = 0
	End If
	
	frm1.txtReplaceMode.value= Replace
	
	Call DbQuery									'��: Query db data
	
	FncQuery = True									'��: Processing is OK
End Function

'========================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'========================================================================================
Function FncNew() 

	Dim IntRetCD 
	
	FncNew = False                                                          					'��: Processing is NG
	
	'-----------------------
	'Erase condition area
	'Erase contents area
	'-----------------------
	
	Call ggoOper.ClearField(Document, "1")                    			'��: Clear Condition Field
	Call ggoOper.LockField(Document, "N")                                       		'��: Lock  Suitable  Field
	Call SetDefaultVal

	With frm1
		.ChartFX1.ToolBar = 0					'���� �����ϱ� 
		.ChartFX1.CloseData 1 Or &H800				'��Ʈ FX���� ������ ä�� �ʱ�ȭ 
		.ChartFX2.ToolBar = 0					'���� �����ϱ� 
		.ChartFX2.CloseData 1 Or &H800				'��Ʈ FX���� ������ ä�� �ʱ�ȭ 
		.ChartFX3.ToolBar = 0					'���� �����ϱ� 
		.ChartFX3.CloseData 1 Or &H800				'��Ʈ FX���� ������ ä�� �ʱ�ȭ 
		.ChartFX4.ToolBar = 0					'���� �����ϱ� 
		.ChartFX4.CloseData 1 Or &H800				'��Ʈ FX���� ������ ä�� �ʱ�ȭ 
	End With
	
	FncNew = True
End Function

'========================================================================================
' Function Name : FncDelete
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncDelete() 
	FncDelete = False
End Function

'========================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncSave() 
	FncSave = False
End Function

'========================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================
Function FncCopy() 
	FncCopy = False
End Function

'========================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================
Function FncCancel() 
    FncCancel = False
End Function

'========================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================
Function FncInsertRow() 
	FncInsertRow = False
End Function

'========================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================
Function FncDeleteRow()
	FncDeleteRow = False
End Function

'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================
Function FncPrint() 
	FncPrint = False
    Call Parent.FncPrint()
    FncPrint = True
End Function

'========================================================================================
' Function Name : FncPrev
' Function Desc : This function is related to Previous Button
'========================================================================================
Function FncPrev() 
    FncPrev = False
End Function

'========================================================================================
' Function Name : FncNext
' Function Desc : This function is related to Next Button
'========================================================================================
Function FncNext() 
    FncNext = False
End Function

'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================
Function FncPrint() 
    FncPrint = False
End Function

'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================
Function FncExcel() 
	FncExcel = False
	Call parent.FncExport(Parent.C_SINGLE)					'��: ȭ�� ���� 
	FncExcel = True
End Function

'========================================================================================
' Function Name : FncExit
' Function Desc : This function is related to Exit 
'========================================================================================
Function FncExit()
	FncExit = True
End Function

'*******************************  5.2 Fnc�Լ����� ȣ��Ǵ� ���� Function  *******************************
'	���� : 
'********************************************************************************************************* 

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbQuery() 
	
	Dim strVal
	
	Call LayerShowHide(1)

	Err.Clear              

	DbQuery = False                                                        					'��: Processing is NG
	
	strVal = BIZ_PGM_QRY_ID & "?txtLotSize=" & frm1.txtLotSize.Text 			'��: ATI��꿡 ���� ��Ʈũ�� 
	strVal = strVal & "&txtSamplesize=" & frm1.txtSamplesize.Text				'��: ATI��꿡 ���� ����ũ��.
	strVal = strVal & "&txtProcessDefectRatio=" & frm1.txtProcessDefectRatio.Text		'��: ATI��꿡 ���� �ҷ��� 
	strVal = strVal & "&txtAcceptCount=" & frm1.txtAcceptCount.Text			'��: ATI��꿡 ���� �պ��������� 
	strVal = strVal & "&txtReplaceMode=" & Trim(frm1.txtReplaceMode.value)				'��: �ҷ�ǰ ��ü���� 
	
	Call RunMyBizASP(MyBizASP, strVal)							'��: �����Ͻ� ASP �� ���� 

	DbQuery = True                                                          					'��: Processing is NG

End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================
Function DbQueryOk()														'��: ��ȸ ������ ������� 
	
End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>
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
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>�˻�Ư�� �׷���</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    	</TR>
						</TABLE>
					</TD>
					<TD WIDTH=* align=right>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=* >
		<TD WIDTH=100% CLASS="Tab11">
			<TABLE WIDTH=100% HEIGHT=100% <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD VALIGN="top" HEIGHT=30% WIDTH="12%">
					<FIELDSET>
						<TABLE WIDTH="100%" CELLSPACING=0 CELLPADDING=0>		
							<TR>
								<TD CLASS="TD5" NOWRAP HEIGHT=5></TD>
								<TD CLASS="TD6" NOWRAP HEIGHT=5></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>�˻�Ư�����</TD>
								<TD CLASS="TD6" NOWRAP></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP HEIGHT=2></TD>
								<TD CLASS="TD6" NOWRAP HEIGHT=2></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>��Ʈũ��</TD>
								<TD CLASS="TD6" NOWRAP>
									<script language =javascript src='./js/q1441ma1_txtLotSize_txtLotSize.js'></script>
								</TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP HEIGHT=2></TD>
								<TD CLASS="TD6" NOWRAP HEIGHT=2></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>����ũ��</TD>
								<TD CLASS="TD6" NOWRAP>
									<script language =javascript src='./js/q1441ma1_txtSamplesize_txtSamplesize.js'></script>
								</TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP HEIGHT=2></TD>
								<TD CLASS="TD6" NOWRAP HEIGHT=2></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>�հ���������</TD>
								<TD CLASS="TD6" NOWRAP>
									<script language =javascript src='./js/q1441ma1_txtAcceptCount_txtAcceptCount.js'></script>
								</TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP HEIGHT=1></TD>
								<TD CLASS="TD6" NOWRAP HEIGHT=1></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>�����ҷ���</TD>
								<TD CLASS="TD6" NOWRAP>
									<script language =javascript src='./js/q1441ma1_txtProcessDefectRatio_txtProcessDefectRatio.js'></script>&nbsp;%&nbsp;
								</TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP HEIGHT=5></TD>
								<TD CLASS="TD6" NOWRAP HEIGHT=5></TD>
							</TR>
						</TABLE>
						</FIELDSET>
					</TD>
					<!--
					<TD HEIGHT=30% WIDTH=12%>
					<FIELDSET>
						<TABLE WIDTH="100%" CELLSPACING=0 CELLPADDING=0>		
							<TR>
								<TD CLASS="TD5" NOWRAP>2ȸ</TD>
								<TD CLASS="TD6" NOWRAP></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>1�� ����ũ��</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtASNSamplesize1" SIZE=20 MAXLENGTH=15 tag="14" STYLE="Text-Align: Right"></TD>
							</TR>
							<TR>
								<TD HEIGHT=8% WIDTH=100%></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>2�� ����ũ��</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtASNSamplesize2" SIZE=20 MAXLENGTH=15 tag="14" STYLE="Text-Align: Right"></TD>
							</TR>
							<TR>
								<TD HEIGHT=8% WIDTH=100%></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>1�� ��������</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtASNAcceptanceCnt1" SIZE=20 MAXLENGTH=15 tag="14" STYLE="Text-Align: Right"></TD>
							</TR>
							<TR>
								<TD HEIGHT=8% WIDTH=100%></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>2�� ��������</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtASNAcceptanceCnt2" SIZE=20 MAXLENGTH=15 tag="14" STYLE="Text-Align: Right"></TD>
							</TR>
							<TR>
								<TD HEIGHT=200 WIDTH=100%></TD>
							</TR>
						</TABLE>
						</FIELDSET>
					</TD>
					-->	
					<TD HEIGHT=* WIDTH="8%">
					</TD>
					<TD WIDTH=100%>
						<TABLE WIDTH=100% HEIGHT=100% CELLSPACING=0 CELLPADDING=0>
							<TR>
								<TD HEIGHT=50% WIDTH=50%>
									<script language =javascript src='./js/q1441ma1_ChartFX2_N655223810.js'></script>
								</TD>
								<TD HEIGHT=50% WIDTH=50%>
									<script language =javascript src='./js/q1441ma1_ChartFX1_N715183144.js'></script>
								</TD>
							</TR>
							<TR>
								<TD HEIGHT=50% WIDTH=50%>
									<script language =javascript src='./js/q1441ma1_ChartFX4_N922628140.js'></script>
								</TD>
								<TD HEIGHT=50% WIDTH=50%>
									<script language =javascript src='./js/q1441ma1_ChartFX3_N657074159.js'></script>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_01%>> </TD>
	</TR>
	<TR HEIGHT="20">
		<TD WIDTH="100%">
    		<TABLE WIDTH="100%" <%=LR_SPACE_TYPE_30%>>
	   			<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<!--'/* Issue: �˻��� �������� Return - START */-->
        			<TD WIDTH=* ALIGN=RIGHT><A href="vbscript:ReturnClick">������ �ý��� ����</A>&nbsp;|&nbsp;<A href="vbscript:Return2Click">�˻��� ����</A></TD>
        			<!--'/* Issue: �˻��� �������� Return - START */-->
        			<TD WIDTH=10>&nbsp;</TD>
       			</TR>
      		</TABLE>
    	</TD>
    </TR>
	<TR>	
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP"  tabindex=-1 WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" tabindex=-1>
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24" tabindex=-1>
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24" tabindex=-1>
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" tabindex=-1>
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24" tabindex=-1>
<INPUT TYPE=HIDDEN NAME="txtCpFlag" tag="24" tabindex=-1>
<INPUT TYPE=HIDDEN NAME="txtReplaceMode" tag="24" tabindex=-1>
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

