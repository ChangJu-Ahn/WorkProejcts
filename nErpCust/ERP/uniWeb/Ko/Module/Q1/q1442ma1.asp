<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Quality Management
'*  2. Function Name        : 
'*  3. Program ID           : Q1442MA1
'*  4. Program Name         : T전문가 시스템 결과 그래프
'*  5. Program Desc         : Quality Configuration
'*  6. Component List       : 
'*  7. Modified date(First) : 2002/05/14
'*  8. Modified date(Last)  : 2003/05/15
'*  9. Modifier (First)     : Koh Jae Woo
'* 10. Modifier (Last)      : Park Hyun Soo
'* 11. Comment
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change" 
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************
-->
<HTML>
<HEAD>
<TITLE>T전문가 시스템 결과 그래프</TITLE>

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

Const BIZ_PGM_QRY_ID = "q1442mb1.asp"							'☆: Query 비지니스 로직 ASP명 
Const PGM_JUMP_ID1 = "q1411ma1"
'/* Issue: 검사방식 적용으로 Return - START */
Const PGM_JUMP_ID2_1 = "q1413ma5.asp"
Const PGM_JUMP_ID2_2 = "q1413ma6.asp"

Dim lgReturnPage
'/* Issue: 검사방식 적용으로 Return - END */

<!-- #Include file="../../inc/lgvariables.inc" -->

Dim IsOpenPop        

'==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'========================================================================================================= 
Sub InitVariables()

    lgIntFlgMode = Parent.OPMD_CMODE                   	'Indicates that current mode is Create mode
    lgIntGrpCount = 0        	              	'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
    
    lgStrPrevKey = ""                           		'initializes Previous Key
    lgLngCurRows = 0                         		'initializes Deleted Rows Count
End Sub

'==========================================  2.2.1 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
'========================================================================================================= 
Sub SetDefaultVal()
	With frm1
		.txtUpperAcceptCount.Text = "<%= Request("txtUpperBound")%>"
		.txtLowerAcceptCount.Text = "<%= Request("txtLowerBound")%>"
		.txtSD.Text = "<%= Request("txtSD")%>"
		.txtSampleSize.Text = "<%= Request("txtSampleSize")%>"
		.txtInsCri.value = "<%= Request("txtInsCri")%>"	
	End With
	'/* Issue: 검사방식 적용으로 Return - START */
	lgReturnPage = "<%=Request("txtPageCode")%>"
	'/* Issue: 검사방식 적용으로 Return - END */
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

'/* Issue: 검사방식 적용으로 Return - START */
'=============================================  2.3.4()  ======================================
'=	Event Name : Return2Click
'=	Event Desc :
'========================================================================================================
Function Return2Click()
	Select Case lgReturnPage
		Case "OV"		'OC
			Location.href = PGM_JUMP_ID2_1
		Case "AV"		'Adjust
			Location.href = PGM_JUMP_ID2_2
	End Select 

End Function
'/* Issue: 검사방식 적용으로 Return - END */

'==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분 
'========================================================================================================= 
Sub Form_Load()
	Call LoadInfTB19029                                                     	'⊙: Load table , B_numeric_format
	Call AppendNumberPlace("6", "3", "2")
	Call AppendNumberPlace("7", "11", "4")
	Call ggoOper.LockField(Document, "N")                                   	'⊙: Lock  Suitable  Field
	Call ggoOper.FormatField(Document, "1", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec)
    Call ggoOper.FormatField(Document, "2", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec)
	
	Call InitVariables																'⊙: Initializes local global variables          
    Call SetDefaultVal
    Call SetToolbar("10000000000111")
    Call fncQuery					'업무로직 시작 
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
	
	FncQuery = False                                                        '⊙: Processing is NG
	
	Err.Clear     	                                                     '☜: Protect system from crashing

	'-----------------------
	'Erase contents area
	'----------------------- 
	Call InitVariables						'⊙: Initializes local global variables
	
	'⊙: This function check indispensable field
	
	frm1.ChartFX1.ToolBar = 0					'툴바 제거하기 
	frm1.ChartFX1.CloseData 1 Or &H800				'차트 FX와의 데이터 채널 초기화 
	
	  '-----------------------
	'Query function call area
	'----------------------- 
	
	If ReadCookie("txtInsReplace") <> "" Then
		Replace = ReadCookie("txtInsReplace")
	End If
		
	Call DbQuery									'☜: Query db data

	FncQuery = True									'⊙: Processing is OK
        	
End Function

'========================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'========================================================================================
Function FncNew() 

	Dim IntRetCD 
	
	FncNew = False                                                          					'⊙: Processing is NG
	
	'-----------------------
	'Erase condition area
	'Erase contents area
	'-----------------------
	
	Call ggoOper.ClearField(Document, "1")                    			'⊙: Clear Condition Field
	Call ggoOper.LockField(Document, "N")                                       		'⊙: Lock  Suitable  Field
	Call SetDefaultVal
	
	frm1.ChartFX1.ToolBar = 0					'툴바 제거하기 
	frm1.ChartFX1.CloseData 1 Or &H800				'차트 FX와의 데이터 채널 초기화 
	
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
	Call parent.FncExport(Parent.C_SINGLE)					'☜: 화면 유형 
	FncExcel = True
End Function

'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================
Function FncExit()
	FncExit = True
End Function

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbQuery() 
	
	Dim strVal
    
    	Err.Clear              

	DbQuery = False                                                        					'⊙: Processing is NG
	
	strVal = BIZ_PGM_QRY_ID & "?txtUpperAcceptCount=" & frm1.txtUpperAcceptCount.Text	'☜: 상한합격판정치 
	strVal = strVal & "&txtLowerAcceptCount=" & frm1.txtLowerAcceptCount.Text		'☜: 하한합격판정치 
	strVal = strVal & "&txtSD=" & frm1.txtSD.Text					'☜: 표준편차 
	strVal = strVal & "&txtSampleSize=" & frm1.txtSampleSize.Text			'☜: 샘플크기.
		
	Call RunMyBizASP(MyBizASP, strVal)							'☜: 비지니스 ASP 를 가동 

	DbQuery = True                                                          					'⊙: Processing is NG

End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================
Function DbQueryOk()														'☆: 조회 성공후 실행로직 
	
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
			<TABLE <%=LR_SPACE_TYPE_10%> BORDER=0>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>검사특성 그래프</font></td>
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
									<TD CLASS="TD5" NOWRAP>검사특성모수</TD>
									<TD CLASS="TD6" NOWRAP></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP HEIGHT=2></TD>
									<TD CLASS="TD6" NOWRAP HEIGHT=2></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>상한합격판정치</TD>
									<TD CLASS="TD6" NOWRAP>
										<!-- /* 8월 정기패치: 수량 포맷 관련 Tag 수정 */ -->
										<script language =javascript src='./js/q1442ma1_txtUpperAcceptCount_txtUpperAcceptCount.js'></script>
									</TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP HEIGHT=2></TD>
									<TD CLASS="TD6" NOWRAP HEIGHT=2></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWPAP>하한합격판정치</TD>
									<TD CLASS="TD6" NOWRAP>
										<!-- /* 8월 정기패치: 수량 포맷 관련 Tag 수정 */ -->
										<script language =javascript src='./js/q1442ma1_txtLowerAcceptCount_txtLowerAcceptCount.js'></script>
									</TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP HEIGHT=2></TD>
									<TD CLASS="TD6" NOWRAP HEIGHT=2></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>표준편차</TD>
									<TD CLASS="TD6" NOWRAP>
										<!-- /* 8월 정기패치: 수량 포맷 관련 Tag 수정 */ -->
										<script language =javascript src='./js/q1442ma1_txtSD_txtSD.js'></script>
									</TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP HEIGHT=3></TD>
									<TD CLASS="TD6" NOWRAP HEIGHT=3></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>샘플크기</TD>
									<TD CLASS="TD6" NOWRAP>
										<script language =javascript src='./js/q1442ma1_txtSampleSize_txtSampleSize.js'></script>
									</TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP HEIGHT=5></TD>
									<TD CLASS="TD6" NOWRAP HEIGHT=5></TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
					<TD HEIGHT=* WIDTH="8%">
					</TD>
					<TD WIDTH=100%>
						<TABLE WIDTH=100% HEIGHT=100% CELLSPACING=0 CELLPADDING=0>
							<TR>
								<TD HEIGHT=100% WIDTH=100%>
									<script language =javascript src='./js/q1442ma1_ChartFX1_N327143085.js'></script>
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
					<!--'/* Issue: 검사방식 적용으로 Return - START */-->
        			<TD WIDTH=* ALIGN=RIGHT><A href="vbscript:ReturnClick">전문가 시스템 질의</A>&nbsp;|&nbsp;<A href="vbscript:Return2Click">검사방식 적용</A></TD>
        			<!--'/* Issue: 검사방식 적용으로 Return - START */-->
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
<INPUT TYPE=HIDDEN NAME="txtInsCri" tag="24" tabindex=-1>

</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
