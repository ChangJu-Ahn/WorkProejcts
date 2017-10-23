<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Quality Management
'*  2. Function Name        : 
'*  3. Program ID           : Q1441MA1
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

Const BIZ_PGM_QRY_ID = "q1441mb1.asp"							'☆: Query 비지니스 로직 ASP명 
Const PGM_JUMP_ID1 = "q1411ma1"
'/* Issue: 검사방식 적용으로 Return - START */
Const PGM_JUMP_ID2_1 = "q1413ma1.asp"
Const PGM_JUMP_ID2_2 = "q1413ma2.asp"
Const PGM_JUMP_ID2_3 = "q1413ma3.asp"
Const PGM_JUMP_ID2_4 = "q1413ma4.asp"

Dim lgReturnPage
'/* Issue: 검사방식 적용으로 Return - END */

Dim IsOpenPop        

<!-- #Include file="../../inc/lgvariables.inc" -->

'==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'========================================================================================================= 
Sub InitVariables()

    lgIntFlgMode = Parent.OPMD_CMODE                   	'Indicates that current mode is Create mode
    lgIntGrpCount = 0        	              	'initializes Group View Size
    lgStrPrevKey = ""                           		'initializes Previous Key
    lgLngCurRows = 0                         		'initializes Deleted Rows Count
End Sub

'==========================================  2.2.1 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
'========================================================================================================= 
Sub SetDefaultVal()
	frm1.txtLotsize.Text = "<%=Request("txtLotSize")%>"
	frm1.txtSamplesize.Text = "<%=Request("txtSamplesize")%>"	
	frm1.txtAcceptCount.Text = "<%=Request("txtAcceptSize")%>"	
	frm1.txtProcessDefectRatio.Text = "<%=Request("txtDefectRate")%>"	
	frm1.txtReplaceMode.value = "<%=Request("txtReplaceMode")%>"
	
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
		Case "OA"		'OC
			Location.href = PGM_JUMP_ID2_1
		Case "SA"		'Screen
			Location.href = PGM_JUMP_ID2_2
		Case "AA"		'Adjust
			Location.href = PGM_JUMP_ID2_3
		Case "OA2"		'OC 2회 
			Location.href = PGM_JUMP_ID2_4
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
	'-----------------------
	'Check condition area
	'----------------------- 
	'⊙: This function check indispensable field
	With frm1
		.ChartFX1.ToolBar = 0					'툴바 제거하기 
		.ChartFX1.CloseData 1 Or &H800				'차트 FX와의 데이터 채널 초기화 
		.ChartFX2.ToolBar = 0					'툴바 제거하기 
		.ChartFX2.CloseData 1 Or &H800				'차트 FX와의 데이터 채널 초기화 
		.ChartFX3.ToolBar = 0					'툴바 제거하기 
		.ChartFX3.CloseData 1 Or &H800				'차트 FX와의 데이터 채널 초기화 
		.ChartFX4.ToolBar = 0					'툴바 제거하기 
		.ChartFX4.CloseData 1 Or &H800				'차트 FX와의 데이터 채널 초기화 
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

	With frm1
		.ChartFX1.ToolBar = 0					'툴바 제거하기 
		.ChartFX1.CloseData 1 Or &H800				'차트 FX와의 데이터 채널 초기화 
		.ChartFX2.ToolBar = 0					'툴바 제거하기 
		.ChartFX2.CloseData 1 Or &H800				'차트 FX와의 데이터 채널 초기화 
		.ChartFX3.ToolBar = 0					'툴바 제거하기 
		.ChartFX3.CloseData 1 Or &H800				'차트 FX와의 데이터 채널 초기화 
		.ChartFX4.ToolBar = 0					'툴바 제거하기 
		.ChartFX4.CloseData 1 Or &H800				'차트 FX와의 데이터 채널 초기화 
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
	Call parent.FncExport(Parent.C_SINGLE)					'☜: 화면 유형 
	FncExcel = True
End Function

'========================================================================================
' Function Name : FncExit
' Function Desc : This function is related to Exit 
'========================================================================================
Function FncExit()
	FncExit = True
End Function

'*******************************  5.2 Fnc함수명에서 호출되는 개발 Function  *******************************
'	설명 : 
'********************************************************************************************************* 

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbQuery() 
	
	Dim strVal
	
	Call LayerShowHide(1)

	Err.Clear              

	DbQuery = False                                                        					'⊙: Processing is NG
	
	strVal = BIZ_PGM_QRY_ID & "?txtLotSize=" & frm1.txtLotSize.Text 			'☜: ATI계산에 사용될 로트크기 
	strVal = strVal & "&txtSamplesize=" & frm1.txtSamplesize.Text				'☜: ATI계산에 사용될 샘플크기.
	strVal = strVal & "&txtProcessDefectRatio=" & frm1.txtProcessDefectRatio.Text		'☜: ATI계산에 사용될 불량률 
	strVal = strVal & "&txtAcceptCount=" & frm1.txtAcceptCount.Text			'☜: ATI계산에 사용될 합부판정개수 
	strVal = strVal & "&txtReplaceMode=" & Trim(frm1.txtReplaceMode.value)				'☜: 불량품 대체여부 
	
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
			<TABLE <%=LR_SPACE_TYPE_10%>>
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
								<TD CLASS="TD5" NOWRAP>로트크기</TD>
								<TD CLASS="TD6" NOWRAP>
									<script language =javascript src='./js/q1441ma1_txtLotSize_txtLotSize.js'></script>
								</TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP HEIGHT=2></TD>
								<TD CLASS="TD6" NOWRAP HEIGHT=2></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>샘플크기</TD>
								<TD CLASS="TD6" NOWRAP>
									<script language =javascript src='./js/q1441ma1_txtSamplesize_txtSamplesize.js'></script>
								</TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP HEIGHT=2></TD>
								<TD CLASS="TD6" NOWRAP HEIGHT=2></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>합격판정개수</TD>
								<TD CLASS="TD6" NOWRAP>
									<script language =javascript src='./js/q1441ma1_txtAcceptCount_txtAcceptCount.js'></script>
								</TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP HEIGHT=1></TD>
								<TD CLASS="TD6" NOWRAP HEIGHT=1></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>공정불량률</TD>
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
								<TD CLASS="TD5" NOWRAP>2회</TD>
								<TD CLASS="TD6" NOWRAP></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>1차 샘플크기</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtASNSamplesize1" SIZE=20 MAXLENGTH=15 tag="14" STYLE="Text-Align: Right"></TD>
							</TR>
							<TR>
								<TD HEIGHT=8% WIDTH=100%></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>2차 샘플크기</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtASNSamplesize2" SIZE=20 MAXLENGTH=15 tag="14" STYLE="Text-Align: Right"></TD>
							</TR>
							<TR>
								<TD HEIGHT=8% WIDTH=100%></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>1차 판정개수</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtASNAcceptanceCnt1" SIZE=20 MAXLENGTH=15 tag="14" STYLE="Text-Align: Right"></TD>
							</TR>
							<TR>
								<TD HEIGHT=8% WIDTH=100%></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>2차 판정개수</TD>
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
<INPUT TYPE=HIDDEN NAME="txtReplaceMode" tag="24" tabindex=-1>
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

