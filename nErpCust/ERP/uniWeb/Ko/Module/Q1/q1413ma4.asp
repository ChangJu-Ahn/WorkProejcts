<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Quality Management
'*  2. Function Name        : 
'*  3. Program ID           : Q1413MA4
'*  4. Program Name         : 
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
<TITLE>T계수 규준형 2회 샘플링 검사방식 적용</TITLE>

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

Const BIZ_PGM_QRY_ID = "q1413Mb1.asp"
Const PGM_JUMP_ID1 = "q1411ma1"
Const PGM_JUMP_ID2 = "Q1441MA1.asp"
Const PGM_JUMP_ID3 = "Q1442MA1.asp"

Dim lgNextNo					'☜: 화면이 Single/SingleMulti 인경우만 해당 
Dim lgPrevNo					' ""

<!-- #Include file="../../inc/lgvariables.inc" -->

Dim lgMpsFirmDate
Dim lgLlcGivenDt								
Dim IsOpenPop          
Dim gSelframeFlg

'==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'========================================================================================================= 
Sub InitVariables()
    lgIntFlgMode = Parent.OPMD_CMODE                                               	'⊙: Indicates that current mode is Create mode
    lgIntGrpCount = 0
    IsOpenPop = False						'☆: 사용자 변수 초기화 
End Sub

'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== 
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "*", "NOCOOKIE","MA") %>
End Sub

'==========================================  2.2.1 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
'========================================================================================================= 
Sub SetDefaultVal()

End Sub

'===========================================  2.3.1 ResultClick()  ==========================================
'=	Name : ResultClick()																					=
'=	Description : Return Array to Opener Window when OK button click									=
'========================================================================================================
Function ResultClick()

	Dim strVal
	
	Call ggoOper.ClearField(Document, "1")										'⊙: Clear Contents  Field
	
	Call LayerShowHide(1)
	
	strVal = BIZ_PGM_QRY_ID & "?txtLotSize=" & frm1.txtLotSize.Text   '☜: '☆: 조회 조건 데이타 
	strVal = strVal & "&txtAlpha=" & frm1.txtAlpha.Text
	strVal = strVal & "&txtBeta=" & frm1.txtBeta.Text
	strVal = strVal & "&txtP1=" & frm1.txtP1.Text
	strVal = strVal & "&txtP2=" & frm1.txtP2.Text
	
	Call RunMyBizASP(MyBizASP, strVal)
End Function

'=========================================  2.3.2 ShowGraphClick()  ========================================
'=	Name : ShowGraphClick()																				=
'=	Description : Return Array to Opener Window for Cancel button click 								=
'========================================================================================================
Function ShowGraphClick()

	strVal = PGM_JUMP_ID2 & "?txtLotSize=" & frm1.txtLotSize.Text
	strVal = strVal & "&txtDefectRate=" & frm1.txtP2.Text			'불량률을 P2로 한다.
	strVal = strVal & "&txtSampleSize1=" & frm1.txtSampleSize1.Text   
	strVal = strVal & "&txtAcceptSize1=" & frm1.txtAcceptSize1.Text   
	strVal = strVal & "&txtSampleSize2=" & frm1.txtSampleSize2.Text   
	strVal = strVal & "&txtAcceptSize2=" & frm1.txtAcceptSize2.Text
	'/* Issue: 검사방식 적용으로 Return - START */
	strVal = strVal & "&txtPageCode=" & "OA2"
	'/* Issue: 검사방식 적용으로 Return - END */
	Location.href = strVal
End Function

'=============================================  2.3.3()  ======================================
'=	Event Name : ReturnClick
'=	Event Desc :
'========================================================================================================
Function ReturnClick()	
	PgmJump(PGM_JUMP_ID1)
End Function

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
	FncQuery = False
End Function

'========================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'========================================================================================
Function FncNew() 
	FncNew = False	
End Function

'========================================================================================
' Function Name : Fnc
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
	'-----------------------
	'Check content area
	'-----------------------
	If Not chkField(Document, "2") Then    				'⊙: Check contents area
    	Exit Function
    End If

	Call ResultClick()
	
	FncSave = True
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
' Function Desc : This function is related to Excel 
'========================================================================================
Function FncExit()
	FncExit = True
End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>
<BODY SCROLL=NO TABINDEX="-1">
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
								<TD BACKGROUND="../../../CShared/image/table/seltab_up_bg.gif"><IMG SRC="../../../CShared/image/table/seltab_up_left.gif" WIDTH="9" HEIGHT="23"></TD>
								<TD BACKGROUND="../../../CShared/image/table/seltab_up_bg.gif" ALIGN="center" CLASS="CLSMTAB"><FONT COLOR=white>계수 규준형 2회 샘플링 검사방식 적용</FONT></TD>
								<TD BACKGROUND="../../../CShared/image/table/seltab_up_bg.gif" ALIGN="right"><IMG SRC="../../../CShared/image/table/seltab_up_right.gif" WIDTH="10" HEIGHT="23"></TD>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* align=right>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR height=*>
		<TD  VALIGN="TOP" WIDTH=100% CLASS="Tab11">
			<TABLE WIDTH="100%" <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD <%=HEIGHT_TYPE_02%> WIDTH=100%> </TD>
				</TR>
				<TR>
					<TD VALIGN="top"  WIDTH="100%">
						<!--<TABLE CLASS="basicTB" CELLSPACING=0 HEIGHT=100%>-->
						<FIELDSET STYLE="margin-left:10px; margin-right:10px;">
							<LEGEND>입력사항</LEGEND>
							<TABLE WIDTH="100%" <%=LR_SPACE_TYPE_20%>>
								<TR>
									<TD CLASS="TD5" HEIGHT=5 NOWRAP></TD>
									<TD CLASS="TD6" HEIGHT=5 NOWRAP></TD>
									<TD CLASS="TD5" HEIGHT=5 NOWRAP></TD>
									<TD CLASS="TD6" HEIGHT=5 NOWRAP></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>생산자 위험(α)</TD>
									<TD CLASS="TD6" NOWRAP>
										<script language =javascript src='./js/q1413ma4_txtAlpha_txtAlpha.js'></script>
									</TD>
									<TD CLASS="TD5" NOWRAP>소비자 위험(β)</TD>
									<TD CLASS="TD6" NOWRAP>
										<script language =javascript src='./js/q1413ma4_txtBeta_txtBeta.js'></script>
									</TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>P1(AQL)</TD>
									<TD CLASS="TD6" NOWRAP>
										<script language =javascript src='./js/q1413ma4_txtP1_txtP1.js'></script>
									</TD>
									<TD CLASS="TD5" NOWRAP>P2(LTPD)</TD>
									<TD CLASS="TD6" NOWRAP>
										<script language =javascript src='./js/q1413ma4_txtP2_txtP2.js'></script>
									</TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>로트크기</TD>
									<TD CLASS="TD6" NOWRAP>
										<script language =javascript src='./js/q1413ma4_txtLotSize_txtLotSize.js'></script>
									</TD>
									<TD CLASS="TD5" NOWRAP></TD>
									<TD CLASS="TD6" NOWRAP></TD>
								</TR>
							</TABLE>
						</FIELDSET>
						<FIELDSET STYLE="margin-left:10px; margin-right:10px;">
							<LEGEND>결과</LEGEND>
							<TABLE CLASS="basicTB" CELLSPACING=0 HEIGHT=100%>
								<TR>
									<TD CLASS="TD5" NOWRAP>1차 샘플크기</TD>
									<TD CLASS="TD6" NOWRAP>
										<script language =javascript src='./js/q1413ma4_txtSampleSize1_txtSampleSize1.js'></script>
									</TD>
									<TD CLASS="TD5" NOWRAP>1차 합격판정개수</TD>
									<TD CLASS="TD6" NOWRAP>
										<script language =javascript src='./js/q1413ma4_txtAcceptSize1_txtAcceptSize1.js'></script>
									</TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>2차 샘플크기</TD>
									<TD CLASS="TD6" NOWRAP>
										<script language =javascript src='./js/q1413ma4_txtSampleSize2_txtSampleSize2.js'></script>
									</TD>
									<TD CLASS="TD5" NOWRAP>2차 합격판정개수</TD>
									<TD CLASS="TD6" NOWRAP>
										<script language =javascript src='./js/q1413ma4_txtAcceptSize2_txtAcceptSize2.js'></script>
									</TD>
								</TR>
							</TABLE>
						</FIELDSET>
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
					<TD><BUTTON NAME="btnResult" CLASS="CLSMBTN" ONCLICK="vbscript:ResultClick()">결과 보기</BUTTON></TD>
					<TD WIDTH=* ALIGN=RIGHT><A href="vbscript:ShowGraphClick">검사특성 그래프</A>&nbsp;|&nbsp;<A href="vbscript:ReturnClick()">전문가 시스템 질의</A></TD>
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
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
