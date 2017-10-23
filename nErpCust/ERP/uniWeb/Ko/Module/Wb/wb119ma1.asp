
<%@ LANGUAGE="VBSCRIPT" %>
<%'======================================================================================================
'*  1. Module Name          : 법인세 
'*  2. Function Name        : 기준정보 
'*  3. Program ID           : WB119MA1
'*  4. Program Name         : WB119MA1.asp
'*  5. Program Desc         : 버전 복사 
'*  6. Modified date(First) : 2005/03/03
'*  7. Modified date(Last)  : 2005/03/03
'*  8. Modifier (First)     : 최영태 
'*  9. Modifier (Last)      : 최영태 
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../wcm/inc_SvrOperation.asp" -->
<%
'=========================  로긴중인 유저의 법인코드를 출력하기 위해  ======================
    Call LoadBasisGlobalInf()
%>
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">
<STYLE>
	.TD51 {
		BACKGROUND-COLOR: #d1e8f9;
		TEXT-ALIGN: center;
		FONT-SIZE: 9pt;
	}
	.TD61 {
		BACKGROUND-COLOR: #eeeeec;
	}
	.STATUS_FLG1 {
		color: red;
	}
	.STATUS_FLG2 {
		color: darkorange;
	}
	.STATUS_FLG3 {
		color: blue;
	}
	.link1 {
		color: black;
	}
</STYLE>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMaMain.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMaOperation.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="JavaScript" SRC="../../inc/TabScript.js"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"	SRC="../WCM/inc_CliGrid.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"	SRC="../WCM/inc_CliOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/Cookie.vbs">   </SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit

'============================================  상수/변수 선언  ====================================

<!-- #Include file="../../inc/lgvariables.inc" -->

Const BIZ_MNU_ID = "wb119ma1"
Const BIZ_PGM_ID = "wb119mb1.asp"											 '☆: 비지니스 로직 ASP명 


Dim IsOpenPop    
Dim gSelframeFlg , lgCurrGrid      
Dim lgStrPrevKey2
Dim lgRefMode	' 레퍼런스 참조상태 : ERP 추출사용유무 

'============================================  초기화 함수  ====================================
Sub InitSpreadPosVariables()

End Sub

Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
    
    lgStrPrevKey = ""                           'initializes Previous Key
    lgStrPrevKey2 = ""                          'initializes Previous Key
    lgLngCurRows = 0                            'initializes Deleted Rows Count
    lgSortKey = 1
    lgRefMode = False
End Sub

Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029A("I", "*","NOCOOKIE","MA") %>
End Sub



'============================================  신고구분 콤보 박스 채우기  ====================================

Sub InitComboBox()
	Dim IntRetCD1
	call CommonQueryRs("MINOR_CD,MINOR_NM"," B_MINOR "," MAJOR_CD = 'W1018' ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboREP_TYPE1 ,lgF0  ,lgF1  ,Chr(11))
    Call SetCombo2(frm1.cboREP_TYPE2 ,lgF0  ,lgF1  ,Chr(11))
End Sub

Sub InitSpreadSheet()
	Dim ret
	
'    Call initSpreadPosVariables()  

End Sub


'============================================  그리드 함수  ====================================

Sub InitSpreadComboBox()
End Sub

Sub SetSpreadLock()

End Sub

Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
 
End Sub

Sub GetSpreadColumnPos(ByVal pvSpdNo)
 
End Sub

Sub InitData()
	With frm1
	
	.txtFISC_YEAR1.text 	= "<%=wgFISC_YEAR%>"
	.txtCO_CD1.value 	= "<%=wgCO_CD%>"
	.txtCO_NM1.value 	= "<%=wgCO_NM%>"
	.cboREP_TYPE1.value 	= "<%=wgREP_TYPE%>"
 
	.txtFISC_YEAR2.text 	= "<%=wgFISC_YEAR-1%>"
    .txtCO_CD2.value = "<%=wgCO_CD%>"
    .txtCO_NM2.value = "<%=wgCO_NM%>"       
 
	End With
End Sub

Sub BtnCopyVer()
	Call FncSave()
End Sub
'============================================  조회조건 함수  ====================================


'====================================== 탭 함수 =========================================


'============================================  폼 함수  ====================================
Sub Form_Load()

    Call LoadInfTB19029     
                                                    <%'Load table , B_numeric_format%>
    Call ggoOper.LockField(Document, "N")                                   <%'Lock  Suitable  Field%>                         
   
    Call InitSpreadSheet                                                    <%'Setup the Spread sheet%>
    Call InitVariables                                                      <%'Initializes local global variables%>
    
    Call SetToolbar("1000000000000111")										<%'버튼 툴바 제어 %>

	' 변경한곳 
    Call InitComboBox
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.FormatDate(frm1.txtFISC_YEAR1, parent.gDateFormat,3)
	Call ggoOper.FormatDate(frm1.txtFISC_YEAR2, parent.gDateFormat,3)

	Call InitData()

     
    
End Sub


'============================================  이벤트 함수  ====================================
Sub txtFISC_YEAR_DblClick(Button)
    If Button = 1 Then
        frm1.txtFISC_YEAR.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtFISC_YEAR.Focus
    End If
End Sub


'============================================  툴바지원 함수  ====================================

Function FncNew() 
    Dim IntRetCD 

    FncNew = False

    FncNew = True

End Function

Function FncQuery() 
    Dim IntRetCD 
    
    FncQuery = False                                                        
    
    Err.Clear                                                               <%'Protect system from crashing%>

<%  '-----------------------
    'Check previous data area
    '----------------------- %>
    
End Function

Function FncSave() 
 
    FncSave = False                                                          
    
    With frm1
    
		If .cboREP_TYPE1.value = "2" Or .cboREP_TYPE2.value = "2" Then
			Call DisplayMsgBox("WC0039", "X", "X", "X")
			Exit Function
		ElseIf .txtFISC_YEAR1.text = .txtFISC_YEAR2.text And _
			.cboREP_TYPE1.value = .cboREP_TYPE2.value Then
			Call DisplayMsgBox("W20001", "X", "신고구분 및 사업연도", "X")
			Exit Function
			
		End If
    End With
    
    If DbSave = False Then Exit Function                                        '☜: Save db data
    
    FncSave = True 
    
End Function

Function FncCopy() 
 
End Function

Function FncCancel() 
                                                '☜: Protect system from crashing
End Function

Function FncInsertRow(ByVal pvRowCnt) 

End Function


Function FncDeleteRow() 

End Function

Function FncPrint()
    Call parent.FncPrint()                                                   '☜: Protect system from crashing
End Function

Function FncExcel() 
    Call parent.FncExport(parent.C_MULTI)											 <%'☜: 화면 유형 %>
End Function

Function FncFind() 
    Call parent.FncFind(parent.C_MULTI, False)                                      <%'☜:화면 유형, Tab 유무 %>
End Function

Function FncExit()
Dim IntRetCD
	
	FncExit = False

    FncExit = True
End Function

'========================================================================================
Function FncDelete() 
    Dim IntRetCD

    FncDelete = False
    If lgIntFlgMode <> parent.OPMD_UMODE Then
        Call DisplayMsgBox("900002", "X", "X", "X")
        Exit Function
    End If

    IntRetCD = DisplayMsgBox("900003",parent.VB_YES_NO, "X", "X")
    IF IntRetCD = vbNO Then
		Exit Function
	End IF

    Call DbDelete

    FncDelete = True
End Function
'============================================  DB 억세스 함수  ====================================
Function DbSave() 

    DbSave = False
    
    Err.Clear                                                               <%'☜: Protect system from crashing%>
	
	Call LayerShowHide(1)
	
	Dim strVal
    
    With Frm1
    
		strVal = BIZ_PGM_ID 
		
		Call ExecMyBizASP(frm1, strVal)   
    End With                                           '☜:  Run biz logic

    DbSave = True  
  
End Function

Function DbSaveOk()													<%'조회 성공후 실행로직 %>
	
  
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	Call DisplayMsgBox("183114", "X", "X", "X")
    '-----------------------
    'Reset variables area
    '-----------------------
		
End Function


Sub txtFISC_YEAR2_DblClick(Button)
    If Button = 1 Then
        frm1.txtFISC_YEAR2.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtFISC_YEAR2.Focus
    End If
End Sub

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->	
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
					<TD CLASS="CLSMTABP_BAK">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 width=300>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" NOWRAP><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB_BAK"><font color=white><%=Request("strASPMnuMnuNm")%></font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD align=right></TD>
					<TD WIDTH=*>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=*>
		<TD WIDTH=100% CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=100% valign=top>
					<TABLE <%=LR_SPACE_TYPE_60%>>
						<TR>
							<TD WIDTH=100% valign=top>
								<TABLE <%=LR_SPACE_TYPE_60%>>
									<TR>
										<TD CLASS="TD5">옵션</TD>
										<TD CLASS="TD6"><INPUT TYPE=RADIO CLASS=RADIO NAME=rdoTYPE ID=rdoTYPE1 checked VALUE="_Master"><LABEL FOR="rdoTYPE1">기준정보</LABEL>
										<INPUT TYPE=RADIO CLASS=RADIO NAME=rdoTYPE ID=rdoTYPE2 Value=""><LABEL FOR="rdoTYPE2">전체(기준정보+서식)</LABEL>
										<TD CLASS="TD5"></TD>
										<TD CLASS="TD6">
										</TD>
										</TD>
									</TR>
								</TABLE>

							</TD>
						</TR>
						<TR>
							<TD WIDTH=100% valign=top>
								<TABLE <%=LR_SPACE_TYPE_60%>>
								<TR>
									<TD CLASS="TD5">생성할 사업연도</TD>
									<TD CLASS="TD6"><script language =javascript src='./js/wb119ma1_txtFISC_YEAR1_txtFISC_YEAR1.js'></script>
									<TD CLASS="TD5">법인명</TD>
									<TD CLASS="TD6"><INPUT TYPE=TEXT NAME="txtCO_CD1" Size=10 tag="24">
										<INPUT TYPE=TEXT NAME="txtCO_NM1" Size=20 tag="24">
									</TD>
									</TD>
								</TR>
								<TR>
									<TD CLASS="TD5">신고구분</TD>
									<TD CLASS="TD6"><SELECT NAME="cboREP_TYPE1" ALT="신고구분" STYLE="WIDTH: 100%" tag="24X"><OPTION VALUE=""></SELECT>
									</TD>
									<TD CLASS="TD5">&nbsp;</TD>
									<TD CLASS="TD6">&nbsp;</TD>
								</TR>
								</TABLE>
									
							</TD>
						</TR>
						<TR>
							<TD WIDTH=100% valign=top>
								<TABLE <%=LR_SPACE_TYPE_60%>>
									<TR>
										<TD CLASS="TD5">가져올 사업연도</TD>
										<TD CLASS="TD6"><script language =javascript src='./js/wb119ma1_txtFISC_YEAR2_txtFISC_YEAR2.js'></script>
										<TD CLASS="TD5">법인명</TD>
										<TD CLASS="TD6"><INPUT TYPE=TEXT NAME="txtCO_CD2" Size=10 tag="24">
											<INPUT TYPE=TEXT NAME="txtCO_NM2" Size=20 tag="24">
										</TD>
										</TD>
									</TR>
									<TR>
										<TD CLASS="TD5">신고구분</TD>
										<TD CLASS="TD6"><SELECT NAME="cboREP_TYPE2" ALT="신고구분" STYLE="WIDTH: 100%" tag="23X">></SELECT>
										</TD>
										<TD CLASS="TD5">&nbsp;</TD>
										<TD CLASS="TD6">&nbsp;</TD>
									</TR>
								</TABLE>
							</TD>
						</TR>
						<TR HEIGHT=70%>
							<TD>&nbsp;</TD>
						</TR>
					</TABLE>
				</DIV>
			</TR>
		</TABLE>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_01%>>
		<TABLE <%=LR_SPACE_TYPE_30%>>
			<TR>
				<TD WIDTH=10>&nbsp;</TD>
				<TD><BUTTON NAME="btn1"  CLASS="CLSSBTN" ONCLICK="vbscript:BtnCopyVer()" Flag=1>복사 실행</BUTTON>&nbsp;
				</TD>
			</TR>
		</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

