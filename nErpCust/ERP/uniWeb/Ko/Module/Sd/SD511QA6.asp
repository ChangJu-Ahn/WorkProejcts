<%@ LANGUAGE="VBSCRIPT" %>
<%
'======================================================================================================
'*  1. Module Name          : 영업 
'*  2. Function Name        : 영업실적관리 
'*  3. Program ID           : SD511QA6
'*  4. Program Name         : 미수금현황조회(수금처)
'*  5. Program Desc         : 
'*  6. Comproxy List        :
'*  7. Modified date(First) : 2003/04/24
'*  8. Modified date(Last)  : 2003/06/09
'*  9. Modifier (First)     : Hwang SeongBae
'* 10. Modifier (Last)      : Hwang Seongbae
'* 11. Comment              :
'=======================================================================================================
%>
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMaMain.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMaOperation.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentA.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentVariables.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"> </SCRIPT>

<Script Language="VBScript">
Option Explicit                             '☜: indicates that All variables must be declared in advance
	
' External ASP File
'========================================

Const BIZ_PGM_ID 		= "SD511QB6.asp"                              '☆: Biz Logic ASP Name

' Constant variables 
'========================================
Const C_MaxKey          = 18				                          '☆: SpreadSheet의 키의 갯수 

' Common variables 
'========================================
<!-- #Include file="../../inc/lgvariables.inc" -->	

' User-defind Variables
'========================================
Dim lgIsOpenPop                                          

Dim lgStrColorFlag

<% 
   BaseDate     = GetSvrDate                                                         'Get DB Server Date
%>  

Dim FirstDateOfDB 

FirstDateOfDB	= UNIConvDateAToB(UNIGetFirstDay("<%=BaseDate%>",parent.gServerDateFormat), parent.gServerDateFormat,parent.gDateFormat)

Dim lgStartRow
Dim lgEndRow

Const C_PopBizArea	=	0
Const C_PopPayer	=	1

'========================================	
Sub InitVariables()

    lgPageNo     = ""
    lgIntFlgMode     = Parent.OPMD_CMODE                      'Indicates that current mode is Create mode
    lgSortKey        = 1

    Call SetToolBar("1100000000001111")
    
End Sub

'========================================
Sub SetDefaultVal()

	Frm1.txtConFromDt.Text		= cstr(FirstDateOfDB)
		
End Sub							
	
'========================================
Sub LoadInfTB19029()

	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "S","NOCOOKIE","QA") %>                                 
	<% Call LoadBNumericFormatA("Q", "S", "NOCOOKIE", "QA") %>
End Sub

'==========================================
Sub InitComboBox()
	' 승인상태 
	Call CommonQueryRs(" MINOR_CD, MINOR_NM ", " B_MINOR ", " MAJOR_CD = " & FilterVar("A1007", "''", "S") & " ORDER BY MINOR_NM ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	Call SetCombo2(frm1.cboConfFg,lgF0,lgF1,parent.gColSep)
End Sub

'========================================
Sub InitSpreadSheet()

    Call SetZAdoSpreadSheet("SD511QA6","S","A", "V20030424", parent.C_SORT_DBAGENT,frm1.vspdData,C_MaxKey, "X", "X")
    Call SetSpreadLock()   
     
End Sub

'========================================
Sub SetSpreadLock()

    With frm1
		.vspdData.ReDraw = False
		ggoSpread.SpreadLock 1 , -1
		.vspdData.ReDraw = True
    End With
    
End Sub

'========================================
Sub Form_Load()

    Call LoadInfTB19029				                                           '⊙: Load table , B_numeric_format

    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
    Call ggoOper.LockField(Document, "N")

	Call InitComboBox()
	Call InitVariables														
	Call SetDefaultVal	
	Call InitSpreadSheet()

    Call ggoOper.FormatDate(frm1.txtConFromDt, Parent.gDateFormat, 3)			'YYYY으로 포멧팅 
    
    Call SetFocusToDocument("M")	
    frm1.txtConFromDt.focus
End Sub

'========================================
Sub Form_QueryUnload(Cancel , UnloadMode )
 
End Sub

'========================================
Function FncQuery() 

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncQuery = False                                                              '⊙: Processing is NG
    
    If Not chkField(Document, "1") Then Exit Function

    Call ggoOper.ClearField(Document, "2")
    
    Call InitVariables
    
    If DbQuery Then Exit Function

    If Err.number = 0 Then FncQuery = True

    Set gActiveElement = document.ActiveElement  
End Function

'========================================
Function FncPrint()

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncPrint = False                                                              '☜: Processing is NG
	Call Parent.FncPrint()                                                        '☜: Protect system from crashing

    If Err.number = 0 Then
       FncPrint = True                                                            '⊙: Processing is OK
    End If   

    Set gActiveElement = document.ActiveElement   

End Function

'========================================
Function FncExcel() 

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncExcel = False                                                              '☜: Processing is NG

	Call Parent.FncExport(parent.C_MULTI)
    
    If Err.number = 0 Then
       FncExcel = True                                                             '⊙: Processing is OK
    End If   

    Set gActiveElement = document.ActiveElement   

End Function

'========================================
Function FncFind() 

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncFind = False                                                               '☜: Processing is NG

	Call Parent.FncFind(parent.C_MULTI, True)
    
    If Err.number = 0 Then
       FncFind = True                                                             '⊙: Processing is OK
    End If   

    Set gActiveElement = document.ActiveElement   

End Function

'========================================
Function FncExit()
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncExit = True                                                             '⊙: Processing is OK
End Function

'========================================
Function DbQuery() 

	Dim strVal

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    DbQuery = False

	Call LayerShowHide(1)

    If lgIntFlgMode  <> parent.OPMD_UMODE Then									'☜: This means that it is first search
		With frm1
			.txtHConLastYear.value	= CDbl(.txtConFromDt.text) - 1
			.txtHConCurrYear.value	= .txtConFromDt.text
			.txtHConFromDt.value	= UniConvYYYYMMDDToDate(parent.gDateFormat, .txtHConLastYear.value, "01", "01")
			.txtHConToDt.value		= UniConvYYYYMMDDToDate(parent.gDateFormat, .txtConFromDt.text, "12", "31")
			.txtHConBizAreaCd.value	= .txtConBizAreaCd.value
			.txtHConPayerCd.value	= .txtConPayerCd.value
			.txtHConConfFg.Value    = .cboConfFg.value				

			.txtHlgSelectListDT.value	= GetSQLSelectListDataType("A") 
			.txtHlgTailList.value		= MakeSQLGroupOrderByList("A")
			.txtHlgSelectList.value		= EnCoding(GetSQLSelectList("A"))
				
		End With	
    End If    
        
    lgStartRow = frm1.vspdData.MaxRows + 1										'포멧팅 적용하는 시작Row
	Call ExecMyBizASP(frm1, BIZ_PGM_ID)

    If Err.number = 0 Then
       DbQuery = True																'⊙: Processing is OK
    End If   

    Set gActiveElement = document.ActiveElement   

End Function

'========================================
Function DbQueryOk()												

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    lgIntFlgMode     = parent.OPMD_UMODE										  '⊙: Indicates that current mode is Update mode
    
    Call SetQuerySpreadColor
    Call SetToolBar("1100000000011111")

	frm1.vspdData.Focus
    Set gActiveElement = document.ActiveElement   

End Function

'========================================
Function OpenConPopup(ByVal pvIntWhere)

	Dim iArrRet
	Dim iArrParam(5), iArrField(6), iArrHeader(6)

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	Select Case pvIntWhere
	'사업장 
	Case C_PopBizArea									
		iArrParam(1) = "B_BIZ_AREA"									' TABLE 명칭 
		iArrParam(2) = Trim(frm1.txtConBizAreaCd.value)				' Code Condition
		iArrParam(3) = ""											' Name Cindition
		iArrParam(4) = ""											' Where Condition
		iArrParam(5) = frm1.txtConBizAreaCd.alt						' TextBox 명칭 
		
		iArrField(0) = "ED15" & Parent.gColSep & "BIZ_AREA_CD"		' Field명(0)
		iArrField(1) = "ED30" & Parent.gColSep & "BIZ_AREA_NM"		' Field명(1)
    
	    iArrHeader(0) = frm1.txtConBizAreaCd.alt					' Header명(0)
	    iArrHeader(1) = frm1.txtConBizAreaNm.alt					' Header명(1)

		frm1.txtConBizAreaCd.focus 
	'수금처		
	Case C_PopPayer
		iArrParam(1) = "B_BIZ_PARTNER_FTN PF, B_BIZ_PARTNER PA"
		iArrParam(2) = Trim(frm1.txtConPayerCd.value)
		iArrParam(3) = ""
		iArrParam(4) = "PF.USAGE_FLAG = " & FilterVar("Y", "''", "S") & "  AND PF.PARTNER_FTN= " & FilterVar("SPA", "''", "S") & "" _
						& "AND PA.BP_CD = PF.PARTNER_BP_CD AND PA.BP_TYPE <=" & FilterVar("CS", "''", "S") & ""
		iArrParam(5) = frm1.txtConPayerCd.alt
		
		iArrField(0) = "ED15" & Parent.gColSep & "PA.BP_CD"
		iArrField(1) = "ED30" & Parent.gColSep & "PA.BP_NM"
    
	    iArrHeader(0) = frm1.txtConPayerCd.alt
	    iArrHeader(1) = frm1.txtConPayerNm.alt

		frm1.txtConPayerCd.focus

	End Select
	
	iArrParam(0) = iArrParam(5)										' 팝업 명칭 
	
	iArrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(iArrParam, iArrField, iArrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False
	
	If iArrRet(0) = "" Then
		Exit Function
	Else
		OpenConPopup = SetConPopup(iArrRet,pvIntWhere)
	End If	

End Function

'========================================
Function SetConPopup(Byval pvArrRet,ByVal pvIntWhere)

	SetConPopup = False

	With frm1
		Select Case pvIntWhere

		Case C_PopBizArea
			.txtConBizAreaCd.value = pvArrRet(0) 
			.txtConBizAreaNm.value = pvArrRet(1)   			
		Case C_PopPayer
			.txtConPayerCd.value = pvArrRet(0) 
			.txtConPayerNm.value = pvArrRet(1) 

		End Select
	End With

	SetConPopup = True		
	
End Function

'========================================
Sub SetQuerySpreadColor()

	Dim iArrColor1, iArrColor2
	Dim iLoopCnt
	Dim Spread
	
	Set Spread = frm1.vspdData
	
	iArrColor1 = Split(lgStrColorFlag,Parent.gRowSep)

	For iLoopCnt=0 to ubound(iArrColor1,1) - 1
		iArrColor2 = Split(iArrColor1(iLoopCnt),Parent.gColSep)

		Spread.Col = -1
		Spread.Row =  iArrColor2(0)
		
		Select Case iArrColor2(1)
			Case "1"
				Spread.BackColor = RGB(204,255,153) '연두 
			Case "2"
				Spread.BackColor = RGB(176,234,244) '하늘색 
			Case "3"
				Spread.BackColor = RGB(224,206,244) '연보라 
			Case "4"  
				Spread.BackColor = RGB(251,226,153) '연주황 
			Case "5" 
				Spread.BackColor = RGB(255,255,153) '연노랑 
		End Select
	Next

End Sub

'========================================
Sub PopZAdoConfigGrid()

  If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
     Exit Sub
  End If

  Call OpenOrderBy("A")
  
End Sub

'========================================
Sub OpenOrderBy(ByVal pvPsdNo)
	Dim arrRet
	
	If lgIsOpenPop = True Then Exit Sub
	lgIsOpenPop = True

	arrRet = window.showModalDialog("../../ComAsp/ZADOSortPopup.asp",Array(ggoSpread.GetXMLData(pvPsdNo),gMethodText),"dialogWidth=" & parent.SORTW_WIDTH & "px; dialogHeight=" & parent.SORTW_HEIGHT & "px;; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False
	
	If arrRet(0) = "X" Then												' Means that nothing is happened!!!
	   Exit Sub
	Else
	   Call ggoSpread.SaveXMLData(pvPsdNo,arrRet(0),arrRet(1))
       Call InitVariables
       Call InitSpreadSheet()       
   End If
End Sub

'========================================
Sub vspdData_Click(ByVal Col, ByVal Row)
    Call SetPopupMenuItemInf("00000000001")
    gMouseClickStatus = "SPC"   

    Set gActiveSpdSheet = frm1.vspdData
    
    If Frm1.vspdData.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If

    If Row <= 0 Then
        ggoSpread.Source = frm1.vspdData
        If lgSortKey = 1 Then
            ggoSpread.SSSort Col
            lgSortKey = 2
        Else
            ggoSpread.SSSort Col, lgSortKey
            lgSortKey = 1
        End If
        Exit Sub
    End If
    
End Sub

'========================================
Sub vspdData_GotFocus()
    ggoSpread.Source = Frm1.vspdData
End Sub

'========================================
Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If
End Sub    

'========================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If

	If CheckRunningBizProcess = True Then
	   Exit Sub
	End If
    
	If Frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(Frm1.vspdData,NewTop) Then	    
    	If lgPageNo <> "" Then								
           Call DisableToolBar(parent.TBC_QUERY)
           Call DbQuery
    	End If
    End If
    
End Sub

'========================================
Sub txtConFromDt_DblClick(Button)
	If Button = 1 Then
       Frm1.txtConFromDt.Action = 7
       Call SetFocusToDocument("M")	
       Frm1.txtConFromDt.Focus
	End If
End Sub

'========================================
Sub txtConToDt_DblClick(Button)
	If Button = 1 Then
       Frm1.txtConToDt.Action = 7
       Call SetFocusToDocument("M")	
       Frm1.txtConToDt.Focus
	End If
End Sub

'========================================
Sub txtConFromDt_KeyPress(KeyAscii)
	If KeyAscii = 13 Then 
	   Call MainQuery()
	End If   
End Sub

'========================================
Sub txtConToDt_KeyPress(KeyAscii)
	If KeyAscii = 13 Then 
	   Call MainQuery()
	End If   
End Sub

</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
</HEAD>

<!--
'########################################################################################################
'#						6. TAG 부																		#
'######################################################################################################## 
-->
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>미수금조회(수금처)</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=*>&nbsp;</TD>
					<TD WIDTH="30" align=right></td>
					<TD WIDTH=10>&nbsp;</TD>
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
						<FIELDSET>
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS="TD5" NOWRAP>기준년도</TD>									
									<TD CLASS="TD6" NOWRAP>							        
									<TABLE CELLSPACING=0 CELLPADDING=0>
										<TR>
											<TD>
											<script language =javascript src='./js/sd511qa6_OBJECT1_txtConFromDt.js'></script>
											</TD>
											<TD>
										</TR>
									</TABLE>							        
							        </TD>
									<TD CLASS="TD5" NOWRAP>승인상태</TD>
									<TD CLASS="TD6"><SELECT Name="cboConfFg" ID="cboConfFg" ALT="승인상태" tag="11XXXU"><OPTION VALUE=""></OPTION></SELECT></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>사업장</TD>
									<TD CLASS="TD6" NOWRAP>	<INPUT TYPE=TEXT NAME="txtConBizAreaCd" SIZE=10 MAXLENGTH=10 tag="11XXXU" ALT="사업장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnConBizArea" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenConPopUp(C_PopBizArea) ">
															<INPUT TYPE=TEXT NAME="txtConBizAreaNm" SIZE=20 tag="14" ALT="사업장명"></TD>
									<TD CLASS="TD5" NOWRAP>수금처</TD>
									<TD CLASS="TD6" NOWRAP>	
															<INPUT TYPE=TEXT NAME="txtConPayerCd" SIZE=10 MAXLENGTH=10 tag="11XXXU" ALT="수금처"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnConPayer" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenConPopUp(C_PopPayer) ">
															<INPUT TYPE=TEXT NAME="txtConPayerNm" SIZE=20 tag="14" ALT="수금처명"></TD>
								</TR>								
							</TABLE>
						</FIELDSET>
					</TD>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=* valign=top>
						<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD HEIGHT="100%">
									<script language =javascript src='./js/sd511qa6_vspdData_vspdData.js'></script>
								</TD>
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
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>

<INPUT TYPE=HIDDEN NAME="txtHConLastYear"	tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHConCurrYear"	tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHConFromDt"		tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHConToDt"		tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHConBizAreaCd"	tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHConPayerCd"	tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHConConfFg"		tag="24" TABINDEX="-1">

<INPUT TYPE=HIDDEN NAME="txtHlgSelectListDT"	tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHlgTailList"		tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHlgSelectList"		tag="24" TABINDEX="-1">				

</FORM>
    <DIV ID="MousePT" NAME="MousePT">
    <iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
    </DIV>
</BODY>
</HTML>
