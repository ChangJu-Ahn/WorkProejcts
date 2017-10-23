<%@ LANGUAGE="VBSCRIPT" %>
<%
'======================================================================================================
'*  1. Module Name          : 영업 
'*  2. Function Name        : 매출실적관리 
'*  3. Program ID           : SD511QA2
'*  4. Program Name         : 매출채권조회(판매유형1)
'*  5. Program Desc         : 
'*  6. Comproxy List        :
'*  7. Modified date(First) : 2003/04/02
'*  8. Modified date(Last)  : 2003/06/09
'*  9. Modifier (First)     :
'* 10. Modifier (Last)      : Hwang Seongbae
'* 11. Comment              : 표준반영 
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

<Script Language="VBScript">
Option Explicit                             '☜: indicates that All variables must be declared in advance
	

' External ASP File
'========================================

Const BIZ_PGM_ID 		= "SD511QB2.asp"                              '☆: Biz Logic ASP Name

' Constant variables 
'========================================
Const C_MaxKey          = 18				                          '☆: SpreadSheet의 키의 갯수 

Const C_PopBizArea		=	0
Const C_PopSalesGrp		=	1
Const C_PopSalesType	=	2

' Common variables 
'========================================
<!-- #Include file="../../inc/lgvariables.inc" -->	

' User-defind Variables
'========================================
Dim lgIsOpenPop                                          
Dim lgStrColorFlag
Dim lgStartRow
Dim lgEndRow

Dim ToDateOfDB

ToDateOfDB = UNIConvDateAToB("<%=GetSvrDate%>", parent.gServerDateFormat,parent.gDateFormat)

'========================================	
Sub InitVariables()

    lgPageNo     = ""
    lgIntFlgMode     = Parent.OPMD_CMODE                      'Indicates that current mode is Create mode
    lgSortKey        = 1

    Call SetToolBar("1100000000001111")
End Sub

'========================================

Sub SetDefaultVal()

    Dim iYear,iMonth,iDay
    
    Call ExtractDateFrom("<%=GetSvrDate%>", parent.gServerDateFormat,parent.gServerDateType ,iYear,iMonth,iDay)

	Frm1.txtConYMFromDt.Year  = iYear
	Frm1.txtConYMToDt.Year    = iYear
		
End Sub							
									
'========================================
Sub LoadInfTB19029()

	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "S","NOCOOKIE","QA") %>                                 
	<% Call LoadBNumericFormatA("Q", "S", "NOCOOKIE", "QA") %>
End Sub

'========================================
Sub InitSpreadSheet()

    Call SetZAdoSpreadSheet("SD511QA2","S","A", "V20030507", parent.C_SORT_DBAGENT,frm1.vspdData,C_MaxKey, "X", "X")
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
    Call ggoOper.FormatDate(frm1.txtConYMFromDt, Parent.gDateFormat, 3)			'YYYYMM으로 포멧팅 
    Call ggoOper.FormatDate(frm1.txtConYMToDt, Parent.gDateFormat, 3)

	Call InitVariables														
	Call SetDefaultVal	
	Call InitSpreadSheet()
    

    Call SetFocusToDocument("M")	
    frm1.txtConYMFromDt.focus
End Sub

'========================================
Sub Form_QueryUnload(Cancel , UnloadMode )
 
End Sub

'========================================
Function FncQuery() 

    On Error Resume Next
    Err.Clear

    FncQuery = False

    If Not chkField(Document, "1") Then Exit Function

	If frm1.txtConYMFromDt.Text > frm1.txtConYMToDt.Text Then
		Call DisplayMsgBox("970023","X", frm1.txtConYMToDt.Alt, frm1.txtConYMFromDt.Alt)
	    Call SetFocusToDocument("M")	
		frm1.txtConYMToDt.Focus
		Set gActiveElement = document.activeElement                            
		Exit Function
	End If
		
    Call ggoOper.ClearField(Document, "2")
    
    Call InitVariables
    
    If DbQuery Then Exit Function

    If Err.number = 0 Then FncQuery = True

    Set gActiveElement = document.ActiveElement  
    
End Function

'========================================
Function FncPrint()

    On Error Resume Next
    Err.Clear

    FncPrint = False
	Call Parent.FncPrint()

    If Err.number = 0 Then
       FncPrint = True
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
	FncExit = True                                                             '⊙: Processing is OK
End Function

'========================================
Function DbQuery() 

	Dim strVal
	Dim strBillConfFlag
	Dim strExceptFlag
	
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    DbQuery = False
    
	Call LayerShowHide(1)

    strVal = BIZ_PGM_ID
        
	If frm1.rdoBillConfFlag(0).checked Then
		strBillConfFlag=""
	ElseIf frm1.rdoBillConfFlag(1).checked Then
		strBillConfFlag="Y"
	ElseIf frm1.rdoBillConfFlag(2).checked Then
		strBillConfFlag="N"
	End If
		
	If frm1.rdoExceptFlag(0).checked Then
		strExceptFlag=""
	ElseIf frm1.rdoExceptFlag(1).checked Then
		strExceptFlag="Y"
	ElseIf frm1.rdoExceptFlag(2).checked Then
		strExceptFlag="N"
	End If

    If lgIntFlgMode  <> parent.OPMD_UMODE Then									'☜: This means that it is first search
			
		'원래는 Get방식이나 조건부가 많으면 POST방식으로 넘김 
		With frm1

			.txtHConYMFromDt.value	= .txtConYMFromDt.text
			.txtHConYMToDt.value	= .txtConYMToDt.text
			
			.txtHConBizAreaCd.value		= .txtConBizAreaCd.value
			.txtHConSalesGrpCd.value	= .txtConSalesGrpCd.value
			.txtHConSalesTypeCd.value	= .txtConSalesTypeCd.value
				
			.txtHlgSelectListDT.value	= GetSQLSelectListDataType("A") 
			.txtHlgTailList.value		= MakeSQLGroupOrderByList("A")
			.txtHlgSelectList.value		= EnCoding(GetSQLSelectList("A"))
				
			.txtHConBillConfFlag.value	= strBillConfFlag
			.txtHExceptFlag.value		= strExceptFlag
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
		iArrParam(1) = "B_BIZ_AREA"									
		iArrParam(2) = Trim(frm1.txtConBizAreaCd.value)				
		iArrParam(3) = ""											
		iArrParam(4) = ""											
		iArrParam(5) = frm1.txtConBizAreaCd.alt						
		
		iArrField(0) = "ED15" & Parent.gColSep & "BIZ_AREA_CD"		
		iArrField(1) = "ED30" & Parent.gColSep & "BIZ_AREA_NM"		
    
	    iArrHeader(0) = frm1.txtConBizAreaCd.alt					
	    iArrHeader(1) = frm1.txtConBizAreaNm.alt					

		frm1.txtConBizAreaCd.focus 
	'영업그룹 
	Case C_PopSalesGrp	
		iArrParam(1) = "B_SALES_GRP"								
		iArrParam(2) = Trim(frm1.txtConSalesGrpCd.value)			
		iArrParam(3) = ""											
		iArrParam(4) = "USAGE_FLAG = " & FilterVar("Y", "''", "S") & " "							
		iArrParam(5) = frm1.txtConSalesGrpCd.alt					
		
		iArrField(0) = "ED15" & Parent.gColSep & "SALES_GRP"		
		iArrField(1) = "ED30" & Parent.gColSep & "SALES_GRP_NM"		
    
	    iArrHeader(0) = frm1.txtConSalesGrpCd.alt					
	    iArrHeader(1) = frm1.txtConSalesGrpNm.alt					

		frm1.txtConSalesGrpCd.focus 
	'판매유형 
	Case C_PopSalesType
		iArrParam(1) = "B_MINOR"									
		iArrParam(2) = Trim(frm1.txtConSalesTypeCd.value)			
		iArrParam(3) = ""											
		iArrParam(4) = "MAJOR_CD = " & FilterVar("S0001", "''", "S") & ""							
		iArrParam(5) = frm1.txtConSalesTypeCd.alt					
		
		iArrField(0) = "ED15" & Parent.gColSep & "MINOR_CD"			
		iArrField(1) = "ED30" & Parent.gColSep & "MINOR_NM"			
    
	    iArrHeader(0) = frm1.txtConSalesTypeCd.alt					
	    iArrHeader(1) = frm1.txtConSalesTypeNm.alt					

		frm1.txtConSalesTypeCd.focus

	End Select
	
	iArrParam(0) = iArrParam(5)										
	
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
		Case C_PopSalesGrp
			.txtConSalesGrpCd.value = pvArrRet(0)
			.txtConSalesGrpNm.value = pvArrRet(1)
		Case C_PopSalesType
			.txtConSalesTypeCd.value = pvArrRet(0) 
			.txtConSalesTypeNm.value = pvArrRet(1)  

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

'==================================================================================
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
Sub txtConYMFromDt_DblClick(Button)
	If Button = 1 Then
       Frm1.txtConYMFromDt.Action = 7
       Call SetFocusToDocument("M")	
       Frm1.txtConYMFromDt.Focus
	End If
End Sub

'========================================
Sub txtConYMToDt_DblClick(Button)
	If Button = 1 Then
       Frm1.txtConYMToDt.Action = 7
       Call SetFocusToDocument("M")	
       Frm1.txtConYMToDt.Focus
	End If
End Sub

'========================================
Sub txtConYMFromDt_KeyPress(KeyAscii)
	If KeyAscii = 13 Then 
	   Call MainQuery()
	End If   
End Sub

'========================================
Sub txtConYMToDt_KeyPress(KeyAscii)
	If KeyAscii = 13 Then 
	   Call MainQuery()
	End If   
End Sub

</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>매출채권조회(판매유형1)</font></td>
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
									<TD CLASS="TD5" NOWRAP>매출년도</TD>									
									<TD CLASS="TD6" NOWRAP>							        
									<TABLE CELLSPACING=0 CELLPADDING=0>
										<TR>
											<TD>
											<script language =javascript src='./js/sd511qa2_OBJECT1_txtConYMFromDt.js'></script>
											</TD>
											<TD>
											&nbsp;~&nbsp;
											</TD>
											<TD>
											<script language =javascript src='./js/sd511qa2_OBJECT2_txtConYMToDt.js'></script>
											</TD>
										</TR>
									</TABLE>							        
							        </TD>
									<TD CLASS="TD5" NOWRAP>사업장</TD>
									<TD CLASS="TD6" NOWRAP>	<INPUT TYPE=TEXT NAME="txtConBizAreaCd" SIZE=10 MAXLENGTH=10 tag="11NXXU" ALT="사업장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnConBizArea" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenConPopUp(C_PopBizArea) ">
															<INPUT TYPE=TEXT NAME="txtConBizAreaNm" SIZE=20 tag="14" ALT="사업장명"></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>영업그룹</TD>
									<TD CLASS="TD6" NOWRAP>	<INPUT TYPE=TEXT NAME="txtConSalesGrpCd" SIZE=10 MAXLENGTH=5 tag="11NXXU" ALT="영업그룹"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnConSalesGrp" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenConPopUp(C_PopSalesGrp) ">
															<INPUT TYPE=TEXT NAME="txtConSalesGrpNm" SIZE=20 tag="14" ALT="영업그룹명"></TD>
									<TD CLASS="TD5" NOWRAP>판매유형</TD>
									<TD CLASS="TD6" NOWRAP>	<INPUT TYPE=TEXT NAME="txtConSalesTypeCd" SIZE=10 MAXLENGTH=5 tag="11NXXU" ALT="판매유형"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnConSalesType" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenConPopUp(C_PopSalesType)  ">
															<INPUT TYPE=TEXT NAME="txtConSalesTypeNm" SIZE=20 tag="14" ALT="판매유형명"></TD>
								</TR>								
								<TR>
									<TD CLASS=TD5 NOWRAP>확정여부</TD>
									<TD CLASS=TD6 NOWRAP>
										<input type=radio CLASS="RADIO" name="rdoBillConfFlag" id="rdoAll" value="A" tag = "11X" checked>
											<label for="rdoAll">전체</label>&nbsp;&nbsp;
										<input type=radio CLASS="RADIO" name="rdoBillConfFlag" id="rdoConf" value="S" tag = "11X">
											<label for="rdoConf">확정</label>&nbsp;&nbsp;
										<input type=radio CLASS = "RADIO" name="rdoBillConfFlag" id="rdoNonConf" value="D" tag = "11X">
											<label for="rdoNonConf">미확정</label>
									</TD>
									<TD CLASS=TD5 NOWRAP>예외여부</TD>
									<TD CLASS=TD6 NOWRAP>
										<input type=radio CLASS="RADIO" name="rdoExceptFlag" id="rdoAll1" value="A" tag = "11X" checked>
											<label for="rdoAll1">전체</label>&nbsp;&nbsp;
										<input type=radio CLASS="RADIO" name="rdoExceptFlag" id="rdoExcept" value="Y" tag = "11X">
											<label for="rdoExcept">예외</label>&nbsp;&nbsp;
										<input type=radio CLASS = "RADIO" name="rdoExceptFlag" id="rdoNormal" value="N" tag = "11X">
											<label for="rdoNormal">정상</label>
									</TD>
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
									<script language =javascript src='./js/sd511qa2_vspdData_vspdData.js'></script>
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

<INPUT TYPE=HIDDEN NAME="txtHConYMFromDt"   tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHConYMToDt"		tag="24" TABINDEX="-1">

<INPUT TYPE=HIDDEN NAME="txtHConBizAreaCd"		tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHConSalesGrpCd"		tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHConSalesTypeCd"	tag="24" TABINDEX="-1">

<INPUT TYPE=HIDDEN NAME="txtHlgSelectListDT"	tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHlgTailList"		tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHlgSelectList"		tag="24" TABINDEX="-1">				

<INPUT TYPE=HIDDEN NAME="txtHConBillConfFlag"	tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHExceptFlag"		tag="24" TABINDEX="-1">

</FORM>
    <DIV ID="MousePT" NAME="MousePT">
    <iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
    </DIV>
</BODY>
</HTML>
