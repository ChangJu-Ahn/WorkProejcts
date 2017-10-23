<%@ LANGUAGE="VBSCRIPT" %>
<%
'=====================================================================================================
'*  1. Module Name          : 영업 
'*  2. Function Name        : 영업실적관리 
'*  3. Program ID           : SD512QA4
'*  4. Program Name         : 매출채권집계조회(영업그룹)
'*  5. Program Desc         : 
'*  6. Comproxy List        :
'*  7. Modified date(First) : 2003/04/25
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

<Script Language="VBScript">
Option Explicit

' External ASP File
'========================================
Const BIZ_PGM_ID 		= "SD512QB4.asp"

' Constant variables 
'========================================
Const C_MaxKey          = 19

Const C_PopSalesGrp		= 0		'영업그룹 
Const C_PopBizArea		= 1		'사업장 

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
    lgIntFlgMode = Parent.OPMD_CMODE
    lgSortKey    = 1

    Call SetToolBar("1100000000001111")										
    
End Sub

'========================================
Sub SetDefaultVal()
	Frm1.txtConFromDt.Text	= ToDateOfDB
	Frm1.txtConToDt.Text	= ToDateOfDB
End Sub							
	
'========================================
Sub LoadInfTB19029()

	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->

	<% Call loadInfTB19029A("Q", "S","NOCOOKIE","QA") %>                                 
	<% Call LoadBNumericFormatA("Q", "S", "NOCOOKIE", "QA") %>
End Sub

'========================================
Sub InitSpreadSheet()

    Call SetZAdoSpreadSheet("SD512QA4","S","A", "V20030425", parent.C_SORT_DBAGENT,frm1.vspdData,C_MaxKey, "X", "X")
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

	Call InitVariables														
	Call SetDefaultVal	
	Call InitSpreadSheet()
    
    Call ggoOper.FormatDate(frm1.txtConFromDt, Parent.gDateFormat, 3)		'년도 
    Call ggoOper.FormatDate(frm1.txtConToDt, Parent.gDateFormat, 3)			'년도 

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

	If frm1.txtConFromDt.Text > frm1.txtConToDt.Text Then
		Call DisplayMsgBox("970023","X", frm1.txtConToDt.Alt, frm1.txtConFromDt.Alt)
	    Call SetFocusToDocument("M")	
		frm1.txtConToDt.Focus
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

    Set gActiveElement = document.ActiveElement   

End Function

'========================================
Function DbQuery() 

	Dim iStrVal

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    DbQuery = False
    
	Call LayerShowHide(1)

    If lgIntFlgMode  <> parent.OPMD_UMODE Then									'☜: This means that it is first search
		With frm1			
			iStrVal = BIZ_PGM_ID & "?ConFromDt=" & UniConvYYYYMMDDToDate(parent.gDateFormat, .txtConFromDt.Text, "01", "01") & _
									"&ConToDt=" & UniConvYYYYMMDDToDate(parent.gDateFormat, .txtConToDt.Text, "12", "31") & _
									"&SalesGrpCd=" & Trim(.txtConSalesGrpCd.value) & _
									"&BizAreaCd=" & Trim(.txtConBizAreaCd.value)
									
			If .rdoExceptFlagA.checked Then
				iStrVal = iStrVal & "&ExceptFlag="
			ElseIf .rdoExceptFlagN.checked Then
				iStrVal = iStrVal & "&ExceptFlag=N"
			Else
				iStrVal = iStrVal & "&ExceptFlag=Y"
			End If
					
			If .rdoPostFlagA.checked Then
				iStrVal = iStrVal & "&PostFlag="
			ElseIf .rdoPostFlagN.checked Then
				iStrVal = iStrVal & "&PostFlag=N"
			Else
				iStrVal = iStrVal & "&PostFlag=Y"
			End If

			iStrVal = iStrVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")			 
			iStrVal = iStrVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
			iStrVal = iStrVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))				

			lgStartRow = .vspdData.MaxRows + 1										'포멧팅 적용하는 시작Row
		End With	
    End If    

    Call RunMyBizASP(MyBizASP, iStrVal)										'☜: 비지니스 ASP 를 가동 

    If Err.number = 0 Then
       DbQuery = True																'⊙: Processing is OK
    End If   

    Set gActiveElement = document.ActiveElement   

End Function

'========================================
Function DbQueryOk()												

    On Error Resume Next
    Err.Clear

    lgIntFlgMode = parent.OPMD_UMODE
    
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
		'영업그룹 
		Case C_PopSalesGrp	
			iArrParam(1) = "B_SALES_GRP"								' TABLE 명칭 
			iArrParam(2) = Trim(frm1.txtConSalesGrpCd.value)			' Code Condition
			iArrParam(3) = ""											' Name Cindition
			iArrParam(4) = "USAGE_FLAG = " & FilterVar("Y", "''", "S") & " "							' Where Condition
			iArrParam(5) = frm1.txtConSalesGrpCd.alt					' TextBox 명칭 
			
			iArrField(0) = "ED15" & Parent.gColSep & "SALES_GRP"		' Field명(0)
			iArrField(1) = "ED30" & Parent.gColSep & "SALES_GRP_NM"		' Field명(1)
    
		    iArrHeader(0) = frm1.txtConSalesGrpCd.alt					' Header명(0)
		    iArrHeader(1) = frm1.txtConSalesGrpNm.alt					' Header명(1)

			frm1.txtConSalesGrpCd.focus 
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
			Case C_PopSalesGrp
				.txtConSalesGrpCd.value = pvArrRet(0)
				.txtConSalesGrpNm.value = pvArrRet(1)

			Case C_PopBizArea
				.txtConBizAreaCd.value = pvArrRet(0) 
				.txtConBizAreaNm.value = pvArrRet(1)   			

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
    
    If OldLeft <> NewLeft Then Exit Sub

	If CheckRunningBizProcess Then  Exit Sub
    
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>매출채권집계조회(영업그룹)</font></td>
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
											<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name=txtConFromDt CLASS=FPDTYYYYMMDD tag="12X1" ALT="시작년도" Title=FPDATETIME id=OBJECT1></OBJECT>');</SCRIPT>
											</TD>
											<TD>
											&nbsp;~&nbsp;
											</TD>
											<TD>
											<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name=txtConToDt CLASS=FPDTYYYYMMDD tag="12X1" ALT="종료년도" Title=FPDATETIME id=OBJECT2></OBJECT>');</SCRIPT>
											</TD>
										</TR>
									</TABLE>							        
									<TD CLASS="TD5" NOWRAP>영업그룹</TD>
									<TD CLASS="TD6" NOWRAP>	<INPUT TYPE=TEXT NAME="txtConSalesGrpCd" SIZE=10 MAXLENGTH=5 tag="11NXXU" ALT="영업그룹"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnConSalesGrp" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenConPopUp(C_PopSalesGrp) ">
															<INPUT TYPE=TEXT NAME="txtConSalesGrpNm" SIZE=20 tag="14" ALT="영업그룹명"></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>사업장</TD>
									<TD CLASS="TD6" NOWRAP>	<INPUT TYPE=TEXT NAME="txtConBizAreaCd" SIZE=10 MAXLENGTH=10 tag="11NXXU" ALT="사업장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnConBizArea" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenConPopUp(C_PopBizArea) ">
															<INPUT TYPE=TEXT NAME="txtConBizAreaNm" SIZE=20 tag="14" ALT="사업장명"></TD>
									<TD CLASS="TD5" NOWRAP>&nbsp;</TD>
									<TD CLASS="TD6" NOWRAP>&nbsp;</TD>
								</TR>								
								<TR>
									<TD CLASS=TD5 NOWRAP>예외여부</TD>
									<TD CLASS=TD6 NOWRAP>
										<input type=radio CLASS="RADIO" name="rdoExceptFlag" id="rdoExceptFlagA" value="A" tag = "11X" checked>
											<label for="rdoExceptFlagA">전체</label>&nbsp;&nbsp;
										<input type=radio CLASS="RADIO" name="rdoExceptFlag" id="rdoExceptFlagY" value="Y" tag = "11X">
											<label for="rdoExceptFlagY">예외</label>&nbsp;&nbsp;
										<input type=radio CLASS = "RADIO" name="rdoExceptFlag" id="rdoExceptFlagN" value="N" tag = "11X">
											<label for="rdoExceptFlagN">정상</label>
									</TD>
									<TD CLASS=TD5 NOWRAP>확정여부</TD>
									<TD CLASS=TD6 NOWRAP>
										<input type=radio CLASS="RADIO" name="rdoPostFlag" id="rdoPostFlagA" value="A" tag = "11X" checked>
											<label for="rdoPostFlagA">전체</label>&nbsp;&nbsp;
										<input type=radio CLASS="RADIO" name="rdoPostFlag" id="rdoPostFlagY" value="Y" tag = "11X">
											<label for="rdoPostFlagY">확정</label>&nbsp;&nbsp;
										<input type=radio CLASS = "RADIO" name="rdoPostFlag" id="rdoPostFlagN" value="N" tag = "11X">
											<label for="rdoPostFlagN">미확정</label>
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
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=vspdData> <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
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

</FORM>
    <DIV ID="MousePT" NAME="MousePT">
    <iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
    </DIV>
</BODY>
</HTML>
