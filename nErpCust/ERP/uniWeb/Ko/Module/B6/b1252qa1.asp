<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Prucurement
'*  2. Function Name        : 
'*  3. Program ID           : b1252qa1
'*  4. Program Name         : 구매조직조회 
'*  5. Program Desc         : 구매조직조회 
'*  6. Component List       : 
'*  7. Modified date(First) : 2000/06/08
'*  8. Modified date(Last)  : 2003/05/23
'*  9. Modifier (First)     : Shin Jin Hyun
'* 10. Modifier (Last)      : Kang Su Hwan
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change" 
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************
-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!--'******************************************  1.1 Inc 선언   **********************************************-->
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->
<!--'==========================================  1.1.1 Style Sheet  ======================================-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'==========================================  1.1.2 공통 Include   ======================================-->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentA.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="vbscript"   SRC="../../inc/Cookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit																	'☜: indicates that All variables must be declared in advance
<!-- #Include file="../../inc/lgvariables.inc" -->

'==========================================  1.2.1 Global 상수 선언  ======================================
Const BIZ_PGM_ID = "b1252qb1.asp"												'☆: 비지니스 로직 ASP명 
Const BIZ_PGM_JUMP_ID = "b1252ma1"

Dim C_ORGCd
Dim C_ORGNm
Dim C_Useflg
Dim C_FrExpiryDt
Dim C_ToExpiryDt

Const C_SHEETMAXROWS=100

'==========================================  1.2.3 Global Variable값 정의  ===============================
dim IsOpenPop          

'========================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)
   If Button = 2 And gMouseClickStatus = "SPC" Then
      gMouseClickStatus = "SPCR"
   End If
End Sub    

'========================================================================================
Function FncSplitColumn()
    
   If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Function
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   
    
End Function

Function CookiePage(Byval Kubun)
	Dim strTemp, arrVal
	Dim IntRetCD

	if Kubun = 1 then
		if frm1.vspdData.ActiveRow > 0 then
			frm1.vspdData.Row = frm1.vspdData.ActiveRow 
			frm1.vspdData.Col = C_ORGCd
			WriteCookie "ORGCd" , frm1.vspdData.Text
		End if
		
		Call PgmJump(BIZ_PGM_JUMP_ID)
	Else
	    If ReadCookie ("Kubun") = "Y" then 
			frm1.txtOrgCd.value	= ReadCookie ("ORGCd")
	    	
	    	WriteCookie "Kubun", ""
	    	WriteCookie "ORGCd", ""
	    	
	    	Call MainQuery()	    	
	    End if
	End if 
	
End Function

'==========================================  2.1.1 InitVariables()  ======================================
Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
    
    lgStrPrevKey = ""                           'initializes Previous Key
    lgLngCurRows = 0                            'initializes Deleted Rows Count
    lgSortKey = 1
End Sub

'==========================================  2.2.1 SetDefaultVal()  ========================================
 Sub SetDefaultVal()
End Sub

'========================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A( "Q", "*", "NOCOOKIE", "QA") %>
End Sub

'========================================================================================================
Sub InitSpreadPosVariables()
	C_ORGCd = 1
	C_ORGNm = 2	
	C_Useflg = 3
	C_FrExpiryDt=4
	C_ToExpiryDt=5
End Sub

'=============================================== 2.2.3 InitSpreadSheet() ========================================
 Sub InitSpreadSheet()
	
	Call InitSpreadPosVariables()
	
	With frm1.vspdData
		ggoSpread.Source = frm1.vspdData
        ggoSpread.Spreadinit "V20021103",, parent.gAllowDragDropSpread
		
		.ReDraw = false
	
		.MaxCols = C_ToExpiryDt + 1												'☜: 최대 Columns의 항상 1개 증가시킴 
		.Col = C_ToExpiryDt + 1														'☆: 사용자 별 Hidden Column
    
		.MaxRows = 0
		
		Call GetSpreadColumnPos("A")
		frm1.vspdData.OperationMode = 0

		ggoSpread.SSSetEdit C_ORGCd, "구매조직", 20
		ggoSpread.SSSetEdit C_ORGNm, "구매조직명", 30
		ggoSpread.SSSetEdit C_Useflg, "사용여부", 20
		ggoSpread.SSSetDate C_FrExpiryDt, "유효시작일", 25, 2, gDateFormat
		ggoSpread.SSSetDate C_ToExpiryDt, "유효종료일", 25, 2, gDateFormat
		
		Call ggoSpread.SSSetColHidden(.MaxCols,	.MaxCols,	True)
			
		Call SetSpreadLock 
    
		.ReDraw = true
    End With
End Sub

'========================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			 C_ORGCd	  = iCurColumnPos(1)
			 C_ORGNm	  = iCurColumnPos(2)
			 C_Useflg	  = iCurColumnPos(3)
			 C_FrExpiryDt = iCurColumnPos(4)
			 C_ToExpiryDt = iCurColumnPos(5)
	 End Select    
End Sub

'================================== 2.2.4 SetSpreadLock() ==================================================
Sub SetSpreadLock()
	ggoSpread.Source = frm1.vspdData
	ggoSpread.SpreadLockWithOddEvenRowColor()
End Sub

'================================== 2.2.5 SetSpreadColor() ==================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
    With frm1
    
    .vspdData.ReDraw = False
    ggoSpread.SSSetRequired		C_ItemCode, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected	C_ItemNm, pvStartRow, pvEndRow
    ggoSpread.SSSetRequired		C_ReqrdQty, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected	C_MpsRefltQty, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected	C_BasicUnit, pvStartRow, pvEndRow
    ggoSpread.SSSetRequired		C_ReqrdDt, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected	C_AvailDt, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected	C_ReqStatus, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected	C_Column10, pvStartRow, pvEndRow
    .vspdData.ReDraw = True
    
    End With
End Sub

'--------------------------------------------------------------------------------------------------------- 
Function OpenORG()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "구매조직"						<%' 팝업 명칭 %>
	arrParam(1) = "b_pur_org"						<%' TABLE 명칭 %>
	
	arrParam(2) = UCase(Trim(frm1.txtORGCd.Value))	<%' Code Condition%>
'	arrParam(3) = Trim(frm1.txtORGNm.Value)	<%' Name Cindition%>
	
	arrParam(4) = ""							<%' Where Condition%>
	arrParam(5) = "구매조직"							<%' TextBox 명칭 %>
	
    arrField(0) = "PUR_ORG"					<%' Field명(0)%>
    arrField(1) = "PUR_ORG_NM"					<%' Field명(1)%>
    
    arrHeader(0) = "구매조직"						<%' Header명(0)%>
    arrHeader(1) = "구매조직명"						<%' Header명(1)%>
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtORGCd.focus
		Exit Function
	Else
		frm1.txtORGCd.value = arrRet(0)
		frm1.txtORGNm.value = arrRet(1)
		frm1.txtORGCd.focus
	End If	
End Function

'==========================================  3.1.1 Form_Load()  ======================================
Sub Form_Load()
    Call LoadInfTB19029
    Call ggoOper.LockField(Document, "N")
    
    Call InitVariables
    Call InitSpreadSheet
'    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,gDateFormat,parent.gComNum1000,parent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
    Call SetToolbar("1100000000001111")										'⊙: 버튼 툴바 제어	
    frm1.txtORGCd.focus
    Set gActiveElement = document.activeElement 
    Call CookiePage(0)
    
End Sub

'********************************************************************************************************* %>
Sub vspdData_Click(ByVal Col, ByVal Row)
	Set gActiveSpdSheet = frm1.vspdData
	frm1.vspdData.OperationMode = 0    
	
	gMouseClickStatus = "SPC"  
	Call SetPopupMenuItemInf("0000111111")
	
	If frm1.vspdData.MaxRows <= 0 Then                                                    'If there is no data.
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
	
	frm1.vspdData.Row = Row
End Sub

'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)
End Sub

Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	Dim strTemp
	Dim intPos1
   
	With frm1.vspdData 
	 
		ggoSpread.Source = frm1.vspdData
		  
		If Row > 0 And Col = C_OrderUnitPopUp Then
		  
			.Col = Col
			.Row = Row
			Call OpenUnit(.text)
		Elseif Row > 0 And Col = C_CurrPopup Then
		  
			.Col = Col
			.Row = Row
			Call OpenCurr(.text)
		End if 
	End With
End Sub

'==========================================================================================
Sub vspdData_GotFocus()
    ggoSpread.Source = frm1.vspdData
End Sub

'========================================================================================================
Sub vspdData_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)

    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    
    Call GetSpreadColumnPos("A")
End Sub

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
    Call InitComboBox
	Call ggoSpread.ReOrderingSpreadData()
	Call InitData()
End Sub

'==========================================================================================
 Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )

    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    <% '----------  Coding part  -------------------------------------------------------------%>   
    if frm1.vspdData.MaxRows < NewTop + C_SHEETMAXROWS Then	'☜: 재쿼리 체크 
		If lgStrPrevKey <> "" Then							'⊙: 다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
			If CheckRunningBizProcess = True Then
				Exit Sub
			End If	
			
			Call DisableToolBar(TBC_QUERY)
			If DBQuery = False Then
				Call RestoreToolBar()
				Exit Sub
			End If
		End If
    End if
    
End Sub

'==========================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)				
    Dim iColumnName
    
	If Row <= 0 Then
		Exit Sub
	End If
	
	If frm1.vspdData.MaxRows = 0 Then
		Exit Sub
	End if
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub

'========================================================================================
Function FncQuery() 
    Dim IntRetCD 
    
    FncQuery = False                                                        '⊙: Processing is NG
    
    Err.Clear                                                               '☜: Protect system from crashing

	ggoSpread.Source = frm1.vspdData
	
    '-----------------------
    'Erase contents area
    '-----------------------
'    Call ggoOper.ClearField(Document, "2")										'⊙: Clear Contents  Field
	frm1.vspdData.MaxRows = 0
    Call InitVariables
    															'⊙: Initializes local global variables
    
    '-----------------------
    'Check condition area
    '-----------------------
'    If Not ChkField(Document, "1") Then									'⊙: This function check indispensable field
'       Exit Function
'    End If
    
    '-----------------------
    'Query function call area
    '-----------------------
    If DbQuery = False Then Exit Function
       
    FncQuery = True																'⊙: Processing is OK
    
End Function

'========================================================================================
Function FncNew() 
    Dim IntRetCD 
    
    FncNew = False                                                          '⊙: Processing is NG
    
    Err.Clear                                                               '☜: Protect system from crashing
    
    ggoSpread.Source = frm1.vspdData
    
    '-----------------------
    'Erase condition area
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "A")                                         '⊙: Clear Condition Field
    Call ggoOper.LockField(Document, "N")                                          '⊙: Lock  Suitable  Field
    Call InitVariables                                                      '⊙: Initializes local global variables
    Call SetDefaultVal
    Call SetToolbar("11100000000000")    
        
    FncNew = True                                                           '⊙: Processing is OK

End Function

'========================================================================================
Function FncPrint()
	ggoSpread.Source = frm1.vspdData 
    Call parent.FncPrint()
End Function

'========================================================================================
Function FncExcel()
	ggoSpread.Source = frm1.vspdData 
     Call parent.FncExport(Parent.C_Multi)	
End Function

'========================================================================================
Function FncFind() 
	ggoSpread.Source = frm1.vspdData
    Call parent.FncFind(Parent.C_Multi , False)                                     <%'☜:화면 유형, Tab 유무 %>
End Function

'========================================================================================
Function FncExit()
    FncExit = True
End Function

'========================================================================================
 Function DbQuery() 
    Dim LngLastRow      
    Dim LngMaxRow       
    Dim LngRow          
    Dim strTemp         
    Dim StrNextKey      
    Dim pP21018         'As New P21018ListIndReqSvr

    DbQuery = False
    
    if LayerShowHide(1) = False then
       Exit Function 
    end if

    Err.Clear                                                               '☜: Protect system from crashing

	Dim strVal
    
    With frm1
    
    strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001
    strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
    strVal = strVal & "&txtORGCd=" & Trim(.txtORGCd.value)
    If .rdoUseflg(0).checked = True  Then
		strVal = strVal & "&rdoUseflg=" & .rdoUseflg(0).value
	Elseif .rdoUseflg(1).checked = True  Then
		strVal = strVal & "&rdoUseflg=" & .rdoUseflg(1).value
	Elseif .rdoUseflg(2).checked = True  Then
		strVal = strVal & "&rdoUseflg=" & .rdoUseflg(2).value
	End if
	 
	Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 
    
    End With
    
    DbQuery = True
End Function

'========================================================================================
Function DbQueryOk()														'☆: 조회 성공후 실행로직 
    '-----------------------
    'Reset variables area
    '-----------------------
    lgIntFlgMode = parent.OPMD_UMODE												'⊙: Indicates that current mode is Update mode

    Call ggoOper.LockField(Document, "Q")									'⊙: This function lock the suitable field

    Call SetToolbar("1100000000011111")										'⊙: 버튼 툴바 제어	
    If frm1.vspdData.MaxRows > 0 Then
		frm1.vspdData.focus
	Else
		frm1.txtORGCd.focus
	End If
	Set gActiveElement = document.activeElement
End Function

</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc"  -->	
</HEAD>
<!--
'#########################################################################################################
'       					6. Tag부 
'######################################################################################################### 
-->
<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR>
		<TD <%=HEIGHT_TYPE_00%>>&nbsp;<% ' 상위 여백 %></TD>
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white><%=Request("strASPMnuMnuNm")%></font></td>
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
								<tr>
									<TD CLASS="TD5">구매조직</TD>
									<TD CLASS="TD6"><INPUT TYPE=TEXT NAME="txtORGCd" ALT="구매조직" SIZE=10 MAXLENGTH=4  tag="1XNXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnORGCd1" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenORG()">
													<INPUT TYPE=TEXT ID="txtORGNm" ALT="구매조직" NAME="arrCond" tag="24X"></TD>
									<TD CLASS="TD5">사용여부</TD>
									<TD CLASS="TD6"><INPUT TYPE=radio CLASS="Radio" NAME="rdoUseflg" ALT="사용여부" id="rdoUseflg1" checked tag="1X" Value="A">
													<label for="rdoUseflg1">전체</label>
													<INPUT TYPE=radio CLASS="Radio" NAME="rdoUseflg" ALT="사용여부" id="rdoUseflg2" tag="1X" Value="Y">
													<label for="rdoUseflg2">예</label>
													<INPUT TYPE=radio CLASS="Radio" NAME="rdoUseflg" ALT="사용여부" id="rdoUseflg3" tag="1X" Value="N">
													<label for="rdoUseflg3">아니오</label></TD>
								</tr>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD WIDTH=100% valign=top>
						<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD HEIGHT="100%">
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=vaSpread1> <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	
    <tr>
		<TD <%=HEIGHT_TYPE_01%>></TD>
    </tr>
    <tr HEIGHT="20">
		<td WIDTH="100%">
			<table <%=LR_SPACE_TYPE_30%>>
				<tr>
					<TD WIDTH=10>&nbsp;</TD>
					<td WIDTH="*" align="right"><a href="VBSCRIPT:CookiePage(1)">구매조직등록</a></td>
					<TD WIDTH=10>&nbsp;</TD>
				</tr>
			</table>
		</td>
    </tr>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
</FORM>

<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>

</BODY>
</HTML>
