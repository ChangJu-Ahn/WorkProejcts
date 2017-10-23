<%@ LANGUAGE="VBSCRIPT" %>
<!--'======================================================================================================
'*  1. Module Name          : COSTING
'*  2. Function Name        : 실제원가관리 
'*  3. Program ID           : C3810MA1
'*  4. Program Name         : 배부근거조회(품목별)
'*  5. Program Desc         : Query of Account Code
'*  6. Component List       : ADO
'*  7. Modified date(First) : 2005.10.10
'*  8. Modified date(Last)  : 2005.10.10
'*  9. Modifier (First)     : Shin Hyun Ho
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'======================================================================================================= -->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- #Include file="../../inc/IncSvrCcm.inc"  -->
<!-- #Include file="../../inc/IncSvrHTML.inc"  -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAMain.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAEvent.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliVariables.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAOperation.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliRdsQuery.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliDBAgentA.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliDBAgentVariables.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/Cookie.vbs">   </SCRIPT>
<SCRIPT LANGUAGE="VBScript">

Option Explicit                              '☜: indicates that All variables must be declared in advance

<!-- #Include file="../../inc/lgvariables.inc" -->	

'========================================================================================
Dim lgIsOpenPop
Dim IsOpenPop                                               '☜: Popup status                           
Dim lgMark                                                  '☜: 마크                                  

'--------------- 개발자 coding part(변수선언,Start)-----------------------------------------------------------
Const BIZ_PGM_ID        = "C3820MB1.asp"
Const C_MaxKey          = 2 

'========================================================================================
Sub InitVariables()
    lgBlnFlgChgValue = False                               'Indicates that no value changed
    lgPageNo     = ""                                  'initializes Previous Key
    'lgSortKey        = 1

End Sub

'========================================================================================
Sub SetDefaultVal()
    

'--------------- 개발자 coding part(실행로직,Start)--------------------------------------------------
    Dim EndDate

    EndDate = "<%=GetSvrDate%>"


    
    frm1.txtDt.Text   = EndDate 
    
    Call ggoOper.FormatDate(frm1.txtDt, parent.gDateFormat,2)
    
    frm1.txtDt.focus

End Sub

'========================================================================================
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% Call loadInfTB19029A("Q", "C","NOCOOKIE","QA") %>
End Sub


'========================================================================================
Sub InitSpreadSheet()
    Call SetZAdoSpreadSheet("C3820MA1","S","A","V20051010","",frm1.vspdData, C_MaxKey, "X","X")
    Call SetSpreadLock("A")
End Sub



'========================================================================================
Sub SetSpreadLock(ByVal pOpt)
    If pOpt = "A" Then
        With frm1

        .vspdData.ReDraw = False
        ggoSpread.SpreadLockWithOddEvenRowColor()
        ggoSpread.SSSetProtected	.vspdData.MaxCols,-1,-1
        .vspdData.ReDraw = True

        End With
    End if
End Sub



'========================================================================================
Sub InitComboBox()	
	Err.clear
	
End Sub
 


'========================================================================================
Function OpenPopUp(Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	'frm1.hOrgChangeId.value = parent.gChangeOrgId

	Select Case iWhere
		
		Case 0
			arrParam(0) = frm1.txtFctrCd.Alt
			arrParam(1) = "C_DSTB_FCTR"
			arrParam(2) = frm1.txtFctrCd.value
			arrParam(3) = ""
			arrParam(4) = ""  
			arrParam(5) = frm1.txtFctrNm.Alt
	
			arrField(0) = "DSTB_FCTR_CD"
			arrField(1) = "DSTB_FCTR_NM"

			arrHeader(0) = "배부요소"
			arrHeader(1) = "배부요소명"	
		
		Case 1
			arrParam(0) = frm1.txtCostCd.Alt
			arrParam(1) = "B_COST_CENTER A"
			arrParam(2) = frm1.txtCostCd.value
			arrParam(3) = ""
			arrParam(4) = ""  
			arrParam(5) = frm1.txtCostNm.Alt
	
			arrField(0) = "COST_CD"
			arrField(1) = "COST_NM"

			arrHeader(0) = "C/C"
			arrHeader(1) = "C/C명"
			
			
		Case 2
			arrParam(0) = frm1.txtItemCd.Alt
			arrParam(1) = "B_ITEM A"
			arrParam(2) = frm1.txtItemCd.value
			arrParam(3) = ""
			arrParam(4) = ""  
			arrParam(5) = frm1.txtItemNm.Alt
	
			arrField(0) = "ITEM_CD"
			arrField(1) = "ITEM_NM"

			arrHeader(0) = "품목"
			arrHeader(1) = "품목명"	

		
		Case Else
			Exit Function
	End Select
    
	IsOpenPop = True
	
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Select Case iWhere
		Case 0					
			frm1.txtFctrCd.focus
			frm1.txtFctrCd.value = Trim(arrRet(0))
			frm1.txtFctrNm.value = arrRet(1)
		Case 1					
			frm1.txtCostCd.focus
			frm1.txtCostCd.value = Trim(arrRet(0))
			frm1.txtCostNm.value = arrRet(1)
		Case 2					
			frm1.txtItemCd.focus
			frm1.txtItemCd.value = Trim(arrRet(0))
			frm1.txtItemNm.value = arrRet(1)		
		End Select
	End If	

End Function

'========================================================================================
Function PopZAdoConfigGrid()

	Dim arrRet
	Dim gPos

	Select Case UCase(Trim(gActiveSpdSheet.Name))
	       Case "VSPDDATA"
	            gPos = "A"
	       End Select

	If lgIsOpenPop = True Then Exit Function
	lgIsOpenPop = True

	arrRet = window.showModalDialog("../../ComAsp/ZADOSortPopup.asp",Array(ggoSpread.GetXMLData(gPos),gMethodText),"dialogWidth=" & parent.SORTW_WIDTH & "px; dialogHeight=" & parent.SORTW_HEIGHT & "px;; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If arrRet(0) = "X" Then
	   Exit Function
	ElseIf arrRet(0) = "R" Then
	   Call ggoOper.ClearField(Document, "2")
	   Exit Function
	Else
	   Call ggoSpread.SaveXMLData(gPos,arrRet(0),arrRet(1))
       Call InitVariables
       Call InitSpreadSheet()
   End If
End Function



'========================================================================================
Function CookiePage(ByVal Kubun)

End Function

'========================================================================================
Sub Form_Load()
    Call LoadInfTB19029
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
	Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)

    Call ggoOper.LockField(Document, "N")

	Call InitVariables
	Call SetDefaultVal	
	Call InitSpreadSheet()

	Call SetToolbar("1100000000011111")	

End Sub

'========================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )
End Sub

'========================================================================================
Sub txtCostCd_onChange()
	If Trim(frm1.txtCostCd.value) <> "" Then
		Call CommonQueryRs("COST_NM", "B_COST_CENTER", "COST_CD = " & Filtervar(Trim(frm1.txtCostCd.value), "''", "S"), lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
		frm1.txtCostNm.value = Replace(lgF0, chr(11), "")
	Else
		frm1.txtCostCd.value = ""
		frm1.txtCostNm.value = ""
	End If
	
End Sub

Sub txtFctrCd_onChange()
	If Trim(frm1.txtFctrCd.value) <> "" Then
		Call CommonQueryRs("DSTB_FCTR_NM", "C_DSTB_FCTR", "DSTB_FCTR_CD = " & Filtervar(Trim(frm1.txtFctrCd.value), "''", "S"), lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
		frm1.txtFctrNm.value = Replace(lgF0, chr(11), "")
	Else
		frm1.txtFctrCd.value = ""
		frm1.txtFctrNm.value = ""
	End If
	
End Sub

Sub txtItemCd_onChange()
	If Trim(frm1.txtItemCd.value) <> "" Then
		Call CommonQueryRs("ITEM_NM", "B_ITEM", "ITEM_CD = " & Filtervar(Trim(frm1.txtItemCd.value), "''", "S"), lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
		frm1.txtItemNm.value = Replace(lgF0, chr(11), "")
	Else
		frm1.txtItemCd.value = ""
		frm1.txtItemNm.value = ""
	End If
	
End Sub

'========================================================================================
Sub txtDt_DblClick(Button)
    If Button = 1 Then
       frm1.txtDt.Action = 7
       Call SetFocusToDocument("M")	
       frm1.txtDt.Focus       
    End If
End Sub
'========================================================================================
Sub txtDt_KeyPress(KeyAscii)
	If KeyAscii = 13 Then 
		frm1.txtDt.focus
	   Call FncQuery
	End If   
End Sub

'========================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
    Call SetPopupMenuItemInf("00000000001")
    gMouseClickStatus = "SPC"	'Split 상태코드 
    
    Set gActiveSpdSheet = frm1.vspdData
    
    If Row <= 0 Then
        ggoSpread.Source = frm1.vspdData
        If lgSortKey = 1 Then
            ggoSpread.SSSort
            lgSortKey = 2
        Else
            ggoSpread.SSSort ,lgSortKey
            lgSortKey = 1
        End If    
        Exit Sub
    End If

	If Row < 1 Then Exit Sub

	Call SetSpreadColumnValue("A", frm1.vspdData, Col, Row)    
End Sub


Sub vspdData_DblClick(ByVal Col, ByVal Row)				
    If Row <= 0 Then
    End If
	
End Sub

'========================================================================================
Sub vspdData_MouseDown(Button, Shift, X, Y)
	If Button = 2 And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If
End Sub
'========================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If

	If CheckRunningBizProcess = True Then
	   Exit Sub
	End If
    
	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	    
    	If lgPageNo <> "" Then								
           Call DisableToolBar(Parent.TBC_QUERY)
           If DbQuery = False Then
              Call RestoreToolBar()
              Exit Sub
           End if
    	End If
    End If
    
End Sub
'==========================================================================================
Function FncQuery() 
    FncQuery = False                                                        '⊙: Processing is NG
    
    Err.Clear                                                               '☜: Protect system from crashing

    Call InitVariables 														'⊙: Initializes local global variables
    
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then										'⊙: This function check indispensable field
		Exit Function
    End If

	
	Call ggoOper.ClearField(Document, "2")
    ggoSpread.Source = frm1.vspdData
    ggoSpread.ClearSpreadData	

	Call FncSetToolBar("New")
    Call DbQuery

    FncQuery = True		
End Function


'========================================================================================
Function FncPrint() 
    Call parent.FncPrint()
End Function


'========================================================================================
Function FncExcel() 
	Call parent.FncExport(parent.C_MULTI)
End Function


'========================================================================================
Function FncFind() 
    Call parent.FncFind(parent.C_MULTI , False)
End Function

'========================================================================================
Sub FncSplitColumn()
    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If
    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)
End Sub


'========================================================================================
Function FncExit()
    FncExit = True
End Function
'========================================================================================
Function DbQuery() 
	Dim strVal, strZeroFg

    DbQuery = False

    Err.Clear
	Call LayerShowHide(1)

    With frm1


'--------------- 개발자 coding part(실행로직,Start)----------------------------------------------
		strVal = ""
        
        If lgIntFlgMode  <> parent.OPMD_UMODE Then										' This means that it is first search
			strVal = strVal		& BIZ_PGM_ID & "?txtDt=" & Trim(.txtDt.text)	'조회 조건 데이타 
			strVal = strVal		& "&txtFctrCd=" & Trim(.txtFctrCd.value)		'조회 조건 데이타 
			strVal = strVal		& "&txtCostCd=" & Trim(.txtCostCd.Value)		'조회 조건 데이타 
			strVal = strVal		& "&txtItemCd=" & Trim(.txtItemCd.Value)		'조회 조건 데이타 
			strVal = strVal		& "&txtFctrCd_Alt=" & Trim(.txtFctrCd.Alt)		'조회 조건 데이타 
			strVal = strVal		& "&txtCostCd_Alt=" & Trim(.txtCostCd.Alt)		'조회 조건 데이타 
			strVal = strVal		& "&txtItemCd_Alt=" & Trim(.txtItemCd.Alt)		'조회 조건 데이타   
        Else																			' This means that it is next search
			strVal = strVal		& BIZ_PGM_ID & "?txtDt=" & Trim(.htxtDt.text)	'조회 조건 데이타 
			strVal = strVal		& "&txtFctrCd=" & Trim(.htxtFctrCd.value)		'조회 조건 데이타 
			strVal = strVal		& "&txtCostCd=" & Trim(.htxtCostCd.Value)		'조회 조건 데이타 
			strVal = strVal		& "&txtItemCd=" & Trim(.htxtItemCd.Value)		'조회 조건 데이타 
			strVal = strVal		& "&txtFctrCd_Alt=" & Trim(.htxtFctrCd.Alt)		'조회 조건 데이타 
			strVal = strVal		& "&txtCostCd_Alt=" & Trim(.htxtCostCd.Alt)		'조회 조건 데이타 
			strVal = strVal		& "&txtItemCd_Alt=" & Trim(.htxtItemCd.Alt)		'조회 조건 데이타      
			
        End If  


'--------------- 개발자 coding part(실행로직,End)------------------------------------------------

		strVal = strVal & "&lgPageNo="   & lgPageNo                      '☜: Next key tag
		strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")
        strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
		strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))
		'msgbox strval
        Call RunMyBizASP(MyBizASP, strVal)

    End With

    DbQuery = True

End Function



'========================================================================================
Function DbQueryOk()
'    Call ggoOper.LockField(Document, "Q")


	Call FncSetToolBar("Query")
		
	frm1.vspdData.focus
	Set gActiveElement = document.ActiveElement
End Function

'==========================================================
'툴바버튼 세팅 
'==========================================================
Function FncSetToolBar(Cond)
	Select Case UCase(Cond)
	Case "NEW"
		Call SetToolbar("1100000000001111")
	Case "QUERY"
		Call SetToolbar("1100000000011111")
	End Select
End Function


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
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>배부근거조회(품목별)</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD>&nbsp;</TD>					
					<TD>&nbsp;</TD>					
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=*>
		<TD WIDTH=100% CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD HEIGHT=20 WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
								<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS="TD5" NOWRAP>작업년월</TD>
									<TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME="txtDt" CLASS=FPDTYYYYMM tag="12" Title="FPDATETIME" ALT="작업년월" id=fpDateTime1></OBJECT>');</SCRIPT></TD>
									<TD CLASS="TD5" NOWRAP>배부요소</TD>
									<TD CLASS="TD6" NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtFctrCd" SIZE=13 MAXLENGTH=10 STYLE="TEXT-ALIGN: Left" tag="11XXXU" ALT="배부요소"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnFctrCd" align=top TYPE="BUTTON"ONCLICK="vbscript:Call OpenPopup(0)">
									                       <INPUT TYPE=TEXT NAME="txtFctrNm" ALT="배부요소명" SIZE=25 tag="14"></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>Cost Center</TD>
									<TD CLASS="TD6" NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtCostCd" SIZE=10 MAXLENGTH=10 STYLE="TEXT-ALIGN: Left" tag="11XXXU" ALT="Cost Center"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCostCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPopup(1)">
									                       <INPUT TYPE=TEXT NAME="txtCostNm" ALT="Cost Center Name" SIZE=30 tag="14"></TD>
									<TD CLASS="TD5" NOWRAP>품목</TD>
									<TD CLASS="TD6" NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtItemCd" SIZE=10 MAXLENGTH=18 STYLE="TEXT-ALIGN: Left" tag="11XXXU" ALT="품목"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPopup(2)">
									                       <INPUT TYPE=TEXT NAME="txtItemNm" ALT="품목명" SIZE=30 tag="14"></TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=* valign=top>
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD HEIGHT="100%" colspan=7>
								<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> ID=vspdData NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" > <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT></TD>
							</TR>
						</TABLE>						
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD <%=HGIEHT_TYPE_01%>></td>
	</TR>
	<tr>	
		<TD WIDTH="100%" HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" WIDTH="100%" HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" TABINDEX="-1">
</TEXTAREA><%' 업무처리ASP로 넘기기 위한 변수를 담고 있는 Tag들 %>
<INPUT TYPE=hidden NAME="htxtDt"  tag="24">
<INPUT TYPE=hidden NAME="htxtFctrCd"  tag="24">
<INPUT TYPE=hidden NAME="htxtCostCd"  tag="24">
<INPUT TYPE=hidden NAME="htxtItemCd"  tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm" TABINDEX="-1"></iframe>
</DIV>
</BODY>
</HTML>
 

