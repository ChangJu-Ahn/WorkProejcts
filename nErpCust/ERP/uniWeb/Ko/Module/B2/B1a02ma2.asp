
<%@ LANGUAGE="VBSCRIPT" %>
<%'======================================================================================================
'*  1. Module Name          : Basis Architect
'*  2. Function Name        : Master Data(Display Minor Code)
'*  3. Program ID           : b1a02ma2.asp
'*  4. Program Name         : b1a02ma2.asp
'*  5. Program Desc         :
'*  6. Modified date(First) : 1999/09/10
'*  7. Modified date(Last)  : 2002/08/06
'*  8. Modifier (First)     : Hwang Jeong-won
'*  9. Modifier (Last)      : Sim Hae Young
'* 11. Comment              :
'**********************************************************************************************
%>
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="vbscript" SRC="../../inc/Cookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit																	'☜: indicates that All variables must be declared in advance

Const BIZ_PGM_ID = "B1a02mb1.asp"												'☆: 비지니스 로직 ASP명 
Const BIZ_PGM_JUMP_ID = "B1a01ma2"												'☆: Jump시 호출 ASP명 

Const CookieSplit = 1233

<!-- #Include file="../../inc/lgvariables.inc" -->

Dim C_Minor        
Dim C_MinorNm      
Dim C_MinorType    

Dim lgStrQueryFlag			  ' "N":Next, "P":Prev, "Q":Query
Dim IsOpenPop          

Sub InitSpreadPosVariables()
    C_Minor        = 1
    C_MinorNm      = 2
    C_MinorType    = 3
End Sub

Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
    
    lgStrPrevKey = ""
   
    lgLngCurRows = 0                            'initializes Deleted Rows Count
    lgStrQueryFlag = "Q"
    
End Sub

Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029A("I", "B","NOCOOKIE","MA") %>
End Sub

Sub InitSpreadSheet()
    Call initSpreadPosVariables()  

	With frm1.vspdData
	
	ggoSpread.Source = frm1.vspdData	
   'patch version
    ggoSpread.Spreadinit "V20021202",,parent.gAllowDragDropSpread    
    
	.ReDraw = false

	.MaxCols = C_MinorType + 1												'☜: 최대 Columns의 항상 1개 증가시킴 
	.Col = .MaxCols														'☆: 사용자 별 Hidden Column
	.ColHidden = True    
	       
    .MaxRows = 0
    ggoSpread.ClearSpreadData
	
    Call GetSpreadColumnPos("A")  

	ggoSpread.SSSetEdit C_Minor, "Minor코드", 26	, , ,10, 2			'1
	ggoSpread.SSSetEdit C_MinorNm, "Minor코드명", 50	, , ,50			'2
	ggoSpread.SSSetCombo C_MinorType, "Minor코드 정의형식", 40

	.ReDraw = true
	
    Call SetSpreadLock 
    
    End With
    
End Sub

Sub SetSpreadLock()
    With frm1
        .vspdData.ReDraw = False
        ggoSpread.SpreadLock C_Minor,     -1, C_Minor	
        ggoSpread.SpreadLock C_MinorNm,   -1, C_MinorNm	
        ggoSpread.SpreadLock C_MinorType, -1, C_MinorType	
		ggoSpread.SSSetProtected .vspdData.MaxCols, -1, -1
        .vspdData.ReDraw = True
    End With
End Sub

Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
    Dim iRow

    With frm1
        .vspdData.ReDraw = False
        ggoSpread.SSSetRequired C_Minor, pvStartRow, pvEndRow
        ggoSpread.SSSetRequired C_MinorNm, pvStartRow, pvEndRow

        If .rdoChargeCd1.Checked = True Then
	        ggoSpread.SSSetRequired C_MinorType, pvStartRow, pvEndRow
	    Else
	        For iRow  = pvStartRow to pvEndRow
	    	    ggoSpread.SSSetProtected C_MinorType, pvStartRow, pvEndRow
	    	    .vspdData.Col  = C_MinorType
	    	    .vspdData.Row  = lRow
	    	    .vspdData.Text = "시스템 정의"
	    	Next
	    End If

        .vspdData.ReDraw = True
    End With
End Sub

Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
            C_Minor        = iCurColumnPos(1)
            C_MinorNm      = iCurColumnPos(2)
            C_MinorType    = iCurColumnPos(3)
    End Select    
End Sub

Sub InitSpreadComboBox()
    ggoSpread.SetCombo "시스템 정의" & vbtab & "사용자 정의", C_MinorType
End Sub

Function OpenMajor()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "Major코드 팝업"			<%' 팝업 명칭 %>
	arrParam(1) = "B_MAJOR"				 		<%' TABLE 명칭 %>
	arrParam(2) = frm1.txtMajor.value			<%' Code Condition%>
	arrParam(3) = ""							<%' Name Cindition%>
	arrParam(4) = ""							<%' Where Condition%>
	arrParam(5) = "Major코드"			
	
    arrField(0) = "major_cd"					<%' Field명(0)%>
    arrField(1) = "major_nm"				<%' Field명(1)%>
    
    arrHeader(0) = "Major코드"						<%' Header명(0)%>
    arrHeader(1) = "Major코드명"						<%' Header명(1)%>
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetMajor(arrRet)
	End If	

End Function

Function SetMajor(Byval arrRet)
	With frm1
		.txtMajor.value = arrRet(0)
		.txtMajorNm.value = arrRet(1)		
	End With
End Function

Function CookiePage(ByVal flgs)

	On Error Resume Next

	Const CookieSplit = 1233						<%'Cookie Split String : CookiePage Function Use%>

	Dim strTemp, arrVal
	
	If flgs = 1 Then
	
		WriteCookie CookieSplit , frm1.txtMajor.value 
		
	ElseIf flgs = 0 Then
	
		strTemp = ReadCookie(CookieSplit)

		If strTemp = "" then Exit Function
		
		arrVal = Split(strTemp, parent.gRowSep)	

		If arrVal(0) = "" then Exit Function

		frm1.txtMajor.value =  arrVal(0)
		
		If Err.number <> 0 Then
			Err.Clear
			WriteCookie CookieSplit , ""
			Exit function
		End If
		
		WriteCookie CookieSplit , ""
		
		FncQuery()
			
	End If
			
End Function

Sub Form_Load()

    Call LoadInfTB19029                                                     <%'Load table , B_numeric_format%>
    
    Call ggoOper.LockField(Document, "N")                                   <%'Lock  Suitable  Field%>                         
                                                                            <%'Format Numeric Contents Field%>                                                                            
        
    Call InitSpreadSheet                                                    <%'Setup the Spread sheet%>
    Call InitVariables                                                      <%'Initializes local global variables%>
    
    Call InitSpreadComboBox
    Call SetToolbar("1100000000001111")									'⊙: 버튼 툴바 제어 
    Call CookiePage(0)
    
    frm1.txtMajor.focus 
    
End Sub

Sub vspdData_Change(ByVal Col, ByVal Row)
	ggoSpread.Source = frm1.vspdData
	ggoSpread.UpdateRow Row
End Sub

Sub vspdData_Click(ByVal Col, ByVal Row)
    Call SetPopupMenuItemInf("0000011111") 

    gMouseClickStatus = "SPC"   

    Set gActiveSpdSheet = frm1.vspdData
   
    If frm1.vspdData.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
   	    
    If Row <= 0 Then
       ggoSpread.Source = frm1.vspdData
       
       If lgSortKey = 1 Then
           ggoSpread.SSSort Col               'Sort in ascending
           lgSortKey = 2
       Else
           ggoSpread.SSSort Col, lgSortKey    'Sort in descending 
           lgSortKey = 1
       End If
       
       Exit Sub
    End If

	frm1.vspdData.Row = Row
End Sub

Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

Sub vspdData_DblClick(ByVal Col, ByVal Row)				
    Dim iColumnName
        
    If Row <= 0 Then
        Exit Sub
    End If
        
    If frm1.vspdData.MaxRows = 0 Then
        Exit Sub
    End If
    	
End Sub

Sub vspdData_GotFocus()
    ggoSpread.Source = Frm1.vspdData
 
End Sub

Sub vspdData_MouseDown(Button , Shift , x , y)
    
    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If
        
End Sub    

Sub vspdData_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    
    Call GetSpreadColumnPos("A")
End Sub

Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
        
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
	If CheckRunningBizProcess = True Then
	   Exit Sub
	End If
        
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	           
    	If lgStrPrevKeyIndex <> "" Then                         
      	   Call DisableToolBar(Parent.TBC_QUERY)
			If DbQuery = False Then
				Call RestoreTooBar()
			    Exit Sub
			End If  				
    	End If
    End if
End Sub

Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

Sub PopRestoreSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()      
    Call InitSpreadComboBox
	Call ggoSpread.ReOrderingSpreadData()
	Call InitData()
End Sub

Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If CheckRunningBizProcess = True Then					'☜: 조회중이면 다음 조회 안하도록 체크 
        Exit Sub
	End If
        
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) And Not(lgStrPrevKey = "") Then
		Call DisableToolBar(parent.TBC_QUERY)					'☜ : Query 버튼을 disable 시킴.
		If DBQuery = False Then 
		   Call RestoreToolBar()
		   Exit Sub 
		End If
    End if
    
End Sub

Function FncQuery() 
    Dim IntRetCD 
    
    lgStrQueryFlag = "Q"
    
    FncQuery = False                                                        '⊙: Processing is NG
    
    Err.Clear                                                               '☜: Protect system from crashing

    '-----------------------
    'Check previous data area
    '-----------------------
    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "X", "X")			    <%'데이타가 변경되었습니다. 조회하시겠습니까?%>
    	If IntRetCD = vbNo Then
      	Exit Function
    	End If
    End If
    
        '-----------------------
    'Erase contents area
    '-----------------------
    frm1.txtMajorNm.value = ""
    Call ggoOper.ClearField(Document, "2")										'⊙: Clear Contents  Field
    ggoSpread.ClearSpreadData
    If lgStrQueryFlag = "Q" Then Call InitVariables							'⊙: Initializes local global variables
    
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then									'⊙: This function check indispensable field
       Exit Function
    End If
    
    '-----------------------
    'Query function call area
    '-----------------------
	If DbQuery = False Then Exit Function										'☜: Query db data
       
    FncQuery = True																'⊙: Processing is OK
    
End Function

Function FncNew() 
    Dim IntRetCD 
    
    FncNew = False                                                          '⊙: Processing is NG
    
    Err.Clear                                                               '☜: Protect system from crashing
    'On Error Resume Next                                                    '☜: Protect system from crashing
    
    
    '-----------------------
    'Check previous data area
    '-----------------------
    If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
        IntRetCD = DisplayMsgBox("900015", parent.VB_YES_NO, "X", "X")    
		If IntRetCD = vbNo Then
			Exit Function
		End If
       
    End If
    
    '-----------------------
    'Erase condition area
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "A")                                     
    ggoSpread.Source = frm1.vspdData
    ggoSpread.ClearSpreadData
    Call ggoOper.LockField(Document, "N")                                          '⊙: Lock  Suitable  Field
    Call InitVariables                                                      '⊙: Initializes local global variables
        
    FncNew = True                                                           '⊙: Processing is OK

End Function

Function FncDelete() 
    Dim IntRetCD 
    
    FncDelete = False                                                       '⊙: Processing is NG
    
    Err.Clear                                                               '☜: Protect system from crashing
    'On Error Resume Next                                                    '☜: Protect system from crashing
    
    
    '-----------------------
    'Precheck area
    '-----------------------
    If lgIntFlgMode <> parent.OPMD_UMODE Then                                      'Check if there is retrived data
        Call DisplayMsgBox("900002", "X", "X", "X")                                  '☆:
        Exit Function
    End If
    
    IntRetCD = DisplayMsgBox("900003", parent.VB_YES_NO, "X", "X")  
    If IntRetCD = vbNo Then
        Exit Function
    End If    
        
    '-----------------------
    'Erase condition area
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "A")                                        
    ggoSpread.Source = frm1.vspdData
    ggoSpread.ClearSpreadData
    
    FncDelete = True                                                        '⊙: Processing is OK
    
End Function

Function FncExit()
	Dim IntRetCD
	FncExit = False
    ggoSpread.Source = frm1.vspdData	
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO, "X", "X")                <%'데이타가 변경되었습니다. 종료 하시겠습니까?%>
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    FncExit = True
End Function

Function FncPrev() 	
	Dim IntRetCD 
	
	lgStrQueryFlag = "P"
	lgStrPrevKey = ""
	lgIntFlgMode = parent.OPMD_CMODE
	
    Err.Clear                                                               '☜: Protect system from crashing

    '-----------------------
    'Check previous data area
    '-----------------------
    ggoSpread.Source = frm1.vspdData	
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO, "X", "X")                <%'데이타가 변경되었습니다. 종료 하시겠습니까?%>
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")										'⊙: Clear Contents  Field
    ggoSpread.ClearSpreadData
    
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then									'⊙: This function check indispensable field
       Exit Function
    End If
    
    '-----------------------
    'Query function call area
    '-----------------------
    If DbQuery = False Then Exit Function								'☜: Query db data
      
End Function

Function FncNext() 	
	Dim IntRetCD 
	
	lgStrQueryFlag = "N"
	lgStrPrevKey = ""
	lgIntFlgMode = parent.OPMD_CMODE
	
    Err.Clear                                                               '☜: Protect system from crashing

    '-----------------------
    'Check previous data area
    '-----------------------
    ggoSpread.Source = frm1.vspdData	
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO, "X", "X")                <%'데이타가 변경되었습니다. 종료 하시겠습니까?%>
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")										'⊙: Clear Contents  Field
    ggoSpread.ClearSpreadData
    
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then									'⊙: This function check indispensable field
       Exit Function
    End If
    
    '-----------------------
    'Query function call area
    '-----------------------
    If DbQuery = False Then Exit Function
    
End Function

Function FncPrint() 
    Call parent.FncPrint()                                                   '☜: Protect system from crashing
End Function

Function FncExcel() 
    Call parent.FncExport(parent.C_MULTI)                                                   <%'☜: Protect system from crashing%>
End Function

Function FncFind() 
    Call parent.FncFind(parent.C_MULTI, False)                                         <%'☜:화면 유형, Tab 유무 %>
End Function

Function Clear()
	Call ggoOper.ClearField(Document, "2")
    ggoSpread.Source = frm1.vspdData
    ggoSpread.ClearSpreadData
	
End Function

Function DbQuery() 
    Dim LngLastRow      
    Dim LngMaxRow       
    Dim LngRow          
    Dim strTemp         
    Dim B1A028         'As New P21018ListIndReqSvr

    DbQuery = False
    
    Call LayerShowHide(1)
    
    Err.Clear                                                               '☜: Protect system from crashing

	Dim strVal
    
    With frm1
        
    If lgIntFlgMode = parent.OPMD_UMODE Then
		strVal = BIZ_PGM_ID & "?txtMode=" & lgStrQueryFlag							'☜: 
		strVal = strVal & "&txtMajor=" & .hMajor.value 
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
    Else
		strVal = BIZ_PGM_ID & "?txtMode=" & lgStrQueryFlag							'☜: 
		strVal = strVal & "&txtMajor=" & Trim(.txtMajor.value)
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
    End If
   
	Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 
        
    End With
    
    DbQuery = True
    

End Function

Function DbQueryOk()														'☆: 조회 성공후 실행로직 
	
    '-----------------------
    'Reset variables area
    '-----------------------
    lgIntFlgMode = parent.OPMD_UMODE												'⊙: Indicates that current mode is Update mode    
    ggoSpread.SpreadLock 3, 1, 3
    Call SetToolbar("1100000011011111")										'⊙: 버튼 툴바 제어 
        
End Function

Function DbPrevNextOk()														'☆: 조회 성공후 실행로직 
	
	lgIntFlgMode = parent.OPMD_UMODE												'⊙: Indicates that current mode is Update mode
	frm1.vspdData.MaxRows = 0
    ggoSpread.Source = frm1.vspdData
    ggoSpread.ClearSpreadData
    Call MainQuery()
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>Minor코드조회</font></td>
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
					<TR>
						<TD CLASS="TD5">Major코드</TD>
						<TD CLASS="TD6">
										<INPUT TYPE=TEXT NAME="txtMajor" SIZE=10 MAXLENGTH=5 tag="12XXXU" ALT="Major코드"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCountryCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenMajor()">
										<INPUT TYPE=TEXT NAME="txtMajorNm" tag="14X"maxlength=30>
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
							<TD CLASS="TD5">Minor코드 길이</TD>
							<TD CLASS="TD6">
							<INPUT TYPE=TEXT NAME="txtMinorLen" SIZE=10 MAXLENGTH=2 tag="14" STYLE="Text-Align:Right" ALT="Minor길이"></TD>
						</TR>
						<TR>
							<TD CLASS="TD5">사용자정의 Minor코드 추가가능여부</TD>
							<TD CLASS="TD6">
								<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoChargeCd" TAG="2X" VALUE="Y" CHECKED ID="rdoChargeCd1" disabled>가능
								<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoChargeCd" TAG="2X" VALUE="N" ID="rdoChargeCd2" disabled>불가능
							</TD>							
						</TR>					
						<TR>
							<TD HEIGHT="100%" WIDTH="100%" COLSPAN=2>
								<script language =javascript src='./js/b1a02ma2_I182269680_vspdData.js'></script>
							</TD>
						</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<tr>
      <td <%=HEIGHT_TYPE_01%>></td>
    </tr>
    <tr HEIGHT="20">
      <td WIDTH="100%">
      <table <%=LR_SPACE_TYPE_30%>>
        <tr>
          <td WIDTH="*" ALIGN=RIGHT><a href = "VBSCRIPT:PgmJump(BIZ_PGM_JUMP_ID)" ONCLICK="VBSCRIPT:CookiePage 1">Major코드조회</a></td>
		  <TD WIDTH=10>&nbsp;</TD>
        </tr>
      </table>
      </td>
    </tr>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="B1a02mb1.asp" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="hMajor" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

