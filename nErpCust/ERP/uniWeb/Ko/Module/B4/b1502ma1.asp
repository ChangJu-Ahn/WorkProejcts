
<%@ LANGUAGE="VBSCRIPT" %>
<%'======================================================================================================
'*  1. Module Name          : Basis Architect
'*  2. Function Name        : Master Data(Common Holiday)
'*  3. Program ID           : B1502ma1.asp
'*  4. Program Name         : B1502ma1.asp
'*  5. Program Desc         : 공통휴일등록 
'*  6. Modified date(First) : 2000/09/14
'*  7. Modified date(Last)  : 2002/12/13
'*  8. Modifier (First)     : Hwang Jeong-won
'*  9. Modifier (Last)      : Sim Hae Young
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================
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
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit
Const BIZ_PGM_ID = "B1502mb1.asp"												'☆: 비지니스 로직 ASP명 

Dim C_Month
Dim C_Day
Dim C_Remark

Const BIZ_PGM_JUMP_ID = "B1503ma1"

<!-- #Include file="../../inc/lgvariables.inc" -->

Dim IsOpenPop          

Sub InitSpreadPosVariables()
    C_Month     = 1
    C_Day       = 2
    C_Remark    = 3
End Sub

Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    
    lgLngCurRows = 0                            'initializes Deleted Rows Count
    
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

    .MaxCols = C_Remark + 1							'☜: 최대 Columns의 항상 1개 증가시킴 
	.Col = .MaxCols														'☆: 사용자 별 Hidden Column
	.ColHidden = True    
	       
    .MaxRows = 0
    ggoSpread.ClearSpreadData
	
    Call GetSpreadColumnPos("A")  

    ggoSpread.SSSetFloat C_Month,"월" ,12,"6",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,,"1","12"
    ggoSpread.SSSetFloat C_Day,"일" ,12,"6",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,,"1","31"
    ggoSpread.SSSetEdit C_Remark, "사유", 92,,,30
	
	.ReDraw = true
	
    Call SetSpreadLock 
    
    End With
    
End Sub

Sub SetSpreadLock()
    With frm1

    .vspdData.ReDraw = False
    ggoSpread.SpreadLock C_Month, -1, C_Month
    ggoSpread.SpreadLock C_Day, -1, C_Day
	ggoSpread.SSSetProtected .vspdData.MaxCols, -1, -1
    .vspdData.ReDraw = True

    End With
End Sub

Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
    With frm1
    
    .vspdData.ReDraw = False
    ggoSpread.SSSetRequired		C_Month, pvStartRow, pvEndRow
    ggoSpread.SSSetRequired		C_Day, pvStartRow, pvEndRow
    .vspdData.ReDraw = True
    
    End With
End Sub

Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
            C_Month     = iCurColumnPos(1)
            C_Day       = iCurColumnPos(2)
            C_Remark    = iCurColumnPos(3)
    End Select    
End Sub

Sub Form_Load()

    Call LoadInfTB19029                                                     '⊙: Load table , B_numeric_format
    Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field
                                                                                <%'Format Numeric Contents Field%>                                                                            
       
    Call InitSpreadSheet                                                    '⊙: Setup the Spread sheet
    Call InitVariables                                                      '⊙: Initializes local global variables

    Call SetToolbar("1100110100101111")										<%'버튼 툴바 제어 %>
    Call FncQuery
    
End Sub

Sub vspdData_Change(ByVal Col , ByVal Row )

   Frm1.vspdData.Row = Row
   Frm1.vspdData.Col = Col

   If Frm1.vspdData.CellType = parent.SS_CELL_TYPE_FLOAT Then
      If CDbl(Frm1.vspdData.text) < CDbl(Frm1.vspdData.TypeFloatMin) Then
         Frm1.vspdData.text = Frm1.vspdData.TypeFloatMin
      End If
   End If
   
   ggoSpread.Source = frm1.vspdData
   ggoSpread.UpdateRow Row

End Sub

Sub vspdData_Click(ByVal Col, ByVal Row)
    Call SetPopupMenuItemInf("1101011111") 

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
    Call ggoSpread.ReOrderingSpreadData()
End Sub

Sub vspdData_DragDropBlock(ByVal Col , ByVal Row , ByVal Col2 , ByVal Row2 , ByVal NewCol , ByVal NewRow , ByVal NewCol2 , ByVal NewRow2 , ByVal Overwrite , Action , DataOnly , Cancel )
    
    'ggoSpread.Source = vspdData
    Row = 0: Row2 = -1: NewRow = 0
    ggoSpread.SwapRange Col, Row, Col2, Row2, NewCol, NewRow, Cancel
    
End Sub

Sub vspdData_GotFocus()
    ggoSpread.Source = frm1.vspdData
End Sub

Function LoadMakeCalendar()
    
    PgmJump(BIZ_PGM_JUMP_ID)

End Function

Function FncQuery() 
    Dim IntRetCD 
    
    FncQuery = False                                                        '⊙: Processing is NG
    
    Err.Clear                                                               '☜: Protect system from crashing

    '-----------------------
    'Check previous data area
    '-----------------------
    
    If lgBlnFlgChgValue = True Then    
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "X", "X") '☜ 바뀐부분 
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")										'⊙: Clear Contents  Field
    ggoSpread.Source = frm1.vspdData
    ggoSpread.ClearSpreadData
    Call InitVariables				   											'⊙: Initializes local global variables
    
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then									'⊙: This function check indispensable field
       Exit Function
    End If
    
    '-----------------------
    'Query function call area
    '-----------------------
    If DbQuery = False Then Exit Function							<%'Query db data%>
       
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
		IntRetCD = DisplayMsgBox("900015", parent.VB_YES_NO, "X", "X") '☜ 바뀐부분 
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
    
    IntRetCD = DisplayMsgBox("900003", parent.VB_YES_NO, "X", "X")  '☜ 바뀐부분 
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

Function FncSave() 
        
    FncSave = False                                                         '⊙: Processing is NG
    
    Err.Clear                                                               '☜: Protect system from crashing
    'On Error Resume Next                                                    '☜: Protect system from crashing
    
    '-----------------------
    'Precheck area
    '-----------------------
    
    If ggoSpread.SSCheckChange = False Then		'lgBlnFlgChgValue = False
        Call DisplayMsgBox("900001", "X", "X", "X")                            '⊙: No data changed!!
        Exit Function
    End If
    
    '-----------------------
    'Check content area
    '-----------------------
    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSDefaultCheck = False Then   'Not chkField(Document, "2") OR  '⊙: Check contents area
       Exit Function
    End If
    
    '-----------------------
    'Save function call area
    '-----------------------
    If DbSave = False Then Exit Function                                        '☜: Save db data
    
    FncSave = True                                                          '⊙: Processing is OK
    
End Function

Function FncCopy() 
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncCopy = False                                                               '☜: Processing is NG

    If Frm1.vspdData.MaxRows < 1 Then
       Exit Function
    End If
 
    ggoSpread.Source = Frm1.vspdData

	With frm1.vspdData
		If .ActiveRow > 0 Then
			.focus
			.ReDraw = false
		
			ggoSpread.CopyRow
            SetSpreadColor .ActiveRow, .ActiveRow    
			'key field clear
			.Col=C_Month
			.Text=""
			
			.Col=C_Day
			.Text=""
		    				
			.ReDraw = true
		End If
	End with

    If Err.number = 0 Then	
       FncCopy = True                                                            '☜: Processing is OK
    End If
	
    Set gActiveElement = document.ActiveElement   
	
End Function

Function FncCancel() 
    ggoSpread.EditUndo                                                  '☜: Protect system from crashing
End Function

Function FncInsertRow(ByVal pvRowCnt)
    Dim IntRetCD
    Dim imRow
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status
    
    FncInsertRow = False                                                         '☜: Processing is NG
    
    If IsNumeric(Trim(pvRowCnt)) Then
        imRow = CInt(pvRowCnt)
    Else    
        imRow = AskSpdSheetAddRowCount()
        
        If imRow = "" Then
            Exit Function
        End If
    
    End If
    
	With frm1
	
	.vspdData.focus
    ggoSpread.Source = .vspdData
    '.vspdData.EditMode = True
    .vspdData.ReDraw = False
    ggoSpread.InsertRow ,imRow
    SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1
    .vspdData.ReDraw = True
    End With
    
    If Err.number = 0 Then
       FncInsertRow = True                                                          '☜: Processing is OK
    End If   
	
    Set gActiveElement = document.ActiveElement   
    
End Function

Function FncDeleteRow() 
    Dim lDelRows
    Dim iDelRowCnt, i
    
    With frm1.vspdData 
    
    .focus
    ggoSpread.Source = frm1.vspdData 
    
	.Col = 3'ggoSpread.SSGetColsIndex(8)
	If .Text = "A" Then
	    Call DisplayMsgBox("183104", "X", "X", "X")              '☆: you must release this line if you change msg into code
	    Exit Function
	End If
    
	lDelRows = ggoSpread.DeleteRow

    lgBlnFlgChgValue = True
    
    End With
End Function

Function FncPrint() 
    Call parent.FncPrint()                                                   '☜: Protect system from crashing
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

Function DbQuery() 
    Dim LngLastRow      
    Dim LngMaxRow       
    Dim LngRow          
    Dim strTemp         
    Dim StrNextKey      
    Dim comB15028         'As New B15028ListCommonHoliday

    DbQuery = False
    
    Call LayerShowHide(1)
    
    Err.Clear                                                               '☜: Protect system from crashing

	Dim strVal
    
    With frm1
    
    strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001							'☜: 
    strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
    
	Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 
    
    End With
    
    DbQuery = True
    

End Function

Function DbQueryOk()														'☆: 조회 성공후 실행로직 
	
    '-----------------------
    'Reset variables area
    '-----------------------
    lgIntFlgMode = parent.OPMD_UMODE												'⊙: Indicates that current mode is Update mode
    
	Call SetToolbar("1100111100111111")										<%'버튼 툴바 제어 %>
End Function

Function DbSave() 
    Dim comB15021     'As New B15021ControlCommonHoliday
    Dim lRow        
    Dim lGrpCnt     
    Dim retVal      
    Dim boolCheck   
    Dim lStartRow   
    Dim lEndRow     
    Dim lRestGrpCnt 
	Dim strVal, strDel	
	Dim sMon, sDay
	Dim iMon, iDay
	
    DbSave = False                                                          '⊙: Processing is NG
    
    'On Error Resume Next                                                   '☜: Protect system from crashing
	
	'Call LayerShowHide(1)
	
	With frm1
		.txtMode.value = parent.UID_M0002
		.txtInsrtUserId.value = parent.gUsrID
		
    
    '-----------------------
    'Data manipulate area
    '-----------------------
    lGrpCnt = 1    
    strVal = ""
    strDel = ""
    
    '-----------------------
    'Data manipulate area
    '-----------------------
    For lRow = 1 To .vspdData.MaxRows
    
        .vspdData.Row = lRow
        .vspdData.Col = 0
        
        Select Case .vspdData.Text        
						
            Case ggoSpread.InsertFlag											'☜: 신규 
				
				.vspdData.Col = C_Month
				sMon = frm1.vspdData.Text
				iMon = CInt(frm1.vspdData.Text)
			
				If Len(sMon) = 1 and CInt(sMon) < 10 Then 
					sMon = "0" & sMon
				End If
											
				.vspdData.Col = C_Day
				sDay = frm1.vspdData.Text
				iDay = CInt(frm1.vspdData.Text)
				
				If Len(sDay) = 1 and CInt(sDay) < 10 Then 
					sDay = "0" & sDay
				End If
				
				If ((iMon = 1 or iMon = 3 or iMon = 5 or iMon = 7 or iMon = 8 or iMon = 10 or iMon = 12) and iDay > 31) OR _
				   (iMon = 2 and iDay > 29) OR _
				   ((iMon = 4 or iMon = 6 or iMon = 9 or iMon = 11)	and iDay > 30) Then
					Call DisplayMsgBox("126404", "X", "X", "X")					
					Call ggoSpread.sserrorcol(lRow, C_Day)
					Exit Function
				End If
												
				strVal = strVal & "C" & parent.gColSep & sMon & sDay & parent.gColSep				'☜: C=Create
                
                .vspdData.Col = C_Remark	'2
                strVal = strVal & Trim(.vspdData.Text) & parent.gRowSep
                
                lGrpCnt = lGrpCnt + 1
                
            Case ggoSpread.UpdateFlag											'☜: 신규 
            
				.vspdData.Col = C_Month
				sMon = frm1.vspdData.Text
			
				If Len(sMon) = 1 and CInt(sMon) < 10 Then 
					sMon = "0" & sMon
				End If
				
				.vspdData.Col = C_Day
				sDay = frm1.vspdData.Text
			
				If Len(sDay) = 1 and CInt(sDay) < 10 Then 
					sDay = "0" & sDay
				End If
				
				strVal = strVal & "U" & parent.gColSep & sMon & sDay & parent.gColSep				'☜: U=Update

                .vspdData.Col = C_Remark	'2
                strVal = strVal & Trim(.vspdData.Text) & parent.gRowSep
                
                lGrpCnt = lGrpCnt + 1
                
            Case ggoSpread.DeleteFlag													'☜: 삭제 
				
				.vspdData.Col = C_Month
				sMon = frm1.vspdData.Text
			
				If Len(sMon) = 1 and CInt(sMon) < 10 Then 
					sMon = "0" & sMon
				End If
											
				.vspdData.Col = C_Day
				sDay = frm1.vspdData.Text
			
				If Len(sDay) = 1 and CInt(sDay) < 10 Then 
					sDay = "0" & sDay
				End If
				
				strDel = strDel & "D" & parent.gColSep & sMon & sDay & parent.gRowSep									
                
                lGrpCnt = lGrpCnt + 1
        End Select
                
    Next
	
	.txtMaxRows.value = lGrpCnt-1
	.txtSpread.value = strDel & strVal
	
	Call LayerShowHide(1)
	
	Call ExecMyBizASP(frm1, BIZ_PGM_ID)										'☜: 비지니스 ASP 를 가동 
	
	End With
	
    DbSave = True                                                           '⊙: Processing is NG
    
End Function

Function DbSaveOk()													'☆: 저장 성공후 실행 로직 
   
	Call InitVariables
	
    Call MainQuery()

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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>공통휴일</font></td>
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
					<TD WIDTH=100% * valign=top>
						<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD HEIGHT="100%">
									<script language =javascript src='./js/b1502ma1_I821282052_vspdData.js'></script>
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
	<TR HEIGHT=20>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_30%>>
				<TD WIDTH=* ALIGN=RIGHT><A href="vbscript:LoadMakeCalendar">칼렌다 생성</TD>
				<TD WIDTH=10>&nbsp;</TD>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="B1502mb1.asp" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
