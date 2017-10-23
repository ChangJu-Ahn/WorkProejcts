<%@ LANGUAGE="VBSCRIPT" %>
<%'======================================================================================================
'*  1. Module Name          : Basis Architect
'*  2. Function Name        : Organization(horg_mas 부서정보등록)
'*  3. Program ID           : B2405ma1.asp
'*  4. Program Name         : B2405ma1.asp
'*  5. Program Desc         : 부서정보등록 
'*  6. Modified date(First) : 2000/10/26
'*  7. Modified date(Last)  : 2005/10/26
'*  8. Modifier (First)     : Hwang Jeong-won
'*  9. Modifier (Last)      : Jeong Yong Kyun
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

Const BIZ_PGM_ID = "B2405mb1.asp"												<%'비지니스 로직 ASP명 %>
Const BIZ_BATCH_PGM_ID= "B2405mb2.asp"											<%'내부부서코드생성 ASP명 %>
 
Dim C_ORGID
Dim C_DEPT
Dim C_PDEPT
Dim C_LDEPTNM
Dim C_BUILDID
Dim C_LVL
Dim C_SEQ
Dim C_SDEPTNM
Dim C_EDEPTNM
Dim C_ENDDEPTYN
Dim C_ENTRY_FG

<!-- #Include file="../../inc/lgvariables.inc" -->

Dim IsOpenPop          
Dim lgStrPrevKey2

Sub InitSpreadPosVariables()
    C_ORGID     = 1
    C_DEPT      = 2
    C_PDEPT     = 3
    C_LDEPTNM   = 4
    C_BUILDID   = 5
    C_LVL       = 6
    C_SEQ       = 7
    C_SDEPTNM   = 8
    C_EDEPTNM   = 9
    C_ENDDEPTYN = 10
    C_ENTRY_FG  = 11
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
End Sub

Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029A("I", "*","NOCOOKIE","MA") %>
End Sub

Sub InitSpreadSheet()
    Call initSpreadPosVariables()  

	With frm1.vspdData
	
	ggoSpread.Source = frm1.vspdData	
   'patch version
    ggoSpread.Spreadinit "V20021202",,parent.gAllowDragDropSpread    
    
	.ReDraw = false

    .MaxCols = C_ENTRY_FG + 1							'☜: 최대 Columns의 항상 1개 증가시킴 
	.Col = .MaxCols														'☆: 사용자 별 Hidden Column
	.ColHidden = True    
	       
    .MaxRows = 0
    ggoSpread.ClearSpreadData
	
	Call AppendNumberPlace("6","2","0")
	Call AppendNumberPlace("7","3","0")
    Call GetSpreadColumnPos("A")  

    ggoSpread.SSSetEdit  C_ORGID    , "부서개편ID", 12,,,5,2
    ggoSpread.SSSetEdit  C_DEPT     , "부서코드"  , 10,,,10,2  
    ggoSpread.SSSetEdit  C_PDEPT    , "모부서"    , 10,,,10,2
    ggoSpread.SSSetEdit  C_LDEPTNM  , "부서명", 20,,,200,1
    ggoSpread.SSSetEdit  C_BUILDID  , "내부부서코드", 10,,,10
    ggoSpread.SSSetFloat C_LVL      , "레벨" ,6,"6",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,,"1","10"
    ggoSpread.SSSetFloat C_SEQ      , "순서" ,6,"7",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,,"1","999"
    ggoSpread.SSSetEdit  C_SDEPTNM  , "약칭부서명", 15,,,24,1
    ggoSpread.SSSetEdit  C_EDEPTNM  , "영문부서명", 15,,,100,1
    ggoSpread.SSSetCheck C_ENDDEPTYN, "말단부서여부", 14, 2, "말단부서", False    
    ggoSpread.SSSetEdit  C_ENTRY_FG  , "", 4
    
	Call ggoSpread.SSSetColHidden(C_ENTRY_FG,C_ENTRY_FG,True)				        
	.ReDraw = true
	
    Call SetSpreadLock 
    
    End With
    
End Sub

Sub SetSpreadLock()
	Dim ii
		
    With frm1
		.vspdData.ReDraw = False
		ggoSpread.SpreadLock C_ORGID, -1, C_ORGID
		ggoSpread.SpreadLock C_DEPT, -1, C_DEPT
		ggoSpread.SSSetRequired C_LDEPTNM, -1, -1
		ggoSpread.SpreadLock C_BUILDID, -1, C_BUILDID
		ggoSpread.SSSetRequired C_LVL, -1, -1
		ggoSpread.SSSetRequired C_SEQ, -1, -1
		ggoSpread.SSSetRequired C_SDEPTNM, -1, -1
		'ggoSpread.SpreadLock C_ENDDEPTYN, -1, C_ENDDEPTYN
		ggoSpread.SSSetProtected .vspdData.MaxCols, -1, -1
		.vspdData.ReDraw = True

		For ii = 1 To .vspdData.MaxRows
			.vspddata.col = C_ENTRY_FG
			.vspddata.row = ii
			
			If Trim(.vspddata.text) = "E" Then			
				ggoSpread.SpreadLock C_ORGID, ii, C_ENTRY_FG ,ii
			End If
		Next
	End With	
End Sub

Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
    With frm1
    
    .vspdData.ReDraw = False
    ggoSpread.SSSetProtected C_ORGID, pvStartRow, pvEndRow
    ggoSpread.SSSetRequired C_DEPT, pvStartRow, pvEndRow
    ggoSpread.SSSetRequired C_LDEPTNM, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_BUILDID, pvStartRow, pvEndRow
    ggoSpread.SSSetRequired C_LVL, pvStartRow, pvEndRow
    ggoSpread.SSSetRequired C_SEQ, pvStartRow, pvEndRow
    ggoSpread.SSSetRequired C_SDEPTNM, pvStartRow, pvEndRow  
    'ggoSpread.SSSetProtected C_ENDDEPTYN, pvStartRow, pvEndRow
    .vspdData.ReDraw = True
    
    End With
End Sub

Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
            C_ORGID     = iCurColumnPos(1)
            C_DEPT      = iCurColumnPos(2)
            C_PDEPT     = iCurColumnPos(3)
            C_LDEPTNM   = iCurColumnPos(4)
            C_BUILDID   = iCurColumnPos(5)
            C_LVL       = iCurColumnPos(6)
            C_SEQ       = iCurColumnPos(7)
            C_SDEPTNM   = iCurColumnPos(8)
            C_EDEPTNM   = iCurColumnPos(9)
            C_ENDDEPTYN = iCurColumnPos(10)
            C_ENTRY_FG  = iCurColumnPos(11)
    End Select    
End Sub

Function OpenOrgId()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "부서개편ID 팝업"				<%' 팝업 명칭 %>
	arrParam(1) = "horg_abs"						<%' TABLE 명칭 %>
	arrParam(2) = frm1.txtOrgId.value			    <%' Code Condition%>
	arrParam(3) = ""								<%' Name Cindition%>
	arrParam(4) = ""								<%' Where Condition%>
	arrParam(5) = "부서개편ID"					<%' 조건필드의 라벨 명칭 %>
	
    arrField(0) = "orgid"							<%' Field명(0)%>
    arrField(1) = "orgnm"							<%' Field명(1)%>
    
    arrHeader(0) = "부서개편ID"					<%' Header명(0)%>
    arrHeader(1) = "부서개편명"					<%' Header명(1)%>
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	frm1.txtOrgId.focus
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetOrgId(arrRet)
	End If	
	
End Function

Function SetOrgId(Byval arrRet)
	With frm1
		.txtOrgId.value = arrRet(0)
		.txtOrgNm.value = arrRet(1)
	End With
End Function

Sub Form_Load()

    Call LoadInfTB19029                                                     <%'Load table , B_numeric_format%>
    
    Call ggoOper.LockField(Document, "N")                                   <%'Lock  Suitable  Field%>                         
                                                                            <%'Format Numeric Contents Field%>                                                                            
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call InitSpreadSheet                                                    <%'Setup the Spread sheet%>
    Call InitVariables                                                      <%'Initializes local global variables%>
    
    Call SetToolbar("1100110100101111")										<%'버튼 툴바 제어 %>
    frm1.txtOrgId.focus
    
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

Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
	Dim index
	
	With frm1.vspdData
		If Col = C_RegionNm And Row > 0 Then
			.Row = Row
			.Col = Col
			index = .TypeComboBoxCurSel
			
			.Col = C_Region
			.TypeComboBoxCurSel = index
		End If
	End With
End Sub

Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If CheckRunningBizProcess = True Then					'☜: 조회중이면 다음 조회 안하도록 체크 
        Exit Sub
	End If
	
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	    <%'☜: 재쿼리 체크 %>
      
    	If lgStrPrevKey <> "" And lgStrPrevKey2 <> "" Then                  <%'다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 %>
      		Call DisableToolBar(parent.TBC_QUERY)					'☜ : Query 버튼을 disable 시킴.
			If DBQuery = False Then 
			   Call RestoreToolBar()
			   Exit Sub 
			End If 
    	End If

    End If
End Sub

Function FncQuery() 
    Dim IntRetCD 
    
    FncQuery = False                                                        
    
    Err.Clear                                                               <%'Protect system from crashing%>

<%  '-----------------------
    'Check previous data area
    '----------------------- %>
    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "X", "X")			    <%'데이타가 변경되었습니다. 조회하시겠습니까?%>
    	If IntRetCD = vbNo Then
      	Exit Function
    	End If
    End If
    
<%  '-----------------------
    'Erase contents area
    '----------------------- %>
    frm1.txtOrgNm.value = ""
    Call ggoOper.ClearField(Document, "2")									<%'Clear Contents  Field%>
    ggoSpread.ClearSpreadData
    Call InitVariables                                                      <%'Initializes local global variables%>
    															
<%  '-----------------------
    'Check condition area
    '----------------------- %>
    If Not chkField(Document, "1") Then								<%'This function check indispensable field%>
       Exit Function
    End If    
	
     
    CALL dbquery()
    
End Function

Function FncSave() 
        
    FncSave = False                                                         
    
    Err.Clear                                                               <%'☜: Protect system from crashing%>
    On Error Resume Next                                                   <%'☜: Protect system from crashing%>    
    
<%  '-----------------------
    'Precheck area
    '----------------------- %>
    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = False Then
        Call DisplayMsgBox("900001", "X", "X", "X")                          <%'No data changed!!%>
        Exit Function
    End If
    
<%  '-----------------------
    'Save function call area
    '----------------------- %>
    If DbSave = False Then Exit Function                                        '☜: Save db data
    
    FncSave = True                                                          
    
End Function

Function FncCopy() 
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncCopy = False                                                               '☜: Processing is NG

    If Frm1.vspdData.MaxRows < 1 Then
       Exit Function
    End If
 
    ggoSpread.Source = Frm1.vspdData

	With frm1
		If .vspdData.ActiveRow > 0 Then
			.vspdData.focus
			.vspdData.ReDraw = False
		
			ggoSpread.CopyRow
			SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow

			.vspdData.Col = C_DEPT
			.vspdData.Text = ""
    
			.vspdData.Col = C_BUILDID
			.vspdData.Text = ""
			
			.vspdData.Col = C_LVL
			.vspdData.Text = ""
			
			.vspdData.Col = C_SEQ
			.vspdData.Text = ""
			
			.vspdData.ReDraw = True
		End If
	End With
	
    If Err.number = 0 Then	
       FncCopy = True                                                            '☜: Processing is OK
    End If
	
    Set gActiveElement = document.ActiveElement   
	
End Function

Function FncCancel() 
    ggoSpread.Source = frm1.vspdData	
    ggoSpread.EditUndo                                                  '☜: Protect system from crashing
End Function

Function FncInsertRow(ByVal pvRowCnt) 
    Dim IntRetCD
    Dim imRow
    Dim iRow

    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
    Dim uCountID

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
    
    '''''''''Insert하기전 부서개편ID존재유무Check(PK이기때문에 체크필수)
    Call CommonQueryRs(" Count(ORGID) "," HORG_ABS "," ORGID =  " & FilterVar(frm1.txtOrgId.value, "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) 
    uCountID = Trim(Replace(lgF0,Chr(11),"")) 
    
    If Cint(uCountID) <= 0 Then
        Call DisplayMsgBox("970029", "X","부서개편ID", "X")
        Exit Function
    End if
    ''''''''''''''''''''''''''''''''''Srh;2002/09/03
	With frm1	
		.vspdData.focus
		ggoSpread.Source = .vspdData
		
		.vspdData.ReDraw = False
        ggoSpread.InsertRow ,imRow
        SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1
		
        For iRow = .vspdData.ActiveRow to .vspdData.ActiveRow + imRow - 1
        .vspdData.Row = iRow
        .vspdData.Col = C_ORGID
		.vspdData.Text = UCase(.txtOrgId.value)
		Next				
		.vspdData.ReDraw = True		
		''SetSpreadColor .vspdData.ActiveRow    
    End With

    If Err.number = 0 Then
       FncInsertRow = True                                                          '☜: Processing is OK
    End If   
	
    Set gActiveElement = document.ActiveElement   
    
End Function

Function FncDeleteRow() 
    Dim lDelRows

    With frm1.vspdData 
    	.focus
    	ggoSpread.Source = frm1.vspdData 
    	lDelRows = ggoSpread.DeleteRow
    End With
    
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

    DbQuery = False
    
    Err.Clear                                                               <%'☜: Protect system from crashing%>
	
	Call LayerShowHide(1)
	
	Dim strVal
    
    With Frm1
    
		strVal = BIZ_PGM_ID & "?txtMode="            & parent.UID_M0001						         
        strVal = strVal     & "&txtOrgid="           & Frm1.txtOrgId.Value      '☜: Query Key        
        strVal = strVal     & "&txtMaxRows="         & .vspdData.MaxRows            

    Call RunMyBizASP(MyBizASP, strVal)   
    End With                                           '☜:  Run biz logic

    DbQuery = True  
  
End Function

Function DbQueryOk()													<%'조회 성공후 실행로직 %>
	
    Dim iNameArr
    Dim iDx    
    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
    
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'부서개편 id nm값 리턴 
    call CommonQueryRs(" ORGNM "," HORG_ABS "," ORGID =  " & FilterVar(frm1.txtOrgId.value, "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)  
    frm1.txtOrgnm.value = Trim(Replace(lgF0,Chr(11),""))    
	
    '-----------------------
    'Reset variables area
    '-----------------------
    lgIntFlgMode = parent.OPMD_UMODE
        
    Call SetToolbar("1100111100111111")										<%'버튼 툴바 제어 %>
	Call SetSpreadLock
End Function

Function DbSave() 
    Dim lRow        
    Dim lGrpCnt  
	Dim strVal, strDel, sPDept
	
	DbSave = False                                                          
    
    Call LayerShowHide(1)
    On Error Resume Next                                                   <%'☜: Protect system from crashing%>

	With frm1
		.txtMode.value = parent.UID_M0002
		
<%  '-----------------------
    'Data manipulate area
    '----------------------- %>
    lGrpCnt = 1
    
    strVal = ""
    strDel = ""
    
<%  '-----------------------
    'Data manipulate area
    '----------------------- %>
    ' Data 연결 규칙 
    ' 0: Flag , 1: Row위치, 2~N: 각 데이타 

    For lRow = 1 To .vspdData.MaxRows
    
		    .vspdData.Row = lRow
		    .vspdData.Col = 0

		    Select Case .vspdData.Text
		        Case ggoSpread.InsertFlag								'☜: 신규 
					strVal = strVal & "C" & parent.gColSep						'☜: C=Create
		        Case ggoSpread.UpdateFlag								'☜: 수정 
					strVal = strVal & "U" & parent.gColSep						'☜: U=Update
			End Select
			
		    Select Case .vspdData.Text

		        Case ggoSpread.InsertFlag, ggoSpread.UpdateFlag			'☜: 수정, 신규 
					
		            .vspdData.Col = C_ORGID	'2
		            strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		            
		            .vspdData.Col = C_DEPT		'3
		            strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		            
		            .vspdData.Col = C_PDEPT		'4		            
		            strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		            sPDept = Trim(.vspdData.Text)
		            
		            .vspdData.Col = C_LDEPTNM		'5
		            If Trim(.vspdData.Text) = "" Then
						Call LayerShowHide(0)
						.vspdData.Row = 0
						Call DisplayMsgBox("970021", "X", .vspdData.Text , "X")
						.vspdData.Focus
						.vspdData.Row = lRow
						.vspdData.Col = C_LDEPTNM
						.vspdData.Action = 0
						Exit Function
		            End If
		            strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		            
		            .vspdData.Col = C_BUILDID		'6
		            strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		            
		            .vspdData.Col = C_LVL		'7
		            strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		            If Trim(.vspdData.Text) <> "1" and sPDept = "" Then
						Call LayerShowHide(0)
						.vspdData.Col = C_PDEPT
						.vspdData.Row = 0
						Call DisplayMsgBox("970021", "X", .vspdData.Text , "X")
						.vspdData.Focus
						.vspdData.Row = lRow
						.vspdData.Action = 0
						Exit Function
		            End If
		            
		            .vspdData.Col = C_SEQ		'8
		            strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		            
		            .vspdData.Col = C_SDEPTNM		'9
		            If Trim(.vspdData.Text) = "" Then
						Call LayerShowHide(0)
						.vspdData.Row = 0
						Call DisplayMsgBox("970021", "X", .vspdData.Text , "X")
						.vspdData.Focus
						.vspdData.Row = lRow
						.vspdData.Col = C_SDEPTNM
						.vspdData.Action = 0
						Exit Function
		            End If
		            strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		            
		            .vspdData.Col = C_EDEPTNM		'10
		            strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		            
		            .vspdData.Col = C_ENDDEPTYN		'11
	                If .vspdData.Value = 1 Then
			            strVal = strVal & "Y" & parent.gRowSep
			        Else
			            strVal = strVal & "N" & parent.gRowSep
					End If
					
		            lGrpCnt = lGrpCnt + 1
		            
		        Case ggoSpread.DeleteFlag								'☜: 삭제 

					strDel = strDel & "D" & parent.gColSep						'☜: U=Update

		            .vspdData.Col = C_ORGID	'2
		            strDel = strDel & Trim(.vspdData.Text) & parent.gColSep
		            
		            .vspdData.Col = C_DEPT		'3
		            strDel = strDel & Trim(.vspdData.Text) & parent.gRowSep
		            
		            lGrpCnt = lGrpCnt + 1
		    End Select
                
    Next
		
	.txtMaxRows.value = lGrpCnt-1	
	.txtSpread.value = strDel & strVal
	
	Call ExecMyBizASP(frm1, BIZ_PGM_ID)										<%'☜: 비지니스 ASP 를 가동 %>
	
	End With
	
    DbSave = True                                                           
    
End Function

Function DbSaveOk()													        <%' 저장 성공후 실행 로직 %>
	Call InitVariables
	frm1.vspdData.MaxRows = 0
    ggoSpread.Source = frm1.vspdData
    ggoSpread.ClearSpreadData
	
    Call MainQuery()
End Function

Function btnBatch_OnClick()
	Dim strVal
	Dim IntRetCD
		
	If frm1.txtOrgId.value = "" Then		
		Call DisplayMsgBox("970029", "X","부서개편ID", "X")
		Exit function
	End If
       
    IntRetCD = DisplayMsgBox("900018", parent.VB_YES_NO, "X", "X")
	
	If IntRetCD = vbNo Then
		Exit Function
	End If
	
	Call LayerShowHide(1)
	
    strVal = BIZ_BATCH_PGM_ID & "?txtMode=Gen"
	strVal = strVal & "&txtOrgId=" & frm1.txtOrgId.value
	
	Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 
    
End Function

Function Batch_OK()
	Call InitVariables
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" NOWRAP><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
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
								<TR>
									<TD CLASS="TD5">부서개편ID</TD>
									<TD CLASS="TD656">
										<INPUT TYPE=TEXT NAME="txtOrgId" SIZE=10 MAXLENGTH=5 tag="12XXXU"  ALT="부서개편ID"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnOrgId" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenOrgId()">
										<INPUT TYPE=TEXT NAME="txtOrgNm" Size=40 tag="14">
									</TD>
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
						<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD HEIGHT="100%">
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=vaSpread1> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
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
				<TR>
					<TD WIDTH=10>&nbsp;</TD>					
					<TD><BUTTON NAME="btnBatch" CLASS="CLSMBTN" Flag=1>내부부서코드 생성</BUTTON></TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="B2405mb1.asp" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="hOrgId" tag="24">
<INPUT TYPE=HIDDEN NAME="hDept" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

