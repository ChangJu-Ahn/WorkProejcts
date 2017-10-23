
<%@ LANGUAGE="VBSCRIPT" %>
<%'======================================================================================================
'*  1. Module Name          : Basis Architect
'*  2. Function Name        : 조직개편후부서정보등록 
'*  3. Program ID           : B2407ma1.asp
'*  4. Program Name         : B2407ma1.asp
'*  5. Program Desc         : 조직개편후부서정보등록 
'*  6. Modified date(First) : 2005/10/25
'*  7. Modified date(Last)  : 
'*  8. Modifier (First)     : Jeong Yong kyun
'*  9. Modifier (Last)      : 
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

Const BIZ_PGM_ID       = "B2407mb1.asp"												<%'비지니스 로직 ASP명 %>

Dim C_ORGID
Dim C_DEPT
Dim C_PDEPT
Dim C_PDEPT_POP
Dim C_PDEPT_NM
Dim C_LDEPTNM
Dim C_BUILDID
Dim C_LVL
Dim C_SEQ
Dim C_SDEPTNM
Dim C_EDEPTNM
Dim C_COST_CD
Dim C_COST_CD_POP
Dim C_COST_NM
Dim C_BIZ_UNIT_CD
Dim C_BIZ_UNIT_CD_POP
Dim C_BIZ_UNIT_NM
Dim C_ENDDEPTYN
Dim C_HIDDEN_FLAG

<!-- #Include file="../../inc/lgvariables.inc" -->

Dim IsOpenPop          
Dim lgStrPrevKey2

Sub InitSpreadPosVariables()
    C_ORGID           = 1
    C_DEPT            = 2
    C_LDEPTNM         = 3
    C_PDEPT           = 4
    C_PDEPT_POP       = 5   
    C_PDEPT_NM        = 6
    C_BUILDID         = 7
    C_LVL             = 8
    C_SEQ             = 9
    C_SDEPTNM         = 10
    C_EDEPTNM         = 11
    C_COST_CD         = 12
    C_COST_CD_POP     = 13   
    C_COST_NM         = 14
    C_BIZ_UNIT_CD     = 15
    C_BIZ_UNIT_CD_POP = 16   
    C_BIZ_UNIT_nm     = 17   
    C_ENDDEPTYN       = 18
    C_HIDDEN_FLAG     = 19
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
		ggoSpread.Spreadinit "V20021202",,parent.gAllowDragDropSpread    
		 
		.ReDraw = False

		.MaxCols = C_HIDDEN_FLAG + 1											'☜: 최대 Columns의 항상 1개 증가시킴 
		.Col = .MaxCols														'☆: 사용자 별 Hidden Column
		.ColHidden = True    
			       
		.MaxRows = 0
		ggoSpread.ClearSpreadData
	
		Call AppendNumberPlace("6","2","0")
		Call AppendNumberPlace("7","3","0")
		Call GetSpreadColumnPos("A")  

		ggoSpread.SSSetEdit   C_ORGID,           "부서개편ID"  , 12,,,5,2
		ggoSpread.SSSetEdit   C_DEPT,            "부서코드"    , 10,,,10,2  
		ggoSpread.SSSetEdit   C_LDEPTNM,         "부서명"      , 15,,,200,1
		ggoSpread.SSSetEdit   C_PDEPT,           "모부서"      , 10,,,10,2
		ggoSpread.SSSetButton C_PDEPT_POP
		ggoSpread.SSSetEdit   C_PDEPT_NM,        "모부서명"    , 15,,,200,1		
		ggoSpread.SSSetEdit   C_BUILDID,         "내부부서코드", 10,,,10
		ggoSpread.SSSetFloat  C_LVL,             "레벨"        ,  6,"6",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,,"","10"
		ggoSpread.SSSetFloat  C_SEQ,             "순서"        ,  6,"7",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,,"","999"
		ggoSpread.SSSetEdit   C_COST_CD,         "코스트센터"  , 10,,,10,1
		ggoSpread.SSSetButton C_COST_CD_POP
		ggoSpread.SSSetEdit   C_COST_NM,         "코스트센터명", 15,,,30,1
		ggoSpread.SSSetEdit   C_BIZ_UNIT_CD,     "사업부"      , 10,,,10,1    
		ggoSpread.SSSetButton C_BIZ_UNIT_CD_POP
		ggoSpread.SSSetEdit   C_BIZ_UNIT_NM,     "사업부명"    , 15,,,30,1
		ggoSpread.SSSetEdit   C_SDEPTNM,         "약칭부서명"  , 15,,,30,1
		ggoSpread.SSSetEdit   C_EDEPTNM,         "영문부서명"  , 15,,,100,1
		ggoSpread.SSSetCheck  C_ENDDEPTYN,       "말단부서여부", 14, 2, "말단부서", False
		ggoSpread.SSSetEdit   C_HIDDEN_FLAG,     ""  , 10,,,10,1

		Call ggoSpread.SSSetColHidden(C_HIDDEN_FLAG,C_HIDDEN_FLAG,True)				    

		.ReDraw = True

		Call SetSpreadLock("F")
    End With

End Sub

Sub SetSpreadLock(ByVal QryFg) 
	Dim ii

    With frm1
		.vspdData.ReDraw = False    
		If QryFg = "F" Then
			ggoSpread.SpreadLock C_ORGID, -1, C_HIDDEN_FLAG
			ggoSpread.SSSetProtected .vspdData.MaxCols, -1, -1		
		Else 
			ggoSpread.SSSetProtected .vspdData.MaxCols, -1, -1

			For ii = 1 To .vspddata.Maxrows
				.vspddata.col = C_ORGID
				.vspddata.row = ii
								
				If .vspddata.text = parent.gChangeOrgId Then
					ggoSpread.SSSetProtected C_ORGID          , ii, ii
'					ggoSpread.SpreadUnLock   C_DEPT           , ii, C_DEPT , ii
'					ggoSpread.SSSetRequired  C_DEPT           , ii, ii
'					ggoSpread.SpreadUnLock   C_PDEPT          , ii, C_PDEPT , ii					
'					ggoSpread.SSSetRequired  C_PDEPT          , ii, ii
'					ggoSpread.SpreadUnLock   C_PDEPT_POP      , ii, C_PDEPT_POP       , ii
					ggoSpread.SpreadLock     C_PDEPT_NM       , ii, C_PDEPT_NM        , ii
					ggoSpread.SpreadUnLock   C_LDEPTNM        , ii, C_LDEPTNM , ii					
					ggoSpread.SSSetRequired  C_LDEPTNM        , ii, ii
'					ggoSpread.SpreadUnLock   C_BUILDID        , ii, C_BUILDID , ii					
'					ggoSpread.SSSetRequired  C_BUILDID        , ii, ii
'					ggoSpread.SpreadUnLock   C_LVL            , ii, C_LVL , ii					
'					ggoSpread.SSSetRequired  C_LVL            , ii, ii
'					ggoSpread.SpreadUnLock   C_SEQ            , ii, C_SEQ , ii					
'					ggoSpread.SSSetRequired  C_SEQ            , ii, ii
					ggoSpread.SpreadUnLock   C_COST_CD        , ii, C_COST_CD , ii
					ggoSpread.SSSetRequired  C_COST_CD        , ii, ii
					ggoSpread.SpreadUnLock   C_COST_CD_POP    , ii, C_COST_CD_POP     , ii
					ggoSpread.SpreadLock     C_COST_NM        , ii, C_COST_NM         , ii    
					ggoSpread.SpreadUnLock   C_BIZ_UNIT_CD    , ii, C_BIZ_UNIT_CD , ii        					
					ggoSpread.SSSetRequired  C_BIZ_UNIT_CD    , ii, ii
					ggoSpread.SpreadUnLock   C_BIZ_UNIT_CD_POP, ii, C_BIZ_UNIT_CD_POP , ii
					ggoSpread.SpreadLock     C_BIZ_UNIT_NM    , ii, C_BIZ_UNIT_NM     , ii    
					ggoSpread.SpreadUnLock   C_SDEPTNM        , ii, C_PDEPT_POP       , ii
					ggoSpread.SSSetRequired  C_SDEPTNM        , ii, ii					
					ggoSpread.SpreadUnLock   C_EDEPTNM        , ii, C_PDEPT_POP       , ii    
'					ggoSpread.SpreadUnLock   C_ENDDEPTYN      , ii, C_ENDDEPTYN       , ii
				Else
					ggoSpread.SpreadLock C_ORGID, -1, C_HIDDEN_FLAG
					ggoSpread.SSSetProtected .vspdData.MaxCols, -1, -1						
				End If	
			Next
		End If
		.vspdData.ReDraw = True
	End With
End Sub

Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
    With frm1
		.vspdData.ReDraw = False

		ggoSpread.SSSetProtected C_ORGID          , pvStartRow, pvEndRow
		ggoSpread.SSSetRequired  C_DEPT           , pvStartRow, pvEndRow
		ggoSpread.SSSetRequired  C_PDEPT          , pvStartRow, pvEndRow
		ggoSpread.SpreadUnLock   C_PDEPT_POP      , pvStartRow, C_PDEPT_POP       , pvEndRow
		ggoSpread.SpreadLock     C_PDEPT_NM       , pvStartRow, C_PDEPT_NM        , pvEndRow
		ggoSpread.SSSetRequired  C_LDEPTNM        , pvStartRow, pvEndRow
		ggoSpread.SSSetProtected  C_BUILDID       , pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_LVL            , pvStartRow, pvEndRow		
		ggoSpread.SSSetProtected  C_SEQ           , pvStartRow, pvEndRow
		ggoSpread.SSSetRequired  C_COST_CD        , pvStartRow, pvEndRow
		ggoSpread.SpreadUnLock   C_COST_CD_POP    , pvStartRow, C_COST_CD_POP     , pvEndRow
		ggoSpread.SpreadLock     C_COST_NM        , pvStartRow, C_COST_NM         , pvEndRow    
		ggoSpread.SSSetRequired  C_BIZ_UNIT_CD    , pvStartRow, pvEndRow    
		ggoSpread.SpreadUnLock   C_BIZ_UNIT_CD_POP, pvStartRow, C_BIZ_UNIT_CD_POP , pvEndRow
		ggoSpread.SpreadLock     C_BIZ_UNIT_NM    , pvStartRow, C_BIZ_UNIT_NM     , pvEndRow    
		ggoSpread.SpreadUnLock   C_SDEPTNM        , pvStartRow, C_PDEPT_POP       , pvEndRow
		ggoSpread.SSSetRequired  C_SDEPTNM        , pvStartRow, pvEndRow							
		ggoSpread.SpreadUnLock   C_EDEPTNM        , pvStartRow, C_PDEPT_POP       , pvEndRow    
		ggoSpread.SSSetProtected C_ENDDEPTYN      , pvStartRow, pvEndRow    
    
		.vspdData.ReDraw = True
    End With
End Sub

Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

			C_ORGID           = iCurColumnPos(1)
			C_DEPT            = iCurColumnPos(2)
			C_LDEPTNM         = iCurColumnPos(3)
			C_PDEPT           = iCurColumnPos(4)
			C_PDEPT_POP       = iCurColumnPos(5)
			C_PDEPT_NM        = iCurColumnPos(6)
			C_BUILDID         = iCurColumnPos(7)
			C_LVL             = iCurColumnPos(8)
			C_SEQ             = iCurColumnPos(9)
			C_SDEPTNM         = iCurColumnPos(10)
			C_EDEPTNM         = iCurColumnPos(11)
			C_COST_CD         = iCurColumnPos(12)
			C_COST_CD_POP     = iCurColumnPos(13)
			C_COST_NM         = iCurColumnPos(14)
			C_BIZ_UNIT_CD     = iCurColumnPos(15)
			C_BIZ_UNIT_CD_POP = iCurColumnPos(16)
			C_BIZ_UNIT_nm     = iCurColumnPos(17)
			C_ENDDEPTYN       = iCurColumnPos(18)
			C_HIDDEN_FLAG     = iCurColumnPos(19)
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

'======================================================================================================
'   Function Name : OpenPopUp()
'   Function Desc : 
'=======================================================================================================
Function  OpenPopUp(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function 
 
	Select Case iWhere
		Case 1
			arrParam(0) = "부서코드팝업"											' 팝업 명칭 
			arrParam(1) = "horg_mas"													' TABLE 명칭 
			arrParam(2) = Trim(strCode)													' Code Condition
			arrParam(3) = " "															' Name Cindition
			arrParam(4) = " orgid = " & FilterVar(parent.gChangeOrgId,"''","S") & " "	' Where Condition
			arrParam(5) = "부서코드"												' 조건필드의 라벨 명칭 

			arrField(0) = "dept"														' Field명(0)
			arrField(1) = "ldeptnm"														' Field명(1)
		 
			arrHeader(0) = "부서코드"												' Header명(0)
			arrHeader(1) = "부서명"														' Header명(1)

		Case 2
			arrParam(0) = "코스트센터팝업"												' 팝업 명칭 
			arrParam(1) = "b_cost_center"												' TABLE 명칭 
			arrParam(2) = Trim(strCode)													' Code Condition
			arrParam(3) = " "															' Name Cindition
			arrParam(4) = " "															' Where Condition
			arrParam(5) = "코스트센터"													' 조건필드의 라벨 명칭 

			arrField(0) = "cost_cd"														' Field명(0)
			arrField(1) = "cost_nm"														' Field명(1)
		 
			arrHeader(0) = "코스트센터"													' Header명(0)
			arrHeader(1) = "코스트센터명"												' Header명(1)

		Case 3
			arrParam(0) = "사업부팝업"													' 팝업 명칭 
			arrParam(1) = "b_biz_unit"													' TABLE 명칭 
			arrParam(2) = strCode						 								' Code Condition
			arrParam(3) = " "															' Name Cindition
			arrParam(4) = " "															' Where Condition
			arrParam(5) = "사업부"			
	
			arrField(0) = "biz_unit_cd"													' Field명(0)
			arrField(1) = "biz_unit_nm"													' Field명(1)
    
    
			arrHeader(0) = "사업부"														' Header명(0)
			arrHeader(1) = "사업부명"													' Header명(1)
	End Select

	IsOpenPop = True
	 
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
	  "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")   
 
	IsOpenPop = False
 
	If arrRet(0) = "" Then     
		Exit Function
	Else
		Call SetPopup(arrRet, iWhere)
	End If
End Function

'======================================================================================================
'   Function Name : SetPopup(Byval arrRet)
'   Function Desc : 
'=======================================================================================================
Function SetPopup(Byval arrRet,Byval iWhere)
	With frm1
		Select Case iWhere
			Case 1 
				.vspdData.Col = C_PDEPT
				.vspdData.Text = arrRet(0)
				.vspdData.Col = C_PDEPT_NM
				.vspdData.Text = arrRet(1)
'				Call vspdData_Change(C_AcctCd, frm1.vspddata.activerow )  ' 변경이 일어났다고 알려줌        
				Call SetActiveCell(.vspdData,C_PDEPT,.vspdData.ActiveRow ,"M","X","X")
			Case 2
				.vspdData.Col = C_COST_CD
				.vspdData.Text = arrRet(0)
				.vspdData.Col = C_COST_NM
				.vspdData.Text = arrRet(1)
'				Call vspdData_Change(C_AcctCd, frm1.vspddata.activerow )  ' 변경이 일어났다고 알려줌        
				Call SetActiveCell(.vspdData,C_COST_CD,.vspdData.ActiveRow ,"M","X","X")
			Case 3 
			    .vspdData.Col = C_BIZ_UNIT_CD
				.vspdData.Text = arrRet(0)
				.vspdData.Col = C_BIZ_UNIT_NM
				.vspdData.Text = arrRet(1) 
'				Call vspdData_Change(.vspdData.Col, .vspdData.Row )  ' 변경이 일어났다고 알려줌       
				Call SetActiveCell(.vspdData,C_BIZ_UNIT_CD,.vspdData.ActiveRow ,"M","X","X")
		End Select    
	End With
 
	If iwhere <> 0 Then
		lgBlnFlgChgValue = True
	End If 
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
	Dim iTmpValue
	
	With frm1
		.vspddata.col = Col
		.vspddata.Row = row
					
		If .vspdData.CellType = parent.SS_CELL_TYPE_FLOAT Then
			If CDbl(.vspdData.text) < CDbl(.vspdData.TypeFloatMin) Then
				Frm1.vspdData.text = .vspdData.TypeFloatMin
			End If
		End If
	
		ggoSpread.Source = .vspdData
		ggoSpread.UpdateRow Row
	End With	
End Sub

'======================================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : 버튼 컬럼을 클릭할 경우 발생 
'=======================================================================================================
Sub  vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	Dim strTemp
	Dim intPos1

	With frm1
		ggoSpread.Source = .vspdData
		      
		If Row > 0 And Col = C_PDEPT_POP Then
		    .vspdData.Col = C_PDEPT
		    .vspdData.Row = Row
		    Call OpenPopup(.vspdData.Text, 1)
		    Call vspdData_Change(Col,Row)
		End If    
		       
		If Row > 0 And Col = C_COST_CD_POP Then
		    .vspdData.Col = C_COST_CD
		    .vspdData.Row = Row
		    Call OpenPopup(.vspdData.Text, 2)
		    Call vspdData_Change(Col,Row)		    
		End If    
		       
		If Row > 0 And Col = C_BIZ_UNIT_CD_POP Then
		    .vspdData.Col = C_BIZ_UNIT_CD
		    .vspdData.Row = Row
		    Call OpenPopup(.vspdData.Text, 3)
		    Call vspdData_Change(Col,Row)		    
		End If    
	End With
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

    If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	    <%'☜: 재쿼리 체크 %>
      
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

    ggoSpread.Source = frm1.vspdData
    If Not ggoSpread.SSDefaultCheck Then                                  '⊙: Check contents area
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

			.vspdData.Col = C_ORGID
			.vspdData.Text = parent.gChangeOrgId 		

			.vspdData.Col = C_DEPT
			.vspdData.Text = ""

			.vspdData.Col = C_LDEPTNM
			.vspdData.Text = ""

			.vspdData.Col = C_SDEPTNM
			.vspdData.Text = ""
			
			.vspdData.Col = C_EDEPTNM
			.vspdData.Text = ""			
    
			.vspdData.Col = C_BUILDID
			.vspdData.Text = ""
			
			.vspdData.Col = C_LVL
			.vspdData.Text = ""
			
			.vspdData.Col = C_SEQ
			.vspdData.Text = ""

			.vspdData.Col = C_ENDDEPTYN
			.vspdData.Text = "0"
			
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
    
	With frm1	
		.vspdData.focus
		ggoSpread.Source = .vspdData
		
		.vspdData.ReDraw = False
        ggoSpread.InsertRow ,imRow
        SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1
		
        For iRow = .vspdData.ActiveRow to .vspdData.ActiveRow + imRow - 1
        .vspdData.Row = iRow
        .vspdData.Col = C_ORGID
		.vspdData.Text = parent.gChangeOrgId
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

    With frm1
    	.vspdData.focus
    	ggoSpread.Source = .vspdData
    	
    	.vspddata.col = C_ORGID
    	If Trim(.vspddata.text) = parent.gChangeOrgId Then
    		lDelRows = ggoSpread.DeleteRow
    	Else
    	    Call DisplayMsgBox("124538", "X", "X", "X")                          <%' %1 ChangeOrgId of dept deleted is not current changeOrgId. %>
    	End If	
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

	Dim strVal
	
    DbQuery = False
    
    Err.Clear                                                               <%'☜: Protect system from crashing%>
	
	Call LayerShowHide(1)
	
	strVal = BIZ_PGM_ID & "?txtMode="& parent.UID_M0001						         
    strVal = strVal     & "&txtOrgid="& frm1.txtOrgId.Value					'☜: Query Key        

	Call RunMyBizASP(MyBizASP, strVal)										'☜:  Run biz logic

    DbQuery = True  
  
End Function

Function DbQueryOk()														<%'조회 성공후 실행로직 %>
    
    '-----------------------
    'Reset variables area
    '-----------------------
    lgIntFlgMode = parent.OPMD_UMODE
        
    Call SetToolbar("1100111100111111")										<%'버튼 툴바 제어 %>
    Call SetSpreadLock("Q")
	
End Function

Function DbSave() 
    Dim lRow        
	Dim strVal, strDel
	
	DbSave = False                                                          
    
    Call LayerShowHide(1)
    On Error Resume Next                                                   <%'☜: Protect system from crashing%>

	With frm1
		.txtMode.value = parent.UID_M0002
		
		strVal = ""
		strDel = ""
    
		For lRow = 1 To .vspdData.MaxRows
    
			.vspdData.Row = lRow
			.vspdData.Col = 0

			Select Case .vspdData.Text
			    Case ggoSpread.InsertFlag										'☜: 신규 
					strVal = strVal & "C" & parent.gColSep						'☜: C=Create
			    Case ggoSpread.UpdateFlag										'☜: 수정 
					strVal = strVal & "U" & parent.gColSep						'☜: U=Update
			End Select
				
			Select Case .vspdData.Text
			    Case ggoSpread.InsertFlag, ggoSpread.UpdateFlag					'☜: 수정, 신규 
						
			        .vspdData.Col = C_ORGID										
			        strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
			            
			        .vspdData.Col = C_DEPT	
			        strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
			            
			        .vspdData.Col = C_LDEPTNM
			        strVal = strVal & Trim(.vspdData.Text) & parent.gColSep

			        .vspdData.Col = C_PDEPT	           
			        strVal = strVal & Trim(.vspdData.Text) & parent.gColSep

			        .vspdData.Col = C_BUILDID
			        strVal = strVal & Trim(.vspdData.Text) & parent.gColSep

			        .vspdData.Col = C_LVL
			        strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
			            
			        .vspdData.Col = C_SEQ
			        strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
			            
			        .vspdData.Col = C_SDEPTNM
			        strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
			            
			        .vspdData.Col = C_EDEPTNM
			        strVal = strVal & Trim(.vspdData.Text) & parent.gColSep

			        .vspdData.Col = C_COST_CD
			        strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
			            
			        .vspdData.Col = C_BIZ_UNIT_CD
			        strVal = strVal & Trim(.vspdData.Text) & parent.gColSep

			        .vspdData.Col = C_ENDDEPTYN
		            If .vspdData.Value = 1 Then
			            strVal = strVal & "Y" & parent.gRowSep
			        Else
			            strVal = strVal & "N" & parent.gRowSep
					End If
			    Case ggoSpread.DeleteFlag										'☜: 삭제 

					strDel = strDel & "D" & parent.gColSep						'☜: D=Delete

			        .vspdData.Col = C_ORGID
			        strDel = strDel & Trim(.vspdData.Text) & parent.gColSep
			            
			        .vspdData.Col = C_DEPT
			        strDel = strDel & Trim(.vspdData.Text) & parent.gRowSep
			            
			End Select
		Next

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
										<INPUT TYPE=TEXT NAME="txtOrgId" SIZE=10 MAXLENGTH=5 tag="11XXXU"  ALT="부서개편ID"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnOrgId" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenOrgId()">
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
<!--	<TR HEIGHT=20>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_30%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>					
					<TD><BUTTON NAME="btnBatch" CLASS="CLSMBTN" Flag=1>회계조직반영</BUTTON></TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR> -->
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

