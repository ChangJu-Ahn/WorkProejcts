<%@ LANGUAGE="VBSCRIPT" %>
<!--
'======================================================================================================
'*  1. Module Name          : Accounting
'*  2. Function Name        : 
'*  3. Program ID           : A5966MA1
'*  4. Program Name         : 월차발생 부서등록 
'*  5. Program Desc         : Single-Multi Sample
'*  6. Component List       :
'*  7. Modified date(First) : 2001/01/30
'*  8. Modified date(Last)  : 2003/05/30
'*  9. Modifier (First)     : PARK JAI HONG
'* 10. Modifier (Last)      : Jung Sung Ki
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'=======================================================================================================-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>



<!-- #Include file="../../inc/incSvrCcm.inc"  -->
<!-- #Include file="../../inc/incSvrHTML.inc"  -->


<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		


<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliMAMain.vbs">  </SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliMAEvent.vbs">  </SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliVariables.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliMAOperation.vbs"> </SCRIPT>

<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/Cookie.vbs"></SCRIPT>

<SCRIPT LANGUAGE="javascript" SRC="../../inc/TabScript.js">   </SCRIPT>

<Script Language="VBScript">
Option Explicit                                                        '☜: indicates that All variables must be declared in advance

'========================================================================================================
Const BIZ_PGM_ID = "A5966MB1.asp"                                      'Biz Logic ASP 

'========================================================================================================
Dim C_BIZ_CD
Dim C_BIZ_PB
Dim C_BIZ_NM
Dim C_DEPT_CD
Dim C_DEPT_PB
Dim C_DEPT_NM
Dim C_ORG_CHANGE_ID
Dim C_INTERNAL_CD


Const COOKIE_SPLIT      = 4877	                                      '☆: Cookie Split String
'========================================================================================================

<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================

Dim IsOpenPop                   
Dim StartDate
<%
StartDate	= GetSvrDate                                               'Get Server DB Date
%>



'========================================================================================================
Sub initSpreadPosVariables()
	 C_BIZ_CD = 1
	 C_BIZ_PB = 2
	 C_BIZ_NM = 3
	 C_DEPT_CD = 4
	 C_DEPT_PB = 5
	 C_DEPT_NM = 6
	 C_ORG_CHANGE_ID = 7
	 C_INTERNAL_CD = 8
End Sub
'========================================================================================================
' Name : InitVariables()	
' Desc : Initialize value
'========================================================================================================
Sub InitVariables()
	lgIntFlgMode      = parent.OPMD_CMODE						        '⊙: Indicates that current mode is Create mode
	lgBlnFlgChgValue  = False								    '⊙: Indicates that no value changed
	lgIntGrpCount     = 0										'⊙: Initializes Group View Size
    lgStrPrevKey      = ""                                      '⊙: initializes Previous Key
    lgStrPrevKeyIndex = ""                                      '⊙: initializes Previous Key Index
    lgSortKey         = 1                                       '⊙: initializes sort direction
		
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub


'========================================================================================================
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029A("I", "*", "NOCOOKIE", "MA") %>
<% Call LoadBNumericFormatA("I", "*","NOCOOKIE","MA") %>
End Sub

'========================================================================================================
' Name : CookiePage()
' Desc : Write or Read cookie value 
'========================================================================================================
Sub CookiePage(Kubun)

End Sub

'========================================================================================================
' Name : MakeKeyStream
' Desc : Make key stream of query or delete condition data
'========================================================================================================
Sub MakeKeyStream(pRow)
    Dim strYYYYMM
    Dim strYear,strMonth,strDay
End Sub        

'========================================================================================================
' Name : InitComboBox()
' Desc : Set ComboBox
'========================================================================================================
Sub InitComboBox()
End Sub

'========================================================================================================
' Name : InitData()
' Desc : Reset Combox
'========================================================================================================
Sub InitData()
	Dim intRow
	Dim intIndex 
	
	With frm1.vspdData
		For intRow = 1 To .MaxRows			
			.Row = intRow
		'	.Col = C_TYPECd
			intIndex = .value
		'	.col = C_TYPENm
			.value = intindex					
		Next	
	End With
End Sub

'======================================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : 컬럼을 클릭할 경우 발생 
'=======================================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	frm1.vspdData.Row = Row
	Select Case Col
	    Case C_BIZ_PB   ' 사업부코드 
	        frm1.vspdData.Col = C_BIZ_CD
	        Call OpenCost(frm1.vspdData.Text, 1,ROW)
	    Case C_DEPT_PB		'부서코드 
	        frm1.vspdData.Col = C_DEPT_CD
	        Call OpenCost(frm1.vspdData.Text, 2,ROW)
	End Select
	Call SetActiveCell(frm1.vspdData,Col - 1,frm1.vspdData.ActiveRow ,"M","X","X")
End Sub
'===========================================================================
' Function Name : OpenCost
' Function Desc : OpenCost Reference Popup
'===========================================================================
Function OpenCost(strCode, iWhere, Row)

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim biz_cd1
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	Select Case iWhere
	    Case 1
	    	arrParam(0) = "사업장 팝업"					 ' 팝업 명칭 
	    	arrParam(1) = "B_BIZ_AREA"	' TABLE 명칭 
	    	arrParam(2) = strCode	                        ' Code Condition
	    	arrParam(3) = ""								' Name Cindition
	    	arrParam(4) = "" 								' Where Condition
	    	arrParam(5) = "사업장코드"		   			' TextBox 명칭 
	
	    	arrField(0) = "BIZ_AREA_CD"		                ' Field명(0)
	    	arrField(1) = "BIZ_AREA_NM"    					' Field명(1)%>
    
	    	arrHeader(0) = "사업장코드"		        	' Header명(0)%>
	    	arrHeader(1) = "사업장명"	        		' Header명(1)%>
			


	    Case 2
			Frm1.vspdData.Col = C_BIZ_CD
			Frm1.vspdData.Row = Row
			biz_cd1 = Trim(Frm1.vspdData.Text)
			arrParam(0) = "부서 팝업"								  ' 팝업 명칭	    
			arrParam(1) = "B_ACCT_DEPT A, B_COST_CENTER B "			<%' TABLE 명칭 %>
			arrParam(2) = strCode									<%' Code Condition%>
			arrParam(3) = "" 		            					<%' Name Cindition%>
			arrParam(4) = "A.COST_CD = B.COST_CD AND A.ORG_CHANGE_ID = (SELECT CUR_ORG_CHANGE_ID FROM B_COMPANY) AND B.BIZ_AREA_CD = " & FilterVar(biz_cd1, "''", "S")
			arrParam(5) = "부서코드"			
					
			arrField(0) = "A.DEPT_CD"	     							<%' Field명(1)%>
			arrField(1) = "A.DEPT_NM"									<%' Field명(0)%>
			arrField(2) = "A.ORG_CHANGE_ID"									<%' Field명(0)%>
			arrField(3) = "A.INTERNAL_CD"									<%' Field명(0)%>
				   
				    
			arrHeader(0) = "부서코드"			   				<%' Header명(0)%>
			arrHeader(1) = "부서명"								<%' Header명(1)%>
			arrHeader(2) = "조직변경ID"			   				<%' Header명(2)%>
			arrHeader(3) = "내부부서코드"								<%' Header명(3)%>
    		
	End Select
  
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetCost(arrRet, iWhere, Row)
	End If	
	
End Function

'------------------------------------------  SetSItemDC()  --------------------------------------------------
'	Name : SetCost()
'	Description : OpenSItemDC Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetCost(arrRet, iWhere, Row)


	With frm1
        .vspdData.Row = Row
		Select Case iWhere
		    Case 1
		        .vspdData.Col = C_BIZ_CD
		    	.vspdData.text = arrRet(0) 
		    	.vspdData.Col = C_BIZ_NM
		    	.vspdData.text = arrRet(1)
		    	Call vspdData_Change(C_BIZ_CD, Row )
		    Case 2
		        .vspdData.Col = C_DEPT_CD
		    	.vspdData.text = arrRet(0) 
		    	.vspdData.Col = C_DEPT_NM
		    	.vspdData.text = arrRet(1)
		    	.vspdData.Col = C_ORG_CHANGE_ID
		    	.vspdData.text = arrRet(2)
		    	.vspdData.Col = C_INTERNAL_CD
		    	.vspdData.text = arrRet(3)
		End Select

		lgBlnFlgChgValue = True

	End With
	ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row
End Function

'========================================================================================================
' Name : InitSpreadSheet
' Desc : This method initializes spread sheet column property
'========================================================================================================
Sub InitSpreadSheet()
	Call InitSpreadPosVariables()
	With frm1.vspdData
		ggoSpread.Source = frm1.vspdData
		ggoSpread.Spreadinit "V20021103",,parent.gAllowDragDropSpread 
	   .ReDraw = false
	   		
       .MaxCols = C_INTERNAL_CD + 1                                                 ' ☜:☜: Add 1 to Maxcols
	   .Col = .MaxCols                                                              ' ☜:☜: Hide maxcols
       .ColHidden = True                                                            ' ☜:☜:
'       .Col = C_ORG_CHANGE_ID                                                       ' ☜:☜: Hide ORG_CHANGE_ID
'       .ColHidden = True                                                            ' ☜:☜:
'       .Col = C_INTERNAL_CD                                                         ' ☜:☜: Hide INTERNAL_CD
'       .ColHidden = True                                                            ' ☜:☜:
       '.MaxRows = 0
		ggoSpread.Source = frm1.vspdData
		ggospread.ClearSpreadData		'Buffer Clear
	
		Call GetSpreadColumnPos("A")
		ggoSpread.SSSetEdit		C_BIZ_CD	, "사업장코드" ,25,   ,, 25
		ggoSpread.SSSetButton	C_BIZ_PB
		ggoSpread.SSSetEdit		C_BIZ_NM	, "사업장명"   ,30,   ,, 30
		ggoSpread.SSSetEdit		C_DEPT_CD	, "부서코드" ,25,   ,, 25
		ggoSpread.SSSetButton	C_DEPT_PB
		ggoSpread.SSSetEdit		C_DEPT_NM	, "부서명"   ,30,   ,, 30
		ggoSpread.SSSetEdit		C_ORG_CHANGE_ID , "조직변경ID"   ,20,,, 35,2
		ggoSpread.SSSetEdit		C_INTERNAL_CD	, "내부부서코드"   ,20,,, 35,2
	
		Call ggoSpread.MakePairsColumn(C_BIZ_CD,C_BIZ_PB)
		Call ggoSpread.MakePairsColumn(C_DEPT_CD,C_DEPT_PB)
		Call ggoSpread.SSSetColHidden(C_ORG_CHANGE_ID,C_ORG_CHANGE_ID,True)
		Call ggoSpread.SSSetColHidden(C_INTERNAL_CD,C_INTERNAL_CD,True)
		
		
	   .ReDraw = true
	
       Call SetSpreadLock 
    End With
End Sub

'======================================================================================================
' Name : SetSpreadLock
' Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadLock()
    With frm1
    
    .vspdData.ReDraw = False
      ggoSpread.SpreadLock    C_BIZ_CD	, -1, C_BIZ_CD
      ggoSpread.SpreadLock	  C_BIZ_PB	, -1, C_BIZ_PB
      ggoSpread.SpreadLock    C_BIZ_NM	, -1, C_BIZ_NM
      ggoSpread.SSSetRequired C_DEPT_CD , -1, C_DEPT_CD
      ggoSpread.SpreadLock    C_DEPT_NM	, -1, C_DEPT_NM
      ggoSpread.SpreadLock	  .vspdData.MaxCols, -1,.vspdData.MaxCols
    .vspdData.ReDraw = True

    End With
End Sub

'======================================================================================================
' Name : SetSpreadColor
' Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadColor(ByVal pvStarRow, ByVal pvEndRow)
    With frm1
	    .vspdData.ReDraw = False
			ggoSpread.SSSetRequired    C_BIZ_CD , pvStarRow, pvEndRow
			ggoSpread.SSSetProtected   C_BIZ_NM , pvStarRow, pvEndRow
			ggoSpread.SSSetRequired    C_DEPT_CD , pvStarRow, pvEndRow
			ggoSpread.SSSetProtected   C_DEPT_NM , pvStarRow, pvEndRow
	    .vspdData.ReDraw = True
    End With
End Sub

Sub GetSpreadColumnPos(ByVal pvSpdNo)
	Dim iCurColumnPos
	Select Case UCase(pvSpdNo)
	Case "A"
		ggoSpread.Source = frm1.vspdData
		Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
		C_BIZ_CD = iCurColumnPos(1)
		C_BIZ_PB = iCurColumnPos(2)
		C_BIZ_NM = iCurColumnPos(3)
		C_DEPT_CD = iCurColumnPos(4)
		C_DEPT_PB = iCurColumnPos(5)
		C_DEPT_NM = iCurColumnPos(6)
		C_ORG_CHANGE_ID = iCurColumnPos(7)
		C_INTERNAL_CD = iCurColumnPos(8)
	End Select
End Sub

'======================================================================================================
' Name : SubSetErrPos
' Desc : This method set focus to position of error
'      : This method is called in MB area
'======================================================================================================
Sub SubSetErrPos(iPosArr)
    Dim iDx
    Dim iRow
    iPosArr = Split(iPosArr,parent.gColSep)
    If IsNumeric(iPosArr(0)) Then
       iRow = CInt(iPosArr(0))
       For iDx = 1 To  frm1.vspdData.MaxCols - 1
           Frm1.vspdData.Col = iDx
           Frm1.vspdData.Row = iRow
           If Frm1.vspdData.ColHidden <> True And Frm1.vspdData.BackColor <> parent.UC_PROTECTED Then
              Frm1.vspdData.Col = iDx
              Frm1.vspdData.Row = iRow
              Frm1.vspdData.Action = 0 ' go to 
              Exit For
           End If
           
       Next
          
    End If   
End Sub

Sub PopSaveSpreadColumnInf()
	ggoSpread.Source = gActiveSpdSheet
	Call ggoSpread.SaveSpreadColumnInf()
End Sub

Sub PopRestoreSpreadColumnInf()
	ggoSpread.Source = gActiveSpdSheet
	Call ggoSpread.RestoreSpreadInf()
	Call InitSpreadSheet()
	Call InitComboBox
	Call ggoSpread.ReOrderingSpreadData()
	Call InitData()
End Sub

'========================================================================================================
Sub Form_Load()
    Err.Clear
	Call LoadInfTB19029
	Call ggoOper.LockField(Document, "N")
    Call InitSpreadSheet
	Call InitVariables
	Call SetToolbar("1100110100101111")                                              '☆: Developer must customize
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Call InitComboBox
'	Call CookiePage (0)                                                              '☜: Check Cookie
			
End Sub
	
'========================================================================================================
' Name : Form_QueryUnload
' Desc : This sub is called from window_Unonload in Common.vbs automatically
'========================================================================================================
Sub Form_QueryUnload(Cancel, UnloadMode)
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    '--------- Developer Coding Part (End) ------------------------------------------------------------
End Sub

'========================================================================================================
' Name : FncQuery
' Desc : This function is called from MainQuery in Common.vbs
'========================================================================================================
Function FncQuery()
    Dim IntRetCD 
    
    FncQuery = False															 '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

    ggoSpread.Source = Frm1.vspdData
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO,"x","x")					 '☜: Data is changed.  Do you want to display it? 
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    Call ggoOper.ClearField(Document, "2")										 '☜: Clear Contents  Field
    ggoSpread.Source = frm1.vspdData
	ggospread.ClearSpreadData		'Buffer Clear
    Call InitVariables																
    If Not chkField(Document, "1") Then									         '☜: This function check required field
       Exit Function
    End If

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 

  
    Call MakeKeyStream("X")
	'------ Developer Coding part (End )   -------------------------------------------------------------- 

    If DbQuery = False Then                                                      '☜: Query db data
       Exit Function
    End If

    Set gActiveElement = document.ActiveElement   
    FncQuery = True                                                              '☜: Processing is OK
End Function
	
'========================================================================================================
' Name : FncNew
' Desc : This function is called from MainNew in Common.vbs
'========================================================================================================
Function FncNew()
    Dim IntRetCD 
    
    FncNew = False																 '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
    Set gActiveElement = document.ActiveElement   
    FncNew = True																 '☜: Processing is OK
End Function
	
'========================================================================================================
' Name : FncDelete
' Desc : This function is called from MainDelete in Common.vbs
'========================================================================================================
Function FncDelete()
    Dim intRetCD
    
    FncDelete = False                                                            '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
    Set gActiveElement = document.ActiveElement       
    FncDelete = True                                                             '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncSave
' Desc : This function is called from MainSave in Common.vbs
'========================================================================================================
Function FncSave()
    Dim IntRetCD 
    
    FncSave = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
    
    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = False Then                                      '☜:match pointer
        IntRetCD = DisplayMsgBox("900001","x","x","x")                           '☜:There is no changed data. 
        Exit Function
    End If
    
    ggoSpread.Source = frm1.vspdData
    If Not ggoSpread.SSDefaultCheck Then                                         '☜: Check contents area
       Exit Function
    End If
    
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 

	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    If DbSave = False Then                                                       '☜: Query db data
       Exit Function
    End If

    Set gActiveElement = document.ActiveElement   
    FncSave = True                                                              '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncCopy
' Desc : This function is called from MainCopy in Common.vbs
' Keep : Make sure to clear primary key area
'========================================================================================================
Function FncCopy()
	Dim IntRetCD

    FncCopy = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

    If Frm1.vspdData.MaxRows < 1 Then
       Exit Function
    End If
    
    ggoSpread.Source = Frm1.vspdData
	With Frm1.VspdData
         .ReDraw = False
		 If .ActiveRow > 0 Then
            ggoSpread.CopyRow
			SetSpreadColor .ActiveRow, .ActiveRow
    
            .ReDraw = True
		    .Focus
		 End If
	End With
	
    Set gActiveElement = document.ActiveElement   
    FncCopy = True                                                               '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncCancel
' Desc : This function is called by MainCancel in Common.vbs
'========================================================================================================
Function FncCancel() 
    FncCancel = False                                                            '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

    ggoSpread.Source = Frm1.vspdData	
    ggoSpread.EditUndo  

    Set gActiveElement = document.ActiveElement   
    FncCancel = False                                                            '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncInsertRow
' Desc : This function is called by MainInsertRow in Common.vbs
'========================================================================================================
Function FncInsertRow(Byval pvRowCnt)
    Err.Clear                                                                    '☜: Clear err status
	Dim imRow
	FncInsertRow = False

	if IsNumeric(Trim(pvRowCnt)) Then
		imRow = CInt(pvRowCnt)
	else
		imRow = AskSpdSheetAddRowcount()

	If ImRow="" then
		Exit Function
	End If
	End If	
	With frm1
        .vspdData.ReDraw = False
        .vspdData.focus
        ggoSpread.Source = .vspdData
        ggoSpread.InsertRow .vspdData.ActiveRow, imRow
        SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1
        .vspdData.ReDraw = True
    End With
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Set gActiveElement = document.ActiveElement   
    FncInsertRow = True                                                          '☜: Processing is OK
End Function

'========================================================================================================
'   Event Name : vspdData_DblClick
'   Event Desc : 
'========================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)				'8) 컬럼 title 변경 
    Dim iColumnName
 	If Row<=0 then
		Exit Sub
	End If
	If Frm1.vspdData.MaxRows =0 then
		Exit Sub
	End if
End Sub

'========================================================================================================
' Name : FncDeleteRow
' Desc : This function is called by MainDeleteRow in Common.vbs
'========================================================================================================
Function FncDeleteRow()
    Dim lDelRows
    FncDeleteRow = False                                                         '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

    If Frm1.vspdData.MaxRows < 1 then
       Exit function
	End if	
	
    With Frm1.vspdData 
    	.focus
    	ggoSpread.Source = frm1.vspdData 
    	lDelRows = ggoSpread.DeleteRow
    End With
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Set gActiveElement = document.ActiveElement   
    FncDeleteRow = True                                                          '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncPrint
' Desc : This function is called by MainPrint in Common.vbs
'========================================================================================================
Function FncPrint()
    FncPrint = False
    Err.Clear
	Call Parent.FncPrint()
    FncPrint = True
End Function

'========================================================================================================
' Name : FncExcel
' Desc : This function is called by MainExcel in Common.vbs
'========================================================================================================
Function FncExcel() 
    FncExcel = False
    Err.Clear
	Call Parent.FncExport(parent.C_MULTI)
    FncExcel = True
End Function

'========================================================================================================
' Name : FncFind
' Desc : This function is called by MainFind in Common.vbs
'========================================================================================================
Function FncFind() 
    FncFind = False 
    Err.Clear
	Call Parent.FncFind(parent.C_MULTI, True)
    FncFind = True
End Function

'========================================================================================
' Function Name : FncSplitColumn
' Function Desc : 
'========================================================================================
Sub FncSplitColumn()

    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   

End Sub

'========================================================================================================
' Name : FncExit
' Desc : This function is called by MainExit in Common.vbs
'========================================================================================================
Function FncExit()
	Dim IntRetCD

    FncExit = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
	
    ggoSpread.Source = frm1.vspdData	
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO,"x","x")		             '⊙: Data is changed.  Do you want to exit? 
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    FncExit = True                                                               '☜: Processing is OK
End Function
'======================================================================================================
Function DbQuery()
	Dim strVal

    Err.Clear                                                                    '☜: Clear err status
    DbQuery = False                                                              '☜: Processing is NG

'    Call LayerShowHide(1)
	If	LayerShowHide(1) = False Then
		Exit Function
	End If
                                                            '☜: Show Processing Message
    
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    With Frm1
		strVal = BIZ_PGM_ID & "?txtMode="            & parent.UID_M0001						         
        strVal = strVal     & "&txtKeyStream="       & lgKeyStream               '☜: Query Key
        strVal = strVal     & "&txtMaxRows="         & .vspdData.MaxRows
        strVal = strVal     & "&lgStrPrevKeyIndex=" & lgStrPrevKeyIndex         '☜: Next key tag
    End With
		
    '--------- Developer Coding Part (End) ------------------------------------------------------------
    Call RunMyBizASP(MyBizASP, strVal)                                           '☜:  Run biz logic

    DbQuery = True                                                               '☜: Processing is OK

    Set gActiveElement = document.ActiveElement   
End Function

'========================================================================================================
Function DbSave()
    Dim pP21011
    Dim lRow        
    Dim lGrpCnt     
    Dim retVal      
    Dim boolCheck   
    Dim lStartRow   
    Dim lEndRow     
    Dim lRestGrpCnt 
	Dim strVal, strDel
	
    Err.Clear                                                                    '☜: Clear err status
    DbSave = False                                                               '☜: Processing is NG
'    Call LayerShowHide(1)                                                        '☜: Show Processing Message
	If	LayerShowHide(1) = False Then
		Exit Function
	End If

    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    strVal = ""
    strDel = ""
    lGrpCnt = 1



	With Frm1
    
       For lRow = 1 To .vspdData.MaxRows
    
           .vspdData.Row = lRow
           .vspdData.Col = 0
        
           Select Case .vspdData.Text
 
               Case ggoSpread.InsertFlag                                      '☜: Update추가 
													  strVal = strVal & "C" & parent.gColSep
													  strVal = strVal & lRow & parent.gColSep
                    .vspdData.Col = C_BIZ_CD		: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_DEPT_CD		: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_ORG_CHANGE_ID	: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_INTERNAL_CD	: strVal = strVal & Trim(.vspdData.Text) & parent.gRowSep  
                     lGrpCnt = lGrpCnt + 1
                     
               Case ggoSpread.UpdateFlag                                      '☜: Update
                                                  strVal = strVal & "U" & parent.gColSep
                                                  strVal = strVal & lRow & parent.gColSep
                    .vspdData.Col = C_BIZ_CD	: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_DEPT_CD	: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_ORG_CHANGE_ID	: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_INTERNAL_CD	: strVal = strVal & Trim(.vspdData.Text) & parent.gRowSep  
                     
                     lGrpCnt = lGrpCnt + 1
               Case ggoSpread.DeleteFlag                                      '☜: Delete
                                                  strDel = strDel & "D" & parent.gColSep
                                                  strDel = strDel & lRow & parent.gColSep
                    .vspdData.Col = C_BIZ_CD	: strDel = strDel & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_DEPT_CD	: strDel = strDel & Trim(.vspdData.Text) & parent.gRowSep
                    lGrpCnt = lGrpCnt + 1
           End Select
       Next
	
       .txtMode.value        = parent.UID_M0002
	   .txtMaxRows.value     = lGrpCnt-1	
	   .txtSpread.value      = strDel & strVal

	End With


	
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
	Call ExecMyBizASP(frm1, BIZ_PGM_ID)

    DbSave = True                                                               '☜: Processing is OK

    Set gActiveElement = document.ActiveElement   
End Function

'========================================================================================================
' Name : DbDelete
' Desc : This sub is called by FncDelete
'========================================================================================================
Function DbDelete()
    Err.Clear
    DbDelete = False
    DbDelete = True
End Function

'========================================================================================================
' Name : DbQueryOk
' Desc : Called by MB Area when query operation is successful
'========================================================================================================
Sub DbQueryOk()
	
    lgIntFlgMode = parent.OPMD_UMODE    
	'------ Developer Coding part (Start)  -------------------------------------------------------------- 
	Call SetToolbar("1100111100011111")                                              '☆: Developer must customize
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Call InitData()
	Call ggoOper.LockField(Document, "Q")
	frm1.vspdData.focus
	Set gActiveElement = document.ActiveElement
End Sub
	
'========================================================================================================
' Name : DbSaveOk
' Desc : Called by MB Area when save operation is successful
'========================================================================================================
Sub DbSaveOk()
    Call InitVariables															     '⊙: Initializes local global variables
	'------ Developer Coding part (Start)  -------------------------------------------------------------- 
    
    Call ggoOper.ClearField(Document, "2")										     '⊙: Clear Contents  Field
    ggoSpread.Source = frm1.vspdData
	ggospread.ClearSpreadData		'Buffer Clear
    Call MakeKeyStream("X")
	Call SetToolbar("1100111100011111")                                              '☆: Developer must customize
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    DBQuery()
    Set gActiveElement = document.ActiveElement   
End Sub
	
'========================================================================================================
' Function Name : DbDeleteOk
' Function Desc : Called by MB Area when delete operation is successful
'========================================================================================================
Sub DbDeleteOk()
	'------ Developer Coding part (Start)  -------------------------------------------------------------- 
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub


'========================================================================================================
Sub fpdtWk_yymm_DblClick(Button)
	If Button = 1 Then
		frm1.fpdtWk_yymm.Action = 7
		Call SetFocusToDocument("M")
		Frm1.fpdtWk_yymm.Focus
	End If
End Sub

'======================================================================================================
'   Event Name : vspdData_Change
'   Event Desc :
'========================================================================================================
Sub vspdData_Change(Col , Row )
    Dim iDx
    Dim IntRetCD,EFlag
	Dim iRow

    Dim BIZ_CD
    Dim DEPT_CD
	Dim biz_cd1
	Dim arrVal1, arrVal2, ii, jj
	
	
    EFlag = False
 
       
   	Frm1.vspdData.Row = Row
   	Frm1.vspdData.Col = Col

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    '=============================사업장 값 체크 시작 ==================================================
   Select Case Col
		Case C_BIZ_CD
			BIZ_CD = Trim(Frm1.vspdData.Text)

			If BIZ_CD <>"" Then
				IntRetCD = CommonQueryRs(" BIZ_AREA_CD "," A_MONTHLY_DEPT ","BIZ_AREA_CD= " & FilterVar(BIZ_CD, "''", "S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
				If IntRetCD = True Then
					Call DisplayMsgBox("124202","X","X","X")
					frm1.vspdData.text = ""
		    		frm1.vspdData.Col = C_BIZ_NM
		    		frm1.vspdData.text = ""
		    		frm1.vspdData.Col = C_BIZ_CD
		    		Frm1.vspdData.Action = 0
					Exit Sub
				End IF
			
				IntRetCD = CommonQueryRs(" BIZ_AREA_NM "," B_BIZ_AREA ","BIZ_AREA_CD= " & FilterVar(BIZ_CD, "''","S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
				If IntRetCD = False Then
				    Call DisplayMsgBox("127900","X","X","X")
				    Frm1.vspdData.Col = C_BIZ_CD
				    Frm1.vspdData.Text = ""
				    Frm1.vspdData.Col = C_BIZ_NM
				    Frm1.vspdData.Text = ""
				    Frm1.vspdData.Action = 0
				    Set gActiveElement = document.activeElement  
				    EFlag = True
				Else
				    Frm1.vspdData.Col = C_BIZ_NM
				    Frm1.vspdData.Text = Trim(Replace(lgF0,Chr(11),""))
				End If
			End If
		'=============================사업장 값 체크 끝 ==================================================


		'=============================부서코드 값 체크 시작 ==================================================
		  
		Case C_DEPT_CD
			DEPT_CD = Frm1.vspdData.Text

			Frm1.vspdData.Col = C_BIZ_CD
			Frm1.vspdData.Row = Row
			biz_cd1 = Trim(Frm1.vspdData.Text)


			If DEPT_CD <>"" Then
			IntRetCD = CommonQueryRs2by2("A.DEPT_NM, A.ORG_CHANGE_ID, A.INTERNAL_CD ","B_ACCT_DEPT A, B_COST_CENTER B","A.COST_CD = B.COST_CD AND A.ORG_CHANGE_ID = (select cur_org_change_id from b_company) AND B.BIZ_AREA_CD = " & FilterVar(biz_cd1, "''", "S" ) & " AND A.DEPT_CD=" & FilterVar(DEPT_CD, "''", "S"),lgF2By2)
				If IntRetCD = False Then
				    Call DisplayMsgBox("800098","X","X","X")
				    Frm1.vspdData.Col = C_DEPT_CD
				    Frm1.vspdData.Text = ""
				    Frm1.vspdData.Col = C_DEPT_NM
				    Frm1.vspdData.Text = ""
				    Frm1.vspdData.Col = C_ORG_CHANGE_ID
				    Frm1.vspdData.Text = ""
				    Frm1.vspdData.Col = C_INTERNAL_CD
				    Frm1.vspdData.Text = ""
				    Frm1.vspdData.Action = 0
				    Set gActiveElement = document.activeElement  
				    EFlag = True
				Else
				    	arrVal1 = Split(lgF2By2, Chr(11) & Chr(12))			
						jj = Ubound(arrVal1,1)
								
						For ii = 0 to jj - 1
							arrVal2 = Split(arrVal1(ii), chr(11))			
							Frm1.vspdData.Col = C_DEPT_NM
							 Frm1.vspdData.Text = Trim(arrVal2(1))
							 Frm1.vspdData.Col = C_ORG_CHANGE_ID
							 Frm1.vspdData.Text = Trim(arrVal2(2))
							 Frm1.vspdData.Col = C_INTERNAL_CD
							 Frm1.vspdData.Text = Trim(arrVal2(3))
						Next	
				End If
			End If
	end select 	
	'=============================부서코드 값 체크 끝 ==================================================

	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
		             
	Call CheckMinNumSpread(frm1.vspdData, Col, Row)

			
	ggoSpread.Source = frm1.vspdData
	ggoSpread.UpdateRow Row
		    
	'------ Developer Coding part (Start ) --------------------------------------------------------------
	'데이터 확인시 틀린데이터에 대해 undo 해준다.
	Frm1.vspdData.Row = Row
	Frm1.vspdData.Col = 0
		        
	If EFlag And Frm1.vspdData.Text <> ggoSpread.InsertFlag Then
	Call FncCancel()				
	End If
	'------ Developer Coding part (End   ) --------------------------------------------------------------
		    
    
End Sub

'========================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'========================================================================================================
Sub vspdData_Click(Col, Row)
	If lgIntFlgMode = Parent.OPMD_CMODE Then
		Call SetPopupMenuItemInf("1001111111")
	Else
		Call SetPopupMenuItemInf("1101111111")
	End If
	gMouseClickStatus = "SPC"
	Set gActiveSpdSheet = frm1.vspdData

 If Row = 0 Then
  ggoSpread.Source = frm1.vspdData
  If lgSortKey = 1 Then
   ggoSpread.SSSort
   lgSortKey = 2
  Else
   ggoSpread.SSSort ,lgSortKey
   lgSortKey = 1
  End If    
 End If
End Sub

'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData_ColWidthChange(ByVal Col1, ByVal Col2)		' 7) 컬럼 width 변경 이벤트 핸들러 
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(Col1,Col2)
End Sub

'========================================================================================================
'   Event Name : vspdData_MouseDown
'   Event Desc :
'========================================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If
End Sub    



'========================================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_TopLeftChange(OldLeft , OldTop , NewLeft , NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	           
    	If lgStrPrevKeyIndex <> "" Then                         
      	   DbQuery
    	End If
    End if
End Sub

'==========================================================================================
'   Event Name : vspdData_LeaveCell
'   Event Desc : Cell을 벗어날시 무조건발생 이벤트 
'==========================================================================================

Sub vspdData_ScriptLeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
    With frm1.vspdData 
  If Row >= NewRow Then
      Exit Sub
  End If
    End With
End Sub


Sub vspdData_ScriptDragDropBlock(Col,Row,Col2,Row2,NewCol,NewRow,NewCol2,NewRow2,Overwrite,Action,DataOnly,Cancel)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col,Row,Col2,Row2,NewCol,NewRow,NewCol2,NewRow2,Overwrite,Action,DataOnly,Cancel)
    Call GetSpreadColumnPos("A")
End Sub


</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->
</HEAD>

<!--
'########################################################################################################
'#						6. TAG 부																		#
'########################################################################################################
-->
<BODY SCROLL="NO" TABINDEX="-1">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR >
		<td nowrap <%=HEIGHT_TYPE_00%>></TD>
	</TR>
	<TR HEIGHT=23>
		<TD>
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<td nowrap WIDTH=10>&nbsp;</TD>
					<td nowrap CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td nowrap background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td nowrap background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>월차발생부서등록</font></td>
								<td nowrap background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<td nowrap WIDTH=*>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=*>
		<td nowrap WIDTH=100% CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<td nowrap WIDTH=100% HEIGHT=* valign=top>
						<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>
								<td nowrap HEIGHT="100%">
									<script language =javascript src='./js/a5966ma1_vaSpread_vspdData.js'></script>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR >
		<td nowrap WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO  noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" TAG="24" TABINDEX="-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMaxRows"      TAG="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtMode"         TAG="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtKeyStream"    TAG="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId"   TAG="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId"  TAG="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtFlgMode"      TAG="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtPrevNext"     TAG="24" TABINDEX="-1">

</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm" TABINDEX="-1"></IFRAME>
</DIV>
</BODY>
</HTML>

