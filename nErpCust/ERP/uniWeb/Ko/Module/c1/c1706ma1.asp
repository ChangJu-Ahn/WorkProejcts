
<%@ LANGUAGE="VBSCRIPT" %>
<!-- '======================================================================================================
'*  1. Module Name          : Cost Accounting
'*  2. Function Name        : Allocation Factor by Account
'*  3. Program ID           : c1706mb1
'*  4. Program Name         : 계정별 배부요소 정보 등록 
'*  5. Program Desc         : 계정코드별 배부요소 관련 정보 
'*  6. Modified date(First) : 2004/03/23
'*  7. Modified date(Last)  : 
'*  8. Modifier (First)     : cho ig sung
'*  9. Modifier (Last)      :
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================  -->


<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
 
<!--
########################################################################################################
#						   3.    External File Include Part
########################################################################################################-->

<!--
========================================================================================================
=                          3.1 Server Side Script
========================================================================================================-->
<!-- #Include file="../../inc/IncSvrCcm.inc"  -->

<!--
========================================================================================================
=                          3.2 Style Sheet
======================================================================================================== -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		

<!--
========================================================================================================
=                          3.3 Client Side Script
======================================================================================================== -->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliRdsQuery.vbs">	</SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>
Option Explicit 

'########################################################################################################
'#                       4.  Data Declaration Part
'########################################################################################################

'========================================================================================================
'=                       4.1 External ASP File
'========================================================================================================
Const BIZ_PGM_ID = "c1706mb1.asp"                             'Biz Logic ASP

'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Dim C_AcctCd  
Dim C_AcctPop   
Dim C_AcctNm  
Dim C_CostCd  	
Dim C_CostPop
Dim C_CostNm    
Dim C_DstbFctrCd  
Dim C_DstbFctrPop  
Dim C_DstbFctrNm  
Dim	C_AdjustFlag
Dim	C_AllcTarget

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$

'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	


'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim lgStrPrevKeyAcctCd
Dim lgStrPrevKeyCostCd

Dim lgQueryFlag	 
Dim IsOpenPop          


'========================================================================================================
'                        5.1 Common Method-1
'========================================================================================================
'========================================================================================================

'========================================================================================================
' Name : initSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================
Sub initSpreadPosVariables()  
	C_AcctCd		= 1
	C_AcctPop		= 2 
	C_AcctNm		= 3
	C_CostCd		= 4	
	C_CostPop		= 5
	C_CostNm		= 6  
	C_DstbFctrCd	= 7
	C_DstbFctrPop	= 8
	C_DstbFctrNm	= 9
	C_AdjustFlag	= 10
	C_AllcTarget	= 11
End Sub


'========================================================================================================
' Name : InitVariables()	
' Desc : Initialize value
'======================================================================================================== 
Sub InitVariables()

    lgIntFlgMode = Parent.OPMD_CMODE 
    lgBlnFlgChgValue = False   
    lgIntGrpCount = 0
     
    lgStrPrevKeyAcctCd = ""		
    lgStrPrevKeyCostCd = ""		
    lgLngCurRows = 0  
	lgSortKey = 1
	    
End Sub

'========================================================================================================
' Name : SetDefaultVal()	
' Desc : Set default value
'========================================================================================================
Sub SetDefaultVal()
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub

'========================================================================================================
' Name : LoadInfTB19029()	
' Desc : Set System Number format
' Para : 1. Currency
'        2. I(Input),Q(Query),P(Print),B(Bacth)
'        3. "*" is for Common module
'           "A" is for Accounting
'           "I" is for Inventory
'           ...
'========================================================================================================
 Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
	<%Call LoadInfTB19029A("I","*", "NOCOOKIE", "MA") %>
End Sub

'========================================================================================================
' Name : CookiePage()
' Desc : Write or Read cookie value 
'========================================================================================================
Sub CookiePage(Kubun)
   '------ Developer Coding part (Start ) -------------------------------------------------------------- 
   '------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub

'========================================================================================================
' Name : MakeKeyStream
' Desc : This method set focus to pos of err
'========================================================================================================
Sub MakeKeyStream(pOpt)
   
   '------ Developer Coding part (Start ) -------------------------------------------------------------- 
       
   '------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub        

'========================================================================================================
' Name : InitSpreadSheet
' Desc : This method initializes spread sheet column property
'========================================================================================================
Sub InitSpreadSheet()
	Call initSpreadPosVariables()    
	
	With frm1.vspdData
	
    .MaxCols = C_AllcTarget+1
	.Col = .MaxCols
    .ColHidden = True
    
    ggoSpread.Source = frm1.vspdData
    ggoSpread.Spreadinit "V20021122",,parent.gAllowDragDropSpread   

	Call ggoSpread.ClearSpreadData()    '☜: Clear spreadsheet data 
	
	.ReDraw = false

	Call GetSpreadColumnPos("A")
	
    ggoSpread.SSSetEdit		C_AcctCd		,"계정코드"			,15,,,20,2
	ggoSpread.SSSetButton	C_AcctPop    
    ggoSpread.SSSetEdit		C_AcctNm		,"계정명"			,22    
    ggoSpread.SSSetEdit		C_CostCd		,"코스트센타코드"	, 15,,,10,2
	ggoSpread.SSSetButton	C_CostPop    
    ggoSpread.SSSetEdit		C_CostNm		,"코스트센타명"		,22    
    ggoSpread.SSSetEdit		C_DstbFctrCd	,"배부요소코드"		, 15,,,2,2
	ggoSpread.SSSetButton	C_DstbFctrPop
    ggoSpread.SSSetEdit		C_DstbFctrNm	,"배부요소명"		,22
    ggoSpread.SSSetEdit		C_AdjustFlag	,""						,10
    ggoSpread.SSSetEdit		C_AllcTarget	,""						,10

	call ggoSpread.MakePairsColumn(C_AcctCd,C_AcctPop)
	call ggoSpread.MakePairsColumn(C_CostCd,C_CostPop)
	call ggoSpread.MakePairsColumn(C_DstbFctrCd,C_DstbFctrPop)	

	Call ggoSpread.SSSetColHidden(C_AdjustFlag,C_AdjustFlag,True)
	Call ggoSpread.SSSetColHidden(C_AdjustFlag,C_AllcTarget,True)

	
	.ReDraw = true
	
    Call SetSpreadLock 

    End With
    
End Sub

'======================================================================================================
' Name : SetSpreadLock
' Desc : This method set color and protect cell in spread sheet
'======================================================================================================
Sub SetSpreadLock()
    With frm1
    
    .vspdData.ReDraw = False
	ggoSpread.SpreadLock C_AcctCd		, -1, C_AcctCd    
	ggoSpread.SpreadLock C_AcctPop		, -1, C_AcctPop    
    ggoSpread.SpreadLock C_AcctNm		, -1, C_AcctNm
    ggoSpread.SpreadLock C_CostCd		, -1, C_CostCd    
    ggoSpread.SpreadLock C_CostPop		, -1, C_CostPop    
    ggoSpread.SpreadLock C_CostNm		, -1, C_CostNm    
    ggoSpread.SSSetRequired		C_DstbFctrCd	,-1	,C_DstbFctrCd
	ggoSpread.SpreadLock C_DstbFctrNm	, -1, C_DstbFctrNm
	ggoSpread.SSSetProtected	.vspdData.MaxCols	,-1,-1
    .vspdData.ReDraw = True

    End With
End Sub


'======================================================================================================
' Name : SetSpreadColor
' Desc : This method set color and protect cell by cell in spread sheet
'======================================================================================================
Sub SetSpreadColor(ByVal pvStartRow,ByVal pvEndRow)
    With frm1
    
    .vspdData.ReDraw = False
    'ggoSpread.SSSetRequired		C_AcctCd	,pvStartRow	,pvEndRow
    ggoSpread.SSSetProtected	C_AcctNm	,pvStartRow	,pvEndRow    
    'ggoSpread.SSSetRequired  	C_CostCd	,pvStartRow	,pvEndRow
    ggoSpread.SSSetProtected	C_CostNm	,pvStartRow	,pvEndRow    
    ggoSpread.SSSetRequired  	C_DstbFctrCd	,pvStartRow	,pvEndRow    
    ggoSpread.SSSetProtected	C_DstbFctrNm,pvStartRow	,pvEndRow
    .vspdData.ReDraw = True
    
    End With
End Sub


'======================================================================================================
' Name : SubSetErrPos
' Desc : This method set focus to position of error
'      : This method is called in MB area
'======================================================================================================
Sub SubSetErrPos(iPosArr)
    Dim iDx
    Dim iRow
    iPosArr = Split(iPosArr,Parent.gColSep)
    If IsNumeric(iPosArr(0)) Then
       iRow = CInt(iPosArr(0))
       For iDx = 1 To  frm1.vspdData.MaxCols - 1
           Frm1.vspdData.Col = iDx
           Frm1.vspdData.Row = iRow
           If Frm1.vspdData.ColHidden <> True And Frm1.vspdData.BackColor <> Parent.UC_PROTECTED Then
              Frm1.vspdData.Col = iDx
              Frm1.vspdData.Row = iRow
              Frm1.vspdData.Action = 0 ' go to 
              Exit For
           End If
           
       Next
          
    End If   
End Sub

'========================================================================================
' Function Name : GetSpreadColumnPos
' Description   : 
'========================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			C_AcctCd			= iCurColumnPos(1)
			C_AcctPop	        = iCurColumnPos(2)
			C_AcctNm			= iCurColumnPos(3)    
			C_CostCd	        = iCurColumnPos(4)
			C_CostPop	        = iCurColumnPos(5)
			C_CostNm			= iCurColumnPos(6)
			C_DstbFctrCd		= iCurColumnPos(7)
			C_DstbFctrPop	    = iCurColumnPos(8)
			C_DstbFctrNm		= iCurColumnPos(9)
			C_AdjustFlag		= iCurColumnPos(10)
			C_AllcTarget		= iCurColumnPos(11)
    End Select    
End Sub

	
'******************************************  2.4 POP-UP 처리함수  ****************************************
'	기능: 기준 POP-UP
'   설명: 기준 POP-UP에 관한 Open은 Include한다. 
'	      하나의 ASP에서 Popup이 중복되면 하나는 링크시켜 사용하고 나머지는 재정의하여 사용한다.
'********************************************************************************************************* 
 
'------------------------------------------  OpenCalType()  ----------------------------------------------
'	Name :
'	Description : 
'--------------------------------------------------------------------------------------------------------- 

Function OpenPopUp(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	Select Case iWhere

		Case 0
			arrParam(0) = "버전팝업"	
			arrParam(1) = "C_DSTB_RULE"				
			arrParam(2) = Trim(strCode)
			arrParam(3) = ""
			arrParam(4) = "work_step = " & FilterVar("06", "''", "S") & " "			
			arrParam(5) = "버전"			
				
			arrField(0) = "VER_CD"
			    
			arrHeader(0) = "버전"		

		Case 1, 3
			arrParam(0) = "계정팝업"
			arrParam(1) = "A_Acct"
			arrParam(2) = strCode
			arrParam(3) = ""	
			arrParam(4) = "temp_fg_3 in (" & FilterVar("M2", "''", "S") & " ," & FilterVar("M3", "''", "S") & " ," & FilterVar("M4", "''", "S") & " )"
			arrParam(5) = "계정"

			arrField(0) = "Acct_CD"
			arrField(1) = "Acct_NM"
    
			arrHeader(0) = "계정코드"
			arrHeader(1) = "계정명"	

		Case 2, 4
			arrParam(0) = "코스트센타팝업"
			arrParam(1) = "B_COST_CENTER a(nolock), B_MINOR b(NOLOCK), B_MINOR c(NOLOCK)"
			arrParam(2) = strCode
			arrParam(3) = ""		
			arrParam(4) = "a.DI_FG = b.MINOR_CD and b.MAJOR_CD = " & FilterVar("C0002", "''", "S") & "  and a.COST_TYPE = c.MINOR_CD and c.MAJOR_CD = " & FilterVar("C2203", "''", "S") & "  "	
			arrParam(5) = "코스트센타" 

			arrField(0) = "a.COST_CD"
			arrField(1) = "a.COST_NM"
			arrField(2) = "b.MINOR_NM"
			arrField(3) = "c.MINOR_NM"
    
			arrHeader(0) = "코스트센타코드"
			arrHeader(1) = "코스트센타명"
			arrHeader(2) = "직/간 구분"
			arrHeader(3) = "코스트센타 종류"

		Case 5
			arrParam(0) = "배부요소팝업"
			arrParam(1) = "C_DSTB_FCTR"		
			arrParam(2) = strCode
			arrParam(3) = ""		
			arrParam(4) = ""	
			arrParam(5) = "배부요소"			
				
			arrField(0) = "DSTB_FCTR_CD"	
			arrField(1) = "DSTB_FCTR_NM"
			    
			arrHeader(0) = "배부요소코드"	
			arrHeader(1) = "배부요소명"	

	End Select
    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetPopUp(arrRet, iWhere)
	End If	

End Function


Function SetPopUp(Byval arrRet, Byval iWhere)
	With frm1
		Select Case iWhere
			Case 0
				.txtVerCd.focus
				.txtVerCd.Value = arrRet(0)		
			Case 1
				.txtAcctCd.focus
				.txtAcctCd.Value    = arrRet(0)		
				.txtAcctNm.Value    = arrRet(1)		
			Case 2
				.txtCostCd.focus
				.txtCostCd.Value    = arrRet(0)		
				.txtCostNm.Value    = arrRet(1)		
			Case 3
				.vspdData.Col = C_AcctCd
				.vspdData.Text = arrRet(0)
				.vspdData.Col = C_AcctNm
				.vspdData.Text = arrRet(1)
				Call vspddata_Change(.vspddata.col, .vspddata.row)
			Case 4
				.vspdData.Col = C_CostCd
				.vspdData.Text = arrRet(0)
				.vspdData.Col = C_CostNm
				.vspdData.Text = arrRet(1)
				Call vspddata_Change(.vspddata.col, .vspddata.row)
			Case 5
				.vspdData.Col = C_DstbFctrCd
				.vspdData.Text = arrRet(0)
				.vspdData.Col = C_DstbFctrNm
				.vspdData.Text = arrRet(1)
				Call vspddata_Change(.vspddata.col, .vspddata.row)
			
		End Select

		lgBlnFlgChgValue = True
	End With
	
End Function

Sub Form_Load()
	
    Call LoadInfTB19029
    Call ggoOper.LockField(Document, "N") 
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)

    Call InitSpreadSheet
    Call InitVariables
    Call SetDefaultVal
    Call SetToolbar("110011010010111")
    frm1.txtAcctCd.focus
   	Set gActiveElement = document.activeElement			    
     
End Sub

Sub Form_QueryUnload(Cancel , UnloadMode )
End Sub

Sub vspdData_Change(ByVal Col, ByVal Row)
    Frm1.vspdData.Row = Row
    Frm1.vspdData.Col = Col

	Call CheckMinNumSpread(frm1.vspdData, Col, Row)
	
	ggoSpread.Source = frm1.vspdData
	ggoSpread.UpdateRow Row
	lgBlnFlgChgValue = True
	
End Sub

'==========================================================================================
'   Event Name : vspdData_Click
'   Event Desc : This event is spread sheet data changed
'==========================================================================================

sub vspdData_Click(ByVal Col, ByVal Row)
  
   	If lgIntFlgMode <> Parent.OPMD_UMODE Then
		Call SetPopupMenuItemInf("1001111111")
	Else 
		Call SetPopupMenuItemInf("1101111111")
	End If	
    
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
    End If
    
End Sub


'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

End Sub

'========================================================================================================
'   Event Name : vspdData_DblClick
'   Event Desc : 
'========================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)				
    

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
   
	
End Sub

'==========================================================================================
'   Event Desc : Spread Split 상태코드 
'==========================================================================================
Sub vspdData_MouseDown(Button, Shift, X, Y)
	If Button = 2 And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If
End Sub

Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	Dim intRetCD
	
	With frm1 
	
	    ggoSpread.Source = frm1.vspdData
   
		If Row = 0 Then Exit Sub
		
		Select Case Col
			Case C_AcctPop
				.vspdData.Col = C_AcctCd
				.vspdData.Row = Row
				
				Call OpenPopup(Trim(.vspdData.Text), 3)

			Case C_CostPop
				.vspdData.Col = C_CostCd
				.vspdData.Row = Row
				
				Call OpenPopup(Trim(.vspdData.Text), 4)

			Case C_DstbFctrPop        
				.vspdData.Col = C_DstbFctrCd
				.vspdData.Row = Row
				  
				Call OpenPopup(Trim(.vspdData.Text), 5)
		End Select

        Call SetActiveCell(.vspdData,Col-1,.vspdData.ActiveRow ,"M","X","X")   	
    End With
    
End Sub

'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
'    If Col <= C_AcctNm Or NewCol <= C_AcctNm Then
'        Cancel = True
'        Exit Sub
'    End If
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
End Sub

Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
	IF CheckRunningBizProcess = True Then
		Exit Sub
	END IF
	
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then
    	If lgStrPrevKeyAcctCd <> "" Then	
	      	DbQuery
    	End If

    End if
    
End Sub

Function FncQuery() 
    Dim IntRetCD
    
    FncQuery = False
    
    Err.Clear
    
    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"x","x")
    	If IntRetCD = vbNo Then
	      	Exit Function
    	End If
    End If
    
    Call ggoOper.ClearField(Document, "2")	
    ggoSpread.Source = frm1.vspdData
	ggoSpread.ClearSpreadData
    
'	Call InitSpreadSheet
    Call InitVariables
'    Call InitComboBox

    if frm1.txtAcctCd.value = "" then
		frm1.txtAcctNm.value = ""
    end if

    if frm1.txtCostCd.value = "" then
		frm1.txtCostNm.value = ""
    end if

    If Not chkField(Document, "1") Then	
       Exit Function
    End If
    
    Call SetToolbar("1100110100101111")
    
    IF DbQuery = false Then
		Exit Function
	END IF
       
    FncQuery = True	
    
End Function

Function FncNew() 
    Dim IntRetCD 
    
    FncNew = False 
    
    Err.Clear 

    If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
        IntRetCD = DisplayMsgBox("900015", Parent.VB_YES_NO,"x","x")
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    Call ggoOper.ClearField(Document, "1")
    Call ggoOper.ClearField(Document, "2")
    
    ggoSpread.Source = frm1.vspdData
	ggoSpread.ClearSpreadData
    
    Call ggoOper.LockField(Document, "N")
    Call InitVariables 
    Call SetDefaultVal
    
    FncNew = True 

End Function

Function FncSave() 
    Dim IntRetCD 
    
    FncSave = False 
    
    Err.Clear 

    ggoSpread.Source = frm1.vspddata
    If ggoSpread.SSCheckChange = False Then
        IntRetCD = DisplayMsgBox("900001","x","x","x")
        Exit Function
    End If
    
    If Not ggoSpread.SSDefaultCheck Then 
       Exit Function
    End If
    
    IF DbSave = False Then
		Exit Function
	END IF	

    FncSave = True
    
End Function

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'========================================================================================================
' Name : FncCopy
' Desc : This function is called from MainCopy in Common.vbs
' Keep : Make sure to clear primary key area
'========================================================================================================
Function FncCopy() 
	frm1.vspdData.ReDraw = False
	
    if frm1.vspdData.maxrows < 1 then exit function 
	   


    ggoSpread.Source = frm1.vspdData	
    ggoSpread.CopyRow
    SetSpreadColor frm1.vspdData.ActiveRow ,frm1.vspdData.ActiveRow
    
	frm1.vspdData.Col = C_CostCd
	frm1.vspdData.Text = ""
	    
	frm1.vspdData.Col = C_CostNm
	frm1.vspdData.Text = ""


	frm1.vspdData.ReDraw = True
End Function


Function FncCancel() 
 
	if frm1.vspdData.maxrows < 1 then exit function 

	ggoSpread.Source = frm1.vspdData	
	ggoSpread.EditUndo
    
End Function

'========================================================================================================
' Name : FncInsertRow
' Desc : This function is called by MainInsertRow in Common.vbs
'========================================================================================================
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
        .vspdData.ReDraw = False
        .vspdData.focus
        ggoSpread.Source = .vspdData
        ggoSpread.InsertRow  .vspdData.ActiveRow, imRow
        SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1
        .vspdData.ReDraw = True
    End With
	'------ Developer Coding part (Start ) --------------------------------------------------------------
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    
    Set gActiveElement = document.ActiveElement  
    If Err.number = 0 Then
       FncInsertRow = True                                                          '☜: Processing is OK
    End If   
End Function


Function FncDeleteRow() 
    Dim lDelRows
    Dim iDelRowCnt, i
    
    if frm1.vspdData.maxrows < 1 then exit function 
	   
    
    With frm1.vspdData 

    	.focus
    	ggoSpread.Source = frm1.vspdData 
    	lDelRows = ggoSpread.DeleteRow

    End With
End Function

Function FncPrint()
    Call parent.FncPrint() 
End Function

Function FncPrev() 
End Function

Function FncNext() 
End Function

Function FncExcel() 
    Call parent.FncExport(Parent.C_MULTI)
End Function

Function FncFind() 
    Call parent.FncFind(Parent.C_MULTI, False) 
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


'========================================================================================
' Function Name : PopSaveSpreadColumnInf
' Description   : 
'========================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

'========================================================================================
' Function Name : PopRestoreSpreadColumnInf
' Description   : 
'========================================================================================
Sub PopRestoreSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()      
'    Call InitComboBox_Two
	Call ggoSpread.ReOrderingSpreadData()
End Sub


Function FncExit()
Dim IntRetCD
	
	FncExit = False
	
    ggoSpread.Source = frm1.vspdData	
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO,"x","x")
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    FncExit = True
End Function

Function DbQuery() 
	Dim strVal

    DbQuery = False
    
    IF LayerShowHide(1) = False Then
		Exit Function
	END IF
    
    Err.Clear
    
    With frm1
		If lgIntFlgMode = Parent.OPMD_UMODE Then
			strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
			strVal = strVal & "&lgStrPrevKeyAcctCd=" & lgStrPrevKeyAcctCd
			strVal = strVal & "&lgStrPrevKeyCostCd=" & lgStrPrevKeyCostCd
			strVal = strVal & "&txtVerCd=" & Trim(.hVerCd.value)
			strVal = strVal & "&txtAcctCd=" & Trim(.hAcctCd.value)	
			strVal = strVal & "&txtCostCd=" & Trim(.hCostCd.value)						
			strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
		Else
			strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
			strVal = strVal & "&lgStrPrevKeyAcctCd=" & lgStrPrevKeyAcctCd
			strVal = strVal & "&lgStrPrevKeyCostCd=" & lgStrPrevKeyCostCd
			strVal = strVal & "&txtVerCd=" & Trim(.txtVerCd.value)
			strVal = strVal & "&txtAcctCd=" & Trim(.txtAcctCd.value)
			strVal = strVal & "&txtCostCd=" & Trim(.txtCostCd.value)			
			strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
		End If
    
		Call RunMyBizASP(MyBizASP, strVal)
        
    End With
    
    DbQuery = True

End Function

Function DbQueryOk()
	
    lgIntFlgMode = Parent.OPMD_UMODE	
    
    Call ggoOper.LockField(Document, "Q")	

	Call SetToolbar("110011110011111")
	Frm1.vspdData.Focus
    Set gActiveElement = document.ActiveElement   	
	
End Function

Function DbSave() 
    Dim lRow        
    Dim lGrpCnt     
	Dim strVal, strDel
    Dim iColSep 
    Dim iRowSep   
	
    DbSave = False 
    
    IF LayerShowHide(1) = False Then
		Exit function
	END IF
	
	iColSep = Parent.gColSep
	iRowSep = Parent.gRowSep	
 
	With frm1
		.txtMode.value = Parent.UID_M0002

		lGrpCnt = 1

		strVal = ""
		strDel = ""
    
		For lRow = 1 To .vspdData.MaxRows
			.vspdData.Row = lRow
			.vspdData.Col = 0
        
			Select Case .vspdData.Text

	            Case ggoSpread.InsertFlag		
					strVal = strVal & "C" & iColSep & lRow & iColSep

					strVal = strVal & Trim(.txtVerCd.Value) & iColSep

					.vspdData.Col = C_AcctCd
					strVal = strVal & Trim(.vspdData.Text) & iColSep
					.vspdData.Col = C_CostCd	
					strVal = strVal & Trim(.vspdData.Text) & iColSep
					.vspdData.Col = C_DstbFctrCd		
					strVal = strVal & Trim(.vspdData.Text) & iColSep
					.vspdData.Col = C_AdjustFlag
					strVal = strVal & Trim(.vspdData.Text) & iColSep
					.vspdData.Col = C_AllcTarget
					strVal = strVal & Trim(.vspdData.Text) & iRowSep
					lGrpCnt = lGrpCnt + 1

	            Case ggoSpread.UpdateFlag			
					strVal = strVal & "U" & iColSep & lRow & iColSep		

					strVal = strVal & Trim(.txtVerCd.Value) & iColSep

					.vspdData.Col = C_AcctCd
					strVal = strVal & Trim(.vspdData.Text) & iColSep
					.vspdData.Col = C_CostCd	
					strVal = strVal & Trim(.vspdData.Text) & iColSep
					.vspdData.Col = C_DstbFctrCd		
					strVal = strVal & Trim(.vspdData.Text) & iColSep
					.vspdData.Col = C_AdjustFlag
					strVal = strVal & Trim(.vspdData.Text) & iColSep
					.vspdData.Col = C_AllcTarget
					strVal = strVal & Trim(.vspdData.Text) & iRowSep
					lGrpCnt = lGrpCnt + 1

	            Case ggoSpread.DeleteFlag	
					strDel = strDel & "D" & iColSep & lRow & iColSep	

					strVal = strVal & Trim(.txtVerCd.Value) & iColSep

					.vspdData.Col = C_AcctCd
					strVal = strVal & Trim(.vspdData.Text) & iColSep
					.vspdData.Col = C_CostCd	
					strVal = strVal & Trim(.vspdData.Text) & iRowSep
					lGrpCnt = lGrpCnt + 1
                
	        End Select
                
		Next
	
		.txtMaxRows.value = lGrpCnt-1
		.txtSpread.value = strDel & strVal

		Call ExecMyBizASP(frm1, BIZ_PGM_ID)		
	
	End With
	
    DbSave = True    
    
End Function


Function DbSaveOk()		
   
	Call InitVariables
	frm1.vspdData.MaxRows = 0

	Call MainQuery()
		
End Function


Function DbDelete() 
End Function


'=======================================================================================================
'	Name : ExeReflect()
'	Description :
'=======================================================================================================
Function ExeReflect()
	Dim IntRetCD

    ExeReflect = False															'⊙: Processing is NG

    
    If Not chkField(Document, "1") Then                             '⊙: Check contents area
       Exit Function
    End If


	IntRetCD = DisplayMsgBox("900018",Parent.VB_YES_NO,"X","X")
	
	If IntRetCD = vbNo Then
		Exit Function
	End If

	IF LayerShowHide(1) = False Then
		Exit Function
	END IF
	    
    Err.Clear                                                               '☜: Protect system from crashing

	With frm1
		.txtMode.value = Parent.UID_M0003

		Call ExecMyBizASP(frm1, BIZ_PGM_ID)		
	
	End With
   
   
    ExeReflect = True         
End Function


Function ExeReflectOk()
Dim IntRetCD 

	window.status = "반영 작업 완료"

	IntRetCD =DisplayMsgBox("990000","X","X","X")

	MainQuery
			
End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->	
</HEAD>

<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR >
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>계정별배부요소등록</font></td>
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
									<TD CLASS="TD5" NOWRAP>버전</TD>
									<TD CLASS="TD6" NOWRAP><INPUT CLASS="clstxt" NAME="txtVerCd" MAXLENGTH=3 SIZE=10 ALT="버전" tag="12XXXU" ALT="버전"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnVerCd" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtVerCd.value,0)">
									</TD>
									<TD CLASS="TD5">계정</TD>
									<TD CLASS="TD6"><INPUT ClASS="clstxt" TYPE=TEXT NAME="txtAcctCd" SIZE=10 MAXLENGTH=20 tag="11XXXU" ALT="계정"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnAcctCd" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtAcctCd.value,1)">
										 <INPUT TYPE=TEXT ID="txtAcctNm" NAME="txtAcctNm" SIZE=30 tag="14X">
									</TD>
								</TR>    
								<TR>
									<TD CLASS="TD5">코스트센타</TD>
									<TD CLASS="TD6"><INPUT ClASS="clstxt" TYPE=TEXT NAME="txtCostCd" SIZE=10 MAXLENGTH=6 tag="11XXXU" ALT="배부요소"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnAcctCd" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtCostCd.value,2)">
										 <INPUT TYPE=TEXT ID="txtCostNm" NAME="txtCostNm" SIZE=30 tag="14X">
									</TD>
									<TD CLASS="TD5"></TD>
									<TD CLASS="TD6"></TD> 
								</TR>               
								           
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=100% valign=top>
						<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD HEIGHT="100%" NOWRAP>
								<script language =javascript src='./js/c1706ma1_vspdData_vspdData.js'></script>
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
<!--
	<TR HEIGHT=20>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_30%>>
			<TR>
				<TD WIDTH=10>&nbsp;</TD>
				<TD><BUTTON NAME="BtnExe" CLASS="CLSSBTN" ONCLICK="ExeReflect()" >자동생성</BUTTON>&nbsp;</TD>
				<TD WIDTH=*>&nbsp;</TD>
			</TR>
			</TABLE>
		</TD>
	</TR>	
-->
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BIZSIZE%>>
			<IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BIZSIZE%> FRAMEBORDER=0 SCROLLING=No noresize framespacing=0 TABINDEX= "-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" TABINDEX= "-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="hVerCd" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="hAcctCd" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="hCostCd" tag="24" TABINDEX= "-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

