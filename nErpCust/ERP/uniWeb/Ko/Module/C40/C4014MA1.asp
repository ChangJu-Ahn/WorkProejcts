<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% session.CodePage=949 %>

<!-- '======================================================================================================
'*  1. Module Name          : Cost Accounting
'*  2. Function Name        : Cost Element By Account
'*  3. Program ID           : c1601mb1
'*  4. Program Name         : 계정별 원가요소 정보 등록 
'*  5. Program Desc         : 계정코드별 원가요소 관련 정보 
'*  6. Modified date(First) : 2000/11/08
'*  7. Modified date(Last)  : 2002/06/16
'*  8. Modifier (First)     : 강창구 
'*  9. Modifier (Last)      : cho Ig sung / Park, Joon-WonO
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
Const BIZ_PGM_ID = "C4014MB1.asp"                             'Biz Logic ASP

'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Dim C_AcctCd  
Dim C_AcctPop   
Dim C_AcctNm  
Dim C_DiFlag  	
Dim C_DiFlagNm    
Dim C_CostElmt  
Dim C_costElmtPop  
Dim C_CostElmtNm  
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$

'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	


'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim lgStrPrevKeyAcctCd
Dim lgStrPrevKeyDiFlag

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
	C_DiFlag		= 4	
	C_DiFlagNm		= 5  
	C_CostElmt		= 6
	C_costElmtPop	= 7
	C_CostElmtNm	= 8

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
    lgStrPrevKeyDiFlag = ""		
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
	
    .MaxCols = C_CostElmtNm+1
	.Col = .MaxCols
    .ColHidden = True
    
    ggoSpread.Source = frm1.vspdData
    ggoSpread.Spreadinit "V20021122",,parent.gAllowDragDropSpread   

	Call ggoSpread.ClearSpreadData()    '☜: Clear spreadsheet data 
	
	.ReDraw = false

	Call GetSpreadColumnPos("A")
	
    ggoSpread.SSSetEdit		C_AcctCd		,"계정코드"	,20,,,20,2
	ggoSpread.SSSetButton	C_AcctPop    
    ggoSpread.SSSetEdit		C_AcctNm		,"계정명"	,30    
    ggoSpread.SSSetCombo		C_DiFlag		,"", 8
    ggoSpread.SSSetCombo		C_DiFlagNm		,"직/간접"	,14
    ggoSpread.SSSetEdit		C_CostElmt		,"원가요소코드"	,20,,,20,2
	ggoSpread.SSSetButton	C_CostElmtPop
    ggoSpread.SSSetEdit		C_CostElmtNm	,"원가요소명"	,30

	call ggoSpread.MakePairsColumn(C_AcctCd,C_AcctPop)
	call ggoSpread.MakePairsColumn(C_CostElmt,C_CostElmtPop)	

	Call ggoSpread.SSSetColHidden(C_DiFlag,C_DiFlag,True)

	
	.ReDraw = true
	
'    ggoSpread.SSSetSplit(C_AcctNm)
    Call SetSpreadLock 
    Call initComboBox_Two
    
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
    ggoSpread.SpreadLock C_DiFlagNm		, -1, C_DiFlagNm    
	ggoSpread.SpreadLock C_CostElmtNm	, -1, C_CostElmtNm
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
    ggoSpread.SSSetRequired		C_AcctCd	,pvStartRow	,pvEndRow
    ggoSpread.SSSetProtected	C_AcctNm	,pvStartRow	,pvEndRow    
    ggoSpread.SSSetRequired  	C_DiFlagNm	,pvStartRow	,pvEndRow
    ggoSpread.SSSetRequired  	C_CostElmt	,pvStartRow	,pvEndRow    
    ggoSpread.SSSetProtected	C_CostElmtNm,pvStartRow	,pvEndRow
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
			C_DiFlag	        = iCurColumnPos(4)
			C_DiFlagNm			= iCurColumnPos(5)
			C_CostElmt			= iCurColumnPos(6)
			C_costElmtPop	    = iCurColumnPos(7)
			C_CostElmtNm		= iCurColumnPos(8)
    End Select    
End Sub


'******************************************  2.4 POP-UP 처리함수  ****************************************
'	기능: 기준 POP-UP
'   설명: 기준 POP-UP에 관한 Open은 Include한다. 
'	      하나의 ASP에서 Popup이 중복되면 하나는 링크시켜 사용하고 나머지는 재정의하여 사용한다.
'********************************************************************************************************* 
 
'------------------------------------------  OpenCalType()  ----------------------------------------------
'	Name :InitComboBox
'	Description : 
'--------------------------------------------------------------------------------------------------------- 
Sub InitComboBox_One
    
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
   
'   ggoSpread.SetCombo "10" & vbtab & "20" & vbtab & "30" & vbtab & "50" , C_ItemAcct
'    ggoSpread.SetCombo "제품" & vbtab & "반제품" & vbtab & "원자재"& vbtab & "상품", C_ItemAcctNm
'    ggoSpread.SetCombo "M" & vbtab & "O" & vbtab & "P", C_ProcurType
'    ggoSpread.SetCombo "사내가공품" & vbtab & "외주가공품" & vbtab & "구매품", C_ProcurTypeNm
   
    Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR ", "MAJOR_CD=" & FilterVar("C0002", "''", "S") & " "  ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

    Call SetCombo2(frm1.cboDiFlag ,lgF0  ,lgF1  ,Chr(11))                       

     
End Sub

Sub InitComboBox_Two
    
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
   
'   ggoSpread.SetCombo "10" & vbtab & "20" & vbtab & "30" & vbtab & "50" , C_ItemAcct
'    ggoSpread.SetCombo "제품" & vbtab & "반제품" & vbtab & "원자재"& vbtab & "상품", C_ItemAcctNm
'    ggoSpread.SetCombo "M" & vbtab & "O" & vbtab & "P", C_ProcurType
'    ggoSpread.SetCombo "사내가공품" & vbtab & "외주가공품" & vbtab & "구매품", C_ProcurTypeNm
   
    Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR ", "MAJOR_CD=" & FilterVar("C0002", "''", "S") & " "  ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

    ggoSpread.SetCombo Replace(lgF0,Chr(11),vbTab), C_DiFlag			'COLM_DATA_TYPE
    ggoSpread.SetCombo Replace(lgF1,Chr(11),vbTab), C_DiFlagNm
     
End Sub



Function OpenAcct()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "계정팝업"
	arrParam(1) = "A_ACCT"
	arrParam(2) = Trim(frm1.txtAcctCd.Value)
	arrParam(3) = ""
	arrParam(4) = "temp_fg_3 in (" & FilterVar("M2", "''", "S") & " ," & FilterVar("M3", "''", "S") & " ," & FilterVar("M4", "''", "S") & " )"
	arrParam(5) = "계정"
	
    arrField(0) = "ACCT_CD"
    arrField(1) = "ACCT_NM"
    
    arrHeader(0) = "계정코드"
    arrHeader(1) = "계정명"
    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtAcctCd.focus
		Exit Function
	Else
		Call SetAcct(arrRet)
	End If
		
End Function

Function OpenCostElmt()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "원가요소팝업"
	arrParam(1) = "C_COST_ELMT_S"
	arrParam(2) = Trim(frm1.txtCostElmtCd.Value)
	arrParam(3) = ""
	arrParam(4) = ""	
	arrParam(5) = "원가요소"
	
    arrField(0) = "cost_elmt_cd"
    arrField(1) = "cost_elmt_nm"
    
    arrHeader(0) = "원가요소코드"
    arrHeader(1) = "원가요소명"
    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtCostElmtCd.focus
		Exit Function
	Else
		Call SetCostElmt(arrRet)
	End If
		
End Function

Function OpenPopUp(Byval strCode, Byval strCode1, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	Select Case iWhere
		Case 0
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

		Case 1
			arrParam(0) = "원가요소팝업"
			arrParam(1) = "C_COST_ELMT_S"
			arrParam(2) = strCode
			arrParam(3) = ""		
			arrParam(4) = "DI_FLAG =  " & FilterVar(strCode1 , "''", "S") & ""	
			arrParam(5) = "원가요소" 

			arrField(0) = "cost_elmt_cd"
			arrField(1) = "cost_elmt_nm"
    
			arrHeader(0) = "원가요소코드"
			arrHeader(1) = "원가요소명"
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
				.vspdData.Col = C_AcctCd
				.vspdData.Text = arrRet(0)
				.vspdData.Col = C_AcctNm
				.vspdData.Text = arrRet(1)
				Call vspddata_Change(.vspddata.col, .vspddata.row)
			Case 1
				.vspdData.Col = C_CostElmt
				.vspdData.Text = arrRet(0)
				.vspdData.Col = C_CostElmtNm
				.vspdData.Text = arrRet(1)
				Call vspddata_Change(.vspddata.col, .vspddata.row)
			
		End Select

		lgBlnFlgChgValue = True
	End With
	
End Function

Function SetAcct(byval arrRet)
	frm1.txtAcctCd.focus
	frm1.txtAcctCd.Value    = arrRet(0)		
	frm1.txtAcctNm.Value    = arrRet(1)		
	lgBlnFlgChgValue = True
	
End Function

Function SetCostElmt(byval arrRet)
	frm1.txtCostElmtCd.focus
	frm1.txtCostElmtCd.Value    = arrRet(0)		
	frm1.txtCostElmtNm.Value  = arrRet(1)		
	lgBlnFlgChgValue = True
	
End Function

Sub Form_Load()
	
    Call LoadInfTB19029
    Call ggoOper.LockField(Document, "N") 
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)

    Call InitSpreadSheet
    Call InitVariables
	Call InitComboBox_One
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

Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
	Dim index
	
	With frm1.vspdData
		.Row = Row

		.Col = Col
		index = .Value
		
		.Col = C_DiFlag
		.Value = index
		
	End With
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
	Dim StrCode1
	Dim intRetCD
	
	With frm1 
	
	    ggoSpread.Source = frm1.vspdData
   
		If Row = 0 Then Exit Sub
		
		Select Case Col
			Case C_AcctPop
				.vspdData.Col = C_AcctCd
				.vspdData.Row = Row
				
				Call OpenPopup(.vspdData.Text, StrCode1, 0)

			Case C_costElmtPop        
				.vspdData.Col = C_DiFlag
				.vspdData.Row = Row
				  
				If .vspdData.Text = "" Then
					intRetCD =  DisplayMsgBox("235115","x","x","x")
					Exit Sub
				End If
				
     			StrCode1 = Trim(.vspdData.Text)
				
				.vspdData.Col = C_CostElmt
				.vspdData.Row = Row
				  
				Call OpenPopup(.vspdData.Text, StrCode1, 1)
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

    if frm1.txtCostElmtCd.value = "" then
		frm1.txtCostElmtNm.value = ""
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
    
    frm1.vspdData.Col = C_AcctCd
    frm1.vspdData.Text = ""
    
    frm1.vspdData.Col = C_AcctNm
    frm1.vspdData.Text = ""


	frm1.vspdData.ReDraw = True
End Function
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$

Function FncCancel() 
 
     if frm1.vspdData.maxrows < 1 then exit function 
	   

    ggoSpread.Source = frm1.vspdData	
    ggoSpread.EditUndo
    
    call InitData

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
    Call InitComboBox_Two
	Call ggoSpread.ReOrderingSpreadData()
	Call InitData()
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
			strVal = strVal & "&lgStrPrevKeyDiFlag=" & lgStrPrevKeyDiFlag
			strVal = strVal & "&txtAcctCd=" & Trim(.hAcctCd.value)	
			strVal = strVal & "&CboDiFlag=" & Trim(.hcboDiFlag.value)
			strVal = strVal & "&txtCostElmt=" & Trim(.txtCostElmtCd.value)						
			strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
		Else
			strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
			strVal = strVal & "&lgStrPrevKeyAcctCd=" & lgStrPrevKeyAcctCd
			strVal = strVal & "&lgStrPrevKeyDiFlag=" & lgStrPrevKeyDiFlag
			strVal = strVal & "&txtAcctCd=" & Trim(.txtAcctCd.value)
			strVal = strVal & "&CboDiFlag=" & Trim(.CboDiFlag.value)
			strVal = strVal & "&txtCostElmt=" & Trim(.txtCostElmtCd.value)			
			strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
		End If
    
		Call RunMyBizASP(MyBizASP, strVal)
        
    End With
    
    DbQuery = True

End Function

Function DbQueryOk()
	
    lgIntFlgMode = Parent.OPMD_UMODE	
    
    Call ggoOper.LockField(Document, "Q")	

   	InitData()	

	Call SetToolbar("110011110011111")
	Frm1.vspdData.Focus
    Set gActiveElement = document.ActiveElement   	
	
End Function

'========================================================================================================
' Name : InitData()
' Desc : Reset ComboBox
'========================================================================================================
 Sub InitData()
	Dim intRow
	Dim intIndex 
	
	With frm1.vspdData
		For intRow = 1 To .MaxRows
			
			.Row = intRow
			
			.Col = C_DiFlag
			intIndex = .value
			.col = C_DiFlagNm
			.value = intindex
					
		Next	
	End With
End Sub

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
					.vspdData.Col = C_AcctCd
					strVal = strVal & Trim(.vspdData.Text) & iColSep
					.vspdData.Col = C_DiFlag	
					strVal = strVal & Trim(.vspdData.Text) & iColSep
					.vspdData.Col = C_CostElmt		
					strVal = strVal & Trim(.vspdData.Text) & iRowSep
					lGrpCnt = lGrpCnt + 1

	            Case ggoSpread.UpdateFlag			
					strVal = strVal & "U" & iColSep & lRow & iColSep		
					.vspdData.Col = C_AcctCd		
					strVal = strVal & Trim(.vspdData.Text) & iColSep
					.vspdData.Col = C_DiFlag		
					strVal = strVal & Trim(.vspdData.Text) & iColSep
					.vspdData.Col = C_CostElmt	
					strVal = strVal & Trim(.vspdData.Text) & iRowSep
					lGrpCnt = lGrpCnt + 1

	            Case ggoSpread.DeleteFlag	
					strDel = strDel & "D" & iColSep & lRow & iColSep	
					.vspdData.Col = C_AcctCd	
					strDel = strDel & Trim(.vspdData.Text) & iColSep
					.vspdData.Col = C_DiFlag		
					strDel = strDel & Trim(.vspdData.Text) & iRowSep
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
'	Description : 평가금액 반영작업 
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>계정별원가요소등록</font></td>
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
									<TD CLASS="TD5">계정</TD>
									<TD CLASS="TD6"><INPUT ClASS="clstxt" TYPE=TEXT NAME="txtAcctCd" SIZE=10 MAXLENGTH=20 tag="11XXXU" ALT="계정"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnAcctCd" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenAcct()">
										 <INPUT TYPE=TEXT ID="txtAcctNm" NAME="txtAcctNm" SIZE=30 tag="14X">
									</TD>
									<TD CLASS="TD5">직/간접구분</TD>
									<TD CLASS="TD6"><SELECT NAME="cboDiFlag" tag="11" STYLE="WIDTH:82px:"><OPTION VALUE="" selected></OPTION></SELECT></TD>
								</TR>    
								<TR>
									<TD CLASS="TD5">원가요소</TD>
									<TD CLASS="TD6"><INPUT ClASS="clstxt" TYPE=TEXT NAME="txtCostElmtCd" SIZE=10 MAXLENGTH=20 tag="11XXXU" ALT="원가요소"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnAcctCd" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenCostElmt()">
										 <INPUT TYPE=TEXT ID="txtCostElmtNm" NAME="txtCostElmtNm" SIZE=30 tag="14X">
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
								<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> ID=vspdData NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD"> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
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
				<TD><BUTTON NAME="BtnExe" CLASS="CLSSBTN" ONCLICK="ExeReflect()" >자동생성</BUTTON>&nbsp;</TD>
				<TD WIDTH=*>&nbsp;</TD>
			</TR>
			</TABLE>
		</TD>
	</TR>	
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BIZSIZE%>>
			<IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BIZSIZE%> FRAMEBORDER=0 SCROLLING=No noresize framespacing=0 TABINDEX= "-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" TABINDEX= "-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="hAcctCd" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="hCboDiFlag" tag="24" TABINDEX= "-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

