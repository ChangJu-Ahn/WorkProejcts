
<%@ LANGUAGE="VBSCRIPT" %>
<!-- '======================================================================================================
'*  1. Module Name          : Cost
'*  2. Function Name        : 재료비 원가요소정보 등록 
'*  3. Program ID           : c1310ma1
'*  4. Program Name         : 품목계정별 재료비 원가요소 등록 
'*  5. Program Desc         : 품목계정별 조달구분별 재료비 원가요소 등록 
'*  6. Modified date(First) : 2000/08/23
'*  7. Modified date(Last)  : 2002/06/08
'*  8. Modifier (First)     : 강창구 
'*  9. Modifier (Last)      : Cho Ig sung / Park, Joon-Won
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'======================================================================================================= -->

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
Option Explicit                                                        '☜: indicates that All variables must be declared in advance
'########################################################################################################
'#                       4.  Data Declaration Part
'########################################################################################################

'========================================================================================================
'=                       4.1 External ASP File
'========================================================================================================


Const BIZ_PGM_ID = "c1310mb1.asp"	                                 'Biz Logic ASP 
'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================

Dim C_ItemAcctNm  
Dim C_ItemAcct  
Dim C_ProcurType  
Dim C_ProcurTypeNm  		
Dim C_CostElmtCd  
Dim C_CostElmtPop  
Dim C_CostElmtNm  	
Dim C_RelCostElmtCd  
Dim C_RelCostElmtPop  			
Dim C_RelCostElmtNm  


'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	


'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================

Dim lgStrPrevKeyItemAcct
Dim lgStrPrevKeyProcurType

Dim lgQueryFlag		
Dim IsOpenPop          

'========================================================================================================
'                        5.1 Common Method-1
'========================================================================================================
'========================================================================================================

'========================================================================================================
Sub initSpreadPosVariables()  
	 C_ItemAcctNm		= 1
	 C_ItemAcct			= 2
	 C_ProcurType		= 3
	 C_ProcurTypeNm		= 4		
	 C_CostElmtCd		= 5
	 C_CostElmtPop		= 6
	 C_CostElmtNm		= 7	
	 C_RelCostElmtCd	= 8
	 C_RelCostElmtPop	= 9			
	 C_RelCostElmtNm	= 10

End Sub


'========================================================================================================
sub InitVariables()

    lgIntFlgMode = Parent.OPMD_CMODE                                  ' ⊙: Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                                   ' ⊙: Indicates that no value changed
    lgIntGrpCount = 0                                            
    
    lgStrPrevKeyItemAcct = ""			                       '⊙: initializes Previous Key
    lgStrPrevKeyProcurType = ""	
    lgLngCurRows = 0  
	lgSortKey = 1
	    
End Sub


'========================================================================================================
	
Sub SetDefaultVal()
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub


'========================================================================================================
sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<%Call LoadInfTB19029A("I","P", "NOCOOKIE", "MA") %>
End Sub


'========================================================================================================
Sub CookiePage(Kubun)
   '------ Developer Coding part (Start ) -------------------------------------------------------------- 
   '------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub


'========================================================================================================
Sub MakeKeyStream(pOpt)
   
   '------ Developer Coding part (Start ) -------------------------------------------------------------- 
       
   '------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub        



'========================================================================================================
sub InitSpreadSheet()
	Call initSpreadPosVariables()    
	
	With frm1.vspdData
	
    .MaxCols = C_RelCostElmtNm+1		
	.Col = .MaxCols		
    .ColHidden = True
    
    ggoSpread.Source = frm1.vspdData
    ggoSpread.Spreadinit "V20021122",,parent.gAllowDragDropSpread    

	Call ggoSpread.ClearSpreadData()    '☜: Clear spreadsheet data 
	
	.ReDraw = false
	
	Call GetSpreadColumnPos("A")
    
    ggoSpread.SSSetCombo		C_ItemAcctNm	,"품목계정"		,17  
    ggoSpread.SSSetCombo		C_ItemAcct		,""	,5
    ggoSpread.SSSetCombo		C_ProcurType	,""	,8
    ggoSpread.SSSetCombo		C_ProcurTypeNm	,"조달구분"		,17    
    ggoSpread.SSSetEdit		C_CostElmtCd	,"원가요소코드"	,20,,,6,2
	ggoSpread.SSSetButton	C_CostElmtPop
    ggoSpread.SSSetEdit		C_CostElmtNm	,"원가요소명"	,20
    ggoSpread.SSSetEdit		C_RelCostElmtCd	,"관련원가요소코드"	,20,,,6,2
	ggoSpread.SSSetButton	C_RelCostElmtPop
    ggoSpread.SSSetEdit		C_RelCostElmtNm	,"관련원가요소명",20

	call ggoSpread.MakePairsColumn(C_CostElmtCd,C_CostElmtPop)
	call ggoSpread.MakePairsColumn(C_RelCostElmtCd,C_RelCostElmtPop)			
	

	Call ggoSpread.SSSetColHidden(C_ItemAcct ,C_ItemAcct	,True)
	Call ggoSpread.SSSetColHidden(C_ProcurType ,C_ProcurType	,True)

	.ReDraw = true
  
'    ggoSpread.SSSetSplit(C_ProcurTypeNm)	
    Call SetSpreadLock
    Call initComboBox 
    
    End With
    
End Sub


'======================================================================================================
sub SetSpreadLock()
    With frm1
    
    .vspdData.ReDraw = False
	ggoSpread.SpreadLock		C_ItemAcctNm	,-1	,C_ItemAcctNm   
	ggoSpread.SpreadLock		C_ProcurTypeNm	,-1	,C_ProcurTypeNm
    ggoSpread.SSSetRequired		C_CostElmtCd	,-1	,-1
	ggoSpread.SpreadLock		C_CostElmtNm	,-1	,C_CostElmtNm    
    ggoSpread.SpreadLock		C_RelCostElmtNm	,-1	,C_RelCostElmtNm
    ggoSpread.SSSetProtected	.vspdData.MaxCols,-1,-1
    .vspdData.ReDraw = True

    End With
End Sub



'======================================================================================================
sub SetSpreadColor(ByVal pvStartRow,ByVal pvEndRow)
    With frm1
    
    .vspdData.ReDraw = False
    ggoSpread.SSSetRequired		C_ItemAcctNm	,pvStartRow	,pvEndRow
    ggoSpread.SSSetRequired		C_ProcurTypeNm	,pvStartRow	,pvEndRow
    ggoSpread.SSSetRequired		C_CostElmtCd	,pvStartRow	,pvEndRow
    ggoSpread.SSSetProtected	C_CostElmtNm	,pvStartRow	,pvEndRow    
    ggoSpread.SSSetProtected	C_RelCostElmtNm	,pvStartRow	,pvEndRow
    .vspdData.ReDraw = True
    
    End With
End Sub

'******************************************  2.4 POP-UP 처리함수  ****************************************
'	기능: 기준 POP-UP
'   설명: 기준 POP-UP에 관한 Open은 Include한다. 
'	      하나의 ASP에서 Popup이 중복되면 하나는 링크시켜 사용하고 나머지는 재정의하여 사용한다.
'********************************************************************************************************* 
 

'--------------------------------------------------------------------------------------------------------- 
Sub InitComboBox
    
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
   
'   ggoSpread.SetCombo "10" & vbtab & "20" & vbtab & "30" & vbtab & "50" , C_ItemAcct
'    ggoSpread.SetCombo "제품" & vbtab & "반제품" & vbtab & "원자재"& vbtab & "상품", C_ItemAcctNm
'    ggoSpread.SetCombo "M" & vbtab & "O" & vbtab & "P", C_ProcurType
'    ggoSpread.SetCombo "사내가공품" & vbtab & "외주가공품" & vbtab & "구매품", C_ProcurTypeNm
   
    Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR ", _
                        " MAJOR_CD = " & FilterVar("P1001", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    ggoSpread.SetCombo Replace(lgF0,Chr(11),vbTab), C_ItemAcct			'COLM_DATA_TYPE
    ggoSpread.SetCombo Replace(lgF1,Chr(11),vbTab), C_ItemAcctNm
 
    
	Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR ", _
                        " MAJOR_CD = " & FilterVar("P1003", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
   	
    ggoSpread.SetCombo Replace(lgF0,Chr(11),vbTab), C_ProcurType			'COLM_DATA_TYPE
    ggoSpread.SetCombo Replace(lgF1,Chr(11),vbTab), C_ProcurTypeNm
    
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub


Function OpenPlant()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "공장팝업"	
	arrParam(1) = "B_PLANT"			
	arrParam(2) = Trim(frm1.txtPlantCd.Value)
	arrParam(3) = ""					
	arrParam(4) = ""						
	arrParam(5) = "공장"			
	
    arrField(0) = "PLANT_CD"	
    arrField(1) = "PLANT_NM"	
    
    arrHeader(0) = "공장코드"		
    arrHeader(1) = "공장명"	
    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtPlantCd.focus
		Exit Function
	Else
		Call SetPlant(arrRet)
	End If
		
End Function

Function OpenPopUp(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "원가요소팝업"	 
	arrParam(1) = "C_Cost_Elmt"		
	arrParam(2) =  strCode
	arrParam(3) = ""			
	arrParam(4) = ""				
	arrParam(5) = "원가요소"  

	arrField(0) = "Cost_Elmt_Cd"	
	arrField(1) = "Cost_Elmt_Nm"	
  
	arrHeader(0) = "원가요소코드"  
	arrHeader(1) = "원가요소명"
    
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
				.vspdData.Col = C_CostElmtCd
				.vspdData.Text = arrRet(0)
				.vspdData.Col = C_CostElmtNm
				.vspdData.Text = arrRet(1)
				Call vspddata_Change(.vspddata.col, .vspddata.row)
				.vspdData.Col = C_CostElmtCd
				.vspdData.Action = 0
				
			Case 1
				.vspdData.Col = C_RelCostElmtCd
				.vspdData.Text = arrRet(0)
				.vspdData.Col = C_RelCostElmtNm
				.vspdData.Text = arrRet(1)
				Call vspddata_Change(.vspddata.col, .vspddata.row)
			
		End Select

		'lgBlnFlgChgValue = True
	End With
	
End Function

Function SetPlant(byval arrRet)
	frm1.txtPlantCd.focus
	frm1.txtPlantCd.Value    = arrRet(0)		
	frm1.txtPlantNm.Value    = arrRet(1)		
	'lgBlnFlgChgValue = True
	
End Function


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
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			C_ItemAcctNm			= iCurColumnPos(1)
			C_ItemAcct		        = iCurColumnPos(2)
			C_ProcurType		    = iCurColumnPos(3)    
			C_ProcurTypeNm		    = iCurColumnPos(4)
			C_CostElmtCd		    = iCurColumnPos(5)
			C_CostElmtPop			= iCurColumnPos(6)
			C_CostElmtNm		    = iCurColumnPos(7)
			C_RelCostElmtCd			= iCurColumnPos(8)
			C_RelCostElmtPop		= iCurColumnPos(9)
			C_RelCostElmtNm		    = iCurColumnPos(10)
    End Select    
End Sub



'========================================================================================================
'========================================================================================================
'                        5.2 Common Method-2
'========================================================================================================
'========================================================================================================

'========================================================================================================

sub Form_Load()

    Call LoadInfTB19029 
    Call ggoOper.LockField(Document, "N")  
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)

    Call InitSpreadSheet  
    Call InitVariables
'	Call InitComboBox
    Call SetDefaultVal
    Call SetToolbar("110011010010111")		
    frm1.txtPlantCd.focus
   	Set gActiveElement = document.activeElement			
     
End Sub


'========================================================================================================
Sub Form_QueryUnload(Cancel, UnloadMode)
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    '--------- Developer Coding Part (End) ------------------------------------------------------------
End Sub



'========================================================================================================
'========================================================================================================
'                        5.5 Tag Event
'========================================================================================================

'========================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'========================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)

	If lgIntFlgMode <> Parent.OPMD_UMODE Then
		Call SetPopupMenuItemInf("1001111111")
	Else 
		Call SetPopupMenuItemInf("1101111111")
	End If	

    gMouseClickStatus = "SPC"	'Split 상태코드 

    Set gActiveSpdSheet = frm1.vspdData


	If frm1.vspdData.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
   	    
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

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    	frm1.vspdData.Row = Row
	'	frm1.vspdData.Col = C_MajorCd
	
	'------ Developer Coding part (End   ) --------------------------------------------------------------     
	
End Sub


'========================================================================================================
'   Event Name : vspdData_MouseDown
'   Event Desc :
'========================================================================================================
Sub vspdData_MouseDown(Button, Shift, X, Y)
	If Button = 2 And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If
End Sub


'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
'    If Col <= C_ProcurTypeNm Or NewCol <= C_ProcurTypeNm Then
'        Cancel = True
'        Exit Sub
'    End If
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
End Sub


'========================================================================================================
'   Event Name : vspdData_Change
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
sub vspdData_Change(ByVal Col, ByVal Row)
    Frm1.vspdData.Row = Row
    Frm1.vspdData.Col = Col

	Call CheckMinNumSpread(frm1.vspdData, Col, Row)


	ggoSpread.Source = frm1.vspdData
	ggoSpread.UpdateRow Row
	'lgBlnFlgChgValue = True
	
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


sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	
	With frm1
	
	    ggoSpread.Source = frm1.vspdData
   
		If Row = 0 Then Exit Sub
		
		Select Case Col
			Case C_CostElmtPop
				.vspdData.Col = Col
				.vspdData.Row = Row
				
				.vspdData.Col = C_CostElmtCd
				Call OpenPopup(.vspdData.Text, 0)

			Case C_RelCostElmtPop        
				.vspdData.Col = Col
				.vspdData.Row = Row
				  
				.vspdData.Col = C_RelCostElmtCd
				Call OpenPopup(.vspdData.Text, 1)
		End Select
        Call SetActiveCell(.vspdData,Col-1,.vspdData.ActiveRow ,"M","X","X")   	
    End With
    
End Sub



sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
	Dim intIndex
	
	With frm1.vspdData
		.Row = Row
		Select Case Col
		
			Case  C_ItemAcctNm
				.Col = Col
				intIndex = .Value
				.Col = C_ItemAcct
				.Value = intIndex
					
				
			Case  C_ItemAcct
				.Col = Col
				intIndex = .Value
				.Col = C_ItemAcctNm
				.Value = intIndex
			
			Case  C_ProcurType
				.Col = Col
				intIndex = .Value
				.Col = C_ProcurTypeNm
				.Value = intIndex	
			
			Case C_ProcurTypeNm
			    .Col = Col
				intIndex = .Value
				.Col = C_ProcurType
				.Value = intIndex	
				
			
		End Select
			
	End With
End Sub

sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )

	If CheckRunningBizProcess = True Then
		Exit Sub
	END If
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	
    	If lgStrPrevKeyItemAcct <> "" Then   
	      	DbQuery
    	End If

    End if
    
End Sub


'========================================================================================================
function FncQuery() 
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
    Call InitComboBox
   
    Call InitVariables 	

    if frm1.txtPlantCd.value = "" then
		frm1.txtPlantNm.value = ""
    end if
    
    If Not chkField(Document, "1") Then		
       Exit Function
    End If
    
    Call SetToolbar("1100110100101111")

    IF DbQuery = False Then
		Exit Function
	END IF			
       
    FncQuery = True				
    
End Function


'========================================================================================================
function FncNew() 
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


'========================================================================================================
function FncSave() 
    Dim IntRetCD 
    
    FncSave = False              
    
    Err.Clear       
    
    ggoSpread.Source = frm1.vspdData
    If lgBlnFlgChgValue = False And ggoSpread.SSCheckChange = False  Then
        IntRetCD = DisplayMsgBox("900001","x","x","x")   
        Exit Function
    End If

    If Not chkField(Document, "1") Then   
       Exit Function
    End If
    
    ggoSpread.Source = frm1.vspdData
    If Not ggoSpread.SSDefaultCheck Then      
       Exit Function
    End If
     
	If DbSave = False Then
		Exit Function
	End If

    FncSave = True    
    
End Function


'========================================================================================================
function FncCopy() 
	frm1.vspdData.ReDraw = False

    if frm1.vspdData.maxrows < 1 then exit function 
'	   

    ggoSpread.Source = frm1.vspdData	
    ggoSpread.CopyRow
    SetSpreadColor frm1.vspdData.ActiveRow, frm1.vspdData.ActiveRow

	With frm1.vspdData
		.col = C_ItemAcct
		.Text= ""
	
		.col = C_ItemAcctNm
		.Text= ""
	
		.col = C_ProcurType
		.Text= ""
					
		.col = C_ProcurTypeNm
		.Text= ""
	End With
    
	frm1.vspdData.ReDraw = True
End Function


function FncCancel() 

        if frm1.vspdData.maxrows < 1 then exit function 
        
    ggoSpread.Source = frm1.vspdData	
    ggoSpread.EditUndo

    call InitData
End Function



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


function FncDeleteRow() 
    Dim lDelRows
    
    if frm1.vspdData.maxrows < 1 then exit function 
	   

    With frm1.vspdData 
    	.focus
    	ggoSpread.Source = frm1.vspdData 
    	lDelRows = ggoSpread.DeleteRow
    End With
End Function


function FncPrint()
    Call parent.FncPrint()
End Function


function FncPrev() 
End Function

function FncNext() 
End Function

function FncExcel() 
    Call parent.FncExport(Parent.C_MULTI)
End Function

function FncFind() 
    Call parent.FncFind(Parent.C_MULTI, False) 
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


'========================================================================================================
function FncExit()
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

'========================================================================================================
'========================================================================================================
'                        5.3 Common Method-3
'========================================================================================================
'========================================================================================================

'========================================================================================================
function DbQuery() 
	Dim strVal

    DbQuery = False
    
    IF LayerShowHide(1) = False Then
		Exit Function
	END IF	

    Err.Clear 
    
    With frm1
		If lgIntFlgMode = Parent.OPMD_UMODE Then
			strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
			strVal = strVal & "&lgStrPrevKeyItemAcct=" & lgStrPrevKeyItemAcct
			strVal = strVal & "&lgStrPrevKeyProcurType=" & lgStrPrevKeyProcurType
			strVal = strVal & "&txtPlantCd=" & Trim(.hPlantCd.value)
			strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
		Else
			strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
			strVal = strVal & "&lgStrPrevKeyItemAcct=" & lgStrPrevKeyItemAcct
			strVal = strVal & "&lgStrPrevKeyProcurType=" & lgStrPrevKeyProcurType
			strVal = strVal & "&txtPlantCd=" & Trim(.txtPlantCd.value)
			strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
		End If
		
		Call RunMyBizASP(MyBizASP, strVal)
        
    End With
    
    DbQuery = True

End Function

'========================================================================================================
Function DbQueryOk()	

    lgIntFlgMode = Parent.OPMD_UMODE		
    
    Call ggoOper.LockField(Document, "Q")	

    Call SetToolbar("110011110011111")
    Frm1.vspdData.Focus
   	Call InitData()	
    Set gActiveElement = document.ActiveElement   

End Function


'========================================================================================================
sub InitData()
	Dim intRow
	Dim intIndex 
	
	With frm1.vspdData
		For intRow = 1 To .MaxRows
			
			.Row = intRow
			
			.Col = C_ItemAcct
			intIndex = .value
			.col = C_ItemAcctNm
			.value = intindex
			
			.Col = C_ProcurType
			intIndex = .value
			.col = C_ProcurTypeNm
			.value = intindex
					
		Next	
	End With
End Sub


'========================================================================================================
function DbSave() 
    Dim lRow        
    Dim lGrpCnt 
    Dim iColSep
    Dim iRowSep         
	Dim strVal, strDel
	
    DbSave = False  
    
    IF LayerShowHide(1) = False Then
		Exit Function
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
					.vspdData.Col = C_ItemAcct		
					strVal = strVal & Trim(.vspdData.Text) & iColSep
					.vspdData.Col = C_ProcurType	
					strVal = strVal & Trim(.vspdData.Text) & iColSep
					.vspdData.Col = C_CostElmtCd	
					strVal = strVal & Trim(.vspdData.Text) & iColSep
					.vspdData.Col = C_RelCostElmtCd	
					strVal = strVal & Trim(.vspdData.Text) & iRowSep
					lGrpCnt = lGrpCnt + 1

	            Case ggoSpread.UpdateFlag						
					strVal = strVal & "U" & iColSep & lRow & iColSep	
					.vspdData.Col = C_ItemAcct	
					strVal = strVal & Trim(.vspdData.Text) & iColSep
					.vspdData.Col = C_ProcurType	
					strVal = strVal & Trim(.vspdData.Text) & iColSep
					.vspdData.Col = C_CostElmtCd	
					strVal = strVal & Trim(.vspdData.Text) & iColSep
					.vspdData.Col = C_RelCostElmtCd
					strVal = strVal & Trim(.vspdData.Text) & iRowSep
					lGrpCnt = lGrpCnt + 1

	            Case ggoSpread.DeleteFlag	
					strDel = strDel & "D" & iColSep & lRow & iColSep	
					.vspdData.Col = C_ItemAcct		'2
					strDel = strDel & Trim(.vspdData.Text) & iColSep
					.vspdData.Col = C_ProcurType	'3
					strDel = strDel & Trim(.vspdData.Text) & iColSep
					.vspdData.Col = C_CostElmtCd	
					strDel = strDel & Trim(.vspdData.Text) & iColSep
					.vspdData.Col = C_RelCostElmtCd
					strDel = strDel & Trim(.vspdData.Text) & iRowSep
					lGrpCnt = lGrpCnt + 1
                
	        End Select
      
		Next

		.txtMaxRows.value = lGrpCnt-1
		.txtSpread.value =  strDel & strVal

		Call ExecMyBizASP(frm1, BIZ_PGM_ID)		
	
	End With
	
    DbSave = True  
    
End Function


'========================================================================================================
Function DbSaveOk()			
   
	Call InitVariables
	frm1.vspdData.MaxRows = 0
	Call MainQuery()
		
End Function


'========================================================================================================
Sub DbDeleteOk()
	'------ Developer Coding part (Start)  -------------------------------------------------------------- 

	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub



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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="left"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>재료비원가요소등록</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="9" height="23"></td>
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
									<TD CLASS="TD5">공장</TD>
									<TD CLASS="TD656"><INPUT ClASS="clstxt" TYPE=TEXT NAME="txtPlantCd" SIZE=10 MAXLENGTH=4 tag="12XXXU" ALT="공장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPlant()">
										 <INPUT TYPE=TEXT ID="txtPlantNm" NAME="txtPlantNm" SIZE=25 tag="14X">
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
					<TD WIDTH=100% HEIGHT=100% valign=top>
						<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD HEIGHT="100%" NOWRAP>
								<script language =javascript src='./js/c1310ma1_vspdData_vspdData.js'></script>
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
		<TD WIDTH=100% HEIGHT=<%=BizSize%>>
			<IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX= "-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" TABINDEX= "-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="hPlantCd" tag="24" TABINDEX= "-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

