
<%@ LANGUAGE="VBSCRIPT" %>
<!-- '======================================================================================================
'*  1. Module Name          : Cost
'*  2. Function Name        : 배부규칙등록 
'*  3. Program ID           : c1702ma1.asp
'*  4. Program Name         : 배부규칙등록 
'*  5. Program Desc         : 배부규칙등록 
'*  6. Modified date(First) : 2000/08/23
'*  7. Modified date(Last)  : 2002/06/18
'*  8. Modifier (First)     : 강창구 
'*  9. Modifier (Last)      : Cho Ig Sung / Park, Joon-won
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

<SCRIPT LANGUAGE=VBSCRIPT>
Option Explicit


'########################################################################################################
'#                       4.  Data Declaration Part
'########################################################################################################

'========================================================================================================
'=                       4.1 External ASP File
'========================================================================================================

Const BIZ_PGM_ID = "c1702mb1.asp"	                          'Biz Logic ASP
 
Dim C_WorkStep  	
Dim C_WorkStepPopup  															
Dim C_MinorNm  
Dim C_DstbFctrCd  
Dim C_DstbFctrPopup  
Dim C_DstbFctrNm  

'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	


'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim IsOpenPop          

'========================================================================================================
'                        5.1 Common Method-1
'========================================================================================================
'========================================================================================================

'========================================================================================================
Sub initSpreadPosVariables()  
	 C_WorkStep			= 1	
	 C_WorkStepPopup	= 2															
	 C_MinorNm			= 3
	 C_DstbFctrCd		= 4
	 C_DstbFctrPopup	= 5
	 C_DstbFctrNm		= 6
End Sub


'========================================================================================================
Sub InitVariables()

    lgIntFlgMode = Parent.OPMD_CMODE
    lgBlnFlgChgValue = False 
    lgIntGrpCount = 0     
    
    lgStrPrevKey = "" 
    lgLngCurRows = 0 
    lgSortKey = 1
End Sub


'========================================================================================================
Sub SetDefaultVal()
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub


'========================================================================================================
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% Call loadInfTB19029A("I", "*", "NOCOOKIE", "MA") %>
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
Sub InitSpreadSheet()
	Call initSpreadPosVariables() 
	
	With frm1.vspdData
	
    .MaxCols = C_DstbFctrNm+1	
    .Col = .MaxCols	
    .ColHidden = True
    
    ggoSpread.Source = frm1.vspdData
    ggoSpread.Spreadinit "V20021122" ,,parent.gAllowDragDropSpread   

	Call ggoSpread.ClearSpreadData()    '☜: Clear spreadsheet data 

	.ReDraw = false

    Call GetSpreadColumnPos("A")

    ggoSpread.SSSetEdit		C_WorkStep		,"작업단계코드"	,20,,,2,2
    ggoSpread.SSSetButton	C_WorkStepPopup
    ggoSpread.SSSetEdit		C_MinorNm		,"작업단계명"	,36
    ggoSpread.SSSetEdit		C_DstbFctrCd	,"배부요소코드"	,20,,,2,2
    ggoSpread.SSSetButton	C_DstbFctrPopup
    ggoSpread.SSSetEdit		C_DstbFctrNm	,"배부요소명"	,37

	call ggoSpread.MakePairsColumn(C_WorkStep,C_WorkStepPopup)
	call ggoSpread.MakePairsColumn(C_DstbFctrCd,C_DstbFctrPopup)

	
	.ReDraw = true

'   ggoSpread.SSSetSplit(C_MinorNm)	
    Call SetSpreadLock 
    
    End With
    
End Sub


'======================================================================================================
Sub SetSpreadLock()
    With frm1
    
    .vspdData.ReDraw = False
    ggoSpread.SpreadLock C_WorkStep			, -1, C_WorkStep
    ggoSpread.SpreadLock C_WorkStepPopup	, -1, C_WorkStepPopup
    ggoSpread.SpreadLock C_MinorNm			, -1, C_MinorNm
    ggoSpread.SSSetRequired	C_DstbFctrCd	, -1, -1
    ggoSpread.SpreadLock C_DstbFctrNm		, -1, C_DstbFctrNm
    ggoSpread.SSSetProtected	.vspdData.MaxCols,-1,-1
    .vspdData.ReDraw = True

    End With
End Sub

'======================================================================================================
Sub SetSpreadColor(ByVal pvStartRow,ByVal pvEndRow)
    With frm1
    
    .vspdData.ReDraw = False
                                         'Col				Row				Row2
    ggoSpread.SSSetRequired		C_WorkStep		,pvStartRow		,pvEndRow
    ggoSpread.SSSetProtected		C_MinorNm		,pvStartRow		,pvEndRow
    ggoSpread.SSSetRequired		C_DstbFctrCd	,pvStartRow		,pvEndRow
    ggoSpread.SSSetProtected		C_DstbFctrNm	,pvStartRow		,pvEndRow
    .vspdData.ReDraw = True
    
    End With
End Sub



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
			C_WorkStep				= iCurColumnPos(1)
			C_WorkStepPopup	        = iCurColumnPos(2)
			C_MinorNm				= iCurColumnPos(3)    
			C_DstbFctrCd	        = iCurColumnPos(4)
			C_DstbFctrPopup			= iCurColumnPos(5)
			C_DstbFctrNm			= iCurColumnPos(6)
    End Select    
End Sub


'******************************************  2.4 POP-UP 처리함수  ****************************************
'	기능: 기준 POP-UP
'   설명: 기준 POP-UP에 관한 Open은 Include한다. 
'	      하나의 ASP에서 Popup이 중복되면 하나는 링크시켜 사용하고 나머지는 재정의하여 사용한다.
'********************************************************************************************************* 
Function OpenWorkStep(Byval strCode)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "작업단계팝업"
	arrParam(1) = "B_MINOR a, B_CONFIGURATION b"	
	arrParam(2) = strCode
	arrParam(3) = ""	
	arrParam(4) = "a.MAJOR_CD = " & FilterVar("C2000", "''", "S") & "  AND a.MAJOR_CD = b.MAJOR_CD AND a.MINOR_CD = b.MINOR_CD and b.SEQ_NO = " & FilterVar("1", "''", "S") & "  AND b.REFERENCE = " & FilterVar("Y", "''", "S") & " "	
	arrParam(5) = "작업단계"			
	
    arrField(0) = "a.MINOR_CD"
    arrField(1) = "a.MINOR_NM"
    
    arrHeader(0) = "작업단계코드"
    arrHeader(1) = "작업단계명"
    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetWorkStep(arrRet)
	End If	

End Function


Function SetWorkStep(Byval arrRet)
	
	With frm1
	
    	.vspdData.Col = C_WorkStep
    	.vspdData.Text = arrRet(0)
    	.vspdData.Col = C_MinorNm
    	.vspdData.Text = arrRet(1)
            
    	Call vspdData_Change(.vspdData.Col, .vspdData.Row)	
	
	End With
	
End Function


Function OpenVerCd()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "버전팝업"	
	arrParam(1) = "C_DSTB_RULE"				
	arrParam(2) = Trim(frm1.txtVerCd.Value)
	arrParam(3) = ""
	arrParam(4) = ""			
	arrParam(5) = "버전"			
		
	arrField(0) = "VER_CD"
'	arrField(1) = ""	
	    
	arrHeader(0) = "버전"		
	
    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtVerCd.focus
		Exit Function
	Else
		Call SetVerCd(arrRet)
	End If	
	
End Function

Function SetVerCd(byval arrRet)
	frm1.txtVerCd.focus
	frm1.txtVerCd.Value = arrRet(0)		
End Function

Function OpenDstbFctrCd(Byval strCode)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

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
    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetDstbFctrCd(arrRet)
	End If	

End Function

Function SetDstbFctrCd(Byval arrRet)
	
	With frm1
	
		.vspdData.Col = C_DstbFctrCd
		.vspdData.Text = arrRet(0)
		.vspdData.Col = C_DstbFctrNm
		.vspdData.Text = arrRet(1)
		            
		Call vspdData_Change(.vspdData.Col, .vspdData.Row)	
	
	End With
	
End Function

'========================================================================================================
'========================================================================================================
'                        5.2 Common Method-2
'========================================================================================================

'========================================================================================================
Sub Form_Load()

    Call LoadInfTB19029     
    
    Call ggoOper.LockField(Document, "N")                      

    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)

    Call InitSpreadSheet 
    Call InitVariables
    
    Call SetDefaultVal

    Call SetToolbar("110011010010111")
    frm1.txtVerCd.focus 
   	Set gActiveElement = document.activeElement			    
    
End Sub


'========================================================================================================
Sub Form_QueryUnload(Cancel, UnloadMode)
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    '--------- Developer Coding Part (End) ------------------------------------------------------------
End Sub

Sub vspdData_Change(ByVal Col , ByVal Row )
    Frm1.vspdData.Row = Row
    Frm1.vspdData.Col = Col

	Call CheckMinNumSpread(frm1.vspdData, Col, Row)

    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row

End Sub

'==========================================================================================
'   Event Name : vspdData_Click
'   Event Desc : This event is spread sheet data changed
'==========================================================================================
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
	Dim strTemp

	With frm1 
	
		ggoSpread.Source = frm1.vspdData
		   
		If Row > 0 And Col = C_WorkStepPopUp Then
		    .vspdData.Col = Col
		    .vspdData.Row = Row
		        
			.vspdData.Col = C_WorkStep
		    Call OpenWorkStep(.vspdData.Text)
		        
		ElseIf Row > 0 And Col = C_DstbFctrPopUp Then
		    .vspdData.Col = Col
		    .vspdData.Row = Row
		        
			.vspdData.Col = C_DstbFctrCd
		    Call OpenDstbFctrCd(.vspdData.Text)

		End If
    Call SetActiveCell(.vspdData,Col-1,.vspdData.ActiveRow ,"M","X","X")   	
    End With
End Sub


'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
'    If Col <= C_MinorNm Or NewCol <= C_MinorNm Then
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
    
    IF frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then
    	If lgStrPrevKey <> "" Then 
    	      	DbQuery
    	End If

    End if
    
End Sub


'========================================================================================================
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
	
    Call InitVariables
    
    If Not chkField(Document, "1") Then
       Exit Function
    End If
    
    IF DbQuery = False Then
		Exit Function
	END IF	
           
    FncQuery = True															
    
End Function


'========================================================================================================
Function FncSave() 
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
    
    IF DbSave = False Then
		Exit Function
	END IF
    
    FncSave = True                                                          
    
End Function

'========================================================================================================
Function FncCopy() 
	frm1.vspdData.ReDraw = False
	
    if frm1.vspdData.maxrows < 1 then exit function 
	   

    ggoSpread.Source = frm1.vspdData	
    ggoSpread.CopyRow
    SetSpreadColor frm1.vspdData.ActiveRow ,frm1.vspdData.ActiveRow
    
    frm1.vspdData.Col = C_WorkStep
    frm1.vspdData.Text = ""
    
    frm1.vspdData.Col = C_MinorNm
    frm1.vspdData.Text = ""

	frm1.vspdData.ReDraw = True
End Function


'========================================================================================================
Function FncCancel() 
 
		if frm1.vspdData.maxrows < 1 then exit function 

    ggoSpread.Source = frm1.vspdData	
    ggoSpread.EditUndo                                                  
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


'========================================================================================================
Function FncDeleteRow() 
    Dim lDelRows

    if frm1.vspdData.maxrows < 1 then exit function 
	   

    With frm1.vspdData 
    	.focus
    	ggoSpread.Source = frm1.vspdData 
    	lDelRows = ggoSpread.DeleteRow
    End With
    
End Function


'========================================================================================================
Function FncPrint()
    Call parent.FncPrint()                                                   
End Function


'========================================================================================================
Function FncPrev() 
    On Error Resume Next
End Function


'========================================================================================================
Function FncNext() 
    On Error Resume Next 
End Function

Function FncExcel() 
    Call parent.FncExport(Parent.C_MULTI)
End Function

Function FncFind() 
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
'   Call InitComboBox
	Call ggoSpread.ReOrderingSpreadData()
'	Call InitData()
End Sub


'========================================================================================================
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

Function ExeCopy() 
	Dim IntRetCD
	Dim strVal

	On Error Resume Next
	Err.Clear 

    If lgIntFlgMode <> parent.OPMD_UMODE Then                                      'Check if there is retrived data
        IntRetCD = DisplayMsgBox("900002","X","X","X")                               
        Exit Function
    End If

	If Not chkField(Document, "2") Then
		Exit Function
	End If	

	IntRetCD = DisplayMsgBox("900018",Parent.VB_YES_NO,"X","X")

	If IntRetCD = vbNo Then
		Exit Function
	End If

	If LayerShowHide(1) = False Then
		Exit Function
	End If

	ExeCopy = False

	strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0003
	strVal = strVal & "&txtVerCd=" & Trim(frm1.hVerCd.value)
	strVal = strVal & "&txtNewVerCd=" & Trim(frm1.txtNewVerCd.value)

	Call RunMyBizASP(MyBizASP, strVal)

	ExeCopy = True
End Function

'========================================================================================================
'========================================================================================================
'                        5.3 Common Method-3
'========================================================================================================
'========================================================================================================

'========================================================================================================
Function DbQuery() 

    DbQuery = False
    
    Call LayerShowHide(1)
    
    Err.Clear 

	Dim strVal
    
    With frm1
    
    If lgIntFlgMode = Parent.OPMD_UMODE Then
		strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001		
		strVal = strVal & "&txtVerCd=" & .hVerCd.value				
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
    Else
		strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001		
		strVal = strVal & "&txtVerCd=" & .txtVerCd.value
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
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
    Dim iColSep 
    Dim iRowSep   
	
    DbSave = False                                                          
    
    Call LayerShowHide(1)
    
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

                .vspdData.Col = C_WorkStep
                strVal = strVal & Trim(.vspdData.Text) & iColSep
                
                .vspdData.Col = C_DstbFctrCd
                strVal = strVal & Trim(.vspdData.Text) & iRowSep
                
                lGrpCnt = lGrpCnt + 1
                
            Case ggoSpread.UpdateFlag
					
				strVal = strVal & "U" & iColSep & lRow & iColSep 

                .vspdData.Col = C_WorkStep
                strVal = strVal & Trim(.vspdData.Text) & iColSep
                
                .vspdData.Col = C_DstbFctrCd	
                strVal = strVal & Trim(.vspdData.Text) & iRowSep
                
                lGrpCnt = lGrpCnt + 1
                
            Case ggoSpread.DeleteFlag

				strDel = strDel & "D" & iColSep & lRow & iColSep

                .vspdData.Col = C_WorkStep	'1
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

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->	
</HEAD>

<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR >
		<TD <%=HEIGHT_TYPE_00%>></TD>
	<TR HEIGHT=23>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" NOWRAP><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>배부규칙등록</font></td>
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
									<TD CLASS="TD656" NOWRAP><INPUT CLASS="clstxt" NAME="txtVerCd" MAXLENGTH=3 SIZE=10 ALT="버전" tag="12XXXU" ALT="버전"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnVerCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenVerCd()">
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
								<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD"> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>				
				<TR >
					<TD WIDTH=100% HEIGHT=100%>
						<FIELDSET CLASS="CLSFLD">					
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS="TD5">신규버전</TD>
									<TD CLASS="TD6"><INPUT ClASS="clstxt" TYPE=TEXT NAME="txtNewVerCd" SIZE=10 MAXLENGTH=3 tag="22XXXU" ALT="신규버전">&nbsp;<BUTTON NAME="btnExe" CLASS="CLSSBTN" onclick="ExeCopy()" Flag=1>복사실행</BUTTON></TD>
									<TD CLASS="TD6"></TD>
									<TD CLASS="TD6"></TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=no noresize framespacing=0 TABINDEX= "-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" TABINDEX= "-1"></TEXTAREA><INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" TABINDEX= "-1"><%' 업무처리ASP로 넘기기 위한 변수를 담고 있는 Tag들 %>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="hVerCd" tag="24" TABINDEX= "-1">

<INPUT  NAME="txtMaxRows3" tag="24">

</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

