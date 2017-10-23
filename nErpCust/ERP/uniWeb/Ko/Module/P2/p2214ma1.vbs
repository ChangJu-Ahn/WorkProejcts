
Const BIZ_PGM_QRY_ID = "p2214mb1.asp"

'==========================================================================================================
Dim C_MpsHistoryNo	
Dim C_Status	
Dim C_MpsExecDt	
Dim C_FixExecFromDt
Dim C_Dtf
Dim C_Ptf
Dim C_PlanDt
Dim C_InvFlag
Dim C_SSFlag
Dim C_MaxFlag
Dim C_MinFlag
Dim C_RoundFlag
Dim C_ConvertDt
Dim C_Approver
Dim C_StartOrderNo
Dim C_EndOrderNo
Dim C_MpsExecDtHD
Dim C_ConvertDtHD


Dim IsOpenPop          
'==========================================================================================================
'	Name : InitVariables()
'	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'========================================================================================================= 
Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE
    lgBlnFlgChgValue = False
    lgIntGrpCount = 0
        
    lgStrPrevKey = ""
    lgLngCurRows = 0
    lgSortKey    = 1
End Sub

'==========================================================================================================
' Name : initSpreadPosVariables()	
' Desc : Initialize Column Const value
'==========================================================================================================
Sub initSpreadPosVariables()  

	C_MpsHistoryNo	=  1
	C_Status        =  2
	C_MpsExecDt     =  3
	C_FixExecFromDt =  4
	C_Dtf			=  5
	C_Ptf			=  6
	C_PlanDt		=  7
	C_InvFlag		=  8
	C_SSFlag		=  9
	C_MaxFlag		= 10
	C_MinFlag		= 11
	C_RoundFlag     = 12
	C_ConvertDt		= 13
	C_Approver		= 14
	C_StartOrderNo	= 15
	C_EndOrderNo	= 16
	C_MpsExecDtHD	= 17
	C_ConvertDtHD	= 18
End Sub


'==========================================================================================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'==========================================================================================================
Sub InitSpreadSheet()
	
	Call initSpreadPosVariables()    
	
	With frm1.vspdData
	
		ggoSpread.Source = frm1.vspdData
		ggoSpread.Spreadinit "V20030224",, parent.gAllowDragDropSpread    
	
		.ReDraw = false
    
		.MaxCols = C_ConvertDtHD + 1
		.MaxRows = 0
    
		Call GetSpreadColumnPos("A")

		ggoSpread.SSSetEdit C_MpsHistoryNo,	"MPS이력 No.", 18
		ggoSpread.SSSetEdit C_Status,		"Status", 12
		ggoSpread.SSSetEdit C_MpsExecDt,	"MPS실행일", 20	
		ggoSpread.SSSetDate C_FixExecFromDt, "기준일자", 11, 2, gDateFormat
		ggoSpread.SSSetDate C_Dtf,			"DTF", 11, 2, gDateFormat
		ggoSpread.SSSetDate C_Ptf,			"PTF", 11, 2, gDateFormat
		ggoSpread.SSSetDate C_PlanDt,		"계획일자", 11, 2, gDateFormat
		ggoSpread.SSSetEdit C_InvFlag,		"가용재고", 8,2
		ggoSpread.SSSetEdit C_SSFlag,		"안전재고", 8,2
		ggoSpread.SSSetEdit C_Maxflag,		"최대Lot", 8,2
		ggoSpread.SSSetEdit C_MinFlag,		"최소Lot", 8,2
		ggoSpread.SSSetEdit C_RoundFlag,	"올림수", 8,2			
		ggoSpread.SSSetEdit C_ConvertDt,	"승인일", 20
		ggoSpread.SSSetEdit C_Approver,		"승인자", 14	
		ggoSpread.SSSetEdit C_StartOrderNo,	"시작MPS No.", 18
		ggoSpread.SSSetEdit C_EndOrderNo,	"종료MPS No.", 18
		ggoSpread.SSSetEdit C_MpsExecDtHD,	"MPS실행일", 20	
		ggoSpread.SSSetEdit C_ConvertDtHD,	"승인일", 20
	
		Call ggoSpread.SSSetColHidden(C_MpsExecDtHD, C_MpsExecDtHD, True)
		Call ggoSpread.SSSetColHidden(C_ConvertDtHD, C_ConvertDtHD, True)
		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
	
		ggoSpread.SSSetSplit2(1)
	
		.ReDraw = true

		Call SetSpreadLock 

    End With
    
End Sub

'==========================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'==========================================================================================================
Sub SetSpreadLock()
    ggoSpread.Source = frm1.vspdData
	ggoSpread.SpreadLockWithOddEvenRowColor()
End Sub

'==========================================================================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'==========================================================================================================
Sub SetSpreadColor(ByVal lRow)
    With frm1
    
    .vspdData.ReDraw = False

	ggoSpread.SSSetProtected C_MpsHistoryNo,	lRow, lRow
	ggoSpread.SSSetProtected C_Status,			lRow, lRow
	ggoSpread.SSSetProtected C_MpsExecDt,		lRow, lRow
	ggoSpread.SSSetProtected C_FixExecFromDt,	lRow, lRow
	ggoSpread.SSSetProtected C_Dtf,				lRow, lRow
	ggoSpread.SSSetProtected C_Ptf,				lRow, lRow
	ggoSpread.SSSetProtected C_PlanDt,			lRow, lRow
	ggoSpread.SSSetProtected C_InvFlag,			lRow, lRow
	ggoSpread.SSSetProtected C_SSFlag,			lRow, lRow
	ggoSpread.SSSetProtected C_MaxFlag,			lRow, lRow
	ggoSpread.SSSetProtected C_MinFlag,			lRow, lRow
	ggoSpread.SSSetProtected C_RoundFlag,		lRow, lRow
	ggoSpread.SSSetProtected C_ConvertDt,		lRow, lRow
	ggoSpread.SSSetProtected C_Approver,		lRow, lRow
	ggoSpread.SSSetProtected C_StartOrderNo,	lRow, lRow
	ggoSpread.SSSetProtected C_EndOrderNo,		lRow, lRow
		
    .vspdData.ReDraw = True
    
    End With
End Sub

'==========================================================================================================
' Function Name : GetSpreadColumnPos
' Description   : 
'==========================================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

			C_MpsHistoryNo	= iCurColumnPos(1)
			C_Status		= iCurColumnPos(2)
			C_MpsExecDt		= iCurColumnPos(3)    
			C_FixExecFromDt	= iCurColumnPos(4)
			C_Dtf			= iCurColumnPos(5)
			C_Ptf			= iCurColumnPos(6)
			C_PlanDt		= iCurColumnPos(7)
			C_InvFlag		= iCurColumnPos(8)
			C_SSFlag		= iCurColumnPos(9)    
			C_MaxFlag		= iCurColumnPos(10)
			C_MinFlag		= iCurColumnPos(11)
			C_RoundFlag		= iCurColumnPos(12)
			C_ConvertDt		= iCurColumnPos(13)    
			C_Approver		= iCurColumnPos(14)
			C_StartOrderNo	= iCurColumnPos(15)
			C_EndOrderNo	= iCurColumnPos(16)
			C_MpsExecDtHD	= iCurColumnPos(17)
			C_ConvertDtHD	= iCurColumnPos(18)
			
    End Select    

End Sub


'------------------------------------------  OpenPlant()  -------------------------------------------------
'	Name : OpenPlant()
'	Description : Condition Plant PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenPlant()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtPlantCd.classname) = "PROTECTED" Then Exit Function

	IsOpenPop = True

	arrParam(0) = "공장팝업"					' 팝업 명칭 
	arrParam(1) = "B_PLANT"							' TABLE 명칭 
	arrParam(2) = Trim(frm1.txtPlantCd.Value)		' Code Condition
	arrParam(3) = ""								' Name Cindition
	arrParam(4) = ""								' Where Condition
	arrParam(5) = "공장"						' TextBox 명칭 
	
    arrField(0) = "PLANT_CD"						' Field명(0)
    arrField(1) = "PLANT_NM"						' Field명(1)
    
    arrHeader(0) = "공장"						' Header명(0)
    arrHeader(1) = "공장명"						' Header명(1)
    
	arrRet = window.showModalDialog("../../comasp/commonpopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetPlant(arrRet)
	End If	
End Function

'------------------------------------------  OpenRunNo()  -------------------------------------------------
'	Name : OpenRunNo()
'	Description : Run No PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenRunNo()
	Dim arrRet
	Dim arrParam(1)
	Dim iCalledAspName, IntRetCD

	If IsOpenPop = True Or UCase(frm1.txtRunNo.className) = "PROTECTED" Then Exit Function

	If frm1.txtPlantCd.value = "" Then
		Call DisplayMsgBox("971012","X", "공장","X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement 
		Exit Function
	End If
	
	IsOpenPop = True

	arrParam(0) = frm1.txtPlantCd.value
	arrParam(1) = frm1.txtRunNo.value
	
	iCalledAspName = AskPRAspName("P2213PA1")
    
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "P2213PA1", "X")
		IsOpenPop = False
		Exit Function
	End If

	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent , arrParam(0), arrParam(1)), _
		"dialogWidth=480px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetRunNo(arrRet)
	End If	
End Function

'------------------------------------------  SetItemCd()  --------------------------------------------------
'	Name : SetItemCd()
'	Description : Item Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetRunNo(ByRef arrRet)
	frm1.txtRunNo.value= arrRet(0)		
	frm1.txtRunNo.focus
	Set gActiveElement = document.activeElement
End Function

'------------------------------------------  SetConPlant()  --------------------------------------------------
'	Name : SetConPlant()
'	Description : Condition Plant Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetPlant(ByRef arrRet)
	frm1.txtPlantCd.Value    = arrRet(0)
	frm1.txtPlantNm.Value    = arrRet(1)
	frm1.txtPlantCd.focus
	Set gActiveElement = document.activeElement
End Function


'=========================================================================================================
'   Event Name : vspdData_Click
'   Event Desc :
'=========================================================================================================
Sub vspdData_Click(ByVal Col , ByVal Row )
	
	Call SetPopupMenuItemInf("0000111111")
	
	gMouseClickStatus = "SPC"
	
	Set gActiveSpdSheet = frm1.vspdData

    If frm1.vspdData.MaxRows = 0 Then
       Exit Sub
   	End If

   	If Row <= 0 Then
   		If Col = C_MpsExecDt Then
   			Col = C_MpsExecDtHD
   		ElseIf Col = C_ConvertDt Then
   			Col = C_ConvertDtHD
   		End If
   		
        ggoSpread.Source = frm1.vspdData
        If lgSortKey = 1 Then
            ggoSpread.SSSort Col
            lgSortKey = 2
        Else
            ggoSpread.SSSort Col, lgSortKey
            lgSortKey = 1
        End If
        Exit Sub
    End If
   	    
    If Row <= 0 Then
       ggoSpread.Source = frm1.vspdData
       Exit Sub
    End If

End Sub

'=========================================================================================================
'   Event Name : vspdData_MouseDown(Button,Shift,x,y)
'   Event Desc :
'=========================================================================================================
Sub vspdData_MouseDown(Button, Shift, x, y)
		
	If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If

End Sub

'=========================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc : 
'=========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

End Sub

'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
End Sub

'==========================================================================================
'   Event Name : vspdData_Change
'   Event Desc :
'==========================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )

    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row

    lgBlnFlgChgValue = True

End Sub

'=========================================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'=========================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop)	Then
		If lgStrPrevKey <> "" Then
			Call DisableToolBar(parent.TBC_QUERY)  
			If DBQuery = False Then 
			   Call RestoreToolBar()
			   Exit Sub
			End If 
		End If
    End if
    
End Sub

'=========================================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'=========================================================================================================
Function FncQuery() 
    Dim IntRetCD 

    FncQuery = False

    Err.Clear

	If frm1.txtPlantCd.value = "" Then
		frm1.txtPlantNm.value = "" 
	End If
	
    Call ggoOper.ClearField(Document, "2")
    Call InitVariables
  
    If Not chkField(Document, "1") Then
       Exit Function
    End If

    If DbQuery = False Then
		Exit Function
	End If

    FncQuery = True
    
End Function

'=========================================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'=========================================================================================================
Function FncNew() 
    Dim IntRetCD 
    
    FncNew = False
    
    Err.Clear

    Call ggoOper.ClearField(Document, "A")
    Call ggoOper.LockField(Document, "N")
    Call InitVariables
        
    FncNew = True

End Function

'=========================================================================================================
' Function Name : FncDelete
' Function Desc : This function is related to Delete Button of Main ToolBar
'=========================================================================================================
Function FncDelete() 
    Dim IntRetCD 
    
    FncDelete = False    
    Err.Clear
    
    If lgIntFlgMode <> parent.OPMD_UMODE Then
        Call DisplayMsgBox("900002", "X", "X", "X")
        Exit Function
    End If

    If DbDelete = False Then
       Exit Function
    End If

    Call ggoOper.ClearField(Document, "A")
    
    FncDelete = True
    
End Function

'=========================================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'=========================================================================================================
Function FncCopy() 
	frm1.vspdData.ReDraw = False
	
    ggoSpread.Source = frm1.vspdData	
    ggoSpread.CopyRow
    SetSpreadColor frm1.vspdData.ActiveRow
    
	frm1.vspdData.ReDraw = True
End Function

'=========================================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'=========================================================================================================
Function FncCancel() 
    ggoSpread.Source = frm1.vspdData	
    ggoSpread.EditUndo
End Function

'=========================================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'=========================================================================================================
Function FncInsertRow() 
	With frm1
	
	.vspdData.focus
    ggoSpread.Source = .vspdData
    .vspdData.ReDraw = False
    ggoSpread.InsertRow
    .vspdData.ReDraw = True
    SetSpreadColor .vspdData.ActiveRow
    
    End With
End Function

'=========================================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'=========================================================================================================
Function FncDeleteRow() 
    Dim lDelRows
    Dim iDelRowCnt, i
    
    With frm1.vspdData 
    
    .focus
    ggoSpread.Source = frm1.vspdData 
    
	lDelRows = ggoSpread.DeleteRow
    
    End With
End Function

'=========================================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'=========================================================================================================
Function FncPrint()     
    Call parent.FncPrint()
End Function

'=========================================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'=========================================================================================================
Function FncExcel() 
    Call parent.FncExport(parent.C_MULTI)
End Function

'=========================================================================================================
' Function Name : FncFind
' Function Desc : 
'=========================================================================================================
Function FncFind() 
    Call parent.FncFind(parent.C_MULTI, False)
End Function

'=========================================================================================================
' Function Name : FncSplitColumn
' Function Desc : 
'=========================================================================================================
Function FncSplitColumn()
    
    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Function
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)
    
End Function

'=========================================================================================================
' Function Name : FncExit
' Function Desc : 
'=========================================================================================================
Function FncExit()
    FncExit = True
End Function

'=========================================================================================================
' Function Name : PopSaveSpreadColumnInf
' Description   : 
'=========================================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

'=========================================================================================================
' Function Name : PopRestoreSpreadColumnInf
' Description   : 
'=========================================================================================================
Sub PopRestoreSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()      
	Call ggoSpread.ReOrderingSpreadData()
	Call InitData()
End Sub

 
'=========================================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'=========================================================================================================
Function DbQuery() 
    DbQuery = False
    
    Call LayerShowHide(1)
    
    Err.Clear

	Dim strVal
    
    With frm1
    If lgIntFlgMode = parent.OPMD_UMODE Then
		strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&txtPlantCd=" & Trim(.hPlantCd.value)
		strVal = strVal & "&txtRunNo=" & Trim(.hRunNo.value)
		strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
    Else
		strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&txtPlantCd=" & Trim(.txtPlantCd.value)
		strVal = strVal & "&txtRunNo=" & Trim(.txtRunNo.value)
		strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
    
    End IF
    Call RunMyBizASP(MyBizASP, strVal)
        
    End With
    
    DbQuery = True

End Function

'=========================================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'=========================================================================================================
Function DbQueryOk()
	Call SetToolBar("11000000000111")

    lgIntFlgMode = parent.OPMD_UMODE
    Call ggoOper.LockField(Document, "Q")
    frm1.vspdData.focus
	lgBlnFlgChgValue = False
End Function
