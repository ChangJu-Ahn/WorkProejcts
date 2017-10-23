
Const BIZ_PGM_ID = "p2349mb1.asp"

'===========================================================================================================
Dim C_ItemCode
Dim C_ItemName  
Dim C_Spec 	     	
Dim C_TrackingNo
Dim C_ReqDate 	
Dim C_ReqQty 	
Dim C_IssuedQty 
Dim C_BasicUnit 
Dim C_RelatedType
Dim C_RelatedNo
Dim C_ItemGroupCd
Dim C_ItemGroupNm


Dim lgStrPrevKey1
Dim lgStrPrevKey2

Dim IsOpenPop          

'========================================================================================================
' Name : initSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================
Sub initSpreadPosVariables()  
    C_ItemCode 		= 1
    C_ItemName 		= 2
    C_Spec 			= 3
    C_TrackingNo 	= 4
    C_ReqDate 		= 5
    C_ReqQty 	    = 6
    C_IssuedQty		= 7
    C_BasicUnit		= 8
    C_RelatedType	= 9
    C_RelatedNo		= 10
    C_ItemGroupCd	= 11
    C_ItemGroupNm	= 12
End Sub

'===========================================================================================================
'	Name : InitVariables()
'	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'=========================================================================================================
Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE
    lgBlnFlgChgValue = False
    lgIntGrpCount = 0

    lgStrPrevKey1 = ""
    lgStrPrevKey2 = ""
    lgLngCurRows = 0
    lgSortKey    = 1
End Sub

'===========================================================================================================
'	Name : SetDefaultVal()
'	Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
'=========================================================================================================
Sub SetDefaultVal()
	frm1.txtFromReqrdDt.text  = StartDate
	frm1.txtToReqrdDt.text	  = LastDate
End Sub


'===========================================================================================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'===========================================================================================================
Sub InitSpreadSheet()
	Call initSpreadPosVariables()    
	
    With frm1.vspdData
    
    ggoSpread.Source = frm1.vspdData
    ggoSpread.Spreadinit "V20030114",,parent.gAllowDragDropSpread    
    
    .Redraw = False
    
	.MaxCols = C_ItemGroupNm +1 
	.MaxRows = 0
    
	Call AppendNumberPlace("6", "6", "0")
	
	Call GetSpreadColumnPos("A")

	ggoSpread.SSSetEdit 	C_ItemCode,		"품목", 18
	ggoSpread.SSSetEdit		C_ItemName,		"품목명", 25
	ggoSpread.SSSetEdit		C_Spec,			"규  격", 25
	ggoSpread.SSSetEdit		C_TrackingNo,	"Tracking No.", 25
	ggoSpread.SSSetEdit		C_BasicUnit,	"단위", 7        
	ggoSpread.SSSetDate		C_ReqDate,		"요구일", 11, 2, parent.gDateFormat	                          
	ggoSpread.SSSetFloat	C_ReqQty,		"필요량", 15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
	ggoSpread.SSSetFloat	C_IssuedQty,	"출고량", 15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
	ggoSpread.SSSetEdit		C_RelatedType,	"생성구분", 10
	ggoSpread.SSSetEdit		C_RelatedNo,	"관련번호", 18
	ggoSpread.SSSetEdit 	C_ItemGroupCd,	"품목그룹",		15
	ggoSpread.SSSetEdit		C_ItemGroupNm,	"품목그룹명",	30
	
	Call ggoSpread.SSSetColHidden(.MaxCols,.MaxCols,True)
	ggoSpread.SSSetSplit2(1)
	
	.ReDraw = true
	
	Call SetSpreadLock 
    
    End With
    
End Sub

'===========================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'===========================================================================================================
Sub SetSpreadLock()
    ggoSpread.Source = frm1.vspdData
	ggoSpread.SpreadLockWithOddEvenRowColor() 
End Sub

'===========================================================================================================
'	Name : InitComboBox()
'	Description : Combo Display
'===========================================================================================================
Sub InitComboBox()
    Call SetCombo(frm1.cboRelatedType, "SO", "수주")
    Call SetCombo(frm1.cboRelatedType, "MP", "MPS")
    Call SetCombo(frm1.cboRelatedType, "OD", "예약")
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

			C_ItemCode		= iCurColumnPos(1)
			C_ItemName		= iCurColumnPos(2)
			C_Spec			= iCurColumnPos(3)
			C_TrackingNo	= iCurColumnPos(4)    
			C_ReqDate		= iCurColumnPos(5)
			C_ReqQty		= iCurColumnPos(6)
			C_IssuedQty		= iCurColumnPos(7)
			C_BasicUnit		= iCurColumnPos(8)
			C_RelatedType	= iCurColumnPos(9)
			C_RelatedNo		= iCurColumnPos(10)
			C_ItemGroupCd	= iCurColumnPos(11)
			C_ItemGroupNm	= iCurColumnPos(12)
    End Select    

End Sub

'------------------------------------------  OpenItemInfo()  -------------------------------------------------
'	Name : OpenItemInfo()
'	Description : Item By Plant PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenItemInfo(Byval strCode)
	Dim arrRet
	Dim arrParam(5), arrField(6)
	Dim iCalledAspName, IntRetCD

	If IsOpenPop = True Then Exit Function
	
	If frm1.txtPlantCd.value = "" Then
		Call DisplayMsgBox("971012","X", "공장","X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement 
		Exit Function
	End If
	
	IsOpenPop = True
	
	arrParam(0) = Trim(frm1.txtPlantCd.value)	' Plant Code
	arrParam(1) = strCode						' Item Code
	arrParam(2) = ""							' Combo Set Data:"1020!MP" -- 품목계정 구분자 조달구분 
	arrParam(3) = ""							' Default Value
	
	arrField(0) = 1								'"ITEM_CD"			
	arrField(1) = 2								'"ITEM_NM"			
    
	iCalledAspName = AskPRAspName("B1B11PA3")
    
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "B1B11PA3", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.parent, arrParam, arrField), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetItemInfo(arrRet)
	End If	

End Function

'------------------------------------------  OpenPlant()  -------------------------------------------------
'	Name : OpenPlant()
'	Description : Plant PopUp
'---------------------------------------------------------------------------------------------------------
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

    arrHeader(0) = "공장"		
    arrHeader(1) = "공장명"		

    arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
 	 "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

    IsOpenPop = False

    If arrRet(0) = "" Then
		Exit Function
    Else
		Call SetPlant(arrRet)
    End If	

End Function

'--------------------------------------  OpenTrackingInfo()  ------------------------------------------
'	Name : OpenTrackingInfo()
'	Description : OpenTracking Info PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenTrackingInfo()
	Dim iCalledAspName, IntRetCD
	
	If frm1.txtPlantCd.value= "" Then
		Call DisplayMsgBox("971012","X", "공장","X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement
		IsOpenPop = False 
		Exit Function
	End If

	Dim arrRet
	Dim arrParam(4)

	If IsOpenPop = True Or UCase(frm1.txtTrackingNo.className) = "PROTECTED" Then Exit Function

	IsOpenPop = True

	arrParam(0) = Trim(frm1.txtPlantCd.value)
	arrParam(1) = Trim(frm1.txtTrackingNo.value)
	arrParam(2) = Trim(frm1.txtItemCd.value)
'	arrParam(3) = UniConvYYYYMMDDToDate(parent.gDateFormat, "1900", "01", "01")'frm1.txtPlanStartDt.Text
'	arrParam(4) = UniConvYYYYMMDDToDate(parent.gDateFormat, "2999", "12", "31")'frm1.txtPlanEndDt.Text
	
	iCalledAspName = AskPRAspName("P4600PA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "P4600PA1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.parent, arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetTrackingNo(arrRet)
	End If
	
End Function
'===========================================================================================================
Function OpenItemGroup()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtItemGroupCd.className) = UCase(parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "품목그룹팝업"
	arrParam(1) = "B_ITEM_GROUP"
	arrParam(2) = Trim(UCase(frm1.txtItemGroupCd.Value))
	arrParam(3) = ""
	arrParam(4) = "DEL_FLG = " & FilterVar("N", "''", "S") & " "
	arrParam(5) = "품목그룹"
	 
	arrField(0) = "ITEM_GROUP_CD"
	arrField(1) = "ITEM_GROUP_NM"
	    
	arrHeader(0) = "품목그룹"
	arrHeader(1) = "품목그룹명"
	    
	arrRet = window.showModalDialog("../../ComAsp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
	"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	 
	IsOpenPop = False
	 
	If arrRet(0) <> "" Then
		Call SetItemGroup(arrRet)
	End If 
	
	Call SetFocusToDocument("M")
	frm1.txtItemGroupCd.focus
 
End Function

'------------------------------------------  SetItemInfo()  --------------------------------------------------
'	Name : SetItemInfo()
'	Description : Item Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetItemInfo(ByRef arrRet)

    With frm1
		.txtItemCd.value = arrRet(0)
		.txtItemnm.value = arrRet(1)
		.txtItemCd.focus
		Set gActiveElement = document.activeElement	
    End With
    
End Function

'------------------------------------------  SetPlant()  --------------------------------------------------
'	Name : SetPlant()
'	Description : Plant Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetPlant(ByRef arrRet)
    frm1.txtPlantCd.Value = arrRet(0)
    frm1.txtPlantNm.value = arrRet(1)
    frm1.txtPlantCd.focus
	Set gActiveElement = document.activeElement	
End Function

'------------------------------------------  SetTrackingNo()  -----------------------------------------
'	Name : SetTrackingNo()
'	Description : Tracking No. Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetTrackingNo(ByRef arrRet)
	frm1.txtTrackingNo.Value = arrRet(0)
	frm1.txtTrackingNo.focus
	Set gActiveElement = document.activeElement	
End Function
'===========================================================================================================
Function SetItemGroup(byval arrRet)
	frm1.txtItemGroupCd.Value    = arrRet(0)  
	frm1.txtItemGroupNm.Value    = arrRet(1)  
End Function


'==========================================================================================
'   Event Name : vspdData_Click
'   Event Desc :
'==========================================================================================
Sub vspdData_Click(ByVal Col , ByVal Row )

    Call SetPopupMenuItemInf("0000111111")
    
	gMouseClickStatus = "SPC"   

    Set gActiveSpdSheet = frm1.vspdData

    If frm1.vspdData.MaxRows = 0 Then
       Exit Sub
   	End If
   	
   	If Row <= 0 Then
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
    
End Sub

'==========================================================================================
'   Event Name : vspdData_MouseDown(Button,Shift,x,y)
'   Event Desc :
'==========================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
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

'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then Exit Sub
   
	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then
		If lgStrPrevKey1 <> "" Then	
			Call DisableToolBar(parent.TBC_QUERY)
            If DBQuery = False Then 
               Call RestoreToolBar()
               Exit Sub
            End If 
		End If
    End if
End Sub

'=======================================================================================================
'   Event Name : txtFromDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtFromReqrdDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtFromReqrdDt.Action = 7
        Call SetFocusToDocument("M")
        frm1.txtFromReqrdDt.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtToDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtToReqrdDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtToReqrdDt.Action = 7
        Call SetFocusToDocument("M")
        frm1.txtToReqrdDt.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtFromReqrdDt_KeyDown(keycode, shift)
'   Event Desc : Enter Event시 FncQuery한다.
'=======================================================================================================
Sub txtFromReqrdDt_KeyDown(keycode, shift)
	If keycode = 13 Then
		Call MainQuery()
	End If
End Sub

'=======================================================================================================
'   Event Name : txtToReqrdDt_KeyDown(keycode, shift)
'   Event Desc : Enter Event시 FncQuery한다.
'=======================================================================================================
Sub txtToReqrdDt_KeyDown(keycode, shift)
	If keycode = 13 Then
		Call MainQuery()
	End If
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
	Call ggoSpread.ReOrderingSpreadData()

End Sub


'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================
Function FncQuery() 
    Dim IntRetCD 
    
    FncQuery = False
    Err.Clear       
    
    If frm1.txtPlantCd.value = "" Then
		frm1.txtPlantNm.value = ""
	End If
	
	If frm1.txtItemCd.value = "" Then
		frm1.txtItemnm.value = ""
	End If

    Call ggoOper.ClearField(Document, "2") 
    Call InitVariables
    
    If Not chkField(Document, "1") Then	
       Exit Function
    End If
    
    If Not(ValidDateCheck(frm1.txtFromReqrdDt, frm1.txtToReqrdDt)) Then
        Exit Function
    End If

    If DbQuery = False Then
		Exit Function
	End If
       
    FncQuery = True
    
End Function

'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================
Function FncFind() 
    Call parent.FncFind(parent.C_MULTI, False)
End Function

'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================
Function FncPrint() 
    Call parent.fncPrint()
End Function

'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================
Function FncExcel() 
    Call parent.FncExport(parent.C_MULTI)
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

Function FncExit()
	FncExit = True
End Function

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbQuery() 
    Dim LngLastRow      
    Dim LngMaxRow       
    Dim LngRow          
    Dim strTemp         
    Dim StrNextKey      

    DbQuery = False
    
    Call LayerShowHide(1)
    
    Err.Clear

    Dim strVal
    
    With frm1
    
    If lgIntFlgMode = parent.OPMD_UMODE Then
		strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001
		strVal = strVal & "&lgStrPrevKey1=" & lgStrPrevKey1
		strVal = strVal & "&lgStrPrevKey2=" & lgStrPrevKey2
		strVal = strVal & "&lgIntFlgMode=" & lgIntFlgMode
		strVal = strVal & "&txtPlantCd=" & Trim(.hPlantCd.value)		
		strVal = strVal & "&txtFromReqrdDt=" & Trim(.hFromReqrdDt.value)
		strVal = strVal & "&txtToReqrdDt=" & Trim(.hToReqrdDt.value)
		strVal = strVal & "&txtItemCd=" & Trim(.hItemCd.value)
		strVal = strVal & "&cboRelatedType=" & Trim(.hRelatedType.value)
		strVal = strVal & "&txtTrackingNo=" & Trim(.hTrackingNo.value)
		strVal = strVal & "&txtItemGroupCd=" & Trim(.hItemGroupCd.value)
		strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
	Else
		strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001
		strVal = strVal & "&lgStrPrevKey1=" & lgStrPrevKey1
		strVal = strVal & "&lgStrPrevKey2=" & lgStrPrevKey2
		strVal = strVal & "&lgIntFlgMode=" & lgIntFlgMode
		strVal = strVal & "&txtPlantCd=" & Trim(.txtPlantCd.value)
		strVal = strVal & "&txtFromReqrdDt=" & Trim(.txtFromReqrdDt.Text)
		strVal = strVal & "&txtToReqrdDt=" & Trim(.txtToReqrdDt.Text)
		strVal = strVal & "&txtItemCd=" & Trim(.txtItemCd.value)
		strVal = strVal & "&cboRelatedType=" & Trim(.cboRelatedType.value)
		strVal = strVal & "&txtTrackingNo=" & Trim(.txtTrackingNo.value)
		strVal = strVal & "&txtItemGroupCd=" & Trim(.txtItemGroupCd.value)
		strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
	
	End IF	

	Call RunMyBizASP(MyBizASP, strVal)
    End With
    
    DbQuery = True
    
End Function
'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================
Function DbQueryOk()
	Call SetToolbar("11000000000111")
    
    lgIntFlgMode = parent.OPMD_UMODE
	lgBlnFlgChgValue = False    
    Call ggoOper.LockField(Document, "Q")

	frm1.vspdData.Focus
End Function
