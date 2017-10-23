
Option Explicit    

Const BIZ_PGM_ID		= "m2111mb1.asp" 
Const BIZ_PGM_ID2		= "m2111mb1_1.asp" 
Const BIZ_PGM_JUMP_ID	= "m2111qa1"

'==========================================  1.2.2 Global 변수 선언  =====================================
Dim lgBlnFlgChgValue    
Dim lgIntGrpCount    
Dim lgIntFlgMode    

Dim IsOpenPop          
Dim lgPageNo2

Dim C_SpplCd
Dim C_SpplNm
Dim C_QuotaRate
Dim C_ApportionQty
Dim C_PlanDt
Dim C_GrpCd
Dim C_GrpNm

'==========================================  2.1.1 InitVariables()  ======================================
Sub InitVariables()
   lgIntFlgMode = Parent.OPMD_CMODE   
   lgBlnFlgChgValue = False    
   lgIntGrpCount = 0           
   IsOpenPop = False   
   lgPageNo2    = ""
End Sub

'----------------------------------------------------------------------------
'  Field의 Tag속성을 Protect로 전환,복구 시키는 함수 Biz Logic에서도 호출 
'----------------------------------------------------------------------------
Function ChangeTag(Byval Changeflg,ByVal Toolbarflg)
	If Changeflg = True then
		Call ggoOper.SetReqAttr(frm1.txtPlantCd,"Q")
		Call ggoOper.SetReqAttr(frm1.txtItemCd,"Q")
		Call ggoOper.SetReqAttr(frm1.txtReqDt,"Q")
		Call ggoOper.SetReqAttr(frm1.txtDlvyDt,"Q")
		Call ggoOper.SetReqAttr(frm1.txtReqQty,"Q")
		Call ggoOper.SetReqAttr(frm1.txtReqUnitCd,"Q")
		Call ggoOper.SetReqAttr(frm1.txtDeptCd,"Q")
		Call ggoOper.SetReqAttr(frm1.txtEmpCd,"Q")
		Call ggoOper.SetReqAttr(frm1.txtStorageCd,"Q")
		Call ggoOper.SetReqAttr(frm1.txtTrackingNo,"Q")
		Call ggoOper.SetReqAttr(frm1.txtOrgCd,"Q")
	
		If Toolbarflg = True then
			Call SetToolBar("11100000001111")
		End if
	Else
		Call ggoOper.SetReqAttr(frm1.txtPlantCd,"N")
		Call ggoOper.SetReqAttr(frm1.txtItemCd,"N")
		Call ggoOper.SetReqAttr(frm1.txtReqDt,"N")
		Call ggoOper.SetReqAttr(frm1.txtDlvyDt,"N")
		Call ggoOper.SetReqAttr(frm1.txtReqQty,"N")
		Call ggoOper.SetReqAttr(frm1.txtReqUnitCd,"N")
		Call ggoOper.SetReqAttr(frm1.txtDeptCd,"D")
		Call ggoOper.SetReqAttr(frm1.txtEmpCd,"D")
		Call ggoOper.SetReqAttr(frm1.txtStorageCd,"D")
		Call ggoOper.SetReqAttr(frm1.txtTrackingNo,"D")
		Call ggoOper.SetReqAttr(frm1.txtOrgCd,"N")
		If Toolbarflg = True then
			Call SetToolBar("11111000001111")
		End if 
	  	Call changeTagTracking()
	End if 
End Function

'======================================================================================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'=======================================================================================================
Sub InitSpreadSheet()
	Call InitSpreadPosVariables()

	With frm1
		.vspdData2.ReDraw = false

		ggoSpread.Source = frm1.vspdData2
        ggoSpread.Spreadinit "V20030513",, parent.gAllowDragDropSpread
	   .vspdData2.MaxCols = C_GrpNm+1
	   .vspdData2.MaxRows = 0

		Call GetSpreadColumnPos("A")

		ggoSpread.SSSetEdit   C_SpplCd, "공급처", 15,,,15,2
		ggoSpread.SSSetEdit	  C_SpplNm, "공급처명", 20
		SetSpreadFloatLocal	  C_QuotaRate, "배분비율(%)",15,1,5
		SetSpreadFloatLocal   C_ApportionQty, "배부량", 15, 1,3
		ggoSpread.SSSetDate	  C_PlanDt, "발주예정일", 15,2,parent.gDateFormat		
		ggoSpread.SSSetEdit	  C_GrpCd, "구매그룹", 10,,,10,2
		ggoSpread.SSSetEdit   C_GrpNm, "구매그룹명", 20
		
		
		Call ggoSpread.MakePairsColumn(C_SpplCd,C_SpplNm)
		Call ggoSpread.MakePairsColumn(C_GrpCd,C_GrpNm)

		Call ggoSpread.SSSetColHidden(.vspdData2.MaxCols,	.vspdData2.MaxCols,	True)	
		
		.vspdData2.ReDraw = True
    End With

    Call SetSpreadLock()
End Sub

Sub SetSpreadLock()
	With frm1
		.vspdData2.ReDraw = False
		ggoSpread.Source = .vspdData2
		ggoSpread.SpreadLock 1 , -1
		.vspdData2.ReDraw = True
	End With
End Sub

'==========================================  2.2.7 InitSpreadPosVariables()  =============================
' Function Name : InitSpreadPosVariables
' Function Desc : This method Assigns Sequential Number to spread sheet column 
'=========================================================================================================
Sub InitSpreadPosVariables()
	C_SpplCd			=	1
	C_SpplNm			=	2
	C_QuotaRate			=	3
	C_ApportionQty		=	4
	C_PlanDt			=	5
	C_GrpCd				=	6
	C_GrpNm				=	7
End Sub

'========================================================================================
' Function Name : GetSpreadColumnPos
' Description   : 
'========================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData2
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			 	C_SpplCd			=	iCurColumnPos(1)
				C_SpplNm			=	iCurColumnPos(2)
				C_QuotaRate			=	iCurColumnPos(3)
				C_ApportionQty		=	iCurColumnPos(4)	
				C_PlanDt			=	iCurColumnPos(5)
				C_GrpCd				=	iCurColumnPos(6)
				C_GrpNm				=	iCurColumnPos(7)
	End Select    
End Sub


Sub SetSpreadFloatLocal(ByVal iCol , ByVal Header , ByVal dColWidth , ByVal HAlign , ByVal iFlag )
   Select Case iFlag
        Case 2                                                              '금액 
            ggoSpread.SSSetFloat iCol, Header, dColWidth, parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec, HAlign
        Case 3                                                              '수량 
            ggoSpread.SSSetFloat iCol, Header, dColWidth, parent.ggQtyNo       ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec, HAlign,,"P"
        Case 4                                                              '단가 
            ggoSpread.SSSetFloat iCol, Header, dColWidth, parent.ggUnitCostNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec, HAlign,,"Z"
        Case 5                                                              '환율 
            ggoSpread.SSSetFloat iCol, Header, dColWidth, parent.ggExchRateNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec, HAlign,,"Z"
    End Select
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
' Function Name : vspdData2_MouseDown
' Function Desc : 
'========================================================================================
Sub vspdData2_MouseDown(Button , Shift , x , y)
   If Button = 2 And gMouseClickStatus = "SPC" Then
      gMouseClickStatus = "SPCR"
   End If
End Sub    

'========================================================================================================
'   Event Name : vspdData2_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData2_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)
End Sub

Sub vspdData2_GotFocus()
    ggoSpread.Source = frm1.vspdData2
End Sub

'========================================================================================================
'   Event Name : vspdData2_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData2_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
End Sub

Sub vspdData2_Click(ByVal Col, ByVal Row)	'###그리드 컨버전 주의부분###
 	gMouseClickStatus = "SPC"   
	 	 	
 	Set gActiveSpdSheet = frm1.vspdData2
	Call SetPopupMenuItemInf("0000111111")         '화면별 설정 
	
	If frm1.vspdData2.MaxRows = 0 Then
 		Exit Sub
 	End If
End Sub

'======================================================================================================
'   Event Name : vspdData2_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'=======================================================================================================
Sub vspdData2_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
	If frm1.vspdData2.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData2, NewTop) Then
    	If lgPageNo2 <> "" Then							
 			If CheckRunningBizProcess = True Then
				Exit Sub
			End If			
			Call DisableToolBar(parent.TBC_QUERY)
			If DBQuery2 = False Then
				Call RestoreToolBar()
				Exit Sub
			End If
    	End If
    End If
End Sub

'========================================================================================
' Function Name : FncSplitColumn
'========================================================================================
Sub FncSplitColumn()
    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   
End Sub

'------------------------------------------  OpenTrackingNo()  -------------------------------------------------
Function OpenTrackingNo()
	Dim arrRet
	Dim arrParam(6)

	If IsOpenPop = True Or UCase(frm1.txtTrackingNo.ClassName)=UCase(Parent.UCN_PROTECTED) Then Exit Function
	 
	If frm1.hdnTrackingflg.value <> "Y" Then Exit Function
	 
	if  Trim(frm1.txtitemCd.Value) = "" then
		Call DisplayMsgBox("17A002", "X", "품목", "X")
		frm1.txtitemCd.focus
		Exit Function
	End if
	 
	If Trim(frm1.txtPlantCd.Value) = "" then
		Call DisplayMsgBox("17A002", "X", "공장", "X")
		frm1.txtPlantCd.focus
		Exit Function
	End if

	IsOpenPop = True 
	 
	arrParam(0) = ""
	arrParam(1) = ""
	arrParam(2) = Trim(frm1.txtPlantCd.Value)
	arrParam(3) = ""
	'arrParam(3) = Trim(frm1.txtitemCd.Value)
	  
	arrParam(4) = ""
	arrParam(5) = " and A.tracking_no not in (" & FilterVar("*", "''", "S") & " ) " 
	arrParam(6) = "M" 
	    
	arrRet = window.showModalDialog("../s3/s3135pa1.asp", Array(window.parent, arrParam), _
	"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	 
	IsOpenPop = False

	If arrRet = "" Then
		frm1.txtTrackingNo.focus
		Exit Function
	Else
		If frm1.hdnTrackingflg.value <> "Y" Then Exit Function
		frm1.txtTrackingNo.Value = arrRet
		frm1.txtTrackingNo.focus	
		Set gActiveElement = document.activeElement
		lgBlnFlgChgValue = True
	End If 
End Function

'------------------------------------------  OpenPlant()  -------------------------------------------------
Function OpenPlant()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtPlantCd.ClassName)=UCase(Parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "공장" 
	arrParam(1) = "B_Plant"    
	 
	arrParam(2) = Trim(frm1.txtPlantCd.Value)
	' arrParam(3) = Trim(frm1.txtPlantNm.Value)
	 
	arrParam(4) = ""   
	arrParam(5) = "공장"   
	 
	arrField(0) = "Plant_CD" 
	arrField(1) = "Plant_NM" 
	    
	arrHeader(0) = "공장"  
	arrHeader(1) = "공장명"  

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
	"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	 
	IsOpenPop = False
	 
	If arrRet(0) = "" Then
		frm1.txtPlantCd.focus
		Exit Function
	Else
		frm1.txtPlantCd.Value= arrRet(0)  
		frm1.txtPlantNm.Value= arrRet(1)
		Call changeItemPlant()
		frm1.txtPlantCd.focus	
		Set gActiveElement = document.activeElement   
		lgBlnFlgChgValue = True
	End If 
 
End Function

'------------------------------------------  OpenItem()  -------------------------------------------------
'Function OpenItem()
'	Dim arrRet
'	Dim arrParam(5), arrField(6), arrHeader(6)
'	Dim iCalledAspName
'	Dim IntRetCD
'	
'	If IsOpenPop = True Then Exit Function
'	If UCase(frm1.txtItemCd.ClassName) = UCase(Parent.UCN_PROTECTED) Then Exit Function
'	 
'	If Trim(frm1.txtPlantCd.Value) = "" Then
'		Call DisplayMsgBox("17A002", "X", "공장", "X")
'		frm1.txtPlantCd.focus
'		Exit Function
'	End If
'	 
'	IsOpenPop = True
'	'***2003.3월 패치분 수정(2003.02.26-Lee,Eun Hee)-유효일추가*****
'	arrParam(0) = "품목"  
'	arrParam(1) = "B_Item_By_Plant,B_Plant,B_Item"
'	 
'	arrParam(2) = Trim(frm1.txtitemCd.Value) 
'
'	arrParam(4) = "B_Item_By_Plant.Plant_Cd = B_Plant.Plant_Cd And "
'	arrParam(4) = arrParam(4) & "B_Item_By_Plant.Item_Cd = B_Item.Item_Cd and B_Item.phantom_flg = " & FilterVar("N", "''", "S") & "  "
'	 
'	if Trim(frm1.txtPlantCd.Value)<>"" then
'		arrParam(4) = arrParam(4) & "And B_Plant.Plant_Cd= " & FilterVar(UCase(frm1.txtPlantCd.Value), "''", "S") & " "  
'	End if 
'	
'	arrParam(4) = arrParam(4) & " AND B_Item_By_Plant.VALID_FROM_DT <=  " & FilterVar(UNIConvDate(Trim(frm1.txtReqDt.text)), "''", "S") & ""
'	arrParam(4) = arrParam(4) & " AND B_Item_By_Plant.VALID_TO_DT   >=  " & FilterVar(UNIConvDate(Trim(frm1.txtReqDt.text)), "''", "S") & " "
'
'	arrParam(5) = "품목"      
'	 
'	arrField(0) = "B_Item.Item_Cd"    
'	arrField(1) = "B_Item.Item_NM" 
'	arrField(2) = "B_Plant.Plant_Cd"   
'	arrField(3) = "B_Plant.Plant_NM"
'	
'	arrHeader(2) = "공장"     
'	arrHeader(3) = "공장명"     
'	    
'	arrHeader(0) = "품목"     
'	arrHeader(1) = "품목명"     
'	    
'	iCalledAspName = AskPRAspName("M1111PA1")
'	
'	If Trim(iCalledAspName) = "" Then
'		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "M1111PA1", "X")
'		IsOpenPop = False
'		Exit Function
'	End If
'	
'	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam, arrField, arrHeader), _
'		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
'
'	IsOpenPop = False
'
'	If arrRet(0) = "" Then
'		frm1.txtItemCd.focus	
'		Exit Function
'	Else
'		frm1.txtItemCd.Value    = arrRet(0)  
'		frm1.txtItemNm.Value    = arrRet(1)  
'		frm1.txtSpec.Value		= arrRet(2) 
'		frm1.txtItemCd.focus	
'		Set gActiveElement = document.activeElement   
'		lgBlnFlgChgValue = True
'		Call changeItemPlant()
'	End If 
'End Function

'------------------------------------------  OpenItem()  -------------------------------------------------
Function OpenItem()
	Dim arrRet
	Dim arrParam(5), arrField(2)
	Dim iCalledAspName
	Dim IntRetCD
	
	If IsOpenPop = True Then Exit Function
	If UCase(frm1.txtItemCd.ClassName) = UCase(Parent.UCN_PROTECTED) Then Exit Function
	 
	If Trim(frm1.txtPlantCd.Value) = "" Then
		Call DisplayMsgBox("17A002", "X", "공장", "X")
		frm1.txtPlantCd.focus
		Exit Function
	End If
	 
	IsOpenPop = True
	'***2003.3월 패치분 수정(2003.02.26-Lee,Eun Hee)-유효일추가*****
	arrParam(0) = Trim(frm1.txtPlantCd.value)	' Plant Code
	arrParam(1) = Trim(frm1.txtItemCd.value)		' Item Code
	arrParam(2) = "!"	' “12!MO"로 변경 -- 품목계정 구분자 조달구분 
	arrParam(3) = "30!P"
	arrParam(4) = ""		'-- 날짜 
	arrParam(5) = ""		'-- 자유(b_item_by_plant a, b_item b: and 부터 시작)
	
	arrField(0) = 1 ' -- 품목코드 
	arrField(1) = 2 ' -- 품목명 
	arrField(2) = 3 ' -- Spec
	    
	iCalledAspName = AskPRAspName("B1B11PA3")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "B1B11PA3", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam, arrField), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtItemCd.focus	
		Exit Function
	Else
		frm1.txtItemCd.Value    = arrRet(0)  
		frm1.txtItemNm.Value    = arrRet(1)  
		frm1.txtSpec.Value		= arrRet(2) 
		frm1.txtItemCd.focus	
		Set gActiveElement = document.activeElement   
		lgBlnFlgChgValue = True
		Call changeItemPlant()
	End If 
End Function

'------------------------------------------  Openunit()  -------------------------------------------------
Function OpenUnit()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtReqUnitCd.ClassName)=UCase(Parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "요청단위"    
	arrParam(1) = "B_Unit_OF_MEASURE"   
	 
	arrParam(2) = Trim(frm1.txtReqUnitCd.Value) 
	 
	arrParam(4) = ""       
	arrParam(5) = "요청단위"     
	 
	arrField(0) = "Unit"    
	arrField(1) = "Unit_Nm"    
	    
	arrHeader(0) = "요청단위"  
	arrHeader(1) = "요청단위명"  
	    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
	"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtReqUnitCd.focus
		Exit Function
	Else
		frm1.txtReqUnitCd.Value    = arrRet(0)  
		frm1.txtReqUnitCd.focus	
		Set gActiveElement = document.activeElement  
		lgBlnFlgChgValue = True  
	End If  
End Function

'------------------------------------------  OpenStorage()  -------------------------------------------------
Function OpenStorage()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function
	if UCase(frm1.txtStorageCd.className) = UCase(Parent.UCN_PROTECTED) then Exit Function
	if Trim(frm1.txtPlantCd.value)="" then
		Call DisplayMsgBox("17A002", "X", "공장", "X")
		frm1.txtPlantCd.focus
		Exit function 
	End if 
	 
	IsOpenPop = True

	arrParam(0) = "입고창고"   
	arrParam(1) = "B_Storage_location,B_Plant" 
	 
	arrParam(2) = Trim(frm1.txtStorageCD.Value) 
	 
	arrParam(4) = "B_Storage_location.Plant_Cd=B_Plant.Plant_Cd And " 
	arrParam(4) = arrParam(4) & "B_Plant.Plant_Cd= " & FilterVar(UCase(frm1.txtPlantCd.Value), "''", "S") & " "
	arrParam(5) = "입고창고"     
	 
	arrField(0) = "B_Storage_location.Sl_Cd" 
	arrField(1) = "B_Storage_location.Sl_Nm" 
	    
	arrHeader(0) = "입고창고"    
	arrHeader(1) = "입고창고명"    
	    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
	"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtStorageCD.focus
		Exit Function
	Else
		frm1.txtStorageCd.Value    = arrRet(0)  
		frm1.txtStorageNm.Value    = arrRet(1) 
		frm1.txtStorageCd.focus	
		Set gActiveElement = document.activeElement    
		lgBlnFlgChgValue = True
	End If 
End Function
'-----------------------------------  OpenReqNo()  -------------------------------------------------
Function OpenReqNo()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim iCalledAspName
	Dim IntRetCD
	
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "요청번호"     
	arrParam(1) = "M_PURCHASE_REQUISITION"  
	 
	arrParam(2) = Trim(frm1.txtReqNo.Value)  
	 
	arrParam(4) = ""       
	arrParam(5) = "요청번호"     
	 
	arrField(0) = "Pr_No"     
	arrField(1) = "F2" & Parent.gColSep & "Convert(varchar(10), req_qty)" 
	arrField(2) = "req_unit"          
	    
	arrHeader(0) = "요청번호"         
	arrHeader(1) = "수량"          
	arrHeader(2) = "단위"     
	    
	iCalledAspName = AskPRAspName("M2111PA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "M2111PA1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent,""), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtReqNo.focus
		Exit Function
	Else
		frm1.txtReqNo.Value= arrRet(0) 
		frm1.txtReqNo.focus	
		Set gActiveElement = document.activeElement 
	End If 
End Function

'------------------------------------------  OpenDept()  -------------------------------------------------
Function OpenDept()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtDeptCd.ClassName)=UCase(Parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "요청부서" 
	arrParam(1) = "B_ACCT_DEPT"    
	 
	arrParam(2) = Trim(frm1.txtDeptCd.Value)
	' arrParam(3) = Trim(frm1.txtDeptNm.Value)
	 
	arrParam(4) = "ORG_CHANGE_ID= " & FilterVar(Parent.gChangeOrgId, "''", "S") & " "
	arrParam(5) = "요청부서"   

	arrField(0) = "DEPT_CD" 
	arrField(1) = "DEPT_NM" 
	    
	arrHeader(0) = "요청부서"  
	arrHeader(1) = "요청부서명"  

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
	"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	 
	IsOpenPop = False
	 
	If arrRet(0) = "" Then
		frm1.txtDeptCd.focus	
		Exit Function
	Else
		frm1.txtDeptCd.Value = arrRet(0)
		frm1.txtDeptNm.Value = arrRet(1)
		frm1.txtDeptCd.focus	
		Set gActiveElement = document.activeElement
		lgBlnFlgChgValue = True
	End If 
 
End Function
'------------------------------------------  OpenORG()  -------------------------------------------------
Function OpenORG()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True or UCase(frm1.txtOrgCd.ClassName)=UCase(Parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "구매조직"   
	arrParam(1) = "B_Pur_Org"   
	 
	arrParam(2) = Trim(frm1.txtOrgCd.Value)
	 
	arrParam(4) = "USAGE_FLG=" & FilterVar("Y", "''", "S") & " "   
	arrParam(5) = "구매조직"    
	 
	arrField(0) = "PUR_ORG"     
	arrField(1) = "PUR_ORG_NM"    
	    
	arrHeader(0) = "구매조직"   
	arrHeader(1) = "구매조직명"   
	    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
	"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtOrgCd.focus	
		Exit Function
	Else
		frm1.txtOrgCd.Value = arrRet(0)
		frm1.txtOrgNm.Value = arrRet(1)
		frm1.txtOrgCd.focus	
		Set gActiveElement = document.activeElement
		lgBlnFlgChgValue = True
	End If
End Function

'++++++++++++++++++++++++++++++++++++++++++  2.5 개발자 정의 함수  +++++++++++++++++++++++++++++++++++++++
'    개별적 프로그램 마다 필요한 개발자 정의 Procedure (Sub, Function, Validation & Calulation 관련 함수)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ !-->
Function changeTagTracking()
	If UCase(Trim(frm1.hdnTrackingflg.Value)) <> "Y" Then
		ggoOper.SetReqAttr frm1.txtTrackingNo, "Q"
	Else
		ggoOper.SetReqAttr frm1.txtTrackingNo, "N"
	End If
End Function


'========================================================================================
' Function Name : changeItemPlant()
'========================================================================================
Function changeItemPlant()
    Dim strVal    
	Err.Clear
	
	If gLookUpEnable = False Then Exit Function
	If CheckRunningBizProcess = True Then Exit Function
	    
	if Trim(frm1.txtPlantCd.Value) = "" or Trim(frm1.txtItemCd.Value) = "" then
		Exit Function
	End if
	    
	changeItemPlant = False                 
	If LayerShowHide(1) = False Then Exit Function
	        
	strVal = BIZ_PGM_ID & "?txtMode=" & "changeItemPlant"
	strVal = strVal & "&txtItemCd=" & Trim(frm1.txtItemCd.Value)
	strVal = strVal & "&txtPlantCd=" & Trim(frm1.txtPlantCd.Value)
	Call RunMyBizASP(MyBizASP, strVal)
	    
	frm1.txtTrackingNo.Value = ""
	changeItemPlant = True                  
End Function

'========================================================================================
' Function Name : CookiePage
'========================================================================================
Sub ReadCookiePage()
	Dim strTemp

	strTemp = ReadCookie("ReqNo")
	If strTemp = "" then Exit sub
	
	frm1.txtReqNo.value = ReadCookie("ReqNo")
	Call WriteCookie("ReqNo" , "")
	Call MainQuery()
End Sub

Function WriteCookiePage()
	Dim IntRetCD

	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900017", Parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If
	    
	Call WriteCookie("m2111ma1_plantcd", frm1.txtPlantCd.Value)
	Call WriteCookie("m2111ma1_itemcd", frm1.txtItemCd.Value)
	 
	Call PgmJump(BIZ_PGM_JUMP_ID)
End Function


'==========================================================================================
'   Event Name : OCX Event    
'==========================================================================================
Sub txtReqDt_DblClick(Button)
	if Button = 1 then
		frm1.txtReqDt.Action = 7
        Call SetFocusToDocument("M")  
        frm1.txtReqDt.Focus
	End if
End Sub

Sub txtDlvyDt_DblClick(Button)
	if Button = 1 then
		frm1.txtDlvyDt.Action = 7
        Call SetFocusToDocument("M")  
        frm1.txtDlvyDt.Focus
	End if
End Sub

Sub txtReqDt_Change()
	lgBlnFlgChgValue = true 
End Sub

Sub txtDlvyDt_Change()
	lgBlnFlgChgValue = true 
End Sub

Sub txtReqQty_Change()
	lgBlnFlgChgValue = true 
End Sub

'========================================================================================
' Function Name : FncQuery
'========================================================================================
Function FncQuery() 
	Dim IntRetCD 
	    
	Err.Clear                                       
	
	FncQuery = False                                
	
	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If
	'-----------------------
	'Erase contents area
	'-----------------------
	Call ggoOper.ClearField(Document, "2")   
	Call InitVariables
	'-----------------------
	'Check condition area
	'----------------------- 
	'If Not chkField(Document, "1") Then  Exit Function
	If Not chkFieldByCell(frm1.txtReqNo,"A",1) Then Exit Function
	'-----------------------
	'Query function call area
	'----------------------- 
	If DbQuery = False Then Exit Function
	FncQuery = True         
End Function

'========================================================================================
' Function Name : FncNew
'========================================================================================
Function FncNew() 
	Dim IntRetCD 
		    
	FncNew = False                                  
	'-----------------------
	'Check previous data area
	'-----------------------
	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900015", Parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If
	'-----------------------
	'Erase condition area
	'Erase contents area
	'-----------------------
	Call ggoOper.ClearField(Document, "A")          

	Call LockObjectField(frm1.txtReqDt, "R")
    Call LockObjectField(frm1.txtReqQty, "R")
    Call LockObjectField(frm1.txtDlvyDt, "R")
    Call LockObjectField(frm1.txtPoQty, "P")
    Call LockObjectField(frm1.txtGmQty, "P")
		       
	Call ChangeTag(False,False)
	Call ggoOper.SetReqAttr(frm1.txtReqNo2, "D")
	Call SetDefaultVal
	Call InitVariables
		       
	FncNew = True         
End Function

'========================================================================================
' Function Name : FncDelete
'========================================================================================
Function FncDelete() 
	Dim IntRetCD

	IntRetCD = DisplayMsgBox("900003", Parent.VB_YES_NO, "X", "X")
	If IntRetCD = vbNo Then Exit Function
	    
	FncDelete = False        
	'-----------------------
	'Precheck area
	'-----------------------
	If lgIntFlgMode <> Parent.OPMD_UMODE Then              
		Call DisplayMsgBox("900002", "X", "X", "X")
		Exit Function
	End If
	'-----------------------
	'Delete function call area
	'-----------------------
	If DbDelete = False Then Exit Function
	FncDelete = True                                
End Function

'========================================================================================
' Function Name : FncSave
'========================================================================================
Function FncSave() 
	Dim IntRetCD 
	    
	Err.Clear
	
	FncSave = False                                 
	If CheckRunningBizProcess = True Then Exit Function
	'-----------------------
	'Precheck area
	'-----------------------
	If lgBlnFlgChgValue = False Then
		IntRetCD = DisplayMsgBox("900001", "X", "X", "X")
		Exit Function
	End If
	'-----------------------
	'Check content area
	'-----------------------
	'If Not chkField(Document, "2") Then Exit Function
	
	If Not chkFieldByCell(frm1.txtPlantCd,"A",1)	Then Exit Function
	If Not chkFieldByCell(frm1.txtItemCd,"A",1)		Then Exit Function
	If Not chkFieldByCell(frm1.txtReqDt,"A",1)		Then Exit Function
	If Not chkFieldByCell(frm1.txtReqQty,"A",1)		Then Exit Function
	If Not chkFieldByCell(frm1.txtDlvyDt,"A",1)		Then Exit Function
	If Not chkFieldByCell(frm1.txtReqUnitCd,"A",1)	Then Exit Function
	If Not chkFieldByCell(frm1.txtOrgCd,"A",1)		Then Exit Function

	    
	If Trim(UNICDbl(frm1.txtReqQty.Text)) = "" Or Trim(UNICDbl(frm1.txtReqQty.Text)) = "0" then
		Call DisplayMsgBox("970021", "X","요청량", "X")
		frm1.txtReqQty.focus
		Set gActiveElement = document.activeElement
		Exit Function
	End if
    '-----------------------
    'Save function call area
    '-----------------------
    If DbSave = False Then Exit Function
    FncSave = True                                  
End Function

'========================================================================================
' Function Name : FncCopy
'========================================================================================
Function FncCopy() 
	Dim IntRetCD
	
	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900017", Parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If
	    
	lgIntFlgMode = Parent.OPMD_CMODE      
	Call ggoOper.ClearField(Document, "1")             
	    
	frm1.txtReqNo2.value = ""
	frm1.txtReqStateCd.Value = ""
	frm1.txtReqStateNm.Value = "" 
	frm1.txtPoQty.Text = "0"
	frm1.txtGmQty.Text = "0"
	frm1.txtReqTypeCd.Value = ""
	frm1.txtReqTypeNm.Value = ""
	frm1.hdnProcurType.value = ""
	
	Call LockObjectField(frm1.txtReqDt, "R")
    Call LockObjectField(frm1.txtReqQty, "R")
    Call LockObjectField(frm1.txtDlvyDt, "R")
    Call LockObjectField(frm1.txtPoQty, "P")
    Call LockObjectField(frm1.txtGmQty, "P")
    
	
	Call ChangeTag(False,False)
	Call SetToolBar("11101000000111")
	Call ggoOper.SetReqAttr(frm1.txtReqNo2, "D")
	'@@수정(Lee,Eun Hee)
	If frm1.vspdData2.MaxRows > 0 Then
		frm1.vspdData2.MaxRows = 0
	End If
	
	frm1.txtReqNo2.focus 	
	Set gActiveElement = document.activeElement
	lgBlnFlgChgValue = True
End Function

'========================================================================================
' Function Name : FncPrint
'========================================================================================
Function FncPrint() 
    Call parent.FncPrint()
End Function

'========================================================================================
' Function Name : FncExcel
'========================================================================================
Function FncExcel() 
    Call parent.FncExport(Parent.C_SINGLE)      
End Function

'========================================================================================
' Function Name : FncFind
'========================================================================================
Function FncFind() 
    Call parent.FncFind(Parent.C_SINGLE , False)               
End Function

'========================================================================================
' Function Name : FncExit
'========================================================================================
Function FncExit()
	Dim IntRetCD
	
	FncExit = False
	
	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If
	FncExit = True
End Function

'========================================================================================
' Function Name : DbDelete
'========================================================================================
Function DbDelete() 
	Dim strVal
	
	Err.Clear                                           
	    
	DbDelete = False         
	    
	If LayerShowHide(1) = False Then Exit Function
	
	strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0003     
	strVal = strVal & "&txtReqNo2=" & Trim(frm1.txtReqNo2.value) 
	Call RunMyBizASP(MyBizASP, strVal)        
	 
	DbDelete = True                                                 
End Function

'========================================================================================
' Function Name : DbDeleteOk
'========================================================================================
Function DbDeleteOk()            
	lgBlnFlgChgValue = False
	Call MainNew()
End Function

'========================================================================================
' Function Name : DbQuery
'========================================================================================
Function DbQuery()
    Dim strVal
    
	Err.Clear                                                       
	    
	DbQuery = False                                                 
	If LayerShowHide(1) = False Then Exit Function
	    
	strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001     
	strVal = strVal & "&txtReqNo=" & Trim(frm1.txtReqNo.value)  
	
	Call RunMyBizASP(MyBizASP, strVal)        
	 
	DbQuery = True                                                  
End Function

'========================================================================================
' Function Name : DbQueryOk
'========================================================================================
Function DbQueryOk()            
	'-----------------------
    'Reset variables area
    '-----------------------
    lgIntFlgMode = Parent.OPMD_UMODE          
    lgBlnFlgChgValue = False
    
    'Call ggoOper.LockField(Document, "Q")    
    Call LockHTMLField(frm1.txtPlantCd, "P")
    Call LockHTMLField(frm1.txtItemCd, "P")
    Call LockObjectField(frm1.txtReqDt, "R")
    Call LockObjectField(frm1.txtReqQty, "R")
    Call LockObjectField(frm1.txtDlvyDt, "R")
    Call LockObjectField(frm1.txtPoQty, "P")
    Call LockObjectField(frm1.txtGmQty, "P")
    
    ggoOper.SetReqAttr frm1.txtReqNo2, "Q"
	Call Dbquery2()
End Function

Function DbQuery2()
    Dim strVal
    
	Err.Clear                                                       
	    
	DbQuery2 = False                                                 
	If LayerShowHide(1) = False Then Exit Function
    
	strVal = BIZ_PGM_ID2 & "?txtPrno=" & Trim(frm1.txtReqNo.value)
	strVal = strVal & "&txtMaxRows=" & frm1.vspdData2.MaxRows
	strVal = strVal & "&lgPageNo="		 & lgPageNo2						'☜: Next key tag 
	
	Call RunMyBizASP(MyBizASP, strVal)        
	 
	DbQuery2 = True                                                  
End Function

Function DbQueryOk2()            
	'-----------------------
    'Reset variables area
    '-----------------------
    lgIntFlgMode = Parent.OPMD_UMODE          
    lgBlnFlgChgValue = False
        
    IF Trim(frm1.txtReqStateCd.value) = "RQ" THEN
		If frm1.vspdData2.MaxRows > 0 Then
			ggoOper.SetReqAttr frm1.txtOrgCd, "Q"
		Else
			ggoOper.SetReqAttr frm1.txtOrgCd, "N"
		End If
	END IF
End Function

'========================================================================================
' Function Name : DBSave
'========================================================================================
Function DbSave() 
	Dim strVal

	Err.Clear              
	DbSave = False             

	If LayerShowHide(1) = False Then Exit Function
	
	With frm1
		.txtMode.value = Parent.UID_M0002         
		.txtFlgMode.value = lgIntFlgMode
		  
		Call ExecMyBizASP(frm1, BIZ_PGM_ID)          
	End With
	DbSave = True                                                   
End Function
'========================================================================================
' Function Name : DbSaveOk
'========================================================================================
Function DbSaveOk()             
	Call InitVariables
	Call SetToolBar("11111000001111")
	lgBlnFlgChgValue = False
	Call MainQuery()
End Function

