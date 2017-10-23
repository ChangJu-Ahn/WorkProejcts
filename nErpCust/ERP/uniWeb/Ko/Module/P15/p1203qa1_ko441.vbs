Dim lgIsOpenPop                                             '☜: Popup status                           
Dim	lgTopLeft_A
Dim lgStrPrevKey_A
Dim lgStrPrevKey_B

'--------------- 개발자 coding part(변수선언,Start)-----------------------------------------------------------
Const BIZ_PGM_ID        = "p1203qb1_ko441.asp"                         '☆: Biz logic spread sheet for #1
Const BIZ_PGM_ID1       = "p1203qb2_ko441.asp"                         '☆: Biz logic spread sheet for #2

Const C_MaxKey            = 4                                    '☆☆☆☆: Max key value

 '==========================================  2.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'========================================================================================================= 
Sub InitVariables()
    lgBlnFlgChgValue = False                               'Indicates that no value changed
End Sub

'==========================================  2.2 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : 서버로 부터 필드 정의 정보를 가져옴 
'========================================================================================================= 
Sub SetDefaultVal()
	frm1.txtFromDt.Text	= StartDate
	frm1.txtToDt.Text	= UniConvYYYYMMDDToDate(parent.gDateFormat, "2999","12","31")
	If lgPLCd <> "" then 
		Call ggoOper.SetReqAttr(frm1.txtPlantCd, "Q") 
		frm1.txtPlantCd.Tag = left(frm1.txtPlantCd.Tag,1) & "4" & mid(frm1.txtPlantCd.Tag,3,len(frm1.txtPlantCd.Tag))
        	frm1.txtPlantCd.value = lgPLCd
	End If
End Sub

'========================================= 2.6 InitSpreadSheet() =========================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'==========================================================================================================
Sub InitSpreadSheet(ByVal pvGridId)
    If pvGridId = "A" Then                                   ' 초기화 Spreadsheet #1 
        Call SetZAdoSpreadSheet("P1203QA1", "S", "A", "V20030330", parent.C_SORT_DBAGENT, frm1.vspdData1, C_MaxKey, "X", "X")
    End If

    Call SetZAdoSpreadSheet("P1203QA1", "S", "B", "V20030330", parent.C_SORT_DBAGENT, frm1.vspdData2, C_MaxKey, "X", "X")

    Call SetSpreadLock(pvGridId)
End Sub

'========================================= 2.7 SetSpreadLock() ===========================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'=========================================================================================================
Sub SetSpreadLock(Byval pvGridId)
    If pvGridId = "A" Then
		ggoSpread.Source = frm1.vspdData1
		ggoSpread.SpreadLockWithOddEvenRowColor()
	End If
	
	ggoSpread.Source = frm1.vspdData2
	ggoSpread.SpreadLockWithOddEvenRowColor()
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
    Call InitSpreadSheet(gActiveSpdSheet.id)      
	Call ggoSpread.ReOrderingSpreadData()
End Sub

'------------------------------------------  OpenConItemCd()  -------------------------------------------------
'	Name : OpenConItemCd()
'	Description : Item PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenConItemCd()
	Dim arrRet
	Dim arrParam(5), arrField(6)
	Dim iCalledAspName, IntRetCD
		
	If frm1.txtPlantCd.value = "" Then
		Call DisplayMsgBox("971012", "X", "공장", "X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement 
		Exit Function
	End If

	If lgIsOpenPop = True Then Exit Function
	
	lgIsOpenPop = True
	
	arrParam(0) = Trim(frm1.txtPlantCd.value)	' Item Code
	arrParam(1) = Trim(frm1.txtItemCd.value) 						
	arrParam(2) = "12!MO"							' Combo Set Data:"12!MP" -- 품목계정 구분자 조달구분 
	arrParam(3) = ""							' Default Value
	
	
    arrField(0) = 1 							' Field명(0) : "ITEM_CD"
    arrField(1) = 2 							' Field명(1) : "ITEM_NM"
    
	iCalledAspName = AskPRAspName("B1B11PA3")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "B1B11PA3", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.parent, arrParam, arrField), _
			"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
	
	lgIsOpenPop = False
	
	If arrRet(0) <> "" Then
		frm1.txtItemCd.Value = arrRet(0)
		frm1.txtItemNm.Value = arrRet(1)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtItemCd.focus

End Function

 '------------------------------------------  OpenSppl()  -------------------------------------------------
'	Name : OpenConRouting()
'	Description : Routing PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenConRouting()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	
	If frm1.txtPlantCd.value = "" Then
		Call DisplayMsgBox("971012", "X", "공장", "X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement 
		Exit Function
	End If

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	arrParam(0) = "라우팅"												' 팝업 명칭 
	arrParam(1) = "P_ROUTING_HEADER"										' TABLE 명칭 
	arrParam(2) = Trim(frm1.txtRoutNo.Value)								' Code Condition
	arrParam(3) = ""														' Name Cindition
	arrParam(4) = "PLANT_CD = " & FilterVar(frm1.txtPlantCd.value, "''", "S")	' Where Condition
	arrParam(5) = "라우팅"												' TextBox 명칭 
	
    arrField(0) = "ROUT_NO"												' Field명(0)
    arrField(1) = "DESCRIPTION"												' Field명(1)
    arrField(2) = "MAJOR_FLG"												' Field명(1)
    
    arrHeader(0) = "라우팅"												' Header명(0)
    arrHeader(1) = "라우팅명"											' Header명(1)
    arrHeader(2) = "주라우팅"										' Header명(1)
    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If arrRet(0) <> "" Then
		frm1.txtRoutNo.Value = arrRet(0)
		frm1.txtRoutNm.Value = arrRet(1)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtRoutNo.focus
	
End Function

 '------------------------------------------  OpenConPlant()  -------------------------------------------------
'	Name : OpenConPlant()
'	Description : Plant PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenConPlant()

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Or UCase(frm1.txtPlantCd.className) = UCase(parent.UCN_PROTECTED) Then Exit Function

	lgIsOpenPop = True

	arrParam(0) = "공장"							' 팝업 명칭 
	arrParam(1) = "B_PLANT"								' TABLE 명칭 
	arrParam(2) = Trim(frm1.txtPlantCd.Value)			' Code Condition
	arrParam(3) = ""									' Name Condition
	arrParam(5) = "공장"							' TextBox 명칭 
	
    arrField(0) = "PLANT_CD"							' Field명(0)
    arrField(1) = "PLANT_NM"							' Field명(1)
        
    arrHeader(0) = "공장"						' Header명(0)
    arrHeader(1) = "공장명"						' Header명(1)
    
    arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If arrRet(0) <> "" Then
		frm1.txtPlantCd.Value = arrRet(0)
		frm1.txtPlantNm.Value = arrRet(1)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtPlantCd.focus
	
End Function

'==================================================================================
' Name : PopZAdoConfigGrid
' Desc :
'==================================================================================
Sub PopZAdoConfigGrid()
	If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
		Exit Sub
	End If

	Call OpenOrderByPopup(gActiveSpdSheet.Id)
End Sub

'===========================================================================
' Function Name : OpenOrderByPopup
' Function Desc : OpenOrderByPopup Reference Popup
'===========================================================================
Function OpenOrderByPopup(ByVal pvGridId)

	Dim arrRet
	
	On Error Resume Next
	
	If lgIsOpenPop = True Then Exit Function
	
	lgIsOpenPop = True

	arrRet = window.showModalDialog("../../ComAsp/ZADOSortPopup.asp", Array(ggoSpread.GetXMLData(pvGridId), gMethodText), _
	         "dialogWidth=" & parent.SORTW_WIDTH & "px; dialogHeight=" & parent.SORTW_HEIGHT & "px;; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If arrRet(0) = "X" Then
		Exit Function
	Else
		Call ggoSpread.SaveXMLData(pvGridId, arrRet(0), arrRet(1))
		Call InitVariables
		Call InitSpreadSheet(pvGridId)
		If pvGridId = "B" Then
			Call vspdData1_Click(frm1.vspdData1.ActiveCol, frm1.vspdData1.ActiveRow)
		End If
   End If
End Function

'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================
Sub Form_QueryUnload(Cancel, UnloadMode )
End Sub

'========================================================================================================
'   Event Name : txtPoFrDt
'   Event Desc :
'=========================================================================================================
Sub txtFromDt_DblClick(Button)
	if Button = 1 then
		frm1.txtFromDt.Action = 7
		Call SetFocusToDocument("M")
		Frm1.txtFromDt.Focus
	End if
End Sub
'========================================================================================================
'   Event Name : txtPoToDt
'   Event Desc :
'========================================================================================================
Sub txtToDt_DblClick(Button)
	if Button = 1 then
		frm1.txtToDt.Action = 7
		Call SetFocusToDocument("M")
		Frm1.txtToDt.Focus
	End if
End Sub

'=======================================================================================================
'   Event Name : txtFromDt_KeyDown(keycode, shift)
'   Event Desc : Enter Event시 FncQuery한다.
'=======================================================================================================
Sub txtFromDt_KeyDown(keycode, shift)
	If keycode = 13 Then
		Call MainQuery()
	End If
End Sub

'=======================================================================================================
'   Event Name : txtToDt_KeyDown(keycode, shift)
'   Event Desc : Enter Event시 FncQuery한다.
'=======================================================================================================
Sub txtToDt_KeyDown(keycode, shift)
	If keycode = 13 Then
		Call MainQuery()
	End If
End Sub

'========================================================================================================
'   Event Name : vspdData1_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'=======================================================================================================
Sub vspdData1_Click(ByVal Col, ByVal Row)
	gMouseClickStatus = "SPC"
	Set gActiveSpdSheet = frm1.vspdData1
	Call SetPopupMenuItemInf("00000000001")
     
	If frm1.vspdData1.MaxRows = 0 Then Exit Sub
	
    If Row <= 0 Then
	ggoSpread.Source = frm1.vspdData1
        Exit Sub
    End If
     lgStrPrevKey_B = "" 
	Call DisableToolBar(parent.TBC_QUERY)  
    Call SetSpreadColumnValue("A", frm1.vspdData1, Col, Row)	

	If DbQuery("B") = False Then
	   Call RestoreToolBar()
	   Exit Sub
	End If

     frm1.vspdData2.MaxRows = 0
     lgStrPrevKey_B = ""                                  'initializes Previous Key
End Sub

'========================================================================================================
'   Event Name : vspdData2_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'=======================================================================================================
Sub vspdData2_Click(ByVal Col, ByVal Row)
	gMouseClickStatus = "SP2C"
	Set gActiveSpdSheet = frm1.vspdData2
	Call SetPopupMenuItemInf("00000000001")
	
	If frm1.vspdData1.MaxRows = 0 Then Exit Sub

    If Row <= 0 Then
		ggoSpread.Source = frm1.vspdData2
        Exit Sub
    End If

    Call SetSpreadColumnValue("B", frm1.vspdData2, Col, Row)	
   
End Sub

'==========================================================================================
'   Event Name : vspdData_MouseDown(Button,Shift,x,y)
'   Event Desc :
'==========================================================================================

Sub vspdData1_MouseDown(Button,Shift,x,y)
		
	If Button <> "1" And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If

End Sub

'==========================================================================================
'   Event Name : vspdData2_MouseDown(Button,Shift,x,y)
'   Event Desc :
'==========================================================================================

Sub vspdData2_MouseDown(Button,Shift,x,y)
		
	If Button <> "1" And gMouseClickStatus = "SP2C" Then
		gMouseClickStatus = "SP2CR"
	End If

End Sub
'==========================================================================================
'   Event Name : vspdData1_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
Sub vspdData1_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    if frm1.vspdData1.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData1, NewTop) Then	'☜: 재쿼리 체크'
		If lgStrPrevKey_A <> "" Then                           '⊙: 다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
			lgTopLeft_A = "Y"
			Call DisableToolBar(parent.TBC_QUERY)  
			If DbQuery("A") = False Then
				Call RestoreToolBar()
				Exit Sub
			End If
		End If
   End if
    
End Sub

'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
Sub vspdData2_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    If frm1.vspdData2.MaxRows =< NewTop + VisibleRowCnt(frm1.vspdData2, NewTop) Then	'☜: 재쿼리 체크'
		If lgStrPrevKey_B <> "" Then                        '⊙: 다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
			Call DisableToolBar(parent.TBC_QUERY)  
			If DbQuery("B") = False Then
				Call RestoreToolBar()
				Exit Sub
			End If

		End If
	End if
    
End Sub

'===========================================================================================
'   Event Name : vspdData1_ColWidthChange
'   Event Desc : 
'===========================================================================================
Sub vspdData1_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
    ggoSpread.Source = frm1.vspdData1
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

'===========================================================================================
'   Event Name : vspdData2_ColWidthChange
'   Event Desc : 
'===========================================================================================
Sub vspdData2_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

Function FncQuery() 

    FncQuery = False                                                        '⊙: Processing is NG
    
    Err.Clear     

    '-----------------------
    'Erase contents area
    '-----------------------
    If frm1.txtPlantCd.value = "" Then
		frm1.txtPlantNm.value = ""
	End If	
	
	If frm1.txtItemCd.value = "" Then
		frm1.txtItemNm.value = ""
	End If
	
	If frm1.txtRoutNo.value = "" Then
		frm1.txtRoutNm.value = ""
	End If

	If ValidDateCheck(frm1.txtFromDt, frm1.txtToDt) = False Then
		Exit Function
	End If
	
    Call ggoOper.ClearField(Document, "2")									'⊙: Clear Contents  Field
    Call InitVariables 														'⊙: Initializes local global variables
    
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then								'⊙: This function check indispensable field
       Exit Function
    End If

    '-----------------------
    'Query function call area
    '-----------------------
    lgStrPrevKey_A = ""
	If DbQuery("A") = False Then   
		Exit Function           
    End If     							

    FncQuery = True		
End Function


'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================

Function FncPrint() 
    Call parent.FncPrint()
End Function


'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================

Function FncExcel() 
	Call parent.FncExport(parent.C_MULTI)
End Function


'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================

Function FncFind() 
    Call parent.FncFind(parent.C_MULTI , False)                                     '☜:화면 유형, Tab 유무 
End Function


'========================================================================================
' Function Name : FncSplitColumn
' Function Desc : 
'========================================================================================
Function FncSplitColumn()
    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Function
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)
End Function

'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================
Function FncExit()
	FncExit = True
End Function

'========================================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'==========================================================================================================
Function DbQuery(ByVal pOpt)
	Dim strVal
	Dim FromDt, ToDt

    DbQuery = False
    
    Err.Clear                
                                                   '☜: Protect system from crashing
	LayerShowHide(1)
		
    With frm1
		If .txtFromDt.Text = "" Then
			FromDt = UniConvYYYYMMDDToDate(parent.gDateFormat, "1900","01","01")
		Else
			FromDt = .txtFromDt.text
		End If
			
		If .txtToDt.Text = "" Then
			ToDt = UniConvYYYYMMDDToDate(parent.gDateFormat, "2999","12","31")
		Else
			ToDt = .txtToDt.text
		End If

        If pOpt = "A" Then
			strVal = BIZ_PGM_ID & "?txtPlantCd=" & Trim(.txtPlantCd.value)
			strVal = strVal & "&txtFromDt=" & Trim(FromDt)
			strVal = strVal & "&txtToDt=" & Trim(ToDt)
			strVal = strVal & "&txtItemCd=" & Trim(.txtItemCd.value)
			strVal = strVal & "&txtRoutNo=" & Trim(.txtRoutNo.value)
			strVal = strVal & "&iOpt=" & pOpt
        Else   
			strVal = BIZ_PGM_ID1 & "?txtPlantCd=" & Trim(.txtPlantCd.value)
			strVal = strVal & "&txtItemCd=" & GetKeyPosVal("A", 1)
			strVal = strVal & "&txtRoutNo=" & GetKeyPosVal("A", 2)
			strVal = strVal & "&iOpt=" & pOpt
        End If   
        
        If pOpt = "A" Then
           strVal = strVal & "&lgStrPrevKey="   & lgStrPrevKey_A                      '☜: Next key tag
           strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType(pOpt)
           strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList(pOpt)
           strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList(pOpt))
        Else   
           strVal = strVal & "&lgStrPrevKey="   & lgStrPrevKey_B                      '☜: Next key tag
           strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType(pOpt)
           strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList(pOpt)
           strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList(pOpt))
        End If   

        Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 
    End With

    DbQuery = True
End Function


'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================

Function DbQueryOk(ByVal pOpt)														'☆: 조회 성공후 실행로직 

	Call SetToolbar("11000000000111")
	Call ggoOper.LockField(Document, "Q")								 '⊙: This function lock the suitable field 
		
	lgBlnFlgChgValue = False
	
	If pOpt = "A" Then
		If lgTopLeft_A <> "Y" Then
			Call SetActiveCell(frm1.vspdData1,1,1,"M","X","X")
			Set gActiveElement = document.activeElement
			Call vspdData1_Click(1, 1)
		End If
		lgTopLeft_A = "N"
	End If								'⊙: This function lock the suitable field

End Function