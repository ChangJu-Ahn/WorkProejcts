'========================================
' External ASP File
'========================================
Const BIZ_PGM_ID        = "s5112qb1.asp"
Const BIZ_PGM_JUMP_ID	= "s5112ma1"

' Constant variables 
'========================================

Const C_MaxKey          = 9                                    '☆☆☆☆: Max key value
															   '☆: Jump시 Cookie로 보낼 Grid value
Const C_PopSoldToParty	= 1
Const C_PopBillType		= 2
Const C_PopItemCd		= 3
Const C_PopSalesGrp		= 4

Dim lgBlnOpenPop
Dim lgIntStartRow
'=========================================
Sub InitVariables()
    lgBlnFlgChgValue = False                               'Indicates that no value changed
    lgPageNo     = ""                                  'initializes Previous Key
    lgSortKey        = 1
End Sub

'=========================================
Sub SetDefaultVal()
	frm1.txtBillFrDt.text = UNIGetFirstDay(EndDate, Parent.gDateFormat)
	frm1.txtBillToDt.text = EndDate
	frm1.txtSoldToParty.focus
End Sub

'==========================================
Sub InitSpreadSheet()
   
	Call SetZAdoSpreadSheet("S5112QA1","S","A","V20030714", Parent.C_SORT_DBAGENT, frm1.vspdData, C_MaxKey, "X", "X" )
	
	Call SetSpreadLock 
	
End Sub

'=========================================
Sub SetSpreadLock()
	ggoSpread.SpreadLockWithOddEvenRowColor()
'	frm1.vspdData.OperationMode = 3
End Sub

'========================================
Sub FormatField()
	Call FormatDATEField(frm1.txtBillFrDt)
	Call FormatDATEField(frm1.txtBillToDt)
End Sub
'=========================================
Sub LockFieldInit()
	Call LockObjectField(frm1.txtBillFrDt, "O")
	Call LockObjectField(frm1.txtBillToDt, "O")
End Sub
'========================================

'========================================
Function OpenConPopup(ByVal pvIntWhere)
	Dim iArrRet
	Dim iArrParam(5), iArrField(6), iArrHeader(6)

	OpenConPopup = False
	
	If lgBlnOpenPop = True Then Exit Function

	lgBlnOpenPop = True

	With frm1
		Select Case pvIntWhere
			' 주문처 
			Case C_PopSoldToParty												
				iArrParam(1) = "dbo.b_biz_partner BP"				' TABLE 명칭 
				iArrParam(2) = Trim(.txtSoldToParty.value)		' Code Condition
				iArrParam(3) = ""									' Name Cindition
'				iArrParam(4) = "BP.bp_type IN (" & FilterVar("C", "''", "S") & " , " & FilterVar("CS", "''", "S") & ") AND BP.usage_flag = " & FilterVar("Y", "''", "S") & " "		' Where Condition
				iArrParam(4) = "BP.bp_type IN (" & FilterVar("C", "''", "S") & " , " & FilterVar("CS", "''", "S") & ") "		' Where Condition
					
				iArrField(0) = "ED15" & Parent.gColSep & "BP.bp_cd"	' Select Column
				iArrField(1) = "ED30" & Parent.gColSep & "BP.bp_nm"
				    
				iArrHeader(0) = .txtSoldtoParty.Alt							' Spread Title명 
				iArrHeader(1) = .txtSoldtoPartyNm.Alt
	
				.txtSoldToParty.focus
				
			' 매출채권형태 
			Case C_PopBillType												
				iArrParam(1) = "s_bill_type_config"
				iArrParam(2) = Trim(.txtBillType.value)
				iArrParam(3) = ""
				iArrParam(4) = "USAGE_FLAG=" & FilterVar("Y", "''", "S") & " "

				iArrField(0) = "ED15" & Parent.gColSep & "bill_type"
				iArrField(1) = "ED30" & Parent.gColSep & "bill_type_nm"

				iArrHeader(0) = .txtBillType.Alt
				iArrHeader(1) = .txtBillTypeNm.Alt
				
				.txtBillType.focus

			' 품목 
			Case C_PopItemCd

				OpenConPopup = OpenConItemPopup(C_PopItemCd, .txtItemCd.value)
				.txtItemCd.focus
				Exit Function


				iArrParam(1) = "b_item"
				iArrParam(2) = Trim(.txtItemCd.value)
				iArrParam(3) = ""
				iArrParam(4) = ""

				iArrField(0) = "ED15" & Parent.gColSep & "Item_Cd"
				iArrField(1) = "ED30" & Parent.gColSep & "Item_Nm"
				iArrField(2) = "ED30" & Parent.gColSep & "spec"

				iArrHeader(0) = .txtItemCd.Alt
				iArrHeader(1) = .txtItemNm.Alt
				iArrHeader(2) = "품목규격"
				
				.txtItemCd.focus
				
			' 영업그룹 
			Case C_PopSalesGrp												
				iArrParam(1) = "dbo.B_SALES_GRP"
				iArrParam(2) = Trim(.txtSalesGrp.value)
				iArrParam(3) = ""
				iArrParam(4) = "USAGE_FLAG=" & FilterVar("Y", "''", "S") & " "
				
				iArrField(0) = "ED15" & Parent.gColSep & "SALES_GRP"
				iArrField(1) = "ED30" & Parent.gColSep & "SALES_GRP_NM"
    
			    iArrHeader(0) = .txtSalesGrp.Alt
			    iArrHeader(1) = .txtSalesGrpNm.Alt
			    
			    .txtSalesGrp.focus

		End Select
	End With
 
	iArrParam(0) = iArrHeader(0)							' 팝업 Title
	iArrParam(5) = iArrHeader(0)							' 조회조건 명칭 

	iArrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(iArrParam, iArrField, iArrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgBlnOpenPop = False

	If iArrRet(0) <> "" Then
		OpenConPopup = SetConPopup(iArrRet,pvIntWhere)
	End If	
	
End Function

' Item Popup
'========================================
Function OpenConItemPopup(ByVal pvIntWhere, ByVal pvStrData)
	Dim iArrRet
	Dim iArrParam(3)
	Dim iCalledAspName

	OpenConItemPopup = False

	iCalledAspName = AskPRAspName("s2210pa1")
	
	If Trim(iCalledAspName) = "" Then
		Call DisplayMsgBox("900040", parent.VB_INFORMATION, "s2210pa1", "X")
		lgBlnOpenPop = False
		Exit Function
	End If

	iArrParam(0) = pvStrData
	
	iArrRet = window.showModalDialog(iCalledAspName, Array(window.parent,iArrParam), _
	 "dialogWidth=850px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgBlnOpenPop = False

	If iArrRet(0) <> "" Then
		OpenConItemPopup = SetConPopup(iArrRet,pvIntWhere)
	End If	
End Function

'========================================
Sub PopZAdoConfigGrid()
	If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
		Exit Sub
	End If
	
	Call OpenOrderByPopup("A")
End Sub

'========================================
Function OpenOrderByPopup(ByVal pSpdNo)
	Dim arrRet
	
	On Error Resume Next 
	
	If lgBlnOpenPop = True Then Exit Function
	lgBlnOpenPop = True

	arrRet = window.showModalDialog("../../ComAsp/ZADOSortPopup.asp",Array(ggoSpread.GetXMLData("A"),gMethodText),"dialogWidth=" & parent.SORTW_WIDTH & "px; dialogHeight=" & parent.SORTW_HEIGHT & "px;; center: Yes; help: No; resizable: No; status: No;")

	lgBlnOpenPop = False
	
	If arrRet(0) = "X" Then
	   Exit Function
	Else
	   Call ggoSpread.SaveXMLData(pSpdNo,arrRet(0),arrRet(1))
       Call InitVariables
       Call InitSpreadSheet()       
   End If
End Function

'========================================
Function SetConPopup(Byval pvArrRet,ByVal pvIntWhere)

	SetConPopup = False

	With frm1
		Select Case pvIntWhere
			Case C_PopSoldToParty
				.txtSoldToParty.value = pvArrRet(0) 
				.txtSoldToPartyNm.value = pvArrRet(1)   

			Case C_PopBillType
				.txtBillType.value = pvArrRet(0) 
				.txtBillTypeNm.value = pvArrRet(1)   

			Case C_PopItemCd
				.txtItemCd.value = pvArrRet(0) 
				.txtItemNm.value = pvArrRet(1)   
				
			Case C_PopSalesGrp
				.txtSalesGrp.value = pvArrRet(0) 
				.txtSalesGrpNm.value = pvArrRet(1)
		End Select
	End With
	
	SetConPopup = True

End Function

'========================================
Function CookiePage(ByVal Kubun)
	Dim strTemp, arrVal

	Const CookieSplit = 4877

	If Kubun = 1 Then
		If frm1.vspdData.ActiveRow > 0 Then
			frm1.vspdData.Row = frm1.vspdData.ActiveRow
			frm1.vspdData.Col = GetKeyPos("A",1)
			WriteCookie CookieSplit , frm1.vspdData.Text
		Else
			WriteCookie CookieSplit , ""
		End If

		Call PgmJump(BIZ_PGM_JUMP_ID)

	ElseIf Kubun = 0 Then

		strTemp = ReadCookie(CookieSplit)

		If strTemp = "" then Exit Function

		arrVal = Split(strTemp, Parent.gRowSep)

		If arrVal(0) = "" Then 
			WriteCookie CookieSplit , ""
			Exit Function
		End If

		frm1.txtSoldToParty.value =  arrVal(0)
		frm1.txtSoldToPartyNm.value =  arrVal(1)
		frm1.txtBillType.value =  arrVal(2)
		frm1.txtBillTypeNm.value = arrVal(3) 
		frm1.txtSalesGrp.value =  arrVal(6)
		frm1.txtSalesGrpNm.value = arrVal(7) 
		frm1.txtItemCd.value =  arrVal(8)
		frm1.txtItemNm.value = arrVal(9)

		If Err.number <> 0 Then
			Err.Clear
			WriteCookie CookieSplit , ""
			Exit Function 
		End If

		Call MainQuery()

		WriteCookie CookieSplit , ""

	End IF

End Function

'=========================================
Sub Form_Load()
    Call LoadInfTB19029														'⊙: Load table , B_numeric_format
'	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
'   Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field
 	Call FormatField()
	Call LockFieldInit()
   
    Call InitVariables														'⊙: Initializes local global variables
	Call SetDefaultVal	
	Call InitSpreadSheet()
    Call SetToolbar("11000000000011")							'⊙: 버튼 툴바 제어 
	Call CookiePage(0)
End Sub

'==========================================
Sub Form_QueryUnload(Cancel , UnloadMode )
End Sub

'==========================================
Sub vspdData_Click(ByVal Col, ByVal Row)

    Call SetPopupMenuItemInf("00000000001")

    gMouseClickStatus = "SPC"
    Set gActiveSpdSheet = frm1.vspdData
 
	frm1.vspdData.ReDraw = False
    If Row = 0 Then
'		frm1.vspdData.OperationMode = 0

        If lgSortKey = 1 Then
            ggoSpread.SSSort Col, lgSortKey
            lgSortKey = 2
        Else
            ggoSpread.SSSort ,lgSortKey
            lgSortKey = 1
        End If    
	Else
'		frm1.vspdData.OperationMode = 3
    End If
End Sub

'==========================================
Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If

End Sub

'==========================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
	If OldLeft <> NewLeft Then Exit Sub

	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	    '☜: 재쿼리 체크	
		If lgPageNo <> "" Then		                                                    '다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
			Call DisableToolBar(Parent.TBC_QUERY)
			Call DbQuery
		End If
	End If
End Sub

'==========================================
Sub rdoTexIssueFlg1_OnClick()
	frm1.txtRadio.value = frm1.rdoTexIssueFlg1.value
End Sub

'==========================================
Sub rdoTexIssueFlg2_OnClick()
	frm1.txtRadio.value = frm1.rdoTexIssueFlg2.value
End Sub

'==========================================
Sub rdoTexIssueFlg3_OnClick()
	frm1.txtRadio.value = frm1.rdoTexIssueFlg3.value
End Sub
	
'==========================================
Sub txtBillFrDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtBillFrDt.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtBillFrDt.Focus
	End If
End Sub

'==========================================
Sub txtBillToDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtBillToDt.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtBillToDt.Focus
	End If
End Sub

'==========================================
Sub txtBillFrDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub

'==========================================
Sub txtBillToDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub

'==========================================
Function FncQuery() 

    FncQuery = False                                                        '⊙: Processing is NG
    
    Err.Clear                                                               '☜: Protect system from crashing

	If ValidDateCheck(frm1.txtBillFrDt, frm1.txtBillToDt) = False Then Exit Function
     lgIntFlgMode     = Parent.OPMD_CMODE 

	ggoSpread.Source = frm1.vspdData
	ggoSpread.ClearSpreadData    

    Call InitVariables 														'⊙: Initializes local global variables
    
    Call DbQuery															'☜: Query db data

    FncQuery = True		
End Function

'========================================
Function FncPrint() 
    Call parent.FncPrint()
End Function

'========================================
Function FncExcel() 
	Call parent.FncExport(Parent.C_MULTI)
End Function

'========================================
Function FncFind() 
    Call parent.FncFind(Parent.C_MULTI , False)                                     
End Function

'========================================
Function FncSplitColumn()
    
    Dim ACol
    Dim ARow
    Dim iRet
    Dim iColumnLimit
    
    iColumnLimit  = frm1.vspdData.MaxCols
    
    ACol = Frm1.vspdData.ActiveCol
    ARow = Frm1.vspdData.ActiveRow

    If ACol > iColumnLimit Then
       iRet = DisplayMsgBox("900030", "X", iColumnLimit , "X")
       Exit Function
    End If   
    
    Frm1.vspdData.ScrollBars = Parent.SS_SCROLLBAR_NONE
    
    ggoSpread.Source = Frm1.vspdData
    
    ggoSpread.SSSetSplit(ACol)    
    
    Frm1.vspdData.Col = ACol
    Frm1.vspdData.Row = ARow
    
    Frm1.vspdData.Action = 0    
    
    Frm1.vspdData.ScrollBars = Parent.SS_SCROLLBAR_BOTH
    
End Function

'========================================
Function FncExit()
    FncExit = True
End Function

'========================================
Function DbQuery() 
	Dim strVal

    DbQuery = False
    
    Err.Clear                                                               '☜: Protect system from crashing
			
	If LayerShowHide(1) = False Then Exit Function 
    
    With frm1
		strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001	
	
		If lgIntFlgMode = Parent.OPMD_UMODE Then
			' Scroll시 
			strVal = strVal & "&txtSoldToParty=" & Trim(.txtHSoldToParty.value)
			strVal = strVal & "&txtBillType=" & Trim(.txtHBillType.value)
			strVal = strVal & "&txtSalesGrp=" & Trim(.txtHSalesGrp.value)
			strVal = strVal & "&txtItemCd=" & Trim(.txtHItemCd.value)
			strVal = strVal & "&txtBillFrDt=" & Trim(.txtHBillFrDt.value)
			strVal = strVal & "&txtBillToDt=" & Trim(.txtHBillToDt.value)
			strVal = strVal & "&txtRadio=" & Trim(.txtHRadio.value)	
 
		Else
			' 처음 조회시 
			strVal = strVal & "&txtSoldToParty=" & Trim(.txtSoldToParty.value)
			strVal = strVal & "&txtBillType=" & Trim(.txtBillType.value)
			strVal = strVal & "&txtSalesGrp=" & Trim(.txtSalesGrp.value)
			strVal = strVal & "&txtItemCd=" & Trim(.txtItemCd.value)
			strVal = strVal & "&txtBillFrDt=" & Trim(.txtBillFrDt.text)
			strVal = strVal & "&txtBillToDt=" & Trim(.txtBillToDt.text)
			strVal = strVal & "&txtRadio=" & Trim(.txtRadio.value)
		End If
	
        strVal = strVal & "&lgPageNo="   & lgPageNo                      '☜: Next key tag
		strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")
        strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
		strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))
		
		lgIntStartRow = .vspdData.MaxRows + 1
		
        Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 
    End With
    
    DbQuery = True
End Function

'========================================
Function DbQueryOk()
	Call SetToolbar("11000000000111")
	If frm1.vspdData.MaxRows > 0 Then
		frm1.vspdData.Focus
		frm1.vspdData.SelModeSelected = True
		If lgIntFlgMode <> Parent.OPMD_UMODE Then
			frm1.vspdData.Row = 1
			lgIntFlgMode = Parent.OPMD_UMODE
		End If
		Call FormatSpreadCellByCurrency()
    Else
       frm1.txtSoldToParty.focus
    End If   
End Function

'========================================
' 화폐별로 Cell Formating을 재설정한다.
Sub FormatSpreadCellByCurrency()
	With frm1
		Call ReFormatSpreadCellByCellByCurrency(.vspdData,lgIntStartRow,.vspdData.MaxRows,GetKeyPos("A",2),GetKeyPos("A",3),"C","Q","X","X")
		Call ReFormatSpreadCellByCellByCurrency(.vspdData,lgIntStartRow,.vspdData.MaxRows,GetKeyPos("A",2),GetKeyPos("A",4),"A","Q","X","X")
		Call ReFormatSpreadCellByCellByCurrency(.vspdData,lgIntStartRow,.vspdData.MaxRows,GetKeyPos("A",2),GetKeyPos("A",5),"A","Q","X","X")
		Call ReFormatSpreadCellByCellByCurrency(.vspdData,lgIntStartRow,.vspdData.MaxRows,GetKeyPos("A",2),GetKeyPos("A",6),"A","Q","X","X")
		Call ReFormatSpreadCellByCellByCurrency2(.vspdData,lgIntStartRow,.vspdData.MaxRows,parent.gCurrency,GetKeyPos("A",7),"A","Q","X","X")
		Call ReFormatSpreadCellByCellByCurrency2(.vspdData,lgIntStartRow,.vspdData.MaxRows,parent.gCurrency,GetKeyPos("A",8),"A","Q","X","X")
		Call ReFormatSpreadCellByCellByCurrency2(.vspdData,lgIntStartRow,.vspdData.MaxRows,parent.gCurrency,GetKeyPos("A",9),"A","Q","X","X")
	End With
End Sub
