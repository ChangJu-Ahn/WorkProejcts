' External ASP File
'========================================
Const BIZ_PGM_ID 		= "s3113MB9.asp"					                       '☆: Biz Logic ASP Name

' Constant variables 
'========================================
Const C_MaxKey          = 1                                           '☆: key count of SpreadSheet

Const C_PopSoldToParty	= 1
Const C_PopSalesGrp		= 2
Const C_PopSoNo			= 3
Const C_PopPlantCd		= 4
Const C_PopItemCd		= 5


' User-defind Variables
'========================================
Dim IsOpenPop  
Dim gblnWinEvent											'☜: ShowModal Dialog(PopUp) 
														    'Window가 여러 개 뜨는 것을 방지하기 위해 
														    'PopUp Window가 사용중인지 여부를 나타냄 
Dim lgBlnOpenedFlag											' 화면 Load완료여부 

Dim	lgBlnSoldToPartyChg
Dim lgBlnSalesGrpChg
Dim	lgBlnPlantCdChg
Dim	lgBlnItemCdChg


'========================================
Function InitVariables()
	lgPageNo         = ""
    lgIntFlgMode     = Parent.OPMD_CMODE                          'Indicates that current mode is Create mode
    gblnWinEvent = False
    lgBlnFlgChgValue = False	                           'Indicates that no value changed
    lgSortKey        = 1   

    lgStrPrevKey = ""										'initializes Previous Key

	lgBlnSoldToPartyChg = False
	lgBlnSalesGrpChg	= False
	lgBlnPlantCdChg		= False
	lgBlnItemCdChg		= False
End Function

'==========================================
Sub SetDefaultVal()
	With frm1
		.txtFromDt.Text = UNIGetFirstDay(EndDate, Parent.gDateFormat)
		.txtToDt.Text = EndDate	

		If Parent.gPlant <> "" Then
			.txtPlantCd.value = Parent.gPlant
			Call GetPlantNm()
		End If
		
		If Parent.gSalesGrp <> "" Then
			.txtSalesGrp.value = Parent.gSalesGrp
			Call GetSalesGrpNm()
		End If

		.txtFromDt.Focus
	End With
	lgBlnFlgChgValue = False
End Sub

'========================================
Sub InitSpreadSheet()
	
	Call SetZAdoSpreadSheet("S3113MA9","S","A","V20021107", Parent.C_SORT_DBAGENT, frm1.vspdData, C_MaxKey, "X", "X" )
	
	Call SetSpreadLock  
	    
End Sub

'========================================
Sub SetSpreadLock()
	ggoSpread.SpreadLockWithOddEvenRowColor()
'	frm1.vspdData.OperationMode = 3
End Sub	
'========================================
Sub FormatField()
	Call FormatDATEField(frm1.txtFromDt)
	Call FormatDATEField(frm1.txtToDt)
End Sub
'=========================================
Sub LockFieldInit()
	Call LockObjectField(frm1.txtFromDt, "O")
	Call LockObjectField(frm1.txtToDt, "O")
End Sub
'========================================

'========================================
Function OpenSoNo()
	
	Dim iCalledAspName
	Dim strRet

	If IsOpenPop = True Then Exit Function
			
	IsOpenPop = True
		
	'20021228 kangjungu dynamic popup
	iCalledAspName = AskPRAspName("s3111pa1")	
	if Trim(iCalledAspName) = "" then
		IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION, "s3111pa1", "x")
		IsOpenPop = False
		exit Function
	end if

	strRet = window.showModalDialog(iCalledAspName, Array(window.parent, ""), _
		"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	frm1.txtSoNo.focus 
	If strRet = "" Then
		Exit Function
	Else
		frm1.txtSoNo.value = strRet 
	End If	

End Function

'========================================
Function OpenConPopup(ByVal pvIntWhere)

	Dim iArrRet
	Dim iArrParam(5), iArrField(6), iArrHeader(6)

	OpenConPopup = False
	
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True


	Select Case pvIntWhere
	Case C_PopSoldToParty												
		iArrParam(1) = "B_BIZ_PARTNER PARTNER"			' TABLE 명칭 
		iArrParam(2) = Trim(frm1.txtSoldToParty.value)	' Code Condition
		iArrParam(3) = ""								' Name Cindition
'		iArrParam(4) = "PARTNER.USAGE_FLAG=" & FilterVar("Y", "''", "S") & "  AND PARTNER.BP_TYPE IN (" & FilterVar("CS", "''", "S") & ", " & FilterVar("C", "''", "S") & " )"	' Where Condition
		iArrParam(4) = "PARTNER.BP_TYPE IN (" & FilterVar("CS", "''", "S") & ", " & FilterVar("C", "''", "S") & " )"	' Where Condition
		iArrParam(5) = "주문처"						' TextBox 명칭 
			
		iArrField(0) = "PARTNER.BP_CD"					' Field명(0)
		iArrField(1) = "PARTNER.BP_NM"					' Field명(1)
		    
		iArrHeader(0) = "주문처"					' Header명(0)
		iArrHeader(1) = "주문처명"					' Header명(1)
		
		frm1.txtSoldToParty.focus

	Case C_PopSalesGrp												
		iArrParam(1) = "B_SALES_GRP"
		iArrParam(2) = Trim(frm1.txtSalesGrp.value)
		iArrParam(3) = ""
		iArrParam(4) = "USAGE_FLAG=" & FilterVar("Y", "''", "S") & " "
		iArrParam(5) = "영업그룹"
		
		iArrField(0) = "SALES_GRP"
		iArrField(1) = "SALES_GRP_NM"
    
	    iArrHeader(0) = "영업그룹"
	    iArrHeader(1) = "영업그룹명"
	    
	    frm1.txtSalesGrp.focus

	Case C_PopPlantCd
		iArrParam(1) = "b_plant A"
		iArrParam(2) = Trim(frm1.txtPlantCd.value)
		iArrParam(3) = ""
		iArrParam(4) = ""
		iArrParam(5) = "공장"
	
		iArrField(0) = "ED15" & Parent.gColSep & "A.plant_cd"
		iArrField(1) = "ED30" & Parent.gColSep & "A.plant_nm"
		    
		iArrHeader(0) = "공장"
		iArrHeader(1) = "공장명"
		
		frm1.txtPlantCd.focus

	Case C_PopItemCd
		iArrParam(1) = "b_item A"
		iArrParam(2) = Trim(frm1.txtItemCd.value)
		iArrParam(3) = ""
		
		If Trim(frm1.txtPlantCd.value) <> "" Then
			iArrParam(4) = "Exists (SELECT * " & _
								"	FROM b_item_by_plant B " & _
								"	WHERE A.item_cd = B.item_cd " & _
								"	AND B.plant_cd =  " & FilterVar(frm1.txtPlantCd.value, "''", "S") & ")"
		Else
			iArrParam(4) = ""
		End If
		iArrParam(5) = "품목"

		iArrField(0) = "ED15" & Parent.gColSep & "A.item_cd"
		iArrField(1) = "ED30" & Parent.gColSep & "A.item_nm"
		iArrField(2) = "ED30" & Parent.gColSep & "A.spec"

		iArrHeader(0) = "품목"
		iArrHeader(1) = "품목명"
		iArrHeader(2) = "품목규격"
		
		frm1.txtItemCd.focus

	End Select
 
	iArrParam(0) = iArrParam(5)							

	iArrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(iArrParam, iArrField, iArrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	
	If iArrRet(0) <> "" Then
		Call SetConPopup(iArrRet,pvIntWhere)
		OpenConPopup = True
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
	
	If IsOpenPop = True Then Exit Function
	IsOpenPop = True

	arrRet = window.showModalDialog("../../ComAsp/ZADOSortPopup.asp",Array(ggoSpread.GetXMLData("A"),gMethodText),"dialogWidth=" & parent.SORTW_WIDTH & "px; dialogHeight=" & parent.SORTW_HEIGHT & "px;; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	
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

	Select Case pvIntWhere
	Case C_PopSoldToParty
		frm1.txtSoldToParty.value = pvArrRet(0) 
		frm1.txtSoldToPartyNm.value = pvArrRet(1)   
	Case C_PopSalesGrp
		frm1.txtSalesGrp.value = pvArrRet(0) 
		frm1.txtSalesGrpNm.value = pvArrRet(1)   
	Case C_PopPlantCd
		frm1.txtPlantCd.value = pvArrRet(0) 
		frm1.txtPlantNm.value = pvArrRet(1)   
	Case C_PopItemCd
		frm1.txtItemCd.value = pvArrRet(0) 
		frm1.txtItemNm.value = pvArrRet(1)   
	End Select

	SetConPopup = True

End Function

'========================================
Sub Form_Load()
    Call LoadInfTB19029											  '⊙: Load table , B_numeric_format
   
    'Html에서 tag 숫자가 1과 2로 시작하는 부분 각각Format
'   Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
'	Call ggoOper.LockField(Document, "N")                         '⊙: Lock  Suitable  Field
 	Call FormatField()
	Call LockFieldInit()
    
	Call InitVariables											  '⊙: Initializes local global variables
	Call SetDefaultVal	
	Call InitSpreadSheet()
	Call SetToolbar("11000000000011")										'⊙: 버튼 툴바 제어 
End Sub

'========================================
Sub Form_QueryUnload(Cancel, UnloadMode)
End Sub

'==========================================
Function GetSalesGrpNm()
	Dim iStrCode
	
	iStrCode = Trim(frm1.txtSalesGrp.value)
	If iStrCode <> "" Then
		iStrCode = " " & FilterVar(iStrCode, "''", "S") & ""
		If Not GetCodeName(iStrCode, "default", "default", "default", "default", "" & FilterVar("SG", "''", "S") & "", C_PopSalesGrp) Then
			frm1.txtSalesGrp.value = ""
			frm1.txtSalesGrpNm.value = ""
		End If
	Else
		frm1.txtSalesGrpNm.value = ""
	End If
End Function

'==========================================
Function GetPlantNm
	Dim iStrCode
	
	iStrCode = Trim(frm1.txtPlantCd.value)
	If iStrCode <> "" Then
		iStrCode = " " & FilterVar(iStrCode, "''", "S") & ""
		If Not GetCodeName(iStrCode, "default", "default", "default", "default", "" & FilterVar("PT", "''", "S") & "", C_PopPlantCd) Then
			frm1.txtPlantCd.value = ""
			frm1.txtPlantNm.value = ""
		End If
	Else
		frm1.txtPlantNm.value = ""
	End If
End Function

'==========================================
Function txtSoldToParty_OnKeyDown()
	lgBlnFlgChgValue = True
	lgBlnSoldToPartyChg = True
End Function

'==========================================
Function txtSalesGrp_OnKeyDown()
	lgBlnFlgChgValue = True
	lgBlnSalesGrpChg = True
End Function

'==========================================
Function txtPlantCd_OnKeyDown()
	lgBlnFlgChgValue = True
	lgBlnPlantCdChg = True
End Function

'==========================================
Function txtItemCd_OnKeyDown()
	lgBlnFlgChgValue = True
	lgBlnItemCdChg = True
End Function

'	Description : 조회조건의 유효성을 Check한다.
'   주의사항 : 화면의 tab order 별로 기술한다. 
'====================================================
Function ChkValidityQueryCon()
	Dim iStrCode, iStrPlantCd

	ChkValidityQueryCon = True

	If lgBlnSoldToPartyChg Then
		iStrCode = Trim(frm1.txtSoldToParty.value)
		If iStrCode <> "" Then
			iStrCode = " " & FilterVar(iStrCode, "''", "S") & ""
			If Not GetCodeName(iStrCode, "" & FilterVar("C%", "''", "S") & "", "default", "default", "default", "" & FilterVar("BP", "''", "S") & "", C_PopSoldToParty) Then
				Call DisplayMsgBox("970000", "X", frm1.txtSoldtoparty.alt, "X")
				frm1.txtSoldToPartyNm.value = ""
				frm1.txtSoldtoparty.focus
				ChkValidityQueryCon = False
				Exit Function
			End If
		Else
			frm1.txtSoldToPartyNm.value = ""
		End If
		lgBlnSoldToPartyChg	= False
	End If

	If lgBlnSalesGrpChg Then
		iStrCode = Trim(frm1.txtSalesGrp.value)
		If iStrCode <> "" Then
			iStrCode = " " & FilterVar(iStrCode, "''", "S") & ""
			If Not GetCodeName(iStrCode, "default", "default", "default", "default", "" & FilterVar("SG", "''", "S") & "", C_PopSalesGrp) Then
				Call DisplayMsgBox("970000", "X", frm1.txtSalesGrp.alt, "X")
				frm1.txtSalesGrpNm.value = ""
				frm1.txtSalesGrp.focus
				ChkValidityQueryCon = False
				Exit Function
			End If
		Else
			frm1.txtSalesGrpNm.value = ""
		End If
		lgBlnSalesGrpChg = False
	End If
			
	If lgBlnPlantCdChg Then
		iStrCode = Trim(frm1.txtPlantCd.value)
		If iStrCode <> "" Then
			iStrCode = " " & FilterVar(iStrCode, "''", "S") & ""
			If Not GetCodeName(iStrCode, "default", "default", "default", "default", "" & FilterVar("PT", "''", "S") & "", C_PopPlantCd) Then
				Call DisplayMsgBox("970000", "X", frm1.txtPlantCd.alt, "X")
				frm1.txtPlantNm.value = ""
				frm1.txtPlantCd.focus
				ChkValidityQueryCon = False
				Exit Function
			End If
		Else
			frm1.txtPlantNm.value = ""
		End If
		lgBlnPlantCdChg = False
	End If

	If lgBlnItemCdChg Then

		iStrPlantCd = Trim(frm1.txtPlantCd.value)

		If iStrPlantCd <> "" Then 
			iStrPlantCd = " " & FilterVar(iStrPlantCd, "''", "S") & ""
		Else
			iStrPlantCd = "default"
		End If

		iStrCode = Trim(frm1.txtItemCd.value)
		If iStrCode <> "" Then
			iStrCode = " " & FilterVar(iStrCode, "''", "S") & ""
			If Not GetCodeName(iStrCode, iStrPlantCd, "default", "default", "default", "" & FilterVar("IT", "''", "S") & "", C_PopItemCd) Then
				Call DisplayMsgBox("970000", "X", frm1.txtItemCd.alt, "X")
				frm1.txtItemNm.value = ""
				frm1.txtItemCd.focus
				ChkValidityQueryCon = False
				Exit Function
			End If
		Else
			frm1.txtItemNm.value = ""
		End If
		lgBlnItemCdChg = False
	End If

End Function

'	Description : 코드값에 해당하는 명을 Display한다.
'====================================================
Function GetCodeName(ByVal pvStrArg1, ByVal pvStrArg2, ByVal pvStrArg3, ByVal pvStrArg4, ByVal pvIntArg5, ByVal pvStrFlag, ByVal pvIntWhere)

	Dim iStrSelectList, iStrFromList, iStrWhereList
	Dim iStrRs
	Dim iArrRs(2), iArrTemp
	
	GetCodeName = False
	
	iStrSelectList = " * "
	iStrFromList = " dbo.ufn_s_GetCodeName (" & pvStrArg1 & ", " & pvStrArg2 & ", " & pvStrArg3 & ", " & pvStrArg4 & ", " & pvIntArg5 & ", " & pvStrFlag & ") "
	iStrWhereList = ""
	
	Err.Clear
	
	If CommonQueryRs2by2(iStrSelectList,iStrFromList,iStrWhereList,iStrRs) Then
		iArrTemp = Split(iStrRs, Chr(11))
		iArrRs(0) = iArrTemp(1)
		iArrRs(1) = iArrTemp(2)
		iArrRs(2) = iArrTemp(3)
		GetCodeName = SetConPopup(iArrRs, pvIntWhere)
	Else
		' 관련 Popup Display
		'If lgBlnOpenedFlag Then	GetCodeName = OpenConPopup(pvIntWhere)
	End if
End Function

'==========================================
Sub vspdData_Click(ByVal Col , ByVal Row)

    Call SetPopupMenuItemInf("00000000001")

	gMouseClickStatus = "SPC"

    Set gActiveSpdSheet = frm1.vspdData
        
	frm1.vspdData.ReDraw = False
	
    If Row = 0 Then
'		frm1.vspdData.OperationMode = 0
		ggoSpread.Source = frm1.vspdData

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
	frm1.vspdData.ReDraw = True
  
End Sub

'==========================================
Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If

End Sub    

'==========================================
Sub vspdData_TopLeftChange(ByVal OldLeft, ByVal OldTop, ByVal NewLeft, ByVal NewTop)
	
	If OldLeft <> NewLeft Then Exit Sub

	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	    '☜: 재쿼리 체크	
		If lgPageNo <> "" Then
			Call DisableToolBar(Parent.TBC_QUERY)
			Call DbQuery
		End If
	End If

End Sub

'========================================
Sub vspdData_Keypress(KeyAscii)
	If KeyAscii = 13	Then Call MainQuery()
End Sub

'========================================
Sub txtFromDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtFromDt.Action = 7		
		Call SetFocusToDocument("M")
		frm1.txtFromDt.Focus
	End If
End Sub

'========================================
Sub txtToDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtToDt.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtToDt.Focus
	End If
End Sub

'==========================================
Sub txtFromDt_Keypress(KeyAscii)
	If KeyAscii = 13	Then Call MainQuery()
End Sub

'========================================
Sub txtToDt_Keypress(KeyAscii)
	If KeyAscii = 13	Then Call MainQuery()
End Sub

'========================================
Function FncQuery() 
    
    FncQuery = False                                                        '⊙: Processing is NG
    
    Err.Clear                                                               '☜: Protect system from crashing
	'** ValidDateCheck(pObjFromDt, pObjToDt) : 'pObjToDt'이 'pObjFromDt'보다 크거나 같아야 할때 **
	If ValidDateCheck(frm1.txtFromDt, frm1.txtToDt) = False Then Exit Function
	
	ggoSpread.Source = frm1.vspdData
	ggoSpread.ClearSpreadData
	
	' 조회조건 유효값 check
	If 	lgBlnFlgChgValue Then
		If Not ChkValidityQueryCon Then	Exit Function
	End If
	
    Call InitVariables 														'⊙: Initializes local global variables
    
	If DbQuery = False Then Exit Function									

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
    
    iColumnLimit  = C_SoldToPartyNm
   
    ACol = Frm1.vspdData.ActiveCol
    ARow = Frm1.vspdData.ActiveRow

    If ACol > iColumnLimit Then
		'◎ Frm1없으면 frm1삭제 
		Frm1.vspdData.Col = iColumnLimit	:	Frm1.vspdData.Row = 0
		iRet = DisplayMsgBox("900030", "X", Trim(frm1.vspdData.Text), "X")
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

	Err.Clear														'☜: Protect system from crashing
	DbQuery = False													'⊙: Processing is NG
	
	If LayerShowHide(1) = False Then
		Exit Function
	End If
	
	Dim strVal
	
    With frm1
		strVal = BIZ_PGM_ID & "?txtHMode=" & Parent.UID_M0001					
		If lgIntFlgMode = Parent.OPMD_UMODE Then
			' Scroll시 
			' strVal = strVal & "&txtBillType=" & Trim(.txtHPlantCd.value) 
			' 박정순 수정. 
			strVal = strVal & "&txtPlantCd=" & Trim(.txtPlantCd.value) 
			strVal = strVal & "&txtSalesGrp=" & Trim(.txtHSalesGrp.value)
			strVal = strVal & "&txtSoldToParty=" & Trim(.txtHSoldToParty.value)
			strVal = strVal & "&txtFromDt=" & Trim(.txtHFromDt.value)
			strVal = strVal & "&txtToDt=" & Trim(.txtHToDt.value)
			strVal = strVal & "&txtSoNo=" & Trim(.txtHSoNo.value)
			strVal = strVal & "&txtItemCd=" & Trim(.txtHItemCd.value)
		Else
			' 처음 조회시 
			strVal = strVal & "&txtPlantCd=" & Trim(.txtPlantCd.value)
			strVal = strVal & "&txtSalesGrp=" & Trim(.txtSalesGrp.value)
			strVal = strVal & "&txtSoldToParty=" & Trim(.txtSoldToParty.value)
			strVal = strVal & "&txtFromDt=" & Trim(.txtFromDt.Text)
			strVal = strVal & "&txtToDt=" & Trim(.txtToDt.Text)
			strVal = strVal & "&txtSoNo=" & Trim(.txtSoNo.value)
			strVal = strVal & "&txtItemCd=" & Trim(.txtItemCd.value)
		End If

        strVal = strVal & "&lgPageNo="		 & lgPageNo					'☜: Next key tag 
		strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")
        strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
		strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))
	End With    
    
	Call RunMyBizASP(MyBizASP, strVal)									
    DbQuery = True    

End Function

'=========================================
Function DbQueryOk()	    												'☆: 조회 성공후 실행로직 

	If frm1.vspdData.MaxRows > 0 Then
		frm1.vspdData.Focus
		frm1.vspdData.SelModeSelected = True
		If lgIntFlgMode <> Parent.OPMD_UMODE Then
			frm1.vspdData.Row = 1
		End If
		lgIntFlgMode = Parent.OPMD_UMODE
	Else
		Call SetFocusToDocument("M")
		frm1.txtFromDt.focus
	End If

End Function

