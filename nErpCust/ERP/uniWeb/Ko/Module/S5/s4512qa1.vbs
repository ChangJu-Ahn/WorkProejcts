Const BIZ_PGM_ID        = "s4512qb1.asp"
Const C_MaxKey          = 1                                    '☆☆☆☆: Max key value

'=========================================
Sub InitVariables()
    lgStrPrevKey     = ""
    lgPageNo         = ""
    lgIntFlgMode     = parent.OPMD_CMODE
    lgBlnFlgChgValue = False                               
    lgSortKey        = 1
End Sub

'=========================================
Sub SetDefaultVal()
	frm1.txtPromiseFrDt.text = StartDate
	frm1.txtPromiseToDt.text = EndDate
End Sub
'=========================================
Sub InitSpreadSheet()
	Call SetZAdoSpreadSheet("S4512QA1","S","A","V20070119", Parent.C_SORT_DBAGENT, frm1.vspdData, C_MaxKey, "X", "X" )
    Call SetSpreadLock 
End Sub

'=========================================
Sub SetSpreadLock()
	ggoSpread.SpreadLockWithOddEvenRowColor()
End Sub
'========================================
Sub FormatField()
	Call FormatDATEField(frm1.txtPromiseFrDt)
	Call FormatDATEField(frm1.txtPromiseToDt)
End Sub
'=========================================
Sub LockFieldInit()
	Call LockObjectField(frm1.txtPromiseFrDt, "O")
	Call LockObjectField(frm1.txtPromiseToDt, "O")
End Sub
'========================================
Function OpenConSItemDC(Byval iWhere)

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	Select Case iWhere
	Case 0
		arrParam(1) = "B_BIZ_PARTNER"						
		arrParam(2) = Trim(frm1.txtconBp_cd.Value)			
		'arrParam(3) = Trim(frm1.txtconBp_Nm.Value)			
		arrParam(4) = "BP_TYPE in (" & FilterVar("C", "''", "S") & " ," & FilterVar("CS", "''", "S") & ")"				
		arrParam(5) = "납품처"							
	
		arrField(0) = "BP_CD"								
		arrField(1) = "BP_NM"								
    
		arrHeader(0) = "납품처"							
		arrHeader(1) = "납품처명"						

		frm1.txtconBp_cd.focus
		
	Case 1
		arrParam(1) = "B_MINOR A, I_MOVETYPE_CONFIGURATION B"				
		arrParam(2) = Trim(frm1.txtDnType.value)
		arrParam(3) = ""
		arrParam(4) = "A.MINOR_CD=B.MOV_TYPE AND (B.TRNS_TYPE = " & FilterVar("DI", "''", "S") & " OR (B.TRNS_TYPE = " & FilterVar("ST", "''", "S") & " AND B.STCK_TYPE_FLAG_DEST = " & FilterVar("T", "''", "S") & " )) AND A.MAJOR_CD=" & FilterVar("I0001", "''", "S") & " "	
		arrParam(5) = "출하형태"
		arrField(0) = "A.MINOR_CD"
		arrField(1) = "A.MINOR_NM"
		arrHeader(0) = "출하형태"					
		arrHeader(1) = "출하형태명"					

		frm1.txtDNType.focus
		
	Case 3
		arrParam(1) = "B_SALES_GRP"							
		arrParam(2) = Trim(frm1.txtSalesGroup.Value)		
		arrParam(4) = ""									
		arrParam(5) = "영업그룹"						
	
		arrField(0) = "SALES_GRP"							
		arrField(1) = "SALES_GRP_NM"							
    
		arrHeader(0) = "영업그룹"						
		arrHeader(1) = "영업그룹명"		
		
		frm1.txtSalesGroup.focus					

	Case 4
		arrParam(1) = "B_PLANT"								
		arrParam(2) = Trim(frm1.txtPlant.value)				
		'arrParam(3) = Trim(frm1.txtPlantNm.value)			
		arrParam(4) = ""									
		arrParam(5) = "공장"							
	
		arrField(0) = "PLANT_CD"							
		arrField(1) = "PLANT_NM"							
    
		arrHeader(0) = "공장"							
		arrHeader(1) = "공장명"
		
		frm1.txtPlant.focus							
	
	Case 5
		arrParam(1) = "B_STORAGE_LOCATION"					
		arrParam(2) = Trim(frm1.txtStoRo_cd.Value)			
		'arrParam(3) = Trim(frm1.txtStoRo_Nm.Value)			
		arrParam(4) = ""									
		arrParam(5) = "창고"							
	
		arrField(0) = "SL_CD"								
		arrField(1) = "SL_NM"								
    
		arrHeader(0) = "창고"							
		arrHeader(1) = "창고명"				
		
		frm1.txtStoRo_cd.focus		
		
	Case 6
		Dim strRet
		
		Dim arrTNParam(5), i

		For i = 0 to UBound(arrTNParam)
			arrTNParam(i) = ""
		Next	

		'20021227 kangjungu dynamic popup
		iCalledAspName = AskPRAspName("s3135pa1")	
		if Trim(iCalledAspName) = "" then
			IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION, "s3135pa1", "x")
			lgIsOpenPop = False
			exit Function
		end if

		strRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrTNParam), _
			"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

		lgIsOpenPop = False

		If strRet = "" Then
			Exit Function
		Else
			frm1.txtTrackingNo.value = strRet 
		End If		
		
		frm1.txtTrackingNo.focus
		Exit Function			
	
	End Select

	arrParam(0) = arrParam(5)								

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetConSItemDC(arrRet, iWhere)
	End If	
	
End Function

'========================================
Function PopZAdoConfigGrid()
	Dim arrRet
	
	On Error Resume Next
	
	If lgIsOpenPop = True Then Exit Function
	lgIsOpenPop = True

	arrRet = window.showModalDialog("../../ComAsp/ZADOSortPopup.asp",Array(ggoSpread.GetXMLData("A"),gMethodText),"dialogWidth=" & parent.SORTW_WIDTH & "px; dialogHeight=" & parent.SORTW_HEIGHT & "px;; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False
	
	If arrRet(0) = "X" Then
	   Exit Function
	Else
	   Call ggoSpread.SaveXMLData("A",arrRet(0),arrRet(1))
       Call InitVariables
       Call InitSpreadSheet()       
   End If
End Function

'=========================================
Function SetConSItemDC(Byval arrRet, Byval iWhere)
	With frm1
		Select Case iWhere
		Case 0
			.txtconBp_cd.value = arrRet(0) 
			.txtconBp_Nm.value = arrRet(1)   
		Case 1
			.txtDNType.value = arrRet(0) 
			.txtDNTypeNm.value = arrRet(1)   
		Case 3
			.txtSalesGroup.value = arrRet(0) 
			.txtSalesGroupNm.value = arrRet(1)   
		Case 4
			.txtPlant.value = arrRet(0) 
			.txtPlantNm.value = arrRet(1)   
		Case 5
			.txtStoRo_cd.value = arrRet(0) 
			.txtStoRo_Nm.value = arrRet(1)   
		End Select
	End With
End Function

'=========================================
Sub Form_Load()
    Call LoadInfTB19029	
'	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
'   Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field
	Call FormatField()
	Call LockFieldInit()
	
	Call InitVariables
	Call SetDefaultVal
	Call InitSpreadSheet()
    Call SetToolbar("11000000000011")							'⊙: 버튼 툴바 제어 
    frm1.txtconBp_cd.focus	   
End Sub

'=========================================
Sub Form_QueryUnload(Cancel , UnloadMode )
End Sub

'=========================================
Sub vspdData_Click(ByVal Col , ByVal Row )

	Call SetPopupMenuItemInf("00000000001")
	gMouseClickStatus = "SPC"
	Set gActiveSpdSheet = frm1.vspdData

	If frm1.vspdData.MaxRows = 0 Then                                                    'If there is no data.
		Exit Sub
	End If

	If Row <= 0 Then
		ggoSpread.Source = frm1.vspdData
		If lgSortKey = 1 Then
			ggoSpread.SSSort Col			'Sort In Assending
			lgSortKey = 2
		Else
			ggoSpread.SSSort Col, lgSortKey	'Sort In Desending
			lgSortKey = 1
		End If
		Exit Sub
	End If
End Sub

'=========================================
Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If

End Sub

'=========================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )

    If OldLeft <> NewLeft Then Exit Sub
    
	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	    '☜: 재쿼리 체크	
    	If lgPageNo <> "" Then		                                                    '다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
			If CheckRunningBizProcess Then Exit Sub
			Call DisableToolBar(parent.TBC_QUERY)
			Call DbQuery
    	End If
    End If
End Sub

'=======================================================
'Sub rdoQueryFlg1_OnClick()
'	frm1.txtRadio.value = frm1.rdoQueryFlg1.value
'End Sub

'=======================================================
'Sub rdoQueryFlg2_OnClick()
'	lblTitle.innerHTML = "출고일"
'	frm1.txtRadio.value = frm1.rdoQueryFlg2.value
'End Sub

'=======================================================
'Sub rdoQueryFlg3_OnClick()
'	lblTitle.innerHTML = "출고예정일"
'	frm1.txtRadio.value = frm1.rdoQueryFlg3.value
'End Sub
	
'=======================================================
Sub txtPromiseFrDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtPromiseFrDt.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtPromiseFrDt.focus
	End If
End Sub

'=======================================================
Sub txtPromiseToDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtPromiseToDt.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtPromiseToDt.focus
	End If
End Sub

'=======================================================
Sub txtPromiseFrDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub

'=======================================================
Sub txtPromiseToDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub

'=======================================================
Function FncQuery() 

    FncQuery = False                                                        
    
    Err.Clear                                                               

    lgIntFlgMode = parent.OPMD_CMODE

	If ValidDateCheck(frm1.txtPromiseFrDt, frm1.txtPromiseToDt) = False Then Exit Function

	ggoSpread.Source = frm1.vspdData
	ggoSpread.ClearSpreadData    
	
	Call InitVariables 														
    
	With frm1
		If .rdoQueryFlg1.checked = True Then
			.txtRadio.value = .rdoQueryFlg1.value
		ElseIf .rdoQueryFlg2.checked = True Then			
			.txtRadio.value = .rdoQueryFlg2.value
		ElseIf .rdoQueryFlg3.checked = True Then
			.txtRadio.value = .rdoQueryFlg3.value
		End If		
	End With

    Call DbQuery															'☜: Query db data

    FncQuery = True		
End Function

'=====================================================
Function FncPrint() 
    Call parent.FncPrint()
End Function

'=====================================================
Function FncExcel() 
	Call parent.FncExport(parent.C_MULTI)
End Function

'=====================================================
Function FncFind() 
    Call parent.FncFind(parent.C_MULTI , False)                                     
End Function

'=====================================================
Function FncExit()
    FncExit = True
End Function

'=====================================================
Function DbQuery() 
	Dim strVal

    DbQuery = False
    
    Err.Clear                                                               

	If LayerShowHide(1) = False Then
		Exit Function
	End If

'post방식 
		frm1.txtFlgMode.Value = lgIntFlgMode
		frm1.OPMD_UMODE.Value = parent.OPMD_UMODE
		frm1.txtMode.Value = Parent.UID_M0001				
		frm1.txt_lgPageNo.Value = lgPageNo                      '☜: Next key tag
		frm1.txt_lgStrPrevKey.Value = lgStrPrevKey                      '☜: Next key tag
		frm1.txt_lgSelectListDT.Value = GetSQLSelectListDataType("A")			 
		frm1.txt_lgTailList.Value = MakeSQLGroupOrderByList("A")
		frm1.txt_lgSelectList.Value = EnCoding(GetSQLSelectList("A"))

		Call ExecMyBizASP(frm1, BIZ_PGM_ID)          

    DbQuery = True
End Function

'=====================================================
Function DbQueryOk()														'☆: 조회 성공후 실행로직 

	lgBlnFlgChgValue = False
    lgIntFlgMode     = parent.OPMD_UMODE												'⊙: Indicates that current mode is Update mode

	Call SetToolbar("11000000000111")							'⊙: 버튼 툴바 제어 
    If frm1.vspdData.MaxRows > 0 Then
       frm1.vspdData.Focus
    End if  	

End Function

