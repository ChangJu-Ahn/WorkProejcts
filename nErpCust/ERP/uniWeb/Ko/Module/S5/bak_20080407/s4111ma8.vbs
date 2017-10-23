Option Explicit   

' External ASP File
'========================================
Const BIZ_PGM_ID = "s4111mb8.asp"											'☆: 비지니스 로직 ASP명 
Const BIZ_PGM_JUMP_ID = "s4111ma1"	     									'☆: JUMP시 비지니스 로직 ASP명 

' Constant variables 
'========================================
Const C_MaxKey          = 1                                           

' User-defind Variables
'========================================
Dim IsOpenPop     
Dim lgIsOpenPop     

'==========================================
Dim lsDnNo         

Dim GridBLNo			
Dim GridBLDocNo
Dim GridApplicantCd
Dim GridApplicantNm
Dim GridCur
Dim GridDocAmt

'=========================================
Sub FormatField()
    With frm1
        ' 날짜 OCX Foramt 설정 
        Call FormatDATEField(.txtReqGiDtFrom)
        Call FormatDATEField(.txtReqGiDtTo)
    End With
End Sub

'=========================================
Sub LockFieldInit(ByVal pvFlag)
    With frm1
        ' 날짜 OCX
        Call LockObjectField(.txtReqGiDtFrom, "O")
        Call LockObjectField(.txtReqGiDtTo, "O")
    End With

End Sub

'=========================================
Sub InitVariables()
    lgStrPrevKey     = ""
    lgPageNo         = ""
	lgBlnFlgChgValue = False
    lgIntFlgMode     = parent.OPMD_CMODE                          
    lgSortKey        = 1
End Sub

'=========================================
Sub SetDefaultVal()

	Set gActiveElement = document.activeElement 
	
	frm1.rdoPostGiFlagAll.checked = True
	frm1.txtPostGiFlag.value = frm1.rdoPostGiFlagAll.value   
	lgBlnFlgChgValue = False
	frm1.txtDn_Type.focus
	frm1.txtSalesGrp.value = parent.gSalesGrp
	frm1.txtReqGiDtFrom.Text = StartDate
	frm1.txtReqGiDtTo.Text = EndDate
End Sub

'=========================================
Sub InitSpreadSheet()
	Call SetZAdoSpreadSheet("S4111MA8","S","A","V20021106", Parent.C_SORT_DBAGENT, frm1.vspdData, C_MaxKey, "X", "X" )
    Call SetSpreadLock 
End Sub

'=========================================
Sub SetSpreadLock()
	ggoSpread.SpreadLockWithOddEvenRowColor()
End Sub

'=========================================
Sub Form_Load()

	Call LoadInfTB19029				                                           '⊙: Load table , B_numeric_format
	Call FormatField()
	Call LockFieldInit("L")
	Call InitVariables														    
	Call SetDefaultVal	
	Call InitSpreadSheet()
	Call SetToolbar("11000000000011")							'⊙: 버튼 툴바 제어 
	
	frm1.txtDn_Type.focus	

End Sub

'=========================================
Sub Form_QueryUnload(Cancel , UnloadMode )
End Sub

'=========================================
Function FncQuery() 

    FncQuery = False                                                        
   
    Err.Clear     
                                                             
	If ValidDateCheck(frm1.txtReqGiDtFrom, frm1.txtReqGiDtTo) = False Then Exit Function

    Call ggoOper.ClearField(Document, "2")
    Call InitVariables 														

    If DbQuery = False Then Exit Function

    FncQuery = True															

End Function

'========================================
Function FncPrint()
    FncPrint = False                                                             
    Err.Clear                                                                    
	Call Parent.FncPrint()                                                       
    FncPrint = True                                                              
End Function

'========================================
Function FncExcel() 
    FncExcel = False                                                             
    Err.Clear                                                                    

	Call Parent.FncExport(parent.C_MULTI)

    FncExcel = True                                                              
End Function

'========================================
Function FncFind() 
    FncFind = False                                                              
    Err.Clear                                                                    

	Call Parent.FncFind(parent.C_MULTI, True)

    FncFind = True                                                               
End Function

'========================================
Function FncExit()
    FncExit = True                                                               
End Function

'========================================
Function DbQuery() 
	Dim strVal

    DbQuery = False
    
    Err.Clear                                                               
	Call LayerShowHide(1)
    
    With frm1

	If lgIntFlgMode = parent.OPMD_UMODE Then  
	    strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001									

		strVal = strVal & "&txtDn_Type=" & Trim(.txtHDn_Type.value)
		strVal = strVal & "&txtSo_no=" & Trim(.txtHSo_no.value)
		strVal = strVal & "&txtShip_to_party=" & Trim(.txtHShip_to_party.value)
		strVal = strVal & "&txtReqGiDtFrom=" & Trim(.txtHReqGiDtFrom.value)
		strVal = strVal & "&txtReqGiDtTo=" & Trim(.txtHReqGiDtTo.value)
		strVal = strVal & "&txtTrans_meth=" & Trim(.txtHTrans_meth.value)
		strVal = strVal & "&txtPostGiFlag=" & Trim(.txtHPostGiFlag.value)
		strVal = strVal & "&txtSalesGrp=" & Trim(frm1.txtHSalesGrp.value)
		
        strVal = strVal & "&lgStrPrevKey="   & lgStrPrevKey                      '☜: Next key tag
   	
   	Else  	
   	
		strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001									
		
		strVal = strVal & "&txtDn_Type=" & Trim(.txtDn_Type.value)
		strVal = strVal & "&txtSo_no=" & Trim(.txtSo_no.value)
		strVal = strVal & "&txtShip_to_party=" & Trim(.txtShip_to_party.value)
		strVal = strVal & "&txtReqGiDtFrom=" & Trim(.txtReqGiDtFrom.Text)
		strVal = strVal & "&txtReqGiDtTo=" & Trim(.txtReqGiDtTo.Text)
		strVal = strVal & "&txtTrans_meth=" & Trim(.txtTrans_meth.value)
		strVal = strVal & "&txtPostGiFlag=" & Trim(.txtPostGiFlag.value)
		strVal = strVal & "&txtSalesGrp=" & Trim(frm1.txtSalesGrp.value)

		strVal = strVal & "&lgStrPrevKey="   & lgStrPrevKey                      '☜: Next key tag
	End If			
        strVal =     strVal & "&lgPageNo="       & lgPageNo                          '☜: Next key tag
		strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")			 
		strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
		strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))
  
    Call RunMyBizASP(MyBizASP, strVal)										

    End With
    
    DbQuery = True
End Function

'========================================
Function DbQueryOk()													

	lgBlnFlgChgValue = False
    lgIntFlgMode     = parent.OPMD_UMODE								
	
	Call SetToolbar("11000000000111")
	
    If frm1.vspdData.MaxRows > 0 Then
       frm1.vspdData.Focus
    Else
       frm1.txtDn_Type.focus	
    End if  	

End Function

'========================================
Function OpenSORef()
	Dim iCalledAspName
	Dim arrRet
	Dim arrParam(0)

	On Error Resume Next

	If lgIntFlgMode <> parent.OPMD_UMODE Then
		Call DisplayMsgBox("900002", "X", "X", "X")
		Exit Function
	End IF

	If IsOpenPop = True Then Exit Function

	Call vspdData_Click(frm1.vspdData.ActiveCol,frm1.vspdData.ActiveRow)

	If lsDnNo = "" Then
		Call DisplayMsgBox("204198", parent.VB_YES_NO, "X", "X")
		Exit Function
	End IF

	IsOpenPop = True

	iCalledAspName = AskPRAspName("S4112RA8")
		
	if Trim(iCalledAspName) = "" then
		IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION, "S4112RA8", "x")
		gblnWinEvent = False
		exit Function
	end if

	arrParam(0) = lsDnNo
   
	arrRet = window.showModalDialog(iCalledAspName,Array(window.parent,arrParam),"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

End Function

'========================================
Function OpenSoNo()
	Dim iCalledAspName
	Dim strRet

	If IsOpenPop = True Then Exit Function
			
	IsOpenPop = True
		
	iCalledAspName = AskPRAspName("S3111PA1")
		
	if Trim(iCalledAspName) = "" then
		IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION, "S3111PA1", "x")
		gblnWinEvent = False
		exit Function
	end if

	strRet = window.showModalDialog(iCalledAspName, array(window.parent,"DN"), _
		"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

    frm1.txtSo_no.focus

	If strRet <> "" Then
		frm1.txtSo_no.value = strRet 
	End If	

End Function

'========================================
Function OpenRequried(ByVal iRequried)

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	Select Case iRequried
	Case 1												
		arrParam(0) = "출하형태"
		arrParam(1) = "B_MINOR A, I_MOVETYPE_CONFIGURATION B"				
		arrParam(2) = Trim(frm1.txtDn_Type.value)		
		arrParam(3) = ""
		arrParam(4) = "A.MINOR_CD=B.MOV_TYPE AND (B.TRNS_TYPE = " & FilterVar("DI", "''", "S") & " OR (B.TRNS_TYPE = " & FilterVar("ST", "''", "S") & " AND B.STCK_TYPE_FLAG_DEST = " & FilterVar("T", "''", "S") & " )) AND A.MAJOR_CD=" & FilterVar("I0001", "''", "S") & " "	
		arrParam(5) = "출하형태"

		arrField(0) = "A.MINOR_CD"
		arrField(1) = "A.MINOR_NM"

		arrHeader(0) = "출하형태"					
		arrHeader(1) = "출하형태명"
		
		frm1.txtDn_Type.focus

	Case 2												
		arrParam(0) = "납품처"						
		arrParam(1) = "B_BIZ_PARTNER_FTN PARTNER_FTN,B_BIZ_PARTNER PARTNER"		
		arrParam(2) = Trim(frm1.txtShip_to_party.value)							
		'arrParam(3) = Trim(frm1.txtShip_to_partyNm.value)						
		arrParam(4) = "PARTNER_FTN.USAGE_FLAG=" & FilterVar("Y", "''", "S") & "  AND PARTNER_FTN.PARTNER_FTN=" & FilterVar("SSH", "''", "S") & " AND PARTNER.BP_CD=PARTNER_FTN.BP_CD AND PARTNER.BP_TYPE <= " & FilterVar("CS", "''", "S") & ""
		arrParam(5) = "납품처"						
		
	    arrField(0) = "PARTNER_FTN.PARTNER_BP_CD"				
	    arrField(1) = "PARTNER.BP_NM"					
	    arrField(2) = "PARTNER_FTN.BP_CD"		
	    arrField(3) = "PARTNER_FTN.PARTNER_FTN"			
	    arrField(4) = "PARTNER_FTN.USAGE_FLAG"					
	    
	    arrHeader(0) = "납품처"						
	    arrHeader(1) = "납품처명"					
	    arrHeader(2) = "거래처코드"					
	    arrHeader(3) = "거래처타입"					
	    arrHeader(4) = "사용여부"	
	    
	    frm1.txtShip_to_party.focus

	Case 3												
		arrParam(0) = "운송방법"					
		arrParam(1) = "B_MINOR"							
		arrParam(2) = Trim(frm1.txtTrans_meth.value)	
		'arrParam(3) = Trim(frm1.txtTrans_meth_nm.value)	
		arrParam(4) = "MAJOR_CD=" & FilterVar("B9009", "''", "S") & ""				
		arrParam(5) = "운송방법"					
		
	    arrField(0) = "MINOR_CD"						
	    arrField(1) = "MINOR_NM"						
	    
	    arrHeader(0) = "운송방법"					
	    arrHeader(1) = "운송방법명"
	    
		frm1.txtTrans_meth.focus					

	End Select
    
	Select Case iRequried
	Case 2
		arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	Case Else
		arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	End Select

	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetRequried(arrRet,iRequried)
	End If	
	
End Function

'========================================
Function OpenConSalesGrp()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "영업그룹"					
	arrParam(1) = "B_SALES_GRP"						
	arrParam(2) = Trim(frm1.txtSalesGrp.value)		
	arrParam(3) = ""
	arrParam(4) = "USAGE_FLAG=" & FilterVar("Y", "''", "S") & " "					
	arrParam(5) = "영업그룹"					
		
	arrField(0) = "SALES_GRP"						
	arrField(1) = "SALES_GRP_NM"					
	    
	arrHeader(0) = "영업그룹"					
	arrHeader(1) = "영업그룹명"					

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	frm1.txtSalesGrp.focus
	
	If arrRet(0) <> "" Then
		frm1.txtSalesGrp.value = arrRet(0)
		frm1.txtSalesGrpNm.value = arrRet(1)
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
Function SetRequried(Byval arrRet,ByVal iRequried)

	Select Case iRequried
	Case 1
		frm1.txtDn_Type.value = arrRet(0) 
		frm1.txtDn_TypeNm.value = arrRet(1)   
	Case 2
		frm1.txtShip_to_party.value = arrRet(0) 
		frm1.txtShip_to_partyNm.value = arrRet(1)   
	Case 3
		frm1.txtTrans_meth.value = arrRet(0) 
		frm1.txtTrans_meth_nm.value = arrRet(1)   
	End Select

End Function

'=========================================
Function CookiePage(ByVal Kubun)

	On Error Resume Next

	Const CookieSplit = 4877

	Call vspdData_Click(frm1.vspdData.ActiveCol,frm1.vspdData.ActiveRow)

    ' Check Ryu
    If Kubun = 1 Then
		WriteCookie CookieSplit , lsDnNo
	End If

End Function

'=========================================
Sub SetQuerySpreadColor(ByVal lRow)
	Dim GCol
    With frm1

		.vspdData.ReDraw = False
		For GCol = 1  To 12
			ggoSpread.SSSetProtected GCol, lRow, .vspdData.MaxRows
		Next
		.vspdData.ReDraw = True
    End With

End Sub

'=========================================
Sub rdoPostGiFlagAll_OnClick()
	frm1.txtPostGiFlag.value = frm1.rdoPostGiFlagAll.value 
End Sub

'=========================================
Sub rdoPostGiFlagYes_OnClick()
	frm1.txtPostGiFlag.value = frm1.rdoPostGiFlagYes.value 
End Sub

'=========================================
Sub rdoPostGiFlagNo_OnClick()
	frm1.txtPostGiFlag.value = frm1.rdoPostGiFlagNo.value 
End Sub

'========================================
Function vspdData_DblClick(ByVal Col, ByVal Row)
	If frm1.vspdData.MaxRows > 0 AND frm1.vspdData.ActiveRow > 0 Then
		If frm1.vspdData.ActiveRow = Row Then
			Call OpenSORef
		End If
	End If
End Function

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

	If Row < 1 Then Exit Sub

	frm1.vspdData.Row = Row
	frm1.vspdData.Col = GetKeyPos("A",1) ' 1
	lsDnNo=frm1.vspdData.Text
  
End Sub

'=======================================================
Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If

End Sub    

'========================================
Sub vspdData_GotFocus()
    ggoSpread.Source = frm1.vspdData
End Sub

'=======================================================
Sub vspdData_TopLeftChange(OldLeft , OldTop , NewLeft , NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	    '☜: 재쿼리 체크	
    	If lgPageNo <> "" Then		                                                    '다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
			If CheckRunningBizProcess Then Exit Sub
			Call DisableToolBar(parent.TBC_QUERY)
			Call DbQuery
    	End If
    End If
    
End Sub

'=======================================================
Sub txtReqGiDtFrom_DblClick(Button)
	If Button = 1 Then
		frm1.txtReqGiDtFrom.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtReqGiDtFrom.focus
	End If
End Sub

'=======================================================
Sub txtReqGiDtTo_DblClick(Button)
	If Button = 1 Then
		frm1.txtReqGiDtTo.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtReqGiDtTo.focus		
	End If
End Sub

'=======================================================
Sub txtReqGiDtFrom_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub

'=======================================================
Sub txtReqGiDtTo_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub


