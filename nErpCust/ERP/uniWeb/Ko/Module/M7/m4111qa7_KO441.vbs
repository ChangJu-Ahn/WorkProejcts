
'================================================================================================================================
Dim lgIsOpenPop      
Dim lgSaveRow  
Dim IsCookieSplit
Dim StartDate, EndDate

Const BIZ_PGM_ID 		= "m4111qb7_KO441.asp"    
Const BIZ_PGM_JUMP_ID 	= "m4111qa8"        
Const C_MaxKey          = 17		

'================================================================================================================================
Sub SetDefaultVal()
	frm1.txtMvFrDt.Text	= StartDate
	frm1.txtMvToDt.Text	= EndDate
	If lgPLCd <> "" then 
		Call ggoOper.SetReqAttr(frm1.txtPlantCd, "Q") 
		frm1.txtPlantCd.Tag = left(frm1.txtPlantCd.Tag,1) & "4" & mid(frm1.txtPlantCd.Tag,3,len(frm1.txtPlantCd.Tag))
        frm1.txtPlantCd.value = lgPLCd
	End If
End Sub
'================================================================================================================================
Sub InitVariables()
    lgPageNo         = ""
	lgBlnFlgChgValue = False
    lgSortKey        = 1
    lgSaveRow        = 0
    lgIntFlgMode = Parent.OPMD_CMODE 
End Sub
'================================================================================================================================
Sub InitSpreadSheet()
 	Call SetZAdoSpreadSheet("M4111QA7","G","A","V20030602", Parent.C_SORT_DBAGENT, frm1.vspdData, C_MaxKey, "X", "X" )
	Call SetSpreadLock("A")       
End Sub
'================================================================================================================================
Sub SetSpreadLock(ByVal pOpt)
    If pOpt = "A" Then
		ggoSpread.Source = frm1.vspdData
		ggoSpread.SpreadLockWithOddEvenRowColor()
	End If
End Sub
'================================================================================================================================
Function OpenPlantCd()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function
    If frm1.txtPlantCd.className = "protected" Then Exit Function

	lgIsOpenPop = True

	arrParam(0) = "����"						
	arrParam(1) = "B_PLANT"      				
	arrParam(2) = Trim(frm1.txtPlantCd.Value)	
'	arrParam(3) = Trim(frm1.txtPlantNm.Value)	
	arrParam(4) = ""							
	arrParam(5) = "����"						
	
    arrField(0) = "PLANT_CD"					
    arrField(1) = "PLANT_NM"					
    
    arrHeader(0) = "����"					
    arrHeader(1) = "�����"					
    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement		
		Exit Function
	Else
		frm1.txtPlantCd.Value = arrRet(0)
		frm1.txtPlantNm.Value = arrRet(1)
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement		
	End If	
	frm1.txtItemCd.value=""
	frm1.txtItemNm.value=""
End Function
'================================================================================================================================
Function OpenItemCd()
	Dim arrRet
	Dim arrParam(5), arrField(1)
	Dim iCalledAspName
	Dim IntRetCD
	
	If lgIsOpenPop = True Then Exit Function
	
	if  Trim(frm1.txtPlantCd.Value) = "" then
		Call DisplayMsgBox("17A002","X" , "����","X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement		
		Exit Function
	End if
	
	lgIsOpenPop = True

	arrParam(0) = Trim(frm1.txtPlantCd.value)	' Plant Code
	arrParam(1) = Trim(frm1.txtItemCd.value)		' Item Code
	arrParam(2) = "!"	' ��12!MO"�� ���� -- ǰ����� ������ ���ޱ��� 
	arrParam(3) = "30!P"
	arrParam(4) = ""		'-- ��¥ 
	arrParam(5) = ""		'-- ����(b_item_by_plant a, b_item b: and ���� ����)
	
	arrField(0) = 1 ' -- ǰ���ڵ� 
	arrField(1) = 2 ' -- ǰ���						
    
	iCalledAspName = AskPRAspName("B1B11PA3")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "B1B11PA3", "X")
		lgIsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(parent.window, arrParam, arrField), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
		
	lgIsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtItemCd.focus
		Set gActiveElement = document.activeElement		
		Exit Function
	Else
		frm1.txtItemCd.Value = arrRet(0)
		frm1.txtItemNm.Value = arrRet(1)
		frm1.txtItemCd.focus	
		Set gActiveElement = document.activeElement		
	End If
	
End Function
'================================================================================================================================
Function OpenBpCd()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	arrParam(0) = "�ŷ�ó"				
	arrParam(1) = "B_Biz_Partner"			
	arrParam(2) = Trim(frm1.txtBpCd.Value)	
'	arrParam(3) = Trim(frm1.txtBpNm.Value)	
	arrParam(4) = "BP_TYPE in (" & FilterVar("S", "''", "S") & " ," & FilterVar("CS", "''", "S") & ") "					
	arrParam(5) = "�ŷ�ó"				
	
    arrField(0) = "BP_CD"					
    arrField(1) = "BP_NM"					
    
    arrHeader(0) = "�ŷ�ó"				
    arrHeader(1) = "�ŷ�ó��"			
    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtBpCd.focus	
		Set gActiveElement = document.activeElement		
		Exit Function
	Else
		frm1.txtBpCd.Value = arrRet(0)
		frm1.txtBpNm.Value = arrRet(1)
		frm1.txtBpCd.focus	
		Set gActiveElement = document.activeElement		
	End If	
End Function
'================================================================================================================================
Function OpenSlCd()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	arrParam(0) = "â��"					
	arrParam(1) = "B_STORAGE_LOCATION"		
	arrParam(2) = Trim(frm1.txtSlCd.Value)	
'	arrParam(3) = Trim(frm1.txtSlNm.Value)	
	arrParam(4) = ""						
	arrParam(5) = "â��"					
	
    arrField(0) = "SL_CD"					
    arrField(1) = "SL_NM"					
    
    arrHeader(0) = "â��"				
    arrHeader(1) = "â���"				
    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtSlCd.focus	
		Set gActiveElement = document.activeElement		
		Exit Function
	Else
		frm1.txtSlCd.Value = arrRet(0)
		frm1.txtSlNm.Value = arrRet(1)
		frm1.txtSlCd.focus	
		Set gActiveElement = document.activeElement		
	End If	
End Function
'================================================================================================================================
Function OpenIoType()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	arrParam(0) = "�������"	
	arrParam(1) = "M_MVMT_TYPE"				
	
	arrParam(2) = Trim(frm1.txtIoType.Value)
'	arrParam(3) = Trim(frm1.txtIoTypeNm.Value)	
			
	arrParam(4) = "RCPT_FLG <> " & FilterVar("Y", "''", "S") & " "			
	arrParam(5) = "�������"			
	
    arrField(0) = "IO_TYPE_CD"	
    arrField(1) = "IO_TYPE_NM"	
    
    arrHeader(0) = "�������"		
    arrHeader(1) = "���������"
    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	lgIsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtIoTYpe.focus	
		Set gActiveElement = document.activeElement		
		Exit Function
	Else
		frm1.txtIoTYpe.Value = arrRet(0)
		frm1.txtIoTypeNm.Value = arrRet(1)
		frm1.txtIoTYpe.focus	
		Set gActiveElement = document.activeElement		
	End If	

End Function 
'================================================================================================================================
Sub PopZAdoConfigGrid()
	If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
		Exit Sub
	End If
	
	Call OpenGroupPopup("A")
End Sub
'================================================================================================================================
Function OpenGroupPopup(ByVal pSpdNo)
	Dim arrRet
	
	On Error Resume Next
	
	If lgIsOpenPop = True Then Exit Function
	lgIsOpenPop = True

	arrRet = window.showModalDialog("../../ComAsp/ZADOGroupPopup.asp",Array(ggoSpread.GetXMLData("A"),gMethodText),"dialogWidth=" & parent.GROUPW_WIDTH & "px; dialogHeight=" & parent.GROUPW_HEIGHT & "px;; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False
	
	If arrRet(0) = "X" Then
	   Exit Function
	Else
	   Call ggoSpread.SaveXMLData(pSpdNo,arrRet(0),arrRet(1))
       Call InitVariables
       Call InitSpreadSheet()       
   End If

End Function
'================================================================================================================================
Function CookiePage(ByVal Kubun)

	Dim strTemp, arrVal
	Dim strCookie, i

	Const CookieSplit = 4877				

	If Kubun = 1 Then						

		Call vspdData_Click(frm1.vspdData.ActiveCol,frm1.vspdData.ActiveRow)		
		WriteCookie CookieSplit , IsCookieSplit	
		
		If Len(Trim(frm1.txtPlantCd.value)) Then		
			WriteCookie "PlantCd",Trim(frm1.txtPlantCd.value) 
		Else
			WriteCookie "PlantCd",""
		End If
		
		If Len(Trim(frm1.txtItemCd.value)) Then
			WriteCookie "ItemCd",Trim(frm1.txtItemCd.value) 
		Else
			WriteCookie "ItemCd",""
		End If				
		
		If Len(Trim(frm1.txtBpCd.value)) Then
			WriteCookie "BpCd",Trim(frm1.txtBpCd.value) 
		Else
			WriteCookie "BpCd",""
		End If
		
		If Len(Trim(frm1.txtMvFrDt.text)) Then
			WriteCookie "MvFrDt",Trim(frm1.txtMvFrDt.text) 
		Else
			WriteCookie "MvFrDt",""
		End If
		
		If Len(Trim(frm1.txtMvToDt.text)) Then
			WriteCookie "MvToDt",Trim(frm1.txtMvToDt.text) 
		Else
			WriteCookie "MvToDt",""
		End If
				
		If Len(Trim(frm1.txtSlCd.value)) Then
			WriteCookie "SlCd",Trim(frm1.txtSlCd.value) 
		Else
			WriteCookie "SlCd",""
		End If
		
		If Len(Trim(frm1.txtIoType.value)) Then
			WriteCookie "IoType",Trim(frm1.txtIoType.value) 
		Else
			WriteCookie "IoType",""
		End If
		
			
		Call PgmJump(BIZ_PGM_JUMP_ID)

	ElseIf Kubun = 0 Then						

		strTemp = ReadCookie(CookieSplit)

		If strTemp = "" then Exit Function

		arrVal = Split(strTemp, gRowSep)

		'If arrVal(0) = "" Then Exit Function
		
		Dim iniSep

		If Err.number <> 0 Then
			Err.Clear
			WriteCookie CookieSplit , ""
			Exit Function 
		End If

		Call MainQuery()

		WriteCookie CookieSplit , ""

	End IF

End Function
'================================================================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)

   If Button = 2 And gMouseClickStatus = "SPC" Then
      gMouseClickStatus = "SPCR"
   End If
End Sub    
'================================================================================================================================
Sub FncSplitColumn()

    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   

End Sub
'================================================================================================================================
Sub txtMvFrDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtMvFrDt.Action = 7
		Call SetFocusToDocument("M") 
		frm1.txtMvFrDt.focus
	End If
End Sub
'================================================================================================================================
Sub txtMvToDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtMvToDt.Action = 7
		Call SetFocusToDocument("M") 
		frm1.txtMvToDt.focus
	End If
End Sub
'================================================================================================================================
Sub txtMvFrDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub

Sub txtMvToDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub
'================================================================================================================================
Sub vspdData_GotFocus()
    ggoSpread.Source = frm1.vspdData
End Sub
'================================================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)
	If Row = 0 Or Frm1.vspdData.MaxRows = 0 Then 
         Exit Sub
    End If
End Sub
'================================================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
    
    Dim iRowSep, ii
    
    Set gActiveSpdSheet = frm1.vspdData
    
    Call SetPopupMenuItemInf("00000000001")		
	gMouseClickStatus = "SPC"   
	
	   
	If frm1.vspdData.MaxRows = 0 Then Exit Sub
	   
	If Row <= 0 Then
		ggoSpread.Source = frm1.vspdData
		If lgSortKey = 1 Then
			ggoSpread.SSSort Col				'Sort in Ascending
			lgSortkey = 2
		Else
			ggoSpread.SSSort Col, lgSortKey	'Sort in Descending
			lgSortkey = 1
		End If
		
		Exit Sub
	End If    
	
	Call SetSpreadColumnValue("A",Frm1.vspdData, Col, Row) 

	IsCookieSplit = ""
	iRowSep = Parent.gRowSep 
		
	For ii = 1 to frm1.vspdData.MaxCols - 1
	    IsCookieSplit = IsCookieSplit & Trim(GetSpreadText(frm1.vspdData,GetKeyPos("A",ii),Row,"X","X")) & iRowSep
	Next
	
End Sub
'================================================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft, ByVal OldTop, ByVal NewLeft, ByVal NewTop)
	
	If OldLeft <> NewLeft Then
	    Exit Sub
	End If
    
	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	    '��: ������ üũ	
		If lgPageNo <> "" Then		                                                    '���� Ű ���� ������ �� �̻� ��������ASP�� ȣ������ ���� 
			If DbQuery = False Then
				Exit Sub
			End if
		End If
	End If
End Sub
'================================================================================================================================
Function FncQuery() 

    FncQuery = False                            
    Err.Clear                                   
    
    with frm1
		if (UniConvDateToYYYYMMDD(.txtMvFrDt.text,Parent.gDateFormat,"") > UniConvDateToYYYYMMDD(.txtMvToDt.text,Parent.gDateFormat,"")) And Trim(.txtMvFrDt.text) <> "" And Trim(.txtMvToDt.text) <> "" then	
			Call DisplayMsgBox("17a003","X","�����","X")			
			Exit Function
		End if   
	End with

    Call InitVariables 							
    
    ggoSpread.Source = frm1.vspdData
	ggoSpread.ClearSpreadData
	
	If DbQuery = False Then Exit Function

    FncQuery = True								

End Function
'================================================================================================================================
Function FncPrint() 
    Call parent.FncPrint()
    Set gActiveElement = document.activeElement
End Function
'================================================================================================================================
Function FncExcel() 
	Call parent.FncExport(parent.C_MULTI)
	Set gActiveElement = document.activeElement
End Function
'================================================================================================================================
Function FncFind() 
    Call parent.FncFind(parent.C_MULTI , False)
    Set gActiveElement = document.activeElement        
End Function
'================================================================================================================================
Function FncExit()
    FncExit = True
    Set gActiveElement = document.activeElement
End Function
'================================================================================================================================
Function DbQuery() 
	Dim strVal
	
    DbQuery = False
    
    Err.Clear                                   
    If CheckRunningBizProcess = True Then
       Exit Function
    End If                                              
    
    Call LayerShowHide(1) 
    
    With frm1
    If lgIntFlgMode = Parent.OPMD_UMODE Then
        strVal = BIZ_PGM_ID & "?txtPlantCd=" & Trim(.hdnPlantCd.value)
	    strVal = strVal & "&txtItemCd=" & Trim(.hdnItemCd.value)
	    strVal = strVal & "&txtBpCd=" & Trim(.hdnBpCd.value)
    	strVal = strVal & "&txtMvFrDt=" & Trim(.hdnMvFrDt.value)
    	strVal = strVal & "&txtMvToDt=" & Trim(.hdnMvToDt.value)    	
    	strVal = strVal & "&txtSlCd=" & Trim(.hdnSlCd.value)
    	strVal = strVal & "&txtIoType=" & Trim(.hdnIoType.value)    	
    else
	    strVal = BIZ_PGM_ID & "?txtPlantCd=" & Trim(.txtPlantCd.value)
	    strVal = strVal & "&txtItemCd=" & Trim(.txtItemCd.value)
	    strVal = strVal & "&txtBpCd=" & Trim(.txtBpCd.value)
    	strVal = strVal & "&txtMvFrDt=" & Trim(.txtMvFrDt.Text)
    	strVal = strVal & "&txtMvToDt=" & Trim(.txtMvToDt.Text)    	
    	strVal = strVal & "&txtSlCd=" & Trim(.txtSlCd.value)
    	strVal = strVal & "&txtIoType=" & Trim(.txtIoType.value)    	
    end if
        strVal = strVal & "&lgPageNo="   & lgPageNo        
        strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")
        strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
		strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))
  
        strVal = strVal & "&gBizArea=" & lgBACd 
        strVal = strVal & "&gPlant=" & lgPLCd 
        strVal = strVal & "&gPurGrp=" & lgPGCd 
        strVal = strVal & "&gPurOrg=" & lgPOCd  

        Call RunMyBizASP(MyBizASP, strVal)										
        
    End With
    
    DbQuery = True
    Call SetToolbar("1100000000011111")				

End Function
'================================================================================================================================
Function DbQueryOk()								

    lgBlnFlgChgValue = False
    lgSaveRow        = 1
    lgIntFlgMode = Parent.OPMD_UMODE
    
    If frm1.vspdData.MaxRows > 0 Then
		frm1.vspdData.Focus
	Else
		frm1.txtPlantCd.focus
	End If
	
	Set gActiveElement = document.activeElement	
	
End Function
'================================================================================================================================
