Option Explicit

Const BIZ_PGM_ID 		= "m5111qb2_KO441.asp"                         '☆: 비지니스 로직 ASP명 
Const BIZ_PGM_JUMP_ID 	= "m5112MA1"                             '☆: Cookie에서 사용할 상수 
Const C_MaxKey          = 26							         '☆☆☆☆: Max key value

'==========================================  2.1.1 InitVariables()  ======================================
Sub InitVariables()
    lgPageNo         = ""
    lgBlnFlgChgValue = False
    lgSortKey        = 1
    lgSaveRow        = 0
    lgIntFlgMode = parent.OPMD_CMODE 
End Sub

'======================= 2.2.3 InitSpreadSheet() ========================================
Sub InitSpreadSheet()
    
	Call SetZAdoSpreadSheet("M5111QA2","S","A","V20030913",parent.C_SORT_DBAGENT,frm1.vspdData, _
									C_MaxKey, "X","X")
    Call SetSpreadLock 
     
End Sub
'================================== 2.2.4 SetSpreadLock() ==================================================
Sub SetSpreadLock()
	ggoSpread.Source = frm1.vspdData
	ggoSpread.SpreadLockWithOddEvenRowColor()
End Sub

'------------------------------------------  OpenBizArea()  -------------------------------------------------
Function OpenBizArea()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function
    If frm1.txtBizArea.className = "protected" Then Exit Function
	
	lgIsOpenPop = True

	arrParam(0) = "사업장"					
	arrParam(1) = "B_BIZ_AREA"					
	arrParam(2) = Trim(frm1.txtBizArea.Value)	
	arrParam(5) = "사업장"					

    arrField(0) = "BIZ_AREA_CD"					
    arrField(1) = "BIZ_AREA_NM"					
    
    
    arrHeader(0) = "사업장"					
    arrHeader(1) = "사업장명"				
    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtBizArea.focus
		Exit Function
	Else
		frm1.txtBizArea.Value	= arrRet(0)
		frm1.txtBizAreaNm.value = arrRet(1)
		frm1.txtBizArea.focus	
		Set gActiveElement = document.activeElement
	End If	
End Function
'------------------------------------------  OpenItemCd()  -------------------------------------------------
Function OpenItemCd()
	Dim arrRet
	Dim arrParam(5), arrField(1)
	Dim iCalledAspName
	Dim IntRetCD
	
	If lgIsOpenPop = True Then Exit Function
	
	if  Trim(frm1.txtPlantCd.Value) = "" then
		Call DisplayMsgBox("17A002","X" , "공장","X")
		Exit Function
	End if
	
	lgIsOpenPop = True

	arrParam(0) = Trim(frm1.txtPlantCd.value)	' Plant Code
	arrParam(1) = Trim(frm1.txtItemCd.value)		' Item Code
	arrParam(2) = "!"	' “12!MO"로 변경 -- 품목계정 구분자 조달구분 
	arrParam(3) = "30!P"
	arrParam(4) = ""		'-- 날짜 
	arrParam(5) = ""		'-- 자유(b_item_by_plant a, b_item b: and 부터 시작)
	
	arrField(0) = 1 ' -- 품목코드 
	arrField(1) = 2 ' -- 품목명				
    
	iCalledAspName = AskPRAspName("B1B11PA3")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "B1B11PA3", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam, arrField), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtItemCd.focus
		Exit Function
	Else
		frm1.txtItemCd.Value = arrRet(0)
		frm1.txtItemNm.Value = arrRet(1)
		frm1.txtItemCd.focus	
		Set gActiveElement = document.activeElement
	End If	
End Function
 '------------------------------------------  OpenPlantCd()  -------------------------------------------------
Function OpenPlantCd()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function
    If frm1.txtPlantCd.className = "protected" Then Exit Function

	lgIsOpenPop = True

	arrParam(0) = "공장"						 
	arrParam(1) = "B_PLANT"      					 
	arrParam(2) = Trim(frm1.txtPlantCd.Value)		 
'	arrParam(3) = Trim(frm1.txtPlantNm.Value)		 
	arrParam(4) = ""								 
	arrParam(5) = "공장"						 
	
    arrField(0) = "PLANT_CD"						 
    arrField(1) = "PLANT_NM"						 
    
    arrHeader(0) = "공장"						 
    arrHeader(1) = "공장명"						 
    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtPlantCd.focus	
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


'------------------------------------------  OpenBpCd()  -------------------------------------------------
Function OpenBpCd()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	arrParam(0) = "공급처"						 
	arrParam(1) = "B_Biz_Partner"					 
	arrParam(2) = Trim(frm1.txtBpCd.Value)		 
'	arrParam(3) = Trim(frm1.txtBpNm.Value)		 
	arrParam(4) = "BP_TYPE In (" & FilterVar("S", "''", "S") & " ," & FilterVar("CS", "''", "S") & ") And usage_flag=" & FilterVar("Y", "''", "S") & " "					 
	arrParam(5) = "공급처"						 
	
    arrField(0) = "BP_CD"							 
    arrField(1) = "BP_NM"						 
    
    arrHeader(0) = "공급처"						 
    arrHeader(1) = "공급처명"					 
    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtBpCd.focus
		Exit Function
	Else
		frm1.txtBpCd.Value = arrRet(0)
		frm1.txtBpNm.Value = arrRet(1)
		frm1.txtBpCd.focus	
		Set gActiveElement = document.activeElement
	End If	
End Function

'------------------------------------------  OpenIvType()  -------------------------------------------------
Function OpenIvType()

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	arrParam(0) = "매입형태"						 
	arrParam(1) = "M_IV_TYPE"							 
	arrParam(2) = Trim(frm1.txtIvType.Value)			 
'	arrParam(3) = Trim(frm1.txtIvTypeNm.Value)			 
	arrParam(4) = ""									 
	arrParam(5) = "매입형태"						 
	
    arrField(0) = "IV_TYPE_CD"							 
    arrField(1) = "IV_TYPE_NM"							 
        
    arrHeader(0) = "매입형태"						 
    arrHeader(1) = "매입형태명"						 
    
    arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtIvType.focus
		Exit Function
	Else
		frm1.txtIvType.Value = arrRet(0)
		frm1.txtIvTypeNm.Value = arrRet(1)
		frm1.txtIvType.focus	
		Set gActiveElement = document.activeElement
	End If	
End Function

 '------------------------------------------  OpenPurGrpCd()  -------------------------------------------------
Function OpenPurGrpCd()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function
    If frm1.txtPurGrpCd.className = "protected" Then Exit Function

	lgIsOpenPop = True

	arrParam(0) = "구매그룹"	
	arrParam(1) = "B_Pur_Grp"				
	
	arrParam(2) = Trim(frm1.txtPurGrpCd.Value)
'	arrParam(3) = Trim(frm1.txtPurGrpNm.Value)	
	
	arrParam(4) = "USAGE_FLG=" & FilterVar("Y", "''", "S") & " "			
	arrParam(5) = "구매그룹"			
	
    arrField(0) = "PUR_GRP"	
    arrField(1) = "PUR_GRP_NM"	
    
    arrHeader(0) = "구매그룹"		
    arrHeader(1) = "구매그룹명"
    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	lgIsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtPurGrpCd.focus	
		Exit Function
	Else
		frm1.txtPurGrpCd.Value = arrRet(0)
		frm1.txtPurGrpNm.Value = arrRet(1)
		frm1.txtPurGrpCd.focus	
		Set gActiveElement = document.activeElement
	End If	

End Function 

'------------------------------------  OpenGroupPopup()  ----------------------------------------------
Function OpenGroupPopup()

	Dim arrRet
	Dim arrParam
	Dim TInf(5)
	Dim ii
	Dim iLoop
	Dim tmpPopUpR
	
	On Error Resume Next
	
	ReDim arrParam(parent.C_MaxSelList * 2 - 1 )

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True
	
    TInf(0) = gMethodText
	
	tmpPopUpR = GetPopUpR("A")
	
	For iLoop = 0 to parent.C_MaxSelList * 2 - 1 Step 2
      arrParam(iLoop + 0 ) = tmpPopUpR(iLoop / 2  , 0)
      arrParam(iLoop + 1 ) = tmpPopUpR(iLoop / 2  , 1)
    Next  
      
	arrRet = window.showModalDialog("../../ComAsp/ADOGrpSortPopup.asp",Array(GetSQLSortFieldCD("A"),GetSQLSortFieldNm("A"),arrParam,TInf),"dialogWidth=420px; dialogHeight=250px;; center: Yes; help: No; resizable: No; status: No;")
	
	lgIsOpenPop = False

	If arrRet(0) = "0" Then
		If Err.Number <> 0 Then
			Err.Clear 
		End If
		Exit Function
	Else
	
	   Call SetPopUpR("A",arrRet,frm.vspdData)   
	   
       Call InitVariables
       Call InitSpreadSheet
       
   End If
   
End Function

'==========================================   CookiePage()  ======================================
Function CookiePage(ByVal Kubun)

	Dim strTemp, arrVal
	Dim strCookie, i

	Const CookieSplit = 4877						

	If Kubun = 1 Then								

		Call vspdData_Click(frm1.vspdData.ActiveCol,frm1.vspdData.ActiveRow)

		WriteCookie CookieSplit , IscookieSplit						
		if frm1.vspdData.ActiveRow > 0 then			
			strTemp = ReadCookie(CookieSplit)
			If strTemp = "" then Exit Function
			arrVal = Split(strTemp, parent.gRowSep)
			frm1.vspdData.Row = frm1.vspdData.ActiveRow 			
			WriteCookie "txtIvNo" , arrVal(0)					
			WriteCookie CookieSplit , ""			
		end if		
		Call PgmJump(BIZ_PGM_JUMP_ID)

	ElseIf Kubun = 0 Then							
		strTemp = ReadCookie(CookieSplit)
					
		If strTemp = "" then Exit Function

		arrVal = Split(strTemp, parent.gRowSep)

		
		Dim iniSep

		
		If Len(ReadCookie ("BpCd")) Then
			frm1.txtBpCd.Value	=  ReadCookie ("BpCd")
			WriteCookie "BpCd",""
		Else
			frm1.txtBpCd.Value		=  arrVal(0)
		End If
		
		frm1.txtBpNm.value			=  arrVal(1)
		
		If Len(ReadCookie ("tBizArea")) Then
			frm1.txtBizArea.Value	=  ReadCookie ("tBizArea")
			WriteCookie "tBizArea",""
		Else
			frm1.txtBizArea.Value	=  arrVal(3)
		End If
		
		frm1.txtBizAreaNm.value		=  arrVal(4)
		
		If Len(ReadCookie ("ItemCd")) Then
			frm1.txtItemCd.Value	=  ReadCookie ("ItemCd")
			WriteCookie "ItemCd",""
		Else
			frm1.txtItemCd.Value	=  arrVal(5)
		End If
		
		frm1.txtItemNm.Value		=  arrVal(6)
		
		If arrVal(7) = "" or arrVal(7) = Null Then
			frm1.txtIvFrDt.Text		=  ReadCookie ("IvFrDt")
			WriteCookie "IvFrDt",""
		Else
			frm1.txtIvFrDt.Text		=  arrVal(7)			
		End If
		
		If arrVal(7) = "" or arrVal(7) = Null Then
			frm1.txtIvToDt.Text		=  ReadCookie ("IvToDt")
			WriteCookie "IvToDt",""
		Else
			frm1.txtIvToDt.Text		=  arrVal(7)			
		End If
		
		If Len(ReadCookie ("PlantCd")) Then
			frm1.txtPlantCd.Value	=  ReadCookie ("PlantCd")
			WriteCookie "PlantCd",""
		Else
			frm1.txtPlantCd.Value	=  arrVal(8)
		End If
		
		frm1.txtPlantNm.value		=  arrVal(9)
		
		If Len(ReadCookie ("IvType")) Then
			frm1.txtIvType.Value	=  ReadCookie ("IvType")
			WriteCookie "IvType",""
		Else
			frm1.txtIvType.Value	=  arrVal(10)
		End If
				
		frm1.txtIvTypeNm.value	=  arrVal(11)
		
		If Len(ReadCookie ("PurGrpCd")) Then
			frm1.txtPurGrpCd.Value	=  ReadCookie ("PurGrpCd")
			WriteCookie "PurGrpCd",""
		Else
			frm1.txtPurGrpCd.Value	=  arrVal(12)
		End If
				
		frm1.txtPurGrpNm.value		=  arrVal(13)		
		
		If Len(ReadCookie ("PstFlg")) Then
			frm1.cboPstFlg.Value	=  ReadCookie ("PstFlg")
			WriteCookie "PstFlg",""
		Else
			frm1.cboPstFlg.Value	=  arrVal(14)
		End If
				
		If Err.number <> 0 Then
			Err.Clear
			WriteCookie CookieSplit , ""
			Exit Function 
		End If

		Call MainQuery()

		WriteCookie CookieSplit , ""

	End IF

End Function

'------------------------------------  PopZAdoConfigGrid()  ----------------------------------------------
Sub PopZAdoConfigGrid()
	If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
		Exit Sub
	End If
	
	Call OpenOrderByPopup("A")
End Sub

'===========================================================================
' Function Name : OpenOrderByPopup
'===========================================================================
 Function OpenOrderByPopup(ByVal pSpdNo)

	Dim arrRet
	
	On Error Resume Next
	
	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True
	
    arrRet = window.showModalDialog("../../ComAsp/ZADOSortPopup.asp",Array(ggoSpread.GetXMLData("A"), gMethodText),"dialogWidth=" & parent.GROUPW_WIDTH & "px; dialogHeight=" & parent.GROUPW_HEIGHT & "px;; center: Yes; help: No; resizable: No; status: No;")
	lgIsOpenPop = False

	If arrRet(0) = "X" Then
		If Err.Number <> 0 Then
			Err.Clear 
		End If
		Exit Function
	Else
	   Call ggoSpread.SaveXMLData(pSpdNo, arrRet(0), arrRet(1))
       Call InitVariables
       Call InitSpreadSheet
   End If
End Function

'=============================  Form_QueryUnload()  ==============================================
Sub Form_QueryUnload(Cancel , UnloadMode )
   
End Sub

'=============================  vspdData_MouseDown()  ==============================================
Sub vspdData_MouseDown(Button , Shift , x , y)

   If Button = 2 And gMouseClickStatus = "SPC" Then
      gMouseClickStatus = "SPCR"
   End If
End Sub    
'=============================  FncSplitColumn()  ==============================================
Function FncSplitColumn()
    
    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Function
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)  
    
End Function

'=============================  txtIvFrDt_DblClick()  ==============================================
Sub txtIvFrDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtIvFrDt.Action = 7
		Call SetFocusToDocument("M")	
        frm1.txtIvFrDt.Focus
	End If
End Sub
'=============================  txtIvToDt_DblClick()  ==============================================
Sub txtIvToDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtIvToDt.Action = 7
		Call SetFocusToDocument("M")	
        frm1.txtIvToDt.Focus
	End If
End Sub

'=============================  vspdData_GotFocus()  ==============================================
Sub vspdData_GotFocus()
    ggoSpread.Source = frm1.vspdData
End Sub

'=============================  vspdData_DblClick()  ==============================================
Function vspdData_DblClick(ByVal Col, ByVal Row)
	If frm1.vspdData.MaxRows > 0 Then
		If frm1.vspdData.ActiveRow = Row Or frm1.vspdData.ActiveRow > 0 Then
		'	Call CookiePage(1)
		End If
	End If
End Function
'=============================  vspdData_Click()  ==============================================	
Sub vspdData_Click(ByVal Col, ByVal Row)
	
	Set gActiveSpdSheet = frm1.vspdData
    SetPopupMenuItemInf("00000000001")
	
	gMouseClickStatus = "SPC"
	
	
	If frm1.vspdData.MaxRows = 0 Then
		Exit Sub
	End If
	
    If Row <= 0 Then
       
       ggoSpread.Source = frm1.vspdData
       If lgSortKey = 1 Then
			ggoSpread.SSSort Col		'Sort in ascending
			lgSortKey = 2
	   Else
			ggoSpread.SSSort Col, lgSortKey		'Sort in descending
			lgSortKey = 1
       End If
       
       Exit Sub
    End If   
    
    Call SetSpreadColumnValue("A",frm1.vspdData,Col,Row)
    
	IsCookieSplit=""
    With frm1.vspddata
		.Row = Row
		.Col = GetKeyPos("A", 1)
		IsCookieSplit =  IsCookieSplit & Trim(.Text) & parent.gRowSep
	End With
End Sub		
'=============================  vspdData_ColWidthChange()  ==============================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub
'=============================  vspdData_TopLeftChange()  ==============================================
Sub vspdData_TopLeftChange(ByVal OldLeft, ByVal OldTop, ByVal NewLeft, ByVal NewTop)
	
	If OldLeft <> NewLeft Then
	    Exit Sub
	End If
    

	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	    '☜: 재쿼리 체크	
		If lgPageNo <> "" Then		                                                    '다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
			If DbQuery = False Then
				Exit Sub
			End if
		End If
	End If
End Sub
'=============================  txtIvFrDt_KeyDown()  ==============================================
Sub txtIvFrDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub

Sub txtIvToDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub
'=============================  FncQuery()  ==============================================
Function FncQuery() 

    FncQuery = False                                                        '⊙: Processing is NG
    
    Err.Clear                                                               '☜: Protect system from crashing

    ggoSpread.Source = frm1.vspdData
    ggoSpread.ClearSpreadData
    Call InitVariables 														'⊙: Initializes local global variables

	with frm1
        If CompareDateByFormat(.txtIvFrDt.text,.txtIvToDt.text,.txtIvFrDt.Alt,.txtIvToDt.Alt, _
                   "970025",.txtIvFrDt.UserDefinedFormat,parent.gComDateType,False) = False And Trim(.txtIvFrDt.text) <> "" And Trim(.txtIvToDt.text) <> "" then	
           Call DisplayMsgBox("17a003","X","매입등록일","X")		      
           Exit Function
        End if
        	
	End with

    Call DbQuery															'☜: Query db data

    FncQuery = True															'⊙: Processing is OK
	Set gActiveElement = document.activeElement
End Function

'=============================  FncSave()  ==============================================
Function FncSave()     
End Function

'=============================  FncPrint()  ==============================================
Function FncPrint() 
    Call parent.FncPrint()
End Function

'=============================  FncExcel()  ==============================================
Function FncExcel() 
	Call parent.FncExport(parent.C_MULTI)
End Function

'=============================  FncFind()  ==============================================
Function FncFind() 
    Call parent.FncFind(parent.C_MULTI , False)                                     '☜:화면 유형, Tab 유무 
End Function
'=============================  FncExit()  ==============================================
Function FncExit()
	
    FncExit = True
End Function
'=============================  DbQuery()  ==============================================
Function DbQuery() 
	Dim strVal

    DbQuery = False
    
    Err.Clear                                                               '☜: Protect system from crashing
	
    If CheckRunningBizProcess = True Then
       Exit Function
	End If
	Call LayerShowHide(1)
    
    With frm1
    
		If lgIntFlgMode = parent.OPMD_UMODE Then

		  	strVal = BIZ_PGM_ID	& "?txtBizArea="	& Trim(.hdnBizArea.value)
		  	strVal = strVal	& "&txtItemCd="			& Trim(.hdnItemCd.value)		
		  	strVal = strVal	& "&txtBpCd="			& Trim(.hdnBpCd.value)
		  	strVal = strVal	& "&txtIvFrDt="			& Trim(.hdnIvFrDt.value)
			strVal = strVal	& "&txtIvToDt="			& Trim(.hdnIvToDt.value)	  	
		  	strVal = strVal	& "&txtPlantCd="		& Trim(.hdnPlantCd.value)
		  	strVal = strVal	& "&txtIvType="			& Trim(.hdnIvType.value)	  	
			strVal = strVal	& "&txtPurGrpCd="		& Trim(.hdnPurGrpCd.value)		
			strVal = strVal	& "&txtPstFlg="			& Trim(.hdncboPstFlg.value)
		Else

		  	strVal = BIZ_PGM_ID	& "?txtBizArea="	& Trim(.txtBizArea.value)
		  	strVal = strVal	& "&txtItemCd="			& Trim(.txtItemCd.value)		
		  	strVal = strVal	& "&txtBpCd="			& Trim(.txtBpCd.value)
		  	strVal = strVal	& "&txtIvFrDt="			& Trim(.txtIvFrDt.Text)
			strVal = strVal	& "&txtIvToDt="			& Trim(.txtIvToDt.Text)	  	
		  	strVal = strVal	& "&txtPlantCd="		& Trim(.txtPlantCd.value)
		  	strVal = strVal	& "&txtIvType="			& Trim(.txtIvType.value)	  	
			strVal = strVal	& "&txtPurGrpCd="		& Trim(.txtPurGrpCd.value)		
			strVal = strVal	& "&txtPstFlg="			& Trim(.cboPstFlg.value)
		End If	
			
			strVal = strVal & "&lgPageNo="		 & lgPageNo   
		    strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")
		    strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
			strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))
        
        Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 
        
    End With
    
    DbQuery = True
    Call SetToolBar("1100000000011111")										'⊙: 버튼 툴바 제어	

End Function

'=============================  DbQueryOk()  ==============================================
Function DbQueryOk()														'☆: 조회 성공후 실행로직 

  	lgBlnFlgChgValue = False
    lgSaveRow        = 1
    lgIntFlgMode = parent.OPMD_UMODE
  
End Function

