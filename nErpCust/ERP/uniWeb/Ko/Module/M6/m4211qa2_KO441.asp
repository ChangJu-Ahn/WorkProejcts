<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Procurement
'*  2. Function Name        : 
'*  3. Program ID           : m4211qa2
'*  4. Program Name         : 통관상세조회 
'*  5. Program Desc         :
'*  6. Modified date(First) : 
'*  7. Modified date(Last)  : 2003.05.29
'*  8. Modifier (First)     : Jin-hyun Shin
'*  9. Modifier (Last)      : Lee Eun Hee
'* 10. Comment              :
'* 11. Common Coding Guide  : 
'**********************************************************************************************
-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- '******************************************  1.1 Inc 선언   **************************************** !-->
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->
<!--'==========================================  1.1.1 Style Sheet  ======================================!-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'==========================================  1.1.2 공통 Include   =================================== !-->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentA.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="vbscript"   SRC="../../inc/Cookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMA_KO441.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript">

Option Explicit					

<!-- #Include file="../../inc/lgvariables.inc" -->	
                               
Dim lgIsOpenPop                                          
Dim lgSaveRow                                           

Const BIZ_PGM_ID 		= "m4211qb2_KO441.asp"                     
Const BIZ_PGM_JUMP_ID1 	= "m4211qa1"
Const BIZ_PGM_JUMP_ID2 	= "m4212ma1"
Const Major_Cd_Incoterms= "B9006"
Const C_MaxKey          = 27					             

Dim iDBSYSDate
Dim EndDate, StartDate

iDBSYSDate = "<%=GetSvrDate%>"

EndDate = UNIConvDateAtoB(iDBSYSDate, parent.gServerDateFormat, parent.gDateFormat)
StartDate = UnIDateAdd("m", -1, EndDate, parent.gDateFormat)	

'==========================================  setCookie()  ======================================
Function setCookie_01()

	if frm1.vspdData.maxrows > 0 then
		frm1.vspdData.row = frm1.vspdData.ActiveRow
		frm1.vspdData.col =  GetKeyPos("A", 11)
		WriteCookie "CCNo", Trim(frm1.vspdData.text)

		WriteCookie "txtBeneficiaryCd", Trim(frm1.txtBeneficiaryCd.Value)
		WriteCookie "txtIncotermsCd", Trim(frm1.txtIncotermsCd.Value)
		WriteCookie "txtPurGrpCd", Trim(frm1.txtPurGrpCd.Value)
		WriteCookie "txtIDFrDt", frm1.txtIDFrDt.Text
		WriteCookie "txtIDToDt", frm1.txtIDToDt.Text
		WriteCookie "txtIPFrDt", frm1.txtIPFrDt.Text
		WriteCookie "txtIPToDt", frm1.txtIPToDt.Text
	end if
	
	Call PgmJump(BIZ_PGM_JUMP_ID1)

End Function

Function setCookie_02()

Const CookieSplit = 4875

	if frm1.vspdData.maxrows > 0 then
		frm1.vspdData.row = frm1.vspdData.ActiveRow
		frm1.vspdData.col =  GetKeyPos("A", 11)
		if Trim(frm1.vspdData.text) <> "" then
			WriteCookie CookieSplit, Trim(frm1.vspdData.text)
		end if
	end if
	
	Call PgmJump(BIZ_PGM_JUMP_ID2)

End Function

Function GetCookies()

Dim strQueryFlg

	if ReadCookie("CCNo") <> "" then
		strQueryFlg					= ReadCookie("CCNo")
		frm1.txtBeneficiaryCd.Value	= ReadCookie("txtBeneficiaryCd")
		frm1.txtPurGrpCd.Value		= ReadCookie("txtPurGrpCd")
		frm1.txtIncotermsCd.Value	= ReadCookie("txtIncotermsCd")
		frm1.txtCCNo.Value	= ReadCookie("CCNo")
		frm1.txtIDFrDt.Text	= ReadCookie("txtIDFrDt")
		frm1.txtIDToDt.Text	= ReadCookie("txtIDToDt")
		frm1.txtIPFrDt.Text	= ReadCookie("txtIPFrDt")
		frm1.txtIPToDt.Text	= ReadCookie("txtIPToDt")
		
		WriteCookie "CCNo",""
		WriteCookie "txtBeneficiaryCd",""
		WriteCookie "txtPurGrpCd",""
		WriteCookie "txtIncotermsCd",""
		WriteCookie "txtIDFrDt",""
		WriteCookie "txtIDToDt",""
		WriteCookie "txtIPFrDt",""
		WriteCookie "txtIPToDt",""
	end if
	
	if strQueryFlg <> "" then Call dbQuery

End Function

'==========================================  2.1.1 InitVariables()  ======================================
Sub InitVariables()
    lgPageNo         = ""
	lgBlnFlgChgValue = False
    lgSortKey        = 1
    lgSaveRow        = 0
    lgIntFlgMode = Parent.OPMD_CMODE 
End Sub
'==========================================  2.2.1 SetDefaultVal()  ========================================
Sub SetDefaultVal()
	
  	frm1.txtIDFrDt.Text	= StartDate
	frm1.txtIDToDt.Text	= EndDate
	frm1.txtIPFrDt.Text	= StartDate
	frm1.txtIPToDt.Text	= EndDate
	frm1.txtBeneficiaryCd.focus	

	If lgPGCd <> "" then 
		Call ggoOper.SetReqAttr(frm1.txtPurGrpCd, "Q") 
		frm1.txtPurGrpCd.Tag = left(frm1.txtPurGrpCd.Tag,1) & "4" & mid(frm1.txtPurGrpCd.Tag,3,len(frm1.txtPurGrpCd.Tag))
        frm1.txtPurGrpCd.value = lgPGCd
	End If

	If lgPLCd <> "" then 
		Call ggoOper.SetReqAttr(frm1.txtPlantCd, "Q") 
		frm1.txtPlantCd.Tag = left(frm1.txtPlantCd.Tag,1) & "4" & mid(frm1.txtPlantCd.Tag,3,len(frm1.txtPlantCd.Tag))
        frm1.txtPlantCd.value = lgPLCd
	End If

	Set gActiveElement = document.activeElement

End Sub
'======================================================================================
' Function Name : InitComboBox()
'========================================================================================
Sub InitComboBox()
	Call SetCombo(frm1.cboPrcFlg, "T", "진단가")
	Call SetCombo(frm1.cboPrcFlg, "F", "가단가")
End Sub
'======================================================================================
' Function Name : LoadInfTB19029
'========================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "M", "NOCOOKIE", "QA") %>
	<% Call LoadBNumericFormatA("Q", "M", "NOCOOKIE", "QA") %>
End Sub

'======================= 2.2.3 InitSpreadSheet() ========================================
Sub InitSpreadSheet()
    
    Call SetZAdoSpreadSheet("M4211QA201","S","A","V20030329", Parent.C_SORT_DBAGENT, frm1.vspdData, C_MaxKey, "X", "X" )
	Call SetSpreadLock("A") 
End Sub

'================================== 2.2.4 SetSpreadLock() ==================================================
Sub SetSpreadLock(ByVal pOpt)
    If pOpt = "A" Then
		ggoSpread.Source = frm1.vspdData
		ggoSpread.SpreadLockWithOddEvenRowColor()
	End If
End Sub

'------------------------------------------  OpenItemCd()  -------------------------------------------------
Function OpenItemCd()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim iCalledAspName
	Dim IntRetCD
	
	If lgIsOpenPop = True Then Exit Function
	
	if  Trim(frm1.txtPlantCd.Value) = "" then
		Call DisplayMsgBox("17A002", "X", "공장", "X")
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
	arrField(2) = 3 ' -- Spec
    
	iCalledAspName = AskPRAspName("B1B11PA3")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "B1B11PA3", "X")
		lgIsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(parent.window, arrParam, arrField, arrHeader), _
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
'------------------------------------------  OpenPlantCd()  -------------------------------------------------
Function OpenPlantCd()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function
    If frm1.txtPlantCd.className = "protected" Then Exit Function
    
	lgIsOpenPop = True

	arrParam(0) = "공장"						<%' 팝업 명칭 %>
	arrParam(1) = "B_PLANT"      					<%' TABLE 명칭 %>
	arrParam(2) = Trim(frm1.txtPlantCd.Value)		<%' Code Condition%>
'	arrParam(3) = Trim(frm1.txtPlantNm.Value)		<%' Name Cindition%>
	arrParam(4) = ""								<%' Where Condition%>
	arrParam(5) = "공장"						<%' TextBox 명칭 %>
	
    arrField(0) = "PLANT_CD"						<%' Field명(0)%>
    arrField(1) = "PLANT_NM"						<%' Field명(1)%>
    
    arrHeader(0) = "공장"						<%' Header명(0)%>
    arrHeader(1) = "공장명"						<%' Header명(1)%>
    
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
'------------------------------------------  OpenBeneficiary()  -------------------------------------------------
Function OpenBeneficiary()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	arrParam(0) = "수출자"					
	arrParam(1) = "B_Biz_Partner"				
	arrParam(2) = Trim(frm1.txtBeneficiaryCd.Value)		
'	arrParam(3) = Trim(frm1.txtBpNm.Value)		
	arrParam(4) = "BP_TYPE in (" & FilterVar("S", "''", "S") & " ," & FilterVar("CS", "''", "S") & ") "					
	arrParam(5) = "수출자"					
	
    arrField(0) = "BP_CD"						
    arrField(1) = "BP_NM"						
    
    arrHeader(0) = "수출자"					
    arrHeader(1) = "수출자명"				
    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtBeneficiaryCd.focus	
		Set gActiveElement = document.activeElement
		Exit Function
	Else
		frm1.txtBeneficiaryCd.Value = arrRet(0)
		frm1.txtBeneficiaryNm.Value = arrRet(1)
		frm1.txtBeneficiaryCd.focus	
		Set gActiveElement = document.activeElement
	End If	
End Function

'------------------------------------------  OpenIncoterms()  -------------------------------------------------
Function OpenIncoterms()

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	arrParam(0) = "가격조건"					
	arrParam(1) = "B_Minor"			
	arrParam(2) = Trim(frm1.txtIncotermsCd.Value)	
'	arrParam(3) = Trim(frm1.txtPoTypeNm.Value)	
	arrParam(4) = "Major_Cd=  " & FilterVar(Major_Cd_Incoterms , "''", "S") & ""
	arrParam(5) = "가격조건"					
	
    arrField(0) = "Minor_Cd"						
    arrField(1) = "Minor_Nm"						
        
    arrHeader(0) = "가격조건"					
    arrHeader(1) = "가격조건명"					
    
    arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtIncotermsCd.focus	
		Set gActiveElement = document.activeElement
		Exit Function
	Else
		frm1.txtIncotermsCd.Value = arrRet(0)
		frm1.txtIncotermsNm.Value = arrRet(1)
		frm1.txtIncotermsCd.focus	
		Set gActiveElement = document.activeElement
	End If	
End Function

'------------------------------------------  OpenPurGrp()  -------------------------------------------------
Function OpenPurGrp()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function
    If frm1.txtPurGrpCd.className = "protected" Then Exit Function

	lgIsOpenPop = True

	arrParam(0) = "구매그룹"	
	arrParam(1) = "B_Pur_Grp"				
	
	arrParam(2) = Trim(frm1.txtPurGrpCd.Value)
'	arrParam(3) = Trim(frm1.txtPurGrpNm.Value)	
	
	arrParam(4) = ""
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
		Set gActiveElement = document.activeElement
		Exit Function
	Else
		frm1.txtPurGrpCd.Value = arrRet(0)
		frm1.txtPurGrpNm.Value = arrRet(1)
		frm1.txtPurGrpCd.focus	
		Set gActiveElement = document.activeElement
	End If	

End Function 
'------------------------------------  PopZAdoConfigGrid()  ----------------------------------------------
Sub PopZAdoConfigGrid()
	If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
		Exit Sub
	End If
	
	Call OpenGroupPopup("A")
End Sub

'========================================================================================================
' Function Name : OpenGroupPopup
'========================================================================================================
Function OpenGroupPopup(ByVal pSpdNo)
	Dim arrRet
	
	On Error Resume Next
	
	If lgIsOpenPop = True Then Exit Function
	lgIsOpenPop = True

	arrRet = window.showModalDialog("../../ComAsp/ZADOSortPopup.asp",Array(ggoSpread.GetXMLData("A"),gMethodText),"dialogWidth=" & parent.SORTW_WIDTH & "px; dialogHeight=" & parent.SORTW_HEIGHT & "px;; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False
	
	If arrRet(0) = "X" Then
	   Exit Function
	Else
	   Call ggoSpread.SaveXMLData(pSpdNo,arrRet(0),arrRet(1))
       Call InitVariables
       Call InitSpreadSheet()       
   End If
End Function

'==========================================  3.1.1 Form_Load()  ======================================
Sub Form_Load()

    Call LoadInfTB19029	
   	Call ggoOper.FormatField(Document, "A",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
    Call ggoOper.LockField(Document, "N")                                   
    Call InitVariables	 													
    Call GetValue_ko441()
    Call SetDefaultVal	
	Call InitSpreadSheet()
	Call SetToolbar("1100000000001111")		
	Call GetCookies()
    'Call InitComboBox()
	Set gActiveElement = document.activeElement
    
End Sub
'========================================  Form_QueryUnload()  ==================================
Sub Form_QueryUnload(Cancel , UnloadMode )
   
End Sub
'========================================  vspdData_MouseDown()  ==================================
Sub vspdData_MouseDown(Button , Shift , x , y)

   If Button = 2 And gMouseClickStatus = "SPC" Then
      gMouseClickStatus = "SPCR"
   End If
End Sub    
'========================================  FncSplitColumn()  ==================================
Sub FncSplitColumn()

    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   

End Sub

'========================================  OCX_EVENT()  ==================================
Sub txtIDFrDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtIDFrDt.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtIDFrDt.focus
	End If
End Sub

Sub txtIDToDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtIDToDt.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtIDToDt.focus
	End If
End Sub

Sub txtIDFrDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub

Sub txtIDToDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub

Sub txtIPFrDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtIPFrDt.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtIPFrDt.focus
	End If
End Sub

Sub txtIPToDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtIPToDt.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtIPToDt.focus
	End If
End Sub

Sub txtIPFrDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub

Sub txtIPToDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub
'========================================  vspdData_GotFocus()  ==================================
Sub vspdData_GotFocus()
    ggoSpread.Source = frm1.vspdData
End Sub
'========================================  vspdData_DblClick()  ==================================
Function vspdData_DblClick(ByVal Col, ByVal Row)
	If Row = 0 Or Frm1.vspdData.MaxRows = 0 Then 
      Exit Function
    End If
	If frm1.vspdData.MaxRows > 0 Then
		If frm1.vspdData.ActiveRow = Row Or frm1.vspdData.ActiveRow > 0 Then
		'	Call CookiePage(1)
		End If
	End If
End Function
'========================================  FncSplitColumn()  ==================================	
Sub vspdData_Click(ByVal Col, ByVal Row)
   
    Call SetPopupMenuItemInf("00000000001")		
	gMouseClickStatus = "SPC"   
	Set gActiveSpdSheet = frm1.vspdData
	   
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
End Sub
'========================================  vspdData_TopLeftChange()  ==================================	
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
'*******************************  5.1 Toolbar(Main)에서 호출되는 Function *******************************
Function FncQuery() 

    FncQuery = False                                            
    
    Err.Clear                                                   

    ggoSpread.Source = frm1.vspdData
    ggoSpread.ClearSpreadData					
    Call InitVariables 											
    

	with frm1
		if (UniConvDateToYYYYMMDD(.txtIDFrDt.text,gDateFormat,"") > UniConvDateToYYYYMMDD(.txtIDToDt.text,Parent.gDateFormat,"")) And Trim(.txtIDFrDt.text) <> "" And Trim(.txtIDToDt.text) <> "" then	
			Call DisplayMsgBox("17a003", "X","신고일", "X")	
			Exit Function
		End if   
		if (UniConvDateToYYYYMMDD(.txtIPFrDt.text,gDateFormat,"") > UniConvDateToYYYYMMDD(.txtIPToDt.text,Parent.gDateFormat,"")) And Trim(.txtIPFrDt.text) <> "" And Trim(.txtIPToDt.text) <> "" then	
			Call DisplayMsgBox("17a003", "X","면허일", "X")	
			Exit Function
		End if   
	End with
    If DbQuery = False Then Exit Function

    FncQuery = True	
    Set gActiveElement = document.activeElement												

End Function
'====================================  FncSave()  ============================================
Function FncSave()     
End Function
'====================================  FncPrint()  ============================================
'========================================================================================
Function FncPrint() 
    Call parent.FncPrint()
End Function
'====================================  FncExcel()  ============================================
Function FncExcel() 
	Call parent.FncExport(Parent.C_MULTI)
End Function
'====================================  FncFind()  ============================================
Function FncFind() 
    Call parent.FncFind(Parent.C_MULTI , False)                            
End Function
'====================================  FncExit()  ============================================
Function FncExit()
    FncExit = True
End Function
'====================================  DbQuery()  ============================================
Function DbQuery() 
	Dim strVal

    DbQuery = False
    
    Err.Clear                                                       
	If LayerShowHide(1) = False Then
	     Exit Function
	End If 
    
    With frm1
    If lgIntFlgMode = parent.OPMD_UMODE Then
		strVal = BIZ_PGM_ID & "?txtBpCd=" & Trim(.hdnBeneficiaryCd.value)
	    strVal = strVal & "&txtIncotermsCd=" & Trim(.hdnIncotermsCd.value)
	    strVal = strVal & "&txtPurGrpCd=" & Trim(.hdnPurGrpCd.value)
    	strVal = strVal & "&txtIDFrDt=" & Trim(.hdnIDFrDt.value)
    	strVal = strVal & "&txtIDToDt=" & Trim(.hdnIDToDt.value)
    	strVal = strVal & "&txtIPFrDt=" & Trim(.hdnIPFrDt.value)    	
    	strVal = strVal & "&txtIPToDt=" & Trim(.hdnIPToDt.value)
	    strVal = strVal & "&txtPlantCd=" & Trim(.hdnPlantCd.value)
	    strVal = strVal & "&txtItemCd=" & Trim(.hdnItemCd.value)
	    strVal = strVal & "&txtCCNo=" & Trim(.hdnCCNo.value)
        
  	else
	    strVal = BIZ_PGM_ID & "?txtBpCd=" & Trim(.txtBeneficiaryCd.value)
	    strVal = strVal & "&txtIncotermsCd=" & Trim(.txtIncotermsCd.value)
	    strVal = strVal & "&txtPurGrpCd=" & Trim(.txtPurGrpCd.value)
    	strVal = strVal & "&txtIDFrDt=" & Trim(.txtIDFrDt.Text)
    	strVal = strVal & "&txtIDToDt=" & Trim(.txtIDToDt.Text)
    	strVal = strVal & "&txtIPFrDt=" & Trim(.txtIPFrDt.Text)    	
    	strVal = strVal & "&txtIPToDt=" & Trim(.txtIPToDt.Text)
	    strVal = strVal & "&txtPlantCd=" & Trim(.txtPlantCd.value)
	    strVal = strVal & "&txtItemCd=" & Trim(.txtItemCd.value)
	    strVal = strVal & "&txtCCNo=" & Trim(.txtCCNo.value)
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
'====================================  DbQueryOk()  ============================================
Function DbQueryOk()												

	lgBlnFlgChgValue = False
    lgSaveRow        = 1
	lgIntFlgMode = Parent.OPMD_UMODE
	Call vspdData_Click(1,1)
    
    If frm1.vspdData.MaxRows > 0 Then
		frm1.vspdData.Focus
	Else
		frm1.txtBeneficiaryCd.focus	
	End If						

End Function

</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
</HEAD>
<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR>
		<TD <%=HEIGHT_TYPE_00%>></TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>통관상세</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=*>&nbsp;</TD>
					<TD WIDTH="*" align=right></td>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=*>
		<TD WIDTH=100% CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH=100%>
						<FIELDSET>
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS="TD5" NOWRAP>수출자</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="수출자" NAME="txtBeneficiaryCd"  SIZE=10 MAXLENGTH=10 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSpplCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenBeneficiary()">
														   <INPUT TYPE=TEXT NAME="txtBeneficiaryNm" SIZE=20 tag="14"></TD>
									<TD CLASS="TD5" NOWRAP>신고일</TD>
									<TD CLASS="TD6" NOWRAP>
										<table cellspacing=0 cellpadding=0>
											<tr>
												<td>
													<script language =javascript src='./js/m4211qa2_fpDateTime2_txtIDFrDt.js'></script>
												</td>
												<td>~</td>
												<td>
													<script language =javascript src='./js/m4211qa2_fpDateTime2_txtIDToDt.js'></script>
												</td>
											</tr>
										</table>
							         </TD>				   
								</TR>					   
								<TR>
								    <TD CLASS="TD5" NOWRAP>공장</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="공장"  NAME="txtPlantCd" SIZE=10 LANG="ko" MAXLENGTH=4 tag="1XNXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenPlantCd() ">
														   <INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=20 tag="14"></TD>
									<TD CLASS="TD5" NOWRAP>면허일</TD>
									<TD CLASS="TD6" NOWRAP>
										<table cellspacing=0 cellpadding=0>
											<tr>
												<td>
													<script language =javascript src='./js/m4211qa2_fpDateTime2_txtIPFrDt.js'></script>
												</td>
												<td>~</td>
												<td>
													<script language =javascript src='./js/m4211qa2_fpDateTime2_txtIPToDt.js'></script>
												</td>
											</tr>
										</table>
							         </TD>
	                            </TR>
	                            <TR>
									<TD CLASS="TD5" NOWRAP>품목</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT Alt="품목" NAME="txtItemCd" SIZE=10 MAXLENGTH=18  tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenItemCd()">
														   <INPUT TYPE=TEXT Alt="품목" NAME="txtItemNm" SIZE=20 tag="14"></TD>
									<TD CLASS="TD5" NOWRAP>구매그룹</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT Alt="구매그룹" NAME="txtPurGrpCd" SIZE=10 MAXLENGTH=4  tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPurGrp" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPurGrp()">
														   <INPUT TYPE=TEXT NAME="txtPurGrpNm" SIZE=20 tag="14"></TD>	
								</TR>								
								<TR>
									<TD CLASS="TD5" NOWRAP>가격조건</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="가격조건"  NAME="txtIncotermsCd" SIZE=10 LANG="ko" MAXLENGTH=5 tag="1XNXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPoType" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenIncoterms() ">
														   <INPUT TYPE=TEXT NAME="txtIncotermsNm" SIZE=20 tag="14"></TD>
									<TD CLASS="TD5" NOWRAP>통관관리번호</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="통관관리번호" NAME="txtCCNo" SIZE="34" MAXLENGTH=18 tag="1XNXXU"></TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=* VALIGN=TOP>
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD CLASS=TD5 NOWRAP>총통관수량</TD>
								<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/m4211qa2_fpDoubleSingle1_txtTotQty.js'></script></TD>
								<TD CLASS=TD6 NOWRAP></TD>
								<TD CLASS=TD6 NOWRAP></TD>
							</TR>
							<TR>
								<TD HEIGHT=100% WIDTH=100% COLSPAN=4>
									<script language =javascript src='./js/m4211qa2_vaSpread1_vspdData.js'></script>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>		
			</TABLE>
		</TD>
	</TR>
	
    <TR>
      <TD <%=HEIGHT_TYPE_01%>></TD>
    </TR>
    <TR HEIGHT="20">
		<TD WIDTH="100%">
			<TABLE <%=LR_SPACE_TYPE_30%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD WIDTH="*" ALIGN="RIGHT"><a ONCLICK="VBSCRIPT:setCookie_01()">통관현황조회</a> | <a ONCLICK="VBSCRIPT:setCookie_02()">통관내역등록</a></TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
    </TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0  TABINDEX=-1></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="hdnBeneficiaryCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnIncotermsCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnPurGrpCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnIDFrDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnIDToDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnIPFrDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnIPToDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnPlantCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnItemCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnCCNo" tag="24">
</FORM>
    <DIV ID="MousePT" NAME="MousePT">
    <iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
    </DIV>
</BODY>
</HTML>
