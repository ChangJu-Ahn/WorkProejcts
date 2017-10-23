<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Procurement
'*  2. Function Name        : 
'*  3. Program ID           : M5114QA3
'*  4. Program Name         : 입고대비매입실적조회 
'*  5. Program Desc         :
'*  6. Modified date(First) : 2003-05-21
'*  7. Modified date(Last)  : 
'*  8. Modifier (First)     : Kim Jin Ha
'*  9. Modifier (Last)      : 
'* 10. Comment              :
'* 11. Common Coding Guide  : 
'**********************************************************************************************
-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- '#########################################################################################################
'												1. 선 언 부 
'##########################################################################################################!-->
<!-- '******************************************  1.1 Inc 선언   **********************************************
'	기능: Inc. Include
'********************************************************************************************************* !-->
<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->
<!--'==========================================  1.1.1 Style Sheet  ======================================
'==========================================================================================================!-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'==========================================  1.1.2 공통 Include   ======================================
'==========================================================================================================!-->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentA.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs">   </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMA_KO441.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript">

Option Explicit					
'==============================================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'==============================================================================================================================

Dim lgIsOpenPop                                          
Dim lgMark                                                
Dim IscookieSplit 
Dim lgSaveRow                                           
Dim iDBSYSDate
Dim EndDate, StartDate

'==============================================================================================================================
Const BIZ_PGM_ID 		= "M5114QB3_KO441.asp"                     
Const BIZ_PGM_JUMP_ID1 	= "m5211qa1"                         
Const C_MaxKey          = 30					             
'==============================================================================================================================
iDBSYSDate = "<%=GetSvrDate%>"

EndDate = UNIConvDateAtoB(iDBSYSDate, parent.gServerDateFormat, parent.gDateFormat)
StartDate = UnIDateAdd("m", -1, EndDate, parent.gDateFormat)
'==============================================================================================================================
Function setCookie_02()

	if frm1.vspdData.maxrows > 0 then
		frm1.vspdData.row = frm1.vspdData.ActiveRow
		frm1.vspdData.col =  GetKeyPos("A", 11)
		WriteCookie "BlNo", Trim(frm1.vspdData.Text)
	end if
	
	Call PgmJump(BIZ_PGM_JUMP_ID2)

End Function
'==============================================================================================================================
Function setCookie_01()

	Dim strCfmFlg

	if frm1.vspdData.maxrows > 0 then
		if frm1.rdoCfmFlg0.checked then
			strCfmFlg = ""
		elseif frm1.rdoCfmFlg1.checked then
			strCfmFlg = "Y"
		else
			strCfmFlg = "N"
		end if		
		frm1.vspdData.row = frm1.vspdData.ActiveRow
		frm1.vspdData.col =  GetKeyPos("A", 10)
		WriteCookie "BlNo", Trim(frm1.vspdData.Text)
		WriteCookie "txtBeneficiaryCd", Trim(frm1.txtBeneficiaryCd.Value)
		WriteCookie "txtIncotermsCd", Trim(frm1.txtIncotermsCd.Value)
		WriteCookie "txtPurGrpCd", Trim(frm1.txtPurGrpCd.Value)
		WriteCookie "rdoCfmFlg", strCfmFlg
		WriteCookie "txtBlIssueFrDt", frm1.txtBlIssueFrDt.Text
		WriteCookie "txtBlIssueToDt", frm1.txtBlIssueToDt.Text
		WriteCookie "txtLoadingFrDt", frm1.txtLoadingFrDt.Text
		WriteCookie "txtLoadingToDt", frm1.txtLoadingToDt.Text
	end if
	
	Call PgmJump(BIZ_PGM_JUMP_ID1)

End Function
'==============================================================================================================================
Function GetCookies()

	Dim strCfmFlg

	if ReadCookie("BlNo") <> "" then
		frm1.txtBlNo.Value			= ReadCookie("BlNo")
		frm1.txtBeneficiaryCd.Value	= ReadCookie("txtBeneficiaryCd")
		frm1.txtPurGrpCd.Value		= ReadCookie("txtPurGrpCd")
		frm1.txtIncotermsCd.Value	= ReadCookie("txtIncotermsCd")
		strCfmFlg					= ReadCookie("rdoCfmFlg")
		frm1.txtBlIssueFrDt.Text	= ReadCookie("txtBlIssueFrDt")
		frm1.txtBlIssueToDt.Text	= ReadCookie("txtBlIssueToDt")
		frm1.txtLoadingFrDt.Text	= ReadCookie("txtLoadingFrDt")
		frm1.txtLoadingToDt.Text	= ReadCookie("txtLoadingToDt")
		
		if	strCfmFlg = "" then
			frm1.rdoCfmFlg0.checked = true
		elseif strCfmFlg = "Y" then
			frm1.rdoCfmFlg1.checked = true
		else
			frm1.rdoCfmFlg2.checked = true
		end if	

		WriteCookie "BlNo",""
		WriteCookie "txtBeneficiaryCd",""
		WriteCookie "txtPurGrpCd",""
		WriteCookie "txtIncotermsCd",""
		WriteCookie "txtBlIssueFrDt",""
		WriteCookie "txtBlIssueToDt",""
		WriteCookie "txtLoadingFrDt",""
		WriteCookie "txtLoadingToDt",""
	end if
	
	if Trim(frm1.txtBLNo.Value) <> "" then Call dbQuery

End Function
'==============================================================================================================================
Sub InitVariables()
    lgStrPrevKey     = ""
	lgBlnFlgChgValue = False
    lgSortKey        = 1
    lgSaveRow        = 0
	lgIntFlgMode = parent.OPMD_CMODE 
    lgPageNo         = ""
End Sub
'==============================================================================================================================
Sub SetDefaultVal()
	frm1.txtMvFrDt.Text	= StartDate
	frm1.txtMvToDt.Text	= EndDate
	If lgPLCd <> "" then 
		Call ggoOper.SetReqAttr(frm1.txtPlantCd, "Q") 
		frm1.txtPlantCd.Tag = left(frm1.txtPlantCd.Tag,1) & "4" & mid(frm1.txtPlantCd.Tag,3,len(frm1.txtPlantCd.Tag))
        frm1.txtPlantCd.value = lgPLCd
	End If
End Sub
'==============================================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q","M","NOCOOKIE","QA") %>
	<% Call LoadBNumericFormatA("Q", "M", "NOCOOKIE", "QA") %>
End Sub
'==============================================================================================================================
Sub InitSpreadSheet()
	Call SetZAdoSpreadSheet("M5114QA3","S","A","V20031201", Parent.C_SORT_DBAGENT, frm1.vspdData, C_MaxKey, "X", "X" )
	Call SetSpreadLock 
End Sub
'==============================================================================================================================
Sub SetSpreadLock()
		frm1.vspdData.ReDraw = False
		'ggoSpread.SpreadLock 1 , -1
		ggoSpread.SpreadLockWithOddEvenRowColor()
		ggoSpread.spreadUnlock 	GetKeyPos("A", 28), -1
		frm1.vspdData.ReDraw = True
End Sub
'==============================================================================================================================
Function OpenMvmtType()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim iCalledAspName
	Dim IntRetCD
	
	If lgIsOpenPop = True Or UCase(frm1.txtMvmtType.className) = UCase(Parent.UCN_PROTECTED) Then Exit Function

	lgIsOpenPop = True
	
	arrParam(0) = "입고형태"	
	'arrParam(1) ="( select distinct  IO_Type_Cd, io_type_nm from  M_CONFIG_PROCESS a,  m_mvmt_type b where a.rcpt_type = b.io_type_cd    and a.sto_flg = 'N' AND a.USAGE_FLG='Y' ) c"
	arrParam(1) = "( select distinct  IO_Type_Cd, io_type_nm from  M_CONFIG_PROCESS a,  m_mvmt_type b where a.rcpt_type = b.io_type_cd    and a.sto_flg = " & FilterVar("N", "''", "S") & "  AND a.USAGE_FLG=" & FilterVar("Y", "''", "S") & "  and ((b.RCPT_FLG=" & FilterVar("Y", "''", "S") & "  AND b.RET_FLG=" & FilterVar("N", "''", "S") & " ) or (b.RET_FLG=" & FilterVar("N", "''", "S") & "  And b.SUBCONTRA_FLG=" & FilterVar("N", "''", "S") & " )) ) c"
	arrParam(2) = Trim(frm1.txtMvmtType.Value)
	arrParam(3) = ""
	arrParam(4) = ""
	arrParam(5) = "입고형태"			
	
    arrField(0) = "IO_Type_Cd"
    arrField(1) = "IO_Type_NM"
    
    arrHeader(0) = "입고형태"		
    arrHeader(1) = "입고형태명"
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
				
	lgIsOpenPop = False
	
	If isEmpty(arrRet) Then Exit Function				'페이지를 찾을 수 없는 에러발생시.
	If arrRet(0) = "" Then
		frm1.txtMvmtType.focus	
		Set gActiveElement = document.activeElement
		Exit Function
	Else
		frm1.txtMvmtType.Value	= arrRet(0)		
		frm1.txtMvmtTypeNm.Value= arrRet(1)
		lgBlnFlgChgValue = True
		frm1.txtMvmtType.focus	
		Set gActiveElement = document.activeElement
	End If	
End Function
'==============================================================================================================================
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
		Set gActiveElement = document.activeElement
		Exit Function
	Else
		frm1.txtIvType.Value = arrRet(0)
		frm1.txtIvTypeNm.Value = arrRet(1)
		frm1.txtIvType.focus	
		Set gActiveElement = document.activeElement
	End If	
End Function
'==============================================================================================================================
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
		Set gActiveElement = document.activeElement
		Exit Function
	Else
		frm1.txtBpCd.Value = arrRet(0)
		frm1.txtBpNm.Value = arrRet(1)
		frm1.txtBpCd.focus	
		Set gActiveElement = document.activeElement
	End If	
End Function
'==============================================================================================================================
Function OpenPoNo()
	
	Dim strRet
	Dim iCalledAspName
	Dim IntRetCD
	Dim arrParam(5), arrField(6), arrHeader(6)
		
	If lgIsOpenPop = True Or UCase(frm1.txtPoNo.className) = UCase(parent.UCN_PROTECTED) Then Exit Function
		
	lgIsOpenPop = True
	
	arrParam(0) = ""					
	arrParam(1) = "Y"
	arrParam(2) = ""
	
	iCalledAspName = AskPRAspName("M3111PA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "M3111PA1", "X")
		lgIsOpenPop = False
		Exit Function
	End If
	
	strRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If strRet(0) = "" Then
		frm1.txtPoNo.focus	
		Set gActiveElement = document.activeElement
		Exit Function
	Else
		frm1.txtPoNo.value = strRet(0)
		frm1.txtPoNo.focus	
		Set gActiveElement = document.activeElement
	End If

End Function
'==============================================================================================================================
Function OpenMvmtNo()
	
		Dim strRet
		Dim arrParam(3)
		Dim iCalledAspName
		Dim IntRetCD
	
		If lgIsOpenPop = True Or UCase(frm1.txtMvmtNo.className) = UCase(Parent.UCN_PROTECTED) Then Exit Function
		
		lgIsOpenPop = True

		arrParam(0) = ""'Trim(frm1.hdnSupplierCd.Value)
		arrParam(1) = ""'Trim(frm1.hdnGroupCd.Value)
		arrParam(2) = ""'Trim(frm1.hdnMvmtType.Value)		
		arrParam(3) = ""'This is for Inspection check, must be nothing.
		
		iCalledAspName = AskPRAspName("M4111PA3")
	
		If Trim(iCalledAspName) = "" Then
			IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "M4111PA3", "X")
			lgIsOpenPop = False
			Exit Function
		End If
	
		strRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
			"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
		
		lgIsOpenPop = False
		
		If isEmpty(strRet) Then Exit Function				'페이지를 찾을 수 없는 에러발생시.
		
		If strRet(0) = "" Then
			frm1.txtMvmtNo.focus	
			Set gActiveElement = document.activeElement
			Exit Function
		Else
			frm1.txtMvmtNo.value = strRet(0)
			frm1.txtMvmtNo.focus	
			Set gActiveElement = document.activeElement
		End If	
		
End Function
'==============================================================================================================================
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
'==============================================================================================================================
Function OpenItemCd()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function
	
	if  Trim(frm1.txtPlantCd.Value) = "" then
		Call DisplayMsgBox("17A002", "X", "공장", "X")
		frm1.txtPlantCd.focus	
		Set gActiveElement = document.activeElement
		Exit Function
	End if
	
	lgIsOpenPop = True

	arrParam(0) = "품목"						
	arrParam(1) = "B_Item_By_Plant,B_Plant,B_Item"	
	arrParam(2) = Trim(frm1.txtItemCd.Value)		
	arrParam(3) = ""
	arrParam(4) = "B_Item_By_Plant.Plant_Cd = B_Plant.Plant_Cd And "
	arrParam(4) = arrParam(4) & "B_Item_By_Plant.Item_Cd = B_Item.Item_Cd and B_Item.phantom_flg = " & FilterVar("N", "''", "S") & "  "
	if Trim(frm1.txtPlantCd.Value)<>"" then
		arrParam(4) = arrParam(4) & "And B_Plant.Plant_Cd= " & FilterVar(frm1.txtPlantCd.Value, "''", "S") & ""    
	End if
	arrParam(5) = "품목"						

    arrField(0) = "B_Item.Item_Cd"					
    arrField(1) = "B_Item.Item_NM"					
    arrField(2) = "B_Plant.Plant_Cd"				
    arrField(3) = "B_Plant.Plant_NM"				
    
    arrHeader(0) = "품목"						
    arrHeader(1) = "품목명"						
    arrHeader(2) = "공장"						
    arrHeader(3) = "공장명"						
    
	arrRet = window.showModalDialog("../m1/m1111pa1.asp", Array(parent.window, arrParam, arrField, arrHeader), _
		"dialogWidth=695px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

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
'==============================================================================================================================
Function OpenGLRef(ByVal strGlType)

	Dim strRet
	Dim arrParam(1)
	Dim iCalledAspName
	Dim IntRetCD
	
	If lgIsOpenPop = True Then Exit Function
		
	lgIsOpenPop = True
	
	frm1.vspdData.Row = frm1.vspdData.ActiveRow
	frm1.vspdData.Col =  GetKeyPos("A", 27)
	
	arrParam(0) = Trim(frm1.vspdData.Text)
	arrParam(1) = ""
	
   If strGlType = "A" Then               '회계전표팝업 
		iCalledAspName = AskPRAspName("a5120ra1")
		If Trim(iCalledAspName) = "" Then
			IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a5120ra1" ,"x")
			lgIsOpenPop = False
			Exit Function
		End If
	   strRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam), _
			"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

    Elseif strGlType = "T" Then          '결의전표팝업 
		iCalledAspName = AskPRAspName("a5130ra1")
		If Trim(iCalledAspName) = "" Then
			IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a5130ra1" ,"x")
			lgIsOpenPop = False
			Exit Function
		End If
	   strRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam), _
			"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

    Elseif strGlType = "B" Then
     	Call DisplayMsgBox("205154","X" , "X","X")   '아직 전표가 생성되지 않았습니다. 
    End if

	lgIsOpenPop = False
	
End Function
'==============================================================================================================================
Function OpenIVRef()
	
	Dim strRet
	Dim arrParam(5)
	Dim iCalledAspName
	Dim iCurRow
	
	if lgIntFlgMode <> Parent.OPMD_UMODE then
		Call DisplayMsgBox("900002", "X","X","매입현황" )
		frm1.txtMvmtType.focus	
		Set gActiveElement = document.activeElement
		Exit Function
	End if 
	
	iCurRow = frm1.vspdData.ActiveRow
	
	Call frm1.vspdData.GetText(GetKeyPos("A", 29),	iCurRow,	arrParam(0))	'입고번호 
	Call frm1.vspdData.GetText(GetKeyPos("A", 22),	iCurRow,	arrParam(1))	'발주번호 
	Call frm1.vspdData.GetText(GetKeyPos("A", 2),	iCurRow,	arrParam(2))	'공장 
	Call frm1.vspdData.GetText(GetKeyPos("A", 3),	iCurRow,	arrParam(3))	'공장명 
	Call frm1.vspdData.GetText(GetKeyPos("A", 4),	iCurRow,	arrParam(4))	'품목 
	Call frm1.vspdData.GetText(GetKeyPos("A", 5),	iCurRow,	arrParam(5))	'품목명 
		
	If lgIsOpenPop = True Then Exit Function
		
	lgIsOpenPop = True
		
	iCalledAspName = AskPRAspName("M5114RA3")
		
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "M5114RA3", "X")
		lgIsOpenPop = False
		Exit Function
	End If
		
	strRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
			
	lgIsOpenPop = False
		
End Function
'==============================================================================================================================
Sub PopZAdoConfigGrid()
	If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
		Exit Sub
	End If
	
	Call OpenGroupPopup("A")
End Sub
'==============================================================================================================================
Function OpenGroupPopup(ByVal pSpdNo)

	Dim arrRet

	On Error Resume Next
	
	If lgIsOpenPop = True Then Exit Function
	lgIsOpenPop = True

	arrRet = window.showModalDialog("../../ComAsp/ZADOSortPopup.asp",Array(ggoSpread.GetXMLData(pSpdNo),gMethodText),"dialogWidth=" & parent.SORTW_WIDTH & "px; dialogHeight=" & parent.SORTW_HEIGHT & "px;; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False
	
	If arrRet(0) = "X" Then
	   Exit Function
	Else
	   Call ggoSpread.SaveXMLData(pSpdNo,arrRet(0),arrRet(1))
       Call InitVariables
       Call InitSpreadSheet()       
   End If
   
End Function
'==============================================================================================================================
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
	Else
		msgbox "^*^__ ING."	
	End IF

End Function
'==============================================================================================================================
Sub Form_Load()

    Call LoadInfTB19029														
    Call ggoOper.FormatField(Document, "A",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
    Call ggoOper.LockField(Document, "N")                                   
	Call InitVariables														
    Call GetValue_ko441()
	Call SetDefaultVal	
	Call InitSpreadSheet()
    Call SetToolbar("1100000000001111")		
	frm1.txtMvmtType.focus
    Set gActiveElement = document.activeElement 
End Sub
'==============================================================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)

   If Button = 2 And gMouseClickStatus = "SPC" Then
      gMouseClickStatus = "SPCR"
   End If
End Sub    
'==============================================================================================================================
Function FncSplitColumn()
    
    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Function
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   
End Function
'==============================================================================================================================
Sub txtMvFrDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtMvFrDt.Action = 7
		Call SetFocusToDocument("M") 
		frm1.txtMvFrDt.focus
	End If
End Sub
'==============================================================================================================================
Sub txtMvToDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtMvToDt.Action = 7
		Call SetFocusToDocument("M") 
		frm1.txtMvToDt.focus
	End If
End Sub
'==============================================================================================================================
Sub txtMvFrDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call FncQuery()
End Sub
'==============================================================================================================================
Sub txtMvToDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call FncQuery()
End Sub
'==============================================================================================================================
Sub vspdData_GotFocus()
    ggoSpread.Source = frm1.vspdData
End Sub
'==============================================================================================================================
Function vspdData_DblClick(ByVal Col, ByVal Row)
	If frm1.vspdData.MaxRows > 0 Then
		If frm1.vspdData.ActiveRow = Row Or frm1.vspdData.ActiveRow > 0 Then
		'	Call CookiePage(1)
		End If
	End If
End Function
'==============================================================================================================================	
Sub vspdData_Click(ByVal Col, ByVal Row)
   
    gMouseClickStatus = "SPC"   
    Set gActiveSpdSheet = frm1.vspdData
    
	Call SetPopupMenuItemInf("00000000001")
    
    If Row = 0 Then
        ggoSpread.Source = frm1.vspdData
        If lgSortKey = 1 Then
            ggoSpread.SSSort, lgSortKey
            lgSortKey = 2
        Else
            ggoSpread.SSSort, lgSortKey
            lgSortKey = 1
        End If    
    End If
	
	IscookieSplit = ""
End Sub
'==============================================================================================================================	
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
   
	ggoSpread.Source = frm1.vspdData
	Call OpenGLRef("A")
End Sub
'==============================================================================================================================	
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
'==============================================================================================================================
Function FncQuery() 

    FncQuery = False                                            
    
    Err.Clear                                                   

    Call ggoOper.ClearField(Document, "2")						
    Call InitVariables 											


	with frm1
		if (UniConvDateToYYYYMMDD(.txtMvFrDt.text,Parent.gDateFormat,"") > UniConvDateToYYYYMMDD(.txtMvToDt.text,Parent.gDateFormat,"")) And Trim(.txtMvFrDt.text) <> "" And Trim(.txtMvToDt.text) <> "" then	
			Call DisplayMsgBox("17a003", "X","입고일", "X")	
			Exit Function
		End if   
	End with


    If DbQuery = False Then Exit Function

    FncQuery = True													
	Set gActiveElement = document.activeElement 
End Function
'==============================================================================================================================
Function FncPrint() 
    Call parent.FncPrint()
    Set gActiveElement = document.activeElement 
End Function
'==============================================================================================================================
Function FncExcel() 
	Call parent.FncExport(parent.C_MULTI)
	Set gActiveElement = document.activeElement 
End Function
'==============================================================================================================================
Function FncFind() 
    Call parent.FncFind(parent.C_MULTI , False)     
    Set gActiveElement = document.activeElement                        
End Function
'==============================================================================================================================
Function FncExit()
    FncExit = True
    Set gActiveElement = document.activeElement 
End Function
'==============================================================================================================================
Function DbQuery() 
	Dim strVal
	Dim strCfmFlg

    DbQuery = False
    
    Err.Clear                                                       
	
    If  LayerShowHide(1) = False Then
       	Exit Function
    End If
	
    
    With frm1
		
		if .rdoCfmFlg0.checked then
			strCfmFlg = ""
		elseif .rdoCfmFlg1.checked then
			strCfmFlg = "Y"
		else
			strCfmFlg = "N"
		end if
	
		If lgIntFlgMode = parent.OPMD_UMODE Then	    
			
			strVal = BIZ_PGM_ID & "?txtMvmtType=" & Trim(.hdnMvmtType.value)
		    strVal = strVal & "&txtIvType=" & Trim(.hdnIvType.value)
		    strVal = strVal & "&txtBpCd=" & Trim(.hdnBpCd.value)
			strVal = strVal & "&txtMvFrDt=" & Trim(.hdnMvFrDt.Value)
			strVal = strVal & "&txtMvToDt=" & Trim(.hdnMvToDt.Value)
			strVal = strVal & "&txtPoNo=" & Trim(.hdnPoNo.value)
			strVal = strVal & "&txtMvmtNo=" & Trim(.hdnMvmtNo.value)
			strVal = strVal & "&txtPlantCd=" & Trim(.hdnPlantCd.value)
		    strVal = strVal & "&txtItemCd=" & Trim(.hdnItemCd.value)
		    strVal = strVal & "&txtCfmFlg=" & Trim(.hdnstrCfmFlg.value)
		Else
			strVal = BIZ_PGM_ID & "?txtMvmtType=" & Trim(.txtMvmtType.value)
		    strVal = strVal & "&txtIvType=" & Trim(.txtIvType.value)
		    strVal = strVal & "&txtBpCd=" & Trim(.txtBpCd.value)
			strVal = strVal & "&txtMvFrDt=" & Trim(.txtMvFrDt.Text)
			strVal = strVal & "&txtMvToDt=" & Trim(.txtMvToDt.Text)
			strVal = strVal & "&txtPoNo=" & Trim(.txtPoNo.value)
			strVal = strVal & "&txtMvmtNo=" & Trim(.txtMvmtNo.value)
			strVal = strVal & "&txtPlantCd=" & Trim(.txtPlantCd.value)
		    strVal = strVal & "&txtItemCd=" & Trim(.txtItemCd.value)
		    strVal = strVal & "&txtCfmFlg=" & strCfmFlg
		End If
		
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
End Function
'==============================================================================================================================
Function DbQueryOk()												

	lgBlnFlgChgValue = False
    lgSaveRow        = 1
	lgIntFlgMode = parent.OPMD_UMODE
	
	Call SetToolbar("1100000000011111")	
	
	If frm1.vspdData.MaxRows > 0 Then
		frm1.vspdData.focus
	Else
		frm1.txtMvmtType.focus
	End IF
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>입고대비매입실적</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=*>&nbsp;</TD>
					<TD WIDTH=* align=right><A href="vbscript:OpenIVRef()" >매입현황</A>&nbsp;</TD>
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
									<TD CLASS="TD5" NOWRAP>입고형태</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT Alt="입고형태" NAME="txtMvmtType" SIZE=10 MAXLENGTH=5 tag="1XNXXU" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnMoveType" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenMvmtType()" OnMouseOver="vbscript:PopUpMouseOver()" OnMouseOut="vbscript:PopUpMouseOut()">
														   <INPUT TYPE=TEXT Alt="입고형태" NAME="txtMvmtTypeNm" SIZE=20 tag="24X"></TD>
									<TD CLASS="TD5" NOWRAP>매입형태</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="매입형태" NAME="txtIvType" SIZE=10 MAXLENGTH=5 tag="1XNXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnIvType" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenIvType() ">
														   <INPUT TYPE=TEXT NAME="txtIvTypeNm" SIZE=20 tag="14"></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>공급처</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="공급처" NAME="txtBpCd" SIZE=10 MAXLENGTH=10 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBpCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenBpCd()">
														   <INPUT TYPE=TEXT NAME="txtBpNm" SIZE=20 tag="14"></TD>			
									<TD CLASS="TD5" NOWRAP>입고일</TD>
									<TD CLASS="TD6" NOWRAP>
										<table cellpadding=0 cellspacing=0>
											<tr>
												<td NOWRAP>
													<script language =javascript src='./js/m5114qa3_fpDateTime2_txtMvFrDt.js'></script>
												</td>
												<td NOWRAP>~</td>
												<td NOWRAP>
												   <script language =javascript src='./js/m5114qa3_fpDateTime2_txtMvToDt.js'></script>
												</td>
											</tr>
										</table>
									</TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>발주번호</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtPoNo" SIZE=32 MAXLENGTH=18 ALT="발주번호" tag="1XNXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnORGCd1" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPoNo()"></TD>
									<TD CLASS="TD5" NOWRAP>입고번호</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtMvmtNo" SIZE=32 MAXLENGTH=18 ALT="입고번호"  tag="1XNXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnMvmtNo" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenMvmtNo()"></TD>
								</TR>
								<TR>
								    <TD CLASS="TD5" NOWRAP>공장</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="공장"  NAME="txtPlantCd" SIZE=10 LANG="ko" MAXLENGTH=4 tag="1XNXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenPlantCd() ">
														   <INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=20 tag="14"></TD>
									<TD CLASS="TD5" NOWRAP>품목</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT Alt="품목" NAME="txtItemCd" SIZE=32 MAXLENGTH=18  tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenItemCd()"></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>매입확정여부</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=radio AlT="확정여부" NAME="rdoCfmFlg" ID="rdoCfmFlg0" CLASS="RADIO" checked tag="11"><label for="rdoCfmFlg0">&nbsp;전체&nbsp;&nbsp;</label>
														   <INPUT TYPE=radio AlT="확정여부" NAME="rdoCfmFlg" ID="rdoCfmFlg1" CLASS="RADIO" tag="11"><label for="rdoCfmFlg1">&nbsp;확정&nbsp;</label>
														   <INPUT TYPE=radio AlT="확정여부" NAME="rdoCfmFlg" ID="rdoCfmFlg2" CLASS="RADIO" tag="11"><label for="rdoCfmFlg2">&nbsp;미확정&nbsp;&nbsp;</label></TD>
									<TD CLASS="TD5" NOWRAP></TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT Alt="품목" NAME="txtItemNm" SIZE=32 tag="14"></TD>
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
								<TD HEIGHT=100% WIDTH=100% COLSPAN=4>
									<script language =javascript src='./js/m5114qa3_vaSpread1_vspdData.js'></script>
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
    <TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>

<INPUT TYPE=HIDDEN NAME="hdnMvmtType" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnIvType" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnBpCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnMvFrDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnMvToDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnPoNo" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnMvmtNo" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnPlantCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnItemCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnstrCfmFlg" tag="24">

</FORM>
    <DIV ID="MousePT" NAME="MousePT">
    <iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
    </DIV>
</BODY>
</HTML>
