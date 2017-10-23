<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Procurement
'*  2. Function Name        : 
'*  3. Program ID           :
'*  4. Program Name         :
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2001/01/18
'*  8. Modified date(Last)  : 2005/11/28
'*  9. Modifier (First)     : Min, Hak-jun
'* 10. Modifier (Last)      : Sim Hae Young
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                            
'**********************************************************************************************
-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!--'******************************************  1.1 Inc 선언   ****************************************-->
<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->
<!--'==========================================  1.1.1 Style Sheet  ======================================-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'==========================================  1.1.2 공통 Include   ====================================-->
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

Option Explicit													'☜: indicates that All variables must be declared in advance

<!-- #Include file="../../inc/lgvariables.inc" -->	

Dim lgIsOpenPop                                             <%'☜: Popup화면의 상태 저장변수               %>
Dim IscookieSplit 
Dim lgSaveRow                                               <%'☜: Cookie용을 변수                         %> 

Dim iDBSYSDate
Dim EndDate, StartDate

iDBSYSDate = "<%=GetSvrDate%>"
EndDate = UNIConvDateAtoB(iDBSYSDate, parent.gServerDateFormat, parent.gDateFormat)
StartDate = UnIDateAdd("m", -1, EndDate, parent.gDateFormat)

Const BIZ_PGM_ID 		= "m5111qb1_KO441.asp"                         '☆: 비지니스 로직 ASP명 
Const BIZ_PGM_JUMP_ID 	= "m5111qa2"                             '☆: Cookie에서 사용할 상수 
Const C_MaxKey          = 23							         '☆☆☆☆: Max key value

'==========================================  2.1.1 InitVariables()  ======================================
Sub InitVariables()
    lgPageNo         = ""
    lgBlnFlgChgValue = False
    lgSortKey        = 1
    lgSaveRow        = 0
    lgIntFlgMode = parent.OPMD_CMODE 
End Sub
'==========================================  2.2.1 SetDefaultVal()  ========================================
Sub SetDefaultVal()
	frm1.txtIvFrDt.Text	= StartDate
	frm1.txtIvToDt.Text	= EndDate
	If lgPLCd <> "" then 
		Call ggoOper.SetReqAttr(frm1.txtPlantCd, "Q") 
		frm1.txtPlantCd.Tag = left(frm1.txtPlantCd.Tag,1) & "4" & mid(frm1.txtPlantCd.Tag,3,len(frm1.txtPlantCd.Tag))
        frm1.txtPlantCd.value = lgPLCd
	End If
	If lgPGCd <> "" then 
		Call ggoOper.SetReqAttr(frm1.txtPurGrpCd, "Q") 
		frm1.txtPurGrpCd.Tag = left(frm1.txtPurGrpCd.Tag,1) & "4" & mid(frm1.txtPurGrpCd.Tag,3,len(frm1.txtPurGrpCd.Tag))
        frm1.txtPurGrpCd.value = lgPGCd
	End If
	If lgBACd <> "" then 
		Call ggoOper.SetReqAttr(frm1.txtBizArea, "Q") 
		frm1.txtBizArea.Tag = left(frm1.txtBizArea.Tag,1) & "4" & mid(frm1.txtBizArea.Tag,3,len(frm1.txtBizArea.Tag))
        frm1.txtBizArea.value = lgBACd
	End If
End Sub
'==========================================  InitComboBox()  =========================================
Sub InitComboBox()
			Call SetCombo(frm1.cboPstFlg, "Y", "Y")
			Call SetCombo(frm1.cboPstFlg, "N", "N")
End Sub

'==========================================  LoadInfTB19029()  =========================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "M", "NOCOOKIE", "QA") %>
	<% Call LoadBNumericFormatA("Q", "M", "NOCOOKIE", "QA")%>
End Sub
'======================= 2.2.3 InitSpreadSheet() ========================================
Sub InitSpreadSheet()
    Call SetZAdoSpreadSheet("M5111QA1","G","A","V20030913",parent.C_GROUP_DBAGENT,frm1.vspdData, _
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

	arrParam(0) = "사업장"					<%' 팝업 명칭 %>
	arrParam(1) = "B_BIZ_AREA"					<%' TABLE 명칭 %>
	arrParam(2) = Trim(frm1.txtBizArea.Value)	<%' Code Condition%>
	arrParam(5) = "사업장"					<%' TextBox 명칭 %>

    arrField(0) = "BIZ_AREA_CD"					<%' Field명(0)%>
    arrField(1) = "BIZ_AREA_NM"					<%' Field명(1)%>
    
    
    arrHeader(0) = "사업장"					<%' Header명(0)%>
    arrHeader(1) = "사업장명"				<%' Header명(1)%>    
    
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

	arrParam(0) = "공급처"						<%' 팝업 명칭 %>
	arrParam(1) = "B_Biz_Partner"					<%' TABLE 명칭 %>
	arrParam(2) = Trim(frm1.txtBpCd.Value)		<%' Code Condition%>
'	arrParam(3) = Trim(frm1.txtBpNm.Value)		<%' Name Cindition%>
	arrParam(4) = "BP_TYPE In (" & FilterVar("S", "''", "S") & " ," & FilterVar("CS", "''", "S") & ") And usage_flag=" & FilterVar("Y", "''", "S") & " "					<%' Where Condition%>
	arrParam(5) = "공급처"						<%' TextBox 명칭 %>
	
    arrField(0) = "BP_CD"							<%' Field명(0)%>
    arrField(1) = "BP_NM"							<%' Field명(1)%>
    
    arrHeader(0) = "공급처"						<%' Header명(0)%>
    arrHeader(1) = "공급처명"					<%' Header명(1)%>
    
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

	arrParam(0) = "매입형태"						<%' 팝업 명칭 %>
	arrParam(1) = "M_IV_TYPE"							<%' TABLE 명칭 %>
	arrParam(2) = Trim(frm1.txtIvType.Value)			<%' Code Condition%>
'	arrParam(3) = Trim(frm1.txtIvTypeNm.Value)			<%' Name Condition%>
	arrParam(4) = ""									<%' Where Condition%>
	arrParam(5) = "매입형태"						<%' TextBox 명칭 %>
	
    arrField(0) = "IV_TYPE_CD"							<%' Field명(0)%>
    arrField(1) = "IV_TYPE_NM"							<%' Field명(1)%>
        
    arrHeader(0) = "매입형태"						<%' Header명(0)%>
    arrHeader(1) = "매입형태명"						<%' Header명(1)%>
    
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

'------------------------------------  PopZAdoConfigGrid()  ----------------------------------------------
Sub PopZAdoConfigGrid()
	If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
		Exit Sub
	End If
	
	Call OpenGroupByPopup("A")
End Sub
'------------------------------------  OpenGroupByPopup()  ----------------------------------------------
Function OpenGroupByPopup(ByVal pSpdNo)

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

'==========================================   CookiePage()  ======================================
Function CookiePage(ByVal Kubun)

	Dim strTemp, arrVal
	Dim strCookie, i
	Dim strToDt
	Dim strAddMonthToDt

	Const CookieSplit = 4877						<% 'Cookie Split String : CookiePage Function Use%>

	If Kubun = 1 Then								<% 'Jump로 화면을 이동할 경우 %>

		Call vspdData_Click(frm1.vspdData.ActiveCol,frm1.vspdData.ActiveRow)		
		WriteCookie CookieSplit , IsCookieSplit		<% 'Jump로 화면을 이동할때 필요한 Cookie 변수정의 %>
		
		If Len(Trim(frm1.txtBizArea.value)) Then		
			WriteCookie "tBizArea",Trim(frm1.txtBizArea.value) 
		Else
			WriteCookie "tBizArea",""
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
		
		If Len(Trim(frm1.txtIvFrDt.text)) Then
			WriteCookie "IvFrDt",Trim(frm1.txtIvFrDt.text) 
		Else
			WriteCookie "IvFrDt",""
		End If
		
		If Len(Trim(frm1.txtIvToDt.text)) Then
			WriteCookie "IvToDt",Trim(frm1.txtIvToDt.text) 
		Else
			WriteCookie "IvToDt",""
		End If
		
		If Len(Trim(frm1.txtPlantCd.value)) Then		
			WriteCookie "PlantCd",Trim(frm1.txtPlantCd.value) 
		Else
			WriteCookie "PlantCd",""
		End If
		
		If Len(Trim(frm1.txtIvType.value)) Then
			WriteCookie "IvType",Trim(frm1.txtIvType.value) 
		Else
			WriteCookie "IvType",""
		End If
		
		If Len(Trim(frm1.txtPurGrpCd.value)) Then
			WriteCookie "PurGrpCd",Trim(frm1.txtPurGrpCd.value) 
		Else
			WriteCookie "PurGrpCd",""
		End If
						
		If Len(Trim(frm1.cboPstFlg.value)) Then
			WriteCookie "PstFlg",Trim(frm1.cboPstFlg.value) 
		Else
			WriteCookie "PstFlg",""
		End If
					
		Call PgmJump(BIZ_PGM_JUMP_ID)
		
		

	ElseIf Kubun = 0 Then							<% 'Jump로 화면이 이동해 왔을경우 %>
		If Trim(ReadCookie("CookieIoIvFlg")) = "Y" Then
	 		frm1.txtIvFrDt.Text	= UNIConvDateAtoB(Trim(ReadCookie("CookieFromDt")), parent.gServerDateFormat, parent.gDateFormat)
		 	strToDt	= UNIConvDateAtoB(Trim(ReadCookie("CookieToDt")), parent.gServerDateFormat, parent.gDateFormat)
		 	strAddMonthToDt = UnIDateAdd("m", 1, strToDt, parent.gDateFormat)
		 	frm1.txtIvToDt.Text	= UnIDateAdd("d", -1, strAddMonthToDt, parent.gDateFormat)
		 	frm1.txtBpCd.value	= Trim(ReadCookie("CookieBpCd"))
		 	frm1.txtBpNm.value	= Trim(ReadCookie("CookieBpNm"))

			WriteCookie "CookieIoIvFlg",""
			WriteCookie "CookieFromDt",""
			WriteCookie "CookieToDt",""
			WriteCookie "CookieBpCd",""
			WriteCookie "CookieBpNm",""

			Call MainQuery()
			Exit Function
		End If
		strTemp = ReadCookie(CookieSplit)

		If strTemp = "" then Exit Function

		arrVal = Split(strTemp, parent.gRowSep)

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

'==========================================  3.1.1 Form_Load()  ======================================
Sub Form_Load()
	Call LoadInfTB19029							
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
    Call ggoOper.LockField(Document, "N")       
	Call InitVariables							
    Call GetValue_ko441()
	Call SetDefaultVal	
	Call InitSpreadSheet()
    Call SetToolbar("1100000000001111")	
    Call InitComboBox()
	Call CookiePage(0)
    frm1.txtBizArea.focus
    Set gActiveElement = document.activeElement
	
End Sub
'===========================================  txtIvFrDt_DblClick()  ====================================
Sub txtIvFrDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtIvFrDt.Action = 7
		Call SetFocusToDocument("M")	
        frm1.txtIvFrDt.Focus
	End If
End Sub
'===========================================  txtIvToDt_DblClick()  ====================================
Sub txtIvToDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtIvToDt.Action = 7
		Call SetFocusToDocument("M")	
        frm1.txtIvToDt.Focus
	End If
End Sub
'===========================================  vspdData_GotFocus()  ====================================
Sub vspdData_GotFocus()
    ggoSpread.Source = frm1.vspdData
End Sub
'===========================================  vspdData_DblClick()  ====================================
Function vspdData_DblClick(ByVal Col, ByVal Row)
	If frm1.vspdData.MaxRows > 0 Then
		If frm1.vspdData.ActiveRow = Row Or frm1.vspdData.ActiveRow > 0 Then
		'	Call CookiePage(1)
		End If
	End If
End Function
'===========================================  vspdData_Click()  ====================================
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
	IscookieSplit = ""	
	Dim ii
    frm1.vspdData.Row = Row
	For ii = 1 to C_MaxKey
        frm1.vspdData.Col = GetKeyPos("A", ii)
		IscookieSplit = IscookieSplit & Trim(frm1.vspdData.Text) & Parent.gRowSep 
	Next
    
End Sub	
'===========================================  vspdData_ColWidthChange()  ====================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub
'===========================================  vspdData_MouseDown()  ====================================
Sub vspdData_MouseDown(Button , Shift , x , y)

   If Button = 2 And gMouseClickStatus = "SPC" Then
      gMouseClickStatus = "SPCR"
   End If
End Sub    
'===========================================  FncSplitColumn()  ====================================
Function FncSplitColumn()
    
    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Function
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)  
    
End Function
'===========================================  vspdData_TopLeftChange()  ====================================
Sub vspdData_TopLeftChange(ByVal OldLeft, ByVal OldTop, ByVal NewLeft, ByVal NewTop)
	
	If OldLeft <> NewLeft Then
	    Exit Sub
	End If
    
    If CheckRunningBizProcess = True Then
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

'===========================================  txtIvFrDt_KeyDown()  ====================================
Sub txtIvFrDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub

Sub txtIvToDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub
'===========================================  FncQuery()  ====================================
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
'===========================================  FncSave()  ====================================
Function FncSave()     
End Function
'===========================================  FncPrint()  ====================================
Function FncPrint() 
    Call parent.FncPrint()
End Function

'===========================================  FncExcel()  ====================================
Function FncExcel() 
	Call parent.FncExport(parent.C_MULTI)
	Set gActiveElement = document.activeElement
End Function
'===========================================  FncFind()  ====================================
Function FncFind() 
    Call parent.FncFind(parent.C_MULTI , False)                                     <%'☜:화면 유형, Tab 유무 %>
    Set gActiveElement = document.activeElement
End Function

'===========================================  FncExit()  ====================================
Function FncExit()
	FncExit = True
	Set gActiveElement = document.activeElement
End Function

'===========================================  DbQuery()  ====================================
Function DbQuery() 
	Dim strVal

    DbQuery = False
    
    Err.Clear                                                               '☜: Protect system from crashing
	
    If LayerShowHide(1) = False then
		Exit function
	End If
    
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

        strVal = strVal & "&gBizArea=" & lgBACd 
        strVal = strVal & "&gPlant=" & lgPLCd 
        strVal = strVal & "&gPurGrp=" & lgPGCd 
        strVal = strVal & "&gPurOrg=" & lgPOCd  
	
        Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 
        
    End With
    
    DbQuery = True
    Call SetToolBar("1100000000011111")										'⊙: 버튼 툴바 제어	

End Function

'===========================================  DbQueryOk()  ====================================
Function DbQueryOk()														'☆: 조회 성공후 실행로직 

	lgBlnFlgChgValue = False
    lgSaveRow        = 1
    lgIntFlgMode = parent.OPMD_UMODE

End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kCM.inc" -->	
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>매입내역집계</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=*>&nbsp;</TD>
					<!--<TD WIDTH="*" align=right><button name="btnAutoSel" class="clsmbtn" ONCLICK="OpenGroupByPopup()">집계순서</button></td>-->
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
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
								    <TD CLASS="TD5" NOWRAP>사업장</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="사업장" NAME="txtBizArea" SIZE=10 LANG="ko" MAXLENGTH=10 tag="1XNXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBizArea" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenBizArea() ">
														   <INPUT TYPE=TEXT NAME="txtBizAreaNm" SIZE=20 tag="14"></TD>
									<TD CLASS="TD5" NOWRAP>품목</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT Alt="품목" NAME="txtItemCd" SIZE=10 MAXLENGTH=18 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenItemCd()">
														   <INPUT TYPE=TEXT Alt="품목" NAME="txtItemNm" SIZE=20 tag="14"></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>공급처</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="공급처" NAME="txtBpCd" SIZE=10 MAXLENGTH=10 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBpCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenBpCd()">
														   <INPUT TYPE=TEXT NAME="txtBpNm" SIZE=20 tag="14"></TD>			
									<TD CLASS="TD5" NOWRAP>매입등록일</TD>
									<TD CLASS="TD6" NOWRAP>
                                        <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT id=fpDateTime2 title=FPDATETIME CLASS=FPDTYYYYMMDD name=txtIvFrDt CLASSID=<%=gCLSIDFPDT%> tag="11X1" ALT="매입등록일"></OBJECT>');</SCRIPT> ~&nbsp
								        <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT id=fpDateTime2 title=FPDATETIME CLASS=FPDTYYYYMMDD name=txtIvToDt CLASSID=<%=gCLSIDFPDT%> ALT="매입등록일" tag="11X1"></OBJECT>');</SCRIPT> </TD>														   


								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>공장</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="공장" NAME="txtPlantCd" SIZE=10 LANG="ko" MAXLENGTH=4 tag="1XNXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenPlantCd() ">
														   <INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=20 tag="14"></TD>
									<TD CLASS="TD5" NOWRAP>매입형태</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="매입형태" NAME="txtIvType" SIZE=10 LANG="ko" MAXLENGTH=5 tag="1XNXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnIvType" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenIvType() ">
														   <INPUT TYPE=TEXT NAME="txtIvTypeNm" SIZE=20 tag="14"></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>구매그룹</TD>							
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT Alt="구매그룹" NAME="txtPurGrpCd" SIZE=10 MAXLENGTH=4 tag="11"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPurGrpCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPurGrpCd()">
														   <INPUT TYPE=TEXT NAME="txtPurGrpNm" SIZE=20 tag="14"></TD>								
									<TD CLASS="TD5" NOWRAP>확정구분</TD>
									<TD CLASS="TD6" NOWRAP><SELECT NAME="cboPstFlg" tag="11"  STYLE="WIDTH: 98px;"><OPTION VALUE="" selected></OPTION></SELECT></TD>
								</TR>								
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=* valign=top>
						<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD HEIGHT="100%">
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=vaSpread1> <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
    <tr>
		<TD <%=HEIGHT_TYPE_01%>></TD>
    </tr>
    <TR HEIGHT="20">
		<TD WIDTH="100%">
			<TABLE <%=LR_SPACE_TYPE_30%>>
				<TR>			 
					<TD WIDTH=10>&nbsp;</TD>
					<TD WIDTH="*" ALIGN="RIGHT"><a ONCLICK="VBSCRIPT:CookiePage(1)">매입내역상세조회</a></TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
    </TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="hdnBizArea" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnItemCd" tag="24">	
<INPUT TYPE=HIDDEN NAME="hdnBpCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnIvFrDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnIvToDt" tag="24">	  	
<INPUT TYPE=HIDDEN NAME="hdnPlantCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnIvType" tag="24">	  	
<INPUT TYPE=HIDDEN NAME="hdnPurGrpCd" tag="24">		
<INPUT TYPE=HIDDEN NAME="hdncboPstFlg" tag="24">
</FORM>
    <DIV ID="MousePT" NAME="MousePT">
    <iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
    </DIV>
</BODY>
</HTML>
