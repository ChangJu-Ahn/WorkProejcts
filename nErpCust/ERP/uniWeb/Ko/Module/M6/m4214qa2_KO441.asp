<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Procurement
'*  2. Function Name        : 
'*  3. Program ID           : M4214QA2
'*  4. Program Name         : 미통관선적상세조회 
'*  5. Program Desc         :
'*  6. Modified date(First) : 
'*  7. Modified date(Last)  : 
'*  8. Modifier (First)     : kangsuhwan
'*  9. Modifier (Last)      : 
'* 10. Comment              :
'* 11. Common Coding Guide  : 
'**********************************************************************************************
-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- '******************************************  1.1 Inc 선언   ****************************************** !-->
<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->
<!--'==========================================  1.1.1 Style Sheet  ======================================!-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'==========================================  1.1.2 공통 Include   ====================================!-->
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

<!-- #Include file="../../inc/lgvariables.inc" -->	

Dim lgIsOpenPop                                          
Dim lgMark                                                
Dim IscookieSplit 
Dim lgSaveRow                                           
Dim iDBSYSDate
Dim EndDate, StartDate


iDBSYSDate = "<%=GetSvrDate%>"

EndDate = UNIConvDateAtoB(iDBSYSDate, parent.gServerDateFormat, parent.gDateFormat)
StartDate = UnIDateAdd("m", -1, EndDate, parent.gDateFormat)

Const BIZ_PGM_ID 		= "M4214QB2_KO441.asp"                     
Const BIZ_PGM_JUMP_ID1 	= "m5211qa1"                         
Const BIZ_PGM_JUMP_ID2 	= "m5212ma1"   
Const Major_Cd_Incoterms= "B9006"
Const C_MaxKey          = 28					             

'==========================================  setCookie()  ======================================
'	Name : setCookie()
'	Description : 
'===============================================================================================
Function setCookie_02()

	if frm1.vspdData.maxrows > 0 then
		frm1.vspdData.row = frm1.vspdData.ActiveRow
		frm1.vspdData.col =  GetKeyPos("A", 11)
		WriteCookie "BlNo", Trim(frm1.vspdData.Text)
	end if
	
	Call PgmJump(BIZ_PGM_JUMP_ID2)

End Function



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

'==========================================  2.1.1 InitVariables()  ======================================
Sub InitVariables()
    lgStrPrevKey     = ""
	lgBlnFlgChgValue = False
    lgSortKey        = 1
    lgSaveRow        = 0
	lgIntFlgMode = parent.OPMD_CMODE 
    lgPageNo         = ""
End Sub

'==========================================  2.2.1 SetDefaultVal()  ========================================
Sub SetDefaultVal()
	frm1.txtBlIssueFrDt.Text	= StartDate
	frm1.txtBlIssueToDt.Text	= EndDate
	frm1.txtLoadingFrDt.Text	= StartDate
	frm1.txtLoadingToDt.Text	= EndDate 

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
<% Call loadInfTB19029A("Q","M","NOCOOKIE","QA") %>
<% Call LoadBNumericFormatA("Q", "M", "NOCOOKIE", "QA") %>
End Sub

'======================= 2.2.3 InitSpreadSheet() ========================================
Sub InitSpreadSheet()
    Call SetZAdoSpreadSheet("M4214QA2","S","A","V20030320", Parent.C_SORT_DBAGENT, frm1.vspdData, C_MaxKey, "X", "X" )

   Call SetSpreadLock 

End Sub

'================================== 2.2.4 SetSpreadLock() ==================================================
Sub SetSpreadLock()
	ggoSpread.Source = frm1.vspdData
	ggoSpread.SpreadLockWithOddEvenRowColor()
End Sub

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
	arrParam(4) = "BP_TYPE In (" & FilterVar("S", "''", "S") & " ," & FilterVar("CS", "''", "S") & ") And usage_flag=" & FilterVar("Y", "''", "S") & " "				
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
		Exit Function
	Else
		frm1.txtIncotermsCd.Value = arrRet(0)
		frm1.txtIncotermsNm.Value = arrRet(1)
		frm1.txtIncotermsCd.focus	
		Set gActiveElement = document.activeElement
	End If	
End Function

'------------------------------------------  OpenPurGrp()  -------------------------------------------------
'	Name : OpenPurGrp()
'	Description : PurGrp PopUp
'---------------------------------------------------------------------------------------------------------
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
		Exit Function
	Else
		frm1.txtPurGrpCd.Value = arrRet(0)
		frm1.txtPurGrpNm.Value = arrRet(1)
		frm1.txtPurGrpCd.focus	
		Set gActiveElement = document.activeElement
	End If	

End Function 
'------------------------------------------  OpenItemCd()  -------------------------------------------------
'	Name : OpenItemCd()
'	Description : Item PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenItemCd()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim iCalledAspName

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
		IsOpenPop = False
		Exit Function
	End If
    
    ' -- 기존 TEMPLATE 과 다름(주석처리)
	'arrRet = window.showModalDialog("../m1/m1111pa1.asp", Array(parent.window, arrParam, arrField, arrHeader), _
	'	"dialogWidth=695px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

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
'	Name : OpenPlantCd()
'	Description : Plant PopUp
'---------------------------------------------------------------------------------------------------------
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

'------------------------------------  PopZAdoConfigGrid()  ----------------------------------------------
Sub PopZAdoConfigGrid()
	If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
		Exit Sub
	End If
	
	Call OpenGroupPopup("A")
End Sub

'------------------------------------  OpenGroupPopup()  ----------------------------------------------
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
<!--
'=======================================  3.2.1 btnBLNoOnClick()  ======================================
-->
Sub btnBLNoOnClick()
	Call OpenBlNoPop()
End Sub
<!--
'+++++++++++++++++++++++++++++++++++++++++++++++  OpenBlNoPop()  +++++++++++++++++++++++++++++++++++++++
-->

Function OpenBlNoPop()
	Dim strRet
	Dim IntRetCD
	Dim iCalledAspName
		
	If lgIsOpenPop = True Then Exit Function

    lgIsOpenPop = True
		
	iCalledAspName = AskPRAspName("M5211PA1")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "M5211PA1", "X")
		lgIsOpenPop = False
		Exit Function
	End If
			
	strRet = window.showModalDialog(iCalledAspName, Array(Window.parent), _
		"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False
		
	If strRet = "" Then
		Exit Function
	Else
		frm1.txtBlNo.value = strRet
		frm1.txtBlNo.focus	
		Set gActiveElement = document.activeElement
	End If	
End Function

'==========================================  3.1.1 Form_Load()  ======================================
Sub Form_Load()

    Call LoadInfTB19029														
    Call ggoOper.FormatField(Document, "A",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
    Call ggoOper.LockField(Document, "N")                                   

	Call InitVariables														
    Call GetValue_ko441()
	Call SetDefaultVal	
	Call InitSpreadSheet()
    Call SetToolbar("1100000000001111")		
	Call GetCookies()
    frm1.txtBeneficiaryCd.focus
     
End Sub
'==========================================================================================
'   Event Name : Form_QueryUnload
'==========================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )
End Sub
'========================================================================================
' Function Name : vspdData_MouseDown
'========================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)

   If Button = 2 And gMouseClickStatus = "SPC" Then
      gMouseClickStatus = "SPCR"
   End If
End Sub    

'========================================================================================
' Function Name : FncSplitColumn
'========================================================================================
Function FncSplitColumn()
    
    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Function
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   
    
End Function


'==========================================================================================
'   Event Name : 
'   Event Desc : Date OCX Double Click, Key Down
'==========================================================================================
Sub txtBlIssueFrDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtBlIssueFrDt.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtBlIssueFrDt.focus
	End If
End Sub

Sub txtBlIssueToDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtBlIssueToDt.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtBlIssueToDt.focus
	End If
End Sub

Sub txtBlIssueFrDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call FncQuery()
End Sub

Sub txtBlIssueToDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call FncQuery()
End Sub


Sub txtLoadingFrDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtLoadingFrDt.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtLoadingFrDt.focus
	End If
End Sub

Sub txtLoadingToDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtLoadingToDt.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtLoadingToDt.focus
	End If
End Sub

Sub txtLoadingFrDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call FncQuery()
End Sub

Sub txtLoadingToDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call FncQuery()
End Sub
'==========================================================================================
'   Event Name : vspdData_GotFocus
'==========================================================================================
Sub vspdData_GotFocus()
    ggoSpread.Source = frm1.vspdData
End Sub

'==========================================================================================
'   Event Name : vspdData_DblClick
'==========================================================================================
Function vspdData_DblClick(ByVal Col, ByVal Row)
	If frm1.vspdData.MaxRows > 0 Then
		If frm1.vspdData.ActiveRow = Row Or frm1.vspdData.ActiveRow > 0 Then
		'	Call CookiePage(1)
		End If
	End If
End Function
	
'======================================================================================================
'   Event Name : vspdData_Click
'=======================================================================================================%>
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
	
'======================================================================================================
'   Event Name : vspdData_TopLeftChange
'=======================================================================================================
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

    Call ggoOper.ClearField(Document, "2")						
    Call InitVariables 											


	with frm1
		if (UniConvDateToYYYYMMDD(.txtBlIssueFrDt.text,Parent.gDateFormat,"") > UniConvDateToYYYYMMDD(.txtBlIssueToDt.text,Parent.gDateFormat,"")) And Trim(.txtBlIssueFrDt.text) <> "" And Trim(.txtBlIssueToDt.text) <> "" then	
			Call DisplayMsgBox("17a003", "X","B/L접수일", "X")	
			Exit Function
		End if   
		if (UniConvDateToYYYYMMDD(.txtLoadingFrDt.text,Parent.gDateFormat,"") > UniConvDateToYYYYMMDD(.txtLoadingToDt.text,Parent.gDateFormat,"")) And Trim(.txtLoadingFrDt.text) <> "" And Trim(.txtLoadingToDt.text) <> "" then	
			Call DisplayMsgBox("17a003", "X","선적일", "X")	
			Exit Function
		End if   
	End with


    If DbQuery = False Then Exit Function

    FncQuery = True													

End Function
'========================================================================================
' Function Name : FncSave
'========================================================================================
Function FncSave()     
End Function
'========================================================================================
' Function Name : FncPrint
'========================================================================================
Function FncPrint() 
    Call parent.FncPrint()
End Function

'========================================================================================
' Function Name : FncExcel
'========================================================================================
Function FncExcel() 
	Call parent.FncExport(parent.C_MULTI)
End Function

'========================================================================================
' Function Name : FncFind
'========================================================================================
Function FncFind() 
    Call parent.FncFind(parent.C_MULTI , False)                            
End Function

'========================================================================================
' Function Name : FncExit
'========================================================================================
Function FncExit()
    FncExit = True
End Function

'========================================================================================
' Function Name : DbQuery
'========================================================================================
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
		
		strVal = BIZ_PGM_ID & "?txtBpCd=" & FilterVar(Trim(.hdnBeneficiaryCd.value),"","SNM")
	    strVal = strVal & "&txtIncotermsCd=" & FilterVar(Trim(.hdnIncotermsCd.value),"","SNM")
	    strVal = strVal & "&txtPurGrpCd=" & FilterVar(Trim(.hdnPurGrpCd.value),"","SNM")
    	strVal = strVal & "&txtBlFrDt=" & Trim(.hdnBlIssueFrDt.value)
    	strVal = strVal & "&txtBlToDt=" & Trim(.hdnBlIssueToDt.value)
    	strVal = strVal & "&txtLoadingFrDt=" & Trim(.hdnLoadingFrDt.value)    	
    	strVal = strVal & "&txtLoadingToDt=" & Trim(.hdnLoadingToDt.value)
    	strVal = strVal & "&txtCfmFlg=" & FilterVar(Trim(.hdnstrCfmFlg.value),"","SNM")
	    strVal = strVal & "&txtItemCd=" & FilterVar(Trim(.hdnItemCd.value),"","SNM")
	    strVal = strVal & "&txtPlantCd=" & FilterVar(Trim(.hdnPlantCd.value),"","SNM")
	    strVal = strVal & "&txtBlNo=" & FilterVar(Trim(.hdnBlNo.value),"","SNM")
        strVal = strVal & "&lgPageNo="   & lgPageNo         
		strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")
        strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
		strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))
	Else
		strVal = BIZ_PGM_ID & "?txtBpCd=" & FilterVar(Trim(.txtBeneficiaryCd.value),"","SNM")
	    strVal = strVal & "&txtIncotermsCd=" & FilterVar(Trim(.txtIncotermsCd.value),"","SNM")
	    strVal = strVal & "&txtPurGrpCd=" & FilterVar(Trim(.txtPurGrpCd.value),"","SNM")
    	strVal = strVal & "&txtBlFrDt=" & Trim(.txtBlIssueFrDt.Text)
    	strVal = strVal & "&txtBlToDt=" & Trim(.txtBlIssueToDt.Text)
    	strVal = strVal & "&txtLoadingFrDt=" & Trim(.txtLoadingFrDt.Text)    	
    	strVal = strVal & "&txtLoadingToDt=" & Trim(.txtLoadingToDt.Text)
    	strVal = strVal & "&txtCfmFlg=" & FilterVar(Trim(strCfmFlg),"","SNM")
	    strVal = strVal & "&txtItemCd=" & FilterVar(Trim(.txtItemCd.value),"","SNM")
	    strVal = strVal & "&txtPlantCd=" & FilterVar(Trim(.txtPlantCd.value),"","SNM")
	    strVal = strVal & "&txtBlNo=" & FilterVar(Trim(.txtBlNo.value),"","SNM")
        strVal = strVal & "&lgPageNo="   & lgPageNo         
		strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")
        strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
		strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))
	End If

        strVal = strVal & "&gBizArea=" & lgBACd 
        strVal = strVal & "&gPlant=" & lgPLCd 
        strVal = strVal & "&gPurGrp=" & lgPGCd 
        strVal = strVal & "&gPurOrg=" & lgPOCd  


        Call RunMyBizASP(MyBizASP, strVal)							
        
    End With
    
    DbQuery = True
    Call SetToolbar("1100000000011111")								

End Function

'========================================================================================
' Function Name : DbQueryOk
'========================================================================================
Function DbQueryOk()												

	lgBlnFlgChgValue = False
    lgSaveRow        = 1
	lgIntFlgMode = parent.OPMD_UMODE

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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>미통관선적상세조회</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=*>&nbsp;</TD>
					<TD WIDTH="*" align=right>&nbsp;</td>
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
									<TD CLASS="TD5" NOWRAP>B/L접수일</TD>
									<TD CLASS="TD6" NOWRAP>
										<table cellspacing=0 cellpadding=0>
											<tr>
												<td>
													<script language =javascript src='./js/m4214qa2_fpDateTime2_txtBlIssueFrDt.js'></script>
												</td>
												<td>~</td>
												<td>
													<script language =javascript src='./js/m4214qa2_fpDateTime2_txtBlIssueToDt.js'></script>
												</td>
											</tr>
										</table>
							         </TD>				   
								</TR>
								<TR>
								    <TD CLASS="TD5" NOWRAP>공장</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="공장"  NAME="txtPlantCd" SIZE=10 LANG="ko" MAXLENGTH=4 tag="1XNXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenPlantCd() ">
														   <INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=20 tag="14"></TD>
									<TD CLASS="TD5" NOWRAP>선적일</TD>
									<TD CLASS="TD6" NOWRAP>
										<table cellspacing=0 cellpadding=0>
											<tr>
												<td>
													<script language =javascript src='./js/m4214qa2_fpDateTime2_txtLoadingFrDt.js'></script>
												</td>
												<td>~</td>
												<td>
													<script language =javascript src='./js/m4214qa2_fpDateTime2_txtLoadingToDt.js'></script>
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
									<TD CLASS="TD5" NOWRAP>B/L관리번호</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="B/L관리번호" NAME="txtBlNo" SIZE="29" MAXLENGTH=18 tag="1XNXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBLNo" ALIGN=top TYPE="BUTTON" onclick="vbscript:btnBLNoOnClick()"></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>확정여부</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=radio AlT="확정여부" NAME="rdoCfmFlg" ID="rdoCfmFlg0" CLASS="RADIO" checked tag="11"><label for="rdoCfmFlg0">&nbsp;전체&nbsp;&nbsp;</label>
														   <INPUT TYPE=radio AlT="확정여부" NAME="rdoCfmFlg" ID="rdoCfmFlg1" CLASS="RADIO" tag="11"><label for="rdoCfmFlg1">&nbsp;확정&nbsp;</label>
														   <INPUT TYPE=radio AlT="확정여부" NAME="rdoCfmFlg" ID="rdoCfmFlg2" CLASS="RADIO" checked tag="11"><label for="rdoCfmFlg2">&nbsp;미확정&nbsp;&nbsp;</label></TD>
									<TD CLASS="TD5" NOWRAP></TD>
									<TD CLASS="TD6" NOWRAP></TD>
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
								<TD CLASS=TD5 NOWRAP>총B/L수량</TD>
								<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/m4214qa2_fpDoubleSingle1_txtTotQty.js'></script></TD>
								<TD CLASS=TD6 NOWRAP></TD>
								<TD CLASS=TD6 NOWRAP></TD>
							</TR>
							<TR>
								<TD HEIGHT=100% WIDTH=100% COLSPAN=4>
									<script language =javascript src='./js/m4214qa2_vaSpread1_vspdData.js'></script>
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
					<TD WIDTH="*" ALIGN="RIGHT">&nbsp;</TD>
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

<INPUT TYPE=HIDDEN NAME="hdnBeneficiaryCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnIncotermsCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnPurGrpCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnBlIssueFrDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnBlIssueToDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnLoadingFrDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnLoadingToDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnstrCfmFlg" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnItemCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnPlantCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnBlNo" tag="24">

</FORM>
    <DIV ID="MousePT" NAME="MousePT">
    <iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
    </DIV>
</BODY>
</HTML>
