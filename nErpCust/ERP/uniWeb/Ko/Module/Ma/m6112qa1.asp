<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Procurement
'*  2. Function Name        : 
'*  3. Program ID           :
'*  4. Program Name         :
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 
'*  8. Modified date(Last)  : 2001/01/19
'*  9. Modifier (First)     : YOON JI YOUNG
'* 10. Modifier (Last)      : 
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
<!--'#########################################################################################################
'												1. 선 언 부 
'##########################################################################################################-->
<!-- '******************************************  1.1 Inc 선언   **********************************************
'	기능: Inc. Include
'********************************************************************************************************* -->
<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->
<!--'==========================================  1.1.1 Style Sheet  ======================================
'==========================================================================================================-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'==========================================  1.1.2 공통 Include   ======================================
'==========================================================================================================-->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentA.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs">   </SCRIPT>
<SCRIPT LANGUAGE="VBScript">

Option Explicit				

'******************************************  1.2 Global 변수/상수 선언  ***********************************
'	1. Constant는 반드시 대문자 표기.
'**********************************************************************************************************
<!-- #Include file="../../inc/lgvariables.inc" -->	
Dim lgIsOpenPop                                            '☜: Popup화면의 상태 저장변수               
Dim IscookieSplit 
Dim lgSaveRow  

Dim iDBSYSDate
Dim EndDate, StartDate

iDBSYSDate = "<%=GetSvrDate%>"
EndDate = UNIConvDateAtoB(iDBSYSDate, parent.gServerDateFormat, parent.gDateFormat)
StartDate = UnIDateAdd("m", -1, EndDate, parent.gDateFormat)

Const BIZ_PGM_ID 		= "m6112qb1.asp"                  
Const BIZ_PGM_JUMP_ID 	= "m6112qa2"                      
Const C_MaxKey          = 14							  
'==============================================================================================================================
Sub InitVariables()
    lgPageNo     = ""
	lgBlnFlgChgValue = False
    lgSortKey        = 1
    lgSaveRow        = 0
    lgIntFlgMode = parent.OPMD_CMODE   
End Sub
'==============================================================================================================================
Sub SetDefaultVal()
	frm1.txtChargeFrDt.Text	= StartDate
	frm1.txtChargeToDt.Text	= EndDate
End Sub
'==============================================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "M", "NOCOOKIE", "QA") %>
End Sub
'==============================================================================================================================
Sub InitSpreadSheet()
	Call SetZAdoSpreadSheet("M6112QA1","G","A","V20030429",parent.C_GROUP_DBAGENT,frm1.vspdData, C_MaxKey, "X","X")
    Call SetSpreadLock 
End Sub
'==============================================================================================================================
Sub SetSpreadLock()
	ggoSpread.Source = frm1.vspdData
	ggoSpread.SpreadLockWithOddEvenRowColor()
End Sub
'==============================================================================================================================
Function OpenBizArea()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function
	
	lgIsOpenPop = True

	arrParam(0) = "사업장"						
	arrParam(1) = "B_BIZ_AREA"	
	arrParam(2) = Trim(frm1.txtBizArea.Value)	
'	arrParam(3) = Trim(frm1.txtBizAreaNm.Value)	
	arrParam(4) = ""							
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
		Set gActiveElement = document.activeElement
		Exit Function
	Else
		frm1.txtBizArea.Value = arrRet(0)
		frm1.txtBizAreaNm.Value = arrRet(1)
		frm1.txtBizArea.focus	
		Set gActiveElement = document.activeElement
	End If	
End Function
'==============================================================================================================================
Function OpenChargeType()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	
	If lgIsOpenPop = True Then Exit Function
	
	lgIsOpenPop = True

	arrParam(0) = "경비항목"			   
	arrParam(1) = "A_JNL_ITEM,b_trade_charge"		
	arrParam(2) = Trim(frm1.txtChargeType.Value)	
'	arrParam(3) = Trim(frm1.txtChargeTypeNm.Value)	
	arrParam(4) = "b_trade_charge.charge_cd=A_JNL_ITEM.JNL_CD And A_JNL_ITEM.JNL_TYPE=" & FilterVar("EC", "''", "S") & " and b_trade_charge.module_type=" & FilterVar("M", "''", "S") & " "
	arrParam(5) = "경비항목"			
	
    arrField(0) = "JNL_CD"				
    arrField(1) = "JNL_NM"				
    
    arrHeader(0) = "경비항목"		
    arrHeader(1) = "경비항목명"		
    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	lgIsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtChargeType.focus	
		Set gActiveElement = document.activeElement		
		Exit Function
	Else
		frm1.txtChargeType.Value = arrRet(0)
		frm1.txtChargeTypeNm.Value = arrRet(1)
		frm1.txtChargeType.focus	
		Set gActiveElement = document.activeElement		
	End If	
End Function
'==============================================================================================================================
Function OpenBpCd()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	arrParam(0) = "지급처"			
	arrParam(1) = "B_Biz_Partner"		
	arrParam(2) = Trim(frm1.txtBpCd.Value)
'	arrParam(3) = Trim(frm1.txtBpNm.Value)
	arrParam(4) = ""					
	arrParam(5) = "지급처"			
	
    arrField(0) = "BP_CD"				
    arrField(1) = "BP_NM"				
    
    arrHeader(0) = "지급처"			
    arrHeader(1) = "지급처명"		
    
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
Function Openprocessstep()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	arrParam(0) = "진행구분"			
	arrParam(1) = "B_minor"				
	arrParam(2) = Trim(frm1.txtprocessstep.value)	
'	arrParam(3) = trim(frm1.txtprocessstepNm.value)	
	arrParam(4) = "major_cd=" & FilterVar("M9014", "''", "S") & ""				
	arrParam(5) = "진행구분"			
	
    arrField(0) = "minor_cd"			
    arrField(1) = "minor_nm"			
    
    arrHeader(0) = "진행구분"		
    arrHeader(1) = "진행구분명"		
    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	lgIsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtprocessstep.focus	
		Set gActiveElement = document.activeElement	
		Exit Function
	Else
		frm1.txtprocessstep.Value = arrRet(0)
		frm1.txtprocessstepNm.Value = arrRet(1)	
		frm1.txtprocessstep.focus	
		Set gActiveElement = document.activeElement	
	End If	

End Function
'==============================================================================================================================
Function OpenItemCd()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim iCalledAspName
	
	If lgIsOpenPop = True Then Exit Function
	
	lgIsOpenPop = True

	arrParam(0) = Trim(frm1.txtItemCd.value)		' Item Code
	arrParam(1) = ""
	arrParam(2) = ""
	arrParam(3) = "30"
	arrParam(4) = ""		'-- 날짜 
	arrParam(5) = ""		'-- 자유(b_item_by_plant a, b_item b: and 부터 시작)
	
	arrField(0) = 1 ' -- 품목코드 
	arrField(1) = 2 ' -- 품목명							' Header명(1)			
    
    iCalledAspName = AskPRAspName("B1B01PA2")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "B1B01PA2", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam,arrField, arrHeader), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtItemCd.focus	
		Set gActiveElement = document.activeElement
		Exit Function
	Else
		frm1.txtItemCd.Value = arrRet(0)
		frm1.txtItemNm.Value = arrRet(1)
	    frm1.hdnplantcd.Value = arrRet(3)
	    frm1.txtItemCd.focus	
		Set gActiveElement = document.activeElement
	End If	
End Function
'==============================================================================================================================
Function OpenPoNo()
	
	Dim strRet
	Dim arrParam(2)
	Dim iCalledAspName
	
	If lgIsOpenPop = True Or UCase(frm1.txtPoNo.className) = UCase(parent.UCN_PROTECTED) Then Exit Function
		
	lgIsOpenPop = True
		
	arrParam(0) = ""  'Return Flag
	arrParam(1) = ""  'Release Flag
	arrParam(2) = ""  'STO Flag
	
	iCalledAspName = AskPRAspName("m3111pa1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", PopupParent.VB_INFORMATION, "m3111pa1", "X")
		lblnWinEvent = False
		Exit Function
	End If
	strRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam), _
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
Sub PopZAdoConfigGrid()
	If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
		Exit Sub
	End If
	
	Call OpenGroupByPopup("A")
End Sub
'==============================================================================================================================
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
'==============================================================================================================================
Function OpenDocumentNo() ' -- 화면에 없는 팝업 - 수불번호팝업 
	Dim arrRet
	Dim Param1, Param2, Param3, Param4
	
	'공장코드가 있는 지 체크 
	If Trim(frm1.txtItemCd.Value) = "" or Trim(frm1.hdnplantcd.Value) = "" then
		Call DisplayMsgBox("17A002","X" , "품목","X")
		Exit Function
	End If
	
	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	Param1 = Trim(frm1.txtDocumentNo.Value)
	Param2 = Trim(frm1.hdnYear.Value)
	Param3 = "R40"
    Param4 = Trim(frm1.hdnplantcd.Value)
	
	arrRet = window.showModalDialog("m6112pa1.asp", Array(Param1, Param2, Param3, Param4), _
		"dialogWidth=705px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	lgIsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtDocumentNo.focus	
		Set gActiveElement           = document.activeElement
		Exit Function
	Else
		frm1.txtDocumentNo.Value    = arrRet(0)	
		frm1.hdnYear.Value          = arrRet(1)
		frm1.txtDocumentNo.focus	
		Set gActiveElement           = document.activeElement
	End If	
End Function
'==============================================================================================================================
Sub setplant()
	Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6, i
	Dim iCodeArr, iNameArr, iRateArr
    Dim item_cd
    Dim stwhere
    
    Err.Clear
    
    item_cd = frm1.txtItemCd.value
	
	stwhere = " B_Item_By_Plant.Plant_Cd = B_Plant.Plant_Cd "
    stwhere = stwhere & " AND  B_Item_By_Plant.Item_Cd = B_Item.Item_Cd " 
    stwhere = stwhere & " AND  B_Item.phantom_flg = " & FilterVar("N", "''", "S") & "  "
    stwhere = stwhere & " AND  B_Item.Item_Cd =  " & FilterVar(item_cd, "''", "S") & "  "

	Call CommonQueryRs(" B_Plant.Plant_Cd,B_Plant.Plant_Nm ","  B_Item_By_Plant,B_Plant,B_Item ",stwhere, lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

	If Err.number <> 0 Then
		MsgBox Err.description 
		Err.Clear 
		Exit Sub
	End If
    
    if Trim(lgF0) <> "" then
        iCodeArr = Split(lgF0, Chr(11))
        iNameArr = Split(lgF1, Chr(11))
        frm1.hdnplantcd.value = iCodeArr(0)
    else 
        frm1.hdnplantcd.value = ""
    end if

End Sub
'==============================================================================================================================
Function CookiePage(ByVal Kubun)

	Dim strTemp, arrVal
	Dim strCookie, i

	Const CookieSplit = 4877					

	If Kubun = 1 Then							

		Call vspdData_Click(frm1.vspdData.ActiveCol,frm1.vspdData.ActiveRow)		
		WriteCookie CookieSplit , IsCookieSplit	
		
		If Len(Trim(frm1.txtBizArea.value)) Then	
			WriteCookie "tBizArea",Trim(frm1.txtBizArea.value) 
		Else
			WriteCookie "tBizArea",""
		End If
		
		If Len(Trim(frm1.txtChargeType.value)) Then
			WriteCookie "ChargeType",Trim(frm1.txtChargeType.value) 
		Else
			WriteCookie "ChargeType",""
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
				
		If Len(Trim(frm1.txtChargeFrDt.text)) Then
			WriteCookie "ChargeFrDt",Trim(frm1.txtChargeFrDt.text) 
		Else
			WriteCookie "ChargeFrDt",""
		End If
		
		If Len(Trim(frm1.txtChargeToDt.text)) Then
			WriteCookie "ChargeToDt",Trim(frm1.txtChargeToDt.text) 
		Else
			WriteCookie "ChargeToDt",""
		End If
				
		If Len(Trim(frm1.txtPoNo.value)) Then
			WriteCookie "PoNo",Trim(frm1.txtPoNo.value) 
		Else
			WriteCookie "PoNo",""
		End If
		
		If Len(Trim(frm1.txtProcessStep.value)) Then
			WriteCookie "ProcessStep",Trim(frm1.txtProcessStep.value) 
		Else
			WriteCookie "ProcessStep",""
		End If
		
		Call PgmJump(BIZ_PGM_JUMP_ID)

	ElseIf Kubun = 0 Then						

		strTemp = ReadCookie(CookieSplit)

		If strTemp = "" then Exit Function

		arrVal = Split(strTemp, parent.gRowSep)

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
'==============================================================================================================================
Sub Form_Load()
	Call LoadInfTB19029							
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
    Call ggoOper.LockField(Document, "N")       

	Call InitVariables							
	Call SetDefaultVal	
	Call InitSpreadSheet()
    Call SetToolbar("1100000000001111")	
    frm1.txtBizArea.focus
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
Sub txtChargeFrDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtChargeFrDt.Action = 7
		Call SetFocusToDocument("M")	
        frm1.txtChargeFrDt.Focus
        Set gActiveElement = document.activeElement
	End If
End Sub
'==============================================================================================================================
Sub txtChargeToDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtChargeToDt.Action = 7
		Call SetFocusToDocument("M")	
        frm1.txtChargeToDt.Focus
        Set gActiveElement = document.activeElement
	End If
End Sub
'==============================================================================================================================
Sub txtChargeFrDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub
'==============================================================================================================================
Sub txtChargeToDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub
'==============================================================================================================================
Sub vspdData_GotFocus()
    ggoSpread.Source = frm1.vspdData
End Sub
'==============================================================================================================================
Function vspdData_DblClick(ByVal Col, ByVal Row)
	If Row = 0 Or Frm1.vspdData.MaxRows = 0 Then 
	      Exit Function
	End If

	If frm1.vspdData.MaxRows > 0 Then
		If frm1.vspdData.ActiveRow = Row Or frm1.vspdData.ActiveRow > 0 Then
		End If
	End If
End Function
'==============================================================================================================================	
Sub vspdData_Click(ByVal Col, ByVal Row)
	Dim ii
	Dim iRowSep
	
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
	iRowSep = Parent.gRowSep 
	
	For ii = 1 to C_MaxKey
	    IsCookieSplit = IsCookieSplit & Trim(GetSpreadText(frm1.vspdData,GetKeyPos("A",ii),Row,"X","X")) & iRowSep
	Next
	
End Sub
'==============================================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
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

    With frm1
         If CompareDateByFormat(.txtChargeFrDt.text,.txtChargeToDt.text,.txtChargeFrDt.Alt,.txtChargeToDt.Alt, _
                   "970025",.txtChargeFrDt.UserDefinedFormat,parent.gComDateType,False) = False And Trim(.txtChargeFrDt.text) <> "" And Trim(.txtChargeToDt.text) <> "" then
			Call DisplayMsgBox("17a003","X","발생일자","X")	
			Exit Function
		End if   
	End with
	
	Call ggoOper.ClearField(Document, "2")	
	
	ggoSpread.Source = frm1.vspdData
	ggoSpread.ClearSpreadData				
		
    Call InitVariables 	
    
    Call DbQuery								

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

    DbQuery = False
    
    Err.Clear                                   

    If  LayerShowHide(1) = False Then
      	Exit Function
    End If

    With frm1
	  	If lgIntFlgMode = parent.OPMD_UMODE Then
	  	    strVal = BIZ_PGM_ID	& "?txtBizArea=" & Trim(.hdnBizArea.value)
		    strVal = strVal	& "&txtChargeType=" &	Trim(.hdnChargeType.value)
		    strVal = strVal	& "&txtBpCd="        &	Trim(.hdnBpCd.value)
		    strVal = strVal	& "&txtItemCd="      &	Trim(.hdnItemCd.value)
		    strVal = strVal	& "&txtChargeFrDt=" & Trim(.hdnChargeFrDt.value)
		    strVal = strVal	& "&txtChargeToDt=" & Trim(.hdnChargeToDt.value)
		    strVal = strVal	& "&txtPoNo="	     & Trim(.hdnPoNo.value)		
		    strVal = strVal	& "&txtProcessStep=" & Trim(.hdnProcessStep.value)
        Else
	  	    strVal = BIZ_PGM_ID	& "?txtBizArea=" & Trim(.txtBizArea.value)
		    strVal = strVal	& "&txtChargeType=" &	Trim(.txtChargeType.value)
		    strVal = strVal	& "&txtBpCd="        &	Trim(.txtBpCd.value)
		    strVal = strVal	& "&txtItemCd="      &	Trim(.txtItemCd.value)
		    strVal = strVal	& "&txtChargeFrDt=" & Trim(.txtChargeFrDt.Text)
		    strVal = strVal	& "&txtChargeToDt=" & Trim(.txtChargeToDt.Text)
		    strVal = strVal	& "&txtPoNo="	     & Trim(.txtPoNo.value)		
		    strVal = strVal	& "&txtProcessStep=" & Trim(.txtProcessStep.value)
        End If
            
            strVal = strVal & "&lgPageNo="		 & lgPageNo                      
		    strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")
            strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
		    strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))

			Call RunMyBizASP(MyBizASP, strVal)									
        
    End With
    
    DbQuery = True
    Call SetToolBar("1100000000011111")										

End Function
'==============================================================================================================================
Function DbQueryOk()														
	lgBlnFlgChgValue = False
    lgSaveRow        = 1
    lgIntFlgMode = parent.OPMD_UMODE   
    
    If frm1.vspdData.MaxRows > 0 Then
		frm1.vspdData.Focus
	Else
		frm1.txtBizArea.focus
	End If
	Set gActiveElement = document.activeElement	
	
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>경비배부집계</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=*>&nbsp;</TD>
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
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="사업장" NAME="txtBizArea" SIZE=10  MAXLENGTH=10 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBizArea" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenBizArea() ">
														   <INPUT TYPE=TEXT NAME="txtBizAreaNm" SIZE=20 tag="14"></TD>
									<TD CLASS="TD5" NOWRAP>경비항목</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT Alt="경비항목" NAME="txtChargeType" SIZE=10 MAXLENGTH=20 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnChargeType" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenChargeType()">
														   <INPUT TYPE=TEXT NAME="txtChargeTypeNm" SIZE=20 tag="14"></TD>					   
								</TR>
								<TR>						   
									<TD CLASS="TD5" NOWRAP>지급처</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="지급처" NAME="txtBpCd" SIZE=10 MAXLENGTH=10 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBpCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenBpCd()">
														   <INPUT TYPE=TEXT NAME="txtBpNm" SIZE=20 tag="14"></TD>			
									<TD CLASS="TD5" NOWRAP>품목</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT Alt="품목" NAME="txtItemCd" SIZE=10 MAXLENGTH=18 MAXLENGTH=18 ONChange="vbscript:setplant()" tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenItemCd()">
														   <INPUT TYPE=TEXT Alt="품목" NAME="txtItemNm" SIZE=20 tag="14"></TD>
	                            </TR>														   
									<TD CLASS="TD5" NOWRAP>진행구분</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT Alt="진행구분" NAME="txtProcessStep" SIZE=10 MAXLENGTH=5  tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnProcessStep" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenProcessStep()">
														   <INPUT TYPE=TEXT NAME="txtProcessStepNm" SIZE=20 tag="14"></TD>	
									<TD CLASS="TD5" NOWRAP>발생일자</TD>
									<TD CLASS="TD6" NOWRAP>
										<table cellspacing=0 cellpadding=0>
											<tr>
												<td>
													<script language =javascript src='./js/m6112qa1_fpDateTime2_txtChargeFrDt.js'></script>
												</td>
												<td>~</td>
												<td>
													<script language =javascript src='./js/m6112qa1_fpDateTime2_txtChargeToDt.js'></script>
												</td>
											<tr>
										</table>
									</TD>
	                            </TR>
								<TR>
								    <TD CLASS="TD5" NOWRAP>발주번호</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT Alt="발주번호" NAME="txtPoNo" SIZE=32 MAXLENGTH=18 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPoNo" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPoNo()">
                                    <TD CLASS="TD5" NOWRAP>&nbsp;</TD>
                                    <TD CLASS="TD6" NOWRAP>&nbsp;</td>
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
									<script language =javascript src='./js/m6112qa1_vaSpread1_vspdData.js'></script>
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
					<TD WIDTH="*" ALIGN="RIGHT"><a ONCLICK="VBSCRIPT:CookiePage(1)">경비배부상세조회</a></TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
    </TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=No noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="hdnYear" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnplantcd" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnBizArea" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnChargeType" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnBpCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnItemCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnChargeFrDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnChargeToDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnPoNo" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnProcessStep" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnDocumentNo" tag="24">
</FORM>
    <DIV ID="MousePT" NAME="MousePT">
    <iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
    </DIV>
</BODY>
</HTML>

