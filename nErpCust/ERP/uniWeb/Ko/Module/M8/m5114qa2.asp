<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% session.CodePage=949 %>

<!--
'**********************************************************************************************
'*  1. Module Name          : Procurement
'*  2. Function Name        : 
'*  3. Program ID           :
'*  4. Program Name         :
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 
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
<!--'#########################################################################################################
'												1. 선 언 부
'##########################################################################################################-->
<!--'******************************************  1.1 Inc 선언   **********************************************
'	기능: Inc. Include
'*********************************************************************************************************-->
<!-- #Include file="../../inc/incSvrCcm.inc" -->
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
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentA.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="vbscript"   SRC="../../inc/Cookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript">

Option Explicit

'******************************************  1.2 Global 변수/상수 선언  ***********************************
'	1. Constant는 반드시 대문자 표기.
'**********************************************************************************************************
<!-- #Include file="../../inc/lgvariables.inc" -->	
Dim lgIsOpenPop      
Dim lgSaveRow  
Dim IscookieSplit
 
Const BIZ_PGM_ID 		= "m5114qb2.asp"               
Const C_MaxKey          = 29					       

Dim StartDate, EndDate
	
StartDate = uniDateAdd("m", -1, "<%=GetSvrDate%>", Parent.gServerDateFormat)
StartDate = UniConvDateAToB(StartDate, Parent.gServerDateFormat, Parent.gDateFormat)
EndDate   = UniConvDateAToB("<%=GetSvrDate%>", Parent.gServerDateFormat, Parent.gDateFormat)
	
'==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'=========================================================================================================
 Sub InitVariables()
    lgPageNo         = ""
	lgBlnFlgChgValue = False
    lgSortKey        = 1
    lgSaveRow        = 0
    lgIntFlgMode = Parent.OPMD_CMODE 
End Sub
'==========================================  2.2.1 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
'=========================================================================================================
Sub SetDefaultVal()
	frm1.txtMvFrDt.Text	= StartDate
	frm1.txtMvToDt.Text	= EndDate
	Set gActiveElement = document.activeElement
End Sub
'======================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'========================================================================================
Sub LoadInfTB19029()
	<!--#Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "M", "NOCOOKIE", "QA") %>
	<% Call LoadBNumericFormatA("Q", "M", "NOCOOKIE", "QA") %>
End Sub

'======================= 2.2.3 InitSpreadSheet() ========================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================
Sub InitSpreadSheet()
 	Call SetZAdoSpreadSheet("M4111QA6","S","A","V20030513", Parent.C_SORT_DBAGENT, frm1.vspdData, C_MaxKey, "X", "X" )
	Call SetSpreadLock("A")       
End Sub

'================================== 2.2.4 SetSpreadLock() ==================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadLock(ByVal pOpt)
    If pOpt = "A" Then
		ggoSpread.Source = frm1.vspdData
		ggoSpread.SpreadLockWithOddEvenRowColor()
	End If
End Sub
'------------------------------------------  OpenPlantCd()  --------------------------------------------
'	Name : OpenPlantCd()
'	Description : Plant PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenPlantCd()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function
	
	lgIsOpenPop = True

	arrParam(0) = "공장"						
	arrParam(1) = "B_Plant"						
	arrParam(2) = Trim(frm1.txtPlantCd.Value)	
'	arrParam(3) = Trim(frm1.txtPlantNm.Value)	
	arrParam(4) = ""							
	arrParam(5) = "공장"						

    arrField(0) = "Plant_Cd"					
    arrField(1) = "Plant_NM"					
    
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
End Function

'------------------------------------------  OpenItemCd()  -------------------------------------------------
'	Name : OpenItemCd()
'	Description : Item PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenItemCd()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim iCalledAspName
	Dim IntRetCD
	
	If lgIsOpenPop = True Then Exit Function
	
	if  Trim(frm1.txtPlantCd.Value) = "" then
		Call DisplayMsgBox("17A002","X" , "공장","X")
		Exit Function
	End if
	
	lgIsOpenPop = True

	arrParam(0) = "품목"						
	arrParam(1) = "B_Item_By_Plant,B_Plant,B_Item"
	arrParam(2) = Trim(frm1.txtItemCd.Value)	
'	arrParam(3) = Trim(frm1.txtItemNm.Value)	
	arrParam(4) = "B_Item_By_Plant.Plant_Cd = B_Plant.Plant_Cd And "
	arrParam(4) = arrParam(4) & "B_Item_By_Plant.Item_Cd = B_Item.Item_Cd and B_Item.phantom_flg = " & FilterVar("N", "''", "S") & "  "
	if Trim(frm1.txtPlantCd.Value)<>"" then
		arrParam(4) = arrParam(4) & "And B_Plant.Plant_Cd= " & FilterVar(frm1.txtPlantCd.value, "''", "S") & " "
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
    
	iCalledAspName = AskPRAspName("M1111PA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "M1111PA1", "X")
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

'------------------------------------------  OpenBpCd()  -------------------------------------------------
'	Name : OpenBpCd()
'	Description : Supplier PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenBpCd()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	arrParam(0) = "거래처"					
	arrParam(1) = "B_Biz_Partner"				
	arrParam(2) = Trim(frm1.txtBpCd.Value)		
'	arrParam(3) = Trim(frm1.txtBpNm.Value)		
	arrParam(4) = "BP_TYPE in (" & FilterVar("S", "''", "S") & " ," & FilterVar("CS", "''", "S") & ") "					
	arrParam(5) = "거래처"					
	
    arrField(0) = "BP_CD"						
    arrField(1) = "BP_NM"						
    
    arrHeader(0) = "거래처"					
    arrHeader(1) = "거래처명"				
    
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
'------------------------------------------  OpenSlCd()  -------------------------------------------------
'	Name : OpenSlCd()
'	Description : Sl PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenSlCd()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	arrParam(0) = "창고"						
	arrParam(1) = "B_STORAGE_LOCATION"			
	arrParam(2) = Trim(frm1.txtSlCd.Value)		
'	arrParam(3) = Trim(frm1.txtSlNm.Value)		
	arrParam(4) = ""							
	arrParam(5) = "창고"						
	
    arrField(0) = "SL_CD"						
    arrField(1) = "SL_NM"						
    
    arrHeader(0) = "창고"					
    arrHeader(1) = "창고명"					
    
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
'------------------------------------------  OpenIoType()  -------------------------------------------------
'	Name : OpenIoType()
'	Description : PurGrp PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenIoType()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	arrParam(0) = "입고유형"	
	arrParam(1) = "M_MVMT_TYPE"				
	
	arrParam(2) = Trim(frm1.txtIoType.Value)
'	arrParam(3) = Trim(frm1.txtIoTypeNm.Value)	
			
	arrParam(4) = "RCPT_FLG <> " & FilterVar("N", "''", "S") & " "			
	arrParam(5) = "입고유형"			
	
    arrField(0) = "IO_TYPE_CD"	
    arrField(1) = "IO_TYPE_NM"	
    
    arrHeader(0) = "입고유형"		
    arrHeader(1) = "입고유형명"
    
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

'------------------------------------  PopZAdoConfigGrid()  ----------------------------------------------
'	Name : PopZAdoConfigGrid()
'	Description : Group Condition PopUp
'---------------------------------------------------------------------------------------------------------
Sub PopZAdoConfigGrid()
	If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
		Exit Sub
	End If
	
	Call OpenSortPopup("A")
End Sub
'========================================================================================================
' Function Name : OpenSortPopup
' Function Desc : OpenSortPopup Reference Popup
'========================================================================================================
Function OpenSortPopup(ByVal pSpdNo)
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
'==========================================   CookiePage()  ======================================
'	Name : CookiePage()
'	Description : JUMP시 Load화면으로 조건부로 Value
'====================================================================================================
Function CookiePage(ByVal Kubun)

	Dim strTemp, arrVal
	Dim strCookie, i
	Dim strToDt
	Dim strAddMonthToDt

	Const CookieSplit = 4877					

	If Kubun = 1 Then							

		Call vspdData_Click(frm1.vspdData.ActiveCol,frm1.vspdData.ActiveRow)

		WriteCookie CookieSplit , IscookieSplit	
		if frm1.vspdData.ActiveRow > 0 then		
			strTemp = ReadCookie(CookieSplit)
			If strTemp = "" then Exit Function
			arrVal = Split(strTemp, parent.gRowSep)
			frm1.vspdData.Row = frm1.vspdData.ActiveRow 			
			WriteCookie "MvmtNo" , arrVal(0)					
			WriteCookie CookieSplit , ""			
		end if		
		Call PgmJump(BIZ_PGM_JUMP_ID)

	ElseIf Kubun = 0 Then						

	  If ReadCookie("From")="PO" Then			
  
		strTemp = ReadCookie(CookieSplit)
		If strTemp = "" then Exit Function

		arrVal = Split(strTemp, parent.gRowSep)		
		If Len(ReadCookie ("ItemCd")) Then
			frm1.txtItemCd.Value	=  ReadCookie ("ItemCd")
			WriteCookie "ItemCd",""
		Else
			frm1.txtItemCd.Value	=  arrVal(0)
		End If
		
		frm1.txtItemNm.Value	=  arrVal(1)
		
		If Len(ReadCookie ("PlantCd")) Then
			frm1.txtPlantCd.Value	=  ReadCookie ("PlantCd")
			WriteCookie "PlantCd",""
		Else
			frm1.txtPlantCd.Value	=  arrVal(2)
		End If
		
		frm1.txtPlantNm.value	=  arrVal(3)
		
		If Len(ReadCookie ("BpCd")) Then
			frm1.txtBpCd.Value	=  ReadCookie ("BpCd")
			WriteCookie "BpCd",""
		Else
			frm1.txtBpCd.Value	=  arrVal(4)
		End If
		
		frm1.txtBpNm.value		=  arrVal(5)
		
		WriteCookie "From",""
	  Else										
		  If Trim(ReadCookie("CookieIoIvFlg")) = "Y" Then
		 		frm1.txtMvFrDt.Text	= UNIConvDateAtoB(Trim(ReadCookie("CookieFromDt")), parent.gServerDateFormat, parent.gDateFormat)
			 	strToDt	= UNIConvDateAtoB(Trim(ReadCookie("CookieToDt")), parent.gServerDateFormat, parent.gDateFormat)
			 	strAddMonthToDt = UnIDateAdd("m", 1, strToDt, parent.gDateFormat)
			 	frm1.txtMvToDt.Text	= UnIDateAdd("d", -1, strAddMonthToDt, parent.gDateFormat)
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

'		If arrVal(0) = "" Then Exit Function
						
		If Len(ReadCookie ("ItemCd")) Then
			frm1.txtItemCd.Value	=  ReadCookie ("ItemCd")
			WriteCookie "ItemCd",""
		Else
			frm1.txtItemCd.Value	=  arrVal(0)
		End If
		
		frm1.txtItemNm.Value	=  arrVal(1)
		
		If Len(ReadCookie ("PlantCd")) Then
			frm1.txtPlantCd.Value	=  ReadCookie ("PlantCd")
			WriteCookie "PlantCd",""
		Else
			frm1.txtPlantCd.Value	=  arrVal(2)
		End If
		
		frm1.txtPlantNm.value	=  arrVal(3)
		
		If Len(ReadCookie ("BpCd")) Then
			frm1.txtBpCd.Value	=  ReadCookie ("BpCd")
			WriteCookie "BpCd",""
		Else
			frm1.txtBpCd.Value	=  arrVal(4)
		End If
		
		frm1.txtBpNm.value		=  arrVal(5)
						
		If arrVal(6) = "" or arrVal(6) = Null Then
			frm1.txtMvFrDt.Text	=  ReadCookie ("MvFrDt")
			WriteCookie "MvFrDt",""
		Else		
			frm1.txtMvFrDt.Text		=  arrVal(6)
		End If
		
		If arrVal(6) = "" or arrVal(6) = Null Then
			frm1.txtMvToDt.Text	=  ReadCookie ("MvToDt")
			WriteCookie "MvToDt",""
		Else
			frm1.txtMvToDt.Text		=  arrVal(6)
		End If
				
		If Len(ReadCookie ("SlCd")) Then
			frm1.txtSlCd.Value	=  ReadCookie ("SlCd")
			WriteCookie "SlCd",""
		Else
			frm1.txtSlCd.Value	=  arrVal(7)
		End If
		
		frm1.txtSlNm.value 	=  arrVal(8)
		
		If Len(ReadCookie ("IoType")) Then
			frm1.txtIoType.Value	=  ReadCookie ("IoType")
			WriteCookie "IoType",""
		Else
			frm1.txtIoType.Value	=  arrVal(9)
		End If
		
		frm1.txtIoTypeNm.value	=  arrVal(10)
	  End if

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
'	Name : Form_Load()
'	Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분
'========================================================================================================= 
Sub Form_Load()

    Call LoadInfTB19029							
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
    Call ggoOper.LockField(Document, "N")       
    Call InitVariables							
	Call SetDefaultVal	
	Call InitSpreadSheet()
    Call SetToolbar("1100000000001111")			
	Call CookiePage(0)
    
    frm1.txtPlantCd.focus
	Set gActiveElement = document.activeElement	
End Sub
'========================================================================================
' Function Name : vspdData_MouseDown
' Function Desc : 
'========================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)

   If Button = 2 And gMouseClickStatus = "SPC" Then
      gMouseClickStatus = "SPCR"
   End If
End Sub    

'========================================================================================
' Function Name : FncSplitColumn
' Function Desc : 
'========================================================================================
Sub FncSplitColumn()

    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   

End Sub

'==========================================================================================
'   Event Name : txtPoFrDt
'   Event Desc : Date OCX Double Click
'==========================================================================================
Sub txtMvFrDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtMvFrDt.Action = 7
		Call SetFocusToDocument("M") 
		frm1.txtMvFrDt.focus
	End If
End Sub
'==========================================================================================
'   Event Name : txtPoToDt
'   Event Desc : Date OCX Double Click
'==========================================================================================
Sub txtMvToDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtMvToDt.Action = 7
		Call SetFocusToDocument("M") 
		frm1.txtMvToDt.focus
	End If
End Sub
'==========================================================================================
'   Event Name : OCX_KeyDown()
'   Event Desc : 
'==========================================================================================
Sub txtMvFrDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub

Sub txtMvToDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub
'==========================================================================================
'   Event Name : vspdData_GotFocus
'   Event Desc : This event is spread sheet data changed
'==========================================================================================
Sub vspdData_GotFocus()
    ggoSpread.Source = frm1.vspdData
End Sub
'==========================================================================================
'   Event Name : vspdData_DblClick
'   Event Desc :
'==========================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)
	If Row = 0 Or Frm1.vspdData.MaxRows = 0 Then 
         Exit Sub
    End If
End Sub
	
'======================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : 컬럼을 클릭할 경우 발생
'=======================================================================================================%>
Sub vspdData_Click(ByVal Col, ByVal Row)
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
	
	IscookieSplit=""
	frm1.vspddata.row = Row
	frm1.vspddata.col = GetKeyPos("A",1)
	IscookieSplit = IscookieSplit & frm1.vspddata.text & parent.gRowSep
	
End Sub
	
'======================================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
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
'	설명 : Fnc함수명 으로 시작하는 모든 Function
'********************************************************************************************************* 
Function FncQuery() 

    FncQuery = False                                        
    
    Err.Clear                                               
	
	with frm1
		if (UniConvDateToYYYYMMDD(.txtMvFrDt.text,Parent.gDateFormat,"") > UniConvDateToYYYYMMDD(.txtMvToDt.text,Parent.gDateFormat,"")) And Trim(.txtMvFrDt.text) <> "" And Trim(.txtMvToDt.text) <> "" then	
			Call DisplayMsgBox("17a003","X","입고일","X")			
			Exit Function
		End if   
	End with
	
    Call ggoOper.ClearField(Document, "2")					
    Call InitVariables 										

	If DbQuery = False Then Exit Function

    FncQuery = True											
	Set gActiveElement = document.activeElement	
End Function
'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================
Function FncPrint() 
    Call parent.FncPrint()
    Set gActiveElement = document.activeElement	
End Function

'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================
Function FncExcel() 
	Call parent.FncExport(parent.C_MULTI)
	Set gActiveElement = document.activeElement	
End Function

'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================
Function FncFind() 
    Call parent.FncFind(parent.C_MULTI , False)  
    Set gActiveElement = document.activeElement	                  
End Function

'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================
Function FncExit()
	FncExit = True
	Set gActiveElement = document.activeElement	 
End Function
'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbQuery() 
	Dim strVal
	
    DbQuery = False
    
    Err.Clear                                               
    If LayerShowHide(1) = False Then
         Exit Function
    End If 
    
    With frm1
    If lgIntFlgMode = Parent.OPMD_UMODE Then
	    strVal = BIZ_PGM_ID & "?txtPlantCd=" & Trim(.hdnPlantCd.value)
    	strVal = strVal & "&txtItemCd=" & Trim(.hdnItemCd.Value)
    	strVal = strVal & "&txtBpCd=" & Trim(.hdnBpCd.Value)
    	strVal = strVal & "&txtMvFrDt=" & Trim(.hdnMvFrDt.value)
    	strVal = strVal & "&txtMvToDt=" & Trim(.hdnMvToDt.value)
    	strVal = strVal & "&txtSlCd=" & Trim(.hdnSlCd.value)
    	strVal = strVal & "&txtIoType=" & Trim(.hdnIoType.value)
        
    else
	    strVal = BIZ_PGM_ID & "?txtPlantCd=" & Ucase(Trim(.txtPlantCd.value))
    	strVal = strVal & "&txtItemCd=" & Trim(.txtItemCd.Value)
    	strVal = strVal & "&txtBpCd=" & Trim(.txtBpCd.Value)
    	strVal = strVal & "&txtMvFrDt=" & Trim(.txtMvFrDt.Text)
    	strVal = strVal & "&txtMvToDt=" & Trim(.txtMvToDt.Text)
    	strVal = strVal & "&txtSlCd=" & Trim(.txtSlCd.value)
    	strVal = strVal & "&txtIoType=" & Trim(.txtIoType.value)
    end if
        strVal = strVal & "&lgPageNo="   & lgPageNo    
        strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")
        'Modified by KSJ 2008-04-11
        'strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
		strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))

        Call RunMyBizASP(MyBizASP, strVal)								
        
    End With
    
    DbQuery = True
    Call SetToolbar("1100000000011111")									

End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김
'========================================================================================
Function DbQueryOk()													

    lgBlnFlgChgValue = False
    lgSaveRow        = 1
    lgIntFlgMode = Parent.OPMD_UMODE
    frm1.vspdData.focus
    Set gActiveElement = document.activeElement
End Function
</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
</HEAD>
<!-- '#########################################################################################################
'       					6. Tag부
'######################################################################################################### -->
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>미매입입고상세</font></td>
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
								    <TD CLASS="TD5" NOWRAP>공장</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="공장" NAME="txtPlantCd" SIZE=10 LANG="ko" MAXLENGTH=4 tag="1XNXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenPlantCd()">
														   <INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=20 tag="14"></TD>
									<TD CLASS="TD5" NOWRAP>품목</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT Alt="품목" NAME="txtItemCd" SIZE=10 MAXLENGTH=18 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItem" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenItemCd()">
														   <INPUT TYPE=TEXT Alt="품목" NAME="txtItemNm" SIZE=20 tag="14"></TD>					   									
								</TR>					   
								<TR>						   
									<TD CLASS="TD5" NOWRAP>거래처</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="거래처" NAME="txtBpCd" SIZE=10 MAXLENGTH=10 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBpCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenBpCd()">
														   <INPUT TYPE=TEXT NAME="txtBpNm" SIZE=20 tag="14"></TD>			
									<TD CLASS="TD5" NOWRAP>입고일</TD>
									<TD CLASS="TD6" NOWRAP>
										<table cellpadding=0 cellspacing=0>
											<tr>
												<td NOWRAP>
													<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT id=fpDateTime2 title=FPDATETIME CLASS=FPDTYYYYMMDD name=txtMvFrDt CLASSID=<%=gCLSIDFPDT%> tag="11X1" ALT="입고일"></OBJECT>');</SCRIPT>
												</td>
												<td NOWRAP>~</td>
												<td NOWRAP>
												   <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT id=fpDateTime2 title=FPDATETIME CLASS=FPDTYYYYMMDD name=txtMvToDt CLASSID=<%=gCLSIDFPDT%> tag="11X1" ALT="입고일"></OBJECT>');</SCRIPT>
												</td>
											</tr>
										</table>
									</TD>
	                            </TR>	
	                            <TR>
									<TD CLASS="TD5" NOWRAP>창고</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT Alt="창고" NAME="txtSlCd" SIZE=10 MAXLENGTH=7 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSlCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenSlCd()">
														   <INPUT TYPE=TEXT NAME="txtSlNm" SIZE=20 tag="14"></TD>											   		   
									<TD CLASS="TD5" NOWRAP>입고유형</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="입고유형" NAME="txtIoType" SIZE=10 LANG="ko" MAXLENGTH=5 tag="1XNXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnIoType" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenIoType() ">
														   <INPUT TYPE=TEXT NAME="txtIoTypeNm" SIZE=20 tag="14"></TD>									
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
	
    <TR>
      <TD <%=HEIGHT_TYPE_01%>></TD>
    </TR>
    
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TabIndex="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="hdnPlantCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnItemCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnBpCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnMvFrDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnMvToDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnSlCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnIoType" tag="24">
</FORM>
    <DIV ID="MousePT" NAME="MousePT">
    <iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
    </DIV>
</BODY>
</
