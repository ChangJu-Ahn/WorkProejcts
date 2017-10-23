<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Procurement
'*  2. Function Name        : 
'*  3. Program ID           : m3521qa1
'*  4. Program Name         : 미발주구매요청현황조회 
'*  5. Program Desc         : 미발주구매요청현황조회 
'*  6. Component List       : 
'*  7. Modified date(First) : 2003/06/27
'*  8. Modified date(Last)  : 2003/06/27
'*  9. Modifier (First)     : KANG SU HWAN
'* 10. Modifier (Last)      : KANG SU HWAN
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change" 
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************
-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- '#########################################################################################################
'												1. 선 언 부 
'##########################################################################################################!-->
<!--'******************************************  1.1 Inc 선언   **********************************************
'	기능: Inc. Include
'*********************************************************************************************************-->
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->
<!--'==========================================  1.1.1 Style Sheet  ========================================
'=======================================================================================================-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'==========================================  1.1.2 공통 Include   ======================================
'=======================================================================================================-->
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
Const BIZ_PGM_ID 		= "m3521qb1.asp"  
Const BIZ_PGM_JUMP_ID 	= "m3111ma1"      
Const C_MaxKey          = 13		

'==========================================  1.2.2 Global 변수 선언  =====================================
'	1. 변수 표준에 따름. prefix로 g를 사용함.
'	2.Array인 경우는 ()를 반드시 사용하여 일반 변수와 구별해 됨 
'========================================================================================================= %>
<!-- #Include file="../../inc/lgvariables.inc" -->	

'----------------  공통 Global 변수값 정의  ----------------------------------------------------------- %>
Dim lgIsOpenPop          
Dim lgSaveRow     

'==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'=========================================================================================================
Sub InitVariables()
    lgStrPrevKey     = ""
	lgBlnFlgChgValue = False
    lgSortKey        = 1
    lgSaveRow        = 0
    lgIntFlgMode = Parent.OPMD_CMODE  
    lgPageNo         = ""
End Sub

'==========================================  2.2.1 SetDefaultVal()  ======================================
'	Name : SetDefaultVal()
'	Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
'=========================================================================================================
Sub SetDefaultVal()
	Dim StartDate, EndDate, EndDate1
	
    StartDate   = uniDateAdd("m", -1, "<%=GetSvrDate%>", parent.gServerDateFormat)
    StartDate   = UniConvDateAToB(StartDate, parent.gServerDateFormat, parent.gDateFormat)
	EndDate     = UniConvDateAToB("<%=GetSvrDate%>", parent.gServerDateFormat, parent.gDateFormat)
	EndDate1	= uniDateAdd("m", +1, "<%=GetSvrDate%>", parent.gServerDateFormat)
    EndDate1   = UniConvDateAToB(EndDate1, parent.gServerDateFormat, parent.gDateFormat)
    
	With frm1
		.txtPrFrDt.Text	= StartDate
		.txtPrToDt.Text	= EndDate
		.txtPdFrDt.Text	= StartDate
		.txtPdToDt.Text	= EndDate1
		.txtPlantCd.value= parent.gPlant
		.txtPlantCd.focus
	End With
	Set gActiveElement = document.activeElement

End Sub

'======================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "M", "NOCOOKIE", "QA") %>
End Sub

'======================= 2.2.3 InitSpreadSheet() ========================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================
Sub InitSpreadSheet()
    
    Call SetZAdoSpreadSheet("M3521QA1","S","A","V20030627", Parent.C_SORT_DBAGENT, frm1.vspdData, C_MaxKey, "X", "X" )
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
	Else
	
	End If
	
End Sub

'------------------------------------------  OpenPlantCd()  --------------------------------------------
'	Name : OpenPlantCd()
'	Description : Plant PopUp
'-------------------------------------------------------------------------------------------------------
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
		Exit Function
	Else
		frm1.txtPlantCd.Value = arrRet(0)
		frm1.txtPlantNm.Value = arrRet(1)
		frm1.txtPlantCd.focus
	End If	
	frm1.txtItemCd.value=""
	frm1.txtItemNm.value=""
End Function

'------------------------------------------  OpenItemCd()  -----------------------------------------------
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
		Call DisplayMsgBox("17A002", "X", "공장", "X")
		frm1.txtPlantCd.focus
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
		arrParam(4) = arrParam(4) & "And B_Plant.Plant_Cd= " & FilterVar(UCase(frm1.txtPlantCd.Value), "''", "S") & " "    
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
    
	'arrRet = window.showModalDialog("../m1/m1111pa1.asp", Array(window.parent, arrParam, arrField, arrHeader), _
	'	"dialogWidth=695px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

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
		Exit Function
	Else
		frm1.txtItemCd.Value = arrRet(0)
		frm1.txtItemNm.Value = arrRet(1)
		frm1.txtItemCd.focus
	End If	
End Function

'------------------------------------------  OpenPrStsCd()  -------------------------------------------------
'	Name : OpenPrStsCd()
'	Description : PrStatus PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenPrStsCd()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	arrParam(0) = "요청진행상태"				
	arrParam(1) = "B_MINOR"						
	arrParam(2) = Trim(frm1.txtPrStsCd.Value)	
'	arrParam(3) = Trim(frm1.txtPrStsNm.Value)	
	arrParam(4) = "MAJOR_CD = " & FilterVar("M2101", "''", "S") & ""			
	arrParam(5) = "요청진행상태"				
	
    arrField(0) = "MINOR_CD"					
    arrField(1) = "MINOR_NM"					
    
    arrHeader(0) = "요청진행상태"			
    arrHeader(1) = "요청진행상태명"			
    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtPrStsCd.focus
		Exit Function
	Else
		frm1.txtPrStsCd.Value = arrRet(0)
		frm1.txtPrStsNm.Value = arrRet(1)
		frm1.txtPrStsCd.focus
	End If	
End Function
'------------------------------------------  OpenRqDeptCd()  -------------------------------------------
'	Name : OpenRqDeptCd()
'	Description : Req Dept PopUp
'-------------------------------------------------------------------------------------------------------
Function OpenRqDeptCd()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	arrParam(0) = "요청부서"					
	arrParam(1) = "B_ACCT_DEPT"					
	arrParam(2) = Trim(frm1.txtRqDeptCd.Value)	
'	arrParam(3) = Trim(frm1.txtRqDeptNm.Value)	
	arrParam(4) = ""							
	arrParam(5) = "요청부서"					
	
    arrField(0) = "DEPT_CD"						
    arrField(1) = "DEPT_NM"						
    
    arrHeader(0) = "요청부서"				
    arrHeader(1) = "요청부서명"				
    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtRqDeptCd.focus
		Exit Function
	Else
		frm1.txtRqDeptCd.Value = arrRet(0)
		frm1.txtRqDeptNm.Value = arrRet(1)
		frm1.txtRqDeptCd.focus
	End If	
End Function


'------------------------------------------  OpenPrTypeCd()  ---------------------------------------------
'	Name : OpenPrTypeCd()
'	Description : PR Type PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenPrTypeCd()

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	arrParam(0) = "구매요청구분"				
	arrParam(1) = "B_MINOR"						
	arrParam(2) = Trim(frm1.txtPrTypeCd.Value)	
'	arrParam(3) = Trim(frm1.txtPrTypeNm.Value)	
	arrParam(4) = "MAJOR_CD = " & FilterVar("M2102", "''", "S") & " "			
	arrParam(5) = "구매요청구분"				
	
    arrField(0) = "MINOR_CD"					
    arrField(1) = "MINOR_NM"					
        
    arrHeader(0) = "구매요청구분"			
    arrHeader(1) = "구매요청구분명"			
    
    arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtPrTypeCd.focus
		Exit Function
	Else
		frm1.txtPrTypeCd.Value = arrRet(0)
		frm1.txtPrTypeNm.Value = arrRet(1)
		frm1.txtPrTypeCd.focus
	End If	
End Function

'------------------------------------------  OpenTrackNo()  ----------------------------------------------
'	Name : OpenTrackNo()
'	Description : TrackNo PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenTrackNo()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function
	
	lgIsOpenPop = True

	arrParam(0) = "Tracking No"				
	arrParam(1) = "S_SO_TRACKING"				
	arrParam(2) = Trim(frm1.txtTrackNo.Value)	
	arrParam(5) = "Tracking No"				

    arrField(0) = "TRACKING_NO"					
    arrField(1) = "SO_NO"						
    arrField(2) = "Convert(VARCHAR(3),SO_SEQ)"	
    
    arrHeader(0) = "Tracking No"			
    arrHeader(1) = "SO NO"					
    arrHeader(2) = "SO SEQ"				
    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=600px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtTrackNo.focus
		Exit Function
	Else
		frm1.txtTrackNo.Value = arrRet(0)
		frm1.txtTrackNo.focus
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
	
	Call OpenGroupPopup("A")
End Sub
'========================================================================================================
' Name : OpenGroupPopup
' Desc :
'========================================================================================================
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

'==========================================   CookiePage()  ======================================
'	Name : CookiePage()
'	Description : JUMP시 Load화면으로 조건부로 Value
'=================================================================================================
Function CookiePage(ByVal Kubun)
	If Kubun = 1 Then						
		Call PgmJump(BIZ_PGM_JUMP_ID)
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
'   Event Name : txtPrFrDt
'   Event Desc : Date OCX Double Click
'==========================================================================================
 Sub txtPrFrDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtPrFrDt.Action = 7
        Call SetFocusToDocument("M")  
        frm1.txtPrFrDt.Focus
	End If
End Sub
'==========================================================================================
'   Event Name : txtPrToDt
'   Event Desc : Date OCX Double Click
'==========================================================================================
 Sub txtPrToDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtPrToDt.Action = 7
        Call SetFocusToDocument("M")  
        frm1.txtPrToDt.Focus
	End If
End Sub

'==========================================================================================
'   Event Name : txtPdFrDt
'   Event Desc : Date OCX Double Click
'==========================================================================================
 Sub txtPdFrDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtPdFrDt.Action = 7
        Call SetFocusToDocument("M")  
        frm1.txtPdFrDt.Focus
	End If
End Sub
'==========================================================================================
'   Event Name : txtPdToDt
'   Event Desc : Date OCX Double Click
'==========================================================================================
Sub txtPdToDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtPdToDt.Action = 7
        Call SetFocusToDocument("M")  
        frm1.txtPdToDt.Focus
	End If
End Sub
'==========================================================================================
'   Event Name : OCX_KeyDown()
'   Event Desc : 
'==========================================================================================
Sub txtPrFrDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub

Sub txtPrToDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub

Sub txtPdFrDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub

Sub txtPdToDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub
'==========================================================================================
'   Event Name : vspdData_GotFocus
'   Event Desc : This event is spread sheet data changed
'==========================================================================================
 Sub vspdData_GotFocus()
    ggoSpread.Source = frm1.vspdData
End Sub

'======================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'======================================================================================================
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
End Sub
	
'======================================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'======================================================================================================
 Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspddata,NewTop) Then	    
    	If lgPageNo <> "" Then								
 			If CheckRunningBizProcess = True Then
				Exit Sub
			End If			
			Call DisableToolBar(parent.TBC_QUERY)
			If DBQuery = False Then
				Call RestoreToolBar()
				Exit Sub
			End If
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
	    If CompareDateByFormat(.txtPrFrDt.text,.txtPrToDt.text,.txtPrFrDt.Alt,.txtPrToDt.Alt, _
                   "970025",.txtPrFrDt.UserDefinedFormat,parent.gComDateType,False) = False And Trim(.txtPrFrDt.text) <> "" And Trim(.txtPrToDt.text) <> "" Then
			Call DisplayMsgBox("17a003", "X","요청일", "X")			
			Exit Function
		End if   
	End with

	with frm1
	    If CompareDateByFormat(.txtPdFrDt.text,.txtPdToDt.text,.txtPdFrDt.Alt,.txtPdToDt.Alt, _
                   "970025",.txtPdFrDt.UserDefinedFormat,parent.gComDateType,False) = False And Trim(.txtPdFrDt.text) <> "" And Trim(.txtPdToDt.text) <> "" Then
			Call DisplayMsgBox("17a003", "X","필요일", "X")
			Exit Function
		End if   
	End with

    '-----------------------
    'Erase contents area
    '-----------------------
'    Call ggoOper.ClearField(Document, "2")						
	ggoSpread.Source = frm1.vspdData	'###그리드 컨버전 주의부분###
    ggoSpread.ClearSpreadData
    Call InitVariables 											
    
    '-----------------------
    'Check condition area
    '-----------------------
'    If Not chkField(Document, "1") Then							
'       Exit Function
'    End If
	
    '-----------------------
    'Query function call area
    '-----------------------

    If DbQuery = False Then Exit Function

    FncQuery = True												

End Function

'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================
Function FncPrint() 
    Call parent.FncPrint()
End Function

'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================
Function FncExcel() 
	Call parent.FncExport(parent.C_MULTI)
End Function

'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================
Function FncFind() 
    Call parent.FncFind(parent.C_MULTI , False)                        
End Function

'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================
Function FncExit()
    FncExit = True
End Function

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbQuery() 
	Dim strVal

    DbQuery = False
    
    Err.Clear                                                   
	If LayerShowHide(1) = False Then Exit Function
    
    With frm1
	If lgIntFlgMode = parent.OPMD_UMODE Then
	    strVal = BIZ_PGM_ID & "?txtPlantCd=" & Trim(.txtHPlantCd.value)
	    strVal = strVal & "&txtItemCd=" & Trim(.txtHItemCd.Value)
	    strVal = strVal & "&txtRqDeptCd=" & Trim(.txtHRqDeptCd.Value)
    	strVal = strVal & "&txtPrFrDt=" & Trim(.txtHPrFrDt.value)
    	strVal = strVal & "&txtPrToDt=" & Trim(.txtHPrToDt.value)    	
    	strVal = strVal & "&txtPdFrDt=" & Trim(.txtHPdFrDt.value)
    	strVal = strVal & "&txtPdToDt=" & Trim(.txtHPdToDt.value)
	else
	    strVal = BIZ_PGM_ID & "?txtPlantCd=" & Trim(.txtPlantCd.value)
	    strVal = strVal & "&txtItemCd=" & Trim(.txtItemCd.Value)
	    strVal = strVal & "&txtRqDeptCd=" & Trim(.txtRqDeptCd.Value)
    	strVal = strVal & "&txtPrFrDt=" & Trim(.txtPrFrDt.Text)
    	strVal = strVal & "&txtPrToDt=" & Trim(.txtPrToDt.Text)    	
    	strVal = strVal & "&txtPdFrDt=" & Trim(.txtPdFrDt.Text)
    	strVal = strVal & "&txtPdToDt=" & Trim(.txtPdToDt.Text)
    end if

	    strVal = strVal & "&txtchangorgid="   & parent.gChangeOrgId         
	    strVal = strVal & "&lgPageNo="   & lgPageNo         
        strVal = strVal & "&lgStrPrevKey="   & lgStrPrevKey          
		strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")
		strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
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

    '-----------------------
    'Reset variables area
    '-----------------------
	lgBlnFlgChgValue = False
    lgSaveRow        = 1
    lgIntFlgMode = parent.OPMD_UMODE
    If frm1.vspdData.MaxRows > 0 Then
		frm1.vspddata.focus
	Else
		frm1.txtPlantCd.focus
	End If
	Set gActiveElement = document.activeElement

End Function

</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
</HEAD>
<!--'#########################################################################################################
'       					6. Tag부 
'#########################################################################################################-->
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>미발주 구매요청현황조회</font></td>
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
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="공장"  NAME="txtPlantCd" SIZE=10 LANG="ko" MAXLENGTH=4 tag="1XNXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenPlantCd() ">
														   <INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=20 tag="14"></TD>

									<TD CLASS="TD5" NOWRAP>품목</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="품목" NAME="txtItemCd" SIZE=10 LANG="ko" MAXLENGTH=18 tag="1XNXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenItemCd() ">
														   <INPUT TYPE=TEXT NAME="txtItemNm" SIZE=20 tag="14"></TD>														   
								</TR>						   
									<TD CLASS="TD5" NOWRAP>요청부서</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT Alt="요청부서" NAME="txtRqDeptCd" SIZE=10 MAXLENGTH=10  MAXLENGTH=10 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnRqDeptCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenRqDeptCd()">
														   <INPUT TYPE=TEXT NAME="txtRqDeptNm" SIZE=20 tag="14"></TD>
									<TD CLASS="TD5" NOWRAP>필요일</TD>
									<TD CLASS="TD6" NOWRAP>
										<table cellpadding=0 cellspacing=0>
											<tr>
												<td NOWRAP>
													<script language =javascript src='./js/m3521qa1_fpDateTime2_txtPdFrDt.js'></script>
												</td>
												<td NOWRAP>~</td>
												<td NOWRAP>
												   <script language =javascript src='./js/m3521qa1_fpDateTime2_txtPdToDt.js'></script>
												</td>
											</tr>
										</table>
									</TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>요청일</TD>
									<TD CLASS="TD6" NOWRAP>
										<table cellpadding=0 cellspacing=0>
											<tr>
												<td NOWRAP>
													<script language =javascript src='./js/m3521qa1_fpDateTime2_txtPrFrDt.js'></script>
												</td>
												<td NOWRAP>~</td>
												<td NOWRAP>
												   <script language =javascript src='./js/m3521qa1_fpDateTime2_txtPrToDt.js'></script>
												</td>
											</tr>
										</table>
									</TD>
									<TD CLASS="TD5" NOWRAP>&nbsp;</TD>
									<TD CLASS="TD6" NOWRAP>&nbsp;</TD>
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
									<script language =javascript src='./js/m3521qa1_A_vspdData.js'></script>
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
					<TD WIDTH="*" ALIGN="RIGHT"><a ONCLICK="VBSCRIPT:CookiePage(1)">발주등록</a></TD>
					<TD WIDTH=10></TD>
				</TR>
			</TABLE>
		</TD>
    </TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0  tabindex="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtHPlantCd" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHItemCd" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHRqDeptCd" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHPdFrDt" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHPdToDt" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHPrFrDt" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHPrToDt" tag="24">
</FORM>
    <DIV ID="MousePT" NAME="MousePT">
    <iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
    </DIV>
</BODY>
</HTML>
