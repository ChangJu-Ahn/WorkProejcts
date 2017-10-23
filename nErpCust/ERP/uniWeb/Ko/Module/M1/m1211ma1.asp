<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Prucurement
'*  2. Function Name        : 
'*  3. Program ID           : M1211MA1
'*  4. Program Name         : 품목별공급처등록 
'*  5. Program Desc         : 품목별공급처등록 
'*  6. Component List       : 
'*  7. Modified date(First) : 2000/05/08
'*  8. Modified date(Last)  : 2003/05/23
'*  9. Modifier (First)     : Shin Jin-hyun
'* 10. Modifier (Last)      : Kang Su Hwan
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
'            1. 선 언 부 
'##########################################################################################################!-->
<!-- '******************************************  1.1 Inc 선언   **********************************************
' 기능: Inc. Include
'********************************************************************************************************* !-->
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->
<!--'==========================================  1.1.1 Style Sheet  ======================================
'==========================================================================================================!-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'==========================================  1.1.2 공통 Include   ======================================
'==========================================================================================================!-->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>
Option Explicit               

<!--'==========================================  1.2.1 Global 상수 선언  ======================================
'==========================================================================================================!-->
Const BIZ_PGM_ID		= "m1211mb1.asp"  
Const BIZ_PGM_JUMP_ID	= "m1211qa1"

<!-- '==========================================  1.2.2 Global 변수 선언  =====================================
' 1. 변수 표준에 따름. prefix로 g를 사용함.
' 2.Array인 경우는 ()를 반드시 사용하여 일반 변수와 구별해 됨 
'========================================================================================================= !-->
Dim lgBlnFlgChgValue  
Dim lgIntFlgMode   

<!-- '++++++++++++++++  Insert Your Code for Global Variables Assign  ++++++++++++++++++++++++++++++++++++++ !-->
Dim lgIsOpenPop
Dim IsOpenPop          
Dim iDBSYSDate
Dim EndDate, StartDate

iDBSYSDate = "<%=GetSvrDate%>"
EndDate = UNIConvDateAtoB(iDBSYSDate, parent.gServerDateFormat, parent.gDateFormat)

'--------------------------------------------------------------------
'  Cookie 사용함수 
'--------------------------------------------------------------------    
Sub ReadCookiePage()
	Dim strTemp, arrVal
	 
	If Trim(ReadCookie("m1211qa1_suppliercd")) = "" then 
		Exit sub
	End if

	frm1.txtPlantCd1.value		= ReadCookie("m1211qa1_plantcd")
	frm1.txtItemCd1.value		= ReadCookie("m1211qa1_itemcd")
	frm1.txtSupplierCd1.value	= ReadCookie("m1211qa1_suppliercd")
	    
	Call WriteCookie("m1211qa1_plantcd" , "")
	Call WriteCookie("m1211qa1_itemcd" , "")
	Call WriteCookie("m1211qa1_suppliercd" , "")
	 
	Call MainQuery() 
End Sub

Function WriteCookiePage()
	Dim IntRetCD
    
	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900017", parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then 
			Exit Function
		End If
	End If

	Call WriteCookie("m1211ma1_plantcd", frm1.txtPlantCd2.Value)
	Call WriteCookie("m1211ma1_itemcd", frm1.txtItemCd2.Value)
	Call WriteCookie("m1211ma1_suppliercd",frm1.txtSupplierCd2.Value)

	Call PgmJump(BIZ_PGM_JUMP_ID) 
End Function

<!-- '==========================================  2.1.1 InitVariables()  ======================================
' Name : InitVariables()
' Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'========================================================================================================= !-->
Sub InitVariables()
    lgIntFlgMode = parent.OPMD_CMODE       
    lgBlnFlgChgValue = False        
    IsOpenPop = False    
End Sub

<!--'=============================== 2.1.2 LoadInfTB19029() ========================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== !-->
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A( "I", "*", "NOCOOKIE", "MA") %>
End Sub

<!-- '==========================================  2.2.1 SetDefaultVal()  ========================================
' Name : SetDefaultVal()
' Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
'========================================================================================================= !-->
Sub SetDefaultVal()
 	frm1.rdoUseflg(0).checked = true 
	frm1.rdoDefFlg(1).checked = true
	frm1.txtPriority.Text	  = 1
	frm1.txtPlantCd1.value	  = parent.gPlant
	frm1.txtPlantNm1.value	  = parent.gPlantNm 
	frm1.txtPlantCd2.value	  = parent.gPlant
	frm1.txtPlantNm2.value	  = parent.gPlantNm
	frm1.txtGroupCd.value	  = parent.gPurGrp
	frm1.txtGroupNm.value	  = ""
	frm1.txtPlantCd1.focus 
	Set gActiveElement = document.activeElement
	Call SetToolbar("1110100000001111")
End Sub

<!-- '------------------------------------------  OpenPlant()  -------------------------------------------------
' Name : OpenPlant()
' Description : Plant PopUp
'--------------------------------------------------------------------------------------------------------- !-->
Function OpenPlant(byval strComp)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtPlantCd1.className) = UCase(parent.UCN_PROTECTED) And strComp="Plant1" Then Exit Function
	If IsOpenPop = True Or UCase(frm1.txtPlantCd2.className) = UCase(parent.UCN_PROTECTED) And strComp="Plant2" Then Exit Function
	IsOpenPop = True

	arrParam(0) = "공장" 
	arrParam(1) = "B_Plant"    
	 
	If strComp="Plant1" Then
	 arrParam(2) = Trim(frm1.txtPlantCd1.Value)
	Else
	 arrParam(2) = Trim(frm1.txtPlantCd2.Value)
	End If 
	 
	arrParam(4) = ""   
	arrParam(5) = "공장"   
	 
	arrField(0) = "Plant_Cd" 
	arrField(1) = "Plant_NM" 
	    
	arrHeader(0) = "공장"  
	arrHeader(1) = "공장명"  
	    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
	 "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	 
	IsOpenPop = False
	 
	If arrRet(0) = "" Then
		If strComp="Plant1" Then
			frm1.txtPlantCd1.focus
		Else
			frm1.txtPlantCd2.focus
		End If
		Exit Function
	Else
		If strComp="Plant1" Then
			frm1.txtPlantCd1.Value= arrRet(0)  
			frm1.txtPlantNm1.Value= arrRet(1)  
			frm1.txtPlantCd1.focus
		Else
			frm1.txtPlantCd2.Value= arrRet(0)  
			frm1.txtPlantNm2.Value= arrRet(1)
			Call ChangeItemPlant()
			lgBlnFlgChgValue = True
			frm1.txtPlantCd2.focus
		End If 
	End If 
	Set gActiveElement = document.activeElement
End Function

<!-- '------------------------------------------  OpenCondPlant()  -------------------------------------------
' Name : OpenCondPlant()
' Description : Condition Plant PopUp
'--------------------------------------------------------------------------------------------------------- !-->

Function OpenItemCd1()
	Dim arrRet
	Dim arrParam(5), arrField(1)
	Dim iCalledAspName
	Dim IntRetCD

	If lgIsOpenPop = True Then Exit Function

	if  Trim(frm1.txtPlantCd1.Value) = "" then
		Call DisplayMsgBox("17A002", "X", "공장", "X")
		frm1.txtPlantCd1.focus
		Exit Function
	End if

	lgIsOpenPop = True

	arrParam(0) = Trim(frm1.txtPlantCd1.value)	' Plant Code
	arrParam(1) = Trim(frm1.txtItemCd1.value)		' Item Code
	arrParam(2) = "!"	' “12!MO"로 변경 -- 품목계정 구분자 조달구분 
	arrParam(3) = "30!P"
	arrParam(4) = ""		'-- 날짜 
	arrParam(5) = ""		'-- 자유(b_item_by_plant a, b_item b: and 부터 시작)

	arrField(0) = 1 ' -- 품목코드 
	arrField(1) = 2 ' -- 품목명 

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
		frm1.txtItemCd1.focus
		Exit Function
	Else
		frm1.txtItemCd1.Value	= arrRet(0)
		frm1.txtItemNm1.Value	= arrRet(1)
		frm1.txtItemCd1.focus
	End If
End Function

'------------------------------------------  OpenItemCd()  -------------------------------------------------
Function OpenItemCd2()
	Dim arrRet
	Dim arrParam(5), arrField(1)
	Dim iCalledAspName
	Dim IntRetCD

	If lgIsOpenPop = True Then Exit Function

	if  Trim(frm1.txtPlantCd2.Value) = "" then
		Call DisplayMsgBox("17A002", "X", "공장", "X")
		frm1.txtPlantCd2.focus
		Exit Function
	End if

	lgIsOpenPop = True

	arrParam(0) = Trim(frm1.txtPlantCd2.value)	' Plant Code
	arrParam(1) = Trim(frm1.txtItemCd2.value)		' Item Code
	arrParam(2) = "!"	' “12!MO"로 변경 -- 품목계정 구분자 조달구분 
	arrParam(3) = "30!P"
	arrParam(4) = ""		'-- 날짜 
	arrParam(5) = ""		'-- 자유(b_item_by_plant a, b_item b: and 부터 시작)

	arrField(0) = 1 ' -- 품목코드 
	arrField(1) = 2 ' -- 품목명 

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
		frm1.txtItemCd2.focus
		Exit Function
	Else
		frm1.txtItemCd2.Value	= arrRet(0)
		frm1.txtItemNm2.Value	= arrRet(1)
		frm1.txtItemCd2.focus
	End If
End Function

Function OpenSupplier(byval strcomp)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtSupplierCd1.className) = UCase(parent.UCN_PROTECTED) And strComp="Supplier1" Then Exit Function
	If IsOpenPop = True Or UCase(frm1.txtSupplierCd2.className) = UCase(parent.UCN_PROTECTED) And strComp="Supplier2" Then Exit Function

	IsOpenPop = True

	arrParam(0) = "공급처"  
	arrParam(1) = "B_Biz_Partner" 
	 
	If strcomp="Supplier1" Then
		arrParam(2) = Trim(frm1.txtSupplierCd1.Value) 
	Else
		arrParam(2) = Trim(frm1.txtSupplierCd2.Value) 
	End if
	  
	If strComp="Supplier1" Then
		arrParam(4) = "BP_TYPE In (" & FilterVar("S", "''", "S") & " ," & FilterVar("CS", "''", "S") & ")"
	Else
		arrParam(4) = "BP_TYPE In (" & FilterVar("S", "''", "S") & " ," & FilterVar("CS", "''", "S") & ") And USAGE_FLAG=" & FilterVar("Y", "''", "S") & " " 
	End if
	 
	arrParam(5) = "공급처"       
	 
	arrField(0) = "BP_CD"   
	arrField(1) = "BP_NM"   
	    
	arrHeader(0) = "공급처"  
	arrHeader(1) = "공급처명" 
	    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
	"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		If strComp="Supplier1" Then
			frm1.txtSupplierCd1.focus
		Else
			frm1.txtSupplierCd2.focus
		End If
		Exit Function
	Else
		If strComp="Supplier1" Then
			frm1.txtSupplierCd1.Value    = arrRet(0)  
			frm1.txtSupplierNm1.Value    = arrRet(1)  
			frm1.txtSupplierCd1.focus
		Else
			frm1.txtSupplierCd2.Value    = arrRet(0)  
			frm1.txtSupplierNm2.Value    = arrRet(1)  
			lgBlnFlgChgValue = True  
			frm1.txtSupplierCd2.focus
		End If 
	End If 
End Function

Function OpenUnit()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtUnit.className) = UCase(parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "발주단위"    
	arrParam(1) = "B_Unit_OF_MEASURE"  
	 
	arrParam(2) = Trim(frm1.txtUnit.Value) 
	 
	arrParam(4) = ""      
	arrParam(5) = "발주단위"    
	 
	arrField(0) = "Unit"     
	arrField(1) = "Unit_Nm"     
	    
	arrHeader(0) = "발주단위"   
	arrHeader(1) = "발주단위명"   
	    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
	"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtUnit.focus
		Exit Function
	Else
		frm1.txtUnit.Value    = arrRet(0)  
		frm1.txtUnit.focus
		lgBlnFlgChgValue = True  	
	End If 
End Function

Function OpenGroup()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtGroupCd.className) = UCase(parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "구매그룹" 
	arrParam(1) = "B_Pur_Grp"    
	 
	arrParam(2) = Trim(frm1.txtGroupCd.Value)
	 
	arrParam(4) = "Usage_flg=" & FilterVar("Y", "''", "S") & " "
	If Trim(frm1.hdnOrg.value) <> "" Then
		arrParam(4) = arrParam(4) & " And pur_org= " & FilterVar(frm1.hdnOrg.value, "''", "S") & ""
	End if
	arrParam(5) = "구매그룹"   
	 
	arrField(0) = "PUR_GRP" 
	arrField(1) = "PUR_GRP_NM" 
	    
	arrHeader(0) = "구매그룹" 
	arrHeader(1) = "구매그룹명"  
	    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
	"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	 
	IsOpenPop = False
	 
	If arrRet(0) = "" Then
		frm1.txtGroupCd.focus
		Exit Function
	Else
		frm1.txtGroupCd.Value= arrRet(0)  
		frm1.txtGroupNm.Value= arrRet(1) 
		frm1.txtGroupCd.focus
		lgBlnFlgChgValue = True 	
	End If 
End Function 

<!-- '==========================================   ChangeItemPlant()  ======================================
' Name : ChangeItemPlant()
' Description : 
'========================================================================================================= !-->
Sub ChangeItemPlant()
    Dim strVal
    Err.Clear                               

	If gLookUpEnable = False Then
		Exit Sub
	End If

	If Trim(frm1.txtPlantCd2.Value) = "" Or Trim(frm1.txtItemCd2.Value) = "" Then
		Exit Sub
	End if
  
    If LayerShowHide(1) = False Then
       Exit Sub 
    End If    
        
    With frm1
		strVal = BIZ_PGM_ID & "?txtMode="	& "LookUpItemPlant"  
		strVal = strVal & "&txtPlantCd="	& Trim(.txtPlantCd2.value) 
		strval = strval & "&txtItemCd="		& Trim(.txtItemCd2.vaLue)
    End With
    
	Call RunMyBizASP(MyBizASP, strVal)       
End Sub


Sub Setchangeflg()
 lgBlnFlgChgValue = True 
End Sub


Sub Changeflg()
	If frm1.rdoUseflg(0).checked = True Then
		frm1.txtUseflg.value= "Y"
	Else
		frm1.txtUseflg.value= "N"
	End If 
End Sub

<!-- 
'#########################################################################################################
'            3. Event부 
' 기능: Event 함수에 관한 처리 
' 설명: Window처리, Single처리, Grid처리 작업.
'         여기서 Validation Check, Calcuration 작업이 가능한 Event가 발생.
'         각 Object단위로 Grouping한다.
'##########################################################################################################!-->

Sub rdoUseflgY_onClick()
 lgBlnFlgChgValue = True  
End Sub

Sub rdoUseflgN_onClick()
 lgBlnFlgChgValue = True  
End Sub

Sub rdoDefFlgY_onClick()
 lgBlnFlgChgValue = True  
End Sub

Sub rdoDefFlgN_onClick()
 lgBlnFlgChgValue = True  
End Sub

<!-- '==========================================  3.1.1 Form_Load()  ======================================
' Name : Form_Load()
' Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분 
'========================================================================================================= !-->
 Sub Form_Load()
    Call LoadInfTB19029
    Call AppendNumberRange("0","1","99")
    Call AppendNumberRange("1","0","999")
    Call ggoOper.FormatField(Document, "2", ggStrIntegeralPart, ggStrDeciPointPart,gDateFormat,parent.gComNum1000,parent.gComNumDec,,, ggStrMinPart,ggStrMaxPart) 
    Call ggoOper.LockField(Document, "N")      
    Call SetDefaultVal
    Call InitVariables           
    Call ggoOper.FormatNumber(frm1.txtQuotaRate,"99999999","0",true,ggExchRate.DecPoint,parent.gComNumDec,parent.gComNum1000)
	Call ReadCookiePage()
	Call Changeflg() 
End Sub

<!--
'==========================================================================================
'   Event Name : OCX Event
'   Event Desc :
'==========================================================================================
-->
Sub txtFrDt_DblClick(Button)
	if Button = 1 then
		frm1.txtFrDt.Action = 7
        Call SetFocusToDocument("M")  
        frm1.txtFrDt.Focus
	End if
End Sub

Sub txtToDt_DblClick(Button)
	if Button = 1 then
		frm1.txtToDt.Action = 7
        Call SetFocusToDocument("M")  
        frm1.txtToDt.Focus
	End if
End Sub

Sub txtFrDt_Change()
	lgBlnFlgChgValue = true 
	 '-- Modify for issue 9055 by Byun Jee Hyun 2004-12-06
	if frm1.txtFrDt.text <> "" and frm1.txtToDt.text <> "" then
		if UniConvDateToYYYYMMDD(frm1.txtFrDt.text,parent.gDateFormat,"") > UniConvDateToYYYYMMDD(frm1.txtToDt.text,parent.gDateFormat,"") then
			Call DisplayMsgBox("970025", "X", "유효시작일", "유효종료일")
			frm1.txtFrDt.text = ""
			frm1.txtFrDt.Focus
		end if
	end if
End Sub

Sub txtToDt_Change()	
	lgBlnFlgChgValue = true 
	' -- Modify for issue 9055 by Byun Jee Hyun 2004-12-06
	if frm1.txtFrDt.text <> "" and frm1.txtToDt.text <> "" then
		if UniConvDateToYYYYMMDD(frm1.txtFrDt.text,parent.gDateFormat,"") > UniConvDateToYYYYMMDD(frm1.txtToDt.text,parent.gDateFormat,"") then
			Call DisplayMsgBox("970023", "X", "유효종료일", "유효시작일")
			frm1.txtToDt.text = ""
			frm1.txtToDt.Focus
		end if
	end if
End Sub

Sub txtPriority_Change()
	lgBlnFlgChgValue = true 
End Sub

Sub txtPurlt_Change()
	lgBlnFlgChgValue = true 
End Sub

Sub txtMinQty_Change()
	lgBlnFlgChgValue = true 
End Sub

Sub txtMaxQty_Change()
	lgBlnFlgChgValue = true 
End Sub

Sub txtOver_Change()
	lgBlnFlgChgValue = true 
End Sub

Sub txtUnder_Change()
	lgBlnFlgChgValue = true 
End Sub

<!--
'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================
-->
Function FncQuery() 
    Dim IntRetCD 
    Err.Clear                                                   
    
    FncQuery = False                                            

    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then
		Exit Function
		End If
    End If
    
    Call ggoOper.ClearField(Document, "2")      
    Call InitVariables           
    
    If Not chkField(Document, "1") Then       
       Exit Function
    End If
    
	Call Changeflg()
    
    If DbQuery = False Then Exit Function
       
    FncQuery = True            
	Set gActiveElement = document.activeElement
End Function

<!--
'========================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'========================================================================================
-->
Function FncNew() 
    Dim IntRetCD 
    
    FncNew = False                                              
    
    If lgBlnFlgChgValue = True Then
        IntRetCD = DisplayMsgBox("900015", parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    Call ggoOper.ClearField(Document, "A")                      
    Call ggoOper.LockField(Document, "N")                       
    Call SetDefaultVal
    Call InitVariables
    
    FncNew = True            
	Set gActiveElement = document.activeElement
End Function

<!--
'========================================================================================
' Function Name : FncDelete
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
-->
Function FncDelete() 
	Dim IntRetCD

    IntRetCD = DisplayMsgBox("900003", parent.VB_YES_NO, "X", "X")
    If IntRetCD = vbNo Then Exit Function

    FncDelete = False     
    
	If lgIntFlgMode <> parent.OPMD_UMODE Then                          
        Call DisplayMsgBox("900002", "X", "X", "X")
        Exit Function
    End If

    If Not chkField(Document, "1") Then       
       Exit Function
    End If
        
    If DbDelete = False Then Exit Function
    
    FncDelete = True                                            
	Set gActiveElement = document.activeElement
End Function

<!--
'========================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
-->
 Function FncSave() 
    Dim IntRetCD 
    Err.Clear                                                   
    
    FncSave = False                                             
    
    If lgBlnFlgChgValue = False Then
        IntRetCD = DisplayMsgBox("900001", "X", "X", "X")
        Exit Function
    End If
    
    If Not chkField(Document, "2") Then                         
       Exit Function
    End If
    
	With frm1

	If CompareDateByFormat(.txtFrDt.text,.txtToDt.text,.txtFrDt.Alt,.txtToDt.Alt, "970025",.txtFrDt.UserDefinedFormat,parent.gComDateType,False) = False And Trim(.txtFrDt.text) <> "" And Trim(.txtToDt.text) <> "" Then
		Call DisplayMsgBox("17a003", "X","유효일","X")
		Exit Function
	End if  
  
	End With    

    If DbSave = False Then Exit Function
    
    FncSave = True                                              
    Call Changeflg()
	Set gActiveElement = document.activeElement
End Function

<!--
'========================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================
-->
Function FncCopy() 
	Dim IntRetCD
    
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900017", parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    call SetDefaultVal
    
    lgIntFlgMode = parent.OPMD_CMODE         
    
    Call ggoOper.ClearField(Document, "1")                                 
    Call ggoOper.LockField(Document, "N")         
    
    Call Changeflg()
    
    frm1.txtSupplierCd2.value = ""
    frm1.txtSupplierNm2.value = ""    
	Set gActiveElement = document.activeElement
End Function

<!--
'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================
-->
Function FncPrint() 
    Call parent.FncPrint()
	Set gActiveElement = document.activeElement
End Function

<!--
'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================
-->
Function FncExcel() 
    Call parent.FncExport(parent.C_SINGLE)          
	Set gActiveElement = document.activeElement
End Function
<!--
'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================
-->
Function FncFind() 
    Call parent.FncFind(parent.C_SINGLE , False)                               
	Set gActiveElement = document.activeElement
End Function

<!--
'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================
-->
Function FncExit()
	Dim IntRetCD
	
	FncExit = False
	
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
	Set gActiveElement = document.activeElement
    FncExit = True
End Function

<!--
'========================================================================================
' Function Name : DbDelete
' Function Desc : This function delete data
'========================================================================================
-->
Function DbDelete() 
    Dim strVal
    Err.Clear                                                           

    DbDelete = False             
    
    If LayerShowHide(1) = False Then
       Exit Function 
    End If
    
    With frm1
    
    strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0003      
    strVal = strVal & "&txtPlantCd1=" & Trim(.txtPlantCd1.value)  
    strVal = strVal & "&txtItemCd1=" & Trim(.txtItemCd1.value)
    strVal = strVal & "&txtSupplierCd1=" & Trim(.txtSupplierCd1.value)
    
    End With
    
	Call RunMyBizASP(MyBizASP, strVal)         
 
    DbDelete = True                                                     
End Function

<!--
'========================================================================================
' Function Name : DbDeleteOk
' Function Desc : DbDelete가 성공적일때 수행 
'========================================================================================
-->
Function DbDeleteOk()             
	Call FncNew()
End Function

<!--
'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
-->
 Function DbQuery() 
    Dim strVal
    Err.Clear                                                           
    
    DbQuery = False                                                     
    
    If LayerShowHide(1) = False Then
       Exit Function 
    End If
    
    With frm1
    strVal = BIZ_PGM_ID & "?txtMode="		& parent.UID_M0001      
    strVal = strVal & "&txtPlantCd1="		& Trim(.txtPlantCd1.value)
    strval = strval & "&txtItemCd1="		& Trim(.txtItemCd1.value)
    strval = strVal & "&txtSupplierCd1="	& Trim(.txtSupplierCd1.value)
    End With
    
	Call RunMyBizASP(MyBizASP, strVal)         
 
    DbQuery = True
End Function
<!--
'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================
-->
Function DbQueryOk()             
    '-----------------------
    'Reset variables area
    '-----------------------
    lgIntFlgMode = parent.OPMD_UMODE           
    lgBlnFlgChgValue = False
    
    Call ggoOper.LockField(Document, "Q")        
	frm1.txtGroupCd.focus
	Call SetToolbar("11111000001111")
End Function

<!--
'========================================================================================
' Function Name : DbQueryOk1
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================
-->
Function DbQueryOk1()             
    '-----------------------
    'Reset variables area
    '-----------------------
	Call ggoOper.LockField(Document, "N")
	frm1.txtPlantCd2.value = frm1.txtPlantCd1.value
	frm1.txtPlantNm2.value = frm1.txtPlantNm1.value
	frm1.txtitemcd2.value = frm1.txtitemcd1.value
	frm1.txtitemNm2.value = frm1.txtitemNm1.value
	frm1.txtSuppliercd2.value = frm1.txtSuppliercd1.value
	frm1.txtSupplierNm2.Value = frm1.txtSupplierNm1.Value
	frm1.txtGroupCd.focus
	
	'조회후 데이터없을때 자동으로 조회부의 데이터를 데이터부에 옮겨주는데 이때 ChangeItemPlant를 거쳐 
	'구매조직 히든값을 가져오도록한다.	200309
	Call ChangeItemPlant()
		
    Call SetToolbar("1110100000001111")  
End Function

<!--
'========================================================================================
' Function Name : DBSave
' Function Desc : 실제 저장 로직을 수행 , 성공적이면 DBSaveOk 호출됨 
'========================================================================================
-->
Function DbSave()
    Dim strVal
    Err.Clear               
	DbSave = False              

	If UNICDbl(frm1.txtMinQty.text) <> 0 And UNICDbl(frm1.txtMaxQty.text) <> 0 Then
    
		If UNICDbl(frm1.txtMinQty.text) > UNICDbl(frm1.txtMaxQty.text) Then
			Call DisplayMsgBox("171324","X","X","X")   
			Exit Function 
		End If
		
    end if
  
    If LayerShowHide(1) = False Then
		Exit Function 
    End If

	With frm1
		.txtMode.value = parent.UID_M0002          
		.txtFlgMode.value = lgIntFlgMode
  
		If .rdoDefFlg(0).checked = true Then
		 .txtDefFlg.Value = "Y"
		Else
		 .txtDefFlg.Value = "N"
		End If
  
		Call ExecMyBizASP(frm1, BIZ_PGM_ID)          
	End With
 
    DbSave = True                                                           
End Function
<!--
'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncSave에 있는것을 옮김 
'========================================================================================
-->
Function DbSaveOk()              
    Call InitVariables
    Call MainQuery()
End Function

</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" --> 
</HEAD>
<BODY TABINDEX="-1" SCROLL="NO">
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
     <TD>
      <TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
       <TR>
        <td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
        <td background="../../../CShared/image/table/seltab_up_bg.gif" CLASS="CLSMTAB" align="center"><font color=white>품목별공급처</font></td>
        <td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
          </TR>
      </TABLE>
     </TD>
     <TD WIDTH=500>&nbsp;</TD>
     <TD WIDTH=*>&nbsp;</TD>
    </TR>
   </TABLE>
  </TD>
 </TR>
 <TR>
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
         <TD CLASS="TD5" NOWRAP>공장</TD>
         <TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="공장"   NAME="txtPlantCd1" SIZE=10 MAXLENGTH=4 tag="12NXXU" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnORGCd1" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPlant('Plant1')" OnMouseOver="vbscript:PopUpMouseOver()" OnMouseOut="vbscript:PopUpMouseOut()">
             <INPUT TYPE=TEXT ALT="공장" ID="txtPlantNm1" NAME="arrCond" tag="14X"></TD>
         <TD CLASS="TD5" NOWRAP>품목</TD>
         <TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="품목" NAME="txtItemCd1"   SIZE=20 MAXLENGTH=18 tag="12NXXU" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnGroup1Cd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenItemCd1()" OnMouseOver="vbscript:PopUpMouseOver()" OnMouseOut="vbscript:PopUpMouseOut()">
             </TD>
        </TR>
        <tr>
         <TD CLASS="TD5" NOWRAP>공급처</TD>
         <TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="공급처"   NAME="txtSupplierCd1" SIZE=10 MAXLENGTH=10 tag="12NXXU" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnGroup1Cd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenSupplier('Supplier1')">
             <INPUT TYPE=TEXT ALT="공급처" ID="txtSupplierNm1" NAME="arrCond" tag="14X"></TD>
         <TD CLASS="TD5" NOWRAP></TD>
         <TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="품목" ID="txtItemNm1" NAME="arrCond" tag="14X" SIZE=35></TD>
        </tr>
       </TABLE>
      </FIELDSET>
     </TD>
    </TR>
    <TR>
     <TD <%=HEIGHT_TYPE_03%>></TD>
    </TR>
    <TR>
     <TD WIDTH=100% valign=top>
       <TABLE <%=LR_SPACE_TYPE_60%>>
        <TR>
         <TD CLASS="TD5" NOWRAP>공장</TD>
         <TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="공장" NAME="txtPlantCd2" SIZE=10 MAXLENGTH=4 tag="23NXXU" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnORG2Cd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPlant('Plant2')" OnMouseOver="vbscript:PopUpMouseOver()" OnMouseOut="vbscript:PopUpMouseOut()">
                 <INPUT TYPE=TEXT ALT="공장" NAME="txtPlantNm2" SIZE=20 tag="24x"></TD>
         <TD CLASS="TD5" NOWRAP>품목</TD>
         <TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="품목"  NAME="txtItemCd2" SIZE=10 MAXLENGTH=18 tag="23NXXU" ONCHANGE="ChangeItemPlant()"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnORG2Cd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenItemCd2()" OnMouseOver="vbscript:PopUpMouseOver()" OnMouseOut="vbscript:PopUpMouseOut()">
                 <INPUT TYPE=TEXT ALT="품목" NAME="txtItemNm2" SIZE=20 tag="24x"></TD>
        </TR>
        <TR>
         <TD CLASS="TD5" NOWRAP>공급처</TD>
         <TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="공급처"   NAME="txtSupplierCd2" SIZE=10 MAXLENGTH=10 tag="23NXXU" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnGroupCd2" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenSupplier('Supplier2')" OnMouseOver="vbscript:PopUpMouseOver()" OnMouseOut="vbscript:PopUpMouseOut()">
                 <INPUT TYPE=TEXT ALT="공급처" NAME="txtSupplierNm2" SIZE=20 tag="24X" ></TD>
         <TD CLASS="TD5" NOWRAP>구매그룹</TD>
         <TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="구매그룹"  NAME="txtGroupCd" SIZE=10 MAXLENGTH=4 tag="22XXXU" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnGroupCd2" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenGroup()" OnMouseOver="vbscript:PopUpMouseOver()" OnMouseOut="vbscript:PopUpMouseOut()">
                 <INPUT TYPE=TEXT ALT="구매그룹" NAME="txtGroupNm" SIZE=20 tag="24X" ></TD>
        </TR>
        <TR>
         <TD CLASS="TD5" NOWRAP>발주배정가중치</TD>
         <TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT="발주배정가중치" NAME="txtPriority" CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle1 style="HEIGHT: 20px; WIDTH: 80px" tag="22XX0" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT>
         &nbsp;배분비율(%)<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT=배분비 NAME="txtQuotaRate" CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle5 style="HEIGHT: 20px; WIDTH: 80px" tag="24X" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT></TD>
         <TD CLASS="TD5" NOWRAP>구매L/T</TD>
         <TD CLASS="TD6" NOWRAP>
          <Table cellpadding=0 cellspacing=0>
           <TR>
            <TD NOWRAP>
             <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT=구매L/T NAME="txtPurlt"CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle2 NAME="txtPurlt" style="HEIGHT: 20px; WIDTH: 80px" tag="21XX1" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT>
            </TD>
            <TD WIDTH="*" NOWRAP>&nbsp;일
            </TD>
           </TR>
          </Table>
         </TD>
        </TR>
        <TR>
         <TD CLASS="TD5" NOWRAP>사용여부</TD>
         <TD CLASS="TD6" NOWRAP><INPUT TYPE=radio ALT="사용여부" class="radio" NAME="rdoUseflg" id="rdoUseflgY" checked Value="Y" tag="21"><label for="rdoUseflgY">예</label>
                 <INPUT TYPE=radio ALT="사용여부" class="radio" NAME="rdoUseflg" id="rdoUseflgN" Value="N" tag="21"><label for="rdoUseflgN">아니오</label></TD>
         <TD CLASS="TD5" NOWRAP>주공급업체</TD>
         <TD CLASS="TD6" NOWRAP><INPUT TYPE=radio ALT="주공급업체" class="radio" NAME="rdoDefFlg" id="rdoDefFlgY" Value="Y" tag="21"><label for="rdoDefFlgY">예</label>
                 <INPUT TYPE=radio ALT="주공급업체" class="radio" NAME="rdoDefFlg" id="rdoDefFlgN" checked Value="N" tag="21"><label for="rdoDefFlgN">아니오</label></TD>
        </TR>
        <TR>
         <TD CLASS="TD5" NOWRAP>최소발주량</TD>
         <TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT=최소발주량 NAME="txtMinQty" CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle3 style="HEIGHT: 20px; WIDTH: 160px" tag="21X3Z" Title="FPDOUBLESINGLE" ALT=최소발주량></OBJECT>');</SCRIPT></td>
         <TD CLASS="TD5" NOWRAP>최대발주량</TD>
         <TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT=최대발주량 NAME="txtMaxQty" CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle3 style="HEIGHT: 20px; WIDTH: 160px" tag="21X3Z" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT></td>
        </TR>
        <TR>
         <TD CLASS="TD5" NOWRAP>발주단위</TD>
         <TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="발주단위"  NAME="txtUnit" SIZE=10 MAXLENGTH=3 tag="21XNXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnGroupCd2" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenUnit()"></td>
         <TD CLASS="TD5" NOWRAP>주창고</TD>
         <TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="주창고" NAME="txtStorageCd" SIZE=10 tag="24XXXU">&nbsp;&nbsp;&nbsp;&nbsp;
                 <INPUT TYPE=TEXT ALT="주창고" NAME="txtStorageNm" SIZE=20 tag="24X" ></TD>
        </TR>
        <TR>
         <TD CLASS="TD5" NOWRAP>과부족허용율(+)</TD>
         <TD CLASS="TD6" NOWRAP>
          <Table cellpadding=0 cellspacing=0>
           <TR>
            <TD NOWRAP>
             <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT=과부족허용율(+) NAME="txtOver" CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle5 style="HEIGHT: 20px; WIDTH: 160px" tag="21X5Z" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT>
            </TD>
             <TD WIDTH="*" NOWRAP>&nbsp;%
            </TD>
           </TR>
          </Table>
         </TD>
         <TD CLASS="TD5" NOWRAP>과부족허용율(-)</TD>
         <TD CLASS="TD6" NOWRAP>
          <Table cellpadding=0 cellspacing=0>
           <TR>
            <TD NOWRAP>
             <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT=과부족허용율(-) NAME="txtUnder" CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle6 style="HEIGHT: 20px; WIDTH: 160px" tag="21X5Z" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT>
            </TD>
             <TD WIDTH="*" NOWRAP>&nbsp;%
            </TD>
           </TR>
          </Table>
         </TD>
        </TR>
        <TR>
         <TD CLASS="TD5" NOWRAP>유효시작일</TD>
         <TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 NAME="txtFrdt" style="HEIGHT: 20px; WIDTH: 100px" tag="21X1" Title="FPDATETIME" ALT=유효시작일></OBJECT>');</SCRIPT></TD>
         <TD CLASS="TD5" NOWRAP>유효종료일</TD>
         <TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime2 NAME="txtTodt" style="HEIGHT: 20px; WIDTH: 100px" tag="21X1" Title="FPDATETIME" ALT=유효종료일></OBJECT>');</SCRIPT></TD>
        </TR>
        <TR>
         <TD CLASS="TD5" NOWRAP>공급처품목코드</TD>
         <TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="공급처품목코드" NAME="txtSpplCd" SIZE=34 MAXLENGTH=20 tag="21"></TD>
         <TD CLASS="TD5" NOWRAP>공급처품목명</TD>
         <TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="공급처품목명" NAME="txtSpplNm" SIZE=34 MAXLENGTH=50 tag="21"></TD>
        </TR>
        <TR>
         <TD CLASS="TD5" NOWRAP>공급처품목규격</TD>
         <TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="공급처품목규격" NAME="txtSpplSpec" SIZE=34 MAXLENGTH=50 tag="21"></TD>
         <TD CLASS="TD5" NOWRAP>제조사</TD>
         <TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="제조사" NAME="txtMakerNm" SIZE=34 MAXLENGTH=50 tag="21"></TD>
        </TR>
        <TR>
         <TD CLASS="TD5" NOWRAP>공급처영업담당</TD>
         <TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="공급처영업담당" NAME="txtSpplPrsn" SIZE=34 MAXLENGTH=50 tag="21"></TD>
         <TD CLASS="TD5" NOWRAP>공급처연락처</TD>
         <TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="공급처긴급연락처" NAME="txtTel" SIZE=34 MAXLENGTH=20 tag="21"></TD>
        </TR>
        <%Call SubFillRemBodyTD5656(5)%>
       </TABLE>
     </TD> 
    </TR>
   </table>    
  </TR>
    <tr>
      <td <%=HEIGHT_TYPE_01%>></td>
    </tr>
    <tr HEIGHT="20">
  <td WIDTH="100%">
   <table <%=LR_SPACE_TYPE_30%>>
    <tr>
     <td WIDTH="10"></td>
     <td WIDTH="*" align="right"><a href="VBScript:WriteCookiePage()">품목별공급처조회</a></td>
     <td WIDTH="10"></td>
    </tr>
   </table>
  </td>
    </tr>
 <TR>
  <TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME SRC="../../blank.htm" NAME="MyBizASP"  WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=No noresize framespacing=0 tabindex = -1></IFRAME>
  </TD>
 </TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" tabindex = -1>
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24" tabindex = -1 >
<INPUT type=hidden name="txtuseflg" tag="24" tabindex = -1>
<INPUT type=hidden name="txtDefflg" tag="24" tabindex = -1>
<INPUT type=hidden name="hdnOrg" tag="24" tabindex = -1>
</FORM>

    <DIV ID="MousePT" NAME="MousePT">
    <iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
    </DIV>

</BODY>
</HTML>
