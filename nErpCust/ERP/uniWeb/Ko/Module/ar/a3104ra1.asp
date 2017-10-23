
<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : 입금정보팝업
*  2. Function Name        : 
*  3. Program ID           : 
*  4. Program Name         : 
*  5. Program Desc         :  Ado query Sample with DBAgent(Sort)
*  6. Comproxy List        :
*  7. Modified date(First) : 2001/04/18
*  8. Modified date(Last)  : 2001/04/18
*  9. Modifier (First)     :
* 10. Modifier (Last)      :
* 11. Comment              :
=======================================================================================================-->
<HTML>
<HEAD>
<TITLE>입금정보팝업</TITLE>

<!--
########################################################################################################
#						   3.    External File Include Part
########################################################################################################-->

<!--
========================================================================================================
=                          3.1 Server Side Script
========================================================================================================-->

<!-- #Include file="../../inc/IncSvrCcm.inc"  -->
<!-- #Include file="../../inc/IncSvrHTML.inc"  -->

<!--
========================================================================================================
=                          3.2 Style Sheet
======================================================================================================== -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		

<!--
========================================================================================================
=                          3.3 Client Side Script
======================================================================================================== -->
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliPAMain.vbs">			</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliPAEvent.vbs">			</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliPAOperation.vbs">		</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliVariables.vbs">		</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliDBAgentA.vbs">			</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliDBAgentVariables.vbs">	</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliRdsQuery.vbs">			</SCRIPT>
<SCRIPT LANGUAGE = "JavaScript"SRC = "../../inc/incImage.js">			</SCRIPT>

<Script Language="VBScript">
Option Explicit                                                        '☜: indicates that All variables must be declared in advance
	

'########################################################################################################
'#                       4.  Data Declaration Part
'########################################################################################################

'========================================================================================================
'=                       4.1 External ASP File
'========================================================================================================

Const BIZ_PGM_ID 		= "a3104rb1.asp"                              '☆: Biz Logic ASP Name

'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
Const C_SHEETMAXROWS_D  = 30                                          '☆: Server에서 한번에 fetch할 최대 데이타 건수 
Const C_MaxKey          = 3					                          '☆: SpreadSheet의 키의 갯수 

'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
Dim lgIsOpenPop                                          
Dim lgPopUpR                                              

Dim Bp_sts
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim IsOpenPop

Dim arrReturn
Dim arrParent
Dim arrParam

' 권한관리 추가 
Dim lgAuthBizAreaCd	' 사업장 
Dim lgInternalCd	' 내부부서 
Dim lgSubInternalCd	' 내부부서(하위포함)
Dim lgAuthUsrID		' 개인 

	arrParent = window.dialogArguments
	Set PopupParent = arrParent(0)
	arrParam = arrParent(1)


	 '------ Set Parameters from Parent ASP ------ 
	
	top.document.title = "입금정보팝업"

'########################################################################################################
'#                       5.Method Declaration Part
'########################################################################################################

'========================================================================================================
'========================================================================================================
'                        5.1 Common Method-1
'========================================================================================================
'========================================================================================================

'========================================================================================================
' Name : InitVariables()	
' Desc : Initialize value
'========================================================================================================	
Sub InitVariables()

    Redim arrReturn(0)

    lgStrPrevKey     = ""
    lgPageNo         = ""
    lgIntFlgMode     = PopupParent.OPMD_CMODE                          'Indicates that current mode is Create mode
	lgBlnFlgChgValue = False
    lgSortKey        = 1
    
	Self.Returnvalue = arrReturn

	' 권한관리 추가 
	If UBound(arrParam) > 5 Then
		lgAuthBizAreaCd		= arrParam(5)
		lgInternalCd		= arrParam(6)
		lgSubInternalCd		= arrParam(7)
		lgAuthUsrID			= arrParam(8)
	End If
	
End Sub

'========================================================================================================
' Name : SetDefaultVal()	
' Desc : Set default value
'========================================================================================================
Sub SetDefaultVal()

	Dim strSvrDate
	Dim frDt, toDt
	strSvrDate = "<%=GetSvrDate%>"
	frDt = UNIDateAdd("M", -1, strSvrDate,PopupParent.gServerDateFormat)
		
	frm1.txtFrPrDt.Text = UniConvDateAToB(frDt ,PopupParent.gServerDateFormat,PopupParent.gDateFormat) 
	frm1.txtToPrDt.Text = UniConvDateAToB("<%=GetSvrDate%>" ,PopupParent.gServerDateFormat,PopupParent.gDateFormat) 
	Bp_sts = "Y"
End Sub
'------------------------------------------  OpenBp()  ---------------------------------------
'	Name : OpenBp()
'	Description : OpenAp Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function OpenBp(Byval strCode, byval iWhere)
	Dim arrRet
	Dim arrParam(5)

	If IsOpenPop = True Or UCase(frm1.txtBPCd.className) = "PROTECTED" Then Exit Function

	IsOpenPop = True

	arrParam(0) = strCode								'  Code Condition
   	arrParam(1) = ""							' 채권과 연계(거래처 유무)
	arrParam(2) = ""									'FrDt
	arrParam(3) = ""									'ToDt
	arrParam(4) = "B"							'B :매출 S: 매입 T: 전체 
	arrParam(5) = "PAYER"									'SUP :공급처 PAYTO: 지급처 SOL:주문처 PAYER :수금처 INV:세금계산 	
	arrRet = window.showModalDialog("../../Module/ar/BpPopup.asp", Array(window.PopupParent,arrParam), _
		"dialogWidth=780px; dialogHeight=450px; : Yes; help: No; resizable: No; status: No;")
		
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtBPCd.focus

		Exit Function
	Else
		frm1.txtBPCd.value=arrRet(0)
		frm1.txtBPNm.value= arrRet(1)
		frm1.txtBPCd.focus

	End If	
End Function
'========================================================================================================
' Name : OpenPopUp()	
' Desc : 
'========================================================================================================
Function OpenPopUp(Byval strCode, Byval iWhere )
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	
		Select Case iWhere
			Case 1
				if frm1.Chk_Bp.checked = false then
					arrParam(0) = "수금처 팝업"				' 팝업 명칭 
					arrParam(1) = "B_BIZ_PARTNER"    			' TABLE 명칭 
					arrParam(2) = strCode						' Code Condition
					arrParam(3) = ""							' Name Cindition
					arrParam(4) = ""							' Where Condition
					arrParam(5) = "수금처코드"					' 조건필드의 라벨 명칭 

					arrField(0) = "BP_CD"	     				' Field명(0)
					arrField(1) = "BP_NM"			    		' Field명(1)
    
					arrHeader(0) = "수금처코드"					' Header명(0)
					arrHeader(1) = "수금처명"				' Header명(1)
				Else
					IsOpenPop = False
					Exit function
				ENd if	
			Case 2
				arrParam(0) = "가수금유형"								' 팝업 명칭 
				arrParam(1) = "a_jnl_item"	 									' TABLE 명칭 
				arrParam(2) = strCode										' Code Condition
				arrParam(3) = ""												' Name Cindition
				arrParam(4) = "jnl_type = " & FilterVar("SR", "''", "S") & " "								' Where Condition
				arrParam(5) = "가수금유형"								' 조건필드의 라벨 명칭 

				arrField(0) = "JNL_CD"											' Field명(0)
				arrField(1) = "JNL_NM"											' Field명(1)
    
				arrHeader(0) = "가수금유형"								' Header명(0)
				arrHeader(1) = "가수금유형명"							' Header명(1)
			
		End Select 


	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
	     "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Call EscPopUp(iWhere)
		Exit Function
	Else
		Call SetPopUp(arrRet,iWhere)
	End If	

End Function
'===========================================================================================
' Name : EscPopUp()
' Description :
'===========================================================================================
Function EscPopUp(iWhere)
	
	With frm1
		Select Case iWhere
			Case 1
				.txtBPCd.focus
			Case 2
				.txtRcptType.focus
		End Select
	End With

End Function
'===========================================================================================
' Name : SetPopUp()
' Description : Plant Popup에서 Return되는 값 setting
'===========================================================================================
Function SetPopUp(arrRet,iWhere)
	
	With frm1
		Select Case iWhere
			Case 1
				.txtBPCd.value	 = arrRet(0)
				.txtBPNm.value	 = arrRet(1)
				.txtBPCd.focus
			Case 2
				.txtRcptType.Value			= arrRet(0)
				.txtRcptTypeNm.Value		= arrRet(1)
				.txtRcptType.focus
		End Select
	End With

End Function

'========================================================================================================
' Name : LoadInfTB19029()	
' Desc : Set System Number format
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "A","NOCOOKIE","RA") %>                                '☆: 
	<% Call LoadBNumericFormatA("Q", "A", "NOCOOKIE", "RA") %>
End Sub

'===========================================  2.3.1 OkClick()  ==========================================
'=	Name : OkClick()																					=
'=	Description : Return Array to Opener Window when OK button click									=
'=				  이 부분에서 컬럼 추가하고 데이타 전송이 일어나야 합니다.   							=
'========================================================================================================
Function OKClick()
	If frm1.vspdData.ActiveRow > 0 Then 				
		Redim arrReturn(1)
		frm1.vspdData.Row	= frm1.vspdData.ActiveRow
		frm1.vspdData.Col	= GetKeyPos("A",1)		
		arrReturn(0)		= frm1.vspdData.Text
	End if			
		
	Self.Returnvalue = arrReturn
	Self.Close()
End Function


'=========================================  2.3.2 CancelClick()  ========================================
'=	Name : CancelClick()																				=
'=	Description : Return Array to Opener Window for Cancel button click 								=
'========================================================================================================

Function CancelClick()
	Self.Close()			
End Function

'========================================================================================================
'	Name : CookiePage()
'	Description : JUMP시 Load화면으로 조건부로 Value
'========================================================================================================
Function CookiePage(ByVal Kubun)
		
End Function

'========================================================================================================
' Name : InitComboBox()
' Desc : Set ComboBox
'========================================================================================================
Sub InitComboBox()
			
End Sub

'========================================================================================================
' Name : InitSpreadSheet
' Desc : This method initializes spread sheet column property
'========================================================================================================
Sub InitSpreadSheet()
    frm1.vspddata.OperationMode = 3 
    If frm1.Chk_Bp.checked= True Then
		Call SetZAdoSpreadSheet("a3104RA2","S","A","V20021211",PopupParent.C_SORT_DBAGENT,frm1.vspdData, C_MaxKey, "X","X")
	Else
		Call SetZAdoSpreadSheet("a3104RA1","S","A","V20021211",PopupParent.C_SORT_DBAGENT,frm1.vspdData, C_MaxKey, "X","X")
	End If	
    Call SetSpreadLock() 
End Sub

'========================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================================
Sub SetSpreadLock()
    With frm1
    
    .vspdData.ReDraw = False
	ggoSpread.Source = frm1.vspdData
	ggoSpread.SpreadLockWithOddEvenRowColor()
    .vspdData.ReDraw = True

    End With
End Sub

'========================================================================================================
'========================================================================================================
'                        5.2 Common Method-2
'========================================================================================================
'========================================================================================================

'========================================================================================================
' Name : Form_Load
' Desc : This sub is called from window_OnLoad in Common.vbs automatically
'========================================================================================================
Sub Form_Load()

    Call LoadInfTB19029														
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
    Call ggoOper.LockField(Document, "N")                                   

	Call InitVariables														
	Call SetDefaultVal	
	Call InitSpreadSheet()


	'------ Developer Coding part (End   ) -------------------------------------------------------------- 

End Sub
'========================================================================================================
' Name : Form_QueryUnload
' Desc : This sub is called from window_Unonload in Common.vbs automatically
'========================================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )
 
End Sub

'========================================================================================================
' Name : FncQuery
' Desc : This function is called from MainQuery in Common.vbs
'========================================================================================================
Function FncQuery() 

    FncQuery = False                                            
    
    Err.Clear                                                   

    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")						
    Call InitVariables 											
    ggoSpread.Source = frm1.vspdData
	ggospread.ClearSpreadData		'Buffer Clear
    
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then							
       Exit Function
    End If

	If CompareDateByFormat(frm1.txtFrPrDt.text,frm1.txtToPrDt.text,frm1.txtFrPrDt.Alt,frm1.txtToPrDt.Alt, _
        	               "970025",frm1.txtFrPrDt.UserDefinedFormat,PopupParent.gComDateType, true) = False Then
	   frm1.txtFrPrDt.focus
	   Exit Function
	End If
	
    '-----------------------
    'Query function call area
    '-----------------------

    If DbQuery = False Then Exit Function

    FncQuery = True													

End Function

'========================================================================================================
' Name : FncNew
' Desc : This function is called from MainNew in Common.vbs
'========================================================================================================
Function FncNew()
    Dim IntRetCD 
    
    FncNew = False																 '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Set gActiveElement = document.ActiveElement   
    FncNew = True																 '☜: Processing is OK

End Function
	
'========================================================================================================
' Name : FncDelete
' Desc : This function is called from MainDelete in Common.vbs
'========================================================================================================
Function FncDelete()
    Dim intRetCD
    
    FncDelete = False                                                            '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Set gActiveElement = document.ActiveElement       
    FncDelete = True                                                             '☜: Processing is OK

End Function


'========================================================================================================
' Name : FncSave
' Desc : This function is called from MainSave in Common.vbs
'========================================================================================================
Function FncSave()
    Dim IntRetCD 
    
    FncSave = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status    
   
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 

	'------ Developer Coding part (End )   -------------------------------------------------------------- 

    Set gActiveElement = document.ActiveElement   
    FncSave = True                                                               '☜: Processing is OK

End Function

'========================================================================================================
' Name : FncCopy
' Desc : This function is called from MainCopy in Common.vbs
' Keep : Make sure to clear primary key area
'========================================================================================================
Function FncCopy()
	Dim IntRetCD

    FncCopy = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
	
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 

	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Set gActiveElement = document.ActiveElement   
    FncCopy = True                                                               '☜: Processing is OK

End Function

'========================================================================================================
' Name : FncCancel
' Desc : This function is called by MainCancel in Common.vbs
'========================================================================================================
Function FncCancel() 
    FncCancel = False                                                            '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Set gActiveElement = document.ActiveElement   
    FncCancel = False                                                            '☜: Processing is OK

End Function

'========================================================================================================
' Name : FncInsertRow
' Desc : This function is called by MainInsertRow in Common.vbs
'========================================================================================================
Function FncInsertRow()
    FncInsertRow = False                                                         '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Set gActiveElement = document.ActiveElement   
    FncInsertRow = True  
End Function

'========================================================================================================
' Name : FncDeleteRow
' Desc : This function is called by MainDeleteRow in Common.vbs
'========================================================================================================
Function FncDeleteRow()
    Dim lDelRows
    FncDeleteRow = False                                                         '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Set gActiveElement = document.ActiveElement   
    FncDeleteRow = True  
End Function

'========================================================================================================
' Name : FncPrint
' Desc : This function is called by MainPrint in Common.vbs
'========================================================================================================
Function FncPrint()
    FncPrint = False                                                             '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    
    '--------- Developer Coding Part (End) ------------------------------------------------------------
	Call Parent.FncPrint()                                                       '☜: Protect system from crashing
    FncPrint = True                                                              '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncPrev
' Desc : This function is called by MainPrev in Common.vbs
'========================================================================================================
Function FncPrev() 
    FncPrev = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    
    '--------- Developer Coding Part (End) ------------------------------------------------------------
    Set gActiveElement = document.ActiveElement   	
    FncPrev = True 
End Function

'========================================================================================================
' Name : FncNext
' Desc : This function is called by MainPrev in Common.vbs
'========================================================================================================
Function FncNext() 
    FncNext = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    
    '--------- Developer Coding Part (End) ------------------------------------------------------------
    Set gActiveElement = document.ActiveElement   	
    FncNext = True 
End Function

'========================================================================================================
' Name : FncExcel
' Desc : This function is called by MainExcel in Common.vbs
'========================================================================================================
Function FncExcel() 
    FncExcel = False                                                             '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

	Call Parent.FncExport(PopupParent.C_MULTI)

    FncExcel = True                                                              '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncFind
' Desc : This function is called by MainFind in Common.vbs
'========================================================================================================
Function FncFind() 
    FncFind = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

	Call Parent.FncFind(PopupParent.C_MULTI, True)

    FncFind = True                                                               '☜: Processing is OK
End Function

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

'========================================================================================================
' Name : FncExit
' Desc : This function is called by MainExit in Common.vbs
'========================================================================================================
Function FncExit()
	Dim IntRetCD

    FncExit = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900016", PopupParent.VB_YES_NO,"x","x")
		If IntRetCD = vbNo Then
		    Exit Function
		End If
    End If
    '--------- Developer Coding Part (End) ------------------------------------------------------------

    FncExit = True                                                               '☜: Processing is OK
End Function

'========================================================================================================
'========================================================================================================
'                        5.3 Common Method-3
'========================================================================================================
'========================================================================================================

'========================================================================================================
' Name : DbQuery
' Desc : This sub is called by FncQuery
'========================================================================================================
Function DbQuery() 
	Dim strVal

    Err.Clear                                                       
    DbQuery = False
    
	Call LayerShowHide(1)
    
    With frm1

        strVal = BIZ_PGM_ID
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
        If lgIntFlgMode  <> PopupParent.OPMD_UMODE Then   ' This means that it is first search
			strVal = strVal & "?txtFrPrDt="		& Trim(.txtFrPrDt.text)
			strVal = strVal & "&txtToPrDt="		& Trim(.txtToPrDt.text)
			strVal = strVal & "&txtFrPrNo="		& Trim(.txtFrPrNo.value)
			strVal = strVal & "&txtToPrNo="		& Trim(.txtToPrNo.value)    	
			strVal = strVal & "&txtBPCd="		& Trim(.txtBPCd.value)  
			strVal = strVal & "&txtBPCd_alt="	& Trim(.txtBPCd.alt)	
			strVal = strVal & "&txtRcptType="	& Trim(.txtRcptType.value)
			strVal = strVal & "&txtRcptType_alt="	& Trim(.txtRcptType.alt)
			strVal = strVal & "&txtChk_Bp="		& Trim(.Chk_Bp.checked)		 
        Else
			strVal = strVal & "?txtFrPrDt="		& Trim(.htxtFrPrDt.value)
			strVal = strVal & "&txtToPrDt="		& Trim(.htxtToPrDt.value)
			strVal = strVal & "&txtFrPrNo="		& Trim(.htxtFrPrNo.value)
			strVal = strVal & "&txtToPrNo="		& Trim(.htxtToPrNo.value)    	
			strVal = strVal & "&txtBPCd="		& Trim(.htxtBPCd.value)   
			strVal = strVal & "&txtBPCd_alt="	& Trim(.txtBPCd.alt)
			strVal = strVal & "&txtRcptType="	& Trim(.htxtRcptType.value)	
			strVal = strVal & "&txtRcptType_alt="	& Trim(.txtRcptType.alt)			
			strVal = strVal & "&txtChk_Bp="		& Trim(.Chk_Bp.checked)				 
        End If   

    '--------- Developer Coding Part (End) ------------------------------------------------------------
        strVal = strVal & "&lgPageNo="       & lgPageNo         
        strVal = strVal & "&lgMaxCount="     & C_SHEETMAXROWS_D
        strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")
         
        strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
		strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))

		' 권한관리 추가 
		strVal = strVal & "&lgAuthBizAreaCd="	& lgAuthBizAreaCd			' 사업장 
		strVal = strVal & "&lgInternalCd="		& lgInternalCd				' 내부부서 
		strVal = strVal & "&lgSubInternalCd="	& lgSubInternalCd			' 내부부서(하위포함)
		strVal = strVal & "&lgAuthUsrID="		& lgAuthUsrID				' 개인 
			
        Call RunMyBizASP(MyBizASP, strVal)							
        
    End With
    
    DbQuery = True

End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================
Function DbQueryOk()												

	lgBlnFlgChgValue = False
    lgIntFlgMode     = PopupParent.OPMD_UMODE												'⊙: Indicates that current mode is Update mode
	
	If frm1.vspdData.maxrows > 0 Then
		frm1.vspdData.focus
	End If
	
End Function

'========================================================================================================
'========================================================================================================
'                        5.4 User-defined Method
'========================================================================================================
'========================================================================================================

'===========================================================================
' Function Name : OpenSortPopup
' Function Desc : OpenSortPopup Reference Popup
'===========================================================================
Function OpenSortPopup()

Dim arrRet
	
	If lgIsOpenPop = True Then Exit Function
	lgIsOpenPop = True

	arrRet = window.showModalDialog("../../ComAsp/ZADOSortPopup.asp",Array(ggoSpread.GetXMLData("A"),gMethodText),"dialogWidth=" & Popupparent.SORTW_WIDTH & "px; dialogHeight=" & Popupparent.SORTW_HEIGHT & "px;; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False
	
	If arrRet(0) = "X" Then
	   Exit Function
	Else
	   Call ggoSpread.SaveXMLData("A",arrRet(0),arrRet(1))
       Call InitVariables
       Call InitSpreadSheet()       
   End If
   
End Function

'========================================================================================================
'========================================================================================================
'                        5.5 Tag Event
'========================================================================================================
'========================================================================================================

'========================================================================================================
' Function Name : vspdData_MouseDown
' Function Desc : 
'========================================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)

   If Button = 2 And gMouseClickStatus = "SPC" Then
      gMouseClickStatus = "SPCR"
   End If
End Sub    

'========================================================================================================
'   Event Name : vspdData_GotFocus
'   Event Desc : This event is spread sheet data changed
'========================================================================================================
Sub vspdData_GotFocus()
    ggoSpread.Source = frm1.vspdData
End Sub

'========================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'========================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
    gMouseClickStatus = "SPC"   
    
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
    
    If frm1.vspdData.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
	
    Call SetSpreadColumnValue("A",frm1.vspdData,Col,Row)		    
End Sub

Function vspdData_KeyPress(KeyAscii)
    If KeyAscii = 13 And Frm1.vspdData.ActiveRow > 0 Then
        Call OKClick()
    ElseIf KeyAscii = 27 Then
        Call CancelClick()
    End If
End Function

'========================================================================================================
'   Event Name : vspdData_DblClick
'   Event Desc :
'========================================================================================================
Function vspdData_DblClick(ByVal Col, ByVal Row)
	If Row = 0 Then              		' Title cell을 dblclick했거나....
		Exit Function
	End If
	If Frm1.vspdData.MaxRows = 0 Then  	'NO Data
		Exit Function
	End If
	Call OKClick
End Function

'========================================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If

	If CheckRunningBizProcess = True Then
	   Exit Sub
	End If
    
	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	    
    	If lgPageNo <> "" Then								
           If DbQuery = False Then
              Exit Sub
           End if
    	End If
    End If
    
End Sub

'==========================================================================================
'   Event Name : txtFrPrDt
'   Event Desc :
'==========================================================================================

Sub txtFrPrDt_DblClick(Button)
	if Button = 1 then
		frm1.txtFrPrDt.Action = 7
		Call SetFocusToDocument("P")
		Frm1.txtFrPrDt.Focus
		
	End if
End Sub

Sub txtFrPrDt_KeyPress(KeyAscii)
	On Error Resume Next
	If KeyAscii = 27 Then 
		Call CancelClick()
	ElseIf KeyAscii = 13 Then 
		Call Fncquery()
	End IF
End Sub

'==========================================================================================
'   Event Name : txtToPrDt
'   Event Desc :
'==========================================================================================

Sub txtToPrDt_DblClick(Button)
	if Button = 1 then
		frm1.txtToPrDt.Action = 7
		Call SetFocusToDocument("P")
		Frm1.txtToPrDt.Focus
		
	End if
End Sub

Sub  txtToPrDt_KeyPress(KeyAscii)
	On Error Resume Next
	If KeyAscii = 27 Then 
		Call CancelClick()
	ElseIf KeyAscii = 13 Then 
		Call Fncquery()
	End IF
End Sub

Function NoBp_onchange()
	If frm1.Chk_Bp.checked= True Then
		Call ggoOper.SetReqAttr(frm1.txtBPCd,   "Q")
		frm1.txtBPCd.value = ""
		frm1.txtBPNm.value = ""
	Else
		Call ggoOper.SetReqAttr(frm1.txtBPCd,   "D")
	End If

	Call InitSpreadSheet()

End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>

<!--
'########################################################################################################
'#						6. TAG 부																		#
'######################################################################################################## 
-->
<BODY TABINDEX="-1" SCROLL="no">

<FORM NAME=frm1 TARGET="MyBizASP" METHOD="post">
<TABLE <%=LR_SPACE_TYPE_20%>>
	<TR>
		<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
	</TR>
	<TR>
		<TD HEIGHT=20>
			<FIELDSET CLASS="CLSFLD">
				<TABLE <%=LR_SPACE_TYPE_40%>>
					<TR>				
						<TD CLASS=TD5 NOWRAP>입금일자</TD>
						<TD CLASS=TD6 NOWRAP>
							<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT id=fpDateTime1 title=FPDATETIME CLASS=FPDTYYYYMMDD name=txtFrPrDt CLASSID=<%=gCLSIDFPDT%> ALT="입금일자" tag="12"> </OBJECT>');</SCRIPT>&nbsp;~&nbsp;
							<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT id=fpDateTime2 title=FPDATETIME CLASS=FPDTYYYYMMDD name=txtToPrDt CLASSID=<%=gCLSIDFPDT%> ALT="입금일자" tag="12"> </OBJECT>');</SCRIPT>
						</TD>												
						<TD CLASS=TD5 NOWRAP>입금번호</TD>				
						<TD CLASS=TD6 NOWRAP>
						<INPUT TYPE="Text" NAME="txtFrPrNo" SIZE=15 MAXLENGTH=20 STYLE="TEXT-ALIGN: left" tag="1XXXXU" ALT="입금번호">&nbsp;~&nbsp;
						<INPUT TYPE="Text" NAME="txtToPrNo" SIZE=15 MAXLENGTH=20 STYLE="TEXT-ALIGN: left" tag="1XXXXU" ALT="입금번호">
					</TR>			
					<TR>
						<TD CLASS=TD5 NOWRAP >수금처코드</TD>
						<TD CLASS=TD6 NOWRAP >
						<INPUT NAME="txtBPCd" ALT="수금처코드"  MAXLENGTH="10" SIZE=10 STYLE="TEXT-ALIGN: left" tag  ="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCostCd" align=top TYPE="BUTTON" ONCLICK="vbscript:CALL OpenBp(frm1.txtBPCd.Value, 1)">&nbsp;
						<INPUT NAME="txtBPNm" ALT="수금처명"  MAXLENGTH="20" SIZE=20 STYLE="TEXT-ALIGN: left" tag="14X"></TD>
						<TD CLASS="TD5" NOWRAP>가수금유형</TD>
						<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtRcptType" SIZE=10 MAXLENGTH=20  tag="11XXXU" ALT="가수금유형"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnRcptType" align=top TYPE="BUTTON"ONCLICK="vbscript:Call OpenPopup(frm1.txtRcptType.value,2)">&nbsp;
												<INPUT TYPE=TEXT NAME="txtRcptTypeNm" SIZE=25 tag="14X" ALT="가수금유형명"></TD>
								
					</TR>
					<TR>
						<TD CLASS=TD5 NOWRAP>&nbsp;</TD>
						<TD CLASS="TD6" NOWRAP COLSPAN=3><INPUT type="checkbox" CLASS="STYLE CHECK"  NAME=Chk_Bp ID=Chk_Bp tag="1" onclick=NoBp_onchange()><LABEL FOR=Chk_Bp>수금처가 지정되지 않은 Data조회</LABEL>&nbsp;
					</TR>
				</TABLE>
			</FIELDSET>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
	</TR>
	<TR HEIGHT=100%>
		<TD WIDTH=100% HEIGHT=* VALIGN=TOP>
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR HEIGHT=100%>
					<TD WIDTH=100%>
						<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% id=vspdData tag="2"> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"><PARAM NAME="ReDraw" VALUE="0"> <PARAM NAME="FontSize" VALUE="10"> </OBJECT>');</SCRIPT>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</TR>
	<TR HEIGHT=20>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_30%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD>
						<IMG SRC="../../../CShared/image/query_d.gif"  Style="CURSOR: hand" ALT="Search" NAME="Search" ONCLICK="Call FncQuery()"></IMG>
						&nbsp;<IMG SRC="../../../CShared/image/zpConfig_d.gif" Style="CURSOR: hand" ALT="Config" NAME="Config" OnClick="OpenSortPopup()"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/zpConfig.gif',1)" ></IMG>
					
					</TD>
					<TD ALIGN=RIGHT>
						<IMG SRC="../../../CShared/image/ok_d.gif" Style="CURSOR: hand" ALT="OK" NAME="pop1" ONCLICK="OkClick()" ></IMG>&nbsp
						<IMG SRC="../../../CShared/image/cancel_d.gif" Style="CURSOR: hand" ALT="CANCEL" NAME="pop2" ONCLICK="CancelClick()" ></IMG>
					</TD>				
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" WIDTH=100% HEIGHT= <%=BizSize%> FRAMEBORDER=0 SCROLLING=No noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="htxtFrPrDt"    tag="24">
<INPUT TYPE=HIDDEN NAME="htxtToPrDt"	tag="24">
<INPUT TYPE=HIDDEN NAME="htxtFrPrNo"	tag="24">
<INPUT TYPE=HIDDEN NAME="htxtToPrNo"    tag="24">
<INPUT TYPE=HIDDEN NAME="htxtBPCd"      tag="24">
<INPUT TYPE=HIDDEN NAME="htxtRcptType"  tag="24">
</FORM>
    <DIV ID="MousePT" NAME="MousePT">
    <iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
    </DIV>
</BODY>
</HTML>




