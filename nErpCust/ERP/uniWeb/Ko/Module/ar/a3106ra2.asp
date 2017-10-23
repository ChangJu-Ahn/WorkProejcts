
<%@ LANGUAGE="VBSCRIPT" %>
<!--
'********************************************************************************************************
'*  1. Module Name          : Basis Architect															*
'*  2. Function Name        : Reference Popup Business Part												*
'*  3. Program ID           : 																			*
'*  4. Program Name         :																			*
'*  5. Program Desc         : Reference Popup															*
'*  7. Modified date(First) : 2000/03/29																*
'*  8. Modified date(Last)  : 2002/03/11																*
'*  9. Modifier (First)     : Kang Tae Bum																*
'* 10. Modifier (Last)      : Heo Chunk Ku																*
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"									*
'*                            this mark(��) Means that "may  change"									*
'*                            this mark(��) Means that "must change"									*
'* 13. History              :																			*
'*                            																			*
'********************************************************************************************************
 -->
<HTML>
<HEAD>
<TITLE>ä�ǹ߻�����</TITLE>

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
Option Explicit                                            '��: indicates that All variables must be declared in advance
	

'########################################################################################################
'#                       4.  Data Declaration Part
'########################################################################################################

'========================================================================================================
'=                       4.1 External ASP File
'========================================================================================================
Const C_BizCd = 9
Const C_BizNm = 10

Const BIZ_PGM_ID 		= "a3106rb2.asp"                              '��: Biz Logic ASP Name

'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
Const C_SHEETMAXROWS_D  = 30                                          '��: Server���� �ѹ��� fetch�� �ִ� ����Ÿ �Ǽ� 
Const C_MaxKey          = 15					                      '��: SpreadSheet�� Ű�� ���� 

'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim lgIsOpenPop                                          

Dim lgPopUpR                                              

Dim IsOpenPop  

Dim  arrReturn
Dim  arrParent
Dim  arrParam					

 '------ Set Parameters from Parent ASP ------ 
	arrParent = window.dialogArguments
	Set popupparent = arrParent(0)
	arrParam = arrParent(1)
		
	top.document.title = "ä�ǹ߻�����"

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

	Redim arrReturn(0,0)

    lgStrPrevKey     = ""
    lgPageNo         = ""
    lgIntFlgMode     = PopupParent.OPMD_CMODE                          'Indicates that current mode is Create mode
	lgBlnFlgChgValue = False
    lgSortKey        = 1

	Self.Returnvalue = arrReturn

End Sub

'========================================================================================================
' Name : SetDefaultVal()	
' Desc : Set default value
'========================================================================================================
Sub SetDefaultVal()

	Dim lsMode 
	Dim strSvrDate
	Dim frDt, toDt
	strSvrDate = "<%=GetSvrDate%>"
	frDt = UNIDateAdd("M", -1, strSvrDate,PopupParent.gServerDateFormat)
	
	txtBpCd.value		= arrParam(0)
	txtBpNm.value		= arrParam(1)
	txtDocCur.value		= arrParam(2)	
	lsMode				= arrParam(3)	
	txtBizCd.value		= arrParam(4)
	txtBizNm.value		= arrParam(5)	
	htxtAllcDt.value	= arrParam(6) 
    htxtAllcAlt.value	= arrParam(7) 	
    
	' SetReqAttr(Object, Option) ; N : Required, Q : Protect, D : Default
	If txtBpCd.value <> "" Then				
		Call ggoOper.SetReqAttr(txtBpCd,   "Q")		
	Else		
		Call ggoOper.SetReqAttr(txtBpCd,   "N")		
	End If
	
	If  txtDocCur.value <> "" Then				
		Call ggoOper.SetReqAttr(txtDocCur,   "Q")		
	Else		
		Call ggoOper.SetReqAttr(txtDocCur,   "N")		
	End If	
	
	If  txtBizCd.value <> "" Then				
		Call ggoOper.SetReqAttr(txtBizCd,   "Q")		
	Else	
		IF lsMode = "Q" Then
			Call ggoOper.SetReqAttr(txtBizCd,   "N")		
		Else	
			Call ggoOper.SetReqAttr(txtBizCd,   "D")		
		END IF	
	End If	
	
End Sub

 '******************************************  2.4 POP-UP ó���Լ�  ****************************************
'	���: ���� POP-UP
'   ����: ���� POP-UP�� ���� Open�� Include�Ѵ�. 
'	      �ϳ��� ASP���� Popup�� �ߺ��Ǹ� �ϳ��� ��ũ���� ����ϰ� �������� �������Ͽ� ����Ѵ�.
'********************************************************************************************************* 
'======================================================================================================
'   Event Name : OpenCurrencyInfo
'   Event Desc : 
'=======================================================================================================
Function  OpenCurrencyInfo()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function
	If txtDocCur.className = "protected" Then Exit Function
	
	IsOpenPop = True

	arrParam(0) = "�ŷ���ȭ�˾�"					' �˾� ��Ī 
	arrParam(1) = "b_currency"							' TABLE ��Ī 
	arrParam(2) = Trim(txtDocCur.value)							 	    ' Code Condition
	arrParam(3) = ""									' Name Cindition
	arrParam(4) = ""									' Where Condition
	arrParam(5) = "�ŷ���ȭ" 			
	
    arrField(0) = "CURRENCY"							' Field��(0)
    arrField(1) = "CURRENCY_DESC"						' Field��(1)
    
    
    arrHeader(0) = "�ŷ���ȭ"						' Header��(0)
    arrHeader(1) = "�ŷ���ȭ��"						' Header��(1)
    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
	    txtDocCur.focus
		Exit Function
	Else
		Call SetCurrencyInfo(arrRet)
	End If	

End Function

'======================================================================================================
'   Event Name : SetCurrencyInfo
'   Event Desc : 
'=======================================================================================================
Function SetCurrencyInfo(Byval arrRet)
	
		txtDocCur.value = arrRet(0)
		txtDocCur.focus
	
End Function
'------------------------------------------  OpenDept()  ---------------------------------------
'	Name : OpenBp()
'	Description : OpenAp Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function OpenBp(Byval strCode, byval iWhere)
	Dim arrRet
	Dim arrParam(5)

	If IsOpenPop = True Then Exit Function
	
	If iWhere = 1 Then
		if UCase(txtBpCd.className) = "PROTECTED" Then Exit Function
	End if
	
	IsOpenPop = True

	arrParam(0) = strCode								'  Code Condition
   	arrParam(1) = "A_OPEN_AR"							' ä�ǰ� ����(�ŷ�ó ����)
	arrParam(2) = txtARDt.Text								'FrDt
	arrParam(3) = txtToARDt.Text									'ToDt
	arrParam(4) = "B"							'B :���� S: ���� T: ��ü 
	Select Case iWhere
		Case 1
			arrParam(5) = "PAYER"									'SUP :����ó PAYTO: ����ó SOL:�ֹ�ó PAYER :����ó INV:���ݰ�� 	
		Case 2
			arrParam(5) = "SOL"									'SUP :����ó PAYTO: ����ó SOL:�ֹ�ó PAYER :����ó INV:���ݰ�� 	
	End Select	
	arrRet = window.showModalDialog("../../Module/ar/BpPopup.asp", Array(window.PopupParent,arrParam), _
		"dialogWidth=780px; dialogHeight=450px; : Yes; help: No; resizable: No; status: No;")
		
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Call EscBpCdPopup(iWhere)
		Exit Function
	Else
		Call SetBpCd(arrRet,iWhere)
	End If	
End Function
 '------------------------------------------  OpenBpCd()  -------------------------------------------------
'	Name : OpenBpCd()
'	Description : Bp PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenBpCd(ByVal BpPos)

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function
	If txtBpCd.className = "protected" Then Exit Function
	
	IsOpenPop = True

	Select Case BpPos
		Case 1
			arrParam(0) = "����ó�˾�"
			arrParam(1) = "(SELECT DISTINCT A.BP_CD,A.BP_NM FROM B_BIZ_PARTNER A, A_OPEN_AR B " 
			arrParam(1) = arrParam(1) & "WHERE  A.BP_CD=B.PAY_BP_CD AND B.CONF_FG = " & FilterVar("C", "''", "S") & "  AND B.AR_STS=" & FilterVar("O", "''", "S") & "  AND B.BAL_AMT <> 0" 
			IF txtARDt.Text<>"" THEN 	arrParam(1) = arrParam(1) & " AND AR_DT >= " & FilterVar(UNIConvDate(txtARDt.Text), "''", "S") & ""
			IF txtToARDt.Text<>"" THEN arrParam(1) = arrParam(1) & " AND AR_DT <= " & FilterVar(UNIConvDate(txtToARDt.Text), "''", "S") & ""
			arrParam(1) = arrParam(1) & ") TMP"
			
			arrParam(2) = Trim(txtBpCd.Value)
			arrParam(3) = ""
			arrParam(4) = ""
			arrParam(5) = "����ó"			
	
			arrField(0) = "TMP.BP_CD"	
			arrField(1) = "TMP.BP_NM"	
	   
			arrHeader(0) = "����ó"		
			arrHeader(1) = "����ó��"	
	   Case 2
			arrParam(0) = "�ֹ�ó�˾�"
			arrParam(1) = "(SELECT DISTINCT A.BP_CD,A.BP_NM FROM B_BIZ_PARTNER A, A_OPEN_AR B " 
			arrParam(1) = arrParam(1) & "WHERE  A.BP_CD=B.PAY_BP_CD AND B.CONF_FG = " & FilterVar("C", "''", "S") & "  AND B.AR_STS=" & FilterVar("O", "''", "S") & "  AND B.BAL_AMT <> 0" 
			IF txtARDt.Text<>"" THEN 	arrParam(1) = arrParam(1) & " AND AR_DT >= " & FilterVar(UNIConvDate(txtARDt.Text), "''", "S") & ""
			IF txtToARDt.Text<>"" THEN arrParam(1) = arrParam(1) & " AND AR_DT <= " & FilterVar(UNIConvDate(txtToARDt.Text), "''", "S") & ""
			arrParam(1) = arrParam(1) & ") TMP"
						
			arrParam(2) = Trim(txtDealBpCd.Value)
			arrParam(3) = ""
			arrParam(4) = ""
			arrParam(5) = "�ֹ�ó"			
	
			arrField(0) = "TMP.BP_CD"	
			arrField(1) = "TMP.BP_NM"	
	   
			arrHeader(0) = "�ֹ�ó"		
			arrHeader(1) = "�ֹ�ó��"	
   End Select
   
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	IF 	arrRet(0) <> "" then			
		Call SetBpCd(arrRet,BpPos)
	Else 
		Call EscBpCdPopup(BpPos)
		Exit Function
	end if
End Function
  '------------------------------------------  EscBpCdPopup()  --------------------------------------------------
'	Name : EscBpCdPopup()
'	Description : Item Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function EscBpCdPopup(ByVal BpPos)'
	
	If BpPos = 1 Then
		txtBpCd.focus
	Else
		txtDealBpCd.focus
	End If
				
	lgBlnFlgChgValue = True
	
End Function
'------------------------------------------  SetBpCd()  --------------------------------------------------
'	Name : SetBpCd()
'	Description : Item Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetBpCd(Byval arrRet,ByVal BpPos)
	
	If BpPos = 1 Then
		txtBpCd.value = arrRet(0)		
		txtBpNm.value = arrRet(1)
		txtBpCd.focus
	Else
		txtDealBpCd.value = arrRet(0)
		txtDealBpNm.value = arrRet(1)
		txtDealBpCd.focus
	End If
				
	lgBlnFlgChgValue = True
		
End Function

'------------------------------------------  OpenBizCd()  -------------------------------------------------
'	Name : OpenBizCd()
'	Description : Cost PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenBizCd()

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function
	If txtBizCd.className = "protected" Then Exit Function
	
	IsOpenPop = True

	arrParam(0) = "������˾�"					' �˾� ��Ī 
	arrParam(1) = "B_BIZ_AREA"						' TABLE ��Ī 
	arrParam(2) = Trim(txtBizCd.Value)			' Code Condition
	arrParam(3) = ""								' Name Cindition
	arrParam(4) = ""								' Where Condition
	arrParam(5) = "�����"			
	
    arrField(0) = "BIZ_AREA_CD"						' Field��(0)
    arrField(1) = "BIZ_AREA_NM"						' Field��(1)
    
    arrHeader(0) = "�����"						' Header��(0)
    arrHeader(1) = "������"					' Header��(1)
    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	IF 	arrRet(0) <> "" then		
		Call SetBizCd(arrRet)
	Else
		txtBizCd.focus
		Exit Function
	end if
	
End Function

'++++++++++++++++  Insert Your Code for PopUp(Open)  ++++++++++++++++++++++++++++++++++++++++++++++++++ 

'------------------------------------------  SetBizCd()  --------------------------------------------------
'	Name : SetBizCd()
'	Description : Item Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetBizCd(Byval arrRet)
	
		txtBizCd.value = arrRet(0)		
		txtBizNm.value = arrRet(1)
		txtBizCd.focus
	lgBlnFlgChgValue = True				
	
End Function

'========================================  2.3 LoadInfTB19029()  =========================================
' Name : LoadInfTB19029()	
' Desc : Set System Number format
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "A","NOCOOKIE","RA") %>                                '��: 
	<% Call LoadBNumericFormatA("I", "A", "NOCOOKIE", "RA") %>
End Sub


'===========================================  2.3.1 OkClick()  ==========================================
'=	Name : OkClick()																					=
'=	Description : Return Array to Opener Window when OK button click									=
'=				  �� �κп��� �÷� �߰��ϰ� ����Ÿ ������ �Ͼ�� �մϴ�.   							=
'========================================================================================================
Function OKClick()
	Dim ii ,jj ,kk
	
	if vspdData.SelModeSelCount > 0 Then 			
		Redim arrReturn(vspdData.SelModeSelCount - 1,C_MaxKey)
		kk = 0
		For ii = 0 To vspdData.MaxRows - 1
			vspdData.Row = ii + 1			
			If vspdData.SelModeSelected Then
				For jj = 0 To C_MaxKey - 1
					vspdData.Col	 = GetKeyPos("A",jj + 1)		
					arrReturn(kk,jj) = vspdData.Text
				Next			
				arrReturn(kk,C_MaxKey) = txtDocCur.value
				kk = kk + 1
			End If
		Next	
	End If			
	
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
'	Description : JUMP�� Loadȭ������ ���Ǻη� Value
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
    vspddata.OperationMode = 5
    Call SetZAdoSpreadSheet("a3106RA2","S","A","V20030215",PopupParent.C_SORT_DBAGENT,vspdData, C_MaxKey, "X","X")
    Call SetSpreadLock() 
End Sub

'========================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================================
Sub SetSpreadLock()
	vspdData.ReDraw = False
	ggoSpread.Source = vspdData
	ggoSpread.SpreadLockWithOddEvenRowColor()
	vspdData.ReDraw = True
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
    Call LoadInfTB19029()														
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
    Call ggoOper.LockField(Document, "N")                                   
    
	Call SetDefaultVal()						'1
	Call InitVariables()						'2		//logic�� 1->2������ ó���Ǿ�� ��.				
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
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then							
       Exit Function
    End If

	If Not ChkQueryDate Then
		Exit Function
    End If
    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")						
    Call InitVariables() 											
	ggoSpread.Source = vspdData
	ggoSpread.ClearSpreadData
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
    
    FncNew = False																 '��: Processing is NG
    Err.Clear                                                                    '��: Clear err status
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Set gActiveElement = document.ActiveElement   
    FncNew = True																 '��: Processing is OK
End Function
	
'========================================================================================================
' Name : FncDelete
' Desc : This function is called from MainDelete in Common.vbs
'========================================================================================================
Function FncDelete()
    Dim intRetCD
    
    FncDelete = False                                                            '��: Processing is NG
    Err.Clear                                                                    '��: Clear err status
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Set gActiveElement = document.ActiveElement       
    FncDelete = True                                                             '��: Processing is OK
End Function


'========================================================================================================
' Name : FncSave
' Desc : This function is called from MainSave in Common.vbs
'========================================================================================================
Function FncSave()
    Dim IntRetCD 
    
    FncSave = False                                                              '��: Processing is NG
    Err.Clear                                                                    '��: Clear err status    
   
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 

	'------ Developer Coding part (End )   -------------------------------------------------------------- 

    Set gActiveElement = document.ActiveElement   
    FncSave = True                                                               '��: Processing is OK
End Function

'========================================================================================================
' Name : FncCopy
' Desc : This function is called from MainCopy in Common.vbs
' Keep : Make sure to clear primary key area
'========================================================================================================
Function FncCopy()
	Dim IntRetCD

    FncCopy = False                                                              '��: Processing is NG
    Err.Clear                                                                    '��: Clear err status
	
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 

	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Set gActiveElement = document.ActiveElement   
    FncCopy = True                                                               '��: Processing is OK
End Function

'========================================================================================================
' Name : FncCancel
' Desc : This function is called by MainCancel in Common.vbs
'========================================================================================================
Function FncCancel() 
    FncCancel = False                                                            '��: Processing is NG
    Err.Clear                                                                    '��: Clear err status

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Set gActiveElement = document.ActiveElement   
    FncCancel = False                                                            '��: Processing is OK
End Function

'========================================================================================================
' Name : FncInsertRow
' Desc : This function is called by MainInsertRow in Common.vbs
'========================================================================================================
Function FncInsertRow()
    FncInsertRow = False                                                         '��: Processing is NG
    Err.Clear                                                                    '��: Clear err status

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Set gActiveElement = document.ActiveElement   
    FncInsertRow = True                                                          '��: Processing is OK
End Function

'========================================================================================================
' Name : FncDeleteRow
' Desc : This function is called by MainDeleteRow in Common.vbs
'========================================================================================================
Function FncDeleteRow()
    Dim lDelRows
    FncDeleteRow = False                                                         '��: Processing is NG
    Err.Clear                                                                    '��: Clear err status

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Set gActiveElement = document.ActiveElement   
    FncDeleteRow = True                                                          '��: Processing is OK
End Function

'========================================================================================================
' Name : FncPrint
' Desc : This function is called by MainPrint in Common.vbs
'========================================================================================================
Function FncPrint()
    FncPrint = False                                                             '��: Processing is NG
    Err.Clear                                                                    '��: Clear err status
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    
    '--------- Developer Coding Part (End) ------------------------------------------------------------
	Call Parent.FncPrint()                                                       '��: Protect system from crashing
    FncPrint = True                                                              '��: Processing is OK
End Function

'========================================================================================================
' Name : FncPrev
' Desc : This function is called by MainPrev in Common.vbs
'========================================================================================================
Function FncPrev() 
    FncPrev = False                                                              '��: Processing is NG
    Err.Clear                                                                    '��: Clear err status
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    
    '--------- Developer Coding Part (End) ------------------------------------------------------------
    Set gActiveElement = document.ActiveElement   	
    FncPrev = True                                                               '��: Processing is OK
End Function

'========================================================================================================
' Name : FncNext
' Desc : This function is called by MainPrev in Common.vbs
'========================================================================================================
Function FncNext() 
    FncNext = False                                                              '��: Processing is NG
    Err.Clear                                                                    '��: Clear err status
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    
    '--------- Developer Coding Part (End) ------------------------------------------------------------
    Set gActiveElement = document.ActiveElement   	
    FncNext = True                                                               '��: Processing is OK
End Function

'========================================================================================================
' Name : FncExcel
' Desc : This function is called by MainExcel in Common.vbs
'========================================================================================================
Function FncExcel() 
    FncExcel = False                                                             '��: Processing is NG
    Err.Clear                                                                    '��: Clear err status

	Call Parent.FncExport(PopupParent.C_MULTI)

    FncExcel = True                                                              '��: Processing is OK
End Function

'========================================================================================================
' Name : FncFind
' Desc : This function is called by MainFind in Common.vbs
'========================================================================================================
Function FncFind() 
    FncFind = False                                                              '��: Processing is NG
    Err.Clear                                                                    '��: Clear err status

	Call Parent.FncFind(PopupParent.C_MULTI, True)

    FncFind = True                                                               '��: Processing is OK
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

    FncExit = False                                                              '��: Processing is NG
    Err.Clear                                                                    '��: Clear err status

    FncExit = True                                                               '��: Processing is OK
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
    
        strVal = BIZ_PGM_ID
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
        If lgIntFlgMode  <> PopupParent.OPMD_UMODE Then   ' This means that it is first search
			strVal = strVal & "?txtBizCd="		& Trim(txtBizCd.value)				'��:     	
			strVal = strVal & "&txtBpCd="		& Trim(txtBpCd.value)					'��: ��ȸ ���� ����Ÿ 
			strVal = strVal & "&txtArDt="		& Trim(txtArDt.text)					'��: ��ȸ ���� ����Ÿ 
			strVal = strVal & "&txtToArDt="		& Trim(txtToArDt.text)				'��: ��ȸ ���� ����Ÿ 
			strVal = strVal & "&txtArDueDt="	& Trim(txtArDueDt.text)					'��: ��ȸ ���� ����Ÿ 
			strVal = strVal & "&txtToArDueDt="	& Trim(txtToArDueDt.text)				'��: ��ȸ ���� ����Ÿ 
			strVal = strVal & "&txtDocCur="		& Trim(txtDocCur.value)				'��: ��ȸ ���� ����Ÿ 
			strVal = strVal & "&txtDealBpCd="	& Trim(txtDealBpCd.value)			'��: ��ȸ ���� ����Ÿ 
			strVal = strVal & "&txtArNo="		& Trim(txtArNo.value)					'��: ��ȸ ���� ����Ÿ 
			strVal = strVal & "&txtRefNo="		& Trim(txtRefNo.value)					'��: ��ȸ ���� ����Ÿ 
			strVal = strVal & "&txtBpCd_Alt="	& Trim(txtBpcd.alt)
			strVal = strVal & "&txtBizCd_Alt="	& Trim(txtBizcd.alt)
			strVal = strVal & "&txtDealBpCd_Alt="	& Trim(txtDealBpcd.alt)
        Else
			strVal = strVal & "?txtBizCd="		& Trim(htxtBizCd.value)				'��:     	
			strVal = strVal & "&txtBpCd="		& Trim(htxtBpCd.value)					'��: ��ȸ ���� ����Ÿ 
			strVal = strVal & "&txtArDt="		& Trim(htxtArDt.value)					'��: ��ȸ ���� ����Ÿ 
			strVal = strVal & "&txtToArDt="		& Trim(htxtToArDt.value)				'��: ��ȸ ���� ����Ÿ 
			strVal = strVal & "&txtArDueDt="	& Trim(htxtArDueDt.value)					'��: ��ȸ ���� ����Ÿ 
			strVal = strVal & "&txtToArDueDt="	& Trim(htxtToArDueDt.value)				'��: ��ȸ ���� ����Ÿ 
			strVal = strVal & "&txtDocCur="		& Trim(htxtDocCur.value)				'��: ��ȸ ���� ����Ÿ 
			strVal = strVal & "&txtDealBpCd="	& Trim(htxtDealBpCd.value)			'��: ��ȸ ���� ����Ÿ 
			strVal = strVal & "&txtArNo="		& Trim(htxtArNo.value)					'��: ��ȸ ���� ����Ÿ 
			strVal = strVal & "&txtRefNo="		& Trim(htxtRefNo.value)					'��: ��ȸ ���� ����Ÿ 
			strVal = strVal & "&txtBpCd_Alt="	& Trim(txtBpcd.alt)
			strVal = strVal & "&txtBizCd_Alt="	& Trim(txtBizcd.alt)
			strVal = strVal & "&txtDealBpCd_Alt="	& Trim(txtDealBpcd.alt)
        End If   
    '--------- Developer Coding Part (End) ------------------------------------------------------------
			strVal = strVal & "&txtAllcDt="	     & Trim(htxtAllcDt.value)
			strVal = strVal & "&lgPageNo="       & lgPageNo         
			strVal = strVal & "&lgMaxCount="     & C_SHEETMAXROWS_D
			strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")
			strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
			strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))

			Call RunMyBizASP(MyBizASP, strVal)							
    
    DbQuery = True

End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================
Function DbQueryOk()												

	lgBlnFlgChgValue = False
    lgIntFlgMode     = PopupParent.OPMD_UMODE												'��: Indicates that current mode is Update mode

	If vspdData.MaxRows > 0 Then
		vspdData.Focus
	End If

End Function


'========================================================================================================
'========================================================================================================
'                        5.4 User-defined Method
'========================================================================================================
'========================================================================================================

'===========================================================================
' Function Name : OpenOrderBy
' Function Desc : OpenOrderBy Reference Popup
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
       Call InitVariables()
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
    ggoSpread.Source = vspdData
End Sub

'========================================================================================================
'   Event Name : vspdData_DblClick
'   Event Desc :
'========================================================================================================
Function vspdData_DblClick(ByVal Col, ByVal Row)
	If vspdData.MaxRows > 0 Then
		If vspdData.ActiveRow = Row Or vspdData.ActiveRow > 0 Then
			Call OkClick()
		End If
	End If
End Function
	
'========================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : �÷��� Ŭ���� ��� �߻� 
'========================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
    gMouseClickStatus = "SPC"   
    
    If Row = 0 Then
        ggoSpread.Source = vspdData
        If lgSortKey = 1 Then
            ggoSpread.SSSort, lgSortKey
            lgSortKey = 2
        Else
            ggoSpread.SSSort, lgSortKey
            lgSortKey = 1
        End If    
    End If
    
	Call SetSpreadColumnValue("A",vspdData,Col,Row)	        
	
    If vspdData.MaxRows = 0 Then                                                    'If there is no data.
		Exit Sub
   	End If
End Sub

'========================================================================================================
'   Event Name : vspdData_KeyPress
'   Event Desc : 
'========================================================================================================
Function vspdData_KeyPress(KeyAscii)
    If KeyAscii = 13 And vspdData.ActiveRow > 0 Then
        Call OKClick()
    ElseIf KeyAscii = 27 Then
        Call CancelClick()
    End If
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
    
	If vspdData.MaxRows < NewTop + VisibleRowCnt(vspdData,NewTop) Then	    
    	If lgPageNo <> "" Then								
           If DbQuery = False Then
              Exit Sub
           End if
    	End If
    End If
End Sub

'=======================================================================================================
'   Event Name : txtArDt_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub  txtArDt_DblClick(Button)
    If Button = 1 Then
        txtArDt.Action = 7                        
        Call SetFocusToDocument("P")
		txtArDt.Focus 
    End If
End Sub
'=======================================================================================================
'   Event Name : txtArDueDt_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub  txtArDueDt_DblClick(Button)
    If Button = 1 Then
        txtArDueDt.Action = 7      
        Call SetFocusToDocument("P")
		txtArDueDt.Focus                           
    End If
End Sub
'=======================================================================================================
'   Event Name : txtArDt_Change()
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub  txtArDt_Change()
    
End Sub

'=======================================================================================================
'   Event Name : txtToArDt_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub  txtToArDt_DblClick(Button)
    If Button = 1 Then
        txtToArDt.Action = 7     
        Call SetFocusToDocument("P")
		txtToArDt.Focus                            
    End If
End Sub

'=======================================================================================================
'   Event Name : txtToArDueDt_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub  txtToArDueDt_DblClick(Button)
    If Button = 1 Then
        txtToArDueDt.Action = 7    
        Call SetFocusToDocument("P")
		txtToArDueDt.Focus                             
    End If
End Sub
'=======================================================================================================
'   Event Name : txtToArDt_Change()
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub  txtToArDt_Change()
    
End Sub

Sub txtArDt_KeyPress(KeyAscii)
	On Error Resume Next
	If KeyAscii = 27 Then 
		Call CancelClick()
	ElseIf KeyAscii = 13 Then 
		Call Fncquery()
	End IF
End Sub

Sub txtToArDt_KeyPress(KeyAscii)
	On Error Resume Next
	If KeyAscii = 27 Then 
		Call CancelClick()
	ElseIf KeyAscii = 13 Then 
		Call Fncquery()
	End IF
End Sub

'========================================================================================================
'   Event Name : txtArDueDt_KeyPress()
'   Event Desc : 
'========================================================================================================
Sub txtArDueDt_KeyPress(KeyAscii)
	On Error Resume Next
	If KeyAscii = 27 Then 
		Call CancelClick()
	ElseIf KeyAscii = 13 Then 
		Call Fncquery()
	End IF
End Sub

Sub txtToArDueDt_KeyPress(KeyAscii)
	On Error Resume Next
	If KeyAscii = 27 Then 
		Call CancelClick()
	ElseIf KeyAscii = 13 Then 
		Call Fncquery()
	End IF
End Sub
Sub txtBpCd_KeyPress(KeyAscii)
	On Error Resume Next
	If KeyAscii = 27 Then 
		Call CancelClick()
	ElseIf KeyAscii = 13 Then 
		Call Fncquery()
	End IF
End Sub

Function DetailConditionClick()
	If DetailCondition.style.display = "none" Then
		DetailCondition.style.display = ""
		Call ggoOper.SetReqAttr(txtBpCd,   "D")
	Else
		DetailCondition.style.display = "none"
		If arrParam(0) <> "" Then
			Call ggoOper.SetReqAttr(txtBpCd,   "Q")
		Else
			Call ggoOper.SetReqAttr(txtBpCd,   "N")
		End If
	End If
End Function

'=======================================================================================================
'   Function Name : ChkQueryDate
'   Function Desc : 
'=======================================================================================================
Function ChkQueryDate()
	chkQueryDate= True
	
	If CompareDateByFormat(txtArDt.text, txtToArDt.text, txtArDt.Alt, txtToArDt.Alt, _
   	           "970025", txtArDt.UserDefinedFormat,PopupParent.gComDateType, true) = False Then
		chkQueryDate= False
		txtArDt.focus
		Exit Function
	End If
	
	If CompareDateByFormat(txtArDueDt.text,txtToArDueDt.text,txtArDueDt.Alt,txtToArDueDt.Alt, _
   	           "970025",txtArDueDt.UserDefinedFormat,PopupParent.gComDateType, true) = False Then
	   chkQueryDate= False
	   txtArDueDt.focus
	   Exit Function
	End If
	
	If CompareDateByFormat(txtArDt.text,htxtAllcDt.Value,txtArDt.Alt,htxtAllcAlt.value, _
   	           "970025",txtArDt.UserDefinedFormat,PopupParent.gComDateType, true) = False Then
	   chkQueryDate= False
	   txtArDt.focus
	   Exit Function
	End If
	
	If CompareDateByFormat(txtToArDt.text,htxtAllcDt.Value,txtToArDt.Alt, htxtAllcAlt.value,_
   	           "970025",txtToArDt.UserDefinedFormat,PopupParent.gComDateType, true) = False Then
	   chkQueryDate= False
	   txtToArDt.focus
	   Exit Function
	End If

End Function


</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>

<!--
'########################################################################################################
'#						6. TAG ��																		#
'######################################################################################################## 
-->
<BODY TABINDEX="-1" SCROLL="no">
<TABLE <%=LR_SPACE_TYPE_20%>>
	<TR>
		<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
	</TR>
	<TR>
		<TD HEIGHT=20>
			<FIELDSET CLASS="CLSFLD">
				<TABLE <%=LR_SPACE_TYPE_40%>>
					<TR>
						<TD CLASS=TD5 NOWRAP>����ó</TD>
						<TD CLASS=TD6 NOWRAP><INPUT TYPE="Text" NAME="txtBpCd" SIZE=10 MAXLENGTH=10 STYLE="TEXT-ALIGN: Left" tag="12NXXU" ALT="����ó"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBizCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:CALL OpenBp(txtBpCd.Value, 1)"> <INPUT TYPE=TEXT NAME="txtBpNm" SIZE=20 MAXLENGTH=20 STYLE="TEXT-ALIGN: Left" tag="14" ALT="����ó��"></TD>
						<TD CLASS=TD5 NOWRAP>�ŷ���ȭ</TD>
						<TD CLASS=TD6 NOWRAP><INPUT NAME="txtDocCur" ALT="�ŷ���ȭ" MAXLENGTH="3" SIZE=10 STYLE="TEXT-ALIGN: Left" tag ="12NXXU"><IMG align=top name=btnCalType onclick="vbscript:OpenCurrencyInfo()" src="../../../CShared/image/btnPopup.gif"  TYPE="BUTTON"></TD>
					</TR>
					<TR>
						<TD CLASS=TD5 NOWRAP>ä������</TD>
						<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/a3106ra2_OBJECT3_txtArDt.js'></script>								
						&nbsp;~&nbsp;<script language =javascript src='./js/a3106ra2_OBJECT4_txtToArDt.js'></script></TD>												
						<TD CLASS=TD5 NOWRAP>��������</TD>
						<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/a3106ra2_OBJECT1_txtArDueDt.js'></script>								
						&nbsp;~&nbsp;<script language =javascript src='./js/a3106ra2_OBJECT2_txtToArDueDt.js'></script></TD>												
					</TR>
					<TR>						
						<TD CLASS=TD5 NOWRAP>�����</TD>
						<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtBizCd" SIZE=10 MAXLENGTH=10 STYLE="TEXT-ALIGN: Left" tag=11NXXU" ALT="�����"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCostCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenBizCd()"> <INPUT TYPE=TEXT NAME="txtBizNm" SIZE=20 MAXLENGTH=20 STYLE="TEXT-ALIGN: left" tag="14" ALT="������"></TD>					
						<TD CLASS=TD5 NOWRAP>������ȣ</TD>
						<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtRefNo" SIZE=20 MAXLENGTH=10 STYLE="TEXT-ALIGN: Left" tag=11NXXU" ALT="������ȣ">
						&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
						<IMG SRC="../../../CShared/image/icon/QualityC.gif" Style="CURSOR: hand" ALT="OK" NAME="pop1" ONCLICK="DetailConditionClick()" ></IMG></TD>
					</TR>
					<TR ID="DetailCondition" style="display: none">
						<TD CLASS=TD5 NOWRAP>�ֹ�ó</TD>
						<TD CLASS=TD6 NOWRAP><INPUT TYPE="Text" NAME="txtDealBpCd" SIZE=10 MAXLENGTH=10 STYLE="TEXT-ALIGN: Left" tag="11NXXU" ALT="�ֹ�ó"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDealBpCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:CALL OpenBp(txtDealBpCd.Value, 2)"> <INPUT TYPE=TEXT NAME="txtDealBpNm" SIZE=20 MAXLENGTH=20 STYLE="TEXT-ALIGN: Left" tag="14" ALT="�ֹ�ó"></TD>
						<TD CLASS=TD5 NOWRAP>ä�ǹ�ȣ</TD>
						<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtArNo" MAXLENGTH=18 STYLE="TEXT-ALIGN: Left" tag=11NXXU" ALT="ä�ǹ�ȣ"></TD>					
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
						<script language =javascript src='./js/a3106ra2_vspdData_vspdData.js'></script>
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
						<IMG SRC="../../../CShared/image/ok_d.gif" Style="CURSOR: hand" ALT="OK" NAME="pop1" ONCLICK="OkClick()" ></IMG>&nbsp;
						<IMG SRC="../../../CShared/image/cancel_d.gif" Style="CURSOR: hand" ALT="CANCEL" NAME="pop2" ONCLICK="CancelClick()" ></IMG>
					</TD>				
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>>
			<IFRAME NAME="MyBizASP"  WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="htxtBizCd"     tag="24">
<INPUT TYPE=HIDDEN NAME="htxtBpCd"		tag="24">
<INPUT TYPE=HIDDEN NAME="htxtArDt"		tag="24">
<INPUT TYPE=HIDDEN NAME="htxtToArDt"    tag="24">
<INPUT TYPE=HIDDEN NAME="htxtArDueDt"		tag="24">
<INPUT TYPE=HIDDEN NAME="htxtToArDueDt"    tag="24">
<INPUT TYPE=HIDDEN NAME="htxtDocCur"	tag="24">
<INPUT TYPE=HIDDEN NAME="htxtDealBpCd"  tag="24">
<INPUT TYPE=HIDDEN NAME="htxtArNo"      tag="24">
<INPUT TYPE=HIDDEN NAME="htxtAllcDt"	tag="14">
<INPUT TYPE=HIDDEN NAME="htxtAllcAlt"      tag="14">
<INPUT TYPE=HIDDEN NAME="htxtRefNo"      tag="14">
    <DIV ID="MousePT" NAME="MousePT">
    <iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
    </DIV>
</BODY>
</HTML>

