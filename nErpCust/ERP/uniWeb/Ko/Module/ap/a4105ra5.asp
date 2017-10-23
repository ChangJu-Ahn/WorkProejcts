
<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : Template
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
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

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
<!--'==========================================  1.1.1 Style Sheet  ===========================================
'============================================================================================================ -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'=====================================  1.1.2 ���� Include   =============================================
'=========================================================================================================== -->
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliPAMain.vbs">			</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliPAEvent.vbs">			</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliPAOperation.vbs">		</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliVariables.vbs">		</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliDBAgentA.vbs">			</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliDBAgentVariables.vbs">	</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliRdsQuery.vbs">			</SCRIPT>
<SCRIPT LANGUAGE = "JavaScript"SRC = "../../inc/incImage.js">				</SCRIPT>

<Script Language="VBScript">
Option Explicit                                                        '��: indicates that All variables must be declared in advance


'########################################################################################################
'#                       4.  Data Declaration Part
'########################################################################################################

'========================================================================================================
'=                       4.1 External ASP File
'========================================================================================================

Const BIZ_PGM_ID 		= "a4105rb5.asp"                              '��: Biz Logic ASP Name

'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
Const C_SHEETMAXROWS_D  = 30                                          '��: Server���� �ѹ��� fetch�� �ִ� ����Ÿ �Ǽ� 
Const C_MaxKey          = 15				                          '��: SpreadSheet�� Ű�� ���� 

'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim  lgIsOpenPop                                          
Dim  lgPopUpR                                              
Dim  lgQueryFlag
Dim  lgCode		
	
Dim  IsOpenPop     
Dim  IsBpPop  
Dim  IsDocPop  
     	
Dim  arrReturn
Dim  arrParent
Dim  arrParam

' ���Ѱ��� �߰� 
Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' ����� 
Dim lgInternalCd, lgDeptCd, lgDeptNm			' ���κμ� 
Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' ���κμ�(��������)
Dim lgAuthUsrID, lgAuthUsrNm					' ���� 

	 '------ Set Parameters from Parent ASP ------ 
arrParent        = window.dialogArguments
Set PopupParent = arrParent(0)	 
arrParam		= arrParent(1)



top.document.title = PopupParent.gActivePRAspName
'	top.document.title = "ä���߻�����"

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

	' ���Ѱ��� �߰� 
	If UBound(arrParam) > 5 Then
		lgAuthBizAreaCd		= arrParam(5)
		lgInternalCd		= arrParam(6)
		lgSubInternalCd		= arrParam(7)
		lgAuthUsrID			= arrParam(8)
	End If	
End Sub

'========================================================================================================
' Name : LoadInfTB19029()	
' Desc : Set System Number format
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "A","NOCOOKIE", "RA") %>                                '��: 
	<% Call LoadBNumericFormatA("I", "A", "NOCOOKIE", "RA") %>
End Sub



'========================================================================================================
' Name : InitComboBox()
' Desc : Set ComboBox
'========================================================================================================
Sub InitComboBox()
			
End Sub

'========================================================================================================
' Name : SetDefaultVal()	
' Desc : Set default value
'========================================================================================================
Sub SetDefaultVal()
	If arrParam(0) <> "" Then
		frm1.txtBpCd.value = arrParam(0)
		frm1.txtBpNm.value = arrParam(1)			
	End If
		
	If arrParam(2) <> "" Then	
		frm1.txtDocCur.value = arrParam(2)				
	End If	
		
	frm1.htxtAllcDt.value	= arrParam(3) 
    frm1.htxtAllcAlt.value	= arrParam(4) 	
    		
	' SetReqAttr(Object, Option) ; N : Required, Q : Protect, D : Default
	If frm1.txtBpCd.value <> "" Then		
		IsBpPop = False
		Call ggoOper.SetReqAttr(frm1.txtBpCd,   "Q")		
	Else
		IsBpPop = True
		Call ggoOper.SetReqAttr(frm1.txtBpCd,   "N")		
	End If
	
	If frm1.txtDocCur.value <> "" Then		
		IsDocPop = False
		Call ggoOper.SetReqAttr(frm1.txtDocCur,   "Q")		
	Else
		IsDocPop = True
		Call ggoOper.SetReqAttr(frm1.txtDocCur,   "N")		
	End If	

	If arrParam(3) =  "Q" Then					
		Call ggoOper.SetReqAttr(frm1.txtBizCd,   "N")		
	Else			
		Call ggoOper.SetReqAttr(frm1.txtBizCd,   "D")		
	End If


	frm1.txtApDt.text	= UNIDateAdd("M", -1, arrParam(3),PopupParent.gDateFormat)
	frm1.txtToApDt.text	= arrParam(3) 
End Sub

'========================================================================================================
' Name : InitSpreadSheet
' Desc : This method initializes spread sheet column property
'========================================================================================================
Sub InitSpreadSheet()
	frm1.vspdData.OperationMode = 5
    Call SetZAdoSpreadSheet("a4105ra5","S","A","V20021211",PopupParent.C_SORT_DBAGENT,frm1.vspdData, C_MaxKey, "X","X")
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

'++++++++++++++++++++++++++++++++++++++++++  2.3 ������ ���� �Լ�  ++++++++++++++++++++++++++++++++++++++
'+	������ ���� Function, Procedure																		+
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++


'===========================================  2.3.1 OkClick()  ==========================================
'=	Name : OkClick()																					=
'=	Description : Return Array to Opener Window when OK button click									=
'=				  �� �κп��� �÷� �߰��ϰ� ����Ÿ ������ �Ͼ�� �մϴ�.   							=
'========================================================================================================
Function OKClick()
	Dim ii,jj,kk

	If frm1.vspdData.SelModeSelCount > 0 Then 			
		Redim arrReturn(frm1.vspdData.SelModeSelCount - 1,C_MaxKey)
		kk = 0
		For ii = 0 To frm1.vspdData.MaxRows - 1
			frm1.vspdData.Row = ii + 1			
			If frm1.vspdData.SelModeSelected Then
				For jj = 0 To C_MaxKey - 1
					frm1.vspdData.Col	 = GetKeyPos("A",jj + 1)		
					arrReturn(kk,jj) = frm1.vspdData.Text
				Next		
				
				arrReturn(kk,C_MaxKey) = frm1.txtDocCur.value
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
'------------------------------------------  OpenDept()  ---------------------------------------
'	Name : OpenBp()
'	Description : OpenAp Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function OpenBp(Byval strCode, byval iWhere)
	Dim arrRet
	Dim arrParam(5)

	If IsOpenPop = True Then Exit Function
	
	If iWhere = 1 Then
		if UCase(frm1.txtBpCd.className) = "PROTECTED" Then Exit Function
	End if
	
	IsOpenPop = True

	arrParam(0) = strCode								'  Code Condition
   	arrParam(1) = "A_OPEN_AP"							' ä�ǰ� ����(�ŷ�ó ����)
	arrParam(2) = frm1.txtApDt.text								'FrDt
	arrParam(3) = frm1.txtToApDt.Text								'ToDt
	arrParam(4) = "S"							'B :���� S: ���� T: ��ü 
	Select Case iWhere
		Case 1
			arrParam(5) = "PAYTO"									'SUP :����ó PAYTO: ����ó SOL:�ֹ�ó PAYER :����ó INV:���ݰ�� 	
		Case 2
			arrParam(5) = "SUP"		
	
	End Select
	arrRet = window.showModalDialog("../../Module/ar/BpPopup.asp", Array(window.PopupParent,arrParam), _
		"dialogWidth=780px; dialogHeight=450px; : Yes; help: No; resizable: No; status: No;")
		
	
	IsOpenPop = False
	
	
	If 	arrRet(0) <> "" Then			
		Call SetBpCd(arrRet, iWhere)
	else
		If iWhere = 1 Then
			frm1.txtBpCd.focus
		Else 
			frm1.txtDealBpCd.focus
		End If
	End If	
End Function
 '******************************************  2.4 POP-UP ó���Լ�  ****************************************
'	���: ���� POP-UP
'   ����: ���� POP-UP�� ���� Open�� Include�Ѵ�. 
'	      �ϳ��� ASP���� Popup�� �ߺ��Ǹ� �ϳ��� ��ũ���� ����ϰ� �������� �������Ͽ� ����Ѵ�.
'********************************************************************************************************* 
 '------------------------------------------  OpenBpCd()  -------------------------------------------------
'	Name : OpenBpCd()
'	Description : Bp PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenBpCd(ByVal BpPos)'
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function
	If frm1.txtBpCd.className = "protected" Then Exit Function	
	
	IsOpenPop = True
	
Select Case BpPos
		Case 1
			arrParam(0) = "����ó�˾�"
			arrParam(1) = "(SELECT DISTINCT A.BP_CD,A.BP_NM,A.CURRENCY FROM B_BIZ_PARTNER A, A_OPEN_AP B " 
			arrParam(1) = arrParam(1) & "WHERE  A.BP_CD=B.PAY_BP_CD AND B.CONF_FG = " & FilterVar("C", "''", "S") & "  AND B.AP_STS=" & FilterVar("O", "''", "S") & "  AND B.BAL_AMT <> 0" 
			
			IF frm1.txtApDt.Text<>"" THEN 	arrParam(1) = arrParam(1) & " AND Ap_DT >= " & FilterVar(UNIConvDate(frm1.txtApDt.Text), "''", "S") & ""
			IF frm1.txtToApDt.Text<>"" THEN arrParam(1) = arrParam(1) & " AND Ap_DT <= " & FilterVar(UNIConvDate(frm1.txtToApDt.Text), "''", "S") & ""

			arrParam(1) = arrParam(1) & ") TMP"
			arrParam(2) = Trim(frm1.txtBpCd.Value)
			arrParam(3) = ""
			arrParam(4) = ""
			arrParam(5) = "����ó"			
	
			arrField(0) = "TMP.BP_CD"	
			arrField(1) = "TMP.BP_NM"
			arrField(2) = "TMP.CURRENCY"	
	   
			arrHeader(0) = "����ó"		
			arrHeader(1) = "����ó��"
			arrHeader(2) = "�ŷ���ȭ"
	   Case 2
			arrParam(0) = "����ó�˾�"
			arrParam(1) = "(SELECT DISTINCT A.BP_CD,A.BP_NM,A.CURRENCY FROM B_BIZ_PARTNER A, A_OPEN_AP B " 
			arrParam(1) = arrParam(1) & "WHERE  A.BP_CD=B.PAY_BP_CD AND B.CONF_FG = " & FilterVar("C", "''", "S") & "  AND B.AP_STS=" & FilterVar("O", "''", "S") & "  AND B.BAL_AMT <> 0" 
			IF frm1.txtApDt.Text<>"" THEN 	arrParam(1) = arrParam(1) & " AND Ap_DT >= " & FilterVar(UNIConvDate(frm1.txtApDt.Text), "''", "S") & ""
			IF frm1.txtToApDt.Text<>"" THEN arrParam(1) = arrParam(1) & " AND Ap_DT <= " & FilterVar(UNIConvDate(frm1.txtToApDt.Text), "''", "S") & ""

			arrParam(1) = arrParam(1) & ") TMP"
			
			arrParam(2) = Trim(frm1.txtBpCd.Value)
			arrParam(3) = ""
			arrParam(4) = ""
			arrParam(5) = "����ó"			
	
			arrField(0) = "TMP.BP_CD"	
			arrField(1) = "TMP.BP_NM"
			arrField(2) = "TMP.CURRENCY"	
	   
			arrHeader(0) = "����ó"		
			arrHeader(1) = "����ó��"
			arrHeader(2) = "�ŷ���ȭ"
   End Select
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If 	arrRet(0) <> "" Then			
		Call SetBpCd(arrRet, BpPos)
	else
		If BpPos = 1 Then
			frm1.txtBpCd.focus
		Else 
			frm1.txtDealBpCd.focus
		End If
	End If
End Function

 '------------------------------------------  OpenBizCd()  -------------------------------------------------
'	Name : OpenBizCd()
'	Description : Cost PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenBizCd()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function
	If frm1.txtBizCd.className = "protected" Then Exit Function
	
	IsOpenPop = True

	arrParam(0) = "������˾�"			' �˾� ��Ī 
	arrParam(1) = "B_Biz_Area"					' TABLE ��Ī 
	arrParam(2) = Trim(frm1.txtBizCd.Value)		' Code Condition
	arrParam(3) = ""								' Name Cindition
	' ���Ѱ��� �߰� 
	If lgAuthBizAreaCd <> "" Then
		arrParam(4) = " BIZ_AREA_CD = " & FilterVar(lgAuthBizAreaCd, "''", "S")			' Where Condition
	Else
		arrParam(4) = ""
	End If
	arrParam(5) = "�����"			
	
    arrField(0) = "BIZ_AREA_CD"							' Field��(0)
    arrField(1) = "BIZ_AREA_NM"							' Field��(1)
    
    arrHeader(0) = "�����"					' Header��(0)
    arrHeader(1) = "������"				' Header��(1)
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If 	arrRet(0) <> "" Then		
		Call SetBizCd(arrRet)
	Else
		frm1.txtBizCd.focus
	End If
End Function

'======================================================================================================
'   Event Name : OpenCurrencyInfo
'   Event Desc : 
'=======================================================================================================
Function  OpenCurrencyInfo(Byval strCode)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function
	IF IsDocPop = False Then Exit Function
	
	IsOpenPop = True

	arrParam(0) = "�ŷ���ȭ�˾�"					' �˾� ��Ī 
	arrParam(1) = "b_currency"							' TABLE ��Ī 
	arrParam(2) = strCode						 	    ' Code Condition
	arrParam(3) = ""									' Name Cindition
	arrParam(4) = ""									' Where Condition
	arrParam(5) = "�ŷ���ȭ" 			
	
    arrField(0) = "CURRENCY"							' Field��(0)
    arrField(1) = "CURRENCY_DESC"						' Field��(1)
        
    arrHeader(0) = "�ŷ���ȭ"						' Header��(0)
    arrHeader(1) = "�ŷ���ȭ��"						' Header��(1)
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtDocCur.focus
		Exit Function
	Else
		Call SetCurrencyInfo(arrRet)
	End If	
End Function

'======================================================================================================
'   Event Name : SetCurrencyInfo
'   Event Desc : 
'=======================================================================================================
Function SetCurrencyInfo(Byval arrRet)'
	frm1.txtDocCur.value = arrRet(0)
End Function

'++++++++++++++++  Insert Your Code for PopUp(Open)  ++++++++++++++++++++++++++++++++++++++++++++++++++ 
 '------------------------------------------  SetBpCd()  --------------------------------------------------
'	Name : SetBpCd()
'	Description : Item Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetBpCd(Byval arrRet,ByVal BpPos)	
	If BpPos = 1 Then
		frm1.txtBpCd.value = arrRet(0)		
		frm1.txtBpNm.value = arrRet(1)
		frm1.txtDocCur.value = arrRet(2)
		frm1.txtBpCd.focus
	Else 
		frm1.txtDealBpCd.value = arrRet(0)
		frm1.txtDealBpNm.value = arrRet(1)
		frm1.txtDealBpCd.focus
	End If
End Function
 '------------------------------------------  SetBizCd()  --------------------------------------------------
'	Name : SetBizCd()
'	Description : Item Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetBizCd(Byval arrRet)
	frm1.txtBizCd.value = arrRet(0)		
	frm1.txtBizNm.value = arrRet(1)
	frm1.txtBizCd.focus
End Function


'===========================================================================
' Function Name : OpenOrderBy
' Function Desc : OpenOrderBy Reference Popup
'===========================================================================
Function  OpenOrderByPopup()

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


'########################################################################################################
'#						3. Event ��																		#
'#	���: Event �Լ��� ���� ó��																		#
'#	����: Windowó��, Singleó��, Gridó�� �۾�.														#
'#		  ���⼭ Validation Check, Calcuration �۾��� ������ Event�� �߻�.								#
'#		  �� Object������ Grouping�Ѵ�.																	#
'########################################################################################################


'********************************************  3.1 Windowó��  ******************************************
'*	Window�� �߻� �ϴ� ��� Even ó��																	*
'********************************************************************************************************


'=========================================  3.1.1 Form_Load()  ==========================================
'=	Name : Form_Load()																					=
'=	Description : Window Load�� �����ʱ�ȭ �� ȭ���ʱ�ȭ�� �ϱ� ���� �Լ��� Call�ϴ� �κ�				=
'========================================================================================================

Sub Form_Load()
	Call LoadInfTB19029()

   Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
    Call ggoOper.LockField(Document, "N")                                 

	Call InitVariables()														
	Call SetDefaultVal()	
	Call InitSpreadSheet()
End Sub

'=========================================  3.1.2 Form_QueryUnload()  ===================================
'   Event Name : Form_QueryUnload																		=
'   Event Desc :																						=
'========================================================================================================
Sub Form_QueryUnload(Cancel, UnloadMode)
		
End Sub

'*********************************************  3.2 Tag ó��  *******************************************
'*	Document�� TAG���� �߻� �ϴ� Event ó��																*
'*	Event�� ��� �Ʒ��� ����� Event�̿��� ����� �����ϸ� �ʿ�� �߰� �����ϳ�							*
'*	Event�� �浹�� ����Ͽ� �ۼ��Ѵ�.																	*
'********************************************************************************************************


'==========================================  3.2.1 FncQuery =======================================
'========================================================================================================
Function FncQuery()
	FncQuery = False                                            
    
    Err.Clear                                                   
    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")						
	ggoSpread.Source = frm1.vspdData
	ggoSpread.ClearSpreadData
	
    Call InitVariables()
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then							
		Exit Function
    End If
	
	lgQueryFlag = "1"	
	lgCode = ""		

	If Not ChkQueryDate Then
		Exit Function
    End If

	
    '-----------------------
    'Query function call area
    '-----------------------
	If DbQuery = False Then Exit Function

    FncQuery = True													
End Function

'========================================================================================================
' Name : FncCancel
' Desc : This function is called by MainCancel in Common.vbs
'========================================================================================================
Function FncCancel() 

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

'*********************************************  3.3 Object Tag ó��  ************************************
'*	Object���� �߻� �ϴ� Event ó��																		*
'********************************************************************************************************
Sub  txtApDt_KeyPress(KeyAscii)
	On Error Resume Next
	If KeyAscii = 27 Then 
		Call CancelClick()
	ElseIf KeyAscii = 13 Then 
		Call Fncquery()
	End IF
End Sub

Sub  txtToApDt_KeyPress(KeyAscii)
	On Error Resume Next
	If KeyAscii = 27 Then 
		Call CancelClick()
	ElseIf KeyAscii = 13 Then 
		Call Fncquery()
	End IF
End Sub

Function vspdData_KeyPress(KeyAscii)
    If KeyAscii = 13 And frm1.vspdData.ActiveRow > 0 Then
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
    
	If frm1.vspdData.MaxRows < NewTop + PopupParent.VisibleRowCnt(frm1.vspdData,NewTop) Then	    
    	If lgPageNo <> "" Then								
           If DbQuery = False Then
              Exit Sub
           End if
    	End If
    End If
End Sub

'========================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : �÷��� Ŭ���� ��� �߻� 
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
    
	Call SetSpreadColumnValue("A",frm1.vspdData,Col,Row)			
	
    If frm1.vspdData.MaxRows = 0 Then                                                    'If there is no data.
		Exit Sub
   	End If
End Sub

'========================================================================================================
'   Event Name : vspdData_DblClick
'   Event Desc :
'========================================================================================================
Sub  vspdData_DblClick(ByVal Col, ByVal Row)
	If Frm1.vspdData.MaxRows > 0 Then
		If Frm1.vspdData.ActiveRow = Row Or Frm1.vspdData.ActiveRow > 0 Then
			Call OKClick()
		End If
	End If
End Sub

'########################################################################################################
'#					     4. Common Function��															#
'########################################################################################################

'=======================================================================================================
'   Event Name : txtApDt_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub  txtApDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtApDt.Action = 7                        
		Call SetFocusToDocument("P")
		Frm1.txtApDt.Focus    
    End If
End Sub

'=======================================================================================================
'   Event Name : txtApDt_Change()
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub  txtApDt_Change()
    
End Sub
'=======================================================================================================
'   Event Name : txtToApDt_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub  txtToApDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtToApDt.Action = 7                        
		Call SetFocusToDocument("P")
		Frm1.txtToApDt.Focus       
    End If
End Sub

'=======================================================================================================
'   Event Name : txtToApDt_Change()
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub  txtToApDt_Change()
    
End Sub


Sub txtRcptAmt_KeyPress(KeyAscii)
	On Error Resume Next
	If KeyAscii = 27 Then 
		Call CancelClick()
	ElseIf KeyAscii = 13 Then 
		Call Fncquery()
	End IF
End Sub
'########################################################################################################
'#						5. Interface ��																	#
'########################################################################################################


'********************************************  5.1 DbQuery()  *******************************************
' Function Name : DbQuery																				*
' Function Desc : This function is data query and display												*
'********************************************************************************************************
Function DbQuery()
	Dim strVal

    Err.Clear                                                       
    DbQuery = False
    
	Call LayerShowHide(1)
    
    With frm1

        strVal = BIZ_PGM_ID
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
        If lgIntFlgMode  <> PopupParent.OPMD_UMODE Then   ' This means that it is first search
			strVal = strVal & "?txtBizCd="		& Trim(.txtBizCd.Value)
			strVal = strVal & "&txtBpCd="		& Trim(.txtBpCd.value)					'��: ��ȸ ���� ����Ÿ 
			strVal = strVal & "&txtApDt="		& Trim(.txtApDt.Text)					'��: ��ȸ ���� ����Ÿ 
			strVal = strVal & "&txtToApDt="		& Trim(.txtToApDt.Text)					'��: ��ȸ ���� ����Ÿ 
			strVal = strVal & "&txtDocCur="		& Trim(.txtDocCur.value)					'��: ��ȸ ���� ����Ÿ 
			strVal = strVal & "&txtBpcd_Alt="	& Trim(.txtBpCd.alt)
			strVal = strVal & "&txtBizCd_Alt="	& Trim(.txtBizCd.alt)    
			strVal = strVal & "&txtDealBpCd="	& Trim(.txtDealBpCd.value)
			strVal = strVal & "&txtApNo="		& Trim(.txtApNo.value)	
        Else
			strVal = strVal & "?txtBizCd="		& Trim(.htxtBizCd.value)
			strVal = strVal & "&txtBpCd="		& Trim(.htxtBpCd.value)					'��: ��ȸ ���� ����Ÿ 
			strVal = strVal & "&txtApDt="		& Trim(.htxtApDt.value)					'��: ��ȸ ���� ����Ÿ 
			strVal = strVal & "&txtToApDt="		& Trim(.htxtToApDt.value)					'��: ��ȸ ���� ����Ÿ 
			strVal = strVal & "&txtDocCur="		& Trim(.htxtDocCur.value)					'��: ��ȸ ���� ����Ÿ 
			strVal = strVal & "&txtBpcd_Alt="	& Trim(.txtBpCd.alt)
			strVal = strVal & "&txtBizCd_Alt="	& Trim(.txtBizCd.alt) 	
			strVal = strVal & "&txtDealBpCd="	& Trim(.htxtDealBpCd.value)			'��: ��ȸ ���� ����Ÿ 
			strVal = strVal & "&txtApNo="		& Trim(.htxtApNo.value)					'��: ��ȸ ���� ����Ÿ 
        End If   
           
    '--------- Developer Coding Part (End) ------------------------------------------------------------
        strVal = strVal & "&txtAllcDt="	     & Trim(.htxtAllcDt.value)
        strVal = strVal & "&txtRcptAmt="	 & Trim(.txtRcptAmt.value)
        strVal = strVal & "&lgPageNo="       & lgPageNo         
        strVal = strVal & "&lgMaxCount="     & C_SHEETMAXROWS_D
        strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")
        strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
		strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))

		' ���Ѱ��� �߰� 
		strVal = strVal & "&lgAuthBizAreaCd="	& lgAuthBizAreaCd			' ����� 
		strVal = strVal & "&lgInternalCd="		& lgInternalCd				' ���κμ� 
		strVal = strVal & "&lgSubInternalCd="	& lgSubInternalCd			' ���κμ�(��������)
		strVal = strVal & "&lgAuthUsrID="		& lgAuthUsrID				' ���� 

	'--------------- ������ coding part(�������,Start)----------------------------------------------
        Call RunMyBizASP(MyBizASP, strVal)							
        
    End With
    
    DbQuery = True
End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================
Function DbQueryOk()												
	lgBlnFlgChgValue = False
    lgIntFlgMode     = PopupParent.OPMD_UMODE												'��: Indicates that current mode is Update mode
    
    If frm1.vspdData.MaxRows > 0 Then
 		frm1.vspdData.Focus
 	End If
End Function

Function DetailConditionClick()
	If DetailCondition.style.display = "none" Then
		DetailCondition.style.display = ""
		Call ggoOper.SetReqAttr(frm1.txtBpCd,   "D")
	Else
		DetailCondition.style.display = "none"
		If arrParam(0) <> "" Then
			Call ggoOper.SetReqAttr(frm1.txtBpCd,   "Q")
		Else
			Call ggoOper.SetReqAttr(frm1.txtBpCd,   "N")
		End If
	End If
End Function

'=======================================================================================================
'   Function Name : ChkQueryDate
'   Function Desc : 
'=======================================================================================================
Function ChkQueryDate()
	chkQueryDate= True
	
	If PopupParent.CompareDateByFormat(frm1.txtApDt.text,frm1.txtToApDt.text,frm1.txtApDt.Alt,frm1.txtToApDt.Alt, _
	               "970025",frm1.txtApDt.UserDefinedFormat,PopupParent.gComDateType, true) = False Then
		chkQueryDate= False
		frm1.txtApDt.focus
		Exit Function
	End If

	
	If CompareDateByFormat(frm1.txtApDt.text,frm1.htxtAllcDt.Value,frm1.txtApDt.Alt,frm1.htxtAllcAlt.value, _
   	           "970025",frm1.txtApDt.UserDefinedFormat,PopupParent.gComDateType, true) = False Then
	   chkQueryDate= False
	   frm1.txtApDt.focus
	   Exit Function
	End If
	
	If CompareDateByFormat(frm1.txtToApDt.text,frm1.htxtAllcDt.Value,frm1.txtToApDt.Alt, frm1.htxtAllcAlt.value,_
   	           "970025",frm1.txtToApDt.UserDefinedFormat,PopupParent.gComDateType, true) = False Then
	   chkQueryDate= False
	   frm1.txtToApDt.focus
	   Exit Function
	End If

End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->	
</HEAD>
<!--
'########################################################################################################
'						6. Tag ��																		
'########################################################################################################
 -->
<BODY SCROLL=NO TABINDEX="-1">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="post">
<TABLE <%=LR_SPACE_TYPE_20%>>

	<TR>
		<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
	</TR>
	<TR>
		<TD HEIGHT=20>
			<FIELDSET CLASS="CLSFLD">
				<TABLE <%=LR_SPACE_TYPE_40%>*>
					<TR>				
						<TD CLASS=TD5 NOWRAP>ä������</TD>
						<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME="txtApDt" CLASS=FPDTYYYYMMDD tag="12" Title="FPDATETIME" ALT="ä����������" id=fpDateTime></OBJECT>');</SCRIPT>								
							    &nbsp;~&nbsp;<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME="txtToApDt" CLASS=FPDTYYYYMMDD tag="12" Title="FPDATETIME" ALT="ä����������" id=fpDateTime1></OBJECT>');</SCRIPT></TD>												
						<TD CLASS=TD5 NOWRAP>�ŷ���ȭ</TD>
						<TD CLASS=TD6 NOWRAP><INPUT NAME="txtDocCur" ALT="�ŷ���ȭ" MAXLENGTH="3" SIZE=10 STYLE="TEXT-ALIGN: left" tag ="12NXXU"><IMG align=top name=btnCalType onclick="vbscript:OpenCurrencyInfo(txtDocCur.Value)" src="../../../CShared/image/btnPopup.gif"  TYPE="BUTTON"></TD>
					</TR>
					<TR>
						<TD CLASS=TD5 NOWRAP>����ó</TD>
						<TD CLASS=TD6 NOWRAP><INPUT TYPE="Text" NAME="txtBpCd" SIZE=10 MAXLENGTH=10 STYLE="TEXT-ALIGN: left" tag="12NXXU" ALT="����ó"><IMG align=top name=btnBpcd onclick="vbscript:Call OpenBp(frm1.txtBpCd.value,1)" src="../../../CShared/image/btnPopup.gif"  TYPE="BUTTON"> <INPUT TYPE=TEXT NAME="txtBpNm" SIZE=20 MAXLENGTH=20 STYLE="TEXT-ALIGN: Left" tag="14" ALT="�ŷ�ó��"></TD>
						<TD CLASS=TD5 NOWRAP>�����</TD>
						<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtBizCd" SIZE=10 MAXLENGTH=10 STYLE="TEXT-ALIGN: left" tag=11NXXU" ALT="�����"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDeptCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenBizCd()"> <INPUT TYPE=TEXT NAME="txtBizNm" SIZE=20 MAXLENGTH=20 STYLE="TEXT-ALIGN: left" tag="14" ALT="������">
											 <IMG SRC="../../../CShared/image/icon/QualityC.gif" Style="CURSOR: hand" ALT="OK" NAME="pop1" ONCLICK="DetailConditionClick()" ></IMG></TD>					
					</TR>
					<TR ID="DetailCondition" style="display: none">
						<TD CLASS=TD5 NOWRAP>����ó</TD>
						<TD CLASS=TD6 NOWRAP><INPUT TYPE="Text" NAME="txtDealBpCd" SIZE=10 MAXLENGTH=10 STYLE="TEXT-ALIGN: Left" tag="11NXXU" ALT="����ó"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDealBpCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenBp(frm1.txtDealBpCd.value,2)"> <INPUT TYPE=TEXT NAME="txtDealBpNm" SIZE=20 MAXLENGTH=20 STYLE="TEXT-ALIGN: Left" tag="14" ALT="�ֹ�ó"></TD>
						<TD CLASS=TD5 NOWRAP>ä����ȣ</TD>
						<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtApNo" MAXLENGTH=18 STYLE="TEXT-ALIGN: Left" tag=11NXXU" ALT="ä����ȣ"></TD>					
					</TR>
					<TR>
						<TD CLASS=TD5 NOWRAP>��ݱݾ�</TD>
						<TD CLASS=TD6 NOWRAP>
						<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name="txtRcptAmt" style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 150px" title=FPDOUBLESINGLE ALT="�Աݱݾ�" tag="11X2" id=OBJECT1></OBJECT>');</SCRIPT>											
						</TD>
						<TD CLASS=TD5 NOWRAP></TD>
						<TD CLASS=TD6 NOWRAP>
						</TD>
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
						<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% tag="2" HEIGHT=100% id=vspdData> <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> <PARAM NAME="ReDraw" VALUE="0"> <PARAM NAME="FontSize" VALUE="10"> </OBJECT>');</SCRIPT>
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
						<IMG SRC="../../../CShared/image/query_d.gif"  Style="CURSOR: hand" ALT="Search" NAME="Search" ONCLICK="Call FncQuery()">	</IMG>&nbsp;
					<IMG SRC="../../../CShared/image/zpConfig_d.gif" Style="CURSOR: hand" ALT="Config" NAME=Config ONMOUSEOUT="javascript:MM_swapImgRestore()" ONMOUSEOVER="javascript:MM_swapImage(this.name,'','../../../CShared/image/zpConfig.gif',1)" ONCLICK="OpenOrderByPopup()"></IMG></TD>
					<TD WIDTH=30% ALIGN=RIGHT>
						<IMG SRC="../../../CShared/image/ok_d.gif" Style="CURSOR: hand" ALT="OK" NAME="pop1" ONCLICK="OkClick()" ></IMG>&nbsp;
						<IMG SRC="../../../CShared/image/cancel_d.gif" Style="CURSOR: hand" ALT="CANCEL" NAME="pop2" ONCLICK="CancelClick()" ></IMG>
					</TD>				
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP"  WIDTH=100% HEIGHT= <%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="htxtBizCd"		tag="24">
<INPUT TYPE=HIDDEN NAME="htxtBpCd"		tag="24">
<INPUT TYPE=HIDDEN NAME="htxtApDt"		tag="24">
<INPUT TYPE=HIDDEN NAME="htxtToApDt"        tag="24">
<INPUT TYPE=HIDDEN NAME="htxtDocCur"        tag="24">
<INPUT TYPE=HIDDEN NAME="htxtDealBpCd"  tag="24">
<INPUT TYPE=HIDDEN NAME="htxtApNo"      tag="24">
<INPUT TYPE=HIDDEN NAME="htxtAllcDt"	tag="14">
<INPUT TYPE=HIDDEN NAME="htxtAllcAlt"      tag="14">
</FORM>

<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

