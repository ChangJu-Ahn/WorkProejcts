
<%@ LANGUAGE="VBSCRIPT" %>
<!--
'********************************************************************************************************
'*  1. Module Name          : Basis Architect															*
'*  2. Function Name        : Reference Popup Business Part												*
'*  3. Program ID           : a4112ra1.asp																			*
'*  4. Program Name         : ȸ�����-ä������-																			*
'*  5. Program Desc         : Reference Popup															*
'*  7. Modified date(First) : 2000/03/29																*
'*  8. Modified date(Last)  : 2000/03/29																*
'*  9. Modifier (First)     : Kang Tae Bum																*
'* 10. Modifier (Last)      : Kang Tae Bum																*
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

Const BIZ_PGM_ID 		= "a4112rb1.asp"                              '��: Biz Logic ASP Name

'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
Const C_SHEETMAXROWS_D  = 30                                          '��: Server���� �ѹ��� fetch�� �ִ� ����Ÿ �Ǽ� 
Const C_MaxKey          = 20					                          '��: SpreadSheet�� Ű�� ���� 

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

Dim  arrReturn
Dim  arrParent
Dim  arrParam		
		
Dim  IsOpenPop     
 
' ���Ѱ��� �߰� 
Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' ����� 
Dim lgInternalCd, lgDeptCd, lgDeptNm			' ���κμ� 
Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' ���κμ�(��������)
Dim lgAuthUsrID, lgAuthUsrNm					' ���� 

	 '------ Set Parameters from Parent ASP ------ 
arrParent        = window.dialogArguments
Set PopupParent = arrParent(0)	 
arrParam		= arrParent(1)


	Dim strYear, strMonth, strDay, dtToday, EndDate, StartDate
		dtToday = "<%=GetSvrDate%>"
	Call PopupParent.ExtractDateFrom(dtToday, PopupParent.gServerDateFormat, PopupParent.gServerDateType, strYear, strMonth, strDay)

	EndDate = PopupParent.UniConvYYYYMMDDToDate(PopupParent.gDateFormat, strYear, strMonth, strDay)
	StartDate = PopupParent.UNIDateAdd("M", -1, EndDate, PopupParent.gDateFormat)

top.document.title = PopupParent.gActivePRAspName
'top.document.title = "ä�ǹ߻�����"

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
	If UBound(arrParam) > 7 Then
		lgAuthBizAreaCd		= arrParam(8)
		lgInternalCd		= arrParam(9)
		lgSubInternalCd		= arrParam(10)
		lgAuthUsrID			= arrParam(11)
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

'==========================================  2.2.1 SetDefaultVal()  =====================================
'=	Name : SetDefaultVal()																				=
'=	Description : ȭ�� �ʱ�ȭ(���� Field�� ȭ���� �� �� Default���� ������� �ϴ� Field�� Setting)		=
'========================================================================================================
Sub SetDefaultVal()
	'lblTitle.innerHTML = "Open AP ����"
	With frm1
		.txtBpCd.value = arrParam(0)
		.txtBpNm.value = arrParam(1)
		.txtDocCur.value = arrParam(2)	
			
		.htxtAllcDt.value	= arrParam(4) 
		.htxtAllcAlt.value	= arrParam(5) 	
		' SetReqAttr(Object, Option) ; N : Required, Q : Protect, D : Default
		If .txtBpCd.value <> "" Then		
			Call ggoOper.SetReqAttr(.txtBpCd,   "Q")		
		Else
			Call ggoOper.SetReqAttr(.txtBpCd,   "N")		
		End If
	
		If  .txtDocCur.value <> "" Then		
			Call ggoOper.SetReqAttr(.txtDocCur,   "Q")		
		Else
			Call ggoOper.SetReqAttr(.txtDocCur,   "N")		
		End If	
	End With
End Sub

'========================================================================================================
' Name : InitSpreadSheet
' Desc : This method initializes spread sheet column property
'========================================================================================================
Sub InitSpreadSheet()

	frm1.vspdData.OperationMode = 5
    Call SetZAdoSpreadSheet("A4112RA1","S","A","V20021211",PopupParent.C_SORT_DBAGENT,frm1.vspdData, C_MaxKey, "X","X")
    Call SetSpreadLock() 
    
End Sub


'========================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================================

Sub  SetSpreadLock()
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
		Redim arrReturn(frm1.vspdData.SelModeSelCount - 1,C_MaxKey+2)
		kk = 0
		For ii = 0 To frm1.vspdData.MaxRows - 1
			frm1.vspdData.Row = ii + 1			
			If frm1.vspdData.SelModeSelected Then
				For jj = 0 To C_MaxKey - 1
					frm1.vspdData.Col	 = GetKeyPos("A",jj + 1)		
					arrReturn(kk,jj) = frm1.vspdData.Text
				Next			
				arrReturn(kk, C_MaxKey)     = frm1.txtBpCd.value
				arrReturn(kk, C_MaxKey + 1) = frm1.txtBpNm.value
				arrReturn(kk, C_MaxKey + 2) = frm1.txtDocCur.value    
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
	if UCase(Frm1.txtDocCur.className) = "PROTECTED" Then Exit Function
	
	IsOpenPop = True

	arrParam(0) = "�ŷ���ȭ�˾�"					' �˾� ��Ī 
	arrParam(1) = "b_currency"							' TABLE ��Ī 
	arrParam(2) = Trim(Frm1.txtDocCur.value)							 	    ' Code Condition
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
	
		Frm1.txtDocCur.value = arrRet(0)
		frm1.txtDocCur.focus
	    lgBlnFlgChgValue = True
	
End Function
'------------------------------------------  OpenDept()  ---------------------------------------
'	Name : OpenBp()
'	Description : OpenAp Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function OpenBp(Byval strCode, byval iWhere)
	Dim arrRet
	Dim arrParam(5)

	If IsOpenPop = True Then Exit Function
	if UCase(Frm1.txtBpCd.className) = "PROTECTED" Then Exit Function
	
	
	IsOpenPop = True

	arrParam(0) = strCode								'  Code Condition
   	arrParam(1) = "A_OPEN_AR"							' ä�ǰ� ����(�ŷ�ó ����)
	arrParam(2) = frm1.txtArDt.text							'FrDt
	arrParam(3) = frm1.txtToArDt.text									'ToDt
	arrParam(4) = "B"									'B :���� S: ���� T: ��ü 
	arrParam(5) = ""									'SUP :����ó PAYTO: ����ó SOL:�ֹ�ó PAYER :����ó INV:���ݰ�� 	
	arrRet = window.showModalDialog("../../Module/ar/BpPopup.asp", Array(window.PopupParent,arrParam), _
		"dialogWidth=780px; dialogHeight=450px; : Yes; help: No; resizable: No; status: No;")
		
	
	IsOpenPop = False
	
	IF 	arrRet(0) <> "" then			
		Call SetBpCd(arrRet)
	Else
		frm1.txtBpCd.focus
	end if
End Function
 '------------------------------------------  OpenBpCd()  -------------------------------------------------
'	Name : OpenBpCd()
'	Description : Bp PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenBpCd()'
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function
	IF IsBpPop  = False Then Exit Function
	
	IsOpenPop = True

	arrParam(0) = "�ŷ�ó�˾�"
	arrParam(1) = "B_BIZ_PARTNER"				
	arrParam(2) = Trim(Frm1.txtBpCd.Value)
	arrParam(3) = ""
	arrParam(4) = ""
	arrParam(5) = "�ŷ�ó"			
	
    arrField(0) = "BP_CD"	
    arrField(1) = "BP_NM"	
    
    arrHeader(0) = "�ŷ�ó"		
    arrHeader(1) = "�ŷ�ó��"	
    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	IF 	arrRet(0) <> "" then			
		Call SetBpCd(arrRet)
	Else
		frm1.txtBpCd.focus
	end if
End Function
 '------------------------------------------  OpenBiztCd()  -------------------------------------------------
'	Name : OpenBiztCd()
'	Description : Cost PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenBiztCd()'
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "������˾�"			' �˾� ��Ī 
	arrParam(1) = "B_BIZ_AREA"					' TABLE ��Ī 
	arrParam(2) = Trim(Frm1.txtBizCd.Value)				' Code Condition
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
    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	IF 	arrRet(0) <> "" then		
		Call SetBizCd(arrRet)
	Else
		frm1.txtBizCd.focus
	end if
	
End Function

'++++++++++++++++  Insert Your Code for PopUp(Open)  ++++++++++++++++++++++++++++++++++++++++++++++++++ 
 '------------------------------------------  SetBpCd()  --------------------------------------------------
'	Name : SetBpCd()
'	Description : Item Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetBpCd(Byval arrRet)'
	
	Frm1.txtBpCd.value = arrRet(0)		
	Frm1.txtBpNm.value = arrRet(1)
	frm1.txtBpCd.focus			
	lgBlnFlgChgValue = True
	
End Function
 '------------------------------------------  SetBizCd()  --------------------------------------------------
'	Name : SetBizCd()
'	Description : Item Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetBizCd(Byval arrRet)'
	
	Frm1.txtBizCd.value = arrRet(0)		
	Frm1.txtBizNm.value = arrRet(1)
	frm1.txtBizCd.focus				
	lgBlnFlgChgValue = True
	
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
	Call ggoOper.FormatField(Document, "1",PopupParent.ggStrIntegeralPart, PopupParent.ggStrDeciPointPart,PopupParent.gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec,,,PopupParent.ggStrMinPart,PopupParent.ggStrMaxPart)
	Call ggoOper.LockField(Document, "N")                                   

	Call InitVariables()
	Call SetDefaultVal()
	Call InitSpreadSheet()

	Call MM_preloadImages("../../../CShared/image/Query.gif","../../../CShared/image/OK.gif","../../../CShared/image/Cancel.gif")		
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
	Dim IntRetCD
		
	FncQuery = False     
	Err.Clear          
	

	'-----------------------
	'Check condition area
	'-----------------------
	If Not chkField(Document, "1") Then									'This function check indispensable field
		Exit Function
	End If

	
	If Not ChkQueryDate Then
		Exit Function
    End If

	Call InitVariables 		

	Call ggoOper.ClearField(Document, "2")						
	ggoSpread.Source= frm1.vspdData
	ggoSpread.ClearSpreadData
	lgQueryFlag = "1"	
	lgCode = ""
		
	If DbQuery = False Then Exit Function
		 
	 FncQuery = True	
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

Function vspdData_KeyPress(KeyAscii)
    If KeyAscii = 13 And Frm1.vspdData.ActiveRow > 0 Then
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
'            ggoSpread.SSSort, lgSortKey
			ggoSpread.SSSort Col 
            lgSortKey = 2
        Else
'            ggoSpread.SSSort, lgSortKey
			ggoSpread.SSSort Col,lgSortKey 
            lgSortKey = 1
        End If    
    End If
    
    If frm1.vspdData.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
	
	Call SetSpreadColumnValue("A",frm1.vspdData,Col,Row)		
End Sub
'========================================================================================================
'   Event Name : vspdData_DblClick
'   Event Desc :
'========================================================================================================
Sub  vspdData_DblClick(ByVal Col, ByVal Row)
	If Row = 0 Then              		' Title cell�� dblclick�߰ų�....
		Exit Sub
	End If
	
	If Frm1.vspdData.MaxRows = 0 Then  	'NO Data
		Exit Sub
	End If
	
	Call OKClick()
End Sub


'########################################################################################################
'#					     4. Common Function��															#
'########################################################################################################

'=======================================================================================================
'   Event Name : txtArDt_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub  txtArDt_DblClick(Button)
    If Button = 1 Then
        Frm1.txtArDt.Action = 7      
		Call SetFocusToDocument("P")
		Frm1.txtArDt.Focus
		                          
    End If
End Sub

'=======================================================================================================
'   Event Name : txtToArDt_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub  txtToArDt_DblClick(Button)
    If Button = 1 Then
        Frm1.txtToArDt.Action = 7    
		Call SetFocusToDocument("P")
		Frm1.txtToArDt.Focus                            
    End If
End Sub

'=======================================================================================================
'   Event Name : txtArDueDt_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub  txtArDueDt_DblClick(Button)
    If Button = 1 Then
        Frm1.txtArDueDt.Action = 7                        
        Call SetFocusToDocument("P")
		Frm1.txtArDueDt.Focus 
    End If
End Sub
'=======================================================================================================
'   Event Name : txtToArDueDt_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub  txtToArDueDt_DblClick(Button)
    If Button = 1 Then
        Frm1.txtToArDueDt.Action = 7                        
        Call SetFocusToDocument("P")
		Frm1.txtToArDueDt.Focus 
    End If
End Sub

Sub txtBpCd_KeyPress(KeyAscii)
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
			strVal = strVal & "?txtBizCd=" & Trim(.txtBizCd.Value)
			strVal = strVal & "&txtBpCd=" & Trim(.txtBpCd.value)					'��: ��ȸ ���� ����Ÿ 
			strVal = strVal & "&txtArDt=" & Trim(.txtArDt.Text)					'��: ��ȸ ���� ����Ÿ 
			strVal = strVal & "&txtToArDt=" & Trim(.txtToArDt.Text)					'��: ��ȸ ���� ����Ÿ 
			strVal = strVal & "&txtArDueDt="	& Trim(.txtArDueDt.text)					'��: ��ȸ ���� ����Ÿ 
			strVal = strVal & "&txtToArDueDt="	& Trim(.txtToArDueDt.text)				'��: ��ȸ ���� ����Ÿ 
			strVal = strVal & "&txtDocCur=" & Trim(.txtDocCur.value)					'��: ��ȸ ���� ����Ÿ 
			strVal = strVal & "&txtBpcd_Alt=" & Trim(.txtBpCd.alt)
			strVal = strVal & "&txtBizCd_Alt=" & Trim(.txtBizCd.alt)    	
        Else
			strVal = strVal & "?txtBizCd=" & Trim(.htxtBizCd.value)
			strVal = strVal & "&txtBpCd=" & Trim(.htxtBpCd.value)					'��: ��ȸ ���� ����Ÿ 
			strVal = strVal & "&txtArDt=" & Trim(.htxtArDt.value)					'��: ��ȸ ���� ����Ÿ 
			strVal = strVal & "&txtToArDt=" & Trim(.htxtToArDt.value)					'��: ��ȸ ���� ����Ÿ 
			strVal = strVal & "&txtArDueDt="	& Trim(.htxtArDueDt.value)					'��: ��ȸ ���� ����Ÿ 
			strVal = strVal & "&txtToArDueDt="	& Trim(.htxtToArDueDt.value)				'��: ��ȸ ���� ����Ÿ 
			strVal = strVal & "&txtDocCur=" & Trim(.htxtDocCur.value)					'��: ��ȸ ���� ����Ÿ 
			strVal = strVal & "&txtBpcd_Alt=" & Trim(.txtBpCd.alt)
			strVal = strVal & "&txtBizCd_Alt=" & Trim(.txtBizCd.alt) 	
        End If   
           
    '--------- Developer Coding Part (End) ------------------------------------------------------------
        strVal = strVal & "&txtAllcDt="	     & Trim(.htxtAllcDt.value)
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

'=======================================================================================================
'   Function Name : ChkQueryDate
'   Function Desc : 
'=======================================================================================================
Function ChkQueryDate()
	chkQueryDate= True
	
	If PopupParent.CompareDateByFormat(Frm1.txtArDt.text,Frm1.txtToArDt.text,Frm1.txtArDt.Alt,Frm1.txtToArDt.Alt, _
	    	               "970025",Frm1.txtArDt.UserDefinedFormat,PopupParent.gComDateType, true) = False Then
		chkQueryDate= False
		Frm1.txtArDt.focus
		Exit Function
	End If
	
	If CompareDateByFormat(Frm1.txtArDueDt.text,Frm1.txtToArDueDt.text,Frm1.txtArDueDt.Alt,Frm1.txtToArDueDt.Alt, _
   	           "970025",Frm1.txtArDueDt.UserDefinedFormat,PopupParent.gComDateType, true) = False Then
	   	chkQueryDate= False
	   frm1.txtArDueDt.focus
	   Exit Function
	End If

	
	If CompareDateByFormat(frm1.txtArDt.text,frm1.htxtAllcDt.Value,frm1.txtArDt.Alt,frm1.htxtAllcAlt.value, _
   	           "970025",frm1.txtArDt.UserDefinedFormat,PopupParent.gComDateType, true) = False Then
	   chkQueryDate= False
	   frm1.txtArDt.focus
	   Exit Function
	End If
	
	If CompareDateByFormat(frm1.txtToArDt.text,frm1.htxtAllcDt.Value,frm1.txtToArDt.Alt, frm1.htxtAllcAlt.value,_
   	           "970025",frm1.txtToArDt.UserDefinedFormat,PopupParent.gComDateType, true) = False Then
	   chkQueryDate= False
	   frm1.txtToArDt.focus
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
				<TABLE <%=LR_SPACE_TYPE_40%>>
					<TR>	
						<TD CLASS=TD5 NOWRAP>�ŷ�ó</TD>
						<TD CLASS=TD6 NOWRAP>
							<INPUT TYPE="Text" NAME="txtBpCd" SIZE=11 MAXLENGTH=10 tag="12NXXU" ALT="�ŷ�ó"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBizCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:CALL OpenBp(frm1.txtBpCd.Value, 1)"> 
							&nbsp;<INPUT TYPE=TEXT NAME="txtBpNm" SIZE=20 MAXLENGTH=20 STYLE="TEXT-ALIGN: Left" tag="14" ALT="�ŷ�ó��">
						</TD>
						<TD CLASS=TD5 NOWRAP>�ŷ���ȭ</TD>
						<TD CLASS=TD6 NOWRAP>
							<INPUT NAME="txtDocCur" ALT="�ŷ���ȭ" MAXLENGTH="3" SIZE=11 tag ="12NXXU"><IMG align=top name=btnCalType onclick="vbscript:OpenCurrencyInfo()" src="../../../CShared/image/btnPopup.gif"  TYPE="BUTTON">
						</TD>
					</TR>
					<TR>
						<TD CLASS=TD5 NOWRAP>ä������</TD>
						<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME="txtArDt" CLASS=FPDTYYYYMMDD tag="11" Title="FPDATETIME" ALT="ä�ǽ�������" id=OBJECT3></OBJECT>');</SCRIPT>								
						&nbsp;~&nbsp;<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME="txtToArDt" CLASS=FPDTYYYYMMDD tag="11" Title="FPDATETIME" ALT="ä����������" id=OBJECT4></OBJECT>');</SCRIPT></TD>												
						<TD CLASS=TD5 NOWRAP>��������</TD>
						<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME="txtArDueDt" CLASS=FPDTYYYYMMDD tag="11" Title="FPDATETIME" ALT="�����������" id=OBJECT1></OBJECT>');</SCRIPT>								
						&nbsp;~&nbsp;<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME="txtToArDueDt" CLASS=FPDTYYYYMMDD tag="11" Title="FPDATETIME" ALT="������������" id=OBJECT2></OBJECT>');</SCRIPT></TD>												
					</TR>
					<TR>
						<TD CLASS=TD5 NOWRAP>�����</TD>
						<TD CLASS=TD6 NOWRAP>
							<INPUT TYPE=TEXT NAME="txtBizCd" SIZE=11 MAXLENGTH=10 tag=11NXXU" ALT="�����"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCostCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenBiztCd()">
							&nbsp;<INPUT TYPE=TEXT NAME="txtBizNm" SIZE=20 MAXLENGTH=20 STYLE="TEXT-ALIGN: left" tag="14" ALT="������">
						</TD>
						<TD CLASS=TD5 NOWRAP>&nbsp;</TD>
						<TD CLASS=TD6 NOWRAP>&nbsp;</TD>									
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
						<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% tag="2" HEIGHT=100% > <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> <PARAM NAME="ReDraw" VALUE="0"> <PARAM NAME="FontSize" VALUE="10"> </OBJECT>');</SCRIPT>
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
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP"  WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="htxtBizCd"		tag="24">
<INPUT TYPE=HIDDEN NAME="htxtBpCd"		tag="24">
<INPUT TYPE=HIDDEN NAME="htxtArDt"		tag="24">
<INPUT TYPE=HIDDEN NAME="htxtToArDt"        tag="24">
<INPUT TYPE=HIDDEN NAME="htxtArDueDt"		tag="24">
<INPUT TYPE=HIDDEN NAME="htxtToArDueDt"    tag="24">
<INPUT TYPE=HIDDEN NAME="htxtDocCur"        tag="24">
<INPUT TYPE=HIDDEN NAME="htxtAllcDt"	tag="14">
<INPUT TYPE=HIDDEN NAME="htxtAllcAlt"      tag="14">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

