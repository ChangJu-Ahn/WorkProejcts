<%@ LANGUAGE="VBSCRIPT" %>
<% Response.Expires = -1 %>
<!--
'********************************************************************************************************
'*  1. Module Name          : Procurement																*
'*  2. Function Name        :																			*
'*  3. Program ID           : M9211RA1														*
'*  4. Program Name         : 
'*  5. Program Desc         : 
'*  6. Comproxy List        : 																			*
'*  7. Modified date(First) : 2000/03/21																*
'*  8. Modified date(Last)  : 2002/05/07																*
'*  9. Modifier (First)     : Shin Jin-hyun																*
'* 10. Modifier (Last)      : KO MYOUNG JIN
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"									*
'*                            this mark(��) Means that "may  change"										*
'*                            this mark(��) Means that "must change"										*
'********************************************************************************************************
-->
<HTML>
<HEAD>
<!--<TITLE>���ֳ�������</TITLE>-->
<TITLE></TITLE>
<!--
'#########################################################################################################
'												1. �� �� �� 
'##########################################################################################################
'******************************************  1.1 Inc ����   **********************************************
'	���: Inc. Include
'*********************************************************************************************************
-->
<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->
<!--
'==========================================  1.1.1 Style Sheet  ======================================
'==========================================================================================================
-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--
'==========================================  1.1.2 ���� Include   ======================================
'==========================================================================================================
-->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentA.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs">   </SCRIPT>

<SCRIPT LANGUAGE="JavaScript" SRC="../../inc/incImage.js"></SCRIPT>
<Script Language="VBScript">

Option Explicit		

'########################################################################################################
'#                       4.  Data Declaration Part
'########################################################################################################

'========================================================================================================
'=                       4.1 External ASP File
'========================================================================================================

Const BIZ_PGM_ID 		= "m9211rb1.asp"                              '��: Biz Logic ASP Name

'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
Const C_SHEETMAXROWS_D  = 30                                          '��: Fetch max count at once
Const C_MaxKey          = 28                                           '��: key count of SpreadSheet
'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim gblnWinEvent
Dim arrReturn					<% '--- Return Parameter Group %>
Dim arrParent
Dim arrParam

Dim C_PO_NO_REF             
Dim C_PO_SEQ_NO_REF         
Dim C_PLANT_CD_REF          
Dim C_PLANT_NM_REF          
Dim C_ITEM_CD_REF           
Dim C_ITEM_NM_REF           
Dim C_spec_REF              
Dim C_GI_QTY_REF            
Dim C_GI_UNIT_REF           
Dim C_PRICE_REF           
Dim C_GI_AMT_REF        
Dim C_bp_cd_REF             
Dim C_bp_nm_REF            
Dim C_DN_NO_REF     
Dim C_DN_SEQ_REF    
Dim C_SL_CD_REF            
Dim C_SL_NM_REF          
Dim C_CUR_REF            
Dim C_TRACKING_NO_REF     
Dim C_trns_lot_no_REF        
Dim C_trns_lot_sub_no_REF  
Dim C_lot_no_REF          
Dim C_lot_sub_no_REF       
Dim C_GI_AMT_LOC_REF        
Dim C_BASE_UNIT_REF         
Dim C_BASE_QTY_REF          
Dim C_PUR_GRP_REF			  
    
arrParent = window.dialogArguments
Set PopupParent = arrParent(0)
top.document.title = PopupParent.gActivePRAspName
arrParam= arrParent(1)

'########################################################################################################
'#						2. Function ��																	#
'#																										#
'#	���� : �����ڰ� ������ �Լ�, �� Event���� �Լ��� ������ ��� ����� ���� �Լ� ���					#
'#	�������� ���� ���� : 1. Sub �Ǵ� Function�� ȣ���� �� �ݵ�� Call�� ����.							#
'#						 2. Sub, Function �̸��� _�� ���� �ʵ��� �Ѵ�. (Event�� �����ϱ� ����)			#
'########################################################################################################
'*******************************************  2.1 ���� �ʱ�ȭ �Լ�  *************************************
'*	���: �����ʱ�ȭ																					*
'*	Description : Global���� ó��, �����ʱ�ȭ ���� �۾��� �Ѵ�.											*
'********************************************************************************************************
'==========================================  2.1.1 InitVariables()  =====================================
'=	Name : InitVariables()																				=
'=	Description : ���� �ʱ�ȭ(Global ����, �ʱ�ȭ�� �ʿ��� ���� �Ǵ� Flag���� Setting�Ѵ�.)				=
'========================================================================================================
Function InitVariables()
    lgStrPrevKey     = ""
    lgPageNo         = ""
	lgBlnFlgChgValue = False
    lgIntFlgMode     = PopupParent.OPMD_CMODE                        'Indicates that current mode is Create mode
    lgSortKey        = 1
						
	frm1.vspdData.MaxRows = 0	
			
	gblnWinEvent = False
	ReDim arrReturn(0,0)
	Self.Returnvalue = arrReturn
End Function
	
	
'==========================================  2.2.1 SetDefaultVal()  =====================================
'=	Name : SetDefaultVal()																				=
'=	Description : ȭ�� �ʱ�ȭ(���� Field�� ȭ���� �� �� Default���� ������� �ϴ� Field�� Setting)		=
'========================================================================================================
Sub SetDefaultVal()
	Dim EndDate, StartDate
	'im lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6, i
	'im iCodeArr
		
	Err.Clear
	
	EndDate = UNIConvDateAtoB("<%=GetSvrDate%>", PopupParent.gServerDateFormat, PopupParent.gDateFormat)
	StartDate = UnIDateAdd("m", -1, EndDate, PopupParent.gDateFormat)
	
	With frm1
		.txtFrSGiDt = StartDate
		.txtToSGiDt = EndDate
	
		.txtFrStoDt = StartDate
		.txtToStoDt = EndDate
	
		'txtFrStoDt.text = StartDate
		'txtToStoDt.text = EndDate
	
		.txtSGiCd.value 		= arrParam(0)
		.txtGroupCd.value 		= arrParam(1)
		'.hdnRefType.value 		= arrParam(8)
		'.hdnRcptType.value 		= arrParam(9)
	End With
	
	'Call CommonQueryRs(" RCPT_FLG", " M_MVMT_TYPE", " IO_TYPE_CD = '" & FilterVar(Trim(frm1.hdnRcptType.value),"","SNM") & "'", lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

    'IF Len(lgF0) Then
	'	iCodeArr = Split(lgF0, Chr(11))
		    
	'	If Err.number <> 0 Then
	'		MsgBox Err.description 
	'		Err.Clear 
	'		Exit Sub
	'	End If
	'	frm1.hdnRcptFlg.value 	= iCodeArr(0)
	'End if	
	
End Sub
	
	

'*******************************************  2.2 ȭ�� �ʱ�ȭ �Լ�  *************************************
'*	���: ȭ���ʱ�ȭ																					*
'*	Description : ȭ���ʱ�ȭ, Combo Display, ȭ�� Clear �� ȭ�� �ʱ�ȭ �۾��� �Ѵ�.						*
'********************************************************************************************************
'==========================================  2.2.1 SetDefaultVal()  =====================================
'=	Name : SetDefaultVal()																				=
'=	Description : ȭ�� �ʱ�ȭ(���� Field�� ȭ���� �� �� Default���� ������� �ϴ� Field�� Setting)		=
'========================================================================================================

'========================================================================================================
' Name : LoadInfTB19029()	
' Desc : Set System Number format
' Para : 1. Currency
'        2. I(Input),Q(Query),P(Print),B(Bacth)
'        3. "*" is for Common module
'           "A" is for Accounting
'           "I" is for Inventory
'           ...
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "*", "NOCOOKIE", "RA") %>
End Sub
'========================================================================================================
' Name : InitSpreadSheet
' Desc : This method initializes spread sheet column property
'========================================================================================================
Sub InitSpreadSheet()

	Call SetZAdoSpreadSheet("m9211ra1","S","A","V20021202",PopupParent.C_SORT_DBAGENT,frm1.vspdData, C_MaxKey, "X","X")
    Call SetSpreadLock("A")
    frm1.vspdData.OperationMode = 5 
End Sub

'========================================================================================================
' Name : SetSpreadLock
' Desc : This method set color and protect in spread sheet celles
'========================================================================================================
Sub SetSpreadLock(ByVal pOpt)
    IF pOpt = "A" Then
		With frm1
			.vspdData.ReDraw = False
			'------ Developer Coding part (Start ) -------------------------------------------------------------- 
			ggoSpread.SpreadLock 1 , -1
			'------ Developer Coding part (End   ) -------------------------------------------------------------- 
			.vspdData.ReDraw = True

		End With
	Else
	
	End IF
End Sub

'==========================================  2.2.6 InitComboBox()  ======================================
'=	Name : InitComboBox()																				=
'=	Description : Combo Display																			=
'========================================================================================================
'++++++++++++++++++++++++++++++++++++++++++  2.3 ������ ���� �Լ�  ++++++++++++++++++++++++++++++++++++++
'+	������ ���� Function, Procedure																		+
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'===========================================  2.3.1 OkClick()  ==========================================
'=	Name : OkClick()																					=
'=	Description : Return Array to Opener Window when OK button click									=
'========================================================================================================
Function OKClick()
	
	Dim intColCnt, intRowCnt, intInsRow, SData
	SData = 27

		If frm1.vspdData.SelModeSelCount > 0 Then 

			intInsRow = 0

			Redim arrReturn(frm1.vspdData.SelModeSelCount-1, frm1.vspdData.MaxCols - 1)

			For intRowCnt = 1 To frm1.vspdData.MaxRows
				frm1.vspdData.Row = intRowCnt
				If frm1.vspdData.SelModeSelected Then
					'For intColCnt = 0 To frm1.vspdData.MaxCols - 1
					For intColCnt = 1 To SData
					'frm1.vspdData.Col = GetKeyPos("A",intColCnt+1)
					frm1.vspdData.Col = intColCnt
					'arrReturn(intInsRow, intColCnt) = frm1.vspdData.Text
					arrReturn(intInsRow, intColCnt - 1) = frm1.vspdData.Text
					Next
					intInsRow = intInsRow + 1
				End IF								
			Next
		End if			

		Self.Returnvalue = arrReturn
		Self.Close()
End Function	


'=========================================  2.3.2 CancelClick()  ========================================
'=	Name : CancelClick()																				=
'=	Description : Return Array to Opener Window for Cancel button click 								=
'========================================================================================================
Function CancelClick()
	Redim arrReturn(1,1)
	arrReturn(0,0) = ""
	Self.Returnvalue = arrReturn
	Self.Close()
End Function
'*******************************************  2.4 POP-UP ó���Լ�  **************************************
'*	���: POP-UP																						*
'*	Description : POP-UP Call�ϴ� �Լ� �� Return Value setting ó��										*
'********************************************************************************************************
'===========================================  2.4.1 POP-UP Open �Լ�()  =================================
'=	Name : Open???()																					=
'=	Description : POP-UP Open																			=
'========================================================================================================
'------------------------------------------  OpenPoNo()  -------------------------------------------------
'	Name : OpenPoNo()
'	Description : Purchase_Order PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenStoNo()
	
	Dim strRet
	Dim iCalledAspName
	Dim IntRetCD
		
	If gblnWinEvent = True Or UCase(frm1.txtstoNo.className) = UCase(PopupParent.UCN_PROTECTED) Then Exit Function
		
	gblnWinEvent = True
	
	'strRet = window.showModalDialog("../m3/m3111pa1.asp", Array(PopupParent,arrParam), _
	'		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	iCalledAspName = AskPRAspName("M9111PA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", PopupParent.VB_INFORMATION, "M3111PA1", "X")
		gblnWinEvent = False
		Exit Function
	End If
	
	strRet = window.showModalDialog(iCalledAspName, Array(PopupParent,arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	gblnWinEvent = False

	If strRet(0) = "" Then
		Exit Function
	Else
		Call SetStoNo(strRet)
	End If

End Function

Function OpenSGICd()
	
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	arrParam(0) = "������"				
	arrParam(1) = "B_Biz_Partner"
	
	arrParam(2) = Trim(frm1.txtsgicd.Value)
	arrParam(3) = ""							
	
	'arrParam(4) = "Bp_Type in ('S','CS') AND usage_flag='Y'"	
	arrParam(4) = "Bp_Type <> " & FilterVar("C", "''", "S") & "  AND USAGE_FLAG = " & FilterVar("Y", "''", "S") & "  AND IN_OUT_FLAG = " & FilterVar("I", "''", "S") & " "	
	arrParam(5) = "������"				
	
    arrField(0) = "BP_CD"					
    arrField(1) = "BP_NM"					

	arrHeader(0) = "������"				
	arrHeader(1) = "�������"			
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetSGICd(arrRet)
	End If

End Function


'=======================================  2.4.2 POP-UP Return�� ���� �Լ�  ==============================
'=	Name : Set???()																						=
'=	Description : Reference �� POP-UP�� Return���� �޴� �κ�											=
'========================================================================================================
'+++++++++++++++++++++++++++++++++++++++++++  SetSoNo()  ++++++++++++++++++++++++++++++++++++++++++++++++
'+	Name : SetSoNo()																					+
'+	Description : Set Return array from SoNo PopUp Window												+
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Function SetSGICd(strRet)
	frm1.txtSGICd.value = strRet(0)
	frm1.txtSGINm.value = strRet(1)
End Function

Function OpenGroup()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	
	arrParam(0) = "���ű׷�"	
	arrParam(1) = "B_Pur_Grp"				
	
	arrParam(2) = Trim(frm1.txtGroupCd.Value)
'	arrParam(3) = Trim(frm1.txtGroupNm.Value)	
	
	arrParam(4) = ""			
	arrParam(5) = "���ű׷�"			
	
    arrField(0) = "PUR_GRP"	
    arrField(1) = "PUR_GRP_NM"	
    
    arrHeader(0) = "���ű׷�"		
    arrHeader(1) = "���ű׷��"		
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetGroup(arrRet)
	End If	

End Function 


Function SetGroup(byval arrRet)
	frm1.txtGroupCd.Value= arrRet(0)		
	frm1.txtGroupNm.Value= arrRet(1)		
End Function


'========================================================================================================
' Function Name : OpenSortPopup
' Function Desc : OpenSortPopup Reference Popup
'========================================================================================================
Function OpenSortPopup()
	Dim arrRet
	
	On Error Resume Next
	
	'If lgIsOpenPop = True Then Exit Function
	lgIsOpenPop = True

	arrRet = window.showModalDialog("../../ComAsp/ZADOSortPopup.asp",Array(ggoSpread.GetXMLData("A"),gMethodText),"dialogWidth=" & PopupParent.SORTW_WIDTH & "px; dialogHeight=" & PopupParent.SORTW_HEIGHT & "px;; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False
	
	If arrRet(0) = "X" Then
	   Exit Function
	Else
	   Call ggoSpread.SaveXMLData("A",arrRet(0),arrRet(1))
       Call InitVariables
       Call InitSpreadSheet()       
   End If
End Function

'=======================================  2.4.2 POP-UP Return�� ���� �Լ�  ==============================
'=	Name : Set???()																						=
'=	Description : Reference �� POP-UP�� Return���� �޴� �κ�											=
'========================================================================================================
'+++++++++++++++++++++++++++++++++++++++++++  SetSoNo()  ++++++++++++++++++++++++++++++++++++++++++++++++
'+	Name : SetSoNo()																					+
'+	Description : Set Return array from SoNo PopUp Window												+
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Function SetStoNo(strRet)
	frm1.txtStoNo.value = strRet(0)
End Function

	
'++++++++++++++++++++++++++++++++++++++++++  2.5 ������ ���� �Լ�  ++++++++++++++++++++++++++++++++++++++
'+	���� ���α׷����� �ʿ��� ������ ���� Procedure(Sub, Function, Validation & Calulation ���� �Լ�)	+
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
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
	'Call GetAdoFieldInf("m3112ra5","S","A")                                     ' G for Group , A for SpreadSheet No('A','B',....    	
	Call LoadInfTB19029				                                           '��: Load table , B_numeric_format
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
	Call ggoOper.LockField(Document, "N")						<% '��: Lock  Suitable  Field %>
	'Call MakePopData(gDefaultT,gFieldNM,gFieldCD,lgPopUpR,lgSortFieldNm,lgSortFieldCD,C_MaxSelList)    ' You must not this line
	Call InitVariables														    '��: Initializes local global variables
	Call SetDefaultVal	
	Call InitSpreadSheet()
	Call MM_preloadImages("../../../CShared/image/Query.gif","../../../CShared/image/OK.gif","../../../CShared/image/Cancel.gif")

	Call FncQuery()
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
'*********************************************  3.3 Object Tag ó��  ************************************
'*	Object���� �߻� �ϴ� Event ó��																		*
'********************************************************************************************************
'=========================================  3.3.1 vspdData_DblClick()  ==================================
'=	Event Name : vspdData_DblClick																		=
'=	Event Desc :																						=
'========================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)
	If Row = 0 Or Frm1.vspdData.MaxRows = 0 Then 
	     Exit Sub
	End If

	If frm1.vspdData.MaxRows > 0 Then
		If frm1.vspdData.ActiveRow = Row Or frm1.vspdData.ActiveRow > 0 Then
			Call OKClick
		End If
	End If
End Sub

'========================================  3.3.2 vspdData_KeyPress()  ===================================
'=	Event Name : vspdData_KeyPress																		=
'=	Event Desc :																						=
'========================================================================================================
Function vspdData_KeyPress(KeyAscii)
     On Error Resume Next
     If KeyAscii = 13 And frm1.vspdData.ActiveRow > 0 Then    'Frm1������ frm1���� 
        Call OKClick()
     ElseIf KeyAscii = 27 Then
        Call CancelClick()
     End If
End Function

'======================================  3.3.3 vspdData_TopLeftChange()  ================================
'=	Event Name : vspdData_TopLeftChange																	=
'=	Event Desc :																						=
'========================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft, ByVal OldTop, ByVal NewLeft, ByVal NewTop)
	
	If OldLeft <> NewLeft Then
	    Exit Sub
	End If
    

	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	    '��: ������ üũ	
		If lgPageNo <> "" Then		                                                    '���� Ű ���� ������ �� �̻� ��������ASP�� ȣ������ ���� 
			If DbQuery = False Then
				Exit Sub
			End if
		End If
	End If
End Sub

'==========================================================================================
'   Event Name : OCX_KeyDown()
'   Event Desc : 
'==========================================================================================
Sub txtFrStoDt_Keypress(KeyAscii)
	On Error Resume Next
	If KeyAscii = 27 Then
		Call CancelClick()
	Elseif KeyAscii = 13 Then
		Call FncQuery()
	End if
End Sub

Sub txtToStoDt_Keypress(KeyAscii)
	On Error Resume Next
	If KeyAscii = 27 Then
		Call CancelClick()
	Elseif KeyAscii = 13 Then
		Call FncQuery()
	End if
End Sub

Sub txtFrSGIDt_Keypress(KeyAscii)
	On Error Resume Next
	If KeyAscii = 27 Then
		Call CancelClick()
	Elseif KeyAscii = 13 Then
		Call FncQuery()
	End if
End Sub

Sub txtToSGIDt_Keypress(KeyAscii)
	On Error Resume Next
	If KeyAscii = 27 Then
		Call CancelClick()
	Elseif KeyAscii = 13 Then
		Call FncQuery()
	End if
End Sub

'==========================================================================================
'   Event Name : txtFrStoDt
'   Event Desc :
'==========================================================================================
Sub txtFrStoDt_DblClick(Button)
	if Button = 1 then
		frm1.txtFrStoDt.Action = 7
	End if
End Sub

'==========================================================================================
'   Event Name : txtToStoDt
'   Event Desc :
'==========================================================================================
Sub txtToStoDt_DblClick(Button)
	if Button = 1 then
		frm1.txtToStoDt.Action = 7
	End if
End Sub

Sub txtFrSGIDt_DblClick(Button)
	if Button = 1 then
		frm1.txtFrSGIDt.Action = 7
	End if
End Sub


Sub txtToSGIDt_DblClick(Button)
	if Button = 1 then
		frm1.txtToSGIDt.Action = 7
	End if
End Sub


'########################################################################################################
'#					     4. Common Function��															#
'########################################################################################################
'########################################################################################################
'#						5. Interface ��																	#
'########################################################################################################
'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================

Function FncQuery() 
    
    FncQuery = False                                                 
    
    Err.Clear                                                        

	'** ValidDateCheck(pObjFromDt, pObjToDt) : 'pObjToDt'�� 'pObjFromDt'���� ũ�ų� ���ƾ� �Ҷ� **
	'If ValidDateCheck(frm1.txtFrStoDt, frm1.txtToStoDt) = False Then Exit Function
    with frm1
		if (UniConvDateToYYYYMMDD(.txtFrStoDt.text,PopupParent.gDateFormat,"") > UniConvDateToYYYYMMDD(.txtToStoDt.text,PopupParent.gDateFormat,"")) And Trim(.txtFrStoDt.text) <> "" And Trim(.txtToStoDt.text) <> "" then	
			Call DisplayMsgBox("17a003", "X","������", "X")	
			.txtToStoDt.Focus()
			Exit Function
		End if   
	End with
	
	Call ggoOper.ClearField(Document, "2")							
	Call InitVariables												

	If Not chkField(Document, "1") Then				
		Exit Function
	End If

    If DbQuery = False Then Exit Function

    FncQuery = True									
        
End Function

'********************************************  5.1 DbQuery()  *******************************************
' Function Name : DbQuery																				*
' Function Desc : This function is data query and display												*
'********************************************************************************************************
Function DbQuery()
	
	Dim strVal
	
	Err.Clear															<%'��: Protect system from crashing%>

	DbQuery = False														<%'��: Processing is NG%>

    If LayerShowHide(1) = False Then
         Exit Function
    End If  

	If lgIntFlgMode = PopupParent.OPMD_UMODE Then
		strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001
		strVal = strVal & "&txtStoNo=" &	Trim(frm1.hdnStoNo.Value)
		strVal = strVal & "&txtFrStoDt=" & Trim(frm1.hdnFrStoDt.value)
		strVal = strVal & "&txtToStoDt=" & Trim(frm1.hdnToStoDt.value)
		strVal = strVal & "&txtSGICd=" &	Trim(frm1.hdnSGICd.Value)
		strVal = strVal & "&txtFrSGIDt=" & Trim(frm1.hdnFrSGIDt.value)
		strVal = strVal & "&txtToSGIDt=" & Trim(frm1.hdnToSGIDt.value)
		
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
	Else
		strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001
		strVal = strVal & "&txtStoNo=" & Trim(frm1.txtStoNo.value)
		strVal = strVal & "&txtFrStoDt=" & Trim(frm1.txtFrStoDt.text)
		strVal = strVal & "&txtToStoDt=" & Trim(frm1.txtToStoDt.text)
		strVal = strVal & "&txtSGICd=" & Trim(frm1.txtSGICd.value)
		strVal = strVal & "&txtFrSGIDt=" & Trim(frm1.txtFrSGIDt.text)
		strVal = strVal & "&txtToSGIDt=" & Trim(frm1.txtToSGIDt.text)
		
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		
	End if 

	'strVal = strVal & "&txtSupplier=" & Trim(frm1.hdnSupplierCd.value)
	strVal = strVal & "&txtGroup=" & Trim(frm1.txtGroupCd.value)
	
'--------- Developer Coding Part (End) ------------------------------------------------------------
    strVal =     strVal & "&lgPageNo="       & lgPageNo                          '��: Next key tag
        strVal =     strVal & "&lgMaxCount="     & C_SHEETMAXROWS_D                  '��: �ѹ��� �����ü� �ִ� ����Ÿ �Ǽ� 
	strVal =     strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")
    strVal =     strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
	strVal =     strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))

	Call RunMyBizASP(MyBizASP, strVal)								<%'��: �����Ͻ� ASP �� ���� %>

	DbQuery = True														<%'��: Processing is NG%>
End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================
Function DbQueryOk()														<%'��: ��ȸ ������ ������� %>

	lgIntFlgMode = PopupParent.OPMD_UMODE
	
	If frm1.vspdData.MaxRows > 0 Then
		frm1.vspdData.Focus
		frm1.vspdData.Row = 1	:	frm1.vspdData.SelModeSelected = True		
	Else
		frm1.txtStoNo.focus
	End If

End Function

</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
</HEAD>
<%
'########################################################################################################
'#						6. TAG ��																		#
'########################################################################################################
%>
<BODY SCROLL=NO TABINDEX="-1">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_20%>>
	<TR>
		<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
	</TR>
	<TR>
		<TD HEIGHT=20 WIDTH=100%>
			<FIELDSET CLASS="CLSFLD">
				<TABLE <%=LR_SPACE_TYPE_40%>>
					<TR>
						<TD CLASS="TD5" NOWRAP>����̵���û��ȣ</TD>
						<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtStoNo" SIZE=32 MAXLENGTH=18 ALT="���ֹ�ȣ" tag="1XNXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnORGCd1" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenStoNo()"><div style="Display:none"><input type="text" name=none></div></TD>
						<TD CLASS="TD5" NOWRAP>����̵���û��</TD>
						<TD CLASS="TD6" NOWRAP>
							<table cellspacing=0 cellpadding=0>
								<tr>
									<td NOWRAP>
										<script language =javascript src='./js/m9211ra1_fpDateTime1_txtFrStoDt.js'></script>
									</td>
									<td NOWRAP>~</td>
									<td NOWRAP>
										<script language =javascript src='./js/m9211ra1_fpDateTime1_txtToStoDt.js'></script>
									</td>
								<tr>
							</table>
						</TD>
					</TR>
					<TR>
						<TD CLASS="TD5" NOWRAP>������</TD>
						<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtSGiCd" SIZE=10 MAXLENGTH=18 ALT="������" tag="1XNXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnORGCd1" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenSGICd()">
						<INPUT TYPE=TEXT AlT="�������" ID="txtSGiNm" NAME="arrCond" tag="14X">
						</TD>
						
						<TD CLASS="TD5" NOWRAP>�����</TD>
						<TD CLASS="TD6" NOWRAP>
							<table cellspacing=0 cellpadding=0>
								<tr>
									<td NOWRAP>
										<script language =javascript src='./js/m9211ra1_fpDateTime1_txtFrSGiDt.js'></script>
									</td>
									<td NOWRAP>~</td>
									<td NOWRAP>
										<script language =javascript src='./js/m9211ra1_fpDateTime1_txtToSGiDt.js'></script>
									</td>
								<tr>
							</table>
						</TD>
					</TR>
					<TR>
						<TD CLASS=TD5 NOWRAP>���ű׷�</TD> 
						<TD CLASS=TD6 colspan=3 NOWRAP>
						<INPUT TYPE=TEXT AlT="���ű׷�" NAME="txtGroupCd" SIZE=10 MAXLENGTH=4 tag="1XNXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnGroup1Cd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenGroup()">
						<INPUT TYPE=TEXT AlT="���ű׷�" ID="txtGroupNm" NAME="arrCond" tag="14X"></TD>
						</TD>
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
					<TD HEIGHT=100% WIDTH=100% COLSPAN=4>
						<script language =javascript src='./js/m9211ra1_vspdData_vspdData.js'></script>
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
					<TD WIDTH=70% NOWRAP>     <IMG SRC="../../../CShared/image/query_d.gif"  Style="CURSOR: hand" ALT="Search" NAME="Search" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Query.gif',1)"  ONCLICK="FncQuery()"   ></IMG>
					<IMG SRC="../../../CShared/image/zpConfig_d.gif" Style="CURSOR: hand" ALT="Config" NAME="Config" OnClick="OpenSortPopup()"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/zpConfig.gif',1)" ></IMG></TD>
					<TD WIDTH=30% ALIGN=RIGHT><IMG SRC="../../../CShared/image/ok_d.gif"     Style="CURSOR: hand" ALT="OK"     NAME="pop1"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/OK.gif',1)"     ONCLICK="OkClick()"    ></IMG>
							                  <IMG SRC="../../../CShared/image/cancel_d.gif" Style="CURSOR: hand" ALT="CANCEL" NAME="pop2"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)" ONCLICK="CancelClick()"></IMG>					</TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=no noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>



<INPUT TYPE=HIDDEN NAME="hdnStoNo" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnFrStoDt" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnToStoDt" tag="14">

<INPUT TYPE=HIDDEN NAME="hdnSGICd" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnFrSGIDt" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnToSGIDt" tag="14">

<INPUT TYPE=HIDDEN NAME="hdnSupplierCd" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnGroupCd" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnClsflg" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnReleaseflg" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnRcptflg" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnRetflg" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnRefType" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnRcptType" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnIvType" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnIvFlg" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnPoType" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnPlant" tag="14">



</FORM>
<DIV ID="MousePT" NAME="MousePT" STYLE="visible:true">
	<IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm"></IFRAME>
</DIV>
</BODY>
</HTML>
