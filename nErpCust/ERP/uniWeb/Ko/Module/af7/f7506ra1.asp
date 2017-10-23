
<%@ LANGUAGE="VBSCRIPT" %>
<!--
'********************************************************************************************************
'*  1. Module Name          : Basis Architect															*
'*  2. Function Name        : Reference Popup Business Part												*
'*  3. Program ID           : a7506ra1.asp																			*
'*  4. Program Name         : ȸ�����-�����ݰ���-																			*
'*  5. Program Desc         : Reference Popup															*
'*  7. Modified date(First) : 2003/01/22																*
'*  8. Modified date(Last)  : 2003/01/22																*
'*  9. Modifier (First)     : Lee Nam Yo																*
'* 10. Modifier (Last)      : Lee Nam Yo															*
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

Const BIZ_PGM_ID 		= "f7506rb1.asp"                              '��: Biz Logic ASP Name

'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
Const C_SHEETMAXROWS_D  = 30                                          '��: Server���� �ѹ��� fetch�� �ִ� ����Ÿ �Ǽ� 
Const C_MaxKey          = 5					                          '��: SpreadSheet�� Ű�� ���� 

'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim lgIsOpenPop                                          

Dim lgSelectList                                         
Dim lgSelectListDT                                       


Dim lgSortFieldNm                                        
Dim lgSortFieldCD                                         

Dim lgMaxFieldCount

Dim lgPopUpR                                              

Dim lgKeyPos                                              
Dim lgKeyPosVal                                         
Dim lgCookValue 


Dim lgSaveRow 

Dim  lgQueryFlag
Dim  lgCode		

Dim  arrReturn
Dim  arrParent
Dim  arrParam		
		
Dim  IsOpenPop     

Dim  IsDocPop  

Const C_ArClsAmt = 10

' ���Ѱ��� �߰� 
Dim lgAuthBizAreaCd	' ����� 
Dim lgInternalCd	' ���κμ� 
Dim lgSubInternalCd	' ���κμ�(��������)
Dim lgAuthUsrID		' ���� 

	 '------ Set Parameters from Parent ASP ------ 
arrParent        = window.dialogArguments
Set PopupParent = arrParent(0)	 
arrParam		= arrParent(1)


	Dim strYear, strMonth, strDay, dtToday, EndDate, StartDate
		dtToday = "<%=GetSvrDate%>"
	Call PopupParent.ExtractDateFrom(dtToday, PopupParent.gServerDateFormat, PopupParent.gServerDateType, strYear, strMonth, strDay)

	EndDate = PopupParent.UniConvYYYYMMDDToDate(PopupParent.gDateFormat, strYear, strMonth, strDay)
	StartDate = PopupParent.UNIDateAdd("M", -1, EndDate, PopupParent.gDateFormat)

top.document.title = "û������"

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
    lgSaveRow        = 0
    
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
' Para : 1. Currency
'        2. I(Input),Q(Query),P(Print),B(Bacth)
'        3. "*" is for Common module
'           "A" is for Accounting
'           "I" is for Inventory
'           ...
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 

	<% Call loadInfTB19029A("Q", "A","NOCOOKIE", "RA") %>                                '��: 
	<% Call LoadBNumericFormatA("Q", "A", "NOCOOKIE", "RA") %>
	
		'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub
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
'==========================================  2.2.1 SetDefaultVal()  =====================================
'=	Name : SetDefaultVal()																				=
'=	Description : ȭ�� �ʱ�ȭ(���� Field�� ȭ���� �� �� Default���� ������� �ϴ� Field�� Setting)		=
'========================================================================================================

Sub SetDefaultVal()
	with frm1
		.txtSttlmentNoFr.value = arrParam(0)
		.txtSttlDtFr.text = StartDate
		.txtSttlDtTo.text = EndDate
	End With
End Sub

'========================================================================================================
' Name : InitSpreadSheet
' Desc : This method initializes spread sheet column property
'========================================================================================================
Sub InitSpreadSheet()
	frm1.vspddata.OperationMode = 3
    Call SetZAdoSpreadSheet("f7506RA1","S","A","V20021211",PopupParent.C_SORT_DBAGENT,frm1.vspdData, C_MaxKey, "X","X")
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
'========================================================================================================

	Function OKClick()
			Dim intColCnt, arrReturn
		Redim arrReturn(C_MaxKey - 1)
		
		If frm1.vspdData.MaxRows < 1 Then
		   Call CancelClick()
		   Exit Function
		End If
		frm1.vspddata.row = frm1.vspdData.ActiveRow
		if frm1.vspdData.ActiveRow > 0 Then
			For intColCnt = 0 To C_MaxKey - 1
				frm1.vspddata.col = GetKeyPos("A", intColCnt+1)
				arrReturn(intColCnt) = Trim(frm1.vspddata.text)
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
		Self.Close()
	End Function

 '******************************************  2.4 POP-UP ó���Լ�  ****************************************
'	���: ���� POP-UP
'   ����: ���� POP-UP�� ���� Open�� Include�Ѵ�. 
'	      �ϳ��� ASP���� Popup�� �ߺ��Ǹ� �ϳ��� ��ũ���� ����ϰ� �������� �������Ͽ� ����Ѵ�.
'********************************************************************************************************* 
'------------------------------------------  OpenBp()  ---------------------------------------
'	Name : OpenBp()
'	Description : OpenAp Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function OpenBp(Byval strCode, byval iWhere)
	Dim arrRet
	Dim arrParam(5)

	If IsOpenPop = True Then Exit Function
		
	IsOpenPop = True

	arrParam(0) = strCode								'  Code Condition
   	arrParam(1) = ""							' ä�ǰ� ����(�ŷ�ó ����)
	arrParam(2) = ""								'FrDt
	arrParam(3) = ""								'ToDt
	arrParam(4) = "T"							'B :���� S: ���� T: ��ü 
	arrParam(5) = ""									'SUP :����ó PAYTO: ����ó SOL:�ֹ�ó PAYER :����ó INV:���ݰ�� 	
	
	arrRet = window.showModalDialog("../../Module/ar/BpPopup.asp", Array(window.PopupParent,arrParam), _
		"dialogWidth=780px; dialogHeight=450px; : Yes; help: No; resizable: No; status: No;")
		
	
	IsOpenPop = False
	
	IF 	arrRet(0) <> "" then			
		Call SetPopUp(arrRet, iWhere)
	Else
		Call EscPopUp(iwhere)
		lgBlnFlgChgValue = True
	end if

End Function
 '------------------------------------------  OpenPopUp()  -------------------------------------------------
'	Name : OpenPopUp()
'	Description : Bp PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenPopUp(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	
	IsOpenPop = True
	Select Case iWhere
		Case 1	
			arrParam(0) = "�ŷ�ó�˾�"
			arrParam(1) = "B_BIZ_PARTNER"				
			arrParam(2) = Trim(strCode)
			arrParam(3) = ""
			arrParam(4) = ""
			arrParam(5) = "�ŷ�ó"			
	
			arrField(0) = "BP_CD"	
			arrField(1) = "BP_NM"	
    
			arrHeader(0) = "�ŷ�ó"		
			arrHeader(1) = "�ŷ�ó��"	
		Case 2,3	
			arrParam(0) = "û���ȣ"
			arrParam(1) = "F_PRRCPT_STTL"				
			arrParam(2) = Trim(strCode)
			arrParam(3) = ""
			arrParam(4) = ""
			arrParam(5) = "û���ȣ"			
	
			arrField(0) = "STTLMENT_NO"	
			arrField(1) = "PRRCPT_NO"	
    
			arrHeader(0) = "û���ȣ"		
			arrHeader(1) = "�����ݹ�ȣ"	
		Case 4,5
			arrParam(0) = "�����ݹ�ȣ"
			arrParam(1) = "F_PRRCPT_STTL"				
			arrParam(2) = Trim(strCode)
			arrParam(3) = ""
			arrParam(4) = ""
			arrParam(5) = "�����ݹ�ȣ"			
	
			arrField(0) = "PRRCPT_NO"	
			arrField(1) = "STTLMENT_NO"	
    
			arrHeader(0) = "�����ݹ�ȣ"		
			arrHeader(1) = "û���ȣ"			
			
	End Select    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	IF 	arrRet(0) <> "" then			
		Call SetPopUp(arrRet, iWhere)
	Else
		Call EScPopUp(iWhere)
	end if
End Function


'++++++++++++++++  Insert Your Code for PopUp(Open)  ++++++++++++++++++++++++++++++++++++++++++++++++++ 
  '------------------------------------------  SetBpCd()  --------------------------------------------------
'	Name : SetPopUp()
'	Description : Item Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function EScPopUp(Byval iWhere)
	
	Select Case iWhere
		Case 1
				frm1.txtBpCd.focus
		Case 2
				frm1.txtSttlmentNoFr.focus	
		Case 3
				frm1.txtSttlmentNoTo.focus
		Case 4
				frm1.txtPrRcptNoFr.focus	
		Case 5
				frm1.txtPrRcptNoTo.focus	
	End Select	
				
	lgBlnFlgChgValue = True
	
End Function
 '------------------------------------------  SetBpCd()  --------------------------------------------------
'	Name : SetPopUp()
'	Description : Item Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetPopUp(Byval arrRet, Byval iWhere)
	
	Select Case iWhere
		Case 1
				Frm1.txtBpCd.value = arrRet(0)		
				Frm1.txtBpNm.value = arrRet(1)
				frm1.txtBpCd.focus
		Case 2
				Frm1.txtSttlmentNoFr.value = arrRet(0)	
				frm1.txtSttlmentNoFr.focus	
		Case 3
				Frm1.txtSttlmentNoTo.value = arrRet(0)		
				frm1.txtSttlmentNoTo.focus
		Case 4
				Frm1.txtPrRcptNoFr.value = arrRet(0)	
				frm1.txtPrRcptNoFr.focus	
		Case 5
				Frm1.txtPrRcptNoTo.value = arrRet(0)	
				frm1.txtPrRcptNoTo.focus	
	End Select	
				
	lgBlnFlgChgValue = True
	
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
	Call LoadInfTB19029														
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	Call ggoOper.FormatField(Document, "1",PopupParent.ggStrIntegeralPart, PopupParent.ggStrDeciPointPart,PopupParent.gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec,,,PopupParent.ggStrMinPart,PopupParent.ggStrMaxPart)
	Call ggoOper.LockField(Document, "N")                                   

	Call InitVariables														
	Call SetDefaultVal
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
	
	Call InitVariables 		
	If PopupParent.CompareDateByFormat(Frm1.txtSttlDtFr.text,Frm1.txtSttlDtTo.text,Frm1.txtSttlDtFr.Alt,Frm1.txtSttlDtTo.Alt, _
	    	               "970025",Frm1.txtSttlDtFr.UserDefinedFormat,PopupParent.gComDateType, true) = False Then
	   Frm1.txtSttlDtFr.focus
	   Exit Function
	End If
	Call ggoOper.ClearField(Document, "2")						
	Frm1.vspdData.MaxRows = 0
	lgQueryFlag = "1"	
	lgCode = ""
	'-----------------------
	'Check condition area
	'-----------------------
	If Not chkField(Document, "1") Then									'This function check indispensable field
		Exit Function
	End If
		
	If DbQuery = False Then Exit Function
		 
	 FncQuery = True	
End Function


'*********************************************  3.3 Object Tag ó��  ************************************
'*	Object���� �߻� �ϴ� Event ó��																		*
'********************************************************************************************************
Sub txtSttlDtFr_KeyPress(KeyAscii)
	On Error Resume Next
	If KeyAscii = 27 Then 
		Call CancelClick()
	ElseIf KeyAscii = 13 Then 
		Call Fncquery()
	End IF
End Sub

Sub txtSttlDtTo_KeyPress(KeyAscii)
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
	Dim ii
	
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
	
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	lgCookValue = ""
	
	
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
    
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
	Call OKClick
End Sub


'########################################################################################################
'#					     4. Common Function��															#
'########################################################################################################

'=======================================================================================================
'   Event Name : txtSttlDtFr_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub  txtSttlDtFr_DblClick(Button)
    If Button = 1 Then
        Frm1.txtSttlDtFr.Action = 7                        
        Call SetFocusToDocument("P")
		Frm1.txtSttlDtFr.Focus
    End If
End Sub


'=======================================================================================================
'   Event Name : txtSttlDtTo_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub  txtSttlDtTo_DblClick(Button)
    If Button = 1 Then
        Frm1.txtSttlDtTo.Action = 7                        
        Call SetFocusToDocument("P")
		Frm1.txtSttlDtTo.Focus
    End If
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
			strVal = strVal & "?txtBpCd=" & Trim(.txtBpCd.Value)
			strVal = strVal & "&txtSttlmentNoFr=" & Trim(.txtSttlmentNoFr.value)
			strVal = strVal & "&txtSttlmentNoTo=" & Trim(.txtSttlmentNoTo.value)
			strVal = strVal & "&txtPrRcptNoFr=" & Trim(.txtPrRcptNoFr.value)
			strVal = strVal & "&txtPrRcptNoTo=" & Trim(.txtPrRcptNoTo.value)
			strVal = strVal & "&txtSttlDtFr=" & Trim(.txtSttlDtFr.Text)					'��: ��ȸ ���� ����Ÿ 
			strVal = strVal & "&txtSttlDtTo=" & Trim(.txtSttlDtTo.Text)					'��: ��ȸ ���� ����Ÿ 
			strVal = strVal & "&txtBpcd_Alt=" & Trim(.txtBpCd.alt)
	    Else
			strVal = strVal & "?txtBpCd=" & Trim(.htxtBpCd.value)
			strVal = strVal & "&txtSttlDtFr=" & Trim(.htxtSttlDtFr.value)					'��: ��ȸ ���� ����Ÿ 
			strVal = strVal & "&txtSttlDtTo=" & Trim(.htxtSttlDtTo.value)					'��: ��ȸ ���� ����Ÿ 
			strVal = strVal & "&txtBpcd_Alt=" & Trim(.txtBpCd.alt)
			strVal = strVal & "&txtSttlmentNoFr=" & Trim(.htxtSttlmentNoFr.value)
			strVal = strVal & "&txtSttlmentNoTo=" & Trim(.htxtSttlmentNoTo.value)
			strVal = strVal & "&txtPrRcptNoFr=" & Trim(.htxtPrRcptNoFr.value)
			strVal = strVal & "&txtPrRcptNoTo=" & Trim(.htxtPrRcptNoTo.value)
		
	    End If   
           
    '--------- Developer Coding Part (End) ------------------------------------------------------------
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
    lgSaveRow        = 1
	
	If frm1.vspdData.MaxRows > 0 Then
 		frm1.vspdData.Focus
 	End If
	
End Function
'===========================================================================
' Function Name : OpenOrderByPopup
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
						<TD CLASS=TD5 NOWRAP>û������</TD>
						<TD CLASS=TD6 NOWRAP>
							<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME="txtSttlDtFr" CLASS=FPDTYYYYMMDD tag="12" Title="FPDATETIME" ALT="��������"></OBJECT>');</SCRIPT>								
							&nbsp;~&nbsp;<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME="txtSttlDtTo" CLASS=FPDTYYYYMMDD tag="12" Title="FPDATETIME" ALT="��������" ></OBJECT>');</SCRIPT></TD>												
						<TD CLASS=TD5 NOWRAP>�ŷ�ó</TD>
						<TD CLASS=TD6 NOWRAP>
							<INPUT TYPE="Text" NAME="txtBpCd" SIZE=10 MAXLENGTH=10 tag="12NXXU" ALT="�ŷ�ó"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBizCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript: CALL OpenBP(frm1.txtBpCd.value,1)">
							<INPUT TYPE=TEXT NAME="txtBpNm" SIZE=20 MAXLENGTH=20 STYLE="TEXT-ALIGN: Left" tag="14" ALT="�ŷ�ó��">
						</TD>
					</TR>
					<TR>
						<TD CLASS=TD5 NOWRAP>û���ȣ</TD>
						<TD CLASS=TD6 NOWRAP>
							<INPUT TYPE="Text" NAME="txtSttlmentNoFr" SIZE=10 MAXLENGTH=10 tag="11NXXU" ALT="û���ȣ"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBizCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript: CALL OpenPopUp(frm1.txtSttlmentNoFr.value,2)"> ~
							<INPUT TYPE="Text" NAME="txtSttlmentNoTo" SIZE=10 MAXLENGTH=10 tag="11NXXU" ALT="û���ȣ"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBizCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript: CALL OpenPopUp(frm1.txtSttlmentNoTo.value,3)"> 
						</TD>
						<TD CLASS=TD5 NOWRAP>�����ݹ�ȣ</TD>
						<TD CLASS=TD6 NOWRAP>
							<INPUT TYPE="Text" NAME="txtPrRcptNoFr" SIZE=10 MAXLENGTH=10 tag="11NXXU" ALT="���ۼ��ޱݹ�ȣ"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBizCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript: CALL OpenPopUp(frm1.txtPrRcptNoFr.value,4)"> ~
							<INPUT TYPE="Text" NAME="txtPrRcptNoTo" SIZE=10 MAXLENGTH=10 tag="11NXXU" ALT="���ἱ�ޱݹ�ȣ"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBizCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript: CALL OpenPopUp(frm1.txtPrRcptNoTo.value,5)"> 
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
					<IMG SRC="../../../CShared/image/zpConfig_d.gif" Style="CURSOR: hand" ALT="Config" NAME=Config ONMOUSEOUT="javascript:MM_swapImgRestore()" ONMOUSEOVER="javascript:MM_swapImage(this.name,'','../../../CShared/image/zpConfig.gif',1)" ONCLICK="OpenOrderByPopup()"></IMG>
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
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP"  WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=no noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>

<INPUT TYPE=HIDDEN NAME="htxtBpCd"		tag="24">
<INPUT TYPE=HIDDEN NAME="htxtSttlDtFr"		tag="24">
<INPUT TYPE=HIDDEN NAME="htxtSttlDtTo"        tag="24">
<INPUT TYPE=HIDDEN NAME="htxtSttlmentNoFr"        tag="24">
<INPUT TYPE=HIDDEN NAME="htxtSttlmentNoTo"		tag="24">
<INPUT TYPE=HIDDEN NAME="htxtPrRcptNoFr"		tag="24">
<INPUT TYPE=HIDDEN NAME="htxtPrRcptNoTo"        tag="24">
				
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

