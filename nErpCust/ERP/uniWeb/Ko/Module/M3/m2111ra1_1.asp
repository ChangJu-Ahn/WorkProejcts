<%@ LANGUAGE="VBSCRIPT" %>
<%
'********************************************************************************************************
'*  1. Module Name          : Procurement																*
'*  2. Function Name        :																			*
'*  3. Program ID           : m2111ra1.asp																*
'*  4. Program Name         :																			*
'*  5. Program Desc         : Open Po Ref Popup ASP														*
'*  6. Comproxy List        : 																			*
'*  7. Modified date(First) : 2000/03/21																*
'*  8. Modified date(Last)  : 2001/05/08																*
'*                            2002/04/30
'*  9. Modifier (First)     : Shin jin hyun																*
'* 10. Modifier (Last)      : Min, HJ															*	
'*                            Kim Jae Soon
'* 11. Comment              :																			*
'* 12. Common Coding Guide  :																			*
'* 13. History              :																			*
'********************************************************************************************************
Response.Expires = -1													'�� : ASP�� ĳ������ �ʵ��� �Ѵ�.
%>
<HTML>
<HEAD>
<!--<TITLE>���ſ�û����</TITLE>-->
<TITLE></TITLE>
<%
'########################################################################################################
'#						1. �� �� ��																		#
'########################################################################################################
%>
<%
'********************************************  1.1 Inc ����  ********************************************
'*	Description : Inc. Include																			*
'********************************************************************************************************
%>
<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->
<!-- #Include file="../../inc/IncSvrVariables.inc" -->
<%
'============================================  1.1.1 Style Sheet  =======================================
'========================================================================================================
%>
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		<% '��: �ش� ��ġ�� ���� �޶���, ��� ��� %>
<%
'============================================  1.1.2 ���� Include  ======================================
'========================================================================================================
%>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentA.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs">   </SCRIPT>

<SCRIPT LANGUAGE="JavaScript" SRC="../../inc/incImage.js"></SCRIPT>

<Script Language="VBS">

Option Explicit					<% '��: indicates that All variables must be declared in advance %>
	
	
<%
'********************************************  1.2 Global ����/��� ����  *******************************
'*	Description : 1. Constant�� �ݵ�� �빮�� ǥ��														*
'********************************************************************************************************
%>

<%
'============================================  1.2.1 Global ��� ����  ==================================
'========================================================================================================
%>
	'��� �������� 
	Const C_PurGrp 			= 1
	Const C_PurGrpNm 		= 2															'��: Spread Sheet�� Column�� ��� 
	Const C_BpCd	 		= 3
	Const C_BpCdNm 			= 4
	Const C_ProCType		= 5
	Const C_ProCTypeNm		= 6
	
	'�ϴ� �������� 
	Const C_ReqNo 			= 1
	Const C_PlantCd 		= 2															'��: Spread Sheet�� Column�� ��� 
	Const C_PlantNm 		= 3
	Const C_ItemCd 			= 4
	Const C_ItemNm			= 5
	Const C_Spec			= 6
	Const C_Qty 			= 7
	Const C_Unit 			= 8
	Const C_DlvyDt 			= 9
	Const C_PlantDt 		= 10	
	Const C_ReqType			= 11
	Const C_ReqTypeNm		= 12
	Const C_SoNo			= 13	
	Const C_SoSeqNo			= 14
	Const C_TrackingNo		= 15	
	Const C_SLCd 			= 16
	Const C_SLNm 			= 17
	Const C_HSCd			= 18
	Const C_HSNm 			= 19
	
	'�̼��� �߰� 
	Const C_hUnderTot		= 20
	Const C_hOverTot		= 21	
	
	
    Const BIZ_PGM_ID 		= "m2111rb1_1.asp"                              '��: Biz Logic ASP Name
     
<%
'========================================================================================================
'=									1.2 Constant variables 
'========================================================================================================
%>
	Const C_SHEETMAXROWS_D  = 100                                          '��: Fetch max count at once
	Const C_MaxKey_1        = 6                                           '��: key count of SpreadSheet
	'�̼��� ���� 
	Const C_MaxKey			= 21
	'Const C_MaxKey          = 19                                           '��: key count of SpreadSheet
<%
'========================================================================================================
'=									1.3 Common variables 
'========================================================================================================
%>
<!-- #Include file="../../inc/lgvariables.inc" -->	
<%
'========================================================================================================
'=									1.4 User-defind Variables
'========================================================================================================
%>


Dim lgStrPrevKey_1			'�ι�° �׸��忡�� ���Ǵ� ���� 
Dim lgPageNo_1				'�ι�° �׸��忡�� ���Ǵ� ���� 
		
Dim lgSelectList                                            '��: SpreadSheet�� �ʱ�  ��ġ�������� ���� 
Dim lgSelectListDT                                          '��: SpreadSheet�� �ʱ�  ��ġ�������� ���� 

Dim lgSortFieldNm                                           '��: Orderby popup�� ����Ÿ(�ʵ弳��)      
Dim lgSortFieldCD                                           '��: Orderby popup�� ����Ÿ(�ʵ��ڵ�)      

Dim lgPopUpR                                                '��: Orderby default ��                    

Dim lgKeyPos                                                '��: Key��ġ                               
Dim lgKeyPosVal                                             '��: Key��ġ Value                         
Dim IscookieSplit 

Dim IsOpenPop  
Dim gblnWinEvent											'��: ShowModal Dialog(PopUp) 
														    'Window�� ���� �� �ߴ� ���� �����ϱ� ���� 
														    'PopUp Window�� ��������� ���θ� ��Ÿ�� 
Dim arrReturn												'��: Return Parameter Group
Dim arrParam
Dim arrParent
					
arrParent = window.dialogArguments
Set PopupParent = ArrParent(0)
top.document.title = PopupParent.gActivePRAspName

Dim iDBSYSDate
Dim EndDate, StartDate

iDBSYSDate = "<%=GetSvrDate%>"

'------ ��: �ʱ�ȭ�鿡 �ѷ����� ������ ��¥ ------
EndDate = UNIConvDateAtoB(iDBSYSDate, PopupParent.gServerDateFormat, PopupParent.gDateFormat)

'------ ��: �ʱ�ȭ�鿡 �ѷ����� ���� ��¥ ------
StartDate = UnIDateAdd("m", -1, EndDate, PopupParent.gDateFormat)

<%
'########################################################################################################
'#						2. Function ��																	#
'#																										#
'#	���� : �����ڰ� ������ �Լ�, �� Event���� �Լ��� ������ ��� ����� ���� �Լ� ���					#
'#	�������� ���� ���� : 1. Sub �Ǵ� Function�� ȣ���� �� �ݵ�� Call�� ����.							#
'#						 2. Sub, Function �̸��� _�� ���� �ʵ��� �Ѵ�. (Event�� �����ϱ� ����)			#
'########################################################################################################
%>
<% 
'*******************************************  2.1 ���� �ʱ�ȭ �Լ�  *************************************
'*	���: �����ʱ�ȭ																					*
'*	Description : Global���� ó��, �����ʱ�ȭ ���� �۾��� �Ѵ�.											*
'********************************************************************************************************
%>
<%
'==========================================  2.1.1 InitVariables()  =====================================
'=	Name : InitVariables()																				=
'=	Description : ���� �ʱ�ȭ(Global ����, �ʱ�ȭ�� �ʿ��� ���� �Ǵ� Flag���� Setting�Ѵ�.)				=
'========================================================================================================
%>
Function InitVariables()
		lgStrPrevKey     = ""								   'initializes Previous Key
		lgPageNo         = ""
		
		lgStrPrevKey_1     = ""								   'initializes Previous Key
		lgPageNo_1         = ""
        
        lgBlnFlgChgValue = False	                           'Indicates that no value changed
        
        lgIntFlgMode     = PopupParent.OPMD_CMODE                          'Indicates that current mode is Create mode
        frm1.vspdData.OperationMode  = 5
        frm1.vspdData1.OperationMode = 3
        
        lgSortKey        = 1   
        
        lgIntGrpCount = 0										<%'��: Initializes Group View Size%>

        gblnWinEvent = False
       
        Redim arrReturn(0,0)        
        Self.Returnvalue = arrReturn     
End Function

<%'=============================== 2.1.2 LoadInfTB19029() ========================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== %>
	Sub LoadInfTB19029()
		<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
		'------ Developer Coding part (Start ) -------------------------------------------------------------- 

		<% Call loadInfTB19029A("I", "*", "NOCOOKIE", "RA") %>                                '��: 

		'------ Developer Coding part (End )   -------------------------------------------------------------- 
	End Sub
<%
'*******************************************  2.2 ȭ�� �ʱ�ȭ �Լ�  *************************************
'*	���: ȭ���ʱ�ȭ																					*
'*	Description : ȭ���ʱ�ȭ, Combo Display, ȭ�� Clear �� ȭ�� �ʱ�ȭ �۾��� �Ѵ�.						*
'********************************************************************************************************
%>
 Sub InitComboBox()
	'-----------------------------------------------------------------------------------------------------
	' List Minor code for Procurement Type(���ޱ���)
	'-----------------------------------------------------------------------------------------------------
	if frm1.hdnSubcontraflg.value  = "N" then
			Call CommonQueryRs(" MINOR_CD, MINOR_NM ", " B_MINOR ", " MAJOR_CD = 'P1003' AND MINOR_CD = 'P' ORDER BY MINOR_CD ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
			Call SetCombo2(frm1.cboProcType, lgF0, lgF1, Chr(11))
	Elseif  frm1.hdnSubcontraflg.value ="Y" then
		Call CommonQueryRs(" MINOR_CD, MINOR_NM ", " B_MINOR ", " MAJOR_CD = 'P1003' AND MINOR_CD != 'P' ORDER BY MINOR_CD DESC", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
		Call SetCombo2(frm1.cboProcType, lgF0, lgF1, Chr(11))
	Else
		Call CommonQueryRs(" MINOR_CD, MINOR_NM ", " B_MINOR ", " MAJOR_CD = 'P1003' ORDER BY MINOR_CD ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
		Call SetCombo2(frm1.cboProcType, lgF0, lgF1, Chr(11))
	End if
End Sub
<%
'==========================================  2.2.3 InitSpreadSheet()  ===================================
'=	Name : InitSpreadSheet()																			=
'=	Description : This method initializes spread sheet column property									=
'========================================================================================================
%>
	Sub InitSpreadSheet()
		Call SetZAdoSpreadSheet("M2111RA1_1_2","S","B","V20030303",PopupParent.C_SORT_DBAGENT,frm1.vspdData1, _
									C_MaxKey_1, "X","X")

	
		Call SetZAdoSpreadSheet("M2111RA1_1_1","S","A","V20030303",PopupParent.C_SORT_DBAGENT,frm1.vspdData, _
									C_MAXKEY , "X","X")
		
		Call SetSpreadLock 
		'Call SetSpreadLock("A")
	End Sub


<%
'================================== 2.2.4 SetSpreadLock() ==================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
%>
	Sub SetSpreadLock()
		ggoSpread.Source = frm1.vspdData1
  	    ggoSpread.SpreadLockWithOddEvenRowColor()

	End Sub	

'================================== 2.2.4 SetSpreadLock() ==================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================

	Sub SetSpreadLock_1()
	
		ggoSpread.Source = frm1.vspdData
  	    ggoSpread.SpreadLockWithOddEvenRowColor()
	End Sub
<%
'++++++++++++++++++++++++++++++++++++++++++  2.3 ������ ���� �Լ�  ++++++++++++++++++++++++++++++++++++++
'+	������ ���� Function, Procedure																		+
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
%>
<%
'===========================================  2.3.1 OkClick()  ==========================================
'=	Name : OkClick()																					=
'=	Description : Return Array to Opener Window when OK button click									=
'========================================================================================================
%>	
	
	Function OKClick()
	
		Dim intColCnt, intRowCnt, intInsRow
		
		with frm1
		If .vspdData.SelModeSelCount > 0 Then 
			
			intInsRow = 0

			'Redim arrReturn(frm1.vspdData.SelModeSelCount-1, frm1.vspdData.MaxCols-2)
			Redim arrReturn(frm1.vspdData.SelModeSelCount, frm1.vspdData.MaxCols-2)
			For intRowCnt = 1 To frm1.vspdData.MaxRows

				frm1.vspdData.Row = intRowCnt
				
				If frm1.vspdData.SelModeSelected Then
					For intColCnt = 0 To frm1.vspdData.MaxCols - 2
						frm1.vspdData.Col = GetKeyPos("A",intColCnt+1)
						arrReturn(intInsRow, intColCnt) = frm1.vspdData.Text
					Next
										
					intInsRow = intInsRow + 1
				End IF								
			Next
			arrReturn(intInsRow, 0) = frm1.hdnSupplierCd.value
			arrReturn(intInsRow, 1) = frm1.hdnGroupCd.value
			arrReturn(intInsRow, 2) = frm1.hdnSubcontraflg.value
			arrReturn(intInsRow, 3) = frm1.hdnGroupNm.value
			
			
		End if		
		
		end with
		
		Self.Returnvalue = arrReturn
		Self.Close()
		
	End Function
<%
'=========================================  2.3.2 CancelClick()  ========================================
'=	Name : CancelClick()																				=
'=	Description : Return Array to Opener Window for Cancel button click 								=
'========================================================================================================
%>
	Function CancelClick()
		Self.Close()
	End Function
<%
'=========================================  2.3.3 Mouse Pointer ó�� �Լ� ===============================
'========================================================================================================
%>
	Function MousePointer(pstr1)
	      Select case UCase(pstr1)
	            case "PON"
					window.document.search.style.cursor = "wait"
	            case "POFF"
					window.document.search.style.cursor = ""
	      End Select
	End Function
	
<% 
'*******************************************  2.4 POP-UP ó���Լ�  **************************************
'*	���: POP-UP																						*
'*	Description : POP-UP Call�ϴ� �Լ� �� Return Value setting ó��										*
'********************************************************************************************************
%>

'===========================================  2.4.1 POP-UP Open �Լ�()  =================================
'=	Name : Open???()																					=
'=	Description : POP-UP Open																			=
'========================================================================================================
'------------------------------------------  OpenGroup()  -------------------------------------------------
'	Name : OpenGroup()
'	Description : 
'--------------------------------------------------------------------------------------------------------- %>
Function OpenGroup()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If gblnWinEvent = True Or UCase(frm1.txtGroupCd.className) = Ucase(PopupParent.UCN_PROTECTED) Then Exit Function

	gblnWinEvent = True

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
	
	gblnWinEvent = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetGroup(arrRet)
	End If	

End Function 


Function SetGroup(byval arrRet)
	frm1.txtGroupCd.Value= arrRet(0)		
	frm1.txtGroupNm.Value= arrRet(1)	
	'frm1.txtGroupCd.focus	
	Set gActiveElement = document.activeElement
	
End Function

'------------------------------------------  OpenPlant()  -------------------------------------------------
'	Name : OpenPlant()
'	Description : Plant PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenPlant()
	Dim arrRet,lgIsOpenPop
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	arrParam(0) = "����"	
	arrParam(1) = "B_PLANT"
	arrParam(2) = Trim(frm1.txtPlantCd.value)
	arrParam(3) = ""
	arrParam(4) = ""			
	arrParam(5) = "����"			
	
    arrField(0) = "PLANT_CD"	
    arrField(1) = "PLANT_NM"	
    
    arrHeader(0) = "����"		
    arrHeader(1) = "�����"		
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	lgIsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else	
		frm1.txtPlantCd.value = arrRet(0)
		frm1.txtPlantNm.value = arrRet(1)
		frm1.txtPlantCd.focus	
		Set gActiveElement = document.activeElement
	End If	
		
End Function	

'------------------------------------------  OpenSupplier()-------------------------------------------------
'	Name : OpenSupplier()
'	Description : Supplier PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenSupplier()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True  Or UCase(frm1.txtSupplierCd.className) = Ucase(PopupParent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "����ó"					
	arrParam(1) = "B_BIZ_PARTNER"				

	arrParam(2) = Trim(frm1.txtSupplierCd.Value)
	'arrParam(3) = Trim(frm1.txtSupplierNm.Value)
	
	arrParam(4) = "BP_TYPE In ('S','CS') And usage_flag='Y'"	
	arrParam(5) = "����ó"						
	
    arrField(0) = "BP_Cd"					
    arrField(1) = "BP_NM"					
    
    arrHeader(0) = "����ó"				
    arrHeader(1) = "����ó��"			
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetSupplier(arrRet)
	End If	
	
End Function

Function SetSupplier(byval arrRet)
	
	frm1.txtSupplierCd.Value    = arrRet(0)		
	frm1.txtSupplierNm.Value    = arrRet(1)		
	lgBlnFlgChgValue = True
	
End Function

'===========================================================================
' Function Name : OpenSoNo
' Function Desc : OpenSoNo Reference Popup
'===========================================================================
 Function OpenSoNo()

	Dim strRet
	Dim iCalledAspName
	Dim IntRetCD
	
	If IsOpenPop = True Then Exit Function
			
	IsOpenPop = True
		
'	strRet = window.showModalDialog("../s3/s3111pa1.asp", "", _
'		"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	iCalledAspName = AskPRAspName("S3111PA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", PopupParent.VB_INFORMATION, "S3111PA1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	strRet = window.showModalDialog(iCalledAspName, Array(PopupParent,""), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If strRet = "" Then
		Exit Function
	Else
		frm1.txtSoNo.value = strRet
	End If	

End Function
<%
'===========================================================================
' Function Name : OpenTrackingNo
' Function Desc : OpenTrackingNo Reference Popup
'===========================================================================
%>

Function OpenTrackingNo()
	Dim arrRet
	Dim arrParam(5)
	Dim iCalledAspName
	Dim IntRetCD
	
	If IsOpenPop = True Then Exit Function
	
	IsOpenPop = True

	arrParam(0) = ""	'�ֹ�ó 
	arrParam(1) = ""	'�����׷� 
    arrParam(2) = ""	'���� 
    arrParam(3) = ""	'��ǰ�� 
    arrParam(4) = ""	'���ֹ�ȣ 
    arrParam(5) = ""	'�߰� Where�� 
    
'	arrRet = window.showModalDialog("../s3/s3135pa1.asp", Array(arrParam), _
'			"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	iCalledAspName = AskPRAspName("S3135PA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", PopupParent.VB_INFORMATION, "S3135PA1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(PopupParent,arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
    
	IsOpenPop = False

	If arrRet = "" Then
		Exit Function
	Else
		frm1.txtTrackingNo.Value = Trim(arrRet)
		lgBlnFlgChgValue = True
	End If	

End Function
 
<%
'=======================================  2.4.2 POP-UP Return�� ���� �Լ�  ==============================
'=	Name : Set???()																						=
'=	Description : Reference �� POP-UP�� Return���� �޴� �κ�											=
'========================================================================================================
%>
'========================================================================================================
' Function Name : OpenOrderBy
' Function Desc : OpenOrderBy Reference Popup
'========================================================================================================
Function OpenOrderBy()
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


<% '++++++++++++++++  Insert Your Code for PopUp(Open)  ++++++++++++++++++++++++++++++++++++++++++++++++++ %>
<% '------------------------------------------  SetSorgCode()  --------------------------------------------------
'	Name : SetBPCd()
'	Description : SetSorgCode Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- %>

<%
'++++++++++++++++++++++++++++++++++++++++++  2.5 ������ ���� �Լ�  ++++++++++++++++++++++++++++++++++++++
'+	���� ���α׷����� �ʿ��� ������ ���� Procedure(Sub, Function, Validation & Calulation ���� �Լ�)	+
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
%>

<%
'########################################################################################################
'#						3. Event ��																		#
'#	���: Event �Լ��� ���� ó��																		#
'#	����: Windowó��, Singleó��, Gridó�� �۾�.														#
'#		  ���⼭ Validation Check, Calcuration �۾��� ������ Event�� �߻�.								#
'#		  �� Object������ Grouping�Ѵ�.																	#
'########################################################################################################
%>
<%
'********************************************  3.1 Windowó��  ******************************************
'*	Window�� �߻� �ϴ� ��� Even ó��																	*
'********************************************************************************************************
%>
<%
'=========================================  3.1.1 Form_Load()  ==========================================
'=	Name : Form_Load()																					=
'=	Description : Window Load�� �����ʱ�ȭ �� ȭ���ʱ�ȭ�� �ϱ� ���� �Լ��� Call�ϴ� �κ�				=
'========================================================================================================
%>
Sub Form_Load()

'parent.msgbox "aaa"

    Call LoadInfTB19029													'��: Load table , B_numeric_format
'    ReDim lgPopUpR(C_MaxSelList - 1,1)
    
	'Call GetAdoFieldInf("M2111RA1_1","S","A")			              '��: spread sheet �ʵ����� query
	'
                                                                  ' 1. Program id
                                                                  ' 2. G is for Qroup , S is for Sort     
                                                                  ' 3. Spreadsheet no     
    'Html���� tag ���ڰ� 1�� 2�� �����ϴ� �κ� ����Format
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
	
	Call ggoOper.LockField(Document, "N")                         '��: Lock  Suitable  Field
    
'    Call MakePopData(gDefaultT,gFieldNM,gFieldCD,lgPopUpR,lgSortFieldNm,lgSortFieldCD,C_MaxSelList)    ' You must not this line    
    Call InitVariables											  '��: Initializes local global variables
	Call SetDefaultVal	
	Call InitComboBox
	Call InitSpreadSheet()
	Call MM_preloadImages("../../../CShared/image/Query.gif","../../../CShared/image/OK.gif","../../../CShared/image/Cancel.gif")
	Call FncQuery()
End Sub

Sub SetDefaultVal()
		Dim arrParam
		
		arrParam = arrParent(1)
		
		frm1.vspdData1.OperationMode = 3 
		frm1.vspdData.OperationMode = 5
		
		frm1.txtSupplierCd.value 	= arrParam(0)
		frm1.txtSupplierNm.value 	= arrParam(1)
		frm1.txtGroupCd.value 		= arrParam(2)
	'	msgbox PopupParent.gPurGrp
		If arrParam(2) = "" then
			frm1.txtGroupCd.value = PopupParent.gPurGrp
		End if

		frm1.txtGroupNm.value 		= arrParam(3)
		
		frm1.hdnSubcontraflg.value 	= arrParam(4)
		'frm1.hdnSubcontraflg.value 	= arrParam(4)
		
		
	'	If ubound(arrParam) = 5 then		'2002-12-04(LJT)
	'		frm1.hdnSTOflg.value = arrParam(5)
	'	Else 
	'		frm1.hdnSTOflg.value = "N"
	'	End If
		
		If arrParam(0) <> "" then		'2002-12-04(LJT)
			ggoOper.SetReqAttr		frm1.txtGroupCd, "Q"
			ggoOper.SetReqAttr		frm1.txtGroupNm, "Q"
		End if
		
		if  arrParam(2) <> "" then
			ggoOper.SetReqAttr		frm1.txtSupplierCd, "Q"
			ggoOper.SetReqAttr		frm1.txtSupplierNm, "Q"
		End if
		'	ggoOper.SetReqAttr		frm1.cboProcType, "Q"
			'ggoOper.SetReqAttr		frm1.txtGroupCd, "Q"
		
		
		frm1.txtFrPoDt.text 	= UnIDateAdd("d", -15, EndDate, PopupParent.gDateFormat)
		frm1.txtToPoDt.text 	= UnIDateAdd("d", +15, EndDate, PopupParent.gDateFormat)
		
		frm1.txtFrDlvyDt.text 	= EndDate
		frm1.txtToDlvyDt.text 	= UnIDateAdd("m", +1, EndDate, PopupParent.gDateFormat)
		
		' Tracker No.9743 �����ڵ� ���� - 2005.07.22 =========================================
		frm1.txtPlantCd.value=PopupParent.gPlant
		frm1.txtPlantNm.value=PopupParent.gPlantNm
		' Tracker No.9743 �����ڵ� ���� - 2005.07.22 =========================================		
		
End Sub

<%
'=========================================  3.1.2 Form_QueryUnload()  ===================================
'   Event Name : Form_QueryUnload																		=
'   Event Desc :																						=
'========================================================================================================
%>
	Sub Form_QueryUnload(Cancel, UnloadMode)
	   
	End Sub
<%
'*********************************************  3.2 Tag ó��  *******************************************
'*	Document�� TAG���� �߻� �ϴ� Event ó��																*
'*	Event�� ��� �Ʒ��� ����� Event�̿��� ����� �����ϸ� �ʿ�� �߰� �����ϳ�							*
'*	Event�� �浹�� ����Ͽ� �ۼ��Ѵ�.																	*
'********************************************************************************************************
%>



<%
'==========================================================================================
'   Event Name : OCX_Keypress()
'   Event Desc : 
'==========================================================================================
%>
	Sub txtFrPoDt_Keypress(KeyAscii)
		On Error Resume Next
		If KeyAscii = 27 Then
			Call CancelClick()
		Elseif KeyAscii = 13 Then
			Call FncQuery()
		End if
	End Sub

	Sub txtToPoDt_Keypress(KeyAscii)
		On Error Resume Next
		If KeyAscii = 27 Then
			Call CancelClick()
		Elseif KeyAscii = 13 Then
			Call FncQuery()
		End if
	End Sub

	Sub txtFrDlvyDt_Keypress(KeyAscii)
		On Error Resume Next
		If KeyAscii = 27 Then
			Call CancelClick()
		Elseif KeyAscii = 13 Then
			Call FncQuery()
		End if
	End Sub

	Sub txtToDlvyDt_Keypress(KeyAscii)
		On Error Resume Next
		If KeyAscii = 27 Then
			Call CancelClick()
		Elseif KeyAscii = 13 Then
			Call FncQuery()
		End if
	End Sub
<%
'==========================================================================================
'   Event Name : txtFrPoDt
'   Event Desc :
'==========================================================================================
%>
Sub txtFrPoDt_DblClick(Button)
	if Button = 1 then
		frm1.txtFrPoDt.Action = 7
	End if
End Sub

<%
'==========================================================================================
'   Event Name : txtToPoDt
'   Event Desc :
'==========================================================================================
%>
Sub txtToPoDt_DblClick(Button)
	if Button = 1 then
		frm1.txtToPoDt.Action = 7
	End if
End Sub

<%
'==========================================================================================
'   Event Name : txtFrDlvyDt
'   Event Desc :
'==========================================================================================
%>
Sub txtFrDlvyDt_DblClick(Button)
	if Button = 1 then
		frm1.txtFrDlvyDt.Action = 7
	End if
End Sub

<%
'==========================================================================================
'   Event Name : txtToDlvyDt
'   Event Desc :
'==========================================================================================
%>
Sub txtToDlvyDt_DblClick(Button)
	if Button = 1 then
		frm1.txtToDlvyDt.Action = 7
	End if
End Sub

<%
'*********************************************  3.3 Object Tag ó��  ************************************
'*	Object���� �߻� �ϴ� Event ó��																		*
'********************************************************************************************************
%>
	Function vspdData_DblClick(ByVal Col, ByVal Row)
	
	 If Row = 0 Or Frm1.vspdData.MaxRows = 0 Then 
          Exit Function
     End If
	With frm1.vspdData 
		If .MaxRows > 0 Then
			If .ActiveRow = Row Or .ActiveRow > 0 Then
				Call OKClick
			End If
		End If
	End With
	End Function
'========================================================================================
' Function Name : vspdData_Click
' Function Desc : 
'========================================================================================
Sub vspdData1_Click(ByVal Col, ByVal Row)
	
	Dim iPrevRows
	Dim strPurGrp, strPurNm, strBpCd, strProcureType, strVal
	ggoSpread.Source = frm1.vspdData1
	gMouseClickStatus = "SPC"   
	
	frm1.vspdData.MaxRows = 0
	
	Set gActiveSpdSheet = frm1.vspdData1
	Call SetPopupMenuItemInf("0000111111")

	If frm1.vspdData1.MaxRows = 0 Then Exit Sub
'	If Row = IgPrevRow Then Exit Sub
	
	If Row <= 0 Then
		If lgSortKey = 1 Then
			ggoSpread.SSSort Col				'Sort in Ascending
			lgSortkey = 2
		Else
			ggoSpread.SSSort Col, lgSortKey		'Sort in Descending
			lgSortkey = 1
		End If
	Else
 		'------ Developer Coding part (Start)
 		
		frm1.vspdData1.Row = row
		
		frm1.vspdData1.Col = C_PurGrp 	
		strPurGrp = frm1.vspdData1.text
		frm1.hdnGroupCd.value = strPurGrp
		
		frm1.vspdData1.Col = C_PurGrpNm 	
		strPurNm = frm1.vspdData1.text
		frm1.hdnGroupNm.value = strPurNm
		
		frm1.vspdData1.Col = C_BpCd 	
		strBpCd = frm1.vspdData1.text
		frm1.hdnSupplierCd.value = strBpCd
		
		frm1.vspdData1.Col = C_ProCType 	
		strProcureType = frm1.vspdData1.text
		frm1.hdnProcuType.value = strProcureType
		
		
		'�̼��� �߰� 
		lgPageNo = ""
			
		If DbQuery2(strPurGrp,strBpCd,strProcureType) = False Then
			'	Call ResetToolBar(lgOldRow)
				Exit Sub 
		End If	
	 	'------ Developer Coding part (End)
 	End If
End Sub


Function vspdData_KeyPress(KeyAscii)
	On Error Resume Next

	If KeyAscii = 13 And frm1.vspdData.ActiveRow > 0 Then
		Call OKClick()
	ElseIf KeyAscii = 27 Then
		Call CancelClick()
	End If
End Function

'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
	Sub vspdData1_TopLeftChange(ByVal OldLeft, ByVal OldTop, ByVal NewLeft, ByVal NewTop)
	
		If OldLeft <> NewLeft Then
		    Exit Sub
		End If		

		If frm1.vspdData1.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData1,NewTop) Then	    '��: ������ üũ	
			If lgPageNo_1 <> "" Then		                                                    '���� Ű ���� ������ �� �̻� ��������ASP�� ȣ������ ���� 
				If DbQuery = False Then
					Exit Sub
				End if
			End If
		End If		 
	End Sub

'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
	Sub vspdData_TopLeftChange(ByVal OldLeft, ByVal OldTop, ByVal NewLeft, ByVal NewTop)
		Dim purGrp, bpCd, procType

		purGrp  = frm1.hdnGroupCd.value
		bpCd	= frm1.hdnSupplierCd.value
		procType = frm1.hdnProcuType.value
		If OldLeft <> NewLeft Then
		    Exit Sub
		End If		
		

		If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	    '��: ������ üũ	
			If lgPageNo <> "" Then                '���� Ű ���� ������ �� �̻� ��������ASP�� ȣ������ ���� 
				If DbQuery2(purGrp,bpCd,procType) = False Then
					Exit Sub
				End if
			End If
		End If		 
	End Sub
<% '#########################################################################################################
'												4. Common Function�� 
'	���: Common Function
'######################################################################################################### %>

<% '#########################################################################################################
'												5. Interface�� 
'	���: Interface
'	����: ������ Toolbar�� ���� ó���� ���Ѵ�. 
'	      Toolbar�� ��ġ������� ����ϴ� ������ �Ѵ�. 
'	<< ���뺯�� ���� �κ� >>
' 	���뺯�� : Global Variables�� �ƴ����� ������ Sub�� Function���� ���� ����ϴ� ������ �������� 
'				�����ϵ��� �Ѵ�.
' 	1. ������Ʈ���� Call�ϴ� ���� 
'    	   ADF (ADS, ADC, ADF�� �״�� ���)
'    	   - ADF�� Set�ϰ� ����� �� �ٷ� Nothing �ϵ��� �Ѵ�.
' 	2. ������Ʈ�ѿ��� Return�� ���� �޴� ���� 
'    		strRetMsg
'######################################################################################################### %>

<% '*******************************  5.1 Toolbar(Main)���� ȣ��Ǵ� Function *******************************
'	���� : Fnc�Լ��� ���� �����ϴ� ��� Function
'********************************************************************************************************* %>
<%
'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================
%>
Function FncQuery() 

	Dim strPlant
    
    FncQuery = False                                                        '��: Processing is NG
    
    Err.Clear                                                               '��: Protect system from crashing
	
	'** ValidDateCheck(pObjFromDt, pObjToDt) : 'pObjToDt'�� 'pObjFromDt'���� ũ�ų� ���ƾ� �Ҷ� **
	If ValidDateCheck(frm1.txtFrPoDt, frm1.txtToPoDt) = False Then Exit Function
	If ValidDateCheck(frm1.txtFrDlvyDt, frm1.txtToDlvyDt) = False Then Exit Function
   
    '-----------------------
    'Erase contents area
    '-----------------------
    'Call ggoOper.ClearField(Document, "2")	         						'��: Clear Contents  Field
    Call InitVariables 														'��: Initializes local global variables
    
    ggoSpread.Source = frm1.vspdData	'###�׸��� ������ ���Ǻκ�###
    ggoSpread.ClearSpreadData

	ggoSpread.Source = frm1.vspdData1
    ggoSpread.ClearSpreadData
    
	'�̼��� �߰� PLANT
	strPlant = frm1.txtPlantCd.value	
	frm1.hdnPlantCd.value = strPlant    
    
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then								'��: This function check indispensable field
       Exit Function
    End If

    '-----------------------
    'Query function call area
    '-----------------------	
	If DbQuery = False Then Exit Function									

    FncQuery = True		
    
End Function
	
<%
'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
%>
Function DbQuery() 

	Err.Clear														'��: Protect system from crashing
	DbQuery = False													'��: Processing is NG
	
	If LayerShowHide(1) = False Then
		Exit Function
	End If
	
	Dim strVal
    
    With frm1
		If lgIntFlgMode = PopupParent.OPMD_UMODE Then		
			strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001					'��: �����Ͻ� ó�� ASP�� ����	
			strVal = strVal & "&txtFrPoDt=" & .hdnFrDt.value
			strVal = strVal & "&txtToPoDt=" & .hdnToDt.value
			strVal = strVal & "&txtFrDlvyDt=" & .hdnFrDt2.value
			strVal = strVal & "&txtToDlvyDt=" & .hdnToDt2.value		
			strVal = strVal & "&txtSoNo=" & .hdnSoNo.value
			strVal = strVal & "&txtTrackingNo=" & .hdnTrackingNo.value		
			strVal = strVal & "&txtSupplier=" & .hdnSupplierCd.value
			strVal = strVal & "&txtGroup=" & .hdnGroupCd.value
			strVal = strVal & "&txtProcure=" & .hdnProcuType.value 
			strVal = strVal & "&txtSubconfraflg=" & .hdnSubcontraflg.value
			strVal = strVal & "&txtSTOflg=" & .hdnSTOflg.value				'2002-12-04(LJT)
			'�̼��� 
			strVal = strVal & "&txtPlantCd=" & .hdnPlantCd.value
						
			strVal = strVal & "&lgStrPrevKey="   & lgStrPrevKey   
		Else
			strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001			
			strVal = strVal & "&txtFrPoDt=" & Trim(.txtFrPoDt.text)
			strVal = strVal & "&txtToPoDt=" & Trim(.txtToPoDt.text)
			strVal = strVal & "&txtFrDlvyDt=" & Trim(.txtFrDlvyDt.text)
			strVal = strVal & "&txtToDlvyDt=" & Trim(.txtToDlvyDt.text)
			strVal = strVal & "&txtSoNo=" & Trim(.txtSoNo.value)
			strVal = strVal & "&txtTrackingNo=" & Trim(.txtTrackingNo.value)
			strVal = strVal & "&txtSupplier=" & Trim(.txtSupplierCd.value)
			strVal = strVal & "&txtGroup=" & Trim(.txtGroupCd.value)
			strVal = strVal & "&txtProcure=" & Trim(.cboProcType.value )
			strVal = strVal & "&txtSubconfraflg=" & .hdnSubcontraflg.value
			strVal = strVal & "&txtSTOflg=" & .hdnSTOflg.value				'2002-12-04(LJT)
			'�̼��� 
			strVal = strVal & "&txtPlantCd=" & .hdnPlantCd.value			
			strVal = strVal & "&lgStrPrevKey="   & lgStrPrevKey
		End If				
	    strVal = strVal & "&lgPageNo="		 & lgPageNo_1						'��: Next key tag 
        strVal = strVal & "&lgMaxCount="     & C_SHEETMAXROWS_D             '��: �ѹ��� �����ü� �ִ� ����Ÿ �Ǽ�  
		strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("B")
		
        strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("B")
		strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("B"))
		strVal = strVal & "&txtGridNum="	 & "B"
		
		Call RunMyBizASP(MyBizASP, strVal)		    						'��: �����Ͻ� ASP �� ���� 
		
    End With
    
    DbQuery = True    

End Function

<%
'=========================================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'=========================================================================================================
%>
Function DbQueryOk()	    												'��: ��ȸ ������ ������� 
	Dim lRow, i, strPurGrp, strBpCd, strProcuType 
		
	'lgIntFlgMode = PopupParent.OPMD_UMODE
	
	If frm1.vspdData1.MaxRows > 0 Then
		frm1.vspdData1.Focus
		frm1.vspdData1.Row = 1	
		
		frm1.vspdData1.col = C_PurGrp
		strPurGrp = frm1.vspdData1.value
		frm1.vspdData1.col = C_BpCd 
		strBpCd = frm1.vspdData1.value
		frm1.vspdData1.col = C_ProCType 
		strProcuType = frm1.vspdData1.value
		
		
	
		frm1.hdnGroupCd.value = strPurGrp
		frm1.hdnSupplierCd.value = strBpCd
		frm1.hdnProcuType.value = strProcuType
		
		If lgIntFlgMode <> PopupParent.OPMD_UMODE Then
			Call DbQuery2(strPurGrp,strBpCd,strProcuType)
			lgIntFlgMode = PopupParent.OPMD_UMODE
		End If
		
		frm1.vspdData1.SelModeSelected = True		
	Else
	'	frm1.txtDnType.focus
	End If
	
	call SetSpreadLock

End Function
'=======================================================================================================
' Function Name : DbQuery2																				
' Function Desc : This function is data query and display												
'=======================================================================================================
Function DbQuery2(ByVal purGrp, ByVal bpCd, ByVal procType)
	Err.Clear														'��: Protect system from crashing
	DbQuery2 = False													'��: Processing is NG
	
	If LayerShowHide(1) = False Then
		Exit Function
	End If
	
	Dim strVal
	
	'frm1.vspdData.MaxRows = 0
	
    With frm1
		If lgIntFlgMode = PopupParent.OPMD_UMODE Then		
			strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001					'��: �����Ͻ� ó�� ASP�� ����	
			strVal = strVal & "&txtFrPoDt=" & .hdnFrDt.value
			strVal = strVal & "&txtToPoDt=" & .hdnToDt.value
			strVal = strVal & "&txtFrDlvyDt=" & .hdnFrDt2.value
			strVal = strVal & "&txtToDlvyDt=" & .hdnToDt2.value		
			strVal = strVal & "&txtSoNo=" & .hdnSoNo.value
			strVal = strVal & "&txtTrackingNo=" & .hdnTrackingNo.value		
			strVal = strVal & "&txtSupplier=" & bpCd
			strVal = strVal & "&txtGroup=" & purGrp
			strVal = strVal & "&txtProcure=" & procType
			strVal = strVal & "&txtSubconfraflg=" & .hdnSubcontraflg.value
			strVal = strVal & "&txtSTOflg=" & .hdnSTOflg.value				'2002-12-04(LJT)
			'�̼��� 
			strVal = strVal & "&txtPlantCd=" & .hdnPlantCd.value
			
			strVal = strVal & "&lgStrPrevKey="   & lgStrPrevKey   
		Else
			strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001			
			strVal = strVal & "&txtFrPoDt=" & Trim(.txtFrPoDt.text)
			strVal = strVal & "&txtToPoDt=" & Trim(.txtToPoDt.text)
			strVal = strVal & "&txtFrDlvyDt=" & Trim(.txtFrDlvyDt.text)
			strVal = strVal & "&txtToDlvyDt=" & Trim(.txtToDlvyDt.text)
			strVal = strVal & "&txtSoNo=" & Trim(.txtSoNo.value)
			strVal = strVal & "&txtTrackingNo=" & Trim(.txtTrackingNo.value)
			strVal = strVal & "&txtSupplier=" & bpCd
			strVal = strVal & "&txtGroup=" & purGrp
			strVal = strVal & "&txtProcure=" & procType
			strVal = strVal & "&txtSubconfraflg=" & .hdnSubcontraflg.value
			strVal = strVal & "&txtSTOflg=" & .hdnSTOflg.value				'2002-12-04(LJT)
			'�̼��� 
			strVal = strVal & "&txtPlantCd=" & .hdnPlantCd.value
			
			strVal = strVal & "&lgStrPrevKey="   & lgStrPrevKey
		End If				
	    strVal = strVal & "&lgPageNo="		 & lgPageNo						'��: Next key tag 
        strVal = strVal & "&lgMaxCount="     & C_SHEETMAXROWS_D             '��: �ѹ��� �����ü� �ִ� ����Ÿ �Ǽ�  
		strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")
		
        strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
		strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))
		strVal = strVal & "&txtGridNum="	 & "A"
		
		Call RunMyBizASP(MyBizASP, strVal)		    						'��: �����Ͻ� ASP �� ���� 
        
    End With
    
    DbQuery2 = True    
End Function

Function DbQuery2Ok()
	DbQuery2Ok = False
	call SetSpreadLock_1
	DbQuery2Ok = true
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
						<TD CLASS="TD5" NOWRAP>���ޱ���</TD>
						<TD CLASS="TD6"><SELECT NAME="cboProcType" ALT="���ޱ���" STYLE="Width: 168px;" ></SELECT></TD>
						<TD CLASS="TD5" NOWRAP>���ű׷�</TD>
						<TD CLASS="TD6">
						<INPUT TYPE=TEXT AlT="���ű׷�" NAME="txtGroupCd" SIZE=10 MAXLENGTH=4 tag="1XNXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnGroup1Cd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenGroup()">
						<INPUT TYPE=TEXT AlT="���ű׷�" ID="txtGroupNm" NAME="arrCond" tag="14X"></TD>
						</TD>
					</TR>
					<TR>
						<TD CLASS="TD5" NOWRAP>����ó</TD>
						<TD CLASS="TD6">
						<INPUT TYPE=TEXT NAME="txtSupplierCd" SIZE=10 MAXLENGTH=10 ALT="����ó" tag="1XNXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenSupplier()">
						<INPUT TYPE=TEXT AlT="����ó" ID="txtSupplierNm" tag="14X">
						</TD>
						<TD CLASS="TD5" NOWRAP>����</TD>
						<TD CLASS="TD6">
						<INPUT TYPE=TEXT NAME="txtPlantCd" SIZE=10 MAXLENGTH=4 ALT="����" tag="1XNXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPlant()">
						<INPUT TYPE=TEXT AlT="����" ID="txtPlantNm" tag="14X">
						</TD>
					</TR>	
					<TR>
						<TD CLASS="TD5" NOWRAP>���ֿ�����</TD>
						<TD CLASS="TD6" NOWRAP>
							<table cellpadding=0 cellspacing=0>
								<tr>
									<td NOWRAP>
										<script language =javascript src='./js/m2111ra1_1_fpDateTime1_txtFrPoDt.js'></script>
									</td>
									<td NOWRAP>~</td>
									<td NOWRAP>
									   <script language =javascript src='./js/m2111ra1_1_fpDateTime1_txtToPoDt.js'></script>
									</td>
								</tr>
							</table>
						</TD>
						<TD CLASS="TD5" NOWRAP>�ʿ���</TD>
						<TD CLASS="TD6" NOWRAP>
							<table cellspacing=0 cellpadding=0>
								<tr>
									<td NOWRAP>
										<script language =javascript src='./js/m2111ra1_1_fpDateTime2_txtFrDlvyDt.js'></script>
									</td>
									<td NOWRAP>~</td>
									<td NOWRAP>
										<script language =javascript src='./js/m2111ra1_1_fpDateTime2_txtToDlvyDt.js'></script>
									</td>
								<tr>
							</table>
						</TD>						
					</TR>
					<TR>
						<TD CLASS="TD5" NOWRAP>���ֹ�ȣ</TD>
						<TD CLASS="TD6"><INPUT NAME="txtSoNo" ALT="���ֹ�ȣ" TYPE="Text" MAXLENGTH=18 SiZE=26 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSoNo" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenSoNo"></TD>
						<TD CLASS="TD5" NOWRAP>Tracking No.</TD>
						<TD CLASS="TD6"><INPUT NAME="txtTrackingNo" ALT="Tracking No." TYPE="Text" MAXLENGTH=25 SiZE=26  tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnTrackingNo" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenTrackingNo"></TD>
					</TR>
				</TABLE>
			</FIELDSET>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=60% valign=top>
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD HEIGHT="100%">
						<script language =javascript src='./js/m2111ra1_1_vaSpread1_vspdData1.js'></script>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=40% valign=top>
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD HEIGHT="100%">
						<script language =javascript src='./js/m2111ra1_1_vaSpread1_vspdData.js'></script>
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
					<TD WIDTH=70% NOWRAP><IMG SRC="../../../CShared/image/query_d.gif"  Style="CURSOR: hand" ALT="Search" NAME="Search" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Query.gif',1)"  ONCLICK="FncQuery()"   ></IMG>
					<IMG SRC="../../../CShared/image/zpConfig_d.gif"  Style="CURSOR: hand" ALT="Config" NAME="Config" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/zpConfig.gif',1)"  ONCLICK="OpenOrderBy()"></IMG></TD>
					</TD>
					<TD WIDTH=30% ALIGN=RIGHT><IMG SRC="../../../CShared/image/ok_d.gif"     Style="CURSOR: hand" ALT="OK"     NAME="pop1"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/OK.gif',1)"     ONCLICK="OkClick()"    ></IMG>
							                  <IMG SRC="../../../CShared/image/cancel_d.gif" Style="CURSOR: hand" ALT="CANCEL" NAME="pop2"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)" ONCLICK="CancelClick()"></IMG>					</TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>	
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>>
		<IFRAME NAME="MyBizASP" WIDTH=100% SRC="../../blank.htm" HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>

<INPUT TYPE=HIDDEN NAME="hdnFrDt" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnToDt" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnFrDt2" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnToDt2" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnSoNo" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnTrackingNo" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnSupplierCd" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnGroupCd" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnGroupNm" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnProcuType" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnSubcontraflg" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnSTOflg" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnPlantCd" tag="14">


</FORM>
<DIV ID="MousePT" NAME="MousePT" STYLE="visible:true">
	<IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm"></IFRAME>
</DIV>
</BODY>
</HTML>                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                     