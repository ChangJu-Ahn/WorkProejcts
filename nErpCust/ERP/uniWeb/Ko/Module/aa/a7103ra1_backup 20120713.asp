 <%@ LANGUAGE="VBSCRIPT" %>
<%Response.Expires = -1%>
<!--'**********************************************************************************************
'*  1. Module Name          : ȸ����� 
'*  2. Function Name        : �ڻ���� 
'*  3. Program ID           : Asset Acquisition Reference Popup
'*  4. Program Name         : �ڻ�Master ���� �˾�(ȸ�����-�ڻ����-�����ڻ�Master����-�ڻ��ȣ ����)
'*  5. Program Desc         :
'*  6. Comproxy List        : ADO�� �ۼ� 
'*  7. Modified date(First) : 2001/05/23
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Kim Hee Jung
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'********************************************************************************************** -->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- '#########################################################################################################
'												1. �� �� �� 
'############################################################################################################
'******************************************  1.1 Inc ����   **********************************************
'	���: Inc. Include
'************************************************************************************************************ -->
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
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/Cookie.vbs">					</SCRIPT>
<SCRIPT LANGUAGE = "JavaScript"SRC = "../../inc/incImage.js">				</SCRIPT>
<SCRIPT LANGUAGE="VBScript">

Option Explicit                              '��: indicates that All variables must be declared in advance

'****************************************  1.2 Global ����/��� ����  ***********************************
'	1. Constant�� �ݵ�� �빮�� ǥ��.
'**********************************************************************************************************
<!-- #Include file="../../inc/lgvariables.inc" -->	

'Dim lgBlnFlgChgValue                                        '��: Variable is for Dirty flag            
'Dim lgStrPrevKey                                            '��: Next Key tag                          
'Dim lgSortKey                                               '��: Sort���� ���庯��                      
Dim lgIsOpenPop                                             '��: Popup status                           

Dim lgSelectList                                            '��: SpreadSheet�� �ʱ�  ��ġ�������� ���� 
Dim lgSelectListDT                                          '��: SpreadSheet�� �ʱ�  ��ġ�������� ���� 

Dim lgTypeCD                                                '��: 'G' is for group , 'S' is for Sort    
Dim lgFieldCD                                               '��: �ʵ� �ڵ尪                           
Dim lgFieldNM                                               '��: �ʵ� ������                           
Dim lgFieldLen                                              '��: �ʵ� ��(Spreadsheet����)              
Dim lgFieldType                                             '��: �ʵ� ������                           
Dim lgDefaultT                                              '��: �ʵ� �⺻��                           
Dim lgNextSeq                                               '��: �ʵ� Pair��                           
Dim lgKeyTag                                                '��: Key ����                                

Dim lgSortFieldNm                                           '��: Orderby popup�� ����Ÿ(�ʵ弳��)      
Dim lgSortFieldCD                                          '��: Orderby popup�� ����Ÿ(�ʵ��ڵ�)      

Dim lgPopUpR                                                '��: Orderby default ��                    
Dim lgMark

Dim IsOpenPop                                                  '��: ��ũ                                  
'---------------  coding part(�������,Start)-----------------------------------------------------------
' EndDate = GetSvrDate                                           '��: �ʱ�ȭ�鿡 �ѷ����� ���� ��¥ -----
'	StartDate = UNIDateAdd("m", -1, EndDate, PopupParent.gServerDateFormat)                          '��: �ʱ�ȭ�鿡 �ѷ����� ���� ��¥ -----
'	Call GetAdoFiledInf("A7103RA1","S","A")                        '��: spread sheet �ʵ����� query   -----		         
                                                                  
'--------------- ������ coding part(�������,End)-------------------------------------------------------------

'--------------- ������ coding part(��������,Start)-----------------------------------------------------------
Const BIZ_PGM_ID        = "a7103rb1.asp"
Const C_SHEETMAXROWS    = 16                                   '��: Spread sheet���� �������� row
Const C_SHEETMAXROWS_D  = 30   
Const C_MaxKey          = 12	
                                '��: Server���� �ѹ��� fetch�� �ִ� ����Ÿ �Ǽ� 
Dim lsPoNo                                                 '��: Jump�� Cookie�� ���� Grid value

Dim arrReturn
Dim arrParent
Dim arrParam					

	 '------ Set Parameters from Parent ASP ------ 
arrParent		= window.dialogArguments
Set PopupParent = arrParent(0)
arrParam		= arrParent(1)

	' ���Ѱ��� �߰� 
	Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' ����� 
	Dim lgInternalCd, lgDeptCd, lgDeptNm			' ���κμ� 
	Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' ���κμ�(��������)
	Dim lgAuthUsrID, lgAuthUsrNm					' ���� 



	top.document.title = "�ڻ�Master �����˾�"

'--------------- ������ coding part(��������,End)-------------------------------------------------------------

 '#########################################################################################################
'												2. Function�� 
'	���� : �����ڰ� ������ �Լ�, �� Event���� �Լ��� ������ ��� ����� ���� �Լ� �⽽ 
'	�������� ���� ���� : 1. Sub �Ǵ� Function�� ȣ���� �� �ݵ�� Call�� ����.
'		     	     	 2. Sub, Function �̸��� _�� ���� �ʵ��� �Ѵ�. (Event�� �����ϱ� ����) 
'######################################################################################################### 

 '==========================================  2.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : ���� �ʱ�ȭ(Global ����, �ʱ�ȭ�� �ʿ��� ���� �Ǵ� Flag���� Setting�Ѵ�.)
'========================================================================================================= 
Sub InitVariables()
    Redim arrReturn(0)
    lgBlnFlgChgValue = False                               'Indicates that no value changed
    lgStrPrevKey     = ""                                  'initializes Previous Key
    lgSortKey        = 1
    
	Self.Returnvalue = arrReturn

	' ���Ѱ��� �߰� 
	If UBound(arrParam) > 5 Then
		lgAuthBizAreaCd	= arrParam(5)
		lgInternalCd	= arrParam(6)
		lgSubInternalCd	= arrParam(7)
		lgAuthUsrID		= arrParam(8)    
	End If

End Sub

 '==========================================  2.2 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : ������ ���� �ʵ� ���� ������ ������ 
'                 lgSort...�� �����ϴ� ���� ������ sort��� ����� ���� 
'                 IsPopUpR ���������� sort ������ �⺻�� �Ǵ� �� ���� 
'========================================================================================================= 
Sub SetDefaultVal()

	Dim strYear, strMonth, strDay, dtToday, EndDate, StartDate
	'##
	dtToday = "<%=GetSvrDate%>"
	Call PopupParent.ExtractDateFrom(dtToday, PopupParent.gServerDateFormat, PopupParent.gServerDateType, strYear, strMonth, strDay)

	EndDate = UniConvYYYYMMDDToDate(PopupParent.gDateFormat, strYear, strMonth, strDay)
	StartDate = UNIDateClientFormat(PopupParent.gFiscStart)
	'StartDate = UNIDateAdd("M", -1, EndDate, PopupParent.gDateFormat)

	frm1.txtFrAcqDt.Text = StartDate
	frm1.txtToAcqDt.Text = EndDate
	frm1.hOrgChangeId.value = PopupParent.gChangeOrgId

End Sub

Function OpenPopUp(Byval PopFg,strCode)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True
	
	Select Case UCase(PopFg)
'		Case "DP"	
'			arrParam(0) = "ȸ��μ� �˾�"				' �˾� ��Ī 
'			arrParam(1) = "B_ACCT_DEPT"    				' TABLE ��Ī 
'			arrParam(2) = strCode						' Code Condition
'			arrParam(3) = ""							' Name Cindition
'			arrParam(4) = ""							' Where Condition
'			arrParam(5) = "ȸ��μ�"				' �����ʵ��� �� ��Ī 
'
'			arrField(0) = "DEPT_CD"	     				' Field��(0)
'			arrField(1) = "DEPT_NM"			    		' Field��(1)
'
'			arrHeader(0) = "ȸ��μ��ڵ�"				' Header��(0)
'			arrHeader(1) = "ȸ��μ���"		  			' Header��(1)
		Case "FA", "TA"
			arrParam(0) = "�ڻ긶���� �˾�"				' �˾� ��Ī 
			arrParam(1) = "A_ASSET_MASTER"    				' TABLE ��Ī 
			arrParam(2) = strCode						' Code Condition
			arrParam(3) = ""							' Name Cindition

			arrParam(4) = " 1 = 1 "							' Where Condition

			' ���Ѱ��� �߰� 
			If lgAuthBizAreaCd <> "" Then
				arrParam(4) = arrParam(4) & " AND BIZ_AREA_CD = " & FilterVar(lgAuthBizAreaCd, "''", "S")			' Where Condition
			End If

			If lgInternalCd <> "" Then
				arrParam(4) = arrParam(4) & " AND INTERNAL_CD = " & FilterVar(lgInternalCd, "''", "S")			' Where Condition
			End If

			If lgSubInternalCd <> "" Then
				arrParam(4) = arrParam(4) & " AND INTERNAL_CD like " & FilterVar(lgSubInternalCd & "%", "''", "S")		' Where Condition
			End If

			If lgAuthUsrID <> "" Then
				arrParam(4) = arrParam(4) & " AND INSRT_USER_ID = " & FilterVar(lgAuthUsrID, "''", "S")			' Where Condition
			End If

			arrParam(5) = "�ڻ�"				' �����ʵ��� �� ��Ī 

			arrField(0) = "ASST_NO"	     				' Field��(0)
			arrField(1) = "ASST_NM"			    		' Field��(1)
    
			arrHeader(0) = "�ڻ��ȣ"				' Header��(0)
			arrHeader(1) = "�ڻ��"		  			' Header��(1)    	

		Case "AC"
			arrParam(0) = "�����˾�"			                            ' �˾� ��Ī 
			arrParam(1) = "A_ACCT A, A_ASSET_ACCT B"  			' TABLE ��Ī 
			arrParam(2) = strCode						' Code Condition
			arrParam(3) = ""							' Name Cindition
			arrParam(4) = "A.ACCT_CD = B.ACCT_CD"		' Where Condition
			arrParam(5) = "����"				' �����ʵ��� �� ��Ī 

			arrField(0)  = "B.ACCT_CD"	     				' Field��(0)
			arrField(1)  = "A.ACCT_NM"	     				' Field��(1)
				
			arrHeader(0) = "�����ڵ�"				' Header��(0)
			arrHeader(1) = "������"					' Header��(1)    
	End Select
	
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
	     "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		With frm1
		select case PopFg
'			case "DP"
'				.txtDeptCd.focus
			case "AC"
				.txtAcctCd.focus
			case "FA"
				.txtFrAsstNo.focus
			case "TA"
				.txtToAsstNo.focus
			end select
		End With
		Exit Function
	Else
		Call SetPopUp(PopFg,arrRet)
	End If	

End Function

Function SetPopUp(Byval PopupFg,Byval arrRet)
	
	With frm1
	select case PopupFg
'		case "DP"
'			.txtDeptCd.focus
'			.txtDeptCd.value	 = Trim(arrRet(0))
'			.txtDeptNm.value	 = Trim(arrRet(1))

		case "AC"
			.txtAcctCd.focus
			.txtAcctCd.value	 = Trim(arrRet(0))
			.txtAcctNm.value	 = Trim(arrRet(1))

		case "FA"
			.txtFrAsstNo.focus
			.txtFrAsstNo.value = Trim(arrRet(0))
'			.txtFrAsstNm.value = arrRet(1)
		case "TA"
			.txtToAsstNo.focus
			.txtToAsstNo.value = Trim(arrRet(0))
'			.txtToAsstNm.value = arrRet(1)
		end select
	End With

End Function
'------------------------------------------  OpenDeptOrgPopup()  ---------------------------------------
'	Name : OpenDeptOrgPopup()
'	Description : OpenAp Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function OpenDeptOrgPopup()
	Dim arrRet
	Dim arrParam(8)

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	arrParam(0) = frm1.txtFrAcqDt.text								'  Code Condition
   	arrParam(1) = frm1.txtToAcqDt.Text
	'arrParam(2) = lgUsrIntCd                            ' �ڷ���� Condition  
	arrParam(3) = frm1.txtDeptCd.value
	arrParam(4) = "F"									' �������� ���� Condition  
	
	' ���Ѱ��� �߰� 
	arrParam(5)		= lgAuthBizAreaCd
	arrParam(6)		= lgInternalCd
	arrParam(7)		= lgSubInternalCd
	arrParam(8)		= lgAuthUsrID

	arrRet = window.showModalDialog("../../comasp/DeptPopupOrg.asp", Array(PopupParent,arrParam), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		
	lgIsOpenPop = False
	
	frm1.txtDeptCd.focus
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetDept(arrRet)
	End If	
End Function

'------------------------------------------  SetDept()  --------------------------------------------------
'	Name : SetDept()
'	Description : CtrlItem Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 

Function SetDept(Byval arrRet)
		frm1.hOrgChangeId.value=arrRet(2)
		
		frm1.txtDeptCd.value = arrRet(0)
		frm1.txtDeptNm.value = arrRet(1)		
		frm1.txtFrAcqDt.text = arrRet(4)
		frm1.txtToAcqDt.text = arrRet(5)
End Function


'========================================  2.3 LoadInfTB19029()  =========================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'===========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "A","NOCOOKIE","RA") %>                                '��: 
	<% Call LoadBNumericFormatA("Q", "A", "NOCOOKIE", "RA") %>
	
End Sub


'===========================================  2.3.1 OkClick()  ==========================================
'=	Name : OkClick()																					=
'=	Description : Return Array to Opener Window when OK button click									=
'=				  �� �κп��� �÷� �߰��ϰ� ����Ÿ ������ �Ͼ�� �մϴ�.   							=
'========================================================================================================
Function OKClick()
	Dim ii

	If frm1.vspdData.ActiveRow > 0 Then
		Redim arrReturn(C_MaxKey)
		frm1.vspdData.Row = frm1.vspdData.ActiveRow
		
		For ii = 0 To C_MaxKey - 1
			frm1.vspdData.Col  = GetKeyPos("A",ii + 1)
			arrReturn(ii) = frm1.vspdData.Text
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

'=========================================  2.3.3 Mouse Pointer ó�� �Լ� ===============================
'========================================================================================================
Function MousePointer(pstr1)
	Select case UCase(pstr1)
        Case "PON"
			window.document.search.style.cursor = "wait"
        Case "POFF"
			window.document.search.style.cursor = ""
    End Select
End Function

'========================================= 2.6 InitSpreadSheet() =========================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'==========================================================================================================
Sub InitSpreadSheet()
    frm1.vspdData.operationmode = 3
    Call SetZAdoSpreadSheet("A7103RA1","S","A","V20021211",PopupParent.C_SORT_DBAGENT,frm1.vspdData, C_MaxKey, "X","X")
    Call SetSpreadLock() 
End Sub

'========================================= 2.7 SetSpreadLock() ===========================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'=========================================================================================================
Sub SetSpreadLock()
    With frm1
		.vspdData.ReDraw = False
		ggoSpread.SpreadLockWithOddEvenRowColor()
		.vspdData.ReDraw = True
    End With
End Sub

'**********************  2.4 POP-UP ó���Լ�  ****************************************
'	���: ���� POP-UP
'   ����: ���� POP-UP�� ���� Open�� Include�Ѵ�. 
'	      �ϳ��� ASP���� Popup�� �ߺ��Ǹ� �ϳ��� ��ũ���� ����ϰ� �������� �������Ͽ� ����Ѵ�.
'************************************************************************************** 

 '-----------------------  OpenItem()  -------------------------------------------------
'	Name : OpenItem()
'	Description : Item PopUp
'--------------------------------------------------------------------------------------- 

'===========================================================================
' Function Name : OpenOrderBy
' Function Desc : OpenOrderBy Reference Popup
'===========================================================================
Function OpenOrderBy()
	Dim arrRet
	Dim arrParam
	Dim TInf(5)
	Dim ii
	
	On Error Resume Next
	
	ReDim arrParam(PopupParent.C_MaxSelList * 2 - 1 )

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True
	
    TInf(0) = gMethodText
  
	For ii = 0 to PopupParent.C_MaxSelList * 2 - 1 Step 2
      arrParam(ii + 0 ) = lgPopUpR(ii / 2  , 0)
      arrParam(ii + 1 ) = lgPopUpR(ii / 2  , 1)
    Next  
  
	arrRet = window.showModalDialog("../../ComAsp/ZADOSortPopup.asp",Array(ggoSpread.GetXMLData("A"),gMethodText),"dialogWidth=" & PopupParent.SORTW_WIDTH & "px; dialogHeight=" & PopupParent.SORTW_HEIGHT & "px;; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If arrRet(0) = "0" Then
		If Err.Number <> 0 Then
			Err.Clear 
		End If
		Exit Function
	Else
		Call ggSpread.SaveXMLData("A",arrRet(0),arrRet(1))
	   For ii = 0 to PopupParent.C_MaxSelList * 2 - 1 Step 2
           lgPopUpR(ii / 2 ,0) = arrRet(ii + 1)  
           lgPopUpR(ii / 2 ,1) = arrRet(ii + 2)
       Next    
	   
       Call InitVariables
       Call InitSpreadSheet
   End If
End Function

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

 '++++++++++++++++++++++++++++++++++++++++++  2.5 ������ ���� �Լ�  +++++++++++++++++++++++++++++++++++++++
'    ������ ���α׷� ���� �ʿ��� ������ ���� Procedure (Sub, Function, Validation & Calulation ���� �Լ�)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ 

 '==========================================   CookiePage()  ======================================
'	Name : CookiePage()
'	Description : JUMP�� Loadȭ������ ���Ǻη� Value
'==================================================================================================== 

 '#########################################################################################################
'												3. Event�� 
'	���: Event �Լ��� ���� ó�� 
'	����: Windowó��, Singleó��, Gridó�� �۾�.
'         ���⼭ Validation Check, Calcuration �۾��� ������ Event�� �߻�.
'         �� Object������ Grouping�Ѵ�.
'##########################################################################################################
 '******************************************  3.1 Window ó��  *********************************************
'	Window�� �߻� �ϴ� ��� Even ó��	
'********************************************************************************************************* 
 '==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(���� Include ���Ͽ� ����)�� �����ʱ�ȭ �� ȭ���ʱ�ȭ�� �ϱ� ���� �Լ��� Call�ϴ� �κ� 
'========================================================================================================= 
Sub Form_Load()

    Call LoadInfTB19029														'��: Load table , B_numeric_format
	Call ggoOper.FormatField(Document, "1",PopupParent.ggStrIntegeralPart, PopupParent.ggStrDeciPointPart,PopupParent.gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec)
	Call ggoOper.FormatField(Document, "2",PopupParent.ggStrIntegeralPart, PopupParent.ggStrDeciPointPart,PopupParent.gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec)
    Call ggoOper.LockField(Document, "N")                                   '��: Lock  Suitable  Field
    
    'ReDim lgPopUpR(parent.C_MaxSelList - 1,1)
	Call InitVariables														'��: Initializes local global variables
	Call SetDefaultVal	
	Call InitSpreadSheet()
    
'--------------- ������ coding part(�������,Start)----------------------------------------------------
   
'--------------- ������ coding part(�������,End)------------------------------------------------------
End Sub

'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================

Sub Form_QueryUnload(Cancel , UnloadMode )
End Sub

 '**************************  3.2 HTML Form Element & Object Eventó��  **********************************
'	Document�� TAG���� �߻� �ϴ� Event ó��	
'	Event�� ��� �Ʒ��� ����� Event�̿��� ����� �����ϸ� �ʿ�� �߰� �����ϳ� 
'	Event�� �浹�� �����Ͽ� �ۼ��Ѵ�.
'********************************************************************************************************* 

 '******************************  3.2.1 Object Tag ó��  *********************************************
'	Window�� �߻� �ϴ� ��� Even ó��	
'********************************************************************************************************* 


'*********************************************  3.3 Object Tag ó��  ************************************
'*	Object���� �߻� �ϴ� Event ó��																		*
'********************************************************************************************************


Function document_onkeypress()
	If window.event.keyCode = 27 Then
        Call CancelClick()
    End If
End Function

Sub ConditionKeypress()
	If window.event.keyCode = 13 Then
		Call Search_OnClick()
	End If
End sub
Sub txtDeptCd_onBlur()
	If frm1.txtDeptCd.value = "" Then
		frm1.txtDeptNm.value = ""
	End If
End sub



'==========================================================================================
'   Event Name : txtFrAcqDt
'   Event Desc :
'==========================================================================================

Sub txtFrAcqDt_DblClick(Button)
	if Button = 1 then
		frm1.txtFrAcqDt.Action = 7
        Call SetFocusToDocument("M")	
        frm1.txtFrAcqDt.Focus  
	End if
End Sub

'==========================================================================================
'   Event Name : txtToAcqDt
'   Event Desc :
'==========================================================================================

Sub txtToAcqDt_DblClick(Button)
	if Button = 1 then
		frm1.txtToAcqDt.Action = 7
        Call SetFocusToDocument("M")	
        frm1.txtToAcqDt.Focus  
	End if
End Sub

Sub  txtFrAcqDt_KeyPress(KeyAscii)
	On Error Resume Next
	If KeyAscii = 27 Then 
		Call CancelClick()
	ElseIf KeyAscii = 13 Then 
		Call Fncquery()
	End IF
End Sub  

Sub  txtToAcqDt_KeyPress(KeyAscii)
	On Error Resume Next
	If KeyAscii = 27 Then 
		Call CancelClick()
	ElseIf KeyAscii = 13 Then 
		Call Fncquery()
	End IF
End Sub  


Function  txtFrAcqDt_change()
	Call txtDeptCD_OnChange()
End Function  

Function  txtToAcqDt_change()
	Call txtDeptCD_OnChange()
End Function  


'==========================================================================================
'   Event Name : txtDeptCd_Onchange
'   Event Desc : 
'==========================================================================================
Sub txtDeptCD_OnChange()
        
    Dim strSelect, strFrom, strWhere 	
    Dim IntRetCD 
	Dim arrVal1, arrVal2
	Dim ii, jj

	if frm1.txtDeptCd.value = "" then
		frm1.txtDeptNm.value = ""
	end if
	
    lgBlnFlgChgValue = True
	
	If TRim(frm1.txtDeptCd.value) <>"" Then
		'----------------------------------------------------------------------------------------
			strSelect = "dept_cd, ORG_CHANGE_ID"
			strFrom =  " B_ACCT_DEPT "
			strWhere = " ORG_CHANGE_DT >= "
			strWhere = strWhere & " (select max(org_change_dt) from B_ACCT_DEPT where org_change_dt<= " & FilterVar(UNIConvDateToYYYYMMDD(frm1.txtFrAcqDt.Text, gDateFormat,PopupParent.gServerDateType), "''", "S") & ")"
			strWhere = strWhere & " AND ORG_CHANGE_DT <= " 
			strWhere = strWhere & " (select max(org_change_dt) from B_ACCT_DEPT where org_change_dt<= " & FilterVar(UNIConvDateToYYYYMMDD(frm1.txtToAcqDt.Text, gDateFormat,PopupParent.gServerDateType), "''", "S") & ") "
			strWhere =	strWhere & " AND dept_Cd = " & FilterVar(LTrim(RTrim(frm1.txtDeptCd.value)), "''", "S")
		
					'msgbox "Select " & strSelect& " from " &strFrom & " where "&strWhere

		If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) = False Then
	
			IntRetCD = DisplayMsgBox("124600","X","X","X")  
			frm1.txtDeptCd.value = ""
			frm1.txtDeptNm.value = ""
			frm1.hOrgChangeId.value = ""
			frm1.txtDeptCd.focus
		Else 
		
			arrVal1 = Split(lgF2By2, Chr(11) & Chr(12))			
			jj = Ubound(arrVal1,1)
					
			For ii = 0 to jj - 1
				arrVal2 = Split(arrVal1(ii), chr(11))			
				frm1.hOrgChangeId.value = Trim(arrVal2(2))
				
			Next	
			
		End If
	End IF
		'----------------------------------------------------------------------------------------

End Sub

'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================

Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
     '----------  Coding part  -------------------------------------------------------------   
    if frm1.vspdData.MaxRows < NewTop + C_SHEETMAXROWS Then	'��: ������ üũ'
		If lgStrPrevKey <> "" Then							'��: ���� Ű ���� ������ �� �̻� ��������ASP�� ȣ������ ���� 
			DbQuery
		End If
   End if
    
End Sub



'======================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : �÷��� Ŭ���� ��� �߻� 
'=======================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
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
    
'--------------- ������ coding part(�������,Start)----------------------------------------------------
	If Row < 1 Then Exit Sub

	'frm1.vspdData.Row = Row
	'lsPoNo=frm1.vspdData.Text
'--------------- ������ coding part(�������,End)------------------------------------------------------
	Call SetSpreadColumnValue("A",frm1.vspdData,Col,Row)		
End Sub

Function vspdData_KeyPress(KeyAscii)
    If KeyAscii = 13 And frm1.vspdData.ActiveRow > 0 Then
        Call OKClick()
    ElseIf KeyAscii = 27 Then
        Call CancelClick()
    End If
End Function



Sub vspdData_DblClick(ByVal Col, ByVal Row)
	If Frm1.vspdData.MaxRows > 0 Then
		If Frm1.vspdData.ActiveRow = Row Or Frm1.vspdData.ActiveRow > 0 Then
			Call OKClick
		End If
	End If
End Sub

 '#########################################################################################################
'												4. Common Function�� 
'	���: Common Function
'	����: ȯ��ó���Լ�, VAT ó�� �Լ� 
'######################################################################################################### 


 '#########################################################################################################
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
'######################################################################################################### 
 '*******************************  5.1 Toolbar(Main)���� ȣ��Ǵ� Function *******************************
'	���� : Fnc�Լ��� ���� �����ϴ� ��� Function
'********************************************************************************************************* 
Function FncQuery() 
	Dim IntRetCD
    Dim strFrAcqDt, strToAcqDt
    
    FncQuery = False                                                        '��: Processing is NG
    
    Err.Clear                                                               '��: Protect system from crashing   

    '-----------------------
    'Erase contents area
    '-----------------------
    'Call ggoOper.ClearField(Document, "2")									'��: Clear Contents  Field
    Call InitVariables 														'��: Initializes local global variables
    
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then								'��: This function check indispensable field
       Exit Function
    End If
	
	'---------------------------------------
	'������� ���� Check
	'---------------------------------------
		
	strFrAcqDt = UniConvDateToYYYYMMDD(frm1.txtFrAcqDt.Text, PopupParent.gDateFormat,"") 
	strToAcqDt = UniConvDateToYYYYMMDD(frm1.txtToAcqDt.Text, PopupParent.gDateFormat,"")
	    
	If strToAcqDt <> "" Then
		If strFrAcqDt > strToAcqDt Then
			Call DisplayMsgBox("970025", "X", frm1.txtFrAcqDt.Alt, frm1.txtToAcqDt.Alt)
			frm1.txtFrAcqDt.focus
			Exit Function
		End If
	End If
	
	'---------------------------------------
	'�ڻ�����ȣ ���� Check
	'---------------------------------------
	frm1.txtFrAsstNo.value = Trim(frm1.txtFrAsstNo.value)
	frm1.txtToAsstNo.value = Trim(frm1.txtToAsstNo.value)
	
	If frm1.txtFrAsstNo.value <> "" And frm1.txtToAsstNo.value <> "" Then
		If frm1.txtFrAsstNo.value > frm1.txtToAsstNo.value Then
			Call DisplayMsgBox("970025", "X", frm1.txtFrAsstNo.Alt, frm1.txtToAsstNo.Alt)
			frm1.txtFrAsstNo.focus 
			Exit Function
		End If
	End If
	
	IF NOT CheckOrgChangeId Then
		  IntRetCD = DisplayMsgBox("124600","X","X","X")           '��: Display Message(There is no changed data.)
		Exit Function
	End if
    '-----------------------
    'Query function call area
    '-----------------------
	'frm1.vspdData.MaxRows = 0                                                   '��: Protect system from crashing
	ggoSpread.Source = frm1.vspdData
	ggospread.ClearSpreadData		'Buffer Clear
	
    Call DbQuery															'��: Query db data

    FncQuery = True		
End Function


'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================

Function FncPrint() 
    Call PopupParent.FncPrint()
End Function


'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================

Function FncExcel() 
	Call PopupParent.FncExport(PopupParent.C_MULTI)
End Function


'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================

Function FncFind() 
    Call PopupParent.FncFind(PopupParent.C_MULTI , False)                                     '��:ȭ�� ����, Tab ���� 
End Function


'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================

Function FncExit()
	Dim IntRetCD
	FncExit = False

    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900016", PopupParent.VB_YES_NO,,"X","X")
		If IntRetCD = vbNo Then
		    Exit Function
		End If
    End If

    FncExit = True
End Function

 '*******************************  5.2 Fnc�Լ������� ȣ��Ǵ� ���� Function  *******************************
'	���� : 
'********************************************************************************************************* 

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================

Function DbQuery() 
	Dim strVal
	Dim lgPid
	
	lgPid = "<%=Request("PID")%>"

    DbQuery = False
    
    Err.Clear            
    
	Call LayerShowHide(1)

    With frm1
'--------------- ������ coding part(�������,Start)----------------------------------------------
		strVal = BIZ_PGM_ID & "?txtFrAcqDt="	& Trim(.txtFrAcqDt.Text)
		strVal = strVal & "&txtToAcqDt="		& Trim(.txtToAcqDt.Text)
		strVal = strVal & "&txtFrAsstNo="		& Trim(.txtFrAsstNo.value)
		strVal = strVal & "&txtToAsstNo="		& Trim(.txtToAsstNo.value) 
		strVal = strVal & "&txtAcctCd="			& Trim(.txtAcctCd.value)
	    strVal = strVal & "&txtDeptCd="			& Trim(.txtDeptCd.value)
		
'--------------- ������ coding part(�������,End)------------------------------------------------
        strVal = strVal & "&lgStrPrevKey="		& lgStrPrevKey                      '��: Next key tag
        strVal = strVal & "&lgMaxCount="		& CStr(C_SHEETMAXROWS_D)            '��: �ѹ��� �����ü� �ִ� ����Ÿ �Ǽ� 
		strVal = strVal & "&lgSelectListDT="	& GetSQLSelectListDataType("A")
	    strVal = strVal & "&lgTailList="		& MakeSQLGroupOrderByList("A")
		strVal = strVal & "&lgSelectList="		& EnCoding(GetSQLSelectList("A"))
		strVal = strVal & "&PID="				& lgPid
		
		' ���Ѱ��� �߰� 
		strVal = strVal & "&lgAuthBizAreaCd="	& lgAuthBizAreaCd			' ����� 
		strVal = strVal & "&lgInternalCd="		& lgInternalCd				' ���κμ� 
		strVal = strVal & "&lgSubInternalCd="	& lgSubInternalCd			' ���κμ�(��������)
		strVal = strVal & "&lgAuthUsrID="		& lgAuthUsrID				' ���� 
		
        
        Call RunMyBizASP(MyBizASP, strVal)										'��: �����Ͻ� ASP �� ���� 
    
    End With
    
    DbQuery = True


End Function


'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================

Function DbQueryOk()														'��: ��ȸ ������ ������� 

    '-----------------------
    'Reset variables area
    '-----------------------
    lgBlnFlgChgValue = True                                                 'Indicates that no value changed
'    Call ggoOper.LockField(Document, "Q")									'��: This function lock the suitable field

End Function
'==========================================================================================
'   Event Name : CheckOrgChangeId
'   Event Desc : 
'==========================================================================================
Function CheckOrgChangeId()
	Dim strSelect, strFrom, strWhere 	
	Dim IntRetCD 
	Dim ii, jj
	Dim arrVal1, arrVal2
 
	CheckOrgChangeId = True
 
	With frm1
	
		If LTrim(RTrim(.txtDeptCd.value)) <> "" Then
			'----------------------------------------------------------------------------------------
			strSelect = "Distinct ORG_CHANGE_ID"
			strFrom =  " B_ACCT_DEPT "
			strWhere = " ORG_CHANGE_DT >= "
			strWhere = strWhere & " (select max(org_change_dt) from B_ACCT_DEPT where org_change_dt<= " & FilterVar(UNIConvDateToYYYYMMDD(.txtFrAcqDt.Text, gDateFormat,PopupParent.gServerDateType), "''", "S") & ")"
			strWhere = strWhere & " AND ORG_CHANGE_DT <= " 
			strWhere = strWhere & " (select max(org_change_dt) from B_ACCT_DEPT where org_change_dt<= " & FilterVar(UNIConvDateToYYYYMMDD(.txtToAcqDt.Text, gDateFormat,PopupParent.gServerDateType), "''", "S") & ") "
			strWhere = strWhere & " AND ORG_CHANGE_ID =  " & FilterVar(.hOrgChangeId.value , "''", "S") & ""
			strWhere =	strWhere & " AND dept_Cd = " & FilterVar(LTrim(RTrim(frm1.txtDeptCd.value)), "''", "S")
					'msgbox "Select " & strSelect& " from " &strFrom & " where "&strWhere

			IntRetCD= CommonQueryRs(strSelect,strFrom,strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
					
			If IntRetCD = False  OR Trim(Replace(lgF0,Chr(11),"")) <> Trim(.hOrgChangeId.value) Then
					.txtDeptCd.value = ""
					.txtDeptNm.value = ""
					.hOrgChangeId.value = ""
					.txtDeptCd.focus
					CheckOrgChangeId = False
			End if
		End If
	End With
		'----------------------------------------------------------------------------------------

End Function
'��: �Ʒ� OBJECT Tag�� InterDev ����ڸ� ���Ѱ����� ���α׷��� �ϼ��Ǹ� �Ʒ� Include �ڵ�� ��ü�Ǿ�� �Ѵ� 
</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->	
</HEAD>
<!-- '#########################################################################################################
'       					6. Tag�� 
'#########################################################################################################  -->
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
						<TD CLASS=TD5 NOWRAP>�������</TD>
						<TD CLASS=TD6 NOWRAP>
							<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT id=fpDateTime1 title=FPDATETIME CLASS=FPDTYYYYMMDD name=txtFrAcqDt CLASSID=<%=gCLSIDFPDT%> ALT="�����������" tag="11"> </OBJECT>');</SCRIPT>&nbsp;~&nbsp;
							<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT id=fpDateTime2 title=FPDATETIME CLASS=FPDTYYYYMMDD name=txtToAcqDt CLASSID=<%=gCLSIDFPDT%> ALT="�����������" tag="11"> </OBJECT>');</SCRIPT>
						</TD>												
						<TD CLASS=TD5 NOWRAP>�ڻ��ȣ</TD>				
						<TD CLASS=TD6 NOWRAP>
						<INPUT TYPE="Text" NAME="txtFrAsstNo" SIZE=15 MAXLENGTH=18 STYLE="TEXT-ALIGN: left" tag="1XXXXU" ALT="�����ڻ��ȣ"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCostCd" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup('FA',frm1.txtFrAsstNo.Value)">&nbsp;~&nbsp;
						<INPUT TYPE="Text" NAME="txtToAsstNo" SIZE=15 MAXLENGTH=18 STYLE="TEXT-ALIGN: left" tag="1XXXXU" ALT="�����ڻ��ȣ"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCostCd" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup('TA',frm1.txtToAsstNo.Value)">
						</TD>
					</TR>			
					<TR>					
						<TD CLASS=TD5 NOWRAP>�μ��ڵ�</TD>
						<TD CLASS=TD6 NOWRAP><INPUT NAME="txtDeptCd" ALT="�μ��ڵ�" MAXLENGTH="10" SIZE=10 STYLE="TEXT-ALIGN: left" tag  ="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCostCd" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenDeptOrgPopup()">
											 <INPUT NAME="txtDeptNm" ALT="�μ���"   MAXLENGTH="20" SIZE=20 STYLE="TEXT-ALIGN: left" tag="14X"></TD>
						<TD CLASS=TD5 NOWRAP>�����ڵ�</TD>
						<TD CLASS=TD6 NOWRAP><INPUT NAME="txtAcctCd" ALT="�����ڵ�" MAXLENGTH="20" SIZE=10 STYLE="TEXT-ALIGN: left" tag  ="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCostCd" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup('AC',frm1.txtAcctCd.Value)">
											 <INPUT NAME="txtAcctNm" ALT="������"   MAXLENGTH="30" SIZE=20 STYLE="TEXT-ALIGN: left" tag="14X"></TD>
					</TR>					
				</TABLE>
			</FIELDSET>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
	</TR>
	<TR height=100%>
		<TD WIDTH=100% HEIGHT=* VALIGN=TOP>
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR HEIGHT=100%>
					<TD WIDTH=100%>
						<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% tag="2" HEIGHT=100% id=vspdData> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"><PARAM NAME="ReDraw" VALUE="0"> <PARAM NAME="FontSize" VALUE="10"> </OBJECT>');</SCRIPT>
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
					<TD >&nbsp;&nbsp;<IMG SRC="../../../CShared/image/query_d.gif"    Style="CURSOR: hand" ALT="Search" NAME="Search" OnClick="FncQuery()"        onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Query.gif',1)" ></IMG>&nbsp;
					                 <IMG SRC="../../../CShared/image/zpConfig_d.gif" Style="CURSOR: hand" ALT="Config" NAME="Config" OnClick="OpenSortPopup()"     onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/zpConfig.gif',1)" ></IMG>
					</TD>
					<TD ALIGN=RIGHT> <IMG SRC="../../../CShared/image/ok_d.gif"       Style="CURSOR: hand" ALT="OK"     NAME="Ok"     OnClick="OkClick()"         onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/OK.gif',1)"    ></IMG>&nbsp;
									 <IMG SRC="../../../CShared/image/cancel_d.gif"   Style="CURSOR: hand" ALT="CANCEL" NAME="Cancel" OnClick="CancelClick()"     onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)"></IMG>&nbsp;&nbsp;
					</TD>
					<!--TD WIDTH=10>&nbsp;</TD>
					<TD>
						<IMG SRC="../../../CShared/image/query_d.gif"  Style="CURSOR: hand" ALT="Search" NAME="Search" ONCLICK="Call FncQuery()">	</IMG>
						&nbsp;&nbsp;<button name="btnAutoSel" class="clsmbtn" ONCLICK="OpenOrderBy()">���ļ���</button>
					</TD>
					<TD ALIGN=RIGHT>
						<IMG SRC="../../../CShared/image/ok_d.gif" Style="CURSOR: hand" ALT="OK" NAME="pop1" ONCLICK="OkClick()" ></IMG>&nbsp;
						<IMG SRC="../../../CShared/image/cancel_d.gif" Style="CURSOR: hand" ALT="CANCEL" NAME="pop2" ONCLICK="CancelClick()" ></IMG>
					</TD>				
					<TD WIDTH=10>&nbsp;</TD-->
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
<INPUT TYPE=hidden NAME="hOrgChangeId"    tag="14" TABINDEX="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
