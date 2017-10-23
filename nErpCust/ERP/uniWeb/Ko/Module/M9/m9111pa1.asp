<%@ LANGUAGE="VBSCRIPT" %>
<%Response.Expires = -1%>
<!--
'**********************************************************************************************
'*  1. Module Name          : 
'*  2. Function Name        : ����̵���û��ȣPOPUP
'*  3. Program ID           : M9111PA1
'*  4. Program Name         : ����̵���û��ȣPOPUP
'*  5. Program Desc         : Open StoNo Popup ASP
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2002/12/10
'*  8. Modified date(Last)  : 																*
'*                            
'*  9. Modifier (First)     : OH,chang won																			*
'* 10. Modifier (Last)      : 																*
'*                            
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"									*
'*                            this mark(��) Means that "may  change"									*
'*                            this mark(��) Means that "must change"									*
'* 13. History              : 																*
'**********************************************************************************************
-->
<HTML>
<HEAD>
<TITLE>����̵���û��ȣ</TITLE>
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<!--
'#########################################################################################################
'												1. �� �� �� 
'##########################################################################################################
'******************************************  1.1 Inc ����   **********************************************
'	���: Inc. Include
'*********************************************************************************************************
-->

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
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentA.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript" SRC="../../inc/incImage.js"></SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit                                                             '��: indicates that All variables must be declared in advance
                                                                            ' ��������� ������ ���� 
<%'****************************************  1.2 Global ����/��� ����  ***********************************
'	1. Constant�� �ݵ�� �빮�� ǥ��.
'**********************************************************************************************************%>
Const BIZ_PGM_ID 		= "m9111pb1.asp"                              '��: Biz Logic ASP Name

'========================================================================================================
'=									1.2 Constant variables 
'========================================================================================================
Const C_SHEETMAXROWS_D  = 30                                          '��: Fetch max count at once
Const C_MaxKey          = 1                                           '��: key count of SpreadSheet

'========================================================================================================
'=									1.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	

'========================================================================================================
'=									1.4 User-defind Variables
'========================================================================================================

Dim C_Po_No 
Dim C_Po_Type
Dim C_Po_TypeNm
Dim C_Release
Dim C_SupplierCd
Dim C_SupplierNm
Dim C_PoDt
Dim C_GroupCd
Dim C_GroupNm

Dim arrReturn
Dim arrParam					
Dim arrField
Dim PlantCd
Dim arrParent

Dim gblnWinEvent
Dim lgKeyPos                                                '��: Key��ġ                               
Dim lgKeyPosVal                                             '��: Key��ġ Value                         
Dim IscookieSplit 

arrParent = window.dialogArguments
Set PopupParent = arrParent(0)

Dim StartDate,EndDate

EndDate = UNIConvDateAToB("<%=GetSvrDate%>", PopupParent.gServerDateFormat, PopupParent.gDateFormat)
StartDate = UnIDateAdd("m", -1, EndDate, PopupParent.gDateFormat)
top.document.title = "����̵���û��ȣ"

'--------------- ������ coding part(��������,End)-------------------------------------------------------------
<% '==========================================  1.2.3 Global Variable�� ����  ===============================
'========================================================================================================= %>
<% '----------------  ���� Global ������ ����  ----------------------------------------------------------- %>

<% '++++++++++++++++  Insert Your Code for Global Variables Assign  ++++++++++++++++++++++++++++++++++++++ %>
 Dim IsOpenPop						' Popup
 Dim arrValue(3)                    ' Popup�Ǵ� â���� �ѱ涧 �μ��� �迭�� �ѱ� 

<% '#########################################################################################################
'												2. Function�� 
'
'	���� : �����ڰ� ������ �Լ�, �� Event���� �Լ��� ������ ��� ����� ���� �Լ� �⽽ 
'	�������� ���� ���� : 1. Sub �Ǵ� Function�� ȣ���� �� �ݵ�� Call�� ����.
'		     	     	 2. Sub, Function �̸��� _�� ���� �ʵ��� �Ѵ�. (Event�� �����ϱ� ����) 
'######################################################################################################### %>


'========================================================================================================
' Name : InitSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================
Sub InitSpreadPosVariables()

    C_Po_No  		= 1
	C_Po_Type 		= 2
	C_Po_TypeNm		= 3
	C_Release 		= 4
	C_SupplierCd 	= 5
	C_SupplierNm 	= 6
	C_PoDt 		    = 7
	C_GroupCd       = 8
	C_GroupNm		= 9
End Sub



<% '==========================================  2.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : ���� �ʱ�ȭ(Global ����, �ʱ�ȭ�� �ʿ��� ���� �Ǵ� Flag���� Setting�Ѵ�.)
'========================================================================================================= %>
Sub InitVariables()
Dim arrParent

	lgStrPrevKeyIndex = ""
	
	lgIntFlgMode = PopupParent.OPMD_CMODE
	gblnWinEvent = False
	
    lgSortKey = 1                                       '��: initializes sort direction
	Redim arrReturn(0)
	Self.Returnvalue = arrReturn

End Sub
<% '==========================================  2.2 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : ������ ���� �ʵ� ���� ������ ������ 
'                 lgSort...�� �����ϴ� ���� ������ sort��� ����� ���� 
'                 IsPopUpR ���������� sort ������ �⺻�� �Ǵ� �� ���� 
'                 ���α׷� ID�� �ְ� go��ư�� �����ų� menu tree���� Ŭ���ϴ� ���� �Ѿ��                  
'========================================================================================================= %>
Sub SetDefaultVal()

<%'--------------- ������ coding part(�������,Start)--------------------------------------------------%>
	frm1.txtFrPoDt.Text = StartDate
	frm1.txtToPoDt.Text = EndDate
<%'--------------- ������ coding part(�������,End)----------------------------------------------------%>

End Sub



<%'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== %>
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029A("I", "*", "NOCOOKIE", "PA") %>     
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
		Dim intColCnt
		With frm1.vspdData	
			Redim arrReturn(.MaxCols - 1)
			If .MaxRows > 0 Then 
			.Row = .ActiveRow
			'For intColCnt = 0 To .MaxCols - 1
			'	.Col = intColCnt + 1
				.Col = C_Po_No
				arrReturn(0) = .Text
			'Next
			end if
		End With
		
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
		Redim arrReturn(0)
		arrReturn(0) = ""
		self.Returnvalue = arrReturn
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
'==========================================================================================
'   Event Name : txtFrPoDt
'   Event Desc :
'==========================================================================================
%>
Sub txtFrPoDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtFrPoDt.Action = 7
		Call SetFocusToDocument("M")	
        frm1.txtFrPoDt.Focus
	End If
End Sub

<%
'==========================================================================================
'   Event Name : txtToPoDt
'   Event Desc :
'==========================================================================================
%>
Sub txtToPoDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtToPoDt.Action = 7
		Call SetFocusToDocument("M")	
        frm1.txtToPoDt.Focus
	End If
End Sub	
	
<% 
'*******************************************  2.4 POP-UP ó���Լ�  **************************************
'*	���: POP-UP																						*
'*	Description : POP-UP Call�ϴ� �Լ� �� Return Value setting ó��										*
'********************************************************************************************************
%>
<%
'===========================================  2.4.1 POP-UP Open �Լ�()  =================================
'=	Name : Open???()																					=
'=	Description : POP-UP Open																			=
'========================================================================================================
%>
<% '------------------------------------------  OpenPoType()  -------------------------------------------------
'	Name : OpenPoType()
'	Description : OpenPoType PopUp
'--------------------------------------------------------------------------------------------------------- %>
Function OpenPotype()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "�̵�����"						<%' �˾� ��Ī %>
	arrParam(1) = "M_CONFIG_PROCESS"						<%' TABLE ��Ī %>
	
	arrParam(2) = Trim(frm1.txtPotypeCd.Value)	<%' Code Condition%>
	'arrParam(3) = Trim(frm1.txtPotypeNm.Value)	<%' Name Cindition%>
	
	arrParam(4) = " sto_flg = " & FilterVar("Y", "''", "S") & "  "							<%' Where Condition%>
	arrParam(5) = "�̵�����"							<%' TextBox ��Ī %>
	
    arrField(0) = "PO_TYPE_CD"					<%' Field��(0)%>
    arrField(1) = "PO_TYPE_NM"					<%' Field��(1)%>
    
    arrHeader(0) = "�̵�����"						<%' Header��(0)%>
    arrHeader(1) = "�̵�������"						<%' Header��(1)%>
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetPotype(arrRet)
	End If	
End Function

Function SetPotype(byval arrRet)	
	frm1.txtPoTypeCd.Value    = arrRet(0)		
	frm1.txtPoTypeNm.Value    = arrRet(1)
End Function

Function OpenSupplier()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "����ó"						<%' �˾� ��Ī %>
	arrParam(1) = "B_BIZ_PARTNER"						<%' TABLE ��Ī %>

	arrParam(2) = Trim(frm1.txtSupplierCd.Value)	<%' Code Condition%>
	'arrParam(3) = Trim(frm1.txtSupplierNm.Value)	<%' Name Cindition%>
	
	arrParam(4) = " IN_OUT_FLAG = " & FilterVar("I", "''", "S") & "  "	
	arrParam(5) = "����ó"							<%' TextBox ��Ī %>
	
    arrField(0) = "BP_Cd"					<%' Field��(0)%>
    arrField(1) = "BP_NM"					<%' Field��(1)%>
    
    arrHeader(0) = "����ó"						<%' Header��(0)%>
    arrHeader(1) = "����ó��"						<%' Header��(1)%>
    
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
	
End Function


Function OpenGroup()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "���ű׷�"	
	arrParam(1) = "B_Pur_Grp"				
	
	arrParam(2) = Trim(frm1.txtGroupCd.Value)
'	arrParam(3) = Trim(frm1.txtGroupNm.Value)	
	
	arrParam(4) = " B_Pur_Grp.USAGE_FLG=" & FilterVar("Y", "''", "S") & "  "			
	arrParam(5) = "���ű׷�"			
	
    arrField(0) = "PUR_GRP"	
    arrField(1) = "PUR_GRP_NM"	
    
    arrHeader(0) = "���ű׷�"		
    arrHeader(1) = "���ű׷��"		
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
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

<%
'=======================================  2.4.2 POP-UP Return�� ���� �Լ�  ==============================
'=	Name : Set???()																						=
'=	Description : Reference �� POP-UP�� Return���� �޴� �κ�											=
'========================================================================================================
%>

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
'==========================================  2.2.3 InitSpreadSheet()  ===================================
' Name : InitSpreadSheet
' Desc : This method initializes spread sheet column property
'========================================================================================================
Sub InitSpreadSheet()
	
	Call InitSpreadPosVariables
	
	ggoSpread.Source = frm1.vspdData
	ggoSpread.Spreadinit "V20021125",, PopupParent.gAllowDragDropSpread
		
	With frm1.vspdData
	
	.ReDraw = false
    .MaxCols = C_GroupNm+1
    .Col = .MaxCols:	.ColHidden = True
    .MaxRows = 0

    Call GetSpreadColumnPos("A")
    ggoSpread.SSSetEdit 		C_Po_No, "����̵���û��ȣ", 20
    ggoSpread.SSSetEdit 		C_Po_Type, "�̵�����", 10,,,4,2
    ggoSpread.SSSetEdit 		C_Po_TypeNm, "�̵�������", 20
    ggoSpread.SSSetEdit 		C_Release, "Ȯ��", 10
    ggoSpread.SSSetEdit 		C_SupplierCd, "����ó", 15,,,4,2
    ggoSpread.SSSetEdit		    C_SupplierNm, "����ó��", 20        'ǰ��԰� �߰� 
    ggoSpread.SSSetDate 		C_PoDt, "�����", 10, 2, PopupParent.gDateFormat
    ggoSpread.SSSetEdit 		C_GroupCd, "���ű׷�", 10,,,4,2
    ggoSpread.SSSetEdit 		C_GroupNm, "���ű׷��", 20
    
	Call ggoSpread.MakePairsColumn(C_Po_Type,C_Po_TypeNm)
	Call ggoSpread.MakePairsColumn(C_SupplierCd,C_SupplierNm)
	Call ggoSpread.MakePairsColumn(C_GroupCd,C_GroupNm)

    ggoSpread.SSSetSplit(1)										'frozen ����߰� 
    
    Call SetSpreadLock
    
    
	.ReDraw = true
	
    End With
	            
End Sub

'============================================ 2.2.4 SetSpreadLock()  ====================================
' Name : SetSpreadLock
' Desc : This method set color and protect in spread sheet celles
'========================================================================================================
Sub SetSpreadLock()
	ggoSpread.Source = frm1.vspdData
			
	frm1.vspdData.ReDraw = False
	ggoSpread.SpreadLock -1, -1
	frm1.vspdData.ReDraw = True
End Sub	

'========================================================================================
' Function Name : GetSpreadColumnPos
' Description   : 
'========================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
	Dim iCurColumnPos
	
	Select Case UCase(pvSpdNo)
		Case "A"
			ggoSpread.Source = frm1.vspdData
			
			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
            
            C_Po_No  		= iCurColumnPos(1)
			C_Po_Type 		= iCurColumnPos(2)
			C_Po_TypeNm		= iCurColumnPos(3)
			C_Release 		= iCurColumnPos(4)
			C_SupplierCd 	= iCurColumnPos(5)
			C_SupplierNm 	= iCurColumnPos(6)
			C_PoDt 		    = iCurColumnPos(7)
			C_GroupCd       = iCurColumnPos(8)
			C_GroupNm		= iCurColumnPos(9)
	
	End Select

End Sub	

'========================================================================================
' Function Name : PopSaveSpreadColumnInf
' Description   : 
'========================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

'========================================================================================
' Function Name : PopRestoreSpreadColumnInf
' Description   : 
'========================================================================================
Sub PopRestoreSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    frm1.vspdData.Redraw = False
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()
	Call ggoSpread.ReOrderingSpreadData()
End Sub

'========================================================================================
' Function Name : FncSplitColumn
' Function Desc : 
'========================================================================================
Function FncSplitColumn()
   If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Function
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)
End Function

<% '#########################################################################################################
'												3. Event�� 
'	���: Event �Լ��� ���� ó�� 
'	����: Windowó��, Singleó��, Gridó�� �۾�.
'         ���⼭ Validation Check, Calcuration �۾��� ������ Event�� �߻�.
'         �� Object������ Grouping�Ѵ�.
'##########################################################################################################%>
<% '******************************************  3.1 Window ó��  *********************************************
'	Window�� �߻� �ϴ� ��� Even ó��	
'********************************************************************************************************* %>
<% '==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(���� Include ���Ͽ� ����)�� �����ʱ�ȭ �� ȭ���ʱ�ȭ�� �ϱ� ���� �Լ��� Call�ϴ� �κ� 
'========================================================================================================= %>
Sub Form_Load()
    Call LoadInfTB19029													'��: Load table , B_numeric_format
	
    'Html���� tag ���ڰ� 1�� 2�� �����ϴ� �κ� ����Format
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)

	Call ggoOper.LockField(Document, "N")                         '��: Lock  Suitable  Field
    
	Call InitVariables											  '��: Initializes local global variables
	Call SetDefaultVal	
	Call InitSpreadSheet()
	Call MM_preloadImages("../../../CShared/image/Query.gif","../../../CShared/image/OK.gif","../../../CShared/image/Cancel.gif")
	
	Call FncQuery()
	
End Sub
<%
'=========================================  3.1.2 Form_QueryUnload()  ===================================
'   Event Name : Form_QueryUnload																		=
'   Event Desc :																						=
'========================================================================================================
%>
	Sub Form_QueryUnload(Cancel, UnloadMode)
	   
	End Sub	

'======================================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'=======================================================================================================
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
    ElseIf KeyAscii = 13 Then
		Call FncQuery
	End if
End Sub

Sub txtToPoDt_Keypress(KeyAscii)
	On Error Resume Next
	If KeyAscii = 27 Then
		Call CancelClick()
    ElseIf KeyAscii = 13 Then
		Call FncQuery		
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

Function vspdData_KeyPress(KeyAscii)
	On Error Resume Next
	If KeyAscii = 13 And frm1.vspdData.ActiveRow > 0 Then
		Call OKClick()
	ElseIf KeyAscii = 27 Then
		Call CancelClick()
	End If
End Function

'======================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : �÷��� Ŭ���� ��� �߻� 
'======================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
   gMouseClickStatus = "SPC"
   Set gActiveSpdSheet = frm1.vspdData   
    
    If frm1.vspdData.MaxRows <= 0 Then Exit Sub
    
    If Row = 0 Then
        ggoSpread.Source = frm1.vspdData
        If lgSortKey = 1 Then
            ggoSpread.SSSort Col
            lgSortKey = 2
        Else
            ggoSpread.SSSort Col, lgSortKey
            lgSortKey = 1
        End If
        Exit Sub
    End If
    
    gMouseClickStatus = "SPC"
	
	'If Row < 1 Then Exit Sub
	
	IscookieSplit = ""
	
	Dim ii
    
     frm1.vspdData.Col = C_Po_No
     frm1.vspdData.Row = Row
	 IscookieSplit = IscookieSplit & Trim(frm1.vspdData.text) & PopupParent.gRowSep
	 Call SetPopupMenuItemInf("0000111111")  
End Sub

'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    
    Call GetSpreadColumnPos("A")
End Sub

'=======================================================================================================
'   Event Name : vspdData_MouseDown
'   Event Desc : 
'=======================================================================================================
Sub vspdData_MouseDown(Button, Shift, X, Y)
	If Button = 2 And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If
End Sub

'========================================================================================================
'   Event Name : vspdData_LeaveCell
'   Event Desc :
'========================================================================================================
Sub vspdData_LeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
	With frm1.vspdData
		If Row >= NewRow Then
			Exit Sub
		End If

		If NewRow = .MaxRows Then
			If lgStrPrevKeyIndex <> "" Then							'��: ���� Ű ���� ������ �� �̻� ��������ASP�� ȣ������ ���� 
				'DbQuery
				If DbQuery = False Then
					Call RestoreToolBar()
					Exit Sub
				End If
			End If
		End If
	End With
End Sub

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
    
    FncQuery = False                                                        '��: Processing is NG
    
    Err.Clear                                                               '��: Protect system from crashing
	
    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")	         						'��: Clear Contents  Field
    Call InitVariables 														'��: Initializes local global variables
    
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then								'��: This function check indispensable field
       Exit Function
    End If
	
	with frm1
		if (UniConvDateToYYYYMMDD(.txtFrPoDt.text,PopupParent.gDateFormat,"") > UniConvDateToYYYYMMDD(.txtToPoDt.text,PopupParent.gDateFormat,"")) And Trim(.txtFrPoDt.text) <> "" And Trim(.txtToPoDt.text) <> "" then	
			Call DisplayMsgBox("17a003", "X","������", "X")	
			
			Exit Function
		End if   
	End with

	'-----------------------
    'Query function call area
    '-----------------------
	If frm1.rdoPostFlg2.checked = True Then
		frm1.hdtxtRadio.value = "Y"
	ElseIf frm1.rdoPostFlg3.checked = True Then
		frm1.hdtxtRadio.value = "N"
	End If
	
    '-----------------------
    'Query function call area
    '-----------------------	
	If DbQuery = False Then Exit Function									

    FncQuery = True		
    
End Function

<% '*******************************  5.2 Fnc�Լ����� ȣ��Ǵ� ���� Function  *******************************
'	���� : 
'********************************************************************************************************* %>


<%
'========================================================================================
' Function Name : DbDelete
' Function Desc : This function delete data
'========================================================================================
%>
<%
'========================================================================================
' Function Name : DbDeleteOk
' Function Desc : DbDelete�� �������϶� ���� 
'========================================================================================
%>
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
		    strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001
		    strVal = strVal & "&txtPotypeCd=" & .hdnPotype.value
		    strVal = strVal & "&txtSupplierCd=" & .hdnSupplier.value
			strVal = strVal & "&txtFrPoDt=" & .hdnFrDt.value
			strVal = strVal & "&txtToPoDt=" & .hdnToDt.value
		    strVal = strVal & "&txtGroupCd=" & .hdnGroup.value
		    strVal = strVal & "&txtRadio=" & Trim(.hdtxtRadio.value) '13�� �߰�	
		    strVal = strVal & "&hdnRetFlg=" & Trim(.hdnRetFlg.value) '��ǰ���� �߰� 
		    strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey     
		else
		    strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001
		    strVal = strVal & "&txtPotypeCd=" & Trim(.txtPotypeCd.value)
		    strVal = strVal & "&txtSupplierCd=" & Trim(.txtSupplierCd.value)
			strVal = strVal & "&txtFrPoDt=" & Trim(.txtFrPoDt.text)
			strVal = strVal & "&txtToPoDt=" & Trim(.txtToPoDt.text)
		    strVal = strVal & "&txtGroupCd=" & Trim(.txtGroupCd.Value)
		    strVal = strVal & "&txtRadio=" & Trim(.hdtxtRadio.value) '13�� �߰�	
		    strVal = strVal & "&hdnRetFlg=" & Trim(.hdnRetFlg.value) '��ǰ���� �߰� 
		    strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey     
		end if 
	
        strVal = strVal & "&lgPageNo="		 & lgPageNo						'��: Next key tag 
        strVal = strVal & "&lgMaxCount="     & C_SHEETMAXROWS_D             '��: �ѹ��� �����ü� �ִ� ����Ÿ �Ǽ�  
		
       
        Call RunMyBizASP(MyBizASP, strVal)		    						'��: �����Ͻ� ASP �� ���� 
        
    End With
    
    DbQuery = True    

End Function

<%
'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================
%>
Function DbQueryOk()	    												'��: ��ȸ ������ ������� 

	lgIntFlgMode = PopupParent.OPMD_UMODE
	
	If frm1.vspdData.MaxRows > 0 Then
		frm1.vspdData.Focus
		frm1.vspdData.Row = 1	
		frm1.vspdData.SelModeSelected = True		
	End If

End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->
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
						<TD CLASS="TD5" NOWRAP>�̵�����</TD>
						<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT AlT="�̵�����" NAME="txtPotypeCd" MAXLENGTH=5 SIZE=10 tag="1XNXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnGroupCd2" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPotype()">
											   <INPUT TYPE=TEXT AlT="�̵�����" NAME="txtPotypeNm" SIZE=20 tag="24X" ></TD>
						<TD CLASS="TD5" NOWRAP>����ó</TD>
						<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT AlT="����ó" NAME="txtSupplierCd" SIZE=10 MAXLENGTH=10 tag="1XNXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnGroup1Cd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenSupplier()">
											   <INPUT TYPE=TEXT AlT="����ó" ID="txtSupplierNm" NAME="arrCond" tag="14X"></TD>
					</TR>	
					<TR>	
						<TD CLASS="TD5" NOWRAP>�����</TD>
						<TD CLASS="TD6" NOWRAP>
							<table cellspacing=0 cellpadding=0>
								<tr>
									<td>
										<script language =javascript src='./js/m9111pa1_fpDateTime1_txtFrPoDt.js'></script>
									</td>
									<td>~</td>
									<td>
										<script language =javascript src='./js/m9111pa1_fpDateTime1_txtToPoDt.js'></script>
									</td>
								<tr>
							</table>
						</TD>
						<TD CLASS="TD5" NOWRAP>���ű׷�</TD>
						<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT AlT="���ű׷�" NAME="txtGroupCd" SIZE=10 MAXLENGTH=4 tag="1XNXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnGroup1Cd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenGroup()">
											   <INPUT TYPE=TEXT AlT="���ű׷�" ID="txtGroupNm" NAME="arrCond" tag="14X"></TD>
					</TR>
					<TR>
						<TD CLASS=TD5 NOWRAP>Ȯ������</TD> 
						<TD CLASS=TD6 colspan=3 NOWRAP>
							<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoPostFlg" TAG="11X" VALUE=""  ID="rdoPostFlg1"><LABEL FOR="rdoPostFlg1">��ü</LABEL>&nbsp;&nbsp;&nbsp;
							<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoPostFlg" TAG="11X" VALUE="Y" ID="rdoPostFlg2"><LABEL FOR="rdoPostFlg2">Ȯ��</LABEL>&nbsp;&nbsp;&nbsp;
							<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoPostFlg" TAG="11X" VALUE="N" CHECKED ID="rdoPostFlg3"><LABEL FOR="rdoPostFlg3">��Ȯ��</LABEL>			
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
						<script language =javascript src='./js/m9111pa1_vaSpread1_vspdData.js'></script>
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
					<TD WIDTH=70% NOWRAP>     <IMG SRC="../../../CShared/image/query_d.gif"  Style="CURSOR: hand" ALT="Search" NAME="Search" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Query.gif',1)"  ONCLICK="FncQuery()"   ></IMG></TD>
					<TD WIDTH=30% ALIGN=RIGHT><IMG SRC="../../../CShared/image/ok_d.gif"     Style="CURSOR: hand" ALT="OK"     NAME="pop1"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/OK.gif',1)"     ONCLICK="OkClick()"    ></IMG>
							                  <IMG SRC="../../../CShared/image/cancel_d.gif" Style="CURSOR: hand" ALT="CANCEL" NAME="pop2"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)" ONCLICK="CancelClick()"></IMG>					</TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=No noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="hdnPotype" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnSupplier" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnFrDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnToDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnGroup" tag="24">
<INPUT TYPE=HIDDEN NAME="hdtxtRadio" TAG="24">
<INPUT TYPE=HIDDEN NAME="hdnRetFlg" TAG="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT" STYLE="visible:true">
	<IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm"></IFRAME>
</DIV>
</BODY>
</HTML>
