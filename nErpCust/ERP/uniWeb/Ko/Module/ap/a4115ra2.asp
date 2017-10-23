
<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : Template
*  2. Function Name        : 
*  3. Program ID           : a4115ra2
*  4. Program Name         : �ϰ���ݵ��-�������� 
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
<SCRIPT LANGUAGE = "JavaScript"SRC = "../../inc/incImage.js">				</SCRIPT>

<Script Language="VBScript">
Option Explicit                                                        '��: indicates that All variables must be declared in advance


'########################################################################################################
'#                       4.  Data Declaration Part
'########################################################################################################

'========================================================================================================
'=                       4.1 External ASP File
'========================================================================================================


'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim lgIsOpenPop                                          
Dim IsOpenPop  
Dim  arrReturn
Dim  arrParent
Dim  arrParam

Dim DueFg

' ���Ѱ��� �߰� 
Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' ����� 
Dim lgInternalCd, lgDeptCd, lgDeptNm			' ���κμ� 
Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' ���κμ�(��������)
Dim lgAuthUsrID, lgAuthUsrNm					' ���� 

'------ Set Parameters from Parent ASP ------ 
arrParent       = window.dialogArguments
Set PopupParent = arrParent(0)	 
arrParam		= arrParent(1)

Dim strYear, strMonth, strDay, dtToday, EndDate, StartDate

<%  
	Dim dtToday 
	dtToday = GetSvrDate 
%>	

Call ExtractDateFrom("<%=dtToday%>", PopupParent.gServerDateFormat, PopupParent.gServerDateType, strYear, strMonth, strDay)

StartDate = UniConvYYYYMMDDToDate(PopupParent.gDateFormat, strYear, strMonth, strDay)

'top.document.title = PopupParent.gActivePRAspName		
top.document.title = "��������"

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

    lgIntFlgMode     = PopupParent.OPMD_CMODE                          'Indicates that current mode is Create mode
	lgBlnFlgChgValue = False

	Self.Returnvalue = arrReturn

	' ���Ѱ��� �߰� 
	If UBound(arrParam,1) > 7 Then
		lgAuthBizAreaCd		= arrParam(8,0)
		lgInternalCd		= arrParam(9,0)
		lgSubInternalCd		= arrParam(10,0)
		lgAuthUsrID			= arrParam(11,0)
	End If	
End Sub

'========================================================================================================
' Name : LoadInfTB19029()	
' Desc : Set System Number format
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "A","NOCOOKIE","RA") %>                                '��: 
End Sub

'========================================================================================================
' Name : SetDefaultVal()	
' Desc : Set default value
'========================================================================================================
Sub SetDefaultVal()
	frm1.txtDueDt.text	 = StartDate
	frm1.txtNoteDueDt.value	 = frm1.txtDueDt.text
	frm1.txtDocCur.value = PopupParent.gCurrency
	
	Call ggoOper.SetReqAttr(frm1.txtNoteDueDt,  "Q")

	If arrParam(0,0) <> ""	Then frm1.txtDueDt.text = arrParam(0,0) 
	If arrParam(0,1) <> ""	Then frm1.txtDocCur.value = arrParam(0,1) 
	If arrParam(1,0) <> ""	Then frm1.txtPayBpCd.value = arrParam(1,0)		: frm1.txtPayBpNm.value=arrParam(1,1)
	If arrParam(2,0) <> ""	Then frm1.txtInputType.value = arrParam(2,0)	: frm1.txtInputTypeNm.value=arrParam(2,1)
	If arrParam(3,0) <> ""	Then frm1.txtBankCd.value = arrParam(3,0)		': frm1.txtBankNm.value=arrParam(3,1)
'	If arrParam(4,0) <> ""	Then frm1.txtBankAcct.value = arrParam(4,0)
'	If arrParam(5,0) <> ""	Then frm1.txtCheckCd.value = arrParam(5,0)	
	If Not arrParam(6,0)	Then frm1.Rb_IntVotl1.checked = False				:	 frm1.Rb_IntVotl2.checked=true

	Call txtInputType_OnChange()
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
	Redim arrReturn(7,1)

	If Not chkField(Document, "1") Then									         '��: This function check required field
		Exit Function
    End If

	If Trim(frm1.txtPayBpCd.value) <> "" then
		If CommonQueryRs( "BP_NM" , "B_BIZ_PARTNER", " BP_CD = " & FilterVar(frm1.txtPayBpCd.value, "''", "S")  _
				, lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then 
			arrReturn(1,0) = Trim(frm1.txtPayBpCd.value)						
			arrReturn(1,1) = Replace(lgF0,Chr(11),"")
		Else
			Call DisplayMsgBox("971001", "X", Trim(frm1.txtPayBpCd.alt), "X")  
			Exit Function
		End if
	End If
	
	If Trim(frm1.txtInputType.value) <> "" then	
		If CommonQueryRs("B_MINOR.MINOR_NM" , "B_MINOR,B_CONFIGURATION", "B_MINOR.MINOR_CD = B_CONFIGURATION.MINOR_CD And B_MINOR.MAJOR_CD = " & FilterVar("A1006", "''", "S") & " " & _
				" AND B_CONFIGURATION.SEQ_NO = 2 AND B_CONFIGURATION.REFERENCE = " & FilterVar("PP", "''", "S") & "  AND B_MINOR.MINOR_CD= " & FilterVar(frm1.txtInputType.value, "''", "S")  _
				, lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then 
			arrReturn(2,0) = Trim(frm1.txtInputType.value)	
			arrReturn(2,1) = Replace(lgF0,Chr(11),"")						
		Else
			Call DisplayMsgBox("971001", "X", Trim(frm1.txtInputType.alt), "X")  
			Exit Function
		End If
	End If	
	
	If Trim(frm1.txtBizAreaCd.value) <> "" then
		If CommonQueryRs( "BIZ_AREA_NM" , "B_BIZ_AREA", " BIZ_AREA_CD = " & FilterVar(frm1.txtBizAreaCd.value, "''", "S")  _
				, lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then 
			arrReturn(3,0) = Trim(frm1.txtBizAreaCd.value)						
			arrReturn(3,1) = Replace(lgF0,Chr(11),"")
		Else
			Call DisplayMsgBox("971001", "X", Trim(frm1.txtBizAreaCd.alt), "X")  
			Exit Function
		End if
	End If
	
	If Trim(frm1.txtBizAreaCd1.value) <> "" then
		If CommonQueryRs( "BIZ_AREA_NM" , "B_BIZ_AREA", " BIZ_AREA_CD = " & FilterVar(frm1.txtBizAreaCd1.value, "''", "S")  _
				, lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then 
			arrReturn(4,0) = Trim(frm1.txtBizAreaCd1.value)						
			arrReturn(4,1) = Replace(lgF0,Chr(11),"")
		Else
			Call DisplayMsgBox("971001", "X", Trim(frm1.txtBizAreaCd1.alt), "X")  
			Exit Function
		End if
	End If
		
'	If Trim(frm1.txtBankCd.value) <> "" then
'		If CommonQueryRs( "B_BANK.BANK_NM" , "B_BANK", "B_BANK.BANK_CD = " & FilterVar(frm1.txtBankCd.value, "''", "S") _
'				, lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then 
'			arrReturn(3,0) = Trim(frm1.txtBankCd.value)							
'			arrReturn(3,1) = Replace(lgF0,Chr(11),"")
'		Else
'			Call DisplayMsgBox("971001", "X", Trim(frm1.txtBankCd.alt) , "X")  
'			Exit Function
'		End If
'	End If
	
	arrReturn(0,0) = Trim(frm1.txtDueDt.text)
	arrReturn(0,1) = Trim(frm1.txtDocCur.value)	
	arrReturn(1,0) = Trim(frm1.txtPayBpCd.value)		
	arrReturn(2,0) = Trim(frm1.txtInputType.value)
	arrReturn(3,0) = Trim(frm1.txtBizAreaCd.value)	
	arrReturn(4,0) = Trim(frm1.txtBizAreaCd1.value)		
'	arrReturn(4,0) = Trim(frm1.txtBankAcct.value)	
'	arrReturn(5,0) = Trim(frm1.txtCheckCd.value)
	arrReturn(5,0) = ""
	arrReturn(6,0) = frm1.Rb_IntVotl1.checked
'	arrReturn(7,0) = Trim(frm1.txtNoteDueDt.value)	

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
 '------------------------------------------  OpenDept()  ---------------------------------------
'	Name : OpenBp()
'	Description : OpenAp Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function OpenBp(Byval strCode, byval iWhere)
	Dim arrRet
	Dim arrParam(5)

	If IsOpenPop = True Then Exit Function
	If frm1.txtPayBpCd.className = "protected" Then Exit Function
	
	IsOpenPop = True

	arrParam(0) = strCode								'  Code Condition
   	arrParam(1) = ""							' ä�ǰ� ����(�ŷ�ó ����)
	arrParam(2) = ""								'FrDt
	arrParam(3) = ""								'ToDt
	arrParam(4) = "S"							'B :���� S: ���� T: ��ü 
	arrParam(5) = "PAYTO"									'SUP :����ó PAYTO: ����ó SOL:�ֹ�ó PAYER :����ó INV:���ݰ�� 	
	arrRet = window.showModalDialog("../../Module/ar/BpPopup.asp", Array(window.PopupParent,arrParam), _
		"dialogWidth=780px; dialogHeight=450px; : Yes; help: No; resizable: No; status: No;")
		
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Call EscPopup(iWhere)	    
		Exit Function
	Else
		Call SetPopup(arrRet, iWhere)
	End If	
End Function
 '======================================================================================================
'   Function Name : OpenPopUp()
'   Function Desc : 
'=======================================================================================================
Function  OpenPopUp(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim strCd	
	
	If IsOpenPop = True Then Exit Function

	Select Case iWhere
		Case 1
			If frm1.txtPayBpCd.className = "protected" Then Exit Function
			
			arrParam(0) = "����ó�˾�"
			arrParam(1) = "B_BIZ_PARTNER"				
			arrParam(2) = Trim(frm1.txtPayBpCd.Value)
			arrParam(3) = ""
			arrParam(4) = ""
			arrParam(5) = "����ó"			
	
			arrField(0) = "BP_CD"	
			arrField(1) = "BP_NM"	
    
			arrHeader(0) = "����ó"		
			arrHeader(1) = "����ó��"	
		Case 3		
			If frm1.txtDocCur.className = "protected" Then Exit Function
			
			arrParam(0) = "�ŷ���ȭ�˾�"			' �˾� ��Ī 
			arrParam(1) = "B_CURRENCY"					' TABLE ��Ī 
			arrParam(2) = Trim(frm1.txtDocCur.Value)		' Code Condition
			arrParam(3) = ""								' Name Cindition
			arrParam(4) = ""								' Where Condition
			arrParam(5) = "�ŷ���ȭ"			
	
			arrField(0) = "CURRENCY"							' Field��(0)
			arrField(1) = "CURRENCY_DESC"							' Field��(1)
    
			arrHeader(0) = "�ŷ���ȭ"					' Header��(0)
			arrHeader(1) = "�ŷ���ȭ��"
	
		
		Case 8 
			If frm1.txtInputType.className = "protected" Then Exit Function    
			
			arrParam(0) = "��������"					' �˾� ��Ī						
			arrParam(1) = "B_MINOR,B_CONFIGURATION "
			arrParam(2) = Trim(frm1.txtInputType.value)		' Code Condition
			arrParam(3) = ""								' Name Cindition
			arrParam(4) = "B_MINOR.MINOR_CD = B_CONFIGURATION.MINOR_CD And B_MINOR.MAJOR_CD = " & FilterVar("A1006", "''", "S") & "  " _
						& "AND B_CONFIGURATION.SEQ_NO = 2 AND B_CONFIGURATION.REFERENCE = " & FilterVar("PP", "''", "S") & " "	' Where Condition					
			arrParam(5) = "��������"					' TextBox ��Ī 
		
			arrField(0) = "B_MINOR.MINOR_CD"				' Field��(0)
			arrField(1) = "B_MINOR.MINOR_NM"				' Field��(1)
	    
			arrHeader(0) = "��������"					' Header��(0)
			arrHeader(1) = "����������"					' Header��(1)		
		
		Case Else				    
			Exit Function
	End Select				
		
	IsOpenPop = True
	
	arrRet = window.showModalDialog("../../comasp/adocommonpopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")			
		
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Call EscPopup(iWhere)	    
		Exit Function
	Else
		Call SetPopup(arrRet, iWhere)
	End If
End Function
'======================================================================================================
'   Function Name : SetPopup(Byval arrRet)
'   Function Desc : 
'=======================================================================================================
Function EscPopup(Byval iWhere)
	With frm1
		Select Case iWhere
			Case 1	
				.txtPayBpCd.focus
			Case 3
				.txtDocCur.focus
			Case 8
				.txtInputType.focus		 	
		End Select				
	End With

	If iwhere  <> 0 Then
		lgBlnFlgChgValue = True
	End If	
End Function

'======================================================================================================
'   Function Name : SetPopup(Byval arrRet)
'   Function Desc : 
'=======================================================================================================
Function SetPopup(Byval arrRet,Byval iWhere)
	With frm1
		Select Case iWhere
			Case 1	
				.txtPayBpCd.value = arrRet(0)		
				.txtPayBpNm.value = arrRet(1)
				.txtPayBpCd.focus
			Case 3
				.txtDocCur.value = arrRet(0)		
				
				Call txtDocCur_OnChange()
				.txtDocCur.focus
			Case 8
				.txtInputType.value = arrRet(0)		 	
				.txtInputTypeNm.value = arrRet(1)
				.txtInputType.focus		 	
		End Select				
	End With

	If iwhere  <> 0 Then
		lgBlnFlgChgValue = True
	End If	
End Function


'----------------------------------------  OpenBizAreaCd()  -------------------------------------------------
'	Name : BizAreaCd()
'	Description : Business PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenBizAreaCd()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	arrParam(0) = "����� �˾�"				' �˾� ��Ī 
	arrParam(1) = "B_BIZ_AREA"					' TABLE ��Ī 
	arrParam(2) = Trim(frm1.txtBizAreaCd.Value)	' Code Condition
	arrParam(3) = ""							' Name Cindition
	' ���Ѱ��� �߰� 
	If lgAuthBizAreaCd <> "" Then
		arrParam(4) = " BIZ_AREA_CD = " & FilterVar(lgAuthBizAreaCd, "''", "S")			' Where Condition
	Else
		arrParam(4) = ""
	End If

	arrParam(5) = "����� �ڵ�"			

    arrField(0) = "BIZ_AREA_CD"					' Field��(0)
    arrField(1) = "BIZ_AREA_NM"					' Field��(1)

    arrHeader(0) = "������ڵ�"				' Header��(0)
	arrHeader(1) = "������"				' Header��(1)

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If arrRet(0) = "" Then
'		frm1.txtBizAreaCd.focus
		Exit Function
	Else
		Call SetReturnVal(arrRet,1)
	End If
End Function


'----------------------------------------  OpenBizAreaCd()  -------------------------------------------------
'	Name : BizAreaCd()
'	Description : Business PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenBizAreaCd1()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	arrParam(0) = "����� �˾�"				' �˾� ��Ī 
	arrParam(1) = "B_BIZ_AREA"					' TABLE ��Ī 
	arrParam(2) = Trim(frm1.txtBizAreaCd1.Value)	' Code Condition
	arrParam(3) = ""							' Name Cindition
	' ���Ѱ��� �߰� 
	If lgAuthBizAreaCd <> "" Then
		arrParam(4) = " BIZ_AREA_CD = " & FilterVar(lgAuthBizAreaCd, "''", "S")			' Where Condition
	Else
		arrParam(4) = ""
	End If
	arrParam(5) = "����� �ڵ�"			

    arrField(0) = "BIZ_AREA_CD"					' Field��(0)
    arrField(1) = "BIZ_AREA_NM"					' Field��(1)

    arrHeader(0) = "������ڵ�"				' Header��(0)
	arrHeader(1) = "������"				' Header��(1)

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If arrRet(0) = "" Then
'		frm1.txtBizAreaCd.focus
		Exit Function
	Else
		Call SetReturnVal(arrRet,2)
	End If
End Function



'=======================================================================================================
'	Name : SetReturnVal()
'	Description : 
'=======================================================================================================
Function SetReturnVal(byval arrRet,Field_fg)
	Select Case Field_fg
		case 1
			frm1.txtBizAreaCd.Value	= arrRet(0)
			frm1.txtBizAreaNm.Value	= arrRet(1)
			frm1.txtBizAreaCd.focus
		case 2
			frm1.txtBizAreaCd1.Value = arrRet(0)
			frm1.txtBizAreaNm1.Value = arrRet(1)
			frm1.txtBizAreaCd1.focus
	End Select
	
	lgBlnFlgChgValue = True
End Function

'=======================================================================================================
'   Event Name : txtInputType_OnChange()
'   Event Desc :  
'=======================================================================================================
Sub txtInputType_OnChange()
	Dim IntRetCD

    lgBlnFlgChgValue = True
	
	' SetReqAttr(Object, Option) ; N : Required, Q : Protect, D : Default
	
	If CommonQueryRs( "REFERENCE" , "B_CONFIGURATION " , "MAJOR_CD = " & FilterVar("A1006", "''", "S") & "  AND MINOR_CD = " & FilterVar(frm1.txtInputType.value, "''", "S") & " AND SEQ_NO = 4 ", lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then 
		Select Case UCase(lgF0)
			Case "CS" & Chr(11)
'				frm1.txtCheckCd.value   = ""
				frm1.txtBankCd.value   = ""  ' : frm1.txtBankNm.value=""
				frm1.txtBankAcct.value   = ""
				Call ggoOper.SetReqAttr(frm1.txtBankCd,   "Q")
				Call ggoOper.SetReqAttr(frm1.txtBankAcct, "Q")
'				Call ggoOper.SetReqAttr(frm1.txtCheckCd,   "Q")
				Call ggoOper.SetReqAttr(frm1.txtPayBpCd,   "D")	
			Case "DP" & Chr(11)			' ������ 
'				frm1.txtCheckCd.value   = ""
				Call ggoOper.SetReqAttr(frm1.txtBankCd,   "N")
				Call ggoOper.SetReqAttr(frm1.txtBankAcct, "N")
'				Call ggoOper.SetReqAttr(frm1.txtCheckCd,   "Q")
				Call ggoOper.SetReqAttr(frm1.txtPayBpCd,   "D")	
			Case "NO" & Chr(11)		'���� 
'				frm1.txtBankCd.value   = "" : frm1.txtBankNm.value=""
				frm1.txtBankAcct.value   = ""				
				Call ggoOper.SetReqAttr(frm1.txtBankCd,   "N")
'				Call ggoOper.SetReqAttr(frm1.txtBankAcct, "N")
'				Call ggoOper.SetReqAttr(frm1.txtCheckCd,   "N")	
'				Call ggoOper.SetReqAttr(frm1.txtPayBpCd,   "N")
				Call ggoOper.SetReqAttr(frm1.txtNoteDueDt,  "N")

			Case Else
				IntRetCD = DisplayMsgBox("141140","X","X","X")
				
				frm1.txtInputType.value = ""
				frm1.txtInputTypeNm.value = ""
				
				Exit Sub
			
'				frm1.txtCheckCd.value   = ""
'				frm1.txtBankCd.value   = "" : frm1.txtBankNm.value=""
'				frm1.txtBankAcct.value   = ""		
'				Call ggoOper.SetReqAttr(frm1.txtBankCd,   "Q")
'				Call ggoOper.SetReqAttr(frm1.txtBankAcct, "Q")
'				Call ggoOper.SetReqAttr(frm1.txtCheckCd,   "Q")
'				Call ggoOper.SetReqAttr(frm1.txtPayBpCd,   "D")	
		End Select
	End If
End Sub

'==========================================================================================
'   Event Name : txtDocCur_OnChange
'   Event Desc : 
'==========================================================================================
Sub txtDocCur_OnChange()
    lgBlnFlgChgValue = True
    If CommonQueryRs( "CURRENCY_DESC" , "B_CURRENCY" ,  " CURRENCY =  " & FilterVar(frm1.txtDocCur.value , "''", "S") & "" , lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then   			        							
    
	End If	    
End Sub

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

End Function

'========================================================================================================
' Name : FncCancel
' Desc : This function is called by MainCancel in Common.vbs
'========================================================================================================
Function FncCancel() 
    FncCancel = False                                                            '��: Processing is NG
    Err.Clear                                                                    '��: Clear err status

    Set gActiveElement = document.ActiveElement   
    FncCancel = False                                                            '��: Processing is OK
End Function

'*********************************************  3.3 Object Tag ó��  ************************************
'*	Object���� �߻� �ϴ� Event ó��																		*
'********************************************************************************************************
Function Radio1_onChange()									'ȯ���������� 
	lgBlnFlgChgValue = True
	DueFg= True
End Function

Function Radio2_onChange()
	lgBlnFlgChgValue = True
	DueFg= False
End Function
'========================================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
End Sub

'========================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : �÷��� Ŭ���� ��� �߻� 
'========================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)

End Sub

'========================================================================================================
'   Event Name : vspdData_DblClick
'   Event Desc :
'========================================================================================================
Sub  vspdData_DblClick(ByVal Col, ByVal Row)

End Sub

'########################################################################################################
'#					     4. Common Function��															#
'########################################################################################################

'=======================================================================================================
'   Event Name : txtDueDt_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub  txtDueDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtDueDt.Action = 7     
        Call SetFocusToDocument("P")
		Frm1.txtDueDt.Focus                   
    End If
End Sub

'=======================================================================================================
'   Event Name : txtNoteDueDt_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub  txtNoteDueDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtNoteDueDt.Action = 7     
        Call SetFocusToDocument("P")
		Frm1.txtNoteDueDt.Focus                   
    End If
End Sub

'=======================================================================================================
'   Event Name : txtApDt_Change()
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub  txtApDt_Change()
    
End Sub

Sub txtDueDt_KeyPress(KeyAscii)
	On Error Resume Next
	If KeyAscii = 27 Then 
		Call CancelClick()
	ElseIf KeyAscii = 13 Then 
		Call okclick()
	End If
End Sub

Sub txtNoteDueDt_KeyPress(KeyAscii)
	On Error Resume Next
	If KeyAscii = 27 Then 
		Call CancelClick()
	ElseIf KeyAscii = 13 Then 
		Call okclick()
	End If
End Sub

Sub txtBizAreaCd_KeyPress(KeyAscii)
	On Error Resume Next
	If KeyAscii = 27 Then 
		Call CancelClick()
	ElseIf KeyAscii = 13 Then 
		Call okclick()
	End If
End Sub

Sub txtBizAreaCd1_KeyPress(KeyAscii)
	On Error Resume Next
	If KeyAscii = 27 Then 
		Call CancelClick()
	ElseIf KeyAscii = 13 Then 
		Call okclick()
	End If
End Sub

Function document_onkeypress()
	If window.event.keyCode = 13 Then
       Call okClick()
    End If
End Function




'########################################################################################################
'#						5. Interface ��																	#
'########################################################################################################


'********************************************  5.1 DbQuery()  *******************************************
' Function Name : DbQuery																				*
' Function Desc : This function is data query and display												*
'********************************************************************************************************
Function DbQuery()

End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================
Function DbQueryOk()												

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
						<TD CLASS=TD5 NOWRAP>������</TD>
						<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME="txtDueDt" CLASS=FPDTYYYYMMDD tag="12" Title="FPDATETIME" ALT="������" id=fpDateTime></OBJECT>');</SCRIPT>&nbsp
											<INPUT TYPE="RADIO" CLASS="Radio" NAME=Radio_IntVotl ID=Rb_IntVotl1 Checked tag = 2 value="X" onclick=radio1_onchange()><LABEL FOR=Rb_IntVotl1>=</LABEL>&nbsp;
														   <INPUT TYPE="RADIO" CLASS="Radio" NAME=Radio_IntVotl ID=Rb_IntVotl2 tag = 1 value="F" onclick=radio2_onchange()><LABEL FOR=Rb_IntVotl2><=</LABEL>&nbsp;</TD>
						</TR>
						<TR>						
						<TD CLASS=TD5 NOWRAP>�ŷ���ȭ</TD>
						<TD CLASS=TD6 NOWRAP>
							<INPUT TYPE=TEXT NAME="txtDocCur" SIZE=10 MAXLENGTH=4 tag="13NXXU" STYLE="TEXT-ALIGN: left" ALT="�ŷ���ȭ"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDocCur" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript: CALL OpenPopup(frm1.txtDocCur.value,3)">
						</TD>
					</TR>
					<TR>
						<TD CLASS=TD5 NOWRAP>��������</TD>
						<TD CLASS="TD6" NOWRAP>
							<INPUT TYPE=TEXT NAME="txtInputType" ALT="��������" style="HEIGHT: 20px; WIDTH: 80px" MAXLENGTH=5 tag="11NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" style="HEIGHT: 21px; WIDTH: 16px" NAME="btnPayMethod" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(frm1.txtInputType.value, 8)">
						   <INPUT TYPE=TEXT NAME="txtInputTypeNm" ALT="��������" style="HEIGHT: 20px; WIDTH: 150px" tag="14X" >
						</TD>
						</TR>
						<TR>																	   
						<TD CLASS=TD5 NOWRAP>����ó</TD>
						<TD CLASS=TD6 NOWRAP>
							<INPUT TYPE="Text" NAME="txtPayBpCd" SIZE=10 MAXLENGTH=10 STYLE="TEXT-ALIGN: left" tag="11NXXU" ALT="����ó"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBizCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:CALL OpenBp(frm1.txtPayBpCd.Value, 1)">
							<INPUT TYPE=TEXT NAME="txtPayBpNm" SIZE=20 MAXLENGTH=20 STYLE="TEXT-ALIGN: Left" tag="14" ALT="����ó��">
						</TD>
					</TR>
					<TR>
						<TD CLASS="TD5" NOWRAP>�����</TD>
						<TD CLASS="TD6" NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtBizAreaCd" SIZE=13 MAXLENGTH=10 STYLE="TEXT-ALIGN: Left" tag="11XXXU" ALT="���ۻ����"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBizAreaCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenBizAreaCd()">&nbsp;<INPUT TYPE=TEXT NAME="txtBizAreaNm" SIZE=25 tag="14">&nbsp;~</TD>
					</TR>
					<TR>
						<TD CLASS="TD5" NOWRAP></TD>
						<TD CLASS="TD6" NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtBizAreaCd1" SIZE=13 MAXLENGTH=10 STYLE="TEXT-ALIGN: Left" tag="11XXXU" ALT="��������"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBizAreaCd1" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenBizAreaCd1()">&nbsp;<INPUT TYPE=TEXT NAME="txtBizAreaNm1" SIZE=25 tag="14"></TD>
					</TR>	
<!--					<TR>
						<TD CLASS=TD5 NOWRAP>����</TD>
						<TD CLASS=TD6 NOWRAP>
							<INPUT TYPE="Text" NAME="txtBankCd" SIZE=10 MAXLENGTH=10 STYLE="TEXT-ALIGN: left" tag="11NXXU" ALT="����"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBankCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(frm1.txtBankCd.value,5)">
							<INPUT TYPE=TEXT NAME="txtBankNm" SIZE=20 MAXLENGTH=20 STYLE="TEXT-ALIGN: Left" tag="14" ALT="�����">
						</TD>																				
						<TD CLASS=TD5 NOWRAP>���¹�ȣ</TD>
						<TD CLASS=TD6 NOWRAP>
							<INPUT  TYPE=TEXT NAME="txtBankAcct" SIZE=30 MAXLENGTH=30 STYLE="TEXT-ALIGN: left" tag="11XXXU" ALT="���¹�ȣ"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBankCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(frm1.txtBankAcct.value,6)">
						</TD>
					</TR>
					<TR>
						<TD CLASS=TD5 NOWRAP>����������</TD>
						<TD CLASS=TD6 NOWRAP>
							<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME="txtNoteDueDt" CLASS=FPDTYYYYMMDD tag="12" Title="FPDATETIME" ALT="����������" id=fpDateTime1></OBJECT>');</SCRIPT>&nbsp;
						</TD>
						<TD CLASS="TD5" NOWRAP>&nbsp;</TD>
						<TD CLASS="TD6" NOWRAP>&nbsp;</TD>
					</TR>-->
				</TABLE>
			</FIELDSET>
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
					<TD WIDTH=10>&nbsp;</TD>
					<TD WIDTH=30% ALIGN=RIGHT>
						<IMG SRC="../../../CShared/image/ok_d.gif" Style="CURSOR: hand" ALT="OK" NAME="pop1" ONCLICK="OkClick()" ></IMG>&nbsp;
						<IMG SRC="../../../CShared/image/cancel_d.gif" Style="CURSOR: hand" ALT="CANCEL" NAME="pop2" ONCLICK="CancelClick()" ></IMG>
					</TD>				
			
				</TR>
			</TABLE>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="htxtBizCd"		tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="htxtPayBpCd"	tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="htxtApDt"		tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="htxtToApDt"    tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="htxtDocCur"    tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtBankCd"     tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtBankAcct"   tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtNoteDueDt"  tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="htxtBizAreaCd"		tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="htxtBizAreaCd1"	tag="24" TABINDEX="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

