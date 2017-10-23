<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
'*  1. Module Name          : �λ�/�޿����� 
'*  2. Function Name        : ����������� 
'*  3. Program ID           : h9104oa1
'*  4. Program Name         : �ٷμҵ�����Ű��� 
'*  5. Program Desc         : �ٷμҵ�����Ű��� 
'*  6. Comproxy List        : +
'*  7. Modified date(First) : 2001/05/27
'*  8. Modified date(Last)  : 2003/06/13
'*  9. Modifier (First)     : Shin Kwang-Ho
'* 10. Modifier (Last)      : Lee SiNa
'* 11. Comment              :
=======================================================================================================-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs">     </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs">    </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/IncHRQuery.vbs">   </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/IncCliRdsQuery.vbs">   </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incEB.vbs"></SCRIPT>

<Script Language="VBScript">
Option Explicit
'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim IsOpenPop          
Dim lsInternal_cd

'========================================================================================================
' Name : InitVariables()	
' Desc : Initialize value
'========================================================================================================
Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
End Sub

'========================================================================================================
' Name : SetDefaultVal()	
' Desc : Set default value
'========================================================================================================
Sub SetDefaultVal()
    frm1.txtBas_dt.focus 
	frm1.txtBas_dt.text	= UniConvDateAToB("<%=GetSvrDate%>" ,parent.gServerDateFormat,parent.gDateFormat)
	
End Sub

'========================================================================================================
' Name : LoadInfTB19029()	
' Desc : Set System Number format
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("P", "H","NOCOOKIE","OA") %>
End Sub

'========================================================================================================
' Name : InitComboBox()
' Desc : Set ComboBox
'========================================================================================================
Sub InitComboBox()
    Dim iCodeArr 
    Dim iNameArr
    Dim iDx
    
    Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0003", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    iCodeArr = lgF0
    iNameArr = lgF1
    Call SetCombo2(frm1.cboOcpt_type,iCodeArr, iNameArr,Chr(11))
    
    Call CommonQueryRs("MINOR_CD, MINOR_NM ","B_MINOR","MAJOR_CD = " & FilterVar("H0005", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    iCodeArr = lgF0
    iNameArr = lgF1
    Call SetCombo2(frm1.cboPay_cd, iCodeArr, iNameArr, Chr(11))
    
    Call CommonQueryRs("MINOR_CD, MINOR_NM ","B_MINOR","MAJOR_CD = " & FilterVar("H0035", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    iCodeArr = lgF0
    iNameArr = lgF1
    Call SetCombo2(frm1.cboSect_cd, iCodeArr, iNameArr, Chr(11))
End Sub

'========================================================================================================
' Name : Form_Load
' Desc : developer describe this line Called by Window_OnLoad() evnt
'========================================================================================================
Sub Form_Load()
    Call LoadInfTB19029                                                '��: Load table , B_numeric_format
    Call ggoOper.LockField(Document, "N")                                   '��: Lock  Suitable  Field
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call InitVariables                                        
    Call SetDefaultVal
    Call InitComboBox
    Call FuncGetAuth(gStrRequestMenuID, Parent.gUsrID, lgUsrIntCd)     ' �ڷ����:lgUsrIntCd ("%", "1%")
    Call SetToolBar("1000000000000111")

End Sub
'========================================================================================================
' Name : Form_QueryUnload
' Desc : developer describe this line Called by Window_OnUnLoad() evnt
'========================================================================================================
Sub Form_QueryUnload(Cancel, UnloadMode)
End Sub

'========================================================================================================
' Name : FncQuery
' Desc : developer describe this line Called by MainQuery in Common.vbs
'========================================================================================================
Function FncQuery()
    Dim IntRetCD 
    
    FncQuery = False															 '��: Processing is NG
    Err.Clear                                                                    '��: Clear err status

    If txtEmp_no_Onchange() Then        'enter key �� ��ȸ�� ����� check�� �ش���� ������ query����...
        Exit Function
    End if
    
    If txtFr_dept_cd_Onchange() Then        'enter key �� ��ȸ�� �μ��ڵ带 check�� �ش���� ������ query����...
        Exit Function
    End if
    
    If txtTo_dept_cd_Onchange() Then        'enter key �� ��ȸ�� �μ��ڵ带 check�� �ش���� ������ query����...
        Exit Function
    End if

    FncQuery = True                                                              '��: Processing is OK

End Function
	
'========================================================================================================
' Name : FncDelete
' Desc : developer describe this line Called by MainDelete in Common.vbs
'========================================================================================================
Function FncDelete()
    Dim intRetCD
    
    FncDelete = False                                                            '��: Processing is NG
    Err.Clear                                                                    '��: Clear err status
    
    FncDelete = True                                                             '��: Processing is OK
End Function

'========================================================================================================
' Name : FncCancel
' Desc : developer describe this line Called by MainCancel in Common.vbs
'========================================================================================================
Function FncCancel() 
	On Error Resume Next                                                        '��: Protect system from crashing
End Function
'========================================================================================================
' Name : FncPrint
' Desc : developer describe this line Called by MainDeleteRow in Common.vbs
'========================================================================================================
Function FncPrint()
	Call Parent.FncPrint()                                                      '��: Protect system from crashing
End Function

'========================================================================================================
' Name : FncExcel
' Desc : developer describe this line Called by MainExcel in Common.vbs
'========================================================================================================
Function FncExcel() 
	Call Parent.FncExport(parent.C_SINGLE)
End Function

'========================================================================================================
' Name : FncFind
' Desc : developer describe this line Called by MainFind in Common.vbs
'========================================================================================================
Function FncFind() 
	Call Parent.FncFind(parent.C_SINGLE, False)
End Function

'========================================================================================================
' Name : FncExit
' Desc : developer describe this line Called by MainExit in Common.vbs
'========================================================================================================
Function FncExit()
	Dim IntRetCD

	FncExit = False

	FncExit = True
End Function
'========================================================================================================
' Name : OpenEmptName()
' Desc : developer describe this line 
'========================================================================================================

Function OpenEmptName(iWhere)
	Dim arrRet
	Dim arrParam(2)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	If iWhere = 0 Then 'TextBox(Condition)
		arrParam(0) = frm1.txtEmp_no.value			' Code Condition
	    arrParam(1) = ""'frm1.txtName.value			' Name Cindition
	End If
	arrParam(2) = lgUsrIntCd                    ' �ڷ���� Condition  

	arrRet = window.showModalDialog(HRAskPRAspName("EmpPopup"), Array(window.parent, arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtEmp_no.focus
		Exit Function
	Else
		Call SubSetCondEmp(arrRet, iWhere)
	End If	
			
End Function

'======================================================================================================
'	Name : SubSetCondEmp()
'	Description : Item Popup���� Return�Ǵ� �� setting
'=======================================================================================================
Sub SubSetCondEmp(Byval arrRet, Byval iWhere)
	With frm1
		If iWhere = 0 Then
			.txtEmp_no.value = arrRet(0)
			.txtName.value = arrRet(1)
			.txtEmp_no.focus
		End If
	End With
End Sub
		
'========================================================================================================
' Name : OpenDept
' Desc : �μ� POPUP
'========================================================================================================
Function OpenDept(iWhere)
	Dim arrRet
	Dim arrParam(2)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	If iWhere = 0 Then
		arrParam(0) = frm1.txtFr_dept_cd.value			            '  Code Condition
	ElseIf iWhere = 1 Then
		arrParam(0) = frm1.txtTo_dept_cd.value			            ' Code Condition
	End If
	
	arrParam(1) = frm1.txtBas_dt.text
	arrParam(2) = lgUsrIntCd                    ' �ڷ���� Condition  

	arrRet = window.showModalDialog(HRAskPRAspName("DeptPopupDt"), Array(window.parent,arrParam), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Select Case iWhere
		     Case "0"
               frm1.txtFr_dept_cd.focus
             Case "1"  
               frm1.txtTo_dept_cd.focus
        End Select	
		Exit Function
	Else
		Call SetDept(arrRet, iWhere)
	End If	
			
End Function

'------------------------------------------  SetDept()  ------------------------------------------------
'	Name : SetDept()
'	Description : Dept Popup���� Return�Ǵ� �� setting
'---------------------------------------------------------------------------------------------------------
Function SetDept(Byval arrRet, Byval iWhere)
		
	With frm1
		Select Case iWhere
		     Case "0"
               .txtFr_dept_cd.value = arrRet(0)
               .txtFr_dept_nm.value = arrRet(1)
               .txtFr_internal_cd.value = arrRet(2)
               .txtFr_dept_cd.focus
             Case "1"  
               .txtTo_dept_cd.value = arrRet(0)
               .txtTo_dept_nm.value = arrRet(1) 
               .txtTo_internal_cd.value = arrRet(2) 
               .txtTo_dept_cd.focus
        End Select
	End With
End Function       		

'======================================================================================================
' Function Name : FncBtnPrint
' Function Desc : This function is related to Print Button
'=======================================================================================================

Function FncBtnPrint() 
    Dim StrEbrFile
    Dim strFrDept
    Dim strToDept
    Dim strMin
    Dim strMax
    Dim condvar
    Dim strYear, strMonth, strDay
	strMin="0"
	strMax="ZZZZ"
    If Not chkField(Document, "1") Then									<%'��: This function check indispensable field%>
       Exit Function
    End If
  
    If txtEmp_no_Onchange() Then        'enter key �� ��ȸ�� ����� check�� �ش���� ������ query����...
        Exit Function
    End if
    
    If txtFr_dept_cd_Onchange() Then        'enter key �� ��ȸ�� �μ��ڵ带 check�� �ش���� ������ query����...
        Exit Function
    End if
    
    If txtTo_dept_cd_Onchange() Then        'enter key �� ��ȸ�� �μ��ڵ带 check�� �ش���� ������ query����...
        Exit Function
    End if  
    strFrDept = Trim(frm1.txtFr_internal_cd.value)
    strToDept = Trim(frm1.txtTo_internal_cd.value)
    If strFrDept = "" AND strToDept = "" Then       
    Else
        If strFrDept = "" then
            strFrDept = strMin
        End if
        If strToDept = "" then
            strToDept = strMax
        ElseIf strFrDept > strToDept then
	        Call DisplayMsgbox("800153","X","X","X")	'���ۺμ��� ����μ����� �۾ƾ� �մϴ�.
            frm1.txtFr_dept_cd.focus
            Set gActiveElement = document.activeElement
            Exit Function
        End if 
    End if 


	Dim bas_dt, emp_no, pay_cd, ocpt_type, sect_cd, fr_dept_cd, to_dept_cd  
	
	Call parent.ExtractDateFrom(frm1.txtBas_dt.Text,frm1.txtBas_dt.UserDefinedFormat,parent.gComDateType,strYear,strMonth,strDay)
		
	bas_dt = strYear & strMonth & strDay
	emp_no = frm1.txtEmp_no.value
	pay_cd = frm1.cboPay_cd.value
	ocpt_type = frm1.cboOcpt_type.value
	sect_cd = frm1.cboSect_cd.value
	fr_dept_cd = frm1.txtFr_internal_cd.value
	to_dept_cd = frm1.txtTo_internal_cd.value

	if emp_no = "" then
		emp_no = "%"
		frm1.txtName.value = ""
	End if	

	if ocpt_type = "" then
		ocpt_type = "%"
	End if	

	if pay_cd = "" then
		pay_cd = "%"
	End if	
	
	if sect_cd = "" then
		sect_cd = "%"
	End if	
	
	if fr_dept_cd = "" then
		fr_dept_cd = strMin
		frm1.txtFr_dept_nm.value = ""
	End if	
	
	if to_dept_cd = "" then
		to_dept_cd = strMax
		frm1.txtTo_dept_nm.value = ""
	End if 

	
	ObjName = AskEBDocumentName("h9104oa1","ebr")
	
	condvar =  "bas_dt|" & bas_dt
	condvar = condvar & "|emp_no|" & emp_no
	condvar = condvar & "|pay_cd|" & pay_cd
	condvar = condvar & "|ocpt_type|" & ocpt_type
	condvar = condvar & "|sect_cd|" & sect_cd
	condvar = condvar & "|fr_dept_cd|" & fr_dept_cd
	condvar = condvar & "|to_dept_cd|" & to_dept_cd

    Call FncEBRPrint(EBAction,ObjName,condvar)
    
    Dim ObjName2 , ObjName3
	ObjName2 = AskEBDocumentName("h9104oa1_2","ebr")
	ObjName3 = AskEBDocumentName("h9104oa1_3","ebr")

	condvar = ""

    Call FncEBRPrint(EBAction,ObjName2,condvar)
    'Call FncEBRPrint(EBAction,ObjName3,condvar)    
	
End Function


'========================================================================================
' Function Name : FncBtnPreview()
' Function Desc : This function is related to Preview Button
'========================================================================================
Function FncBtnPreview()
    Dim strFrDept
    Dim strToDept
    Dim strMin
    Dim strMax
    Dim condvar
    Dim strYear, strMonth, strDay
  	dim arrParam, arrField, arrHeader
    Dim StrEbrFile
    dim bas_dt, emp_no, pay_cd, ocpt_type, sect_cd, fr_dept_cd, to_dept_cd 
	
	strMin="0"
	strMax="ZZZZ"
    If Not chkField(Document, "1") Then									<%'��: This function check indispensable field%>
       Exit Function
    End If
    If txtEmp_no_Onchange() Then        'enter key �� ��ȸ�� ����� check�� �ش���� ������ query����...
        Exit Function
    End if
    
    If txtFr_dept_cd_Onchange() Then        'enter key �� ��ȸ�� �μ��ڵ带 check�� �ش���� ������ query����...
        Exit Function
    End if
    
    If txtTo_dept_cd_Onchange() Then        'enter key �� ��ȸ�� �μ��ڵ带 check�� �ش���� ������ query����...
        Exit Function
    End if
        
    strFrDept = Trim(frm1.txtFr_internal_cd.value)
    strToDept = Trim(frm1.txtTo_internal_cd.value)
    If strFrDept = "" AND strToDept = "" Then       
    Else
        If strFrDept = "" then
            strFrDept = strMin
        End if
        If strToDept = "" then
            strToDept = strMax
        ElseIf strFrDept > strToDept then
	        Call DisplayMsgbox("800153","X","X","X")	'���ۺμ��� ����μ����� �۾ƾ� �մϴ�.
            frm1.txtFr_dept_cd.focus
            Set gActiveElement = document.activeElement
            Exit Function
        End if 
    End if 
	
	Call parent.ExtractDateFrom(frm1.txtBas_dt.Text,frm1.txtBas_dt.UserDefinedFormat,parent.gComDateType,strYear,strMonth,strDay)
	
	bas_dt = strYear & strMonth & strDay
	emp_no = frm1.txtEmp_no.value
	pay_cd = frm1.cboPay_cd.value
	ocpt_type = frm1.cboOcpt_type.value
	sect_cd = frm1.cboSect_cd.value
	fr_dept_cd = frm1.txtFr_internal_cd.value
	to_dept_cd = frm1.txtTo_internal_cd.value

	if emp_no = "" then
		emp_no = "%"
		frm1.txtName.value = ""
	End if	

	if ocpt_type = "" then
		ocpt_type = "%"
	End if	

	if pay_cd = "" then
		pay_cd = "%"
	End if	
	
	if sect_cd = "" then
		sect_cd = "%"
	End if	
	
	if fr_dept_cd = "" then
		fr_dept_cd = strMin
		frm1.txtFr_dept_nm.value = ""
	End if	
	
	if to_dept_cd = "" then
		to_dept_cd = strMax
		frm1.txtTo_dept_nm.value = ""
	End if 


	ObjName = AskEBDocumentName("h9104oa1","ebr")

		condvar =  "bas_dt|" & bas_dt
		condvar = condvar & "|emp_no|" & emp_no
		condvar = condvar & "|pay_cd|" & pay_cd
		condvar = condvar & "|ocpt_type|" & ocpt_type
		condvar = condvar & "|sect_cd|" & sect_cd
		condvar = condvar & "|fr_dept_cd|" & fr_dept_cd
		condvar = condvar & "|to_dept_cd|" & to_dept_cd

	Call FncEBRPreview(ObjName,condvar)
	
    Dim ObjName2 , ObjName3
	ObjName2 = AskEBDocumentName("h9104oa1_2","ebr")
	ObjName3 = AskEBDocumentName("h9104oa1_3","ebr")

	condvar = ""

	Call FncEBRPreview(ObjName2,condvar)
	'Call FncEBRPreview(ObjName3,condvar)
End Function

Function FncPrint() 
    Call parent.FncPrint()
End Function

Function FncFind() 
	Call parent.FncFind(parent.C_SINGLE,False)
End Function

'========================================================================================================
'   Event Name : txtEmp_no_change             '<==�λ縶���Ϳ� �ִ� ������� Ȯ�� 
'   Event Desc :
'========================================================================================================
Function txtEmp_no_Onchange()
    Dim IntRetCd
    Dim strName
    Dim strDept_nm
    Dim strRoll_pstn
    Dim strPay_grd1
    Dim strPay_grd2
    Dim strEntr_dt
    Dim strInternal_cd
    Dim strVal
    
    If frm1.txtEmp_no.value = "" Then
		frm1.txtName.value = ""
    Else
	    IntRetCd = FuncGetEmpInf2(frm1.txtEmp_no.value,lgUsrIntCd,strName,strDept_nm,_
	                strRoll_pstn, strPay_grd1, strPay_grd2, strEntr_dt, strInternal_cd)
	    if  IntRetCd < 0 then
	        if  IntRetCd = -1 then
    			Call DisplayMsgbox("800048","X","X","X")	'�ش����� �������� �ʽ��ϴ�.
            else
                Call DisplayMsgbox("800454","X","X","X")	'�ڷῡ ���� ������ �����ϴ�.
            end if
			frm1.txtName.value = ""
            frm1.txtEmp_no.focus
            Set gActiveElement = document.ActiveElement
            txtEmp_no_Onchange = true
            Exit Function      
        Else
            frm1.txtName.value = strName
        End if 
    End if  
    
End Function 

'========================================================================================================
'   Event Name : txtFr_dept_cd_change
'   Event Desc :
'========================================================================================================
Function txtFr_dept_cd_Onchange()
    Dim IntRetCd
    Dim strDept_nm

    If frm1.txtFr_dept_cd.value = "" Then
		frm1.txtFr_dept_nm.value = ""
		frm1.txtFr_internal_cd.value = ""
    Else
        IntRetCd = FuncDeptName(frm1.txtFr_dept_cd.value,UNIConvDate(frm1.txtBas_dt.Text),lgUsrIntCd,strDept_nm,lsInternal_cd)
        
        if  IntRetCd < 0 then
            if  IntRetCd = -1 then
                Call DisplayMsgbox("800098", "x","x","x")   '�μ��ڵ������� ��ϵ��� ���� �ڵ��Դϴ�.
            else
                Call DisplayMsgbox("800455", "x","x","x")   ' �ڷ������ �����ϴ�.
            end if
		    frm1.txtFr_dept_nm.value = ""
		    frm1.txtFr_internal_cd.value = ""
            lsInternal_cd = ""
            frm1.txtFr_dept_cd.focus
            Set gActiveElement = document.ActiveElement
            txtFr_dept_cd_Onchange = true
            Exit Function      
        else
            frm1.txtFr_dept_nm.value = strDept_nm
            frm1.txtFr_internal_cd.value = lsInternal_cd
        end if
    End if  
        
End Function

'========================================================================================================
'   Event Name : txtTo_dept_cd_Onchange
'   Event Desc :
'========================================================================================================
Function txtTo_dept_cd_Onchange()
    Dim IntRetCd
    Dim strDept_nm
    
    If frm1.txtTo_dept_cd.value = "" Then
		frm1.txtTo_dept_nm.value = ""
		frm1.txtTo_internal_cd.value = ""
    Else
        IntRetCd = FuncDeptName(frm1.txtTo_dept_cd.value,UNIConvDate(frm1.txtBas_dt.Text),lgUsrIntCd,strDept_nm,lsInternal_cd)
        
        if  IntRetCd < 0 then
            if  IntRetCd = -1 then
                Call DisplayMsgbox("800098", "x","x","x")   '�μ��ڵ������� ��ϵ��� ���� �ڵ��Դϴ�.
            else
                Call DisplayMsgbox("800455", "x","x","x")   ' �ڷ������ �����ϴ�.
            end if
		    frm1.txtTo_dept_nm.value = ""
		    frm1.txtTo_internal_cd.value = ""
            lsInternal_cd = ""
            frm1.txtTo_dept_cd.focus
            Set gActiveElement = document.ActiveElement
            txtTo_dept_cd_Onchange = true
            Exit Function      
        else
            frm1.txtTo_dept_nm.value = strDept_nm
            frm1.txtTo_internal_cd.value = lsInternal_cd
        end if
    End if      
End Function

'========================================================================================================
' Name : txtBas_dt_DblClick
' Desc : �޷� Popup�� ȣ�� 
'========================================================================================================
Sub txtBas_dt_DblClick(Button)
	If Button = 1 Then
		Call SetFocusToDocument("M")	
		frm1.txtBas_dt.Action = 7
		frm1.txtBas_dt.focus
	End If
End Sub

'=======================================================================================================
'   Event Name : txtFr_year_Keypress(Key)
'   Event Desc : enter key down�ÿ� ��ȸ�Ѵ�.
'=======================================================================================================
Sub txtBas_dt_Keypress(Key)
    If Key = 13 Then
        FncQuery()
    End If
End Sub

</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	

<SCRIPT LANGUAGE="JavaScript">
<!-- Hide script from old browsers
function setCookie(name, value, expire)
{
	document.cookie = name + "=" + escape(value)
		+ ((expire == null) ? "" : ("; expires=" + expire.toGMTString()))
	document.cookie = name + "=" + escape(value)
		+ "; path=/EasyBaseWeb/bin"
		+ ((expire == null) ? "" : ("; expires=" + expire.toGMTString()))
	document.cookie = name + "=" + escape(value)
		+ "; path=/EasyBaseWeb/lib"
		+ ((expire == null) ? "" : ("; expires=" + expire.toGMTString()))
}

setCookie("client", "-1", null)
setCookie("owner", "admin", null)
setCookie("identity", "admin", null)
-->
</SCRIPT>

</HEAD>

<BODY TABINDEX="-1" SCROLL="NO">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE CLASS="BatchTB1" CELLSPACING=0 CELLPADDING=0>
	<TR>
		<TD <%=HEIGHT_TYPE_00%>></TD>
	</TR>
	<TR HEIGHT=23>
		<TD>
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSLTAB"><font color=white>�ٷμҵ�����Ű���</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* HEIGHT="right">&nbsp;</TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=*>
		<TD CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD HEIGHT=20>
						<TABLE <%=LR_SPACE_TYPE_60%>>   
							<TR>
								<TD CLASS="TD5" NOWRAP>��������</TD>
								<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=txtBas_dt NAME="txtBas_dt" CLASS=FPDTYYYYMMDD tag="12X1" ALT="��������" Title="FPDATETIME"></OBJECT>');</SCRIPT></td>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>�޿�����</TD>
	                        	<TD CLASS="TD6" NOWRAP><SELECT Name="cboPay_cd" ALT="�޿�����" STYLE="WIDTH: 120px" tag="11"><OPTION Value=""></OPTION></SELECT></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>�ٹ�����</TD>
	                        	<TD CLASS="TD6" NOWRAP><SELECT Name="cboSect_cd" ALT="�ٹ�����" STYLE="WIDTH: 120px" tag="11"><OPTION Value=""></OPTION></SELECT></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>�����</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" ID = "txtEmp_no" NAME="txtEmp_no"  SIZE=13 MAXLENGTH=13 tag="11XXXU" ALT="�����"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnAcctCdFR" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenEmptName (0)">
								                       <INPUT TYPE="Text" NAME="txtName" SIZE=20 MAXLENGTH=30  tag="14XXXU" ALT="������ڵ��"></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>�μ��ڵ�</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" ID = "txtFr_dept_cd" NAME="txtFr_dept_cd" SIZE=10 MAXLENGTH=10  tag="11XXXU" ALT="���ۺμ��ڵ�"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnAcctCdFR" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenDept(0)">
								                       <INPUT TYPE="Text" NAME="txtFr_dept_nm" SIZE=20 MAXLENGTH=40  tag="14XXXU" ALT="�μ��ڵ��">&nbsp;~&nbsp;
								                       <INPUT NAME="txtFr_internal_cd" ALT="���κμ��ڵ�" TYPE="hidden" SiZE=7 MAXLENGTH=30  tag="14XXXU"></TD>
							</TR>			
							<TR>
								<TD CLASS="TD5" NOWRAP></TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" ID = "txtTo_dept_cd" NAME="txtTo_dept_cd" SIZE=10 MAXLENGTH=10  tag="11XXXU" ALT="����μ��ڵ�"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnAcctCdFR" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenDept(1)">
								                       <INPUT TYPE="Text" NAME="txtTo_dept_nm" SIZE=20 MAXLENGTH=40  tag="14XXXU" ALT="�μ��ڵ��">&nbsp;
								                       <INPUT NAME="txtTo_internal_cd" ALT="���κμ��ڵ�" TYPE="hidden" SiZE=7 MAXLENGTH=30  tag="14XXXU"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>����</TD>
								<TD CLASS=TD6 NOWRAP><SELECT Name="cboOcpt_type" ALT="��������" STYLE="WIDTH: 120px" tag="11"><OPTION Value=""></OPTION></SELECT></TD>
								
							</TR>
    					</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	
	<TR>
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</TR>
	<TR HEIGHT=20>
		<TD>
		    <TABLE <%=LR_SPACE_TYPE_30%>>
		        <TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD>
                         <BUTTON NAME="btnRun"   CLASS="CLSSBTN" ONCLICK="vbscript:FncBtnPreview()" Flag=1>�̸�����</BUTTON>&nbsp;
                         <BUTTON NAME="btnPrint" CLASS="CLSSBTN" ONCLICK="vbscript:FncBtnPrint()" Flag=1>�μ�</BUTTON>
		            </TD>
					<TD WIDTH=* ALIGN="right">&nbsp;</TD>
					<TD WIDTH=10>&nbsp;</TD>
		        </TR>
		    </TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
<FORM NAME=EBAction TARGET="MyBizASP" METHOD="POST">
	<INPUT TYPE="HIDDEN" NAME="uname">
	<INPUT TYPE="HIDDEN" NAME="dbname">
	<INPUT TYPE="HIDDEN" NAME="filename">
	<INPUT TYPE="HIDDEN" NAME="condvar">
	<INPUT TYPE="HIDDEN" NAME="date">	
</FORM>
</BODY>
</HTML>

