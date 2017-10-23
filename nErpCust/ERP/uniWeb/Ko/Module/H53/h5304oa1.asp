<%@ LANGUAGE="VBSCRIPT" %>
<!--
'======================================================================================================
'*  1. Module Name          : �λ�/�޿����� 
'*  2. Function Name        : ��/�󿩰����������� 
'*  3. Program ID           : h5304oa1
'*  4. Program Name         : �ǰ������ڰ����/�����Ű� 
'*  5. Program Desc         : �ǰ������ڰ����/�����Ű� 
'*  6. Comproxy List        : +
'*  7. Modified date(First) : 2001/05/30
'*  8. Modified date(Last)  : 2003/06/11
'*  9. Modifier (First)     : BongKyu, Song
'* 10. Modifier (Last)      : Lee SiNa
'* 11. Comment              :
'=======================================================================================================-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/IncServer.asp" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incHRQuery.vbs"></SCRIPT>
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
'========================================================================================================
' Name : InitVariables()
' Desc : Initialize value
'========================================================================================================
Sub InitVariables()

    lgIntFlgMode = Parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
End Sub

'========================================================================================================
' Name : SetDefaultVal()
' Desc : Set default value
'========================================================================================================
Sub SetDefaultVal()
	frm1.txtFr_acq_dt.focus
	frm1.txtFr_acq_dt.text	= UniConvDateAToB("<%=GetSvrDate%>",Parent.gServerDateFormat,Parent.gDateFormat)
	frm1.txtTo_acq_dt.text	= frm1.txtFr_acq_dt.text
	frm1.txtRprt_dt.text	= frm1.txtFr_acq_dt.text
	
End Sub

'========================================================================================================
' Name : LoadInfTB19029()
' Desc : Set System Number format
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "H", "NOCOOKIE", "OA") %>
End Sub

'========================================================================================================
' Name : Form_Load
' Desc : developer describe this line Called by Window_OnLoad() evnt
'========================================================================================================
Sub Form_Load()

    Call LoadInfTB19029                                                     '��: Load table , B_numeric_format
    Call ggoOper.LockField(Document, "N")                                   '��: Lock  Suitable  Field
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
    Call InitVariables

    Call SetDefaultVal
	Call ggoOper.FormatDate(frm1.txtFr_acq_dt, Parent.gDateFormat, 1)
	Call ggoOper.FormatDate(frm1.txtTo_acq_dt, Parent.gDateFormat, 1)
	Call ggoOper.FormatDate(frm1.txtRprt_dt, Parent.gDateFormat, 1)
    Call SetToolbar("1000000000000111")
End Sub
'========================================================================================================
' Name : Form_QueryUnload
' Desc : developer describe this line Called by Window_OnUnLoad() evnt
'========================================================================================================
Sub Form_QueryUnload(Cancel, UnloadMode)
End Sub

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
	Call Parent.FncExport(Parent.C_SINGLE)
End Function

'========================================================================================================
' Name : FncFind
' Desc : developer describe this line Called by MainFind in Common.vbs
'========================================================================================================
Function FncFind()
	Call Parent.FncFind(Parent.C_SINGLE, False)
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

'======================================================================================================
'	Name : OpenCode()
'	Description : Code PopUp at vspdData
'=======================================================================================================
Function OpenCode(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True  Then
	   Exit Function
	End If

	IsOpenPop = True

	Select Case iWhere
	    Case "SECT_CD_POP"
	        arrParam(0) = "�ٹ����� �˾�"			        ' �˾� ��Ī 
	    	arrParam(1) = "B_MINOR"							    ' TABLE ��Ī 
	    	arrParam(2) = frm1.txtSect_cd.value     			' Code Condition
	    	arrParam(3) = ""'frm1.txtSect_nm.value    				' Name Cindition
	    	arrParam(4) = "MAJOR_CD = " & FilterVar("H0035", "''", "S") & ""	               	' Where Condition
	    	arrParam(5) = "�ٹ������ڵ�"  		            ' TextBox ��Ī 

	    	arrField(0) = "MINOR_CD"						   	' Field��(0)
	    	arrField(1) = "MINOR_NM"    				  		' Field��(1)
	    	arrField(2) = ""    				        		' Field��(2)

	    	arrHeader(0) = "�ٹ������ڵ�"	     			' Header��(0)
	    	arrHeader(1) = "�ٹ������ڵ��"	   		        ' Header��(1)
	    	arrHeader(2) = ""   	    						' Header��(1)
	    Case "CUST_CD_POP"
	        arrParam(0) = "�Ű����� �˾�"			        ' �˾� ��Ī 
	    	arrParam(1) = "B_BIZ_AREA"							    ' TABLE ��Ī 
	    	arrParam(2) = frm1.txtCust_cd.value     			' Code Condition
	    	arrParam(3) = ""'frm1.txtCust_nm.value									' Name Cindition
	    	arrParam(4) = ""	               	                ' Where Condition
	    	arrParam(5) = "�Ű������ڵ�"  		        ' TextBox ��Ī 

	    	arrField(0) = "BIZ_AREA_CD"						   	' Field��(0)
	    	arrField(1) = "BIZ_AREA_NM"    				  		' Field��(1)
	    	arrField(2) = ""    				        		' Field��(2)

	    	arrHeader(0) = "�Ű������ڵ�"	     			' Header��(0)
	    	arrHeader(1) = "�Ű������ڵ��"		        ' Header��(1)
	    	arrHeader(2) = ""   	    						' Header��(1)
	End Select

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then

		Select Case iWhere
		    Case "SECT_CD_POP"
		    	frm1.txtSect_cd.focus
		    Case "CUST_CD_POP"
		    	frm1.txtCust_cd.focus
        End Select	
		Exit Function
	Else
		Call SetCode(arrRet,iWhere)
	End If

End Function

'======================================================================================================
'	Name : SetCode()
'	Description : Code PopUp���� Return�Ǵ� �� setting
'=======================================================================================================
Function SetCode(Byval arrRet, Byval iWhere)

	With Frm1

		Select Case iWhere
		    Case "SECT_CD_POP"
		        .txtSect_cd.value = arrRet(0)
		    	.txtSect_nm.value = arrRet(1)
		    	.txtSect_cd.focus
		    Case "CUST_CD_POP"
		        .txtCust_cd.value = arrRet(0)
		    	.txtCust_nm.value = arrRet(1)
		    	.txtCust_cd.focus
        End Select

	End With

End Function

'========================================================================================================
' Name : OpenEmpName()
' Desc : developer describe this line(grid�ܿ��� ���)
'========================================================================================================
Function OpenEmpName(iWhere)
	Dim arrRet
	Dim arrParam(2)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

    If  iWhere = 0 Then
	    arrParam(0) = frm1.txtEmp_no.value			' Code Condition
	    arrParam(1) = ""'frm1.txtName.value			' Name Cindition
	    arrParam(2) = ""                    ' �ڷ���� Condition
    Else
	    arrParam(0) = frm1.txtEmp_no.value			' Code Condition
	    arrParam(1) = ""'frm1.txtName.value			' Name Cindition
	    arrParam(2) = ""                   ' �ڷ���� Condition
	End If

	arrRet = window.showModalDialog(HRAskPRAspName("EmpPopup"), Array(window.parent, arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtEmp_no.focus
		Exit Function
	Else
		Call SetEmpName(arrRet)
	End If

End Function

'======================================================================================================
'	Name : SetEmpName()
'	Description : Item Popup���� Return�Ǵ� �� setting(grid�ܿ��� ���)
'=======================================================================================================
Sub SetEmpName(arrRet)
	With frm1
		.txtEmp_no.value = arrRet(0)
		.txtName.value = arrRet(1)
		Call ggoOper.ClearField(Document, "2")					 '��: Clear Contents  Field
		Set gActiveElement = document.ActiveElement
		.txtEmp_no.focus
		lgBlnFlgChgValue = False
	End With
End Sub

'========================================================================================================
' Name : FncBtnPrint
' Desc : developer describe this line
'========================================================================================================
Function FncBtnPrint()
	Dim condvar
	Dim lngPos
	Dim intCnt
    Dim StrEbrFile
    Dim ObjName

	dim emp_no, fr_dt, to_dt, singo_dt, sect_cd, cust_cd
    
    If Not chkField(Document, "1") Then									<%'��: This function check indispensable field%>
       Exit Function
    End If

	If frm1.txtFamYn(0).checked Then
		StrEbrFile = "h5304oa1_1"
	Else
		StrEbrFile = "h5304oa1_2"
	End If

    If (frm1.txtFr_acq_dt.text <> "") AND (frm1.txtTo_acq_dt.text <> "") Then
    	IF CompareDateByFormat(frm1.txtFr_acq_dt.Text,frm1.txtTo_acq_dt.Text,frm1.txtFr_acq_dt.Alt,frm1.txtTo_acq_dt.Alt,"970025",frm1.txtFr_acq_dt.UserDefinedFormat,Parent.gComDateType,False)=False Then
            Call DisplayMsgbox("970025","X","��������","��������")	
            frm1.txtFr_acq_dt.focus
            Set gActiveElement = document.activeElement
            Exit Function
        End if 
    End if 

	emp_no = frm1.txtEmp_no.value
	
	fr_dt = UniConvDateToYYYYMMDD(frm1.txtFr_acq_dt.text,Parent.gDateFormat,Parent.gServerDateType)	
	to_dt = UniConvDateToYYYYMMDD(frm1.txtTo_acq_dt.text,Parent.gDateFormat,Parent.gServerDateType)
	sect_cd = frm1.txtSect_cd.value
	cust_cd = frm1.txtCust_cd.value	
	singo_dt = UniConvDateToYYYYMMDD(frm1.txtRprt_dt.text,Parent.gDateFormat,"")

	if emp_no = "" then
		emp_no = "%"
		frm1.txtName.value = ""
	End if
    If txtSect_cd_Onchange() Then                                                'enter key �� ��ȸ�� ���ۺμ��ڵ带 check�� �ش���� ������ query����...
        Exit Function
    End if
    If txtCust_cd_Onchange() Then                                                'enter key �� ��ȸ�� ���ۺμ��ڵ带 check�� �ش���� ������ query����...
        Exit Function
    End if 
    If txtEmp_no_Onchange() Then                                                'enter key �� ��ȸ�� ���ۺμ��ڵ带 check�� �ش���� ������ query����...
        Exit Function
    End if

	if sect_cd = "" then
		sect_cd = "%"
		frm1.txtSect_nm.value = ""
	End if

	condvar = "Emp_no|" & emp_no
	condvar = condvar & "|Fromdt|" & fr_dt
	condvar = condvar & "|Todt|" & to_dt
	condvar = condvar & "|Regdt|" & singo_dt
	condvar = condvar & "|Sect_cd|" & sect_cd
	condvar = condvar & "|Cust_cd|" & cust_cd

    ObjName = AskEBDocumentName(StrEbrFile, "ebr")

 	call FncEBRPrint(EBAction , ObjName , condvar)

End Function


'========================================================================================================
' Name : BtnPreview
' Desc : This function is related to Preview Button
'========================================================================================================
Function BtnPreview()

    Dim strYear, strMonth, strDay

	dim condvar
	dim arrParam, arrField, arrHeader
    Dim StrEbrFile
    Dim ObjName

	dim emp_no, fr_dt, to_dt, singo_dt, sect_cd, cust_cd

    If Not chkField(Document, "1") Then									<%'��: This function check indispensable field%>
       Exit Function
    End If

	If frm1.txtFamYn(0).checked Then
		StrEbrFile = "h5304oa1_1"
	Else
		StrEbrFile = "h5304oa1_2"
	End If

    If (frm1.txtFr_acq_dt.text <> "") AND (frm1.txtTo_acq_dt.text <> "") Then
    	IF CompareDateByFormat(frm1.txtFr_acq_dt.Text,frm1.txtTo_acq_dt.Text,frm1.txtFr_acq_dt.Alt,frm1.txtTo_acq_dt.Alt,"970025",frm1.txtFr_acq_dt.UserDefinedFormat,Parent.gComDateType,False)=False Then
            Call DisplayMsgbox("970025","X","��������","��������")	
            frm1.txtFr_acq_dt.focus
            Set gActiveElement = document.activeElement
            Exit Function
        End if 
    End if 

	emp_no = frm1.txtEmp_no.value

	fr_dt = UniConvDateToYYYYMMDD(frm1.txtFr_acq_dt.text,Parent.gDateFormat,Parent.gServerDateType)	
	to_dt = UniConvDateToYYYYMMDD(frm1.txtTo_acq_dt.text,Parent.gDateFormat,Parent.gServerDateType)	
	sect_cd = Trim(UCase(frm1.txtSect_cd.value))
	cust_cd = Trim(UCase(frm1.txtCust_cd.value))
	singo_dt = UniConvDateToYYYYMMDD(frm1.txtRprt_dt.text,Parent.gDateFormat,"")
	
	if emp_no = "" then
		emp_no = "%"
		frm1.txtName.value = ""
	End if
    If txtSect_cd_Onchange() Then                                                'enter key �� ��ȸ�� ���ۺμ��ڵ带 check�� �ش���� ������ query����...
        Exit Function
    End if
    If txtCust_cd_Onchange() Then                                                'enter key �� ��ȸ�� ���ۺμ��ڵ带 check�� �ش���� ������ query����...
        Exit Function
    End if    
    If txtEmp_no_Onchange() Then                                                'enter key �� ��ȸ�� ���ۺμ��ڵ带 check�� �ش���� ������ query����...
        Exit Function
    End if

	if sect_cd = "" then
		sect_cd = "%"
		frm1.txtSect_nm.value = ""
	End if

	condvar = "Emp_no|" & emp_no
	condvar = condvar & "|Fromdt|" & fr_dt
	condvar = condvar & "|Todt|" & to_dt
	condvar = condvar & "|Regdt|" & singo_dt
	condvar = condvar & "|Sect_cd|" & sect_cd
	condvar = condvar & "|Cust_cd|" & cust_cd

    ObjName = AskEBDocumentName(StrEbrFile, "ebr")

	call FncEBRPreview(ObjName , condvar)

End Function

'========================================================================================================
'   Event Name : txtEmp_no_change             '<==�λ縶���Ϳ� �ִ� ������� Ȯ�� 
'   Event Desc :
'========================================================================================================
Function txtEmp_no_Onchange()
    Dim IntRetCd

    If frm1.txtEmp_no.value = "" Then
		frm1.txtName.value = ""
    Else

        IntRetCd = CommonQueryRs(" NAME "," HAA010T "," EMP_NO =  " & FilterVar(frm1.txtEmp_no.value , "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
        If  IntRetCd = false then
            Call DisplayMsgbox("800048","X","X","X")	'�ش����� �������� �ʽ��ϴ�.
			frm1.txtName.value = ""
            frm1.txtEmp_no.focus
            Set gActiveElement = document.ActiveElement
            txtEmp_no_Onchange = true
            Exit Function
        Else
            frm1.txtName.value = Trim(Replace(lgF0,Chr(11),""))
        End if
    End if

End Function


'========================================================================================================
'   Event Name : txtSect_cd_Onchange
'   Event Desc :
'========================================================================================================
Function txtSect_cd_Onchange()
    Dim IntRetCd
    If frm1.txtSect_cd.value = "" Then
		frm1.txtSect_nm.value = ""
    Else
        IntRetCd = CommonQueryRs(" MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0035", "''", "S") & " AND MINOR_CD =  " & FilterVar(frm1.txtSect_cd.value , "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
        If IntRetCd = false then
			Call DisplayMsgbox("800054","X","X","X")	'��ϵ��� ���� �ڵ��Դϴ�.
			 frm1.txtSect_nm.value = ""
             frm1.txtSect_cd.focus
            Set gActiveElement = document.ActiveElement
            txtSect_cd_Onchange = true
        Else
			frm1.txtSect_nm.value = Trim(Replace(lgF0,Chr(11),""))
        End if
    End if

End Function

'========================================================================================================
'   Event Name : txtCust_cd_Onchange
'   Event Desc :
'========================================================================================================
Function txtCust_cd_Onchange()
    Dim IntRetCd
    If frm1.txtCust_cd.value = "" Then
		frm1.txtCust_nm.value = ""
    Else
        IntRetCd = CommonQueryRs(" BIZ_AREA_NM "," B_BIZ_AREA "," BIZ_AREA_CD =  " & FilterVar(frm1.txtCust_cd.value , "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
        If IntRetCd = false then
			Call DisplayMsgbox("800054","X","X","X")	'��ϵ��� ���� �ڵ��Դϴ�.
			 frm1.txtCust_nm.value = ""
             frm1.txtCust_cd.focus
            Set gActiveElement = document.ActiveElement
            txtCust_cd_Onchange = true            
        Else
			frm1.txtCust_nm.value = Trim(Replace(lgF0,Chr(11),""))
        End if
    End if

End Function

'======================================================================================================
'   Event Name : txtYyyymm_DblClick
'   Event Desc : �޷� Popup�� ȣ�� 
'=======================================================================================================

Sub txtFr_acq_dt_DblClick(Button)
	If Button = 1 Then
		Call SetFocusToDocument("M")
		frm1.txtFr_acq_dt.Action = 7
		frm1.txtFr_acq_dt.focus
	End If
End Sub

Sub txtTo_acq_dt_DblClick(Button)
	If Button = 1 Then
		Call SetFocusToDocument("M")	
		frm1.txtTo_acq_dt.Action = 7
		frm1.txtTo_acq_dt.focus
	End If
End Sub

Sub txtRprt_dt_DblClick(Button)
	If Button = 1 Then
		Call SetFocusToDocument("M")	
		frm1.txtRprt_dt.Action = 7
		frm1.txtRprt_dt.focus
	End If
End Sub

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->

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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSLTAB"><font color=white>���尡������溯���Ű�</font></td>
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
								<TD CLASS=TD5 NOWRAP>�Ǻξ���</TD>
				        	    <TD CLASS=TD6><INPUT TYPE="RADIO" CLASS="RADIO" NAME="txtFamYn" TAG="2X" VALUE="�ֽ�" CHECKED ID="txtFamYn1"><LABEL FOR="txtFamYn1">�ֽ�</LABEL>&nbsp;
  				        	                  <INPUT TYPE="RADIO" CLASS="RADIO" NAME="txtFamYn" TAG="2X" VALUE="����" ID="txtFamYn2"><LABEL FOR="txtFamYn2">����</LABEL></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>������</TD>
								<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/h5304oa1_txtFr_acq_dt_txtFr_acq_dt.js'></script>&nbsp;~</TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP></TD>
								<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/h5304oa1_txtTo_acq_dt_txtTo_acq_dt.js'></script>&nbsp;</TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>�Ű���</TD>
								<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/h5304oa1_txtRprt_dt_txtRprt_dt.js'></script></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>�ٹ�����</TD>
								<TD CLASS=TD6 NOWRAP>
								    <INPUT ID="txtSect_cd" NAME="txtSect_cd" ALT="�ٹ�����" TYPE="Text" SiZE=10 MAXLENGTH=10 tag="1XXXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSItemDC" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenCode frm1.txtSect_cd.value,'SECT_CD_POP'">
								    <INPUT NAME="txtSect_nm" ALT="�ٹ�����" TYPE="Text" SiZE=20 MAXLENGTH=50 tag="14XXXU"></td>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>�Ű�����</TD>
								<TD CLASS=TD6 NOWRAP>
								    <INPUT ID="txtCust_cd" NAME="txtCust_cd" ALT="�Ű�����" TYPE="Text" SiZE=10 MAXLENGTH=10 tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSItemDC" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenCode frm1.txtCust_cd.value,'CUST_CD_POP'">
								    <INPUT NAME="txtCust_nm" ALT="�Ű�����" TYPE="Text" SiZE=20 MAXLENGTH=100 tag="14XXXU"></td>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>�����</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtEmp_no" ALT="�����" TYPE="Text" MAXLENGTH="13" SIZE=13  tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBP_CD" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenEmpName (1)">
								                     <INPUT NAME="txtName" TYPE="Text" MAXLENGTH="30" SIZE=20  tag="14XXXU"></TD>
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
                         <BUTTON NAME="btnRun"   CLASS="CLSSBTN" ONCLICK="vbscript:BtnPreview()" Flag=1>�̸�����</BUTTON>&nbsp;
                         <BUTTON NAME="btnPrint" CLASS="CLSSBTN" ONCLICK="vbscript:FncBtnPrint()" Flag=1>�μ�</BUTTON>
		            </TD>
					<TD WIDTH=* ALIGN="right">&nbsp;</TD>
					<TD WIDTH=10>&nbsp;</TD>
		        </TR>
		    </TABLE>
		</TD>
	</TR>
	<TR>
		<TD HEIGHT=20><IFRAME NAME="MyBizASP" SRC = "../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
<FORM NAME="EBAction" TARGET = "MyBizASP" METHOD="POST">
	<INPUT TYPE="HIDDEN" NAME="uname">
	<INPUT TYPE="HIDDEN" NAME="dbname">
	<INPUT TYPE="HIDDEN" NAME="filename">
	<INPUT TYPE="HIDDEN" NAME="condvar">
	<INPUT TYPE="HIDDEN" NAME="date">
</FORM>
</BODY>
</HTML>


