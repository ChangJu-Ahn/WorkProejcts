<%@ LANGUAGE="VBSCRIPT" %>
<!--
'======================================================================================================
'*  1. Module Name          : 인사/급여관리 
'*  2. Function Name        : 급/상여공제관리관리 
'*  3. Program ID           : h6015oa1
'*  4. Program Name         : 현금지급명세서출력 
'*  5. Program Desc         : 현금지급명세서출력 
'*  6. Comproxy List        : +
'*  7. Modified date(First) : 2001/05/31
'*  8. Modified date(Last)  : 2003/06/13
'*  9. Modifier (First)     : BongKyu, Song
'* 10. Modifier (Last)      : Lee SiNa
'* 11. Comment              :
'=======================================================================================================-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

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
Dim lsInternal_cd

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
    Dim strYear
	Dim strMonth
	Dim strDay
	
	Call ExtractDateFrom("<%=GetsvrDate%>",Parent.gServerDateFormat , Parent.gServerDateType ,strYear,strMonth,strDay)
	
	frm1.txtYyyymm.Focus			'년월 default value setting
	frm1.txtYyyymm.Year = strYear 
	frm1.txtYyyymm.Month = strMonth

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

    Call LoadInfTB19029                                                     '⊙: Load table , B_numeric_format
    Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
    Call InitVariables
	
	Call ggoOper.FormatDate(frm1.txtYyyymm, Parent.gDateFormat, 2)

    Call FuncGetAuth(gStrRequestMenuID, Parent.gUsrID, lgUsrIntCd)                                ' 자료권한:lgUsrIntCd ("%", "1%")
                                            
    Call SetDefaultVal

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
    
    FncDelete = False                                                            '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
    
    FncDelete = True                                                             '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncCancel
' Desc : developer describe this line Called by MainCancel in Common.vbs
'========================================================================================================
Function FncCancel() 
	On Error Resume Next                                                        '☜: Protect system from crashing
End Function

'========================================================================================================
' Name : FncPrint
' Desc : developer describe this line Called by MainDeleteRow in Common.vbs
'========================================================================================================
Function FncPrint()
	Call Parent.FncPrint()                                                      '☜: Protect system from crashing
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
	    Case "PROV_CD_POP"
			arrParam(0) = "지급구분팝업"			' 팝업 명칭 
			arrParam(1) = "B_MINOR"				 		' TABLE 명칭 
			arrParam(2) = strCode		                ' Code Condition
			arrParam(3) = ""'frm1.txtProvNm.value				' Name Cindition
			arrParam(4) = "MAJOR_CD = " & FilterVar("H0040", "''", "S") & ""			' Where Condition
			arrParam(5) = "지급구분"			    ' TextBox 명칭 
	
			arrField(0) = "MINOR_CD"					' Field명(0)
			arrField(1) = "MINOR_NM"				    ' Field명(1)
    
			arrHeader(0) = "지급구분"				' Header명(0)
			arrHeader(1) = "지급명"			        ' Header명(1)
	End Select   

    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtProvCd.focus
		Exit Function
	Else
		Call SetCode(arrRet,iWhere)
	End If	

End Function

'======================================================================================================
'	Name : SetCode()
'	Description : Code PopUp에서 Return되는 값 setting
'=======================================================================================================
Function SetCode(Byval arrRet, Byval iWhere)

	With Frm1

		Select Case iWhere
	    Case "PROV_CD_POP"
			.txtProvCd.value = arrRet(0)
			.txtProvNm.value = arrRet(1)		
			.txtProvCd.focus
		End Select

	End With

End Function

'========================================================================================================
' Name : OpenEmpName()
' Desc : developer describe this line(grid외에서 사용) 
'========================================================================================================
Function OpenEmpName(iWhere)
	Dim arrRet
	Dim arrParam(2)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

    If  iWhere = 0 Then
	    arrParam(0) = frm1.txtEmp_no.value			' Code Condition
	    arrParam(1) = ""			' Name Cindition
	    arrParam(2) = lgUsrIntCd                    ' 자료권한 Condition  
    Else
	    arrParam(0) = frm1.txtEmp_no.value			' Code Condition
	    arrParam(1) = ""			' Name Cindition
	    arrParam(2) = lgUsrIntCd                    ' 자료권한 Condition  
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
'	Description : Item Popup에서 Return되는 값 setting(grid외에서 사용)
'=======================================================================================================
Sub SetEmpName(arrRet)
	With frm1
		.txtEmp_no.value = arrRet(0)
		.txtName.value = arrRet(1)
		Call ggoOper.ClearField(Document, "2")					 '☜: Clear Contents  Field
		Set gActiveElement = document.ActiveElement
		.txtEmp_no.focus
		lgBlnFlgChgValue = False
	End With
End Sub


'========================================================================================================
' Name : OpenDept
' Desc : 부서 POPUP
'========================================================================================================
Function OpenDept(iWhere)
	Dim arrRet
	Dim arrParam(2)
    Dim strBasDt 
    
	strBasDt = UNIGetLastDay(frm1.txtYyyymm.Text,Parent.gDateFormatYYYYMM)
	
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	If iWhere = 0 Then
		arrParam(0) = frm1.txtFr_dept_cd.value			            '  Code Condition
	ElseIf iWhere = 1 Then
		arrParam(0) = frm1.txtTo_dept_cd.value			            ' Code Condition
	End If
	
    arrParam(1) = strBasDt
	arrParam(2) = lgUsrIntCd                              ' 자료권한 Condition  
	
	arrRet = window.showModalDialog(HRAskPRAspName("DeptPopupDt"), Array(window.parent, arrParam), _
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
'	Description : Dept Popup에서 Return되는 값 setting
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
		
'========================================================================================================
' Name : FncBtnPrint
' Desc : developer describe this line 
'========================================================================================================
Function FncBtnPrint() 

	Dim strUrl
	Dim lngPos
	Dim intCnt
    Dim StrEbrFile
    Dim strMin
    Dim strMax
    Dim rDate
    Dim ObjName

    rDate = UNIGetLastDay(frm1.txtYyyymm.Text, Parent.gDateFormatYYYYMM)                     '해당년월의 마지막 날을 가지고 온다.
    Call FuncGetTermDept(lgUsrIntCd,UNIConvDate(rDate),strMin,strMax)

    If Not chkField(Document, "1") Then									<%'⊙: This function check indispensable field%>
		Call BtnDisabled(0)
       Exit Function
    End If

	dim yyyymm, prov_cd, fr_dept_cd, to_dept_cd, prov_type, standard_price
	
	StrEbrFile = "h6015oa1"

    yyyymm = frm1.txtYyyymm.year & Right("0" & frm1.txtYyyymm.month , 2)
	prov_cd = frm1.txtProvCd.value
	fr_dept_cd = frm1.txtFr_internal_cd.value
	to_dept_cd = frm1.txtTo_internal_cd.value
	
	standard_price = UNICDbl(frm1.txtStandPrice.Text)
	standard_price = Replace(standard_price, Parent.gClientNumDec, ".")

	If frm1.txtProvType(0).checked Then
		prov_type = "1"
	ElseIf frm1.txtProvType(1).checked Then
		prov_type = "2"
	ElseIf frm1.txtProvType(2).checked Then	
		prov_type = "3"
	End if	

	if standard_price = "" then
		standard_price = "0"
	End if	
	
	if txtProvCd_Onchange() then
		Exit Function
	end if
	if txtFr_dept_cd_Onchange() then
		Exit Function
	end if
	if txtTo_dept_cd_Onchange() then
		Exit Function
	end if			
	if fr_dept_cd = "" then
		fr_dept_cd = strMin
		frm1.txtFr_dept_nm.value = ""
	End if	

	if to_dept_cd = "" then
		to_dept_cd = strMax
		frm1.txtTo_dept_nm.value = ""
 	End if	

    If (fr_dept_cd = "") AND (to_dept_cd = "") Then     
    Else
        If fr_dept_cd > to_dept_cd then
	        Call DisplayMsgbox("800153","X","X","X")	    '시작부서코드는 종료부서코드보다 작아야합니다.
            frm1.txtFr_dept_cd.focus
            Set gActiveElement = document.activeElement
            Exit Function
        End IF        
    END IF   

	Call BtnDisabled(1)
	
	strUrl = "Pay_Yymm|" & yyyymm
	strUrl = strUrl & "|Prov_Type|" & prov_cd
	strUrl = strUrl & "|Fr_Dept_cd|" & fr_dept_cd
	strUrl = strUrl & "|To_Dept_cd|" & to_dept_cd
	strUrl = strUrl & "|Gigup_Type1|" & prov_type
	strUrl = strUrl & "|Stand_amt|" & standard_price

    ObjName = AskEBDocumentName(StrEbrFile, "ebr")

 	call FncEBRPrint(EBAction , ObjName , strUrl)

End Function

'========================================================================================================
' Name : BtnPreview
' Desc : This function is related to Preview Button
'========================================================================================================
Function BtnPreview() 

	Dim strMin
    Dim strMax
    Dim rDate

	dim strUrl
	dim arrParam, arrField, arrHeader
    Dim StrEbrFile, ObjName
		
	dim yyyymm, prov_cd, fr_dept_cd, to_dept_cd, prov_type, standard_price

    
    rDate = UNIGetLastDay(frm1.txtYyyymm.Text, Parent.gDateFormatYYYYMM)                     '해당년월의 마지막 날을 가지고 온다.
    Call FuncGetTermDept(lgUsrIntCd,UNIConvDateCompanyToDB(rDate,Parent.gDateFormat),strMin,strMax)

    If Not chkField(Document, "1") Then									<%'⊙: This function check indispensable field%>
       Call BtnDisabled(0)
	   Exit Function
    End If
	
	StrEbrFile = "h6015oa1"

    yyyymm = frm1.txtYyyymm.year & Right("0" & frm1.txtYyyymm.month , 2)
	prov_cd = frm1.txtProvCd.value
	fr_dept_cd = frm1.txtFr_internal_cd.value
	to_dept_cd = frm1.txtTo_internal_cd.value

	standard_price = UNICDbl(frm1.txtStandPrice.Text)
	standard_price = Replace(standard_price, Parent.gClientNumDec, ".")

	If frm1.txtProvType(0).checked Then
		prov_type = "1"
	ElseIf frm1.txtProvType(1).checked Then
		prov_type = "2"
	ElseIf frm1.txtProvType(2).checked Then	
		prov_type = "3"
	End if	

	if standard_price = "" then
		standard_price = "0"
	End if	

	if txtProvCd_Onchange() then
		Exit Function
	end if
	if txtFr_dept_cd_Onchange() then
		Exit Function
	end if
	if txtTo_dept_cd_Onchange() then
		Exit Function
	end if			
	if fr_dept_cd = "" then
		fr_dept_cd = strMin
		frm1.txtFr_dept_nm.value = ""
	End if	

	if to_dept_cd = "" then
		to_dept_cd = strMax
		frm1.txtTo_dept_nm.value = ""
	End if	

    If (fr_dept_cd = "") AND (to_dept_cd = "") Then     
    Else
        If fr_dept_cd > to_dept_cd then
	        Call DisplayMsgbox("800153","X","X","X")	'시작부서코드는 종료부서코드보다 작아야합니다.
            frm1.txtFr_dept_cd.focus
            Set gActiveElement = document.activeElement
            Exit Function
        End IF        
    END IF   
	Call BtnDisabled(1)
	
	strUrl = "Pay_Yymm|" & yyyymm
	strUrl = strUrl & "|Prov_Type|" & prov_cd
	strUrl = strUrl & "|Fr_Dept_cd|" & fr_dept_cd
	strUrl = strUrl & "|To_Dept_cd|" & to_dept_cd
	strUrl = strUrl & "|Gigup_Type1|" & prov_type
	strUrl = strUrl & "|Stand_amt|" & standard_price

    ObjName = AskEBDocumentName(StrEbrFile, "ebr")
	call FncEBRPreview(ObjName , strUrl)

End Function

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
    			Call DisplayMsgbox("800048","X","X","X")	'해당사원은 존재하지 않습니다.
            else
                Call DisplayMsgbox("800454","X","X","X")	'자료에 대한 권한이 없습니다.
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
'   Event Name : txtProvCd_Onchange             
'   Event Desc :
'========================================================================================================
Function txtProvCd_Onchange()
    Dim IntRetCd
    
    If frm1.txtProvCd.value = "" Then
		frm1.txtProvNm.value = ""
    Else
        IntRetCd = CommonQueryRs(" MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0040", "''", "S") & " AND MINOR_CD =  " & FilterVar(frm1.txtProvCd.value , "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
        If IntRetCd = false then
			Call DisplayMsgbox("800054","X","X","X")	'등록되지 않은 코드입니다.
			 frm1.txtProvNm.value = ""
             frm1.txtProvCd.focus
            Set gActiveElement = document.ActiveElement   
            txtProvCd_Onchange = true    
        Else
			frm1.txtProvNm.value = Trim(Replace(lgF0,Chr(11),""))
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
	Dim rDate
	
	rDate = UNIGetLastDay(frm1.txtYyyymm.Text, Parent.gDateFormatYYYYMM)            '해당년월의 마지막 날을 가지고 온다.
	If frm1.txtFr_dept_cd.value = "" Then
		frm1.txtFr_dept_nm.value = ""
		frm1.txtFr_internal_cd.value = ""
    Else
        IntRetCd = FuncDeptName(frm1.txtFr_dept_cd.value,UNIConvDate(rDate),lgUsrIntCd,strDept_nm,lsInternal_cd)
        
        If  IntRetCd < 0 then
            if  IntRetCd = -1 then
                Call DisplayMsgBox("800012", "x","x","x")   ' 부서코드정보에 등록되지 않은 코드입니다.
            else
                Call DisplayMsgBox("800455", "x","x","x")   ' 자료권한이 없습니다.
            end if
		    frm1.txtFr_dept_nm.value = ""
		    frm1.txtFr_internal_cd.value = ""
            lsInternal_cd = ""
            frm1.txtFr_dept_cd.focus
            Set gActiveElement = document.ActiveElement
            txtFr_dept_cd_Onchange = true        
            Exit Function      
        Else
            frm1.txtFr_dept_nm.value = strDept_nm
            frm1.txtFr_internal_cd.value = lsInternal_cd
        End if
    End if
End Function

'========================================================================================================
'   Event Name : txtTo_dept_cd_Onchange
'   Event Desc :
'========================================================================================================
Function txtTo_dept_cd_Onchange()
    Dim IntRetCd
    Dim strDept_nm
	Dim rDate
	
	rDate = UNIGetLastDay(frm1.txtYyyymm.Text, Parent.gDateFormatYYYYMM)            '해당년월의 마지막 날을 가지고 온다.

	If frm1.txtTo_dept_cd.value = "" Then
		frm1.txtTo_dept_nm.value = ""
		frm1.txtTo_internal_cd.value = ""
    Else
        IntRetCd = FuncDeptName(frm1.txtTo_dept_cd.value,UNIConvDate(rDate),lgUsrIntCd,strDept_nm,lsInternal_cd)
        
        If  IntRetCd < 0 then
            if  IntRetCd = -1 then
                Call DisplayMsgBox("800012", "x","x","x")   ' 부서코드정보에 등록되지 않은 코드입니다.
            else
                Call DisplayMsgBox("800455", "x","x","x")   ' 자료권한이 없습니다.
            end if
		    frm1.txtTo_dept_nm.value = ""
		    frm1.txtTo_internal_cd.value = ""
            lsInternal_cd = ""
            frm1.txtTo_dept_cd.focus
            Set gActiveElement = document.ActiveElement
            txtTo_dept_cd_Onchange = true        
            Exit Function      
        Else           
            frm1.txtTo_dept_nm.value = strDept_nm
            frm1.txtTo_internal_cd.value = lsInternal_cd
        End if
    End if  
    
End Function

'========================================================================================================
' Name : txtYyyymm_DblClick
' Desc : 달력 Popup을 호출 
'========================================================================================================
Sub txtYyyymm_DblClick(Button)
	If Button = 1 Then
		Call SetFocusToDocument("M")	
		frm1.txtYyyymm.Action = 7
		frm1.txtYyyymm.focus
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>현금지급명세서출력</font></td>
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
								<TD CLASS="TD5" NOWRAP>작업년월</TD>
								<TD CLASS="TD6" NOWRAP>
								<script language =javascript src='./js/h6015oa1_txtYyyymm_txtYyyymm.js'></script>
								</TD>															
							</TR>
							<TR>
						  		<TD CLASS=TD5 NOWRAP>지급구분</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtProvCd" MAXLENGTH="2" SIZE="10"  ALT ="지급구분"   tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnProvCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenCode frm1.txtProvCd.value,'PROV_CD_POP'">
						                             <INPUT NAME="txtProvNm" MAXLENGTH="20" SIZE="20" ALT ="지급구분명" tag="14XXXU"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>부서코드</TD>
								<TD CLASS=TD6 NOWRAP><INPUT ID="txtFr_dept_cd" NAME="txtFr_dept_cd" ALT="부서코드" TYPE="Text" MAXLENGTH="10" SIZE=10 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBP_CD" align=top TYPE="BUTTON"ONCLICK="vbscript:Call OpenDept(0)">&nbsp;
								                     <INPUT ID="txtFr_dept_nm" NAME="txtFr_dept_nm" TYPE="Text" MAXLENGTH="50" SIZE=30 tag="14XXXU">&nbsp;~</TD>								
		                                             <INPUT NAME="txtFr_Internal_cd" ALT="내부부서코드" TYPE="HIDDEN" SiZE=7 MAXLENGTH=7 tag="14XXXU">
						    </TR>
						    <TR>
								<TD CLASS=TD5 NOWRAP></TD>
								<TD CLASS=TD6 NOWRAP><INPUT ID="txtTo_dept_cd" NAME="txtTo_dept_cd" ALT="" TYPE="Text" MAXLENGTH="18" SIZE=10 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnITEM_CD" align=top TYPE="BUTTON"ONCLICK="vbscript:Call OpenDept(1)">&nbsp;
								                     <INPUT ID="txtTo_dept_nm" NAME="txtTo_dept_nm" TYPE="Text" MAXLENGTH="40" SIZE=30 tag="14XXXU">&nbsp;</TD>
    			                                     <INPUT NAME="txtTo_Internal_cd" ALT="내부부서코드" TYPE="HIDDEN" SiZE=7 MAXLENGTH=7 tag="14XXXU"></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>지급방식</TD>
				        	    <TD CLASS="TD6"><INPUT TYPE="RADIO" CLASS="RADIO" NAME="txtProvType" TAG="1X" VALUE="기준금액 제외한 은행이체" CHECKED ID="ProvType1"><LABEL FOR="txtProvType1">기준금액 제외한 은행이체</LABEL></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP></TD>
  				        	    <TD CLASS="TD6"><INPUT TYPE="RADIO" CLASS="RADIO" NAME="txtProvType" TAG="1X" VALUE="기준금액 미만금액만 은행이체" ID="ProvType2"><LABEL FOR="txtProvType2">기준금액 미만 금액만 은행이체</LABEL></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP></TD>
  				        	    <TD CLASS="TD6"><INPUT TYPE="RADIO" CLASS="RADIO" NAME="txtProvType" TAG="1X" VALUE="모든 금액 은행이체" ID="ProvType3"><LABEL FOR="txtProvType3">모든 금액 은행이체</LABEL></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>기준금액</TD>
								<TD CLASS="TD6"><script language =javascript src='./js/h6015oa1_txtStandPrice_txtStandPrice.js'></script>&nbsp;원</TD>
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
                         <BUTTON NAME="btnRun"   CLASS="CLSSBTN" ONCLICK="vbscript:BtnPreview()" Flag=1>미리보기</BUTTON>&nbsp;
                         <BUTTON NAME="btnPrint" CLASS="CLSSBTN" ONCLICK="vbscript:FncBtnPrint()" Flag=1>인쇄</BUTTON>
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


