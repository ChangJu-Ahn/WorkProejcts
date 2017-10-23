<%@ LANGUAGE="VBSCRIPT" %>
<!--
'======================================================================================================
'*  1. Module Name          : 인사/급여관리 
'*  2. Function Name        : 급/상여공제관리관리 
'*  3. Program ID           : h5602oa1
'*  4. Program Name         : 통합상실신고(국민/고용)
'*  5. Program Desc         : 통합상실신고(국민/고용)
'*  6. Comproxy List        : +
'*  7. Modified date(First) : 2003/07/31
'*  8. Modified date(Last)  : 2003/07/31
'*  9. Modifier (First)     : choi yong chuel
'* 10. Modifier (Last)      : 
'* 11. Comment              : uniERP 2.5
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

<SCRIPT LANGUAGE=VBSCRIPT>
Option Explicit  
'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim IsOpenPop
Dim lgOldRow

<%StartDate	= GetSvrDate%>

'========================================================================================================
' Name : InitVariables()	
' Desc : Initialize value
'========================================================================================================
Sub InitVariables()

    lgIntFlgMode = Parent.OPMD_CMODE            'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
End Sub

'========================================================================================================
' Name : SetDefaultVal()	
' Desc : Set default value
'========================================================================================================
Sub SetDefaultVal()
    frm1.txtCust_cd.focus
    frm1.txtFr_acq_dt.text = UniConvDateAToB("<%=GetSvrDate%>",Parent.gServerDateFormat,Parent.gDateFormat)
    frm1.txtTo_acq_dt.text = frm1.txtFr_acq_dt.text
	frm1.txtRprt_dt.text   = frm1.txtFr_acq_dt.text
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
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)

	Call ggoOper.FormatDate(frm1.txtTo_acq_dt, Parent.gDateFormat, 1)
    Call ggoOper.FormatDate(frm1.txtFr_acq_dt, Parent.gDateFormat, 1)
	Call ggoOper.FormatDate(frm1.txtRprt_dt, Parent.gDateFormat, 1)
    
    Call InitVariables 
    Call FuncGetAuth(gStrRequestMenuID , Parent.gUsrID, lgUsrIntCd)     ' 자료권한:lgUsrIntCd ("%", "1%")       
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
' Name : FncQuery
' Desc : developer describe this line Called by MainQuery in Common.vbs
'========================================================================================================
Function FncQuery()
    Dim IntRetCD 
    
    FncQuery = False															 '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
    
    If  txtEmp_no_Onchange()  then
        Exit Function
    End If
    If  txtFr_Dept_cd_Onchange()  then
        Exit Function
    End If
    If  txtTo_Dept_cd_Onchange()  then
        Exit Function
    End If
    
    FncQuery = True                                                              '☜: Processing is OK

End Function
	
'========================================================================================================
' Name : FncNew
' Desc : developer describe this line Called by MainNew in Common.vbs
'========================================================================================================
Function FncNew()
    Dim IntRetCD 
    
    FncNew = False																 '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
    
    FncNew = True																 '☜: Processing is OK
End Function
	
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
' Name : FncSave
' Desc : developer describe this line Called by MainSave in Common.vbs
'========================================================================================================
Function FncSave()
    Dim IntRetCD 
    
    FncSave = False                                                              '☜: Processing is NG
    
    Err.Clear                                                                    '☜: Clear err status
    
    FncSave = True                                                               '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncCopy
' Desc : developer describe this line Called by MainSave in Common.vbs
' Keep : Make sure to clear primary key area
'========================================================================================================
Function FncCopy()
	Dim IntRetCD

    FncCopy = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
	
    FncCopy = True                                                               '☜: Processing is OK
    
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
	    arrParam(1) = ""'frm1.txtName.value			' Name Cindition
    Else
	    arrParam(0) = frm1.txtEmp_no.value			' Code Condition
	    arrParam(1) = ""'frm1.txtName.value		' Name Cindition
	End If
    
    arrParam(2) = lgUsrIntcd
    
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

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	If iWhere = 0 Then
		arrParam(0) = frm1.txtFr_dept_cd.value			            '  Code Condition
	Else
		arrParam(0) = frm1.txtTo_dept_cd.value			            ' Code Condition
	End If
    arrParam(1) = frm1.txtRprt_dt.Text
	arrParam(2) = lgUsrIntcd
	
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


'======================================================================================================
'	Name : OpenCode()
'	Description : Code PopUp at vspdData
'=======================================================================================================
Function OpenCode(Byval strCode, Byval iWhere, ByVal Row)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	Select Case iWhere
	    Case "SECT_CD_POP"
	        arrParam(0) = "근무구역 팝업"			        ' 팝업 명칭 
	    	arrParam(1) = "B_MINOR"							    ' TABLE 명칭 
	    	arrParam(2) = frm1.txtSect_cd.value     			' Code Condition
	    	arrParam(3) = ""'frm1.txtSect_nm.value    				' Name Cindition
	    	arrParam(4) = "MAJOR_CD = " & FilterVar("H0035", "''", "S") & ""	               	' Where Condition
	    	arrParam(5) = "근무구역코드"  		            ' TextBox 명칭 

	    	arrField(0) = "MINOR_CD"						   	' Field명(0)
	    	arrField(1) = "MINOR_NM"    				  		' Field명(1)
	    	arrField(2) = ""    				        		' Field명(2)

	    	arrHeader(0) = "근무구역코드"	     			' Header명(0)
	    	arrHeader(1) = "근무구역코드명"	   		        ' Header명(1)
	    	arrHeader(2) = ""   	    						' Header명(1)
	    Case "CUST_CD_POP"
	        arrParam(0) = "신고사업장 팝업"			        ' 팝업 명칭 
'	    	arrParam(1) = "B_BIZ_AREA"							    ' TABLE 명칭 
	    	arrParam(1) = "HFA100T"							    ' TABLE 명칭 
	    	arrParam(2) = frm1.txtCust_cd.value     			' Code Condition
	    	arrParam(3) = ""'frm1.txtCust_nm.value									' Name Cindition
	    	arrParam(4) = ""	               	                ' Where Condition
	    	arrParam(5) = "신고사업장코드"  		        ' TextBox 명칭 

'	    	arrField(0) = "BIZ_AREA_CD"						   	' Field명(0)
'	    	arrField(1) = "BIZ_AREA_NM"    				  		' Field명(1)
	    	arrField(0) = "YEAR_AREA_CD"						   	' Field명(0)
	    	arrField(1) = "YEAR_AREA_NM"    				  		' Field명(1)

	    	arrField(2) = ""    				        		' Field명(2)

	    	arrHeader(0) = "신고사업장코드"	     			' Header명(0)
	    	arrHeader(1) = "신고사업장코드명"		        ' Header명(1)
	    	arrHeader(2) = ""   	    						' Header명(1)
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
		Call SetCode(arrRet, iWhere)
	End If	

End Function


'======================================================================================================
'	Name : SetCode()
'	Description : Code PopUp에서 Return되는 값 setting
'=======================================================================================================
Function SetCode(Byval arrRet, Byval iWhere)

	With frm1

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
'======================================================================================================
' Function Name : FncBtnPrint
' Function Desc : This function is related to Print Button
'=======================================================================================================%>

Function FncBtnPrint() 
	Dim condvar
	Dim lngPos
	Dim intCnt
	Dim emp_no, fr_dt, to_dt, sect_cd ,cust_cd , singo_dt, singo_org_no, print_emp_no ,singo_nm
	Dim Fr_dept_cd , To_dept_cd, rFrDept ,rToDept ,IntRetCd
    Dim StrEbrFile , StrEbrFile2 ,StrEbrFile3
	Dim StrEbrimage , h_image_id
	Dim strFromDt , strToDt
		    	 
    If Not chkField(Document, "1") Then									<%'⊙: This function check indispensable field%>
       Exit Function
    End If
   
    '-----------------------시작날짜부터 종료날짜에러 체크--------------------------------------------------	
    strFromDt = frm1.txtFr_acq_dt.Text
    strToDt   = frm1.txtTo_acq_dt.Text
    
    If (strFromDt <> "") AND (strToDt <> "") Then
    	IF CompareDateByFormat(frm1.txtFr_acq_dt.Text,frm1.txtTo_acq_dt.Text,frm1.txtFr_acq_dt.Alt,frm1.txtTo_acq_dt.Alt,"970025",frm1.txtFr_acq_dt.UserDefinedFormat,Parent.gComDateType,False)=False Then
            Call DisplayMsgbox("970025","X","시작일자","종료일자")	
            frm1.txtTo_acq_dt.focus
            Set gActiveElement = document.activeElement
            Exit Function
        End if 
    End if 
'------------------------------------------------------------------------------------------------	

	If frm1.txtCap_type(0).checked Then
		StrEbrFile  = "h5602oa1_1.ebr"
		StrEbrFile2 = "h5602oa1_1T.ebr"		
		StrEbrFile3 = "h5602oa1_2.ebr"
	Else
		StrEbrFile  = "h5602oa1_3.ebr"
		StrEbrFile2 = "h5602oa1_3T.ebr"		
	End If
	
    singo_dt= UNIConvDateToYYYYMMDD(frm1.txtRprt_dt.text,Parent.gDateFormat,"")
	fr_dt = UNIConvDateToYYYYMMDD(frm1.txtFr_acq_dt.text, Parent.gDateFormat, Parent.gServerDateType)
	to_dt = UNIConvDateToYYYYMMDD(frm1.txtTo_acq_dt.text, Parent.gDateFormat, Parent.gServerDateType)


	emp_no = frm1.txtEmp_no.value
 
	sect_cd = frm1.txtSect_cd.value
	cust_cd = frm1.txtCust_cd.value
		
	fr_dept_cd = frm1.txtFr_dept_cd.value
	to_dept_cd = frm1.txtTo_dept_cd.value
	singo_dt = singo_dt											' ggoOper.RetFormat(frm1.txtRprt_dt.text, "yyyyMMdd")
	singo_nm = frm1.txtWrt_prsn.value


	If sect_cd = "" then
		sect_cd = "%"
	End if	
			
	If emp_no = "" then
		emp_no = "%"
		frm1.txtName.value = ""
	End if	

	If singo_nm = "" Then
		singo_nm = ""
	End if
    
    If  txtEmp_no_Onchange()  then
        Exit Function
    End If

    If  txtFr_Dept_cd_Onchange()  then
        Exit Function
    End If
    If  txtTo_Dept_cd_Onchange() then
        Exit Function
    End If
    If  txtSect_cd_OnChange()  then
        Exit Function
    End If
    If  txtCust_cd_Onchange()  then
        Exit Function
    End If        
    
    
    Fr_dept_cd = frm1.txtFr_internal_cd.value
    To_dept_cd = frm1.txtTo_internal_cd.value
    
    If fr_dept_cd = "" then
        IntRetCd = FuncGetTermDept(lgUsrIntCd ,"", rFrDept ,rToDept)
		frm1.txtFr_internal_cd.value = rFrDept
		frm1.txtFr_dept_nm.value = ""
		frm1.txtFr_dept_cd.focus
	End If	
	
	If to_dept_cd = "" then
        IntRetCd = FuncGetTermDept(lgUsrIntCd ,"", rFrDept ,rToDept)
		frm1.txtTo_internal_cd.value = rToDept
		frm1.txtTo_dept_nm.value = ""
		frm1.txtto_dept_cd.focus
	End If  
    
    Fr_dept_cd = frm1.txtFr_internal_cd.value
    To_dept_cd = frm1.txtTo_internal_cd.value
    
    If (Fr_dept_cd<> "") AND (To_dept_cd<>"") Then       
        If Fr_dept_cd > To_dept_cd then
	        Call DisplayMsgbox("800153","X","X","X")	'시작부서코드는 종료부서코드보다 작아야합니다.
            frm1.txtFr_dept_cd.focus()
            Set gActiveElement = document.activeElement
            Exit Function
        End IF 
    END IF   
    '----------------------------------------------------------------------------------------------						
   	
	condvar = condvar & "emp_no|" & emp_no
	condvar = condvar & "|fr_dt|" & fr_dt
	condvar = condvar & "|to_dt|" & to_dt
	condvar = condvar & "|sect_cd|" & sect_cd
	condvar = condvar & "|cust_cd|" & cust_cd
	condvar = condvar & "|fr_dept_cd|" & fr_dept_cd
	condvar = condvar & "|to_dept_cd|" & to_dept_cd
	condvar = condvar & "|Regdt|" & singo_dt
	condvar = condvar & "|singo_nm|" & singo_nm
	
	IF StrEbrFile <> "" Then
 	    call FncEBRPrint(EBAction , StrEbrFile , condvar)
	End if
		
	IF StrEbrFile2 <> "" Then
 	    call FncEBRPrint(EBAction , StrEbrFile2 , condvar)
	End if

	IF StrEbrFile3 <> "" Then
 	    call FncEBRPrint(EBAction , StrEbrFile3 , condvar)
	End if

End Function


'========================================================================================
' Function Name : FncBtnPreview()
' Function Desc : This function is related to Preview Button
'========================================================================================

Function FncBtnPreview()
	Dim condvar
	Dim lngPos
	Dim intCnt
	Dim emp_no, fr_dt, to_dt, sect_cd ,cust_cd , singo_dt, singo_org_no, print_emp_no  ,singo_nm
	Dim Fr_dept_cd , To_dept_cd, rFrDept ,rToDept ,IntRetCd
    Dim StrEbrFile , StrEbrFile2 ,StrEbrFile3
	Dim StrEbrimage , h_image_id
	Dim strFromDt , strToDt
		    	 
    If Not chkField(Document, "1") Then									<%'⊙: This function check indispensable field%>
       Exit Function
    End If
   
    '-----------------------시작날짜부터 종료날짜에러 체크--------------------------------------------------	
    strFromDt = frm1.txtFr_acq_dt.Text
    strToDt   = frm1.txtTo_acq_dt.Text
    
    If (strFromDt <> "") AND (strToDt <> "") Then
    	IF CompareDateByFormat(frm1.txtFr_acq_dt.Text,frm1.txtTo_acq_dt.Text,frm1.txtFr_acq_dt.Alt,frm1.txtTo_acq_dt.Alt,"970025",frm1.txtFr_acq_dt.UserDefinedFormat,Parent.gComDateType,False)=False Then
            Call DisplayMsgbox("970025","X","시작일자","종료일자")	
            frm1.txtTo_acq_dt.focus
            Set gActiveElement = document.activeElement
            Exit Function
        End if 
    End if 
'------------------------------------------------------------------------------------------------	

	If frm1.txtCap_type(0).checked Then
		StrEbrFile  = "h5602oa1_1.ebr"
		StrEbrFile2 = "h5602oa1_1T.ebr"		
		StrEbrFile3 = "h5602oa1_2.ebr"
	Else
		StrEbrFile  = "h5602oa1_3.ebr"
		StrEbrFile2 = "h5602oa1_3T.ebr"		
	End If
	
	
    singo_dt= UniConvDateToYYYYMMDD(frm1.txtRprt_dt.text,Parent.gDateFormat,"")
	fr_dt = UNIConvDateToYYYYMMDD(frm1.txtFr_acq_dt.text, Parent.gDateFormat, Parent.gServerDateType)
	to_dt = UNIConvDateToYYYYMMDD(frm1.txtTo_acq_dt.text, Parent.gDateFormat, Parent.gServerDateType)


	emp_no = frm1.txtEmp_no.value
	sect_cd = frm1.txtSect_cd.value
	cust_cd = frm1.txtCust_cd.value
	fr_dept_cd = frm1.txtFr_dept_cd.value
	to_dept_cd = frm1.txtTo_dept_cd.value
	singo_dt = singo_dt											' ggoOper.RetFormat(frm1.txtRprt_dt.text, "yyyyMMdd")
	singo_nm = frm1.txtWrt_prsn.value

	If sect_cd = "" then
		sect_cd = "%"
	End if	
			
	If emp_no = "" then
		emp_no = "%"
		frm1.txtName.value = ""
	End if	

	If singo_nm = "" Then
		singo_nm = ""
	End if
    
    If  txtEmp_no_Onchange()  then
        Exit Function
    End If

    If  txtFr_Dept_cd_Onchange()  then
        Exit Function
    End If
    If  txtTo_Dept_cd_Onchange() then
        Exit Function
    End If
    If  txtSect_cd_OnChange()  then
        Exit Function
    End If
    If  txtCust_cd_Onchange()  then
        Exit Function
    End If  
       
    
    Fr_dept_cd = frm1.txtFr_internal_cd.value
    To_dept_cd = frm1.txtTo_internal_cd.value
    
    If fr_dept_cd = "" then
        IntRetCd = FuncGetTermDept(lgUsrIntCd ,"", rFrDept ,rToDept)
		frm1.txtFr_internal_cd.value = rFrDept
		frm1.txtFr_dept_nm.value = ""
		frm1.txtfr_dept_cd.focus
	End If	

	If to_dept_cd = "" then
        IntRetCd = FuncGetTermDept(lgUsrIntCd ,"", rFrDept ,rToDept)
		frm1.txtTo_internal_cd.value = rToDept
		frm1.txtTo_dept_nm.value = ""
		frm1.txtto_dept_cd.focus
	End If  
    
    Fr_dept_cd = frm1.txtFr_internal_cd.value
    To_dept_cd = frm1.txtTo_internal_cd.value

    If (Fr_dept_cd<> "") AND (To_dept_cd<>"") Then       
        If Fr_dept_cd > To_dept_cd then
	        Call DisplayMsgbox("800153","X","X","X")	'시작부서코드는 종료부서코드보다 작아야합니다.
            frm1.txtFr_dept_cd.focus()
            Set gActiveElement = document.activeElement
            Exit Function
        End IF 
    END IF   
    '----------------------------------------------------------------------------------------------						
	condvar = condvar & "emp_no|" & emp_no
	condvar = condvar & "|fr_dt|" & fr_dt
	condvar = condvar & "|to_dt|" & to_dt
	condvar = condvar & "|sect_cd|" & sect_cd
	condvar = condvar & "|cust_cd|" & cust_cd
	condvar = condvar & "|fr_dept_cd|" & fr_dept_cd
	condvar = condvar & "|to_dept_cd|" & to_dept_cd
	condvar = condvar & "|Regdt|" & singo_dt
	condvar = condvar & "|singo_nm|" & singo_nm

	IF StrEbrFile <> "" Then
 	    call FncEBRPreview(StrEbrFile , condvar)
	End if
	
	IF StrEbrFile2 <> "" Then
 	    call FncEBRPreview(StrEbrFile2 , condvar)
	End if
	
	IF StrEbrFile3 <> "" Then
 	    call FncEBRPreview(StrEbrFile3 , condvar)
	End if

End Function
'======================================================================================================
'   Event Name : txtSect_cd_OnChange
'   Event Desc : 근무구역코드 에러체크 
'=======================================================================================================
Function txtSect_cd_OnChange()
    Dim iDx
    Dim IntRetCd
        
    If frm1.txtSect_cd.value = "" Then
        frm1.txtSect_nm.value = ""
    ELSE
        IntRetCd = CommonQueryRs(" MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0035", "''", "S") & " AND MINOR_CD =  " & FilterVar(frm1.txtSect_cd.value , "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) 
        IF IntRetCd = False THEN
            Call DisplayMsgBox("800054", "x","x","x")   '등록되지 않은 코드입니다 
                
	        frm1.txtSect_nm.value = ""
	        frm1.txtSect_cd.focus
	        Set gActiveElement = document.ActiveElement
	        txtSect_cd_OnChange = true
	    Else
	        frm1.txtSect_nm.value = Trim(Replace(lgF0, Chr(11), ""))
	    End If
    End If
    	    
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
        IntRetCd = CommonQueryRs(" YEAR_AREA_NM "," HFA100T "," YEAR_AREA_CD =  " & FilterVar(frm1.txtCust_cd.value , "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
        If IntRetCd = false then
			Call DisplayMsgbox("800054","X","X","X")	'등록되지 않은 코드입니다.
			 frm1.txtCust_nm.value = ""
             frm1.txtCust_cd.focus
            Set gActiveElement = document.ActiveElement
            txtCust_cd_Onchange = true            
        Else
			frm1.txtCust_nm.value = Trim(Replace(lgF0,Chr(11),""))
        End if
    End if

End Function    
'========================================================================================================
'   Event Name : txtEmp_no_Onchange           
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

    If  frm1.txtEmp_no.value = "" Then
		frm1.txtName.value = ""
    Else
	    IntRetCd = FuncGetEmpInf2(frm1.txtEmp_no.value,lgUsrIntCd,strName,strDept_nm,_
	                strRoll_pstn, strPay_grd1, strPay_grd2, strEntr_dt, strInternal_cd)
	                
	    If  IntRetCd < 0 then
	        If  IntRetCd = -1 then
    			Call DisplayMsgbox("800048","X","X","X")	'해당사원은 존재하지 않습니다.
            Else
                Call DisplayMsgbox("800454","X","X","X")	'자료에 대한 권한이 없습니다.
            End if
			frm1.txtName.value = ""
            Frm1.txtEmp_no.focus 
            Set gActiveElement = document.ActiveElement
            txtEmp_no_Onchange = true
        Else
			frm1.txtName.value = strName
        End if 
    End if  
End Function

'========================================================================================================
'   Event Name : txtFr_dept_cd_Onchange
'   Event Desc :
'========================================================================================================
Function txtFr_dept_cd_Onchange()
    Dim IntRetCd,Dept_Nm,Internal_cd
    
    If frm1.txtFr_dept_cd.value = "" Then
		frm1.txtFr_dept_nm.value = ""
		frm1.txtFr_internal_cd.value = ""
    Else
        IntRetCd = FuncDeptName(frm1.txtFr_dept_cd.value , UNIConvDate(frm1.txtRprt_dt.Text), lgUsrIntCd,Dept_Nm , Internal_cd)
	    If  IntRetCd < 0 then
	        If  IntRetCd = -1 then
    			Call DisplayMsgbox("800098","X","X","X")	'부서정보코드에 등록되지 않은 코드입니다.
            Else
                Call DisplayMsgbox("800454","X","X","X")	'자료에 대한 권한이 없습니다.
            End if
		    frm1.txtFr_dept_nm.value = ""
		    frm1.txtFr_internal_cd.value = ""
            frm1.txtFr_dept_cd.focus
            Set gActiveElement = document.ActiveElement 
            
            txtFr_dept_cd_Onchange = true
            Exit Function      
        Else
			frm1.txtFr_dept_nm.value = Dept_Nm
		    frm1.txtFr_internal_cd.value = Internal_cd
        End if 
    End if  
    
End Function

'========================================================================================================
'   Event Name : txtTo_dept_cd_Onchange
'   Event Desc :
'========================================================================================================
Function txtTo_dept_cd_Onchange()
    Dim IntRetCd,Dept_Nm,Internal_cd
    
    If frm1.txtTo_dept_cd.value = "" Then
		frm1.txtTo_dept_nm.value = ""
		frm1.txtTo_internal_cd.value = ""
    Else
        IntRetCd = FuncDeptName(frm1.txtTo_dept_cd.value , UNIConvDate(frm1.txtRprt_dt.Text), lgUsrIntCd,Dept_Nm , Internal_cd)
	    If  IntRetCd < 0 then
	        If  IntRetCd = -1 then
    			Call DisplayMsgbox("800098","X","X","X")	'부서정보코드에 등록되지 않은 코드입니다.
            Else
                Call DisplayMsgbox("800454","X","X","X")	'자료에 대한 권한이 없습니다.
            End if
		    frm1.txtTo_dept_nm.value = ""
		    frm1.txtTo_internal_cd.value = ""
            frm1.txtTo_dept_cd.focus
            Set gActiveElement = document.ActiveElement 
            
            txtTo_dept_cd_Onchange = true
            Exit Function      
        Else
			frm1.txtTo_dept_nm.value = Dept_Nm
		    frm1.txtTo_internal_cd.value = Internal_cd
        End if 
    End if  
    
End Function
'======================================================================================================
'   Event Name : txtYyyymm_DblClick
'   Event Desc : 달력 Popup을 호출 
'=======================================================================================================%>

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
'=======================================================================================================
'   Event Name : txt________Keypress
'   Event Desc : enter key down시에 조회한다.
'=======================================================================================================
Sub txtEmp_no_Keypress(Key)
    If Key = 13 Then
        Call MainQuery()
    End If
End Sub     

Sub txtFr_dept_cd_Keypress(Key)
    If Key = 13 Then
        Call MainQuery()
    End If
End Sub

Sub txtto_dept_cd_Keypress(Key)
    If Key = 13 Then
        Call MainQuery()
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
								<td background="../../image/table/seltab_up_bg.gif"><img src="../../image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../image/table/seltab_up_bg.gif" align="center" CLASS="CLSLTAB"><font color=white>통합상실신고(국민/고용)</font></td>
								<td background="../../image/table/seltab_up_bg.gif" align="right"><img src="../../image/table/seltab_up_right.gif" width="10" height="23"></td>
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
								<TD CLASS="TD5" NOWRAP>자격상실</TD>
				        	    <TD CLASS="TD6"><INPUT TYPE="RADIO" CLASS="RADIO" NAME="txtCap_type" TAG=11 VALUE="국민연금.고용보험" CHECKED ID="med_entr_flag1"><LABEL FOR="txtCap_type1">국민연금.고용보험</LABEL>&nbsp;
  				        	                    <INPUT TYPE="RADIO" CLASS="RADIO" NAME="txtCap_type" TAG=11 VALUE="건강보험)" ID="med_entr_flag2"><LABEL FOR="txtCap_type2">건강보험</LABEL></TD>
							</TR>						   
							<TR>
								<TD CLASS=TD5 NOWRAP>신고사업장</TD>
								<TD CLASS=TD6 NOWRAP><INPUT ID="txtCust_cd" NAME="txtCust_cd" ALT="신고사업장" TYPE="Text" SiZE=10 MAXLENGTH=10 tag="12XXXU"><IMG SRC="../../image/btnPopup.gif" NAME="btnSItemDC" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenCode('x', 'CUST_CD_POP', 'x')">
								                     <INPUT NAME="txtCust_nm" ALT="신고사업장" TYPE="Text" SiZE=20 MAXLENGTH=100 tag="14XXXU"></td>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>근무구역</TD>
								<TD CLASS=TD6 NOWRAP><INPUT ID="txtSect_cd" NAME="txtSect_cd" ALT="근무구역" TYPE="Text" SiZE="10" MAXLENGTH="7" tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSItemDC" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenCode('x', 'SECT_CD_POP', 'x')">
								                     <INPUT NAME="txtSect_nm" ALT="근무구역" TYPE="Text" SiZE="20" MAXLENGTH="20" tag="14XXXU"></td>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>해당일</TD>
								<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/h5602oa1_txtFr_acq_dt_txtFr_acq_dt.js'></script>&nbsp;~&nbsp;
								                    <script language =javascript src='./js/h5602oa1_txtTo_acq_dt_txtTo_acq_dt.js'></script></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>작성자</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" ID="txtWrt_prsn" NAME="txtwrt_prsn" SIZE="15" MAXLENGTH="20" tag="11XXXU" ALT="작성자"></TD>
							</TR>	

							<TR>
								<TD CLASS=TD5 NOWRAP>신고일</TD>
								<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/h5602oa1_txtRprt_dt_txtRprt_dt.js'></script></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>대상자</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtEmp_no" ALT="대상자" TYPE="Text" MAXLENGTH="13" SIZE="13" tag="11XXXU"><IMG SRC="../../image/btnPopup.gif" NAME="btnBP_CD" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenEmpName (1)">
								                     <INPUT NAME="txtName" TYPE="Text" MAXLENGTH="30" SIZE="20" tag="14XXXU"></TD>	
						    </TR>
							<TR>			
							    <TD CLASS=TD5 NOWRAP>부서</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtFr_dept_cd" ALT="부서코드" TYPE="Text" SiZE="10" MAXLENGTH="10" tag="11XXXU"><IMG SRC="../../image/btnPopup.gif" NAME="btnSItemDC" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenDept(0)">
			                                         <INPUT NAME="txtFr_dept_nm" ALT="부서코드명" TYPE="Text" SiZE="20" MAXLENGTH="40" tag="14XXXU">&nbsp;~
		                                             <INPUT NAME="txtFr_internal_cd" ALT="내부부서코드" TYPE="hidden" SiZE="7" MAXLENGTH="7" tag="14XXXU">
		                    </TR>
		                    <TR>    
		                        <TD CLASS=TD5 NOWRAP></TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtto_dept_cd" ALT="부서코드" TYPE="Text" SIZE="10" MAXLENGTH="10"  tag="11XXXU"><IMG SRC="../../image/btnPopup.gif" NAME="btnname" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenDept(1)">
							                         <INPUT NAME="txtto_dept_nm" ALT="부서코드명" TYPE="Text" SIZE="20" MAXLENGTH="40"  tag="14XXXU">
    			                                     <INPUT NAME="txtTo_internal_cd" ALT="내부부서코드" TYPE="hidden" SiZE="7" MAXLENGTH="7" tag="14XXXU"></TD>
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
		                <BUTTON NAME="btnPreview" CLASS="CLSSBTN" onclick="VBScript:FncBtnPreview()">미리보기</BUTTON>&nbsp;
		                <BUTTON NAME="btnPrint"   CLASS="CLSSBTN" OnClick="VBScript:FncBtnPrint()">인쇄</BUTTON></TD>
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

