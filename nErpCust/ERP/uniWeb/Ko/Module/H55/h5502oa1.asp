<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
'*  1. Module Name          : 인사/급여관리 
'*  2. Function Name        : 근무이력관리 
'*  3. Program ID           : h5502oa1
'*  4. Program Name         : 고용보험자격취득/상실신고서 
'*  5. Program Desc         : 고용보험자격취득/상실신고서 
'*  6. Comproxy List        : +
'*  7. Modified date(First) : 2001/09/03
'*  8. Modified date(Last)  : 2001/09/03
'*  9. Modifier (First)     : Shin Kwang-Ho/mok yong bin
'* 10. Modifier (Last)      : 신광호 
'* 11. Comment              :
=======================================================================================================-->
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
    frm1.txtSect_cd.focus 
	frm1.txtFr_dt.text	= UniConvDateAToB("<%=GetSvrDate%>",Parent.gServerDateFormat,Parent.gDateFormat)
	Call ggoOper.FormatDate(frm1.txtFr_dt, Parent.gDateFormat, 1)
	
	frm1.txtTo_dt.text	= UniConvDateAToB("<%=GetSvrDate%>",Parent.gServerDateFormat,Parent.gDateFormat)
	Call ggoOper.FormatDate(frm1.txtTo_dt, Parent.gDateFormat, 1)

	frm1.txtrprt_dt.text	= UniConvDateAToB("<%=GetSvrDate%>",Parent.gServerDateFormat,Parent.gDateFormat)
	Call ggoOper.FormatDate(frm1.txtrprt_dt, Parent.gDateFormat, 1)
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
    Call FuncGetAuth(gStrRequestMenuID , Parent.gUsrID, lgUsrIntCd)                ' 자료권한:lgUsrIntCd ("%", "1%")
                                         
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
    
    If Not chkField(Document, "1") Then									<%'⊙: This function check indispensable field%>
       Exit Function
    End If

    If txtEmp_no_Onchange() Then        'enter key 로 조회시 사원을 check후 해당사항 없으면 query종료...
        Exit Function
    End if
    
    If txtFr_dept_cd_Onchange() Then        'enter key 로 조회시 부서코드를 check후 해당사항 없으면 query종료...
        Exit Function
    End if
    
    If txtTo_dept_cd_Onchange() Then        'enter key 로 조회시 부서코드를 check후 해당사항 없으면 query종료...
        Exit Function
    End if

    If txtSect_cd_Onchange() Then        'enter key 로 조회시 부서코드를 check후 해당사항 없으면 query종료...
        Exit Function
    End if
    
    FncQuery = True                                                              '☜: Processing is OK

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
	arrParam(2) = lgUsrIntCd                    ' 자료권한 Condition  

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
'	Description : Item Popup에서 Return되는 값 setting
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
' Desc : 부서 POPUP
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
    arrParam(1) = ""
	arrParam(2) = lgUsrIntCd                    ' 자료권한 Condition  

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

Function btnPrint_Print()
    Dim strFrDt
    Dim strToDt
    Dim strFrDept
    Dim strToDept
    Dim strMin
    Dim strMax
            
    If Not chkField(Document, "1") Then									<%'⊙: This function check indispensable field%>
       Exit Function
    End If
    
    strFrDt = Trim(frm1.txtFr_dt.Text)
    strToDt = Trim(frm1.txtTo_dt.Text)
    
    If strFrDt = "" then
        strFrDt =  UniConvYYYYMMDDToDate(Parent.gDateFormat,"1950","01","01")
    End if
    If strToDt = "" then
        strToDt = UniConvYYYYMMDDToDate(Parent.gDateFormat,"2100","12","31")
    End If
    If  ValidDateCheck(frm1.txtFr_dt,frm1.txtTo_dt)= False Then   'strFrDt > strToDt then
        Exit Function
    End if 

    If txtEmp_no_Onchange() Then        'enter key 로 조회시 사원을 check후 해당사항 없으면 query종료...
        Exit Function
    End if
    
    If txtFr_dept_cd_Onchange() Then        'enter key 로 조회시 부서코드를 check후 해당사항 없으면 query종료...
        Exit Function
    End if
    
    If txtTo_dept_cd_Onchange() Then        'enter key 로 조회시 부서코드를 check후 해당사항 없으면 query종료...
        Exit Function
    End if

    If txtSect_cd_Onchange() Then        'enter key 로 조회시 부서코드를 check후 해당사항 없으면 query종료...
        Exit Function
    End if
    
    Call FuncGetTermDept(lgUsrIntCd,"",strMin,strMax)
    
    strFrDept = Trim(frm1.txtFr_internal_cd.value)
    strToDept = Trim(frm1.txtTo_internal_cd.value)

    If strFrDept = "" Then       
       strFrDept = strMin
    End if

    If strToDept = "" then
       strToDept = strMax
    End If

    If strFrDept > strToDept then
	    Call DisplayMsgbox("970025","X","시작부서","종료부서")	'시작부서는 종료부서보다 작아야 합니다.
        frm1.txtFr_dept_cd.focus
        Set gActiveElement = document.activeElement
        Exit Function
    End if 
    
	with frm1
	    If .txtcap_type(0).Checked Then
	        Call FncBtnPrint("h5502oa1_1") 
	    Else	        
	        Call FncBtnPrint("h5502oa1_2") 
	        Call FncBtnPrint("h5502oa1_3") 
	    End If
	End with

End Function
		
'======================================================================================================
' Function Name : FncBtnPrint
' Function Desc : This function is related to Print Button
'=======================================================================================================

Function FncBtnPrint(iValue) 
	Dim condvar
	Dim lngPos
	Dim intCnt
    Dim StrEbrFile
    Dim strFrDt
    Dim strToDt
    Dim strFrDept
    Dim strToDept
    Dim strMin
    Dim strMax
    Dim strSect_cd,rel_cd1, rel_cd2
    Dim ls_work_area,ls_comp_na,ls_tel_no,ls_repre_nm,ls_comp_addr
    Dim IntRetCD
    Dim strYear,strMonth,strDay
	Dim strYear2,strMonth2,strDay2
	Dim strYear3,strMonth3,strDay3	
    Dim ls_comp_no1,ls_comp_ser,ls_saupjang1,ls_repre_nm1,ls_juso1,ls_cust_telno1
	Dim fr_dt, to_dt, fr_dept_cd, to_dept_cd, rprt_dt, wrt_prsn, emp_no
	Dim ObjName
          
    strFrDt = Trim(frm1.txtFr_dt.Text)
    strToDt = Trim(frm1.txtTo_dt.Text)
    
    Call FuncGetTermDept(lgUsrIntCd,"",strMin,strMax)
    
    strFrDept = Trim(frm1.txtFr_internal_cd.value)
    strToDept = Trim(frm1.txtTo_internal_cd.value)

    If strFrDept = "" Then       
       strFrDept = strMin
    End if
    If strToDept = "" then
       strToDept = strMax
    End If
    If txtEmp_no_Onchange() Then        'enter key 로 조회시 사원을 check후 해당사항 없으면 query종료...
        Exit Function
    End if
    
    If txtFr_dept_cd_Onchange() Then        'enter key 로 조회시 부서코드를 check후 해당사항 없으면 query종료...
        Exit Function
    End if
    
    If txtTo_dept_cd_Onchange() Then        'enter key 로 조회시 부서코드를 check후 해당사항 없으면 query종료...
        Exit Function
    End if

    If txtSect_cd_Onchange() Then        'enter key 로 조회시 부서코드를 check후 해당사항 없으면 query종료...
        Exit Function
    End if
        
    strSect_cd = frm1.txtSect_cd.value
    rel_cd1 = strSect_cd

    IntRetCD = CommonQueryRs(" reference "," B_CONFIGURATION "," MAJOR_CD=" & FilterVar("H0035", "''", "S") & " and MINOR_CD= " & FilterVar(strSect_cd, "''", "S") & " and seq_no=1 ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
 
    If IntRetCD=False  Then
       Call DisplayMsgBox("800490","X","X","X")      '공통기준정보-환경등록에서 근무구역코드(H0035) 각각에 연결되는 사업장코드를 reference 컬럼에 등록하세요.
       frm1.txtSect_cd.focus()
       Exit Function
    Else
       rel_cd2=Trim(Replace(lgF0,Chr(11),""))
       IntRetCD = CommonQueryRs(" biz_area_nm,tel_no,repre_nm,addr "," b_biz_area "," biz_area_cd= " & FilterVar(rel_cd2, "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
                
       If IntRetCD=False  Then
           Call DisplayMsgBox("800491","X","X","X")  '공통기준정보-환경등록에서 근무구역코드(H0035)의 reference에 등록된 사업장코드가 존재하지 않습니다.
           frm1.txtSect_cd.focus()
           Exit Function
       Else
           ls_comp_na   =Trim(Replace(lgF0,Chr(11),"")) 
           ls_tel_no    =Trim(Replace(lgF1,Chr(11),""))
           ls_repre_nm  =Trim(Replace(lgF2,Chr(11),""))
           ls_comp_addr =Trim(Replace(lgF3,Chr(11),""))
       End If
    End If
           
            IntRetCD = CommonQueryRs(" MINOR_NM "," B_MINOR "," MAJOR_CD=" & FilterVar("H0069", "''", "S") & " and MINOR_CD='" & strSect_cd & "1'",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
           If IntRetCD=True  Then
                    ls_comp_no1   =Trim(Replace(lgF0,Chr(11),""))     
           Else
                    ls_comp_no1   = "%"
           End If

              IntRetCD = CommonQueryRs(" MINOR_NM "," B_MINOR "," MAJOR_CD=" & FilterVar("H0069", "''", "S") & " and MINOR_CD='" & strSect_cd & "2'",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
           If IntRetCD=True  Then
                    ls_comp_ser   =Trim(Replace(lgF0,Chr(11),"")) 
           Else
                    ls_comp_ser   = "%"
           End If

              IntRetCD = CommonQueryRs(" MINOR_NM "," B_MINOR "," MAJOR_CD=" & FilterVar("H0069", "''", "S") & " and MINOR_CD='" & strSect_cd & "3'",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
           If IntRetCD=True  Then
                    ls_saupjang1   =Trim(Replace(lgF0,Chr(11),"")) 
           Else
                    ls_saupjang1   = "%"
           End If

              IntRetCD = CommonQueryRs(" MINOR_NM "," B_MINOR "," MAJOR_CD=" & FilterVar("H0069", "''", "S") & " and MINOR_CD='" & strSect_cd & "4'",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
           If IntRetCD=True  Then
                    ls_repre_nm1   =Trim(Replace(lgF0,Chr(11),"")) 
           Else
                    ls_repre_nm1   = "%"
           End If

              IntRetCD = CommonQueryRs(" MINOR_NM "," B_MINOR "," MAJOR_CD=" & FilterVar("H0069", "''", "S") & " and MINOR_CD='" & strSect_cd & "5'",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
           If IntRetCD=True  Then
                    ls_juso1   =Trim(Replace(lgF0,Chr(11),"")) 
           Else
                    ls_juso1   = "%"
           End If
		
              IntRetCD = CommonQueryRs(" MINOR_NM "," B_MINOR "," MAJOR_CD=" & FilterVar("H0069", "''", "S") & " and MINOR_CD='" & strSect_cd & "6'",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
           If IntRetCD=True  Then
                    ls_cust_telno1   =Trim(Replace(lgF0,Chr(11),"")) 
           Else
                    ls_cust_telno1   = "%"
           End If
	
	Call ExtractDateFrom(frm1.txtFr_dt.text,frm1.txtFr_dt.UserDefinedFormat,Parent.gComDateType,strYear,strMonth,strDay)
	Call ExtractDateFrom(frm1.txtTo_dt.text,frm1.txtTo_dt.UserDefinedFormat,Parent.gComDateType,strYear2,strMonth2,strDay2)
	Call ExtractDateFrom(frm1.txtrprt_dt.text,frm1.txtrprt_dt.UserDefinedFormat,Parent.gComDateType,strYear3,strMonth3,strDay3) 
	
	fr_dt   = strYear&strMonth&strDay     'ggoOper.RetFormat(frm1.txtFr_dt.Text,   "yyyyMMdd")
	to_dt   = strYear2&strMonth2&strDay2     'ggoOper.RetFormat(frm1.txtTo_dt.Text,   "yyyyMMdd")
	rprt_dt = strYear3&strMonth3&strDay3     'ggoOper.RetFormat(frm1.txtrprt_dt.Text, "yyyyMMdd")
	
	fr_dept_cd = strFrDept
	to_dept_cd = strToDept
	wrt_prsn   = frm1.txtwrt_prsn.value
	emp_no     = frm1.txtemp_no.value
    
	If wrt_prsn = "" Then
		wrt_prsn = "%"
		frm1.txtwrt_prsn.value = ""
	End If	
	
	If emp_no = "" Then
		emp_no = "%"
		frm1.txtName.value = ""
	End If	

    StrEbrFile =  iValue
    lngPos = 0
        		
	For intCnt = 1 To 3
	    lngPos = InStr(lngPos + 1, GetUserPath, "/")
	Next
		
	condvar = "Sect_cd|" & strSect_cd
	condvar = condvar & "|fr_dt|" & fr_dt	
	condvar = condvar & "|to_dt|" & to_dt	
	condvar = condvar & "|rprt_dt|" & rprt_dt	
	condvar = condvar & "|emp_no|" & emp_no
	condvar = condvar & "|fr_dept_cd|" & fr_dept_cd
	condvar = condvar & "|to_dept_cd|" & to_dept_cd

    If Trim(iValue) = "h5502oa1_1" Then
        ObjName = AskEBDocumentName("h5502oa1_1", "ebr")
	Elseif  Trim(iValue) = "h5502oa1_2" Then  
        ObjName = AskEBDocumentName("h5502oa1_2", "ebr")
	Elseif	Trim(iValue) = "h5502oa1_3" Then  
        ObjName = AskEBDocumentName("h5502oa1_3", "ebr")
	End If

    Call FncEBRPrint(EBAction, ObjName, condvar)
	
End Function

'========================================================================================
' Function Name : FncBtnPreview()
' Function Desc : This function is related to Preview Button
'========================================================================================

Function btnRun_Preview()
    Dim strFrDt
    Dim strToDt
    Dim strFrDept
    Dim strToDept    
    Dim strMin
    Dim strMax
   
        
    If Not chkField(Document, "1") Then									'⊙: This function check indispensable field%>
       Exit Function
    End If
    If txtEmp_no_Onchange() Then        'enter key 로 조회시 사원을 check후 해당사항 없으면 query종료...
        Exit Function
    End if
    
    If txtFr_dept_cd_Onchange() Then        'enter key 로 조회시 부서코드를 check후 해당사항 없으면 query종료...
        Exit Function
    End if
    
    If txtTo_dept_cd_Onchange() Then        'enter key 로 조회시 부서코드를 check후 해당사항 없으면 query종료...
        Exit Function
    End if

    If txtSect_cd_Onchange() Then        'enter key 로 조회시 부서코드를 check후 해당사항 없으면 query종료...
        Exit Function
    End if    
    strFrDt = Trim(frm1.txtFr_dt.Text)
    strToDt = Trim(frm1.txtTo_dt.Text)
    
    If strFrDt = "" AND strToDt ="" Then       
    Else
        If strFrDt = "" then
            strFrDt =  "1950" &  Parent.gComDateType & "01" & Parent.gComDateType  & "01"
        End if

        If strToDt = "" then
            strToDt = "2100" &  Parent.gComDateType & "12" & Parent.gComDateType  & "31"
        ElseIf  ValidDateCheck(frm1.txtFr_dt,frm1.txtTo_dt)= False Then  'strFrDt > strToDt then
            Exit Function
        End if 
    End if 
    
    Call FuncGetTermDept(lgUsrIntCd,"",strMin,strMax)
    
    strFrDept = Trim(frm1.txtFr_internal_cd.value)
    strToDept = Trim(frm1.txtTo_internal_cd.value)

    If strFrDept = "" Then       
       strFrDept = strMin
    End if

    If strToDept = "" then
       strToDept = strMax
    End If
    
    If strFrDept > strToDept then
	    Call DisplayMsgbox("970025","X","시작부서","종료부서")	'시작부서는 종료부서보다 작아야 합니다.
        frm1.txtFr_dept_cd.focus
        Set gActiveElement = document.activeElement
        Exit Function
    End if 
    
    with frm1
	    If .txtcap_type(0).Checked Then      
	        Call FncBtnPreview("h5502oa1_1")
	    Else
	        Call FncBtnPreview("h5502oa1_2")
	        Call FncBtnPreview("h5502oa1_3")
	    End If
	End with
End Function

'========================================================================================
' Function Name : FncBtnPreview()
' Function Desc : This function is related to Preview Button
'========================================================================================

Function FncBtnPreview(iValue)
	Dim condvar
	Dim lngPos
	Dim intCnt
    Dim StrEbrFile
    Dim strFrDt
    Dim strToDt
    Dim strFrDept
    Dim strToDept
    Dim strMin
    Dim strMax
    Dim strSect_cd,rel_cd1, rel_cd2
    Dim ls_work_area,ls_comp_na,ls_tel_no,ls_repre_nm,ls_comp_addr
    Dim IntRetCD
	Dim arrParam, arrField, arrHeader
	Dim strYear,strMonth,strDay
	Dim strYear2,strMonth2,strDay2
	Dim strYear3,strMonth3,strDay3	
    Dim ls_comp_no1,ls_comp_ser,ls_saupjang1,ls_repre_nm1,ls_juso1,ls_cust_telno1
	Dim fr_dt, to_dt, fr_dept_cd, to_dept_cd, rprt_dt, wrt_prsn, emp_no 
	Dim ObjName
     If txtEmp_no_Onchange() Then        'enter key 로 조회시 사원을 check후 해당사항 없으면 query종료...
        Exit Function
    End if
    
    If txtFr_dept_cd_Onchange() Then        'enter key 로 조회시 부서코드를 check후 해당사항 없으면 query종료...
        Exit Function
    End if
    
    If txtTo_dept_cd_Onchange() Then        'enter key 로 조회시 부서코드를 check후 해당사항 없으면 query종료...
        Exit Function
    End if

    If txtSect_cd_Onchange() Then        'enter key 로 조회시 부서코드를 check후 해당사항 없으면 query종료...
        Exit Function
    End if  
    Call FuncGetTermDept(lgUsrIntCd,"",strMin,strMax)
    
    strFrDt = Trim(frm1.txtFr_dt.Text)
    strToDt = Trim(frm1.txtTo_dt.Text)
    
    Call FuncGetTermDept(lgUsrIntCd,"",strMin,strMax)
    
    strFrDept = Trim(frm1.txtFr_internal_cd.value)
    strToDept = Trim(frm1.txtTo_internal_cd.value)
    If strFrDept = "" Then       
       strFrDept = strMin
    End if
    If strToDept = "" then
       strToDept = strMax
    End If
        
    strSect_cd = frm1.txtSect_cd.value
    rel_cd1 = strSect_cd
    
    IntRetCD = CommonQueryRs(" reference "," B_CONFIGURATION "," MAJOR_CD=" & FilterVar("H0035", "''", "S") & " and MINOR_CD= " & FilterVar(strSect_cd, "''", "S") & " and seq_no=1 ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
 
    If IntRetCD=False  Then
       Call DisplayMsgBox("800490","X","X","X")                             ''해당 자료가 없습니다.
       frm1.txtSect_cd.focus()
       Exit Function
    Else
       rel_cd2=Trim(Replace(lgF0,Chr(11),""))
       IntRetCD = CommonQueryRs(" biz_area_nm,tel_no,repre_nm,addr "," b_biz_area "," biz_area_cd= " & FilterVar(rel_cd2, "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
                
       If IntRetCD=False  Then
           Call DisplayMsgBox("800491","X","X","X")                         ''해당 자료가 없습니다.
           frm1.txtSect_cd.focus()
           Exit Function
       Else
           ls_comp_na   =Trim(Replace(lgF0,Chr(11),"")) 
           ls_tel_no    =Trim(Replace(lgF1,Chr(11),""))
           ls_repre_nm  =Trim(Replace(lgF2,Chr(11),""))
           ls_comp_addr =Trim(Replace(lgF3,Chr(11),""))
       End If
    End If
 
    IntRetCD = CommonQueryRs(" MINOR_NM "," B_MINOR "," MAJOR_CD=" & FilterVar("H0069", "''", "S") & " and MINOR_CD='" & strSect_cd & "1'",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
 
    If IntRetCD=True  Then
        ls_comp_no1   =Trim(Replace(lgF0,Chr(11),"")) 
    Else
        ls_comp_no1   = "%"
    End If

              IntRetCD = CommonQueryRs(" MINOR_NM "," B_MINOR "," MAJOR_CD=" & FilterVar("H0069", "''", "S") & " and MINOR_CD='" & strSect_cd & "2'",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
           If IntRetCD=True  Then
                    ls_comp_ser   =Trim(Replace(lgF0,Chr(11),"")) 
           Else
                   ls_comp_ser   = "%"
           End If

              IntRetCD = CommonQueryRs(" MINOR_NM "," B_MINOR "," MAJOR_CD=" & FilterVar("H0069", "''", "S") & " and MINOR_CD='" & strSect_cd & "3'",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
           If IntRetCD=True  Then
                    ls_saupjang1   =Trim(Replace(lgF0,Chr(11),"")) 
           Else
                    ls_saupjang1   = "%"
           End If

              IntRetCD = CommonQueryRs(" MINOR_NM "," B_MINOR "," MAJOR_CD=" & FilterVar("H0069", "''", "S") & " and MINOR_CD='" & strSect_cd & "4'",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
           If IntRetCD=True  Then
                    ls_repre_nm1   =Trim(Replace(lgF0,Chr(11),"")) 
           Else
                    ls_repre_nm1   = "%"
           End If

              IntRetCD = CommonQueryRs(" MINOR_NM "," B_MINOR "," MAJOR_CD=" & FilterVar("H0069", "''", "S") & " and MINOR_CD='" & strSect_cd & "5'",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
           If IntRetCD=True  Then
                    ls_juso1   =Trim(Replace(lgF0,Chr(11),"")) 
           Else
                    ls_juso1   = "%"
           End If
		
              IntRetCD = CommonQueryRs(" MINOR_NM "," B_MINOR "," MAJOR_CD=" & FilterVar("H0069", "''", "S") & " and MINOR_CD='" & strSect_cd & "6'",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
           If IntRetCD=True  Then
                    ls_cust_telno1   =Trim(Replace(lgF0,Chr(11),"")) 
           Else
                    ls_cust_telno1   = "%"
           End If
	
	Call ExtractDateFrom(frm1.txtFr_dt.text,frm1.txtFr_dt.UserDefinedFormat,Parent.gComDateType,strYear,strMonth,strDay)
	Call ExtractDateFrom(frm1.txtTo_dt.text,frm1.txtTo_dt.UserDefinedFormat,Parent.gComDateType,strYear2,strMonth2,strDay2)
	Call ExtractDateFrom(frm1.txtrprt_dt.text,frm1.txtrprt_dt.UserDefinedFormat,Parent.gComDateType,strYear3,strMonth3,strDay3)
	
	fr_dt   = strYear&strMonth&strDay 		'ggoOper.RetFormat(frm1.txtFr_dt.Text,   "yyyyMMdd")
	to_dt   = strYear2&strMonth2&strDay2 	'ggoOper.RetFormat(frm1.txtTo_dt.Text,   "yyyyMMdd")
	rprt_dt = strYear3&strMonth3&strDay3 	'ggoOper.RetFormat(frm1.txtrprt_dt.Text, "yyyyMMdd")
	
	fr_dept_cd = strFrDept
	to_dept_cd = strToDept
	wrt_prsn   = frm1.txtwrt_prsn.value
	emp_no     = frm1.txtemp_no.value

	If wrt_prsn = "" Then
		wrt_prsn = "%"
		frm1.txtwrt_prsn.value = ""
	End If	

	If emp_no = "" Then
		emp_no = "%"
		frm1.txtName.value = ""
	End If	

    StrEbrFile = iValue 

	condvar = "emp_no|" & emp_no
	condvar = condvar & "|Sect_cd|" & strSect_cd
	condvar = condvar & "|fr_dt|" & fr_dt	
	condvar = condvar & "|to_dt|" & to_dt	
	condvar = condvar & "|rprt_dt|" & rprt_dt		
	condvar = condvar & "|fr_dept_cd|" & fr_dept_cd
	condvar = condvar & "|to_dept_cd|" & to_dept_cd

    ObjName = AskEBDocumentName(StrEbrFile, "ebr")

	call FncEBRPreview(ObjName , condvar)

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
	    	arrParam(1) = "B_BIZ_AREA"							    ' TABLE 명칭 
	    	arrParam(2) = frm1.txtCust_cd.value     			' Code Condition
	    	arrParam(3) = ""'frm1.txtCust_nm.value									' Name Cindition
	    	arrParam(4) = ""	               	                ' Where Condition
	    	arrParam(5) = "신고사업장코드"  		        ' TextBox 명칭 
	
	    	arrField(0) = "BIZ_AREA_CD"						   	' Field명(0)
	    	arrField(1) = "BIZ_AREA_NM"    				  		' Field명(1)
	    	arrField(2) = ""    				        		' Field명(2)
    
	    	arrHeader(0) = "신고사업장코드"	     			' Header명(0)
	    	arrHeader(1) = "신고사업장코드명"		        ' Header명(1)
	    	arrHeader(2) = ""   	    						' Header명(1)
	End Select   
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	
	If arrRet(0) = "" Then
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
		    Case "SECT_CD_POP"
		        .txtSect_cd.value = arrRet(0) 
		    	.txtSect_nm.value = arrRet(1) 
		    Case "CUST_CD_POP"
		        .txtCust_cd.value = arrRet(0) 
		    	.txtCust_nm.value = arrRet(1) 
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
	    arrParam(1) = ""'frm1.txtName.value			' Name Cindition
	    arrParam(2) = lgUsrIntCd                    ' 자료권한 Condition  
    Else
	    arrParam(0) = frm1.txtEmp_no.value			' Code Condition
	    arrParam(1) = ""'frm1.txtName.value			' Name Cindition
	    arrParam(2) = lgUsrIntCd                    ' 자료권한 Condition  
	End If

	arrRet = window.showModalDialog(HRAskPRAspName("EmpPopup"), Array(window.parent, arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False
	
	If arrRet(0) = "" Then
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

		lgBlnFlgChgValue = False
	End With
End Sub

'========================================================================================================
'   Event Name : txtEmp_no_change             '<==인사마스터에 있는 사원인지 확인 
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
        IntRetCd = FuncDeptName(frm1.txtFr_dept_cd.value,"",lgUsrIntCd,strDept_nm,lsInternal_cd)
        
        if  IntRetCd < 0 then
            if  IntRetCd = -1 then
                Call DisplayMsgBox("800012", "x","x","x")   '부서코드정보에 등록되지 않은 코드입니다.
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
        IntRetCd = FuncDeptName(frm1.txtTo_dept_cd.value,"",lgUsrIntCd,strDept_nm,lsInternal_cd)
        
        if  IntRetCd < 0 then
            if  IntRetCd = -1 then
                Call DisplayMsgBox("800012", "x","x","x")   '부서코드정보에 등록되지 않은 코드입니다.
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
        else
            frm1.txtTo_dept_nm.value = strDept_nm
            frm1.txtTo_internal_cd.value = lsInternal_cd
        end if
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
			Call DisplayMsgbox("800054","X","X","X")	
			 frm1.txtSect_nm.value = ""
             frm1.txtSect_cd.focus
            Set gActiveElement = document.ActiveElement       
            txtSect_cd_Onchange = true
            Exit Function      
        Else
			frm1.txtSect_nm.value = Trim(Replace(lgF0,Chr(11),""))
        End if 
    End if  
    
End Function 


'========================================================================================================
' Name : txtFr_dt_DblClick
' Desc : 달력 Popup을 호출 
'========================================================================================================
Sub txtFr_dt_DblClick(Button)
	If Button = 1 Then
		Call SetFocusToDocument("M")	
		frm1.txtFr_dt.Action = 7
		frm1.txtFr_dt.focus
	End If
End Sub
'========================================================================================================
' Name : txtFr_dt_DblClick
' Desc : 달력 Popup을 호출 
'========================================================================================================
Sub txtTo_dt_DblClick(Button)
	If Button = 1 Then
		Call SetFocusToDocument("M")	
		frm1.txtTo_dt.Action = 7
		frm1.txtTo_dt.focus
	End If
End Sub
'========================================================================================================
' Name : txtrprt_dt_DblClick
' Desc : 달력 Popup을 호출 
'========================================================================================================
Sub txtrprt_dt_DblClick(Button)
	If Button = 1 Then
		Call SetFocusToDocument("M")	
		frm1.txtrprt_dt.Action = 7
		frm1.txtrprt_dt.focus
	End If
End Sub
'=======================================================================================================
'   Event Name : txtFr_dt_Keypress(Key)
'   Event Desc : enter key down시에 조회한다.
'=======================================================================================================
Sub txtFr_dt_Keypress(Key)
    If Key = 13 Then
        FncQuery()
    End If
End Sub
'=======================================================================================================
'   Event Name : txtTo_dt_Keypress(Key)
'   Event Desc : enter key down시에 조회한다.
'=======================================================================================================
Sub txtTo_dt_Keypress(Key)
    If Key = 13 Then
        FncQuery()
    End If
End Sub
'=======================================================================================================
'   Event Name : txtrprt_dt_Keypress(Key)
'   Event Desc : enter key down시에 조회한다.
'=======================================================================================================
Sub txtrprt_dt_Keypress(Key)
    If Key = 13 Then
        FncQuery()
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>고용보험취득/상실신고</font></td>
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
							    	<TD CLASS="TD5" NOWRAP>자격구분</TD>
				        	        <TD CLASS="TD6">
				        	            <INPUT TYPE="RADIO" CLASS="RADIO" NAME="txtcap_type" TAG="11" VALUE="자격취득" CHECKED ID="med_entr_flag1"><LABEL FOR="txtcap_type1">자격취득</LABEL>&nbsp;
				        	            <INPUT TYPE="RADIO" CLASS="RADIO" NAME="txtcap_type" TAG="11" VALUE="자격상실" ID="med_entr_flag2"><LABEL FOR="txtcap_type2">자격상실</LABEL>
 				        	        </TD>
								</TR>
							    <TR>
							    	<TD CLASS=TD5 NOWRAP>근무구역</TD>
									<TD CLASS=TD6 NOWRAP>
								        <INPUT ID="txtSect_cd" NAME="txtSect_cd" ALT="근무구역" TYPE="Text" SiZE=10 MAXLENGTH=2  tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSItemDC" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenCode frm1.txtSect_cd.value,'SECT_CD_POP'">
								        <INPUT NAME="txtSect_nm" ALT="근무구역명" TYPE="Text" SiZE=20 MAXLENGTH=50  tag="14XXXU">
									</TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>해당일</TD>
									<TD CLASS=TD6 NOWRAP>
										<script language =javascript src='./js/h5502oa1_fpDateTime1_txtFr_dt.js'></script>&nbsp;~&nbsp;
										<script language =javascript src='./js/h5502oa1_fpDateTime1_txtTo_dt.js'></script>
								    </TD>
								</TR>
								<TR>	
									<TD CLASS=TD5 NOWRAP>작성자</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtwrt_prsn" ALT="작성자" TYPE="Text" MAXLENGTH="10" SIZE=13 tag="11XXXU"></TD>
								</TR>
								<TR>	
									<TD CLASS=TD5 NOWRAP>신고일</TD>
									<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/h5502oa1_rprt_dt_txtrprt_dt.js'></script></TD>
								</TR>
								<TR>
								    <TD CLASS="TD5" NOWRAP>대상자</TD>
								    <TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" ID = "txtEmp_no" NAME="txtEmp_no"  SIZE=13 MAXLENGTH=13 tag="11XXXU" ALT="대상자"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnAcctCdFR" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenEmptName (0)">
								                           <INPUT TYPE="Text" NAME="txtName" SIZE=20 MAXLENGTH=30  tag="14XXXU" ALT="대상자코드명"></TD>
							    </TR>
								
							    <TR>
							    	<TD CLASS="TD5" NOWRAP>부서코드</TD>
							    	<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" ID = "txtFr_dept_cd" NAME="txtFr_dept_cd" SIZE=10 MAXLENGTH=10  tag="11XXXU" ALT="시작부서코드"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnAcctCdFR" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenDept(0)">
							    	                       <INPUT TYPE="Text" NAME="txtFr_dept_nm" SIZE=20 MAXLENGTH=40 tag="14XXXU" ALT="부서코드명">&nbsp;~&nbsp;
							    	                       <INPUT NAME="txtFr_internal_cd" ALT="내부부서코드" TYPE="hidden" SiZE=7 MAXLENGTH=30  tag="14XXXU"></TD>
							    </TR>			
							    <TR>
							    	<TD CLASS="TD5" NOWRAP></TD>
							    	<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" ID = "txtTo_dept_cd" NAME="txtTo_dept_cd" SIZE=10 MAXLENGTH=10  tag="11XXXU" ALT="종료부서코드"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnAcctCdFR" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenDept(1)">
							    	                       <INPUT TYPE="Text" NAME="txtTo_dept_nm" SIZE=20 MAXLENGTH=40  tag="14XXXU" ALT="부서코드명">&nbsp;
							    	                       <INPUT NAME="txtTo_internal_cd" ALT="내부부서코드" TYPE="hidden" SiZE=7 MAXLENGTH=30  tag="14XXXU"></TD>
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
                         <BUTTON NAME="btnRun"   CLASS="CLSSBTN" ONCLICK="vbscript:btnRun_Preview()" Flag=1>미리보기</BUTTON>&nbsp;
                         <BUTTON NAME="btnPrint" CLASS="CLSSBTN" ONCLICK="vbscript:btnPrint_Print()" Flag=1>인쇄</BUTTON>
		            </TD>
					<TD WIDTH=* ALIGN="right">&nbsp;</TD>
					<TD WIDTH=10>&nbsp;</TD>
		        </TR>
		    </TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		                                   <IFRAME NAME="MyBizASP1" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
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


