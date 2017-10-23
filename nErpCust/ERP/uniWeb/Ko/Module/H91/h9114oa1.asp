<%@ LANGUAGE="VBSCRIPT" %>
<!--
'======================================================================================================
'*  1. Module Name          : 인사/급여관리 
'*  2. Function Name        : 인사기본자료관리 
'*  3. Program ID           : h9114oa1
'*  4. Program Name         : 근로소둑영수증출력 
'*  5. Program Desc         : 근로소둑영수증출력 
'*  6. Comproxy List        : +
'*  7. Modified date(First) : 2001/05/27
'*  8. Modified date(Last)  : 2003/06/13
'*  9. Modifier (First)     : Shin Kwang-Ho
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

<SCRIPT LANGUAGE="VBsCRIPT"   SRC="../../inc/incEB.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incHRQuery.vbs"> </SCRIPT>

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
Dim lgOldRow
Dim PrintNum

<% EndDate   = GetSvrDate %>
'========================================================================================================
' Name : InitVariables()	
' Desc : Initialize value
'========================================================================================================
Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    PrintNum = 0
        
End Sub
'========================================================================================================
' Name : SetDefaultVal()	
' Desc : Set default value
'========================================================================================================
Sub SetDefaultVal()
	frm1.txtbas_yy.focus
	frm1.txtbas_yy.text = Year(Date)
	frm1.txtprov_dt.text = UniConvDateAToB("<%=EndDate%>", parent.gServerDateFormat, parent.gDateFormat)
End Sub
'========================================================================================================
' Name : InitComboBox()
' Desc : Set ComboBox
'========================================================================================================
Sub InitComboBox()
    Dim iCodeArr 
    Dim iNameArr
    Dim iDx
    
    '신고 사업장    
    Call CommonQueryRs("YEAR_AREA_NM,YEAR_AREA_CD","HFA100T","",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    iNameArr = lgF0
    iCodeArr = lgF1   
    Call SetCombo2(frm1.txtcust_cd,iCodeArr,iNameArr,Chr(11))     
    
    ' 직종 
    Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0003", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    iCodeArr =  lgF0
    iNameArr =  lgF1
    Call SetCombo2(frm1.cboOcpt_type, iCodeArr, iNameArr, Chr(11))
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
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
	Call FuncGetAuth(gStrRequestMenuID , parent.gUsrID, lgUsrIntCd)     ' 자료권한:lgUsrIntCd ("%", "1%")
    Call InitVariables
    Call InitComboBox 
    Call ggoOper.FormatDate(frm1.txtbas_yy, parent.gDateFormat,3) 
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
       
    Err.Clear                                                                    '☜: Clear err status
    If txtEmp_no_Onchange() Then        'enter key 로 조회시 사원을 check후 해당사항 없으면 query종료...
        Exit Function
    End if
    
    If txtFr_dept_cd_Onchange() Then        'enter key 로 조회시 부서코드를 check후 해당사항 없으면 query종료...
        Exit Function
    End if
    
    If txtTo_dept_cd_Onchange() Then        'enter key 로 조회시 부서코드를 check후 해당사항 없으면 query종료...
        Exit Function
    End if   

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
' Name : OpenEmpName()
' Desc : developer describe this line(grid외에서 사용) 
'========================================================================================================
Function OpenEmp()
	Dim arrRet
	Dim arrParam(2)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = frm1.txtEmp_no.value			' Code Condition
	arrParam(1) = ""'frm1.txtName.value			' Name Cindition
	arrParam(2) = lgUsrIntCd        			' Internal_cd
	
	arrRet = window.showModalDialog(HRAskPRAspName("EmpPopup"), Array(window.parent,arrParam), _
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
    Dim strBasDtAdd
    
	strBasDt = frm1.txtbas_yy.Text & parent.gComDateType & "12" & parent.gComDateType & "31"		
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	If iWhere = 0 Then
		arrParam(0) = frm1.txtFr_dept_cd.value			            '  Code Condition
	ElseIf iWhere = 1 Then
		arrParam(0) = frm1.txtTo_dept_cd.value			            ' Code Condition
	End If
	
    arrParam(1) = UNIConvDate(strBasDt)
    arrParam(2) = lgUsrIntcd
	
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
    Dim strVal
    
    If frm1.txtEmp_no.value = "" Then
		frm1.txtName.value = ""
    Else
        IntRetCd = FuncGetEmpInf2(frm1.txtEmp_no.value,lgUsrIntCd,strName,strDept_nm,_
	                strRoll_pstn, strPay_grd1, strPay_grd2, strEntr_dt, strInternal_cd)    
    
        if  IntRetCd < 0 then
	        if  IntRetCd = -1 then
    			Call DisplayMsgBox("800048","X","X","X")	'해당사원은 존재하지 않습니다.
            else
                Call DisplayMsgBox("800454","X","X","X")	'자료에 대한 권한이 없습니다.
            end if
			frm1.txtName.value = ""
            Call ggoOper.ClearField(Document, "2")
            call InitVariables()
            frm1.txtEmp_no.focus
            Set gActiveElement = document.ActiveElement
            txtEmp_no_Onchange = true
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
    Dim IntRetCd,Dept_Nm,Internal_cd
    Dim strBasDt 
   
	strBasDt = frm1.txtbas_yy.Text & parent.gComDateType & "12" & parent.gComDateType & "31"
	    
    If frm1.txtFr_dept_cd.value = "" Then
		frm1.txtFr_dept_nm.value = ""
		frm1.txtFr_internal_cd.value = ""        
    Else
        IntRetCd = FuncDeptName(frm1.txtFr_dept_cd.value ,strBasDt  , lgUsrIntCd,Dept_Nm , Internal_cd)
	    If  IntRetCd < 0 then
	        If  IntRetCd = -1 then
    			Call DisplayMsgBox("800098","X","X","X")	'부서정보코드에 등록되지 않은 코드입니다.
            Else
                Call DisplayMsgBox("800454","X","X","X")	'자료에 대한 권한이 없습니다.
            End if
		    frm1.txtFr_dept_nm.value = ""
		    frm1.txtFr_internal_cd.value = ""
            frm1.txtFr_dept_cd.focus
            Set gActiveElement = document.ActiveElement             
            txtFr_dept_cd_Onchange = True
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
    Dim strBasDt 
   
	strBasDt = frm1.txtbas_yy.Text & parent.gComDateType & "12" & parent.gComDateType & "31"
	    
    If frm1.txtTo_dept_cd.value = "" Then
		frm1.txtTo_dept_nm.value = ""
		frm1.txtTo_internal_cd.value = ""
    Else
        IntRetCd = FuncDeptName(frm1.txtTo_dept_cd.value , strBasDt , lgUsrIntCd,Dept_Nm , Internal_cd)
	    If  IntRetCd < 0 then
	        If  IntRetCd = -1 then
    			Call DisplayMsgBox("800098","X","X","X")	'부서정보코드에 등록되지 않은 코드입니다.
            Else
                Call DisplayMsgBox("800454","X","X","X")	'자료에 대한 권한이 없습니다.
            End if
		    frm1.txtTo_dept_nm.value = ""
		    frm1.txtTo_internal_cd.value = ""
            frm1.txtTo_dept_cd.focus
            Set gActiveElement = document.ActiveElement 
            txtTo_dept_cd_Onchange = True
            Exit Function
        Else
			frm1.txtTo_dept_nm.value = Dept_Nm
		    frm1.txtTo_internal_cd.value = Internal_cd
        End if 
    End if  
    
End Function

'========================================================================================
' Function Name : FncBtnPrint()
' Function Desc : This function is related to Preview Button
'========================================================================================
Function FncBtnPrint()
	Dim StrEbrFile, ObjName
	Dim strUrl, strNext
	dim stryyyy
	dim gbn

	
	stryyyy = frm1.txtbas_yy.year
	call FncSub(strUrl,strNext)
	if stryyyy="2005" then
    else
		stryyyy="2006"
    end if
    
    With frm1
		gbn="1" 
		If .retire_check0.checked = True Then	
			
			If .prt_check1_flag.checked = True  Then
			    StrEbrFile = "h9114oa1_3" & stryyyy
				strUrl = strUrl & "|prt_check_flag|1"					    
					
   				ObjName = AskEBDocumentName(StrEbrFile,"ebr")
 				call FncEBRPrint(EBAction , ObjName , strUrl)
			End If			

			If .prt_check2_flag.checked = True  Then
				strUrl = strUrl & "|prt_check_flag|2"
			    StrEbrFile = "h9114oa1_3" & stryyyy
   				ObjName = AskEBDocumentName(StrEbrFile,"ebr")
 				call FncEBRPrint(EBAction , ObjName , strUrl)
			End If

			If .prt_check3_flag.checked = True  Then
				strUrl = strUrl & "|prt_check_flag|3"
			    StrEbrFile = "h9114oa1_3" & stryyyy
   				ObjName = AskEBDocumentName(StrEbrFile,"ebr")
 				call FncEBRPrint(EBAction , ObjName , strUrl)
			End If
											
			If .prt_check4_flag.checked = True  Then
				strUrl = strUrl & "|prt_check_flag|4"
			    StrEbrFile = "h9114oa1_4" & stryyyy
   				ObjName = AskEBDocumentName(StrEbrFile,"ebr")
 				call FncEBRPrint(EBAction , ObjName , strUrl)
			End If
		Else
			gbn = "2"
			If .prt_check1_flag.checked = True  Then
				strUrl = strUrl & "|prt_check_flag|1"
			    StrEbrFile = "h9114oa1_1" & stryyyy
   				ObjName = AskEBDocumentName(StrEbrFile,"ebr")
 				call FncEBRPrint(EBAction , ObjName , strUrl)
			End If			

			If .prt_check2_flag.checked = True  Then
				strUrl = strUrl & "|prt_check_flag|2"
			    StrEbrFile = "h9114oa1_1" & stryyyy
   				ObjName = AskEBDocumentName(StrEbrFile,"ebr")
 				call FncEBRPrint(EBAction , ObjName , strUrl)
			End If

			If .prt_check3_flag.checked = True  Then
				strUrl = strUrl & "|prt_check_flag|3"
			    StrEbrFile = "h9114oa1_1" & stryyyy
   				ObjName = AskEBDocumentName(StrEbrFile,"ebr")
 				call FncEBRPrint(EBAction , ObjName , strUrl)
			End If
											
			If .prt_check4_flag.checked = True  Then
				strUrl = strUrl & "|prt_check_flag|4"
			    StrEbrFile = "h9114oa1_2" & stryyyy
   				ObjName = AskEBDocumentName(StrEbrFile,"ebr")
 				call FncEBRPrint(EBAction , ObjName , strUrl)
			End If	
					
		End If
		
		 if .prt_check1_flag.checked or .prt_check2_flag.checked or .prt_check3_flag.checked then
		
			if stryyyy="2006" then
			
				strUrl = strUrl & "|pg|1"
				strUrl = strUrl & "|gbn|"&gbn
			    StrEbrFile = "h9114oa1_12006p2"
   				ObjName = AskEBDocumentName(StrEbrFile,"ebr")
 				call FncEBRPrint(EBAction , ObjName , strUrl)
 			end if	
 			
 		end if	

    End With

End Function

'========================================================================================
' Function Name : FncBtnPreview()
' Function Desc : This function is related to Preview Button
'========================================================================================
Function FncBtnPreview()
	Dim StrEbrFile, ObjName
	Dim strUrl, strNext
	dim stryyyy,gbn

	
	stryyyy = frm1.txtbas_yy.year
	call FncSub(strUrl,strNext)
	
    if stryyyy="2005" then
    else
		stryyyy="2006"
    end if
    
    
    With frm1 
		If .retire_check0.checked = True Then	
			gbn="1"
			If .prt_check1_flag.checked = True  Then
			    StrEbrFile = "h9114oa1_3" & stryyyy
				strUrl = strUrl & "|prt_check_flag|1"					    
					
   				ObjName = AskEBDocumentName(StrEbrFile,"ebr")
 				call FncEBRPreview(ObjName , strUrl)
			End If			

			If .prt_check2_flag.checked = True  Then
				strUrl = strUrl & "|prt_check_flag|2"
			    StrEbrFile = "h9114oa1_3" & stryyyy
   				ObjName = AskEBDocumentName(StrEbrFile,"ebr")
 				call FncEBRPreview(ObjName , strUrl)
			End If

			If .prt_check3_flag.checked = True  Then
				strUrl = strUrl & "|prt_check_flag|3"
			    StrEbrFile = "h9114oa1_3" & stryyyy
   				ObjName = AskEBDocumentName(StrEbrFile,"ebr")
 				call FncEBRPreview(ObjName , strUrl)
			End If
											
			If .prt_check4_flag.checked = True  Then
				strUrl = strUrl & "|prt_check_flag|4"
			    StrEbrFile = "h9114oa1_4" & stryyyy
   				ObjName = AskEBDocumentName(StrEbrFile,"ebr")
 				call FncEBRPreview(ObjName , strUrl)
			End If
			
			
 				
		Else
		gbn="2"
			If .prt_check1_flag.checked = True  Then
				strUrl = strUrl & "|prt_check_flag|1"
			    StrEbrFile = "h9114oa1_1" & stryyyy
   				ObjName = AskEBDocumentName(StrEbrFile,"ebr")
 				call FncEBRPreview(ObjName , strUrl)
			End If			

			If .prt_check2_flag.checked = True  Then
				strUrl = strUrl & "|prt_check_flag|2"
			    StrEbrFile = "h9114oa1_1" & stryyyy
   				ObjName = AskEBDocumentName(StrEbrFile,"ebr")
 				call FncEBRPreview(ObjName , strUrl)
			End If

			If .prt_check3_flag.checked = True  Then
				strUrl = strUrl & "|prt_check_flag|3"
			    StrEbrFile = "h9114oa1_12006"
   				ObjName = AskEBDocumentName(StrEbrFile,"ebr")
 				call FncEBRPreview(ObjName , strUrl)
			End If
											
			If .prt_check4_flag.checked = True  Then
				strUrl = strUrl & "|prt_check_flag|4"
			    StrEbrFile = "h9114oa1_2" & stryyyy
   				ObjName = AskEBDocumentName(StrEbrFile,"ebr")
 				call FncEBRPreview(ObjName , strUrl)
			End If	
			
		
		End If

         if .prt_check1_flag.checked or .prt_check2_flag.checked or .prt_check3_flag.checked then
			if stryyyy="2006" then
			
				strUrl = strUrl & "|pg|1"
				strUrl = strUrl & "|gbn|"&gbn
			    StrEbrFile = "h9114oa1_12006p2"
   				ObjName = AskEBDocumentName(StrEbrFile,"ebr")
   				
 				call FncEBRPreview(ObjName , strUrl)
 			end if	
 		end if	
 			
    End With

End Function

'========================================================================================
' Function Name : FncBtnPreview(StrEbrFile,strUrl)
'========================================================================================
Function FncSub(strUrl,strNext)
	Dim txtretire_check, prt_check_flag
    Dim strYear,strMonth,strDay 	
	Dim bas_dt, bas_yy, biz_area_cd, emp_no, fr_dept_cd, ocpt_type, prov_dt, to_dept_cd
	Dim std_sub
	
    If Not chkField(Document, "1") Then									'⊙: This function check indispensable field
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
	Dim strFrDept ,strToDept    

    With frm1 	
		strFrDept = Trim(.txtFr_internal_cd.value)
		strToDept = Trim(.txtTo_internal_cd.value)
		If strFrDept = "" AND strToDept = "" Then       
		Else
	
		    If strFrDept = "" then
		        .txtFr_internal_cd.value = "00000000000"
		    End if
		    If strToDept = "" then
		        .txtTo_internal_cd.value = "zzzzzzzzzz"
		    ElseIf strFrDept > strToDept then
		        Call DisplayMsgBox("800153","X","X","X")	'시작부서는 종료부서보다 작아야 합니다.
		        .txtFr_dept_cd.focus
		        Set gActiveElement = document.activeElement
		        Exit Function
		    End if 
		End if 

   '----------------------------------------------------------------------------------------------    
	    bas_dt = .txtbas_yy.text & "1231"	    
	    bas_yy = .txtbas_yy.text
	    biz_area_cd = .txtcust_cd.value 
	    emp_no = .txtEmp_no.value 
	    ocpt_type = .cboOcpt_type.value 
	    prov_dt = .txtprov_dt.Year & Right("0" & .txtprov_dt.Month, 2) & Right("0" & .txtprov_dt.Day, 2) 

		if Trim(.txtFr_internal_cd.value ) = "" Then
			fr_dept_cd = "0"
		else
			fr_dept_cd = .txtFr_internal_cd.value 
		end if
		
		if Trim(.txtTo_internal_cd.value  ) = "" Then
			to_dept_cd = "ZZZZ"
		else
			 to_dept_cd = .txtTo_internal_cd.value  
		end if		

	    if emp_no = "" then
	    	emp_no = "%"
	    	.txtName.value = ""
	    End if	

	    if ocpt_type = "" then
	    	ocpt_type = "%"
	    	.cboOcpt_type.value = ""
	    End if	

    	    
	    if prov_dt = "" then
	        Call ExtractDateFrom("<%=GetsvrDate%>",parent.gServerDateFormat , parent.gServerDateType ,strYear,strMonth,strDay)
	    	prov_dt = strYear & Right("0" & strMonth,2) & Right("0" & strDay,2)
	    End if  

		Call CommonQueryRs(" std_sub "," hfa020t ", " comp_cd='1' ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

		std_sub = Trim(Replace(lgF0,Chr(11),""))
		
		strUrl = "bas_dt|" & bas_dt
		strUrl = strUrl & "|bas_yy|" & bas_yy 
		strUrl = strUrl & "|biz_area_cd|" & biz_area_cd
		strUrl = strUrl & "|emp_no|" & emp_no
		strUrl = strUrl & "|fr_dept_cd|" & fr_dept_cd
		strUrl = strUrl & "|ocpt_type|" & ocpt_type
		strUrl = strUrl & "|prov_dt|" & prov_dt
		strUrl = strUrl & "|to_dept_cd|" & to_dept_cd		
		strUrl = strUrl & "|std_sub|" & std_sub	    

    End With
End Function

'========================================================================================================
' Name : FncPrint	
' Desc : developer describe this line Called by MainDeleteRow in Common.vbs
'========================================================================================================
Function FncPrint()
	Call Parent.FncPrint()                                                      '☜: Protect system from crashing
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
	FncExit = True
End Function
'========================================================================================================
'   Event Name :txtbas_yy_DblClick
'   Event Desc : 달력을 호출한다.
'========================================================================================================
Sub txtbas_yy_DblClick(Button) 
    If Button = 1 Then
		Call SetFocusToDocument("M")    
        frm1.txtbas_yy.Action = 7
        frm1.txtbas_yy.focus
    End If
End Sub
'========================================================================================================
'   Event Name :txtrprt_dt_DblClick
'   Event Desc : 달력을 호출한다.
'========================================================================================================
Sub txtprov_dt_DblClick(Button) 
    If Button = 1 Then
		Call SetFocusToDocument("M")    
        frm1.txtprov_dt.Action = 7
        frm1.txtprov_dt.focus
    End If
End Sub

</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>근로소득영수증</font></td>
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
									<TD CLASS="TD5" NOWRAP>기준년</TD>
									<TD CLASS=TD6 NOWRAP COLSPAN=3><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 NAME="txtbas_yy" CLASS=FPDTYYYY tag="12X1" ALT="기준년" Title="FPDATETIME"></OBJECT>');</SCRIPT></TD>
								</TR>
							    <TR>
							    	<TD CLASS=TD5 NOWRAP>신고사업장</TD>
									<TD CLASS=TD6 NOWRAP COLSPAN=3><Select NAME="txtcust_cd" ALT="신고사업장" STYLE=" WIDTH: 200px" tag="12"></SELECT></TD>
								</TR>
								<TR>	
									<TD CLASS=TD5 NOWRAP>신고일</TD>
									<TD CLASS=TD6 NOWRAP COLSPAN=3><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=prov_dt NAME="txtprov_dt" CLASS=FPDTYYYYMMDD tag="12X1" ALT="신고일" Title="FPDATETIME"></OBJECT>');</SCRIPT></TD>
								</TR>
								<TR>
								<TD CLASS="TD5" NOWRAP>사번/성명</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" ID = "txtEmp_no" NAME="txtEmp_no" SIZE=10 MAXLENGTH=13 tag="11XXXU" ALT="사번"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnAcctCdFR" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenEmp ()">
								                       <INPUT TYPE="Text" NAME="txtName" SIZE=15 MAXLENGTH=30 tag="14XXXU" ALT="성명"></TD>
							    </TR>
							    <TD CLASS=TD5 NOWRAP>직종</TD>
						  	    	    <TD CLASS=TD6 NOWRAP><SELECT NAME="cboOcpt_type" ALT="직종" STYLE="WIDTH: 100px" TAG="11"><OPTION VALUE=""></OPTION></SELECT>		
							    <TR>
							    <TD CLASS=TD5 NOWRAP>부서</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtFr_dept_cd" ALT="부서코드" TYPE="Text" SiZE="10" MAXLENGTH="10" tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSItemDC" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenDept(0)">
			                            <INPUT NAME="txtFr_dept_nm" ALT="부서코드명" TYPE="Text" SiZE="20" MAXLENGTH="40" tag="14XXXU">
		                                <INPUT NAME="txtFr_Internal_cd" ALT="내부부서코드" TYPE="HIDDEN" SiZE="7" MAXLENGTH="7" tag="14XXXU">~</TD>
							    </TR>			
							    <TR>
							    	<TD CLASS=TD5 NOWRAP></TD>
							    	<TD CLASS=TD6 NOWRAP><INPUT NAME="txtto_dept_cd" ALT="부서코드" TYPE="Text" SIZE="10" MAXLENGTH="10" tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnname" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenDept(1)">
							                <INPUT NAME="txtto_dept_nm" ALT="부서코드명" TYPE="Text"SIZE="20" MAXLENGTH="40" tag="14XXXU">
							                <INPUT NAME="txtTo_Internal_cd" ALT="내부부서코드" TYPE="HIDDEN" SiZE="7" MAXLENGTH="7" tag="14XXXU"></TD>
							    </TR>
								<TR>	
									<TD CLASS=TD5 NOWRAP>&nbsp;</TD>
									<TD CLASS=TD6 NOWRAP COLSPAN=3>&nbsp;</TD>
								</TR>							    
				        	    <TR>
				        	        <TD CLASS="TD5" NOWRAP>당해퇴직자</TD>
				        	        <TD CLASS="TD6" COLSPAN=3><INPUT TYPE="RADIO" CLASS="RADIO" NAME="txtretire_check" TAG="1X" VALUE="포함" ID="retire_check0"><LABEL FOR="retire_check0">포함</LABEL>&nbsp;
				        	                                  <INPUT TYPE="RADIO" CLASS="RADIO" NAME="txtretire_check" TAG="1X" VALUE="미포함" CHECKED ID="retire_check1"><LABEL FOR="retire_check1">미포함</LABEL></TD>
				        	    </TR>							    
    			                <TR>
							    	<TD CLASS="TD5" NOWRAP>출력선택</TD>
				        	        <TD CLASS="TD6" COLSPAN=3><INPUT TYPE="CHECKBOX" CLASS="CHECK" NAME="prt_check1_flag" TAG="1X" VALUE="YES" ID="prt_check1_flag" CHECKED><LABEL FOR="prt_check1_flag">발행자보관용</LABEL>&nbsp;
				        	                                  <INPUT TYPE="CHECKBOX" CLASS="CHECK" NAME="prt_check2_flag" TAG="1X" VALUE="YES" ID="prt_check2_flag" CHECKED><LABEL FOR="prt_check2_flag">소득자보관용</LABEL>&nbsp;</TD>
                                </TR>
                                <TR><TD CLASS="TD5" NOWRAP>&nbsp;</TD>				        	                                  
				        	        <TD CLASS="TD6" COLSPAN=3><INPUT TYPE="CHECKBOX" CLASS="CHECK" NAME="prt_check3_flag" TAG="1X" VALUE="YES" ID="prt_check3_flag" CHECKED><LABEL FOR="prt_check3_flag">발행자보고용</LABEL>&nbsp;
				        	                                  <INPUT TYPE="CHECKBOX" CLASS="CHECK" NAME="prt_check4_flag" TAG="1X" VALUE="YES" ID="prt_check4_flag" CHECKED><LABEL FOR="prt_check4_flag">총액 출력</LABEL>&nbsp;</TD>
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
		                <BUTTON NAME="btnPreview" CLASS="CLSSBTN" ONCLICK="vbscript:FncBtnPreview()">미리보기</BUTTON>&nbsp;
		                <BUTTON NAME="btnPrint"   CLASS="CLSSBTN" ONCLICK="vbscript:FncBtnPrint()">인쇄</BUTTON>

		            </TD>
					<TD WIDTH=* ALIGN="right">&nbsp;</TD>
					<TD WIDTH=10>&nbsp;</TD>
		        </TR>
		    </TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC = "../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=no noresize framespacing=0></IFRAME>
		</TD>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP1" SRC = "../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=no noresize framespacing=0></IFRAME>
		</TD>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP2" SRC = "../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=no noresize framespacing=0></IFRAME>
		</TD>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP3" SRC = "../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=no noresize framespacing=0></IFRAME>		            
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

