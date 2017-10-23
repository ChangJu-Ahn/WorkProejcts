<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% session.CodePage=949 %>

<!--
'**********************************************************************************************
'*  1. Module명          : Human Resources
'*  2. Function명        : 급여관리 
'*  3. Program ID        : h6017oa1.asp
'*  4. Program 이름      : 급여지급대장출력 
'*  5. Program 설명      : 급여지급대장출력 
'*  6. Comproxy 리스트   : 
'*  7. 최초 작성년월일   : 2001/06/01
'*  8. 최종 수정년월일   : 2003/06/13
'*  9. 최초 작성자       : TGS 최용철 
'* 10. 최종 작성자       : Lee SiNa
'* 11. 전체 comment      :
'**********************************************************************************************-->
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
	
	frm1.txtpay_yymm.focus			'년월 default value setting
	
	frm1.txtpay_yymm.Year = strYear 		 '년월일 default value setting
	frm1.txtpay_yymm.Month = strMonth
End Sub

'========================================================================================================
' Name : LoadInfTB19029()	
' Desc : Set System Number format
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "H", "NOCOOKIE", "MA") %>
End Sub
'========================================================================================================
' Name : InitComboBox()
' Desc : Set ComboBox
'========================================================================================================
Sub InitComboBox()
    Dim iCodeArr 
    Dim iNameArr
    Dim iDx
    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6

    Call CommonQueryRs("MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0005", "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
 
    iCodeArr = lgF0
    iNameArr = lgF1

    Call SetCombo2(frm1.cboPay_cd,iCodeArr, iNameArr,Chr(11))

    Call  CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0122", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) 
    iCodeArr = lgF0
    iNameArr = lgF1

    Call SetCombo2(frm1.cboEmp_type,iCodeArr, iNameArr,Chr(11))

End Sub
'========================================================================================================
' Name : Form_Load
' Desc : developer describe this line Called by Window_OnLoad() evnt
'========================================================================================================
Sub Form_Load()
    Call LoadInfTB19029                                                     '⊙: Load table , B_numeric_format
    Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
    
	Call ggoOper.FormatDate(frm1.txtpay_yymm, Parent.gDateFormat, 2)

	Call InitVariables 
        
    Call FuncGetAuth(gStrRequestMenuID , Parent.gUsrID, lgUsrIntCd)     ' 자료권한:lgUsrIntCd ("%", "1%") 
      
    Call SetDefaultVal
    Call SetToolbar("1000000000000111")
    
    Call InitComboBox
End Sub

'========================================================================================================
' Name : Form_QueryUnload
' Desc : developer describe this line Called by Window_OnUnLoad() evnt
'========================================================================================================
Sub Form_QueryUnload(Cancel, UnloadMode)

End Sub
'========================================================================================================
' Name : FncBtnPrint
' Desc : developer describe this line 
'========================================================================================================
Function FncBtnPrint() 
	Call BtnDisabled(1)

	Dim strUrl
	Dim lngPos
	Dim intCnt
    Dim StrEbrFile
    Dim ObjName
    Dim strWhere
    
	Dim Pay_yymm, Pay_cd, Prov_type, Biz_area_cd
	Dim Fr_dept_cd , To_dept_cd, rFrDept ,rToDept ,IntRetCd
    Dim org_change_dt
'    Dim Emp_type
	
    If Not chkField(Document, "1") Then									<%'⊙: This function check indispensable field%>
        Call BtnDisabled(0)
		Exit Function
    End If

	StrEbrFile = "h6017oa1"
	
    Pay_yymm = frm1.txtPay_yymm.year & Right("0" & frm1.txtPay_yymm.month , 2)
	Pay_cd = frm1.cboPay_cd.value
	Prov_type = frm1.txtProv_type.value
'	Emp_type = frm1.cboEmp_type.value
	Biz_area_cd = frm1.txtBizAreaCd.value
	
	If Pay_cd = "" then
		Pay_cd = "%"
		frm1.cboPay_cd.value = ""
	End If	
'	If Emp_type = "" then
'		Emp_type = "%"
'		frm1.cboEmp_type.value = ""
'	End If	
	If Biz_area_cd = "" then
		Biz_area_cd = "%"
	End If	
    
    If  txtProv_Type_Onchange()  then
        Exit Function
    End If
    If  txtBizAreaCd_Onchange() then
        Exit Function
    End If    
    If  txtFr_Dept_cd_Onchange() then
        Exit Function
    End If
    If  txtTo_Dept_cd_Onchange() then
        Exit Function
    End If
    
    Fr_dept_cd = frm1.txtFr_internal_cd.value
    To_dept_cd = frm1.txtTo_internal_cd.value
    
    If fr_dept_cd = "" then
        IntRetCd = FuncGetTermDept(lgUsrIntCd ,"", rFrDept ,rToDept)
		frm1.txtFr_internal_cd.value = rFrDept
		frm1.txtFr_dept_nm.value = ""
	End If	
	
	If to_dept_cd = "" then
        IntRetCd = FuncGetTermDept(lgUsrIntCd ,"", rFrDept ,rToDept)
		frm1.txtTo_internal_cd.value = rToDept
		frm1.txtTo_dept_nm.value = ""
	End If  
    
    Fr_dept_cd = frm1.txtFr_internal_cd.value
    To_dept_cd = frm1.txtTo_internal_cd.value
    
    If (Fr_dept_cd<> "") AND (To_dept_cd<>"") Then       
        If Fr_dept_cd > To_dept_cd then
	        Call DisplayMsgbox("800153","X","X","X")	'시작부서코드는 종료부서코드보다 작아야합니다.
            frm1.txtFr_dept_cd.focus
            Set gActiveElement = document.activeElement
			Call BtnDisabled(0)
            Exit Function
        End IF 
    END IF   

'    strWhere = " org_change_dt <= (SELECT DISTINCT  top 1 prov_dt FROM HDF070T "
'    strWhere = strWhere +         " WHERE pay_yymm =  " & FilterVar(Pay_yymm, "''", "S")
'    strWhere = strWhere +         "   AND Prov_type =  " & FilterVar(Prov_type, "''", "S") & ")"

' 지급일자 기준이 아닌, 급여년월기준으로 부서정보가져오도록 수정 2007.04.13 
    strWhere = " convert(varchar(6),org_change_dt ,112)  <=    " & FilterVar(pay_yymm, "''", "S") & " "

    Call CommonQueryRs(" MAX(org_change_dt) "," b_acct_dept ", strWhere ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

    org_change_dt = Trim(Replace(lgF0,Chr(11),""))
    
	strUrl = "Pay_yymm|" & Pay_yymm
	strUrl = strUrl & "|Pay_cd|" & Pay_cd 
	strUrl = strUrl & "|Prov_type|" & Prov_type 
'	strUrl = strUrl & "|Emp_type|" & Emp_type 
	strUrl = strUrl & "|Biz_area_cd|" & Biz_area_cd 
	strUrl = strUrl & "|org_change_dt|" & org_change_dt 
	strUrl = strUrl & "|fr_dept_cd|" & fr_dept_cd
	strUrl = strUrl & "|to_dept_cd|" & to_dept_cd

    ObjName = AskEBDocumentName(StrEbrFile, "ebr")
   
 	call FncEBRPrint(EBAction , ObjName , strUrl)

End Function

'========================================================================================
' Function Name : FncBtnPreview()
' Function Desc : This function is related to Preview Button
'========================================================================================
Function FncBtnPreview()

	Dim strUrl
	Dim arrParam, arrField, arrHeader
    Dim StrEbrFile, ObjName
    Dim strWhere
	Dim Pay_yymm, Pay_cd, Prov_type, Biz_area_cd
    Dim Fr_dept_cd , To_dept_cd, rFrDept ,rToDept ,IntRetCd 
    Dim org_change_dt
'    Dim Emp_type


    If Not chkField(Document, "1") Then									'⊙: This function check indispensable field%>		
	   Exit Function
    End If
	
	StrEbrFile = "h6017oa1"
		
    Pay_yymm = frm1.txtPay_yymm.year & Right("0" & frm1.txtPay_yymm.month , 2)
	Pay_cd = frm1.cboPay_cd.value
	Prov_type = frm1.txtProv_type.value
'	Emp_type = frm1.cboEmp_type.value
	Biz_area_cd = frm1.txtBizAreaCd.value

	If Pay_cd = "" then
		Pay_cd = "%"
		frm1.cboPay_cd.value = ""
	End If	
'	If Emp_type = "" then
'		Emp_type = "%"
'		frm1.cboEmp_type.value = ""
'	End If	
	If Biz_area_cd = "" then
		Biz_area_cd = "%"
	End If	
	
    If  txtProv_Type_Onchange() then
        Exit Function
    End If
    If  txtBizAreaCd_Onchange() then
        Exit Function
    End If
    If  txtFr_Dept_cd_Onchange()  then
        Exit Function
    End If

    If  txtTo_Dept_cd_Onchange()  then
        Exit Function
    End If
   
    Fr_dept_cd = frm1.txtFr_internal_cd.value
    To_dept_cd = frm1.txtTo_internal_cd.value
    
    If fr_dept_cd = "" then
        IntRetCd = FuncGetTermDept(lgUsrIntCd ,"", rFrDept ,rToDept)
		frm1.txtFr_internal_cd.value = rFrDept
		frm1.txtFr_dept_nm.value = ""
	End If	
	
	If to_dept_cd = "" then
        IntRetCd = FuncGetTermDept(lgUsrIntCd ,"", rFrDept ,rToDept)
		frm1.txtTo_internal_cd.value = rToDept
		frm1.txtTo_dept_nm.value = ""
	End If  
    
    Fr_dept_cd = frm1.txtFr_internal_cd.value
    To_dept_cd = frm1.txtTo_internal_cd.value
    
    If (Fr_dept_cd<> "") AND (To_dept_cd<>"") Then       
        If Fr_dept_cd > To_dept_cd then
	        Call DisplayMsgbox("800153","X","X","X")	'시작부서코드는 종료부서코드보다 작아야합니다.
            frm1.txtFr_dept_cd.focus
            Set gActiveElement = document.activeElement
		
            Exit Function
        End IF 
    END IF     

'    strWhere = " org_change_dt <= (SELECT DISTINCT  top 1 prov_dt FROM HDF070T "
'    strWhere = strWhere +         " WHERE pay_yymm =  " & FilterVar(Pay_yymm, "''", "S")
'    strWhere = strWhere +         "   AND Prov_type =  " & FilterVar(Prov_type, "''", "S") & ")"

' 지급일자 기준이 아닌, 급여년월기준으로 부서정보가져오도록 수정 2007.04.13 
    strWhere = " convert(varchar(6),org_change_dt ,112)  <=    " & FilterVar(pay_yymm, "''", "S") & " "

    IntRetCd = CommonQueryRs(" MAX(org_change_dt) "," b_acct_dept ", strWhere ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

    org_change_dt = Trim(Replace(lgF0,Chr(11),""))

	strUrl = "Pay_yymm|" & Pay_yymm
	strUrl = strUrl & "|Pay_cd|" & Pay_cd 
	strUrl = strUrl & "|Prov_type|" & Prov_type 
'	strUrl = strUrl & "|Emp_type|" & Emp_type 
	strUrl = strUrl & "|Biz_area_cd|" & Biz_area_cd 
	strUrl = strUrl & "|org_change_dt|" & org_change_dt 
	strUrl = strUrl & "|fr_dept_cd|" & fr_dept_cd
	strUrl = strUrl & "|to_dept_cd|" & to_dept_cd

    ObjName = AskEBDocumentName(StrEbrFile, "ebr")
   	
	call FncEBRPreview(ObjName , strUrl)

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
	Call Parent.FncFind(Parent.C_SINGLE, False)
End Function

'========================================================================================================
' Name : FncExit
' Desc : developer describe this line Called by MainExit in Common.vbs
'========================================================================================================
Function FncExit()
	FncExit = True
End Function

'========================================================================================================
' Name : OpenDept
' Desc : 부서 POPUP
'========================================================================================================
Function OpenDept(iWhere)
	Dim arrRet
	Dim arrParam(2)
    Dim strBasDt 
    
	strBasDt = UNIGetLastDay(frm1.txtPay_yymm.text,Parent.gDateFormatYYYYMM)
    
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	If iWhere = 0 Then
		arrParam(0) = frm1.txtFr_dept_cd.value			            '  Code Condition
	ElseIf iWhere = 1 Then
		arrParam(0) = frm1.txtTo_dept_cd.value			            ' Code Condition
	End If
	
    arrParam(1) = strBasDt
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
'========================================================================================================
' Name : OpenCondAreaPopup()        
' Desc : developer describe this line 
'========================================================================================================
Function OpenCondAreaPopup(Byval iWhere)
	On Error Resume Next
    Err.Clear
    
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True  Then  
	   Exit Function
	End If   

	IsOpenPop = True
	Select Case iWhere
        Case "1"
	        arrParam(0) = "사업장팝업"			' 팝업 명칭 
	        arrParam(1) = "B_BIZ_AREA"				 		' TABLE 명칭 
	        arrParam(2) = frm1.txtBizAreaCd.value		    ' Code Condition
	        arrParam(3) = ""		' Name Cindition
	        arrParam(4) = ""        ' Where Condition
	        arrParam(5) = "사업장코드"			    ' TextBox 명칭 
	
            arrField(0) = "BIZ_AREA_CD"					' Field명(0)
            arrField(1) = "BIZ_AREA_NM"				    ' Field명(1)
    
            arrHeader(0) = "사업장코드"				' Header명(0)
            arrHeader(1) = "사업장명"			    ' Header명(1)
	   
        Case "2"
            arrParam(0) = "지급구분 팝업"			' 팝업 명칭 
	        arrParam(1) = "B_MINOR"				 		' TABLE 명칭 
	        arrParam(2) = frm1.txtProv_Type.value		    ' Code Condition
	        arrParam(3) = ""'frm1.txtProv_TypeNm.value			' Name Cindition
	        arrParam(4) = " MAJOR_CD = " & FilterVar("H0040", "''", "S") & " "    ' Where Condition							' Where Condition
	        arrParam(5) = "지급구분"			    ' TextBox 명칭 
	
            arrField(0) = "minor_cd"					' Field명(0)
            arrField(1) = "minor_nm"				    ' Field명(1)
    
            arrHeader(0) = "지급구분코드"				' Header명(0)
            arrHeader(1) = "지급구분명"

	End Select
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtProv_Type.focus	
		Exit Function
	Else
		Call SubSetCondArea(arrRet,iWhere)
	End If	
	
End Function

'======================================================================================================
'	Name : SetCondArea()
'	Description : Item Popup에서 Return되는 값 setting
'=======================================================================================================
Sub SubSetCondArea(Byval arrRet, Byval iWhere)
	On Error Resume Next
    Err.Clear
    
	With Frm1
		Select Case iWhere
		    Case "1"
		        .txtBizAreaCd.value = arrRet(0)
		        .txtBizAreaNm.value = arrRet(1)		
		        .txtBizAreaCd.focus
		    Case "2"
		        .txtProv_Type.value   = arrRet(0)
		        .txtProv_TypeNm.value = arrRet(1)
				.txtProv_Type.focus
        End Select
	End With

End Sub

'========================================================================================================
'   Event Name : txtBizAreaCd_change
'   Event Desc :
'========================================================================================================
Function txtBizAreaCd_Onchange()
    Dim IntRetCd
    
    If frm1.txtBizAreaCd.value = "" Then
		frm1.txtBizAreaNm.value = ""
    Else
        IntRetCd = CommonQueryRs(" BIZ_AREA_NM "," B_BIZ_AREA "," BIZ_AREA_CD= " & FilterVar(frm1.txtBizAreaCd.value, "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    
        If IntRetCd = false then
			Call DisplayMsgBox("124200","X","X","X")	
			 frm1.txtBizAreaNm.value = ""
             frm1.txtBizAreaCd.focus
            Set gActiveElement = document.ActiveElement
            txtBizAreaCd_Onchange = true 
            
            Exit Function          
        Else
			frm1.txtBizAreaNm.value = Trim(Replace(lgF0,Chr(11),""))
        End if 
    End if  
End Function

'========================================================================================================
'   Event Name : txtProv_Type_Onchange()            '<==지급구분 에러체크 
'   Event Desc :
'========================================================================================================
Function txtProv_Type_Onchange()
    On Error Resume Next
    Err.Clear
    
    Dim iDx
    Dim IntRetCd
    
    IF frm1.txtProv_Type.value = "" THEN
        frm1.txtProv_TypeNm.value = ""
    ELSE
        IntRetCd = CommonQueryRs(" minor_nm "," b_minor "," major_cd = " & FilterVar("H0040", "''", "S") & " and minor_cd =  " & FilterVar(frm1.txtProv_Type.value , "''", "S") & "" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
        If IntRetCd = false then
			Call DisplayMsgbox("800140","X","X","X")	'지급내역코드에 등록되지 않은 코드입니다.
            frm1.txtProv_TypeNm.value = ""
            frm1.txtProv_Type.focus
            txtProv_Type_Onchange = true
        ELSE    
            frm1.txtProv_TypeNm.value = Trim(Replace(lgF0,Chr(11),""))   '수당코드 
        END IF
    END IF 
End Function 
'========================================================================================================
'   Event Name : txtFr_dept_cd_Onchange
'   Event Desc :
'========================================================================================================
Function txtFr_dept_cd_Onchange()
    Dim IntRetCd,Dept_Nm,Internal_cd
	Dim rDate

	rDate = UNIGetLastDay(frm1.txtPay_yymm.Text, Parent.gDateFormatYYYYMM)

    If frm1.txtFr_dept_cd.value = "" Then
		frm1.txtFr_dept_nm.value = ""
		frm1.txtFr_internal_cd.value = ""
    Else
        IntRetCd = FuncDeptName(frm1.txtFr_dept_cd.value , UNIConvDate(rDate), lgUsrIntCd,Dept_Nm , Internal_cd)
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
    Dim rDate 
	
	rDate  = UNIGetLastDay(frm1.txtPay_yymm.Text, Parent.gDateFormatYYYYMM)

    If frm1.txtTo_dept_cd.value = "" Then
		frm1.txtTo_dept_nm.value = ""
		frm1.txtTo_internal_cd.value = ""
    Else
        IntRetCd = FuncDeptName(frm1.txtTo_dept_cd.value ,UNIConvDate(rDate), lgUsrIntCd,Dept_Nm , Internal_cd)
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


'========================================================================================================
' Name : txtPay_yymm_DblClick
' Desc : 달력 Popup을 호출 
'========================================================================================================
Sub txtPay_yymm_DblClick(Button)
	If Button = 1 Then
		Call SetFocusToDocument("M")	
		frm1.txtPay_yymm.Action = 7
		frm1.txtPay_yymm.focus
	End If
End Sub
</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>급여지급대장출력</font></td>
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
								<TD CLASS=TD5  NOWRAP>해당년월</TD>
								<TD CLASS=TD6  NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=txtPay_yymm name=txtPay_yymm CLASS=FPDTYYYYMM title=FPDATETIME ALT="해당년월" tag="12X1" VIEWASTEXT></OBJECT>');</SCRIPT></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>지급구분</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" ID ="txtProv_Type" NAME="txtProv_Type" SIZE=10 MAXLENGTH=3 tag="12XXXU" ALT="지급구분"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnAcctCdFR" ALIGN=top TYPE="BUTTON" ONCLICK="VBScript: OpenCondAreaPopup(2)">
								                       <INPUT TYPE="Text" NAME="txtProv_TypeNm" SIZE=20 MAXLENGTH=30 tag="14XXXU" ALT="지급구분"></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>급여구분</TD>
								<TD CLASS="TD6" NOWRAP><SELECT NAME="cboPay_cd" ALT="급여구분" CLASS=cboNormal TAG="1XN"><OPTION Value=""></OPTION></SELECT></TD>
							</TR>
							<TR>			
								<TD CLASS="TD5" NOWRAP>사업장</TD>
								<TD CLASS="TD6" NOWRAP><INPUT NAME="txtBizAreaCd" MAXLENGTH="10" SIZE=10 ALT ="사업장코드" tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBizAreaCd" align=top TYPE="BUTTON" onclick="vbscript: OpenCondAreaPopup('1')">
											           <INPUT NAME="txtBizAreaNm" MAXLENGTH="50" SIZE=20 ALT ="사업장명" tag="14X"></TD>
							</TR>
							<TR>			
							    <TD CLASS=TD5 NOWRAP>부서</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtFr_dept_cd" ALT="부서코드" TYPE="Text" SiZE=10 MAXLENGTH=10 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSItemDC" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenDept(0)">
			                                         <INPUT NAME="txtFr_dept_nm" ALT="부서코드명" TYPE="Text" SiZE=20 MAXLENGTH=40 tag="14XXXU">&nbsp;~
		                                             <INPUT NAME="txtFr_Internal_cd" ALT="내부부서코드" TYPE="HIDDEN" SiZE=7 MAXLENGTH=7 tag="14XXXU">
		                    </TR>
		                    <TR>    
		                        <TD CLASS=TD5 NOWRAP></TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtto_dept_cd" MAXLENGTH="10" SIZE=10 ALT ="Order ID" tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnname" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenDept(1)">
							                         <INPUT NAME="txtto_dept_nm" MAXLENGTH="40" SIZE=20 ALT ="Order ID" tag="14XXXU">
    			                                     <INPUT NAME="txtTo_Internal_cd" ALT="내부부서코드" TYPE="HIDDEN" SiZE=7 MAXLENGTH=7 tag="14XXXU"></TD>
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
		                <BUTTON NAME="btnPrint"   CLASS="CLSSBTN" OnClick="VBScript:FncBtnPrint()">인쇄</BUTTON>

		            </TD>
					<TD WIDTH=* ALIGN="right">&nbsp;</TD>
					<TD WIDTH=10>&nbsp;</TD>
		        </TR>
		    </TABLE>
		</TD>
	</TR>
	<TR>
		<TD HEIGHT=20><IFRAME NAME="MyBizASP" SRC = "../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=no noresize framespacing=0></IFRAME>
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

