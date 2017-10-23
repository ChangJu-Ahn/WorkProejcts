<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : Human Resources
*  2. Function Name        : 급여계산 
*  3. Program ID           : H6009ba1
*  4. Program Name         : H6009ba1
*  5. Program Desc         : 급여관리/급여계산 
*  6. Comproxy List        :
*  7. Modified date(First) : 2001/06/07
*  8. Modified date(Last)  : 2003/06/13
*  9. Modifier (First)     : YBI
* 10. Modifier (Last)      : Lee SiNa
* 11. Comment              :
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

<Script Language="VBScript">
Option Explicit
'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
Const BIZ_PGM_ID = "H6009bb1.asp"
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
    lgLngCurRows = 0                            'initializes Deleted Rows Count
End Sub

'========================================================================================================
' Name : SetDefaultVal()	
' Desc : Set default value
'========================================================================================================
	
Sub SetDefaultVal()
	
	Dim strYear
	Dim strMonth
	Dim strDay
	
    Call ggoOper.FormatDate(frm1.txtpay_yymm_dt, Parent.gDateFormat, 2)

	Call ExtractDateFrom("<%=GetsvrDate%>",Parent.gServerDateFormat , Parent.gServerDateType ,strYear,strMonth,strDay)
	
	frm1.txtpay_yymm_dt.Year = strYear 
	frm1.txtpay_yymm_dt.Month = strMonth 
	
	frm1.txtbas_dt.Year = strYear 
	frm1.txtbas_dt.Month = strMonth 
	frm1.txtbas_dt.Day = strDay
	
	frm1.txtprov_dt.Year = strYear 
	frm1.txtprov_dt.Month = strMonth 
	frm1.txtprov_dt.Day = strDay
	
	frm1.txtprov_cd.value = "1" ' 급여 
	
End Sub
	
'========================================================================================================
' Name : LoadInfTB19029()	
' Desc : Set System Number format
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "H", "NOCOOKIE", "BA") %>
End Sub
'========================================================================================================
'	Name : CookiePage()
'	Description : Item Popup에서 Return되는 값 setting
'========================================================================================================
Function CookiePage(ByVal flgs)
End Function

'========================================================================================================
' Function Name : MakeKeyStream
' Function Desc : This method set focus to pos of err
'========================================================================================================
Sub MakeKeyStream(pOpt)
   
   If pOpt = "Q" Then
      lgKeyStream = Frm1.txtWarrentNo.Value & Parent.gColSep       'You Must append one character(Parent.gColSep)
   Else
      lgKeyStream = Frm1.txtMajorCd.Value & Parent.gColSep         'You Must append one character(Parent.gColSep)
   End If   

End Sub        
	
'========================================================================================================
' Name : InitComboBox()
' Desc : developer describe this line Initialize ComboBox
'========================================================================================================
Sub InitComboBox()
    Dim iCodeArr 
    Dim iNameArr
    Dim iDx
    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
    
    Call CommonQueryRs("MINOR_CD, MINOR_NM ","B_MINOR","MAJOR_CD = " & FilterVar("H0005", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.txtPay_cd, lgF0, lgF1, Chr(11))

    Call CommonQueryRs("MINOR_CD, MINOR_NM ","B_MINOR","MAJOR_CD = " & FilterVar("H0040", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.txtProv_cd, lgF0, lgF1, Chr(11))
End Sub

'========================================================================================================
' Name : Form_Load
' Desc : developer describe this line Called by Window_OnLoad() evnt
'========================================================================================================
Sub Form_Load()

    Err.Clear                                                                       '☜: Clear err status
		
	Call LoadInfTB19029                                                     '⊙: Load table , B_numeric_format
	Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
	Call InitVariables                                                     '⊙: Setup the Spread sheet
	
	Call InitComboBox()

    Call FuncGetAuth(gStrRequestMenuID, Parent.gUsrID, lgUsrIntCd)     ' 자료권한:lgUsrIntCd ("%", "1%")
    frm1.txtpay_yymm_dt.focus
	Call SetDefaultVal
	Call SetToolbar("1000000000000111")										'⊙: 버튼 툴바 제어 
			
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
' Name : DbQuery
' Desc : This function is called by FncQuery
'========================================================================================================
Function DbQuery()
    Dim strVal
    Err.Clear                                                                    '☜: Clear err status

    DbQuery = True                                                               '☜: Processing is NG
End Function
'========================================================================================================
' Name : DbSave
' Desc : This function is called by FncSave
'========================================================================================================
Function DbSave()
	Dim strVal
    Err.Clear                                                                    '☜: Clear err status
		
	DbSave = False														         '☜: Processing is NG
		
    DbSave  = True                                                               '☜: Processing is NG
End Function
'========================================================================================================
' Name : DbDelete
' Desc : This function is called by FncDelete
'========================================================================================================
Function DbDelete()
	Dim strVal
    Err.Clear                                                                    '☜: Clear err status
		
	DbDelete = False			                                                 '☜: Processing is NG
		
	DbDelete = True                                                              '⊙: Processing is NG
End Function
'========================================================================================================
' Function Name : DbQueryOk
' Function Desc : Called by MB Area when query operation is successful
'========================================================================================================
Function DbQueryOk()


End Function
	
'========================================================================================================
' Function Name : DbSaveOk
' Function Desc : Called by MB Area when save operation is successful
'========================================================================================================
Function DbSaveOk()
    Call FncQuery()
End Function
	
'========================================================================================================
' Function Name : DbDeleteOk
' Function Desc : Called by MB Area when delete operation is successful
'========================================================================================================
Function DbDeleteOk()
	Call InitVariables()
End Function

'========================================================================================================
' Name : OpenEmp()
' Desc : developer describe this line 
'========================================================================================================

Function OpenEmp()
	Dim arrRet
	Dim arrParam(2)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = frm1.txtEmp_no.value			' Code Condition
	arrParam(1) = ""'frm1.txtName.value			' Name Cindition
	arrParam(2) = lgUsrIntCd
	
	arrRet = window.showModalDialog(HRAskPRAspName("EmpPopup"), Array(window.parent, arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtEmp_no.focus
		Exit Function
	Else
		Call SetEmp(arrRet)
	End If	
			
End Function

'======================================================================================================
'	Name : SetEmp()
'	Description : Item Popup에서 Return되는 값 setting
'=======================================================================================================
Sub SetEmp(arrRet)
	With frm1
		.txtEmp_no.value = arrRet(0)
		.txtName.value = arrRet(1)

		Call ggoOper.ClearField(Document, "2")					 '☜: Clear Contents  Field
		.txtEmp_no.focus
		lgBlnFlgChgValue = False
	End With
End Sub


'======================================================================================================
' Function Name : ExeReflect
' Function Desc : 
'=======================================================================================================
Function ExeReflect()
	 
	Dim strVal
	Dim IntRetCD
    Dim strChk_pay_cd
    Dim strNum
    Dim intCnt
    Dim li_yes  ' 12월 월차수당 option
    Dim strSQL, strSQL2
    
	Dim strWkYear
	Dim strWkMonth
	Dim strYYYYMMDD,strYYYYMM
	
	Dim strPayYYMMDt
	Dim strBaseDt
	Dim strProvDt
	Dim strPreDate

	ExeReflect = False                                                          '⊙: Processing is NG
    
	On Error Resume Next                                                   '☜: Protect system from crashing

	If Not chkField(Document, "1") Then
		Call BtnDisabled(0)
		Exit Function
	End If

	If txtEmp_no_OnChange() Then
		Exit Function
	End If

    strWkYear = frm1.txtpay_yymm_dt.Year
    strWkMonth = Right("0" & frm1.txtpay_yymm_dt.Month, 2)
    
    strYYYYMM = strWkYear & strWkMonth
	strYYYYMMDD = UniConvDateAToB(frm1.txtpay_yymm_dt,Parent.gDateFormatYYYYMM,Parent.gServerDateFormat)

    IF  FuncAuthority("!", strYYYYMMDD, Parent.gUsrID) = "N" THEN
        Call DisplayMsgBox("800294","X","X","X")
        Call BtnDisabled(0)
        exit function
    END IF

    strChk_pay_cd = frm1.txtPay_cd.value
    if  strChk_pay_cd = "" then
        strChk_pay_cd = "*"
    end if

    intCnt = 0
    If 	CommonQueryRs(" COUNT(*) "," hdf400t "," prov_type=" & FilterVar("1", "''", "S") & "  AND pay_cd= " & FilterVar(strChk_pay_cd, "''", "S") & "" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = true Then
        intCnt = CInt(Trim(Replace(lgF0, Chr(11), "")))
    end if
    if  intCnt <= 0 then
        Call DisplayMsgBox("800404","X","X","X")' 급여계산 수당/공제를 등록하여주십시오.
        Call BtnDisabled(0)
        Exit function
    end if

    intCnt = 0
    If 	CommonQueryRs(" COUNT(*) "," hdf400t "," prov_type=" & FilterVar("2", "''", "S") & " AND pay_cd= " & FilterVar(strChk_pay_cd, "''", "S") & "" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = true Then
        intCnt = CInt(Replace(lgF0, Chr(11), ""))
    end if
    if  intCnt <= 0 then
        Call DisplayMsgBox("800404","X","X","X")' 급여계산 수당/공제를 등록하여주십시오.
        Call BtnDisabled(0)
        Exit function
    end if


    li_yes = 2
    
    IF  strWkMonth = "12" OR strWkMonth = "01" THEN

        If 	CommonQueryRs(" COUNT(*) "," HDA140t "," prov_type=" & FilterVar("2", "''", "S") & " AND mm_accum=-1 AND use_mm=0",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = true Then
            intCnt = CInt(Replace(lgF0, Chr(11), ""))
        end if
    	IF  intCnt > 0 THEN
            IntRetCD = DisplayMsgBox("800266", Parent.VB_YES_NO, "X", "X")
	        If  IntRetCD = vbYes Then
	        	li_yes = 1
	        else
	            li_yes = 2
	        End If
       END IF   
    END IF

    strPreDate = UNIDateAdd("m", -1, frm1.txtbas_dt.text, Parent.gDateFormat)
    ' 수정하였음 
    if frm1.txtEmp_no.value = "" then
        StrSQL2 = " emp_no LIKE " & FilterVar("%", "''", "S") & ""
    else
		StrSQL2 = " emp_no =  " & FilterVar(frm1.txtemp_no.value, "''", "S") & ""
	end if
		
    strSQL = "wk_yymm =  " & FilterVar(strYYYYMM , "''", "S") & " and " & StrSQL2
    If 	CommonQueryRs(" count(emp_no) "," hca090t ", strSQL, lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = true Then    
        IF  Trim(Replace(lgF0, Chr(11), "")) = 0 THEN        
             ' 월근태집계를 하십시요.
            Call DisplayMsgBox("800503", "X", "X", "X")
	    	Call BtnDisabled(0)
	    	Exit Function
        END IF
    End if

    StrSQL = StrSQL2 & " AND (pay_cd IS NULL OR tax_cd IS NULL)"
    StrSQL = StrSQL & " AND (retire_dt IS NULL OR retire_dt > " & UNIConvDate(strPreDate) & ")"
    StrSQL = StrSQL & " AND entr_dt <=  " & FilterVar(UNIConvDate(frm1.txtbas_dt.text), "''", "S") & ""

    If 	CommonQueryRs(" MIN(emp_no) "," hdf020t ", strSQL, lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = true Then    
        IF  Trim(Replace(lgF0, Chr(11), "")) <> "" THEN        
             ' 급여마스터 정보에서 급여구분,세액구분을 입력해 주십시오.
            Call DisplayMsgBox("800387", "X", "X", "X")
	    	Call BtnDisabled(0)
	    	Exit Function
        END IF
    End if
    
	IntRetCD = DisplayMsgBox("900018",Parent.VB_YES_NO,"X","X")
	If IntRetCD = vbNo Then
		Call BtnDisabled(0)
		Exit Function
	End If

	If   LayerShowHide(1) = False Then
	     Call BtnDisabled(0)
	     Exit Function
	End If
	Call BtnDisabled(1)
	
	strPayYYMMDt = strWkYear & strWkMonth
	strBaseDt = frm1.txtbas_dt.Year & Right("0" & frm1.txtbas_dt.Month, 2) & Right("0" & frm1.txtbas_dt.Day, 2)  
	strProvDt = frm1.txtprov_dt.Year & Right("0" & frm1.txtprov_dt.Month, 2) & Right("0" & frm1.txtprov_dt.Day, 2)  
	
	strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0006
	strVal = strVal & "&txtpay_yymm_dt=" & Replace(strPayYYMMDt, Parent.gComDateType, "")
	strVal = strVal & "&txtbas_dt=" & strBaseDt
	strVal = strVal & "&txtprov_dt=" & strProvDt
	strVal = strVal & "&txtprov_cd=" & frm1.txtProv_cd.value
    strVal = strVal & "&txtpay_cd=" & frm1.txtpay_cd.value
    strVal = strVal & "&txtstand=1"
    strVal = strVal & "&txtLi_yes=" & Li_yes

    strVal = strVal & "&txtEmp_no=" & frm1.txtemp_no.value

	Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 

	ExeReflect = True                                                           '⊙: Processing is NG
	Call BtnDisabled(0)
End Function
'======================================================================================================
' Function Name : ExeDelete
' Function Desc : 
'=======================================================================================================
Function ExeDelete()
	 
	Dim strVal
	Dim IntRetCD
    Dim strChk_pay_cd
    Dim strNum
    Dim intCnt
    Dim li_yes  ' 12월 월차수당 option
    Dim strSQL, strSQL2
    
	Dim strWkYear
	Dim strWkMonth
	Dim strYYYYMMDD,strYYYYMM
	
	Dim strPayYYMMDt
	Dim strBaseDt
	Dim strProvDt
	Dim strPreDate
    Dim strEmp_no
    
	ExeDelete = False                                                          '⊙: Processing is NG
    
	On Error Resume Next                                                   '☜: Protect system from crashing

	If Not chkField(Document, "1") Then
		Call BtnDisabled(0)
		Exit Function
	End If

	If txtEmp_no_OnChange() Then
		Exit Function
	End If

    strWkYear = frm1.txtpay_yymm_dt.Year
    strWkMonth = Right("0" & frm1.txtpay_yymm_dt.Month, 2)
    
    strYYYYMM = strWkYear & strWkMonth
	strYYYYMMDD = UniConvDateAToB(frm1.txtpay_yymm_dt,Parent.gDateFormatYYYYMM,Parent.gServerDateFormat)

    IF  FuncAuthority("!", strYYYYMMDD, Parent.gUsrID) = "N" THEN
        Call DisplayMsgBox("800294","X","X","X")
        Call BtnDisabled(0)
        exit function
    END IF


    strEmp_no = Trim(frm1.txtemp_no.value)


    Dim strpay_cd , strType

    strChk_pay_cd = frm1.txtPay_cd.value
    if  strChk_pay_cd = "" then

		intCnt = 0
		If 	CommonQueryRs(" COUNT(*) "," hdf070t "," prov_type='1' AND pay_yymm = '" & strYYYYMM & "' AND pay_cd like'" & strChk_pay_cd & "%'" & "AND emp_no like'" & strEmp_no & "%'" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = true Then
		    intCnt = CInt(Trim(Replace(lgF0, Chr(11), "")))
		end if
        strChk_pay_cd = "*"

		strpay_cd  = "*"
	    Call  CommonQueryRs(" MINOR_NM "," H_PAY_CD "," MINOR_CD = '" & strpay_cd & "'",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	    strpay_cd = Trim(Replace(lgF0, Chr(11), ""))

	Else
		Call CommonQueryRs("MINOR_NM ","B_MINOR","MAJOR_CD = 'H0005'AND MINOR_CD = '" & strChk_pay_cd & "'" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
		strpay_cd = Trim(Replace(lgF0, Chr(11), ""))

		intCnt = 0
		If 	CommonQueryRs(" COUNT(*) "," hdf070t "," prov_type='1' AND pay_yymm = '" & strYYYYMM & "' AND pay_cd ='" & strChk_pay_cd & "'"  & "AND emp_no like'" & strEmp_no & "%'" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = true Then
		    intCnt = CInt(Trim(Replace(lgF0, Chr(11), "")))
		end if		
		
    End If
    
	If intCnt = 0 Then
		Call DisplayMsgbox("800161","X","X","X")
		Exit Function

	End If
	  
    Call CommonQueryRs(" MINOR_NM "," B_MINOR "," MAJOR_CD = 'H0040'  AND MINOR_CD = '" & frm1.txtProv_cd.value & "'" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) 
    strType = Trim(Replace(lgF0, Chr(11), ""))

	           
	IntRetCD = DisplayMsgBox("800805", Parent.VB_YES_NO , "(" & frm1.txtPay_yymm_dt.text &") "& strpay_cd , strType & " " & intCnt)  '☜: Data is changed.  Do you want to display it? 
	If IntRetCD <> vbYes Then
	    lgBlnFlgChgValue = False
	    Call BtnDisabled(0)
		Exit Function
	End If


	If   LayerShowHide(1) = False Then
	     Call BtnDisabled(0)
	     Exit Function
	End If
	
	Call BtnDisabled(1)
	
	strPayYYMMDt = strWkYear & strWkMonth
	strBaseDt = frm1.txtbas_dt.Year & Right("0" & frm1.txtbas_dt.Month, 2) & Right("0" & frm1.txtbas_dt.Day, 2)  
	strProvDt = frm1.txtprov_dt.Year & Right("0" & frm1.txtprov_dt.Month, 2) & Right("0" & frm1.txtprov_dt.Day, 2)  
	
	strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0003
	strVal = strVal & "&txtpay_yymm_dt=" & Replace(strPayYYMMDt, Parent.gComDateType, "")
	strVal = strVal & "&txtprov_cd=" & frm1.txtProv_cd.value
    strVal = strVal & "&txtpay_cd=" & frm1.txtpay_cd.value
    strVal = strVal & "&txtemp_no=" & frm1.txtemp_no.value

	Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 

	ExeDelete = True                                                           '⊙: Processing is NG
	Call BtnDisabled(0)
End Function
'======================================================================================================
' Function Name : ExeReflectOk
' Function Desc : ExeReflect가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncSave에 있는것을 옮김 
'=======================================================================================================
Function ExeReflectOk()				            '☆: 저장 성공후 실행 로직 
	Dim IntRetCD 

	IntRetCD =DisplayMsgBox("990000","X","X","X")

End Function
Function ExeReflectNo()				            '☆: 실행된 자료가 없습니다 
	Dim IntRetCD 

    Call DisplayMsgBox("800161","X","X","X")

End Function

Function FuncAuthority(Pay_gubun, Pay_yymmdd, Emp_no)

    Dim strRet
    Dim IntRetCD

    strRet = "N"
    IntRetCD = CommonQueryRs("close_type, CONVERT(CHAR(10),close_dt, 20), emp_no","hda270t","org_cd=" & FilterVar("1", "''", "S") & "  and pay_gubun=" & FilterVar("Z", "''", "S") & "  and pay_type= " & FilterVar(Pay_gubun, "''", "S") & "" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    if  IntRetCD = false then
        strRet = "Y"
    else
        SELECT CASE Replace(lgF0, Chr(11), "")
        	CASE "1" '마감형태 : 정상 
        	    IF  Replace(lgF1, Chr(11), "") <= Pay_yymmdd THEN 
        	        strRet = "Y"
        		ELSE
        	        strRet = "N" 
        		END IF
           CASE "2" '마감형태 : 마감 
        	    IF  Replace(lgF1, Chr(11), "") < Pay_yymmdd THEN 
        	        strRet = "Y" 
        		ELSE
        	        strRet = "N" 
        	    END IF
        END SELECT
        
    end if

    FuncAuthority = strRet

End Function

Sub txtpay_yymm_dt_DblClick(Button)
	If Button = 1 Then
		Call SetFocusToDocument("M")	
		frm1.txtpay_yymm_dt.Action = 7
		frm1.txtpay_yymm_dt.focus
	End If
End Sub

Sub txtbas_dt_DblClick(Button)
	If Button = 1 Then
		Call SetFocusToDocument("M")	
		frm1.txtbas_dt.Action = 7
		frm1.txtbas_dt.focus
	End If
End Sub

Sub txtprov_dt_DblClick(Button)
	If Button = 1 Then
		Call SetFocusToDocument("M")	
		frm1.txtprov_dt.Action = 7
		frm1.txtprov_dt.focus
	End If
End Sub
'========================================================================================================
'   Event Name : txtEmp_no_Onchange             
'   Event Desc :
'========================================================================================================
Function txtEmp_no_OnChange()

    Dim IntRetCd
    Dim strName
    Dim strDept_nm
    Dim strRoll_pstn
    Dim strPay_grd1
    Dim strPay_grd2
    Dim strEntr_dt
    Dim strInternal_cd

	frm1.txtName.value = ""
           
    If  frm1.txtEmp_no.value = "" Then
    Else
	    IntRetCd = FuncGetEmpInf2(frm1.txtEmp_no.value,lgUsrIntCd,strName,strDept_nm,_
	                strRoll_pstn, strPay_grd1, strPay_grd2, strEntr_dt, strInternal_cd)
	    if  IntRetCd < 0 then
	        if  IntRetCd = -1 then
    			Call DisplayMsgBox("800048","X","X","X")	'해당사원은 존재하지 않습니다.
            else
                Call DisplayMsgBox("800454","X","X","X")	'자료에 대한 권한이 없습니다.
            end if
            frm1.txtEmp_no.focus
            Set gActiveElement = document.ActiveElement
            txtEmp_no_OnChange = true
        Else
            frm1.txtName.value = strName
        End if 
    End if

End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>
 
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>급여계산</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=*>&nbsp;</TD>
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
								<TD CLASS="TD5" NOWRAP>급여년월</TD>
								<TD CLASS="TD6" NOWRAP><script language =javascript src='./js/h6009ba1_txtPay_yymm_dt_txtPay_yymm_dt.js'></script>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>지급구분</TD>
								<TD CLASS="TD6" NOWRAP><SELECT Name="txtProv_cd" ALT="지급구분" CLASS ="cbonormal" tag="14"><OPTION Value=""></OPTION></SELECT></TD>
							</TR>
							<TR>
							    <TD CLASS=TD5  NOWRAP>기준일</TD>
								<TD CLASS=TD6  NOWRAP><script language =javascript src='./js/h6009ba1_txtBas_dt_txtBas_dt.js'></script>
							</TR>
							<TR>
							    <TD CLASS=TD5  NOWRAP>지급일</TD>
								<TD CLASS=TD6  NOWRAP><script language =javascript src='./js/h6009ba1_txtProv_dt_txtProv_dt.js'></script></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>급여구분</TD>
								<TD CLASS=TD6 NOWRAP><SELECT Name="txtPay_cd" ALT="급여구분" CLASS ="cbonormal" tag="11"><OPTION Value=""></OPTION></SELECT></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>대상자</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtEmp_no" ALT="대상자" TYPE="Text" MAXLENGTH=13 SIZE=13 tag=11XXXU><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBP_CD" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenEmp()">&nbsp;<INPUT NAME="txtName" TYPE="Text" MAXLENGTH=30 SIZE=20 tag=14XXXU></TD>
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
					<TD><BUTTON NAME="btnExe"  CLASS="CLSSBTN" onclick="ExeReflect()" Flag=1>실행</BUTTON>&nbsp;
					    <BUTTON NAME="btnExe2" CLASS="CLSSBTN" onclick="ExeDelete()"  Flag=1>취소</BUTTON></TD>
					<TD WIDTH=* ALIGN="right">&nbsp;</TD>
					<TD WIDTH=10>&nbsp;</TD>
		        </TR>
		    </TABLE>
		</TD>
	</TR>
	<TR>
		<TD HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC = "../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>



