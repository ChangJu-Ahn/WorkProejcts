<%@ LANGUAGE="VBSCRIPT" %>
<!--======================================================================================================
*  1. Module Name          : �λ�/�޿����� 
*  2. Function Name        : �޿����� 
*  3. Program ID           : h6020ma1
*  4. Program Name         : ������ü���ϻ��� 
*  5. Program Desc         : ������ü���ϻ��� 
*  6. Comproxy List        : +
*  7. Modified date(First) : 2001/05/27
*  8. Modified date(Last)  : 2003/06/13
*  9. Modifier (First)     : Shin Kwang-Ho
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

<SCRIPT LANGUAGE=VBSCRIPT>
Option Explicit
'==============================================================================================
'=                       4.3 Common variables 
'==============================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->				           '��: Biz Logic ASP Name

'==============================================================================================
'							1.2.3 Global Variable�� ����  
'==============================================================================================
Const BIZ_PGM_ID      = "h6020mb1.asp"		

Dim IsOpenPop
Dim lgOldRow    
'==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : ���� �ʱ�ȭ(Global ����, �ʱ�ȭ�� �ʿ��� ���� �Ǵ� Flag���� Setting�Ѵ�.)
'=========================================================================================================
Sub InitVariables()
    lgIntFlgMode = Parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
End Sub
'==========================================  2.2.1 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : ȭ�� �ʱ�ȭ(���� Field�� �� �� ȭ���� �� �� Default���� ������� �ϴ� Field�� Setting)
'=========================================================================================================
Sub SetDefaultVal()
	Dim strYear
	Dim strMonth
	Dim strDay
	
	Call ExtractDateFrom("<%=GetsvrDate%>",Parent.gServerDateFormat , Parent.gServerDateType ,strYear,strMonth,strDay)
	
	frm1.txtPay_yymm_dt.Focus			'��� default value setting
	
	frm1.txtPay_yymm_dt.Year = strYear 
	frm1.txtPay_yymm_dt.Month = strMonth 

	frm1.txtYy_mm_dd_dt.Year = strYear 		 '����� default value setting
	frm1.txtYy_mm_dd_dt.Month = strMonth 
	frm1.txtYy_mm_dd_dt.Day = strDay
	
End Sub
'========================================================================================================
' Name : InitComboBox()
' Desc : Set ComboBox
'========================================================================================================
Sub InitComboBox()
    Dim iNameArr1,iCodeArr1
        
    Call CommonQueryRs("MINOR_NM,MINOR_CD","B_MINOR","MAJOR_CD = " & FilterVar("H0005", "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    iNameArr1 = lgF0
    iCodeArr1 = lgF1
    Call SetCombo2(frm1.txtcboPay_cd,iCodeArr1, iNameArr1,Chr(11))            ''''''''DB���� �ҷ� condition����        
End Sub
'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'========================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "H", "NOCOOKIE", "BA") %>
End Sub
'========================================================================================================
' Name : OpenDept
' Desc : �μ� POPUP
'========================================================================================================
Function OpenDept(iWhere)
	Dim arrRet
	Dim arrParam(2)
    Dim strBasDt 
    Dim strBasDtAdd
    	
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	If iWhere = 0 Then
		arrParam(0) = frm1.txtFr_dept_cd.value			            '  Code Condition
	ElseIf iWhere = 1 Then
		arrParam(0) = frm1.txtTo_dept_cd.value			            ' Code Condition
	End If
	
    arrParam(1) = ""
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
'	Name : OpenPopUp
'	Description : Code PopUp at vspdData
'=======================================================================================================
Function OpenPopUp(Byval strCode, Byval iWhere)

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function
	IsOpenPop = True
		
	Select Case iWhere
	    Case "PROV_TYPE"
			arrParam(0) = "���ޱ��� �˾�"			        ' �˾� ��Ī 
	    	arrParam(1) = "B_minor"							    ' TABLE ��Ī 
	    	arrParam(2) = frm1.txtProv_type.value     			' Code Condition
	    	arrParam(3) = ""'frm1.txtProv_type_nm.value									' Name Cindition
	    	arrParam(4) = "major_cd = " & FilterVar("H0040", "''", "S") & ""	   		    	' Where Condition
	    	arrParam(5) = "�����ڵ�"  			            ' TextBox ��Ī 
	
	    	arrField(0) = "minor_cd"						    ' Field��(0)
	    	arrField(1) = "minor_nm"    					  	' Field��(1)
	    	arrField(2) = ""    				        		' Field��(2)
    
	    	arrHeader(0) = "�����ڵ�"	   		    	    ' Header��(0)
	    	arrHeader(1) = "�����ڵ��"        		        ' Header��(1)
	    	arrHeader(2) = ""	    							' Header��(1)
	   Case "SECT_CD"
			arrParam(0) = "�ٹ����� �˾�"			        ' �˾� ��Ī 
	    	arrParam(1) = "B_minor"							    ' TABLE ��Ī 
	    	arrParam(2) = frm1.txtSect_cd.value     			' Code Condition
	    	arrParam(3) = ""'frm1.txtSect_cd_nm.value									' Name Cindition
	    	arrParam(4) = "major_cd = " & FilterVar("H0035", "''", "S") & ""	   		    	' Where Condition
	    	arrParam(5) = "�ٹ������ڵ�"  			            ' TextBox ��Ī 
	
	    	arrField(0) = "minor_cd"						    ' Field��(0)
	    	arrField(1) = "minor_nm"    					  	' Field��(1)
	    	arrField(2) = ""    				        		' Field��(2)
    
	    	arrHeader(0) = "�ٹ������ڵ�"	   		    	    ' Header��(0)
	    	arrHeader(1) = "�ٹ������ڵ��"        		        ' Header��(1)
	    	arrHeader(2) = ""	    							' Header��(1)
	    Case "OCPT_TYPE"
	        arrParam(0) = "���� �˾�"			        ' �˾� ��Ī 
	    	arrParam(1) = "B_minor"							    ' TABLE ��Ī 
	    	arrParam(2) = frm1.txtOcpt_type.value     			' Code Condition
	    	arrParam(3) = ""'frm1.txtOcpt_type_nm.value									' Name Cindition
	    	arrParam(4) = "major_cd = " & FilterVar("H0003", "''", "S") & ""	   		    	' Where Condition
	    	arrParam(5) = "�ٹ������ڵ�"  			            ' TextBox ��Ī 
	
	    	arrField(0) = "minor_cd"						    ' Field��(0)
	    	arrField(1) = "minor_nm"    					  	' Field��(1)
	    	arrField(2) = ""    				        		' Field��(2)
    
	    	arrHeader(0) = "�����ڵ�"	   		    	    ' Header��(0)
	    	arrHeader(1) = "�����ڵ��"        		        ' Header��(1)
	    	arrHeader(2) = ""	    							' Header��(1)	    
	End Select	    
    
    arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		    "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
    	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Select Case iWhere
		    Case "PROV_TYPE"
				frm1.txtProv_type.focus
		    Case "SECT_CD"
				frm1.txtSect_cd.focus
		    Case "OCPT_TYPE"
				frm1.txtOcpt_type.focus
        End Select	
		Exit Function
	Else
		Call SetCode(arrRet, iWhere)
	End If	

End Function
'======================================================================================================
'	Name : SetCode()
'	Description : Code PopUp���� Return�Ǵ� �� setting
'=======================================================================================================
Function SetCode(Byval arrRet, Byval iWhere)

	With frm1

		Select Case iWhere
		    Case "PROV_TYPE"
				.txtProv_type.value = arrRet(0) 
				.txtProv_type_nm.value  = arrRet(1) 
				.txtProv_Oldtype.value = arrRet(0) 
				.txtProv_type.focus
		    Case "SECT_CD"
				.txtSect_cd.value = arrRet(0) 
				.txtSect_cd_nm.value  = arrRet(1)
				.txtSect_Oldcd.value = arrRet(0) 
				.txtSect_cd.focus
		    Case "OCPT_TYPE"
		        .txtOcpt_type.value = arrRet(0) 
				.txtOcpt_type_nm.value  = arrRet(1)
				.txtOcpt_Oldtype.value = arrRet(0) 
				.txtOcpt_type.focus
        End Select
	End With

End Function
'==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(���� Include ���Ͽ� ����)�� �����ʱ�ȭ �� ȭ���ʱ�ȭ�� �ϱ� ���� �Լ��� Call�ϴ� �κ� 
'=========================================================================================================
Sub Form_Load()

    Call LoadInfTB19029                                                     '��: Load table , B_numeric_format
    Call ggoOper.LockField(Document, "N")                                   '��: Lock  Suitable  Field
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)

	Call ggoOper.FormatDate(frm1.txtPay_yymm_dt, Parent.gDateFormat, 2)   '��� 
    Call ggoOper.FormatDate(frm1.txtYy_mm_dd_dt, Parent.gDateFormat, 1)             '����� 
	
	Call FuncGetAuth(gStrRequestMenuID , Parent.gUsrID, lgUsrIntCd)     ' �ڷ����:lgUsrIntCd ("%", "1%")
	
    Call InitVariables
    Call InitComboBox                                        
    Call SetDefaultVal
    Call SetToolbar("1000000000000111")    
     
End Sub
'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================
Sub Form_QueryUnload(Cancel, UnloadMode)
End Sub
'======================================================================================================
'   Event Name : txtYyyymm_DblClick
'   Event Desc : �޷� Popup�� ȣ�� 
'=======================================================================================================
Sub txtPay_yymm_dt_DblClick(Button) 
	If Button = 1 Then
		Call SetFocusToDocument("M")	
		frm1.txtPay_yymm_dt.Action = 7
		frm1.txtPay_yymm_dt.focus
	End If
End Sub
'======================================================================================================
'   Event Name : txtYyyymm_DblClick
'   Event Desc : �޷� Popup�� ȣ�� 
'=======================================================================================================
Sub txtYy_mm_dd_dt_DblClick(Button) 
	If Button = 1 Then
		Call SetFocusToDocument("M")	
		frm1.txtYy_mm_dd_dt.Action = 7
		frm1.txtYy_mm_dd_dt.focus
	End If
End Sub
'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================
Function FncQuery() 

  With frm1
	
    If  txtProv_type_Onchange() Then   
        Exit Function
    End if
    
    If  txtSect_cd_OnChange() Then
        Exit Function
    End if	
    If  txtOcpt_type_Onchange() Then
        Exit Function
    End if	    	
    If  txtFr_dept_cd_Onchange() Then      'enter key �� ��ȸ�� ���ۺμ��ڵ带 check�� �ش���� ������ query����...
        Exit Function
    End if
    
    If  txtTo_dept_cd_Onchange() Then     'enter key �� ��ȸ�� ����μ��ڵ带 check�� �ش���� ������ query����...
        Exit Function
    End if
 
    Dim Fr_dept_cd , To_dept_cd, rFrDept ,rToDept ,IntRetCd
    
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
    
    If (Fr_dept_cd= "") AND (To_dept_cd="") Then       
    Else
        If Fr_dept_cd > To_dept_cd then
	        Call DisplayMsgBox("800153","X","X","X")	'���ۺμ��ڵ�� ����μ��ڵ庸�� �۾ƾ��մϴ�.
            frm1.txtfr_dept_cd.focus()
            Set gActiveElement = document.activeElement
        End IF 
        
    END IF   
    End with
End Function

'========================================================================================
' Function Name : txtProv_type_OnChange()
' Function Desc : 
'========================================================================================
Function txtProv_type_OnChange()
    Dim iDx
    Dim IntRetCd
    
    IF frm1.txtProv_type.value = "" THEN
        frm1.txtProv_type_nm.value = ""
    ELSE
        IntRetCd = CommonQueryRs(" minor_nm "," b_minor "," major_cd = " & FilterVar("H0040", "''", "S") & " and minor_cd =  " & FilterVar(frm1.txtProv_type.value , "''", "S") & "" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
        If IntRetCd = false then
			Call DisplayMsgBox("800140","X","X","X")	'���޳����ڵ忡 ��ϵ��� ���� �ڵ��Դϴ�.
            frm1.txtProv_type_nm.value = ""
            frm1.txtProv_type.focus
            txtProv_type_Onchange = true
            Exit Function
        ELSE    
            frm1.txtProv_type_nm.value = Trim(Replace(lgF0,Chr(11),""))   '�����ڵ� 
        END IF
    END IF 
End Function
'========================================================================================
' Function Name : txtSect_cd_OnChange()
' Function Desc : 
'========================================================================================
Function txtSect_cd_OnChange()
   Dim iDx
    Dim IntRetCd
        
    If frm1.txtSect_cd.value = "" Then
        frm1.txtSect_cd_nm.value = ""
    ELSE
        IntRetCd = CommonQueryRs(" MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0035", "''", "S") & " AND MINOR_CD =  " & FilterVar(frm1.txtSect_cd.value , "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) 
        IF IntRetCd = False THEN
            Call DisplayMsgBox("970000","X","�ٹ������ڵ�","X")
	        frm1.txtSect_cd_nm.value = ""
	        Set gActiveElement = document.ActiveElement
	        txtSect_cd_OnChange = true
	        Exit Function
	    Else
	        frm1.txtSect_cd_nm.value = Trim(Replace(lgF0, Chr(11), ""))

	    End If
    End If
    	    
End Function
'========================================================================================
' Function Name : txtProv_type_OnChange()
' Function Desc : 
'========================================================================================
Function txtOcpt_type_OnChange()
   Dim iDx
    Dim IntRetCd
        
    If frm1.txtOcpt_type.value = "" Then
        frm1.txtOcpt_type_nm.value = ""
    ELSE
        IntRetCd = CommonQueryRs(" MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0003", "''", "S") & " AND MINOR_CD =  " & FilterVar(frm1.txtOcpt_type.value , "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) 
        IF IntRetCd = False THEN
            Call DisplayMsgBox("970000","X","�����ڵ�","X")
	        frm1.txtOcpt_type_nm.value = ""
	        Set gActiveElement = document.ActiveElement
	        txtOcpt_type_OnChange = true
	        Exit Function
	    Else
	        frm1.txtOcpt_type_nm.value = Trim(Replace(lgF0, Chr(11), ""))
	    End If
    End If
End Function

'========================================================================================================
'   Event Name : txtFr_dept_cd_change
'   Event Desc :
'========================================================================================================
Function txtFr_dept_cd_Onchange()
    Dim IntRetCd,Dept_Nm,Internal_cd
    
    If frm1.txtFr_dept_cd.value = "" Then
		frm1.txtFr_dept_nm.value = ""
		frm1.txtFr_internal_cd.value = ""        
    Else
        IntRetCd = FuncDeptName(frm1.txtFr_dept_cd.value , "" , lgUsrIntCd,Dept_Nm , Internal_cd)
	    If  IntRetCd < 0 then
	        If  IntRetCd = -1 then
    			Call DisplayMsgBox("800098","X","X","X")	'�μ������ڵ忡 ��ϵ��� ���� �ڵ��Դϴ�.
            Else
                Call DisplayMsgBox("800454","X","X","X")	'�ڷῡ ���� ������ �����ϴ�.
            End if
		    frm1.txtFr_dept_nm.value = ""
		    frm1.txtFr_internal_cd.value = ""
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
    
    If frm1.txtTo_dept_cd.value = "" Then
		frm1.txtTo_dept_nm.value = ""
		frm1.txtTo_internal_cd.value = ""        
    Else
        IntRetCd = FuncDeptName(frm1.txtTo_dept_cd.value , "" , lgUsrIntCd,Dept_Nm , Internal_cd)
	    If  IntRetCd < 0 then
	        If  IntRetCd = -1 then
    			Call DisplayMsgBox("800098","X","X","X")	'�μ������ڵ忡 ��ϵ��� ���� �ڵ��Դϴ�.
            Else
                Call DisplayMsgBox("800454","X","X","X")	'�ڷῡ ���� ������ �����ϴ�.
            End if
		    frm1.txtTo_dept_nm.value = ""
		    frm1.txtTo_internal_cd.value = ""
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
' Function Name : btnAction_OnClick()
' Function Desc : This function is related to Preview Button
'========================================================================================
Function btnAction_OnClick()
    Dim Fr_dept_cd, To_dept_cd, rFrDept ,rToDept
	dim IntRetCd
    If Not chkField(Document, "1") Then									'��: This function check indispensable field
       Exit Function
    End If

    Dim strVal, RetFlag, pay_yymm, prov_type, gigup_type1, strDate 
	
	pay_yymm  = frm1.txtPay_yymm_dt.Year & Right("0" & frm1.txtPay_yymm_dt.Month, 2)
	strDate   = frm1.txtYy_mm_dd_dt.Year & Right("0" & frm1.txtYy_mm_dd_dt.Month, 2) & Right("0" & frm1.txtYy_mm_dd_dt.Day, 2) 
	prov_type = frm1.txtProv_type.value

	If frm1.txtGigup_type(0).checked Then 
		gigup_type1 = "1"
	Elseif frm1.txtGigup_type(1).checked Then 
		gigup_type1 = "2"
	Else 
		gigup_type1 = "3" 
	End if	
		
	If  txtProv_type_Onchange() Then   
	    Exit Function
	End if
		    
	If  txtSect_cd_OnChange() Then
	    Exit Function
	End if	
	If  txtOcpt_type_Onchange() Then
	    Exit Function
	End if	    	
	If  txtFr_dept_cd_Onchange() Then      'enter key �� ��ȸ�� ���ۺμ��ڵ带 check�� �ش���� ������ query����...
	    Exit Function
	End if
		    
	If  txtTo_dept_cd_Onchange() Then     'enter key �� ��ȸ�� ����μ��ڵ带 check�� �ش���� ������ query����...
	    Exit Function
	End if

	
	With frm1
       Fr_dept_cd = .txtFr_internal_cd.value
        To_dept_cd = .txtTo_internal_cd.value
        
        If Fr_dept_cd = "" then
            IntRetCd = FuncGetTermDept(lgUsrIntCd ,"", rFrDept ,rToDept)
            .txtFr_internal_cd.value = rFrDept
            .txtFr_dept_nm.value = ""
        End If	
            
        If To_dept_cd = "" then
            IntRetCd = FuncGetTermDept(lgUsrIntCd ,"", rFrDept ,rToDept)
            .txtTo_internal_cd.value = rToDept
            .txtTo_dept_nm.value = ""
        End If  

        Fr_dept_cd = .txtFr_internal_cd.value
        To_dept_cd = .txtTo_internal_cd.value

        If (Fr_dept_cd<> "") AND (To_dept_cd<>"") Then       
    
            If Fr_dept_cd > To_dept_cd then
	            Call DisplayMsgBox("800153","X","X","X")	'���ۺμ��ڵ�� ����μ��ڵ庸�� �۾ƾ��մϴ�.
                .txtFr_dept_cd.focus()
                Set gActiveElement = document.activeElement
                Exit Function
            End IF 
            
        END IF   

	    RetFlag = DisplayMsgBox("900018", Parent.VB_YES_NO,"x","x")   '�� �ٲ�κ�	

	    If RetFlag = VBNO Then
	    	Exit Function
	    End IF

	    If LayerShowHide(1) = False Then
	       Exit Function
	    End If		
	    strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001							'��: �����Ͻ� ó�� ASP�� ���� 	    	    
		strVal = strVal & "&txtPay_yymm_dt=" & Trim(pay_yymm)
		strVal = strVal & "&txtProv_type=" & Trim(.txtProv_type.value)
		strVal = strVal & "&txtYy_mm_dd_dt=" & Trim(strDate)
		strVal = strVal & "&txtcboPay_cd=" & Trim(.txtcboPay_cd.value)
		strVal = strVal & "&txtSect_cd=" & Trim(.txtSect_cd.value)
		strVal = strVal & "&txtOcpt_type=" & Trim(.txtOcpt_type.value)
		strVal = strVal & "&txtFr_dept_cd=" & Trim(.txtFr_Internal_cd.value)
		strVal = strVal & "&txtTo_dept_cd=" & Trim(.txtTo_Internal_cd.value)
		strVal = strVal & "&txtGigup_type=" & Trim(gigup_type1)
		strVal = strVal & "&txtStand_amt=" & Trim(.txtStand_amt.Text)				

		Call RunMyBizASP(MyBizASP, strVal)

    End With    
End Function

Function subVatDiskOK(ByVal pFileName) 
Dim strVal
    Err.Clear                                                               '��: Protect system from crashing
    If pFileName <> "" Then
	    strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0002							'��: �����Ͻ� ó�� ASP�� ���� 
	    strVal = strVal & "&txtFileName=" & pFileName							'��: ��ȸ ���� ����Ÿ	
	    Call RunMyBizASP(MyBizASP, strVal)										'��: �����Ͻ� ASP �� ���� 
	End If
End Function


Sub txtProv_type_Keypress(Key)
Dim pflag
    If Key = 13 Then
        Call FncQuery()
    End If
End Sub     

Sub txtSect_cd_Keypress(Key)
Dim pflag
    If Key = 13 Then
        Call FncQuery()
    End If
End Sub     

Sub txtOcpt_type_Keypress(Key)
Dim pflag
    If Key = 13 Then
        Call FncQuery()
    End If
End Sub     

Sub txtFr_dept_cd_Keypress(Key)
Dim pflag
    If Key = 13 Then
        Call FncQuery()
    End If
End Sub     

Sub txtto_dept_cd_Keypress(Key)
Dim pflag
    If Key = 13 Then
        Call FncQuery()
    End If
End Sub     


'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================

Function FncExit()
    FncExit = True
End Function
</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	

<SCRIPT LANGUAGE="JavaScript">
<!-- Hide script from old browsers

function setCookie(name, value, expire)
{
}
-->
</SCRIPT>
<BODY TABINDEX="-1" SCROLL="NO">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE CLASS="BatchTB2">
	<TR>
		<TD <%=HEIGHT_TYPE_00%>></TD>
	</TR>
	<TR HEIGHT=23>
		<TD>
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD>
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>������ü���ϻ���</font></td>
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
					    <FIELDSET CLASS="CLSFLD" STYLE="HEIGHT:100%">
						<TABLE <%=LR_SPACE_TYPE_40%>>   
						    <TR>
								<TD CLASS=TD5  NOWRAP>�ش���</TD>
								<TD CLASS=TD6  NOWRAP><script language =javascript src='./js/h6020ma1_txtPay_yymm_dt_txtPay_yymm_dt.js'></script></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>���ޱ���</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE="Text" ID = "txtProv_type" NAME="txtProv_type" SIZE=7 MAXLENGTH=1 tag="12XXXU" ALT="���ޱ����ڵ�"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnAcctCdFR" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(frm1.txtProv_type.value, 'PROV_TYPE')"> 
								<INPUT TYPE="Text" ID=txtProv_type_nm NAME="txtProv_type_nm" SIZE=20 MAXLENGTH=30 tag="14XXXU" ALT="���ޱ����ڵ��">
								<INPUT TYPE="HIDDEN" ID=txtProv_Oldtype NAME="txtProv_Oldtype" SIZE=20 MAXLENGTH=30 tag="14XXXU" ALT="���ޱ����ڵ��"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5>�Ա������</TD>
								<TD CLASS=TD6><script language =javascript src='./js/h6020ma1_txtYy_mm_dd_dt_txtYy_mm_dd_dt.js'></script></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>�޿�����</TD>
	                            <TD CLASS=TD6 NOWRAP><SELECT Name="txtcboPay_cd" ALT="�޿�����" CLASS=cboNormal tag="11"><OPTION Value=""></OPTION></SELECT></TD>
							</TR>			
							<TR>
							    <TD CLASS=TD5 NOWRAP>�ٹ�����</TD>
							    <TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT ID="txtSect_cd" NAME="txtSect_cd" SIZE=10  tag="11XXXU" ALT="�ٹ������ڵ�"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btn" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtSect_cd.Value, 'SECT_CD')">
							                           <INPUT TYPE=TEXT ID="txtSect_cd_nm" NAME="txtSect_cd_nm" SIZE=15  tag="14XXXU" ALT="�ٹ�������">
							                           <INPUT TYPE="HIDDEN" ID="txtSect_Oldcd" NAME="txtSect_Oldcd" SIZE=20 MAXLENGTH=30 tag="14XXXU" ALT="�ٹ������ڵ�"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>����</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT ID="txtOcpt_type" NAME="txtOcpt_type" SIZE=10 MAXLENGTH=10 ALT="����" tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btn" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtOcpt_type.Value, 'OCPT_TYPE')">
								                     <INPUT TYPE=TEXT ID="txtOcpt_type_nm" NAME="txtOcpt_type_nm" SIZE=15  tag="14XXXU" ALT="����">
							                         <INPUT TYPE="HIDDEN" ID="txtOcpt_Oldtype" NAME="txtOcpt_Oldtype" SIZE=20 MAXLENGTH=30 tag="14XXXU" ALT="����"></TD></TD>
							</TR>			
							<TR>
							    <TD CLASS=TD5 NOWRAP>�μ�</TD>
								<TD CLASS=TD6 NOWRAP><INPUT ID="txtFr_dept_cd" NAME="txtFr_dept_cd" ALT="�μ��ڵ�" TYPE="Text" SiZE="10" MAXLENGTH="10" tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSItemDC" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenDept(0)">
			                            <INPUT ID="txtFr_dept_nm" NAME="txtFr_dept_nm" ALT="�μ��ڵ��" TYPE="Text" SiZE="20" MAXLENGTH="40" tag="14XXXU">
		                                <INPUT ID="txtFr_Internal_cd" NAME="txtFr_Internal_cd" ALT="���κμ��ڵ�" TYPE="HIDDEN" SiZE="7" MAXLENGTH="7" tag="14XXXU">&nbsp;~</TD>
							</TR>			
							<TR>
								<TD CLASS=TD5 NOWRAP></TD>
								<TD CLASS=TD6 NOWRAP><INPUT ID="txtto_dept_cd" NAME="txtto_dept_cd" ALT="�μ��ڵ�" TYPE="Text" SIZE="10" MAXLENGTH="10" tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnname" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenDept(1)">
							            <INPUT ID="txtto_dept_nm" NAME="txtto_dept_nm" ALT="�μ��ڵ��" TYPE="Text"SIZE="20" MAXLENGTH="40" tag="14XXXU">
							            <INPUT ID="txtTo_Internal_cd" NAME="txtTo_Internal_cd" ALT="���κμ��ڵ�" TYPE="HIDDEN" SiZE="7" MAXLENGTH="7" tag="14XXXU"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>���޹��</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE="RADIO" CLASS="Radio" NAME=txtGigup_type ID=Rb_tot tag="12"><LABEL FOR=Rb_tot>���رݾ� ������ ������ü</LABEL></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP></TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE="RADIO" CLASS="Radio" NAME=txtGigup_type ID=Rb_dur tag="12"><LABEL FOR=Rb_dur>���رݾ� �̸� �ݾ׸� ������ü</LABEL></TD>
							</TR>			
							<TR>
								<TD CLASS=TD5 NOWRAP></TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE="RADIO" CLASS="Radio" NAME=txtGigup_type ID=Rb_dept Checked tag="12"><LABEL FOR=Rb_dept>��� �ݾ� ������ü</LABEL></TD>
							</TR>
							
	    					<TR>
              				    <TD CLASS=TD5 NOWRAP>���رݾ�</TD>
	                   			<TD CLASS=TD6><script language =javascript src='./js/h6020ma1_txtStand_amt_txtStand_amt.js'></script></TD>
	                   			
	                   	    </TR>		    		
    					</TABLE>
						</FIELDSET>
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
		                <BUTTON NAME="btnAction" CLASS="CLSMBTN" >����</BUTTON>
		            </TD>
					<TD WIDTH=* ALIGN="right">&nbsp;</TD>
					<TD WIDTH=10>&nbsp;</TD>
		        </TR>
		    </TABLE>
		</TD>
	</TR>
	<TR>
		<TD HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC = "../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
