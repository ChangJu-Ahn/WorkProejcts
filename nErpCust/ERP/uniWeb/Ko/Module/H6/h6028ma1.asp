<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          	: Human Resources
*  2. Function Name        	: Multi Sample
*  3. Program ID           	: h6028ma1_lko320 
*  4. Program Name         	: h6028ma1_lko320
*  5. Program Desc         	: �޿�����ڼ��� 
*  6. Comproxy List        	:
*  7. Modified date(First) 	: 2005/06/20
*  8. Modified date(Last)  	: 2005/06/20
*  9. Modifier (First)     	:  
* 10. Modifier (Last)      	: Lee SiNa
* 11. Comment              	: 
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
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs">   </SCRIPT>

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incHRQuery.vbs"> </SCRIPT>

<Script Language="VBScript">
Option Explicit
'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
Const BIZ_PGM_ID = "h6028mb1.asp"                                      '�����Ͻ� ���� ASP�� 
Const C_SHEETMAXROWS    = 21	                                      '�� ȭ�鿡 �������� �ִ밹��*1.5%>

'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim lsConcd

Dim C_CHK_TYPE
Dim C_EMPNO
Dim C_EMPNM
Dim C_DEPT
Dim C_DEPT_NM
Dim C_ROLL_PSTN
Dim C_PAY_CD
Dim C_PAY_GRD

Dim IsOpenPop     

'========================================================================================================
' Name : initSpreadPosVariables()	
' Desc : Initialize Column  value
'========================================================================================================

Sub initSpreadPosVariables()  
	C_CHK_TYPE	= 1
	C_EMPNO		= 2
	C_EMPNM		= 3
	C_DEPT		= 4
	C_DEPT_NM	= 5
	C_ROLL_PSTN	= 6
	C_PAY_CD	= 7
	C_PAY_GRD	= 8
End Sub
'========================================================================================================
' Name : InitVariables()	
' Desc : Initialize value
'========================================================================================================
Sub InitVariables()
	lgIntFlgMode      = parent.OPMD_CMODE						        '��: Indicates that current mode is Create mode
	lgBlnFlgChgValue  = False								    '��: Indicates that no value changed
	lgIntGrpCount     = 0										'��: Initializes Group View Size
    lgStrPrevKey      = ""                                      '��: initializes Previous Key
    lgSortKey         = 1                                       '��: initializes sort direction
End Sub

'========================================================================================================
' Name : SetDefaultVal()	
' Desc : Set default value
'========================================================================================================	
Sub SetDefaultVal()

 
 
End Sub
	
'========================================================================================================
' Name : LoadInfTB19029()	
' Desc : Set System Number format
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "H","NOCOOKIE","MA") %>
End Sub

'========================================================================================================
'	Name : CookiePage()
'	Description : Item Popup���� Return�Ǵ� �� setting
'========================================================================================================
Function CookiePage(ByVal flgs)

End Function

'========================================================================================================
' Function Name : MakeKeyStream
' Function Desc : This method set focus to pos of err
'========================================================================================================
Sub MakeKeyStream()

    lgKeyStream   = Frm1.txtFr_internal_cd.Value & parent.gColSep
    lgKeyStream   = lgKeyStream & Frm1.txtTo_internal_cd.Value & parent.gColSep
    lgKeyStream   = lgKeyStream & Frm1.txtEmpNo.Value & parent.gColSep
    lgKeyStream   = lgKeyStream & Frm1.txtFrPay_grd1.Value & parent.gColSep    
    lgKeyStream   = lgKeyStream & Frm1.txtToPay_grd1.Value & parent.gColSep        
    lgKeyStream   = lgKeyStream & Frm1.txtProv_type.Value & parent.gColSep
End Sub        

'========================================================================================================
' Name : InitData()
' Desc : Reset Combox
'========================================================================================================
Sub InitData()
	Dim intRow
	Dim strFlag
	
	With frm1.vspdData
		For intRow = 1 To .MaxRows			
 							
		Next	
	End With	
End Sub
'======================================================================================================
'   Event Name : vspdData_ComboSelChange
'   Event Desc : Combo ���� �̺�Ʈ 
'=======================================================================================================
Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)

End Sub

'========================================================================================================
' Name : InitComboBox()
' Desc : Set ComboBox
'========================================================================================================
Sub InitComboBox()
	Dim iCodeArr 
    Dim iNameArr
    
    iCodeArr = "" & chr(11)  & "Y" & chr(11) & "N" & chr(11)
    iNameArr = "��ü" & chr(11)  & "�����" & chr(11) & "������" & chr(11)

    Call SetCombo2(frm1.txtProv_type, iCodeArr, iNameArr,Chr(11))

End Sub

'========================================================================================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================================
Sub InitSpreadSheet()
	Call initSpreadPosVariables()  

	With frm1.vspdData
	
        ggoSpread.Source = frm1.vspdData
		ggoSpread.Spreadinit "V20021125",,parent.gAllowDragDropSpread    
	    .ReDraw = false
        .MaxCols = C_PAY_GRD + 1												<%'��: �ִ� Columns�� �׻� 1�� ������Ŵ %>
	    .Col = .MaxCols															<%'������Ʈ�� ��� Hidden Column%>
        .ColHidden = True
        .MaxRows = 0
		Call GetSpreadColumnPos("A")  
	
	    Call AppendNumberPlace("6","2","0")

        ggoSpread.SSSetCheck C_CHK_TYPE, "��󿩺�", 08,2
        ggoSpread.SSSetEdit  C_EMPNO            , "���", 10,,, 13, 2
        ggoSpread.SSSetEdit  C_EMPNM            , "����", 16
        ggoSpread.SSSetEdit  C_DEPT             , "�μ��ڵ�", 12
        ggoSpread.SSSetEdit  C_DEPT_NM          , "�μ�", 18
        ggoSpread.SSSetEdit  C_ROLL_PSTN        , "����", 18
        ggoSpread.SSSetEdit  C_PAY_CD			, "�޿�����", 18
        ggoSpread.SSSetEdit  C_PAY_GRD          , "��ȣ", 18
        
        Call ggoSpread.SSSetColHidden(C_DEPT,C_DEPT,True)	
	   .ReDraw = true
	
       Call SetSpreadLock 
    
    End With
    
End Sub
'========================================================================================
' Function Name : GetSpreadColumnPos
' Description   : 
'========================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
            C_CHK_TYPE		= iCurColumnPos(1)
			C_EMPNO			= iCurColumnPos(2)
			C_EMPNM			= iCurColumnPos(3)
			C_DEPT			= iCurColumnPos(4)
			C_DEPT_NM		= iCurColumnPos(5)
			C_ROLL_PSTN		= iCurColumnPos(6)
			C_PAY_CD		= iCurColumnPos(7)	
			C_PAY_GRD		= iCurColumnPos(8) 
    End Select    
End Sub
'======================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadLock()
    With frm1
    
        .vspdData.ReDraw = False

        ggoSpread.SpreadLock    C_EMPNO, -1, C_EMPNO
        ggoSpread.SpreadLock    C_EMPNM, -1, C_EMPNM
        ggoSpread.SpreadLock    C_DEPT, -1, C_DEPT         
        ggoSpread.SpreadLock    C_DEPT_NM, -1, C_DEPT_NM 
        ggoSpread.SpreadLock    C_ROLL_PSTN, -1, C_ROLL_PSTN   
        ggoSpread.SpreadLock    C_PAY_CD, -1, C_PAY_CD   
        ggoSpread.SpreadLock    C_PAY_GRD, -1, C_PAY_GRD   
                                        
   	    ggoSpread.SSSetProtected	.vspdData.MaxCols,-1,-1 
        .vspdData.ReDraw = True

    End With
End Sub

'======================================================================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadColor(ByVal pvStartRow,ByVal pvEndRow)
    With frm1
    
       .vspdData.ReDraw = False
         ggoSpread.SSSetProtected		C_EMPNO, pvStartRow, pvEndRow
         ggoSpread.SSSetProtected		C_EMPNM, pvStartRow, pvEndRow
         ggoSpread.SSSetProtected		C_DEPT, pvStartRow, pvEndRow         
         ggoSpread.SSSetProtected		C_DEPT_NM, pvStartRow, pvEndRow
         ggoSpread.SSSetProtected		C_ROLL_PSTN, pvStartRow, pvEndRow
         ggoSpread.SSSetProtected		C_PAY_CD, pvStartRow, pvEndRow
         ggoSpread.SSSetProtected		C_PAY_GRD, pvStartRow, pvEndRow
         
       .vspdData.ReDraw = True
    
    End With
End Sub

'======================================================================================================
' Function Name : SubSetErrPos
' Function Desc : This method set focus to pos of err
'======================================================================================================
Sub SubSetErrPos(iPosArr)
    Dim iDx
    Dim iRow
    iPosArr = Split(iPosArr,parent.gColSep)
    If IsNumeric(iPosArr(0)) Then
       iRow = CInt(iPosArr(0))
       For iDx = 1 To  frm1.vspdData.MaxCols - 1
           Frm1.vspdData.Col = iDx
           Frm1.vspdData.Row = iRow
           If Frm1.vspdData.ColHidden <> True And Frm1.vspdData.BackColor <> parent.UC_PROTECTED Then
              Frm1.vspdData.Col = iDx
              Frm1.vspdData.Row = iRow
              Frm1.vspdData.Action = 0 ' go to 
              Exit For
           End If
           
       Next
          
    End If   
End Sub

'========================================================================================================
' Name : Form_Load
' Desc : developer describe this line Called by Window_OnLoad() evnt
'========================================================================================================
Sub Form_Load()
    
    Err.Clear                                                                       '��: Clear err status
	Call LoadInfTB19029                                                             '��: Load table , B_numeric_format
 		
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
	Call ggoOper.LockField(Document, "N")											'��: Lock Field
  
    Call InitSpreadSheet                                                            'Setup the Spread sheet
    Call InitVariables                                                              'Initializes local global variables
    
    Call FuncGetAuth(gStrRequestMenuID , parent.gUsrID, lgUsrIntCd)     ' �ڷ����:lgUsrIntCd ("%", "1%")
    
    Call SetDefaultVal
    Call InitComboBox
  
    Call SetToolbar("1100000000001111")										        '��ư ���� ���� 
       
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
    Dim strFrDept, strToDept
    
    FncQuery = False															 '��: Processing is NG
    Err.Clear                                                                    '��: Clear err status
 
    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO,"x","x")					 '��: Data is changed.  Do you want to display it? 
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If   

    ggoSpread.Source = Frm1.vspdData    
	ggoSpread.ClearSpreadData      															
    If Not chkField(Document, "1") Then									         '��: This function check required field
       Exit Function
    End If
    
    Dim Fr_dept_cd , To_dept_cd, rFrDept ,rToDept
    
    Fr_dept_cd = frm1.txtFrDept.value
    To_dept_cd = frm1.txtToDept.value

    If txtFrDept_Onchange() Then
        Exit Function
    End if    
    If txtToDept_Onchange() Then
        Exit Function
    End if    
    If txtEmpNo_Onchange() Then
        Exit Function
    End if  
   
    If txtFrPay_grd1_OnChange() Then          'enter key �� ��ȸ�� ����� check�� �ش���� ������ query����...
        Exit Function
    End if

    If txtToPay_grd1_OnChange() Then          'enter key �� ��ȸ�� ����� check�� �ش���� ������ query����...
        Exit Function
    End if   

    If fr_dept_cd = "" then    
        IntRetCd = FuncGetTermDept(lgUsrIntCd ,"", rFrDept ,rToDept)
		frm1.txtFr_internal_cd.value = rFrDept				
		frm1.txtFrDeptNm.value = ""
	End If	

	If to_dept_cd = "" then
        IntRetCd = FuncGetTermDept(lgUsrIntCd ,"", rFrDept ,rToDept)
		frm1.txtTo_internal_cd.value = rToDept
		frm1.txtToDeptNm.value = ""
	End If  

    Fr_dept_cd = frm1.txtFr_internal_cd.value
    To_dept_cd = frm1.txtTo_internal_cd.value
    
    If (Fr_dept_cd<> "") AND (To_dept_cd<>"") Then       
        If Fr_dept_cd > To_dept_cd then
	        Call DisplayMsgBox("800153","X","X","X")	'���ۺμ��ڵ�� ����μ��ڵ庸�� �۾ƾ��մϴ�.
            frm1.txtFr_internal_cd.value = ""
            frm1.txtTo_internal_cd.value = ""
            frm1.txtFrDept.focus
            Set gActiveElement = document.activeElement
            Exit Function
        End IF 
    END IF 

    Call InitVariables                                                        '��: Initializes local global variables
    Call MakeKeyStream()

    Call DisableToolBar(parent.TBC_QUERY)
	If DbQuery = False Then
		Call RestoreToolBar()
        Exit Function
    End If

    FncQuery = True                                                              '��: Processing is OK

End Function

'========================================================================================================
'	Name : OpenDept()
'	Description : Dept PopUp
'========================================================================================================
Function OpenDept(iWhere)
	Dim arrRet
	Dim arrParam(2)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	If iWhere = 0 Then 'TextBox(Condition)
		arrParam(0) = UCase(Trim(frm1.txtfrDept.value))			' from ���Ǻο��� ���� ��� Code Condition
	Else 
		arrParam(0) = UCase(Trim(frm1.txttoDept.value))			' to ���Ǻο��� ���� ��� Code Condition
	End If
	
	arrParam(1) = ""
	arrParam(2) = lgUsrIntcd
	
	arrRet = window.showModalDialog(HRAskPRAspName("DeptPopupDt"), Array(window.parent,arrParam), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		If iWhere = 0 Then 
			frm1.txtfrDept.focus
		Else 
			frm1.txttoDept.focus
		End If	
		Exit Function
	Else
		If iWhere = 0 Then 
			frm1.txtfrDept.value = arrRet(0)
			frm1.txtfrDeptNm.value = arrRet(1)
			frm1.txtfrDept.focus
		Else 
			frm1.txttoDept.value = arrRet(0)
			frm1.txttoDeptNm.value = arrRet(1)
			frm1.txttoDept.focus
		End If	
	End If	
			
End Function
'========================================================================================================
'   Event Name : txtFrDept_Onchange
'   Event Desc :
'========================================================================================================
Function txtFrDept_Onchange()
    Dim IntRetCd,Dept_Nm,Internal_cd
   
    If frm1.txtFrDept.value = "" Then
		frm1.txtFrDeptNm.value = ""
		frm1.txtFr_internal_cd.value = ""
    Else
        IntRetCd = FuncDeptName(frm1.txtFrDept.value , "" , lgUsrIntCd,Dept_Nm , Internal_cd)
    
	    If  IntRetCd < 0 then
	        If  IntRetCd = -1 then
    			Call DisplayMsgBox("800098","X","X","X")	'�μ������ڵ忡 ��ϵ��� ���� �ڵ��Դϴ�.
            Else
                Call DisplayMsgBox("800454","X","X","X")	'�ڷῡ ���� ������ �����ϴ�.
            End if
		    frm1.txtFrDeptNm.value = ""
		    frm1.txtFr_internal_cd.value = ""
            frm1.txtFrDept.focus
            Set gActiveElement = document.ActiveElement 
            txtFrDept_Onchange = true
            Exit Function      
        Else
			frm1.txtFrDeptNm.value = Dept_Nm
		    frm1.txtFr_internal_cd.value = Internal_cd
        End if 
    End if  
    
End Function

'========================================================================================================
'   Event Name : txtToDept_Onchange
'   Event Desc :
'========================================================================================================
Function txtToDept_Onchange()
    Dim IntRetCd,Dept_Nm,Internal_cd
    
    If frm1.txtToDept.value = "" Then
		frm1.txtToDeptNm.value = ""
		frm1.txtTo_internal_cd.value = ""
    Else
        IntRetCd = FuncDeptName(frm1.txtToDept.value , "" , lgUsrIntCd,Dept_Nm , Internal_cd)
	    If  IntRetCd < 0 then
	        If  IntRetCd = -1 then
    			Call DisplayMsgBox("800098","X","X","X")	'�μ������ڵ忡 ��ϵ��� ���� �ڵ��Դϴ�.
            Else
                Call DisplayMsgBox("800454","X","X","X")	'�ڷῡ ���� ������ �����ϴ�.
            End if
		    frm1.txtToDeptNm.value = ""
		    frm1.txtTo_internal_cd.value = ""
            frm1.txtToDept.focus
            Set gActiveElement = document.ActiveElement 
            txtToDept_Onchange = true
            Exit Function      
        Else
			frm1.txtToDeptNm.value = Dept_Nm
		    frm1.txtTo_internal_cd.value = Internal_cd
        End if 
    End if  
    
End Function

'========================================================================================================
'   Event Name : txtEmpNo_OnChange
'   Event Desc :
'========================================================================================================
Function txtEmpNo_Onchange()
    Dim IntRetCd
    Dim strName
    Dim strDept_nm
    Dim strRoll_pstn
    Dim strPay_grd1
    Dim strPay_grd2
    Dim strEntr_dt
    Dim strInternal_cd

    If  frm1.txtEmpNo.value = "" Then
		frm1.txtEmpNm.value = ""
    Else
	    IntRetCd = FuncGetEmpInf2(frm1.txtEmpNo.value,lgUsrIntCd,strName,strDept_nm,_
	                strRoll_pstn, strPay_grd1, strPay_grd2, strEntr_dt, strInternal_cd)
	                
	    If  IntRetCd < 0 then
	        If  IntRetCd = -1 then
    			Call DisplayMsgBox("800048","X","X","X")	'�ش����� �������� �ʽ��ϴ�.
            Else
                Call DisplayMsgBox("800454","X","X","X")	'�ڷῡ ���� ������ �����ϴ�.
            End if
			frm1.txtEmpNm.value = ""
            Frm1.txtEmpNo.focus 
            Set gActiveElement = document.ActiveElement
			txtEmpNo_Onchange = true
        Else
			frm1.txtEmpNm.value = strName
        End if 
    End if  
End Function

'========================================================================================================
'	Name : OpenEmp()
'	Description : Employee PopUp
'========================================================================================================
Function OpenEmp(iWhere)
	Dim arrRet
	Dim arrParam(2)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	If iWhere = 0 Then 'TextBox(Condition)
		arrParam(0) = UCase(Trim(frm1.txtEmpNo.value))			<%' Code Condition%>
		arrParam(1) = ""'frm1.txtEmpNm.value		    ' Name Cindition
	End If
	
	arrParam(2) = lgUsrIntcd
	
	arrRet = window.showModalDialog(HRAskPRAspName("EmpPopup"), Array(window.parent,arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		If iWhere = 0 Then 'TextBox(Condition)
			frm1.txtEmpNo.focus
		End If	
		Exit Function
	Else
		With frm1
			If iWhere = 0 Then 'TextBox(Condition)
				.txtEmpNo.value = arrRet(0)
				.txtEmpNm.value = arrRet(1)
				.txtEmpNo.focus
			End If
		End With
	End If	
			
End Function

'========================================================================================================
' Name : FncDelete
' Desc : developer describe this line Called by MainDelete in Common.vbs
'========================================================================================================
Function FncDelete()
End Function

'========================================================================================================
' Name : FncSave
' Desc : developer describe this line Called by MainSave in Common.vbs
'========================================================================================================
Function FncSave()
    Dim IntRetCD 
    Dim lRow

    FncSave = False                                                              '��: Processing is NG
    Err.Clear                                                                    '��: Clear err status
    
    ggoSpread.Source = frm1.vspdData
    If  ggoSpread.SSCheckChange = False Then
        IntRetCD =  DisplayMsgBox("900001","x","x","x")                           '��:There is no changed data. 
        Exit Function
    End If
    
     ggoSpread.Source = frm1.vspdData
    If Not  ggoSpread.SSDefaultCheck Then                                         '��: Check contents area
       Exit Function
    End If
    
    If Not chkField(Document, "1") Then									         '��: This function check required field
       Exit Function
    End If

    'Call MakeKeyStream("X")
           
    If DbSave = False Then
        Exit Function
    End If
            
    FncSave = True                                                              '��: Processing is OK
    
End Function
'========================================================================================================
' Name : DbSave
' Desc : This function is called by FncSave
'========================================================================================================
Function DbSave()
    Dim lRow        
    Dim lGrpCnt     
    Dim retVal      
	Dim strVal, strDel
    Err.Clear                                                                    '��: Clear err status
		
	DbSave = False														         '��: Processing is NG
		
	If LayerShowHide(1) = False Then
			Exit Function
	End If
		
	With frm1
		.txtMode.value        =  parent.UID_M0002                                        '��: Delete
		.txtFlgMode.value     = lgIntFlgMode
        .txtKeyStream.Value   = lgKeyStream                                      '��: Save Key
	End With

    strVal = ""
    strDel = ""
    lGrpCnt = 1

	With Frm1
    
       For lRow = 1 To .vspdData.MaxRows
    
           .vspdData.Row = lRow
           .vspdData.Col = 0

           Select Case .vspdData.Text
               Case  ggoSpread.UpdateFlag                                      '��: Update
                                                   strVal = strVal & "U" & parent.gColSep
                                                   strVal = strVal & lRow & parent.gColSep
                    .vspdData.Col = C_EMPNO		 : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep     
                    .vspdData.Col = C_CHK_TYPE  : strVal = strVal & Trim(.vspdData.Text) & parent.gRowSep                      
                    lGrpCnt = lGrpCnt + 1
           End Select
       Next
	
	End With

    Frm1.txtMaxRows.value = lGrpCnt-1	
	Frm1.txtSpread.value = strDel & strVal

	Call ExecMyBizASP(frm1, BIZ_PGM_ID)
		
    DbSave  = True                                                               '��: Processing is NG
End Function
'========================================================================================================
' Function Name : DbSaveOk
' Function Desc : Called by MB Area when save operation is successful
'========================================================================================================
Function DbSaveOk()
    lgBlnFlgChgValue = False
	Call InitVariables
     ggoSpread.Source = Frm1.vspdData
    Frm1.vspdData.MaxRows = 0
    ggoSpread.ClearSpreadData
    Call MainQuery()
End Function
	
'========================================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================================
Function FncCopy()
End Function


'========================================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================================
Function FncCancel() 
    ggoSpread.Source = frm1.vspdData	
    ggoSpread.EditUndo  
End Function

'========================================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================================
Function FncInsertRow(ByVal pvRowCnt) 
	
End Function

'========================================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================================
Function FncDeleteRow() 
End Function

'========================================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================================
Function FncPrint()
    Call parent.FncPrint()
End Function

'========================================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================================
Function FncExcel() 
    Call parent.FncExport(parent.C_MULTI)                                         '��: ȭ�� ���� 
End Function

'========================================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================================
Function FncFind() 
    Call parent.FncFind(parent.C_MULTI, False)                                    '��:ȭ�� ����, Tab ���� 
End Function

'========================================================================================
' Function Name : FncSplitColumn
' Function Desc : 
'========================================================================================
Sub FncSplitColumn()

    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   

End Sub

'========================================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================================
Function FncExit()
    Dim IntRetCD
	
	FncExit = False
	
    ggoSpread.Source = frm1.vspdData	
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO,"x","x")			'��: Data is changed.  Do you want to exit? 
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    FncExit = True
End Function
'========================================================================================
' Function Name : PopSaveSpreadColumnInf
' Description   : 
'========================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

'========================================================================================
' Function Name : PopRestoreSpreadColumnInf
' Description   : 
'========================================================================================
Sub PopRestoreSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()      
	Call ggoSpread.ReOrderingSpreadData()
	Call InitData()
End Sub

'========================================================================================================
' Name : DbQuery
' Desc : This function is called by FncQuery
'========================================================================================================
Function DbQuery() 

    DbQuery = False
    
    Err.Clear                                                                        '��: Clear err status
 
	If LayerShowHide(1) = False then
    		Exit Function 
    End if
	
	Dim strVal
 
    With Frm1
		strVal = BIZ_PGM_ID & "?txtMode="            & parent.UID_M0001						         
        strVal = strVal     & "&txtKeyStream="       & lgKeyStream                       '��: Query Key
        strVal = strVal     & "&txtMaxRows="         & .vspdData.MaxRows
        strVal = strVal     & "&lgStrPrevKey=" & lgStrPrevKey                 '��: Next key tag
    End With
 	
    If lgIntFlgMode = parent.OPMD_UMODE Then
    Else
    End If
 
	Call RunMyBizASP(MyBizASP, strVal)                                               '��: Run Biz Logic

    DbQuery = True
    
End Function

'========================================================================================================
' Name : DbDelete
' Desc : This function is called by FncDelete
'========================================================================================================
Function DbDelete()

End Function
'========================================================================================================
' Function Name : DbQueryOk
' Function Desc : Called by MB Area when query operation is successful
'========================================================================================================
Function DbQueryOk()	
	Dim intRow	
	Dim strt_dt,end_dt											     
    lgIntFlgMode = parent.OPMD_UMODE
       
    Call ggoOper.LockField(Document, "Q")										'��: Lock field
    Call InitData()
    
	With frm1.vspdData	    
		For intRow = 1 To .MaxRows			
	   		.Row = intRow
 
		Next	    
    End With
  
	Call SetToolbar("1100100000001111")									
	frm1.vspdData.focus	
End Function

'===========================================================================
' Function Name : OpenPayGrd
' Function Desc : OpenPayGrd Reference Popup
'===========================================================================
Function OpenPayGrd(iWhere)

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True


	arrParam(1) = "B_minor"				            	' TABLE ��Ī 
	
	If iWhere = 0 Then
		arrParam(2) = Trim(frm1.txtFrPay_grd1.Value)	        ' Code Condition
	Else 	
		arrParam(2) = Trim(frm1.txtToPay_grd1.Value)	        ' Code Condition
	End If		
	arrParam(3) = ""									' Name Cindition
	arrParam(4) = "major_cd=" & FilterVar("H0001", "''", "S") & ""		    		' Where Condition
	arrParam(5) = "��ȣ"		    				    ' TextBox ��Ī 
	
	arrField(0) = "minor_cd"							' Field��(0)
	arrField(1) = "minor_nm"    						' Field��(1)%>
    
	arrHeader(0) = "��ȣ�ڵ�"			        		' Header��(0)%>
	arrHeader(1) = "��ȣ��"	        					' Header��(1)%>

    arrParam(3) = ""	
	arrParam(0) = arrParam(5)								    ' �˾� ��Ī 

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		If iWhere = 0 Then 	
			frm1.txtFrPay_grd1.focus
		Else 
			frm1.txtToPay_grd1.focus	
		End If					
		Exit Function
	Else
		If iWhere = 0 Then 	
			frm1.txtFrPay_grd1.value = arrRet(0)
			frm1.txtFrPay_grd1_nm.value = arrRet(1)  
			frm1.txtFrPay_grd1.focus
		Else
			frm1.txtToPay_grd1.value = arrRet(0)
			frm1.txtToPay_grd1_nm.value = arrRet(1)  
			frm1.txtToPay_grd1.focus		 
		End If						
	End If	
	
End Function

'========================================================================================================
'   Event Name : txtFrPay_grd1_OnChange 
'   Event Desc :
'========================================================================================================
Function txtFrPay_grd1_OnChange()

    If  frm1.txtFrPay_grd1.value = "" Then
        frm1.txtFrPay_grd1_nm.value = ""
        frm1.txtFrPay_grd1.focus
        Set gActiveElement = document.ActiveElement
    Else
  
        if   CommonQueryRs(" MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0001", "''", "S") & " AND MINOR_CD =  " & FilterVar(frm1.txtFrPay_grd1.value , "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = false then
            frm1.txtFrPay_grd1_nm.value = ""
            
            Call  DisplayMsgBox("970000", "x","��ȣ�ڵ�","x")
	        frm1.txtFrPay_grd1.focus
	        Set gActiveElement = document.ActiveElement
			txtFrPay_grd1_OnChange = true
			Exit Function				       
	    Else
	    
	        frm1.txtFrPay_grd1_nm.value = Replace(lgF0, Chr(11), "")
	    End If
    End If
		    
End Function

'========================================================================================================
'   Event Name : txtToPay_grd1_OnChange 
'   Event Desc :
'========================================================================================================
Function txtToPay_grd1_OnChange()

    If  frm1.txtToPay_grd1.value = "" Then
        frm1.txtToPay_grd1_nm.value = ""
        frm1.txtToPay_grd1.focus
        Set gActiveElement = document.ActiveElement
    Else
        if   CommonQueryRs(" MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0001", "''", "S") & " AND MINOR_CD =  " & FilterVar(frm1.txtToPay_grd1.value , "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = false then
            frm1.txtToPay_grd1_nm.value = ""
            Call  DisplayMsgBox("970000", "x","��ȣ�ڵ�","x")
	        frm1.txtToPay_grd1.focus
	        Set gActiveElement = document.ActiveElement
			txtToPay_grd1_OnChange = true	        
			Exit Function
	    Else
	        frm1.txtToPay_grd1_nm.value = Replace(lgF0, Chr(11), "")
	    End If
    End If
	    
End Function

'========================================================================================================
'   Event Name : vspdData_Change 
'   Event Desc :
'========================================================================================================
Function vspdData_Change(ByVal Col , ByVal Row )

   	Frm1.vspdData.Row = Row
   	Frm1.vspdData.Col = Col

    Select Case Col
         Case  C_CHK_TYPE

' 			Frm1.vspdData.Col = 0
'			Frm1.vspdData.text = "1"
		
    End Select    
             
   	If Frm1.vspdData.CellType = parent.SS_CELL_TYPE_FLOAT Then
      If CDbl(Frm1.vspdData.text) < CDbl(Frm1.vspdData.TypeFloatMin) Then
         Frm1.vspdData.text = Frm1.vspdData.TypeFloatMin
      End If
	End If
	
	ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row
End Function

'========================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : �÷��� Ŭ���� ��� �߻� 
'========================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)

	Call SetPopupMenuItemInf("1101111111")       

    gMouseClickStatus = "SPC"   
    Set gActiveSpdSheet = frm1.vspdData
    
    If frm1.vspdData.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If  
    If Row = 0 Then
        ggoSpread.Source = frm1.vspdData
        If lgSortKey = 1 Then
            ggoSpread.SSSort
            lgSortKey = 2
        Else
            ggoSpread.SSSort ,lgSortKey
            lgSortKey = 1
        End If
    End If      	
    frm1.vspdData.Row = Row   	
   	
   	
End Sub
'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

End Sub

'========================================================================================================
'   Event Name : vspdData_DblClick
'   Event Desc : 
'========================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)				
    Dim iColumnName
    
    If Row <= 0 Then
		Exit Sub
	End if
	
	If frm1.vspdData.MaxRows = 0 then
		Exit Sub
	End if
End Sub
'========================================================================================================
'   Event Name : vspdData_GotFocus
'   Event Desc : This event is spread sheet data changed
'========================================================================================================
Sub vspdData_GotFocus()
    ggoSpread.Source = Frm1.vspdData
End Sub
'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
End Sub
'========================================================================================================
'   Event Name : vspdData_MouseDown
'   Event Desc : 
'========================================================================================================

Sub vspdData_MouseDown(Button , Shift , x , y)

       If Button = 2 And gMouseClickStatus = "SPC" Then
          gMouseClickStatus = "SPCR"
        End If
End Sub    

'========================================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
 
 	
End Sub

'======================================================================================================
'	Name : ButtonClicked()
'=======================================================================================================
Sub ButtonClicked(Byval ButtonDown)
	Dim intRow
	Dim chk
	Dim strt_dt,end_dt
	
	With frm1.vspdData	
	
	Select Case ButtonDown
		Case 1
			chk = "1"
			For intRow = 1 To .MaxRows			
   				.Row = intRow
				.Col = C_CHK_TYPE
				If .Text = "1" Then
					chk = "0"
					ggoSpread.Source = frm1.vspdData
					ggoSpread.UpdateRow intRow					
'					Exit For
				End If
			Next		
		
			For intRow = 1 To .MaxRows	
   				.Row = intRow			
 
				If 1=1 Then

					.Col = C_CHK_TYPE
					.Text = chk		
					ggoSpread.Source = frm1.vspdData					
					ggoSpread.UpdateRow intRow				
				End If	
			Next
	End Select 
	End With

End Sub


</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>

<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR>
		<TD <%=HEIGHT_TYPE_00%>></TD>
	</TR>
	<!-- space Area-->

	<TR HEIGHT=23>
		<TD WIDTH="100%">
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><IMG src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23" ></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>�ӱ����޴���ڼ���</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><IMG src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23" ></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* ALIGN=LIGHT>&nbsp;</TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
		
    <TR HEIGHT=*>
		<TD width=100% CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
			    <TR>
			        <TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
			    </TR>
				<TR>
					<TD HEIGHT=20 WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
						<TABLE <%=LR_SPACE_TYPE_40%>>
							<TR>
								<TD CLASS="TD5" NOWRAP>�μ�</TD>
								<TD CLASS="TD6" NOWRAP><INPUT NAME="txtFrDept" ALT="�μ��ڵ�" TYPE="Text" SiZE=10 MAXLENGTH=10 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSItemDC" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenDept(0)">
			                                    <INPUT NAME="txtFrDeptNm" ALT="�μ��ڵ��" TYPE="Text" SiZE=20 MAXLENGTH=40 tag="14XXXU">&nbsp;~
		                                        <INPUT NAME="txtFr_Internal_cd" ALT="���κμ��ڵ�" TYPE="HIDDEN" SiZE=7 MAXLENGTH=7 tag="14XXXU"></TD>
					    		<TD CLASS="TD5" NOWRAP>��ȣ</TD>
					    		<TD CLASS="TD6"><INPUT NAME="txtFrPay_grd1" ALT="��ȣ����" TYPE="Text" MAXLENGTH=2 SiZE=5 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSItemDC" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenPayGrd(0)">&nbsp;<INPUT NAME="txtFrPay_grd1_nm" TYPE="Text" MAXLENGTH=20 SIZE=20 tag="24">&nbsp;~</TD>
		                                        
							</TR>
							<TR>
					    		<TD CLASS="TD5" NOWRAP></TD>					    		
								<TD CLASS="TD6" NOWRAP><INPUT NAME="txtToDept" MAXLENGTH="10" SIZE=10 ALT ="Order ID" tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnname" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenDept(1)">
							                    <INPUT NAME="txtToDeptNm" MAXLENGTH="40" SIZE=20 ALT ="Order ID" tag="14XXXU">
    			                                <INPUT NAME="txtTo_Internal_cd" ALT="���κμ��ڵ�" TYPE="HIDDEN" SiZE=7 MAXLENGTH=7 tag="14XXXU"></TD>
							    <TD CLASS="TD5" NOWRAP></TD>
								<TD CLASS="TD6" NOWRAP><INPUT NAME="txtToPay_grd1" ALT="��ȣ����" TYPE="Text" MAXLENGTH=2 SiZE=5 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSItemDC" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenPayGrd(1)">&nbsp;<INPUT NAME="txtToPay_grd1_nm" TYPE="Text" MAXLENGTH=20 SIZE=20 tag="24"></TD>
							</TR>
							<TR>
							    <TD CLASS="TD5" NOWRAP>���</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtEmpNo" SIZE=13 MAXLENGTH=13 tag="11XXXU" ALT="���"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCountryCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenEmp(0)">
									            <INPUT TYPE=TEXT NAME="txtEmpNm" tag="14XXXU">
							    <TD CLASS="TD5" NOWRAP>����ڿ���</TD>
								<TD CLASS="TD6" NOWRAP><SELECT Name="txtProv_type" ALT="����ڿ���" CLASS ="cbonormal" tag="11"></TD>
							</TR>
						</TABLE>
						</FIELDSET>
					</TD>
				</TR>	
				<TR>
				    <TD <%=HEIGHT_TYPE_03%>></TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=* VALIGN=TOP>
						<TABLE <%=LR_SPACE_TYPE_20%> >
							<TR>
								<TD HEIGHT=100% WIDTH=100% >
									<script language =javascript src='./js/h6028ma1_vaSpread_vspdData.js'></script>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=20>
	    <TD>
	        <TABLE <%=LR_SPACE_TYPE_30%>>
	            <TR>
	                <TD WIDTH=10>&nbsp;</TD>
	                <TD><BUTTON NAME="btnCb_autoisrt" CLASS="CLSMBTN" ONCLICK="VBScript: ButtonClicked('1')" flag=1>�ϰ�����/����</BUTTON></TD>
	                <TD WIDTH=* ALIGN="right"></TD>
	                <TD WIDTH=10>&nbsp;</TD>
	            </TR>
	        </TABLE>
	    </TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=0><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=0 FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode"        TAG="24">
<INPUT TYPE=HIDDEN NAME="txtKeyStream"   TAG="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode"     TAG="24">
<INPUT TYPE=HIDDEN NAME="txtPrevNext"    TAG="24">
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA><INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24"><%' ����ó��ASP�� �ѱ�� ���� ������ ��� �ִ� Tag�� %>
</FORM>

<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

