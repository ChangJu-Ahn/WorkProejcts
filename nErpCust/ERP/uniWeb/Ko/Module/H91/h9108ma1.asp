<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : Human Resources
*  2. Function Name        : ����������� 
*  3. Program ID           : h9108ma1.asp
*  4. Program Name         : h9108ma1.asp
*  5. Program Desc         : �����ڷ��� 
*  6. Modified date(First) : 2001/06/09
*  7. Modified date(Last)  : 2003/06/13
*  8. Modifier (First)     : Bong-kyu Song
*  9. Modifier (First)     : Lee SiNa
* 10. Comment              :
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
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incHRQuery.vbs"> </SCRIPT>

<Script Language="VBScript">
Option Explicit

'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
Const WARRANT_TYPE_MAJOR = "S0002"
Const DEL_TYPE_MAJOR     = "S0003"
Const BIZ_PGM_ID      = "h9108mb1.asp"						           '��: Biz Logic ASP Name
Const BIZ_PGM_ID2     = "h9108bb1.asp"						           '��: Biz Logic ASP Name
Const BIZ_PGM_ID3     = "h9108bb2.asp"						           '��: Biz Logic ASP Name
Const BIZ_PGM_ID4     = "h9108bb3.asp"						           '��: Biz Logic ASP Name

'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim IsOpenPop          

Dim gSelframeFlg                                                       '���� TAB�� ��ġ�� ��Ÿ���� Flag %>
Dim gblnWinEvent                                                       'ShowModal Dialog(PopUp) Window�� ���� �� �ߴ� ���� �����ϱ� ���� 
Dim lgBlnFlawChgFlg	
Dim gtxtChargeType
Dim lsInternal_cd

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

	gblnWinEvent      = False
	lgBlnFlawChgFlg   = False
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
Sub MakeKeyStream(pOpt)
    lgKeyStream       = Frm1.txtYear.Year & parent.gColSep		            'You Must append one character(parent.gColSep)
    lgKeyStream       = lgKeyStream       & Frm1.txtEmp_no.Value & parent.gColSep         'You Must append one character(parent.gColSep)
    lgKeyStream       = lgKeyStream       & lgUsrIntCd & parent.gColSep ' �ڷ���� 
End Sub        

'========================================================================================================
' Name : Form_Load
' Desc : developer describe this line Called by Window_OnLoad() evnt
'========================================================================================================
Sub Form_Load()

    Dim strYear, strMonth, strDay
    Err.Clear                                                                       '��: Clear err status
	Call LoadInfTB19029                                                             '��: Load table , B_numeric_format
		
	Call AppendNumberPlace("7", "3", "0")
	  
	Call ggoOper.FormatField(Document, "A",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.FormatDate(frm1.txtYear, parent.gDateFormat, 3)	

	Call ggoOper.LockField(Document, "N")											'��: Lock Field
	
	Call SetToolbar("1100000000011111")												'��: Set ToolBar
	
	Call InitVariables

    Call FuncGetAuth(gStrRequestMenuID, parent.gUsrID, lgUsrIntCd)                                ' �ڷ����:lgUsrIntCd ("%", "1%")

    Call ExtractDateFrom("<%=GetSvrDate%>", parent.gServerDateFormat, parent.gServerDateType, strYear, strMonth, strDay)
	frm1.txtYear.focus
	frm1.txtYear.Year = strYear
 
	Call CookiePage (0)                                                             '��: Check Cookie
			
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
    Dim RetStatus
    Dim strName
    Dim strDept_nm
    Dim strRoll_pstn
    Dim strPay_grd1
    Dim strPay_grd2
    Dim strEntr_dt
    
    FncQuery = False															 '��: Processing is NG
    Err.Clear                                                                    '��: Clear err status
    If txtEmp_no_Onchange() Then         'ENTER KEY �� ��ȸ�� ����� ����� CHECK �Ѵ� 
        Exit Function
    End if

    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO,"x","x")					 '��: Data is changed.  Do you want to display it? 
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
   
    Call ggoOper.ClearField(Document, "2")										 '��: Clear Contents  Field
    
    If Not chkField(Document, "1") Then									         '��: This function check required field
       Exit Function
    End If

    Call InitVariables                                                           '��: Initializes local global variables
    Call MakeKeyStream("X")

    If DbQuery = False Then  
		Exit Function
	End If                                                                 '��: Query db data

    FncQuery = True                                                              '��: Processing is OK

End Function
	
'========================================================================================================
' Name : FncNew
' Desc : developer describe this line Called by MainNew in Common.vbs
'========================================================================================================
Function FncNew()
    Dim IntRetCD 
    
    FncNew = False																 '��: Processing is NG
    Err.Clear                                                                    '��: Clear err status
    
    If lgBlnFlgChgValue = True Then
       IntRetCD = DisplayMsgBox("900015", parent.VB_YES_NO,"x","x")					 '��: Data is changed.  Do you want to make it new? 
       If IntRetCD = vbNo Then
          Exit Function
       End If
    End If
    
    Call ggoOper.ClearField(Document, "A")                                       '��: Clear Condition Field
    Call ggoOper.LockField(Document , "N")                                       '��: Lock  Field
    
    Call SetToolbar("11001000000111")

    Call InitVariables                                                        '��: Initializes local global variables
    
    Set gActiveElement = document.ActiveElement   
    
    FncNew = True																 '��: Processing is OK
End Function
	
'========================================================================================================
' Name : FncDelete
' Desc : developer describe this line Called by MainDelete in Common.vbs
'========================================================================================================
Function FncDelete()
    Dim intRetCD
    
    FncDelete = False                                                            '��: Processing is NG
    Err.Clear        
                                                                '��: Clear err status
    If lgIntFlgMode <> parent.OPMD_UMODE Then                                           '��: Please do Display first. 
        Call DisplayMsgBox("900002","x","x","x")                                
        Exit Function
    End If
    
    IntRetCD = DisplayMsgBox("900003", parent.VB_YES_NO,"x","x")                        '��: Do you want to delete? 
	If IntRetCD = vbNo Then
		Exit Function
	End If
    
    Call MakeKeyStream("D")
    If DbDelete = False Then
		Exit Function
	End If

    Set gActiveElement = document.ActiveElement   
    
    FncDelete = True                                                            '��: Processing is OK
End Function

'========================================================================================================
' Name : FncSave
' Desc : developer describe this line Called by MainSave in Common.vbs
'========================================================================================================
Function FncSave()
    Dim IntRetCD 
    
    FncSave = False                                                              '��: Processing is NG
    
    Err.Clear                                                                    '��: Clear err status

    If lgBlnFlgChgValue = False Then 
        IntRetCD = DisplayMsgBox("900001","x","x","x")                           '��:There is no changed data. 
        Exit Function
    End If
    
    If Not chkField(Document, "2") Then                                          '��: Check contents area
       Exit Function
    End If
    
    If Not chkField(Document, "1") Then									         '��: This function check required field
       Exit Function
    End If

    Call MakeKeyStream("X")
    If DbSave = False Then
		Exit Function
	End If
    
    FncSave = True                                                              '��: Processing is OK
End Function

'========================================================================================================
' Name : FncCopy
' Desc : developer describe this line Called by MainSave in Common.vbs
' Keep : Make sure to clear primary key area
'========================================================================================================
Function FncCopy()
	Dim IntRetCD

    FncCopy = False                                                              '��: Processing is NG
    Err.Clear                                                                    '��: Clear err status
	
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900017", parent.VB_YES_NO,"x","x")				     '��: Data is changed.  Do you want to continue? 
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    lgIntFlgMode = parent.OPMD_CMODE												     '��: Indicates that current mode is Crate mode
    
    Call ggoOper.ClearField(Document, "1")                                       '��: Clear Condition Field
    Call ggoOper.LockField(Document, "N")									     '��: This function lock the suitable field
    Call SetToolbar("11001000000011")

    Set gActiveElement = document.ActiveElement   
    FncCopy = True                                                            '��: Processing is OK
    
End Function

'========================================================================================================
' Name : FncCancel
' Desc : developer describe this line Called by MainCancel in Common.vbs
'========================================================================================================
Function FncCancel() 
	On Error Resume Next                                                      '��: Protect system from crashing
End Function

'========================================================================================================
' Name : FncInsertRow
' Desc : developer describe this line Called by MainInsertRow in Common.vbs
'========================================================================================================
Function FncInsertRow()
	On Error Resume Next                                                      '��: Protect system from crashing
End Function

'========================================================================================================
' Name : FncDeleteRow
' Desc : developer describe this line Called by MainDeleteRow in Common.vbs
'========================================================================================================
Function FncDeleteRow()
	On Error Resume Next                                                      '��: Protect system from crashing
End Function

'========================================================================================================
' Name : FncPrint
' Desc : developer describe this line Called by MainDeleteRow in Common.vbs
'========================================================================================================
Function FncPrint()
	Call Parent.FncPrint()                                                    '��: Protect system from crashing
End Function

'========================================================================================================
' Name : FncPrev
' Desc : developer describe this line Called by MainPrev in Common.vbs
'========================================================================================================
Function FncPrev() 

    Dim strVal
    Dim IntRetCD

    FncPrev = False                                                              '��: Processing is OK
    Err.Clear                                                                    '��: Clear err status

    If lgIntFlgMode <> parent.OPMD_UMODE Then                                           '��: Please do Display first. 
        Call DisplayMsgBox("900002","x","x","x")
        Exit Function
    End If
	
	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900017", parent.VB_YES_NO,"x","x")					 '��: Will you destory previous data
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If
	
    Call MakeKeyStream("P")
    Call ggoOper.ClearField(Document, "2")										 '��: Clear Contents Area
    
    Call InitVariables														 '��: Initializes local global variables

    Call LayerShowHide(1)


    strVal = BIZ_PGM_ID & "?txtMode="          & parent.UID_M0001                       '��: Query
    strVal = strVal     & "&txtKeyStream="     & lgKeyStream                     '��: Query Key
    strVal = strVal     & "&txtPrevNext="      & "P"	                         '��: Direction

	Call RunMyBizASP(MyBizASP, strVal)										     '��: Run Biz 
	
    FncPrev = True                                                               '��: Processing is OK

End Function

'========================================================================================================
' Name : FncNext
' Desc : developer describe this line Called by MainNext in Common.vbs
'========================================================================================================
Function FncNext() 
    Dim strVal
    Dim IntRetCD

    FncNext = False                                                              '��: Processing is OK
    Err.Clear                                                                    '��: Clear err status
    
    If lgIntFlgMode <> parent.OPMD_UMODE Then                                           '��: Please do Display first. 
        Call DisplayMsgBox("900002","x","x","x")
        Exit Function
    End If
	
	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900017", parent.VB_YES_NO,"x","x")					 '��: Will you destory previous data
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If
	
    Call MakeKeyStream("N")

    Call ggoOper.ClearField(Document, "2")										 '��: Clear Contents Area
    
    Call InitVariables														     '��: Initializes local global variables

    Call LayerShowHide(1)

    strVal = BIZ_PGM_ID & "?txtMode="          & parent.UID_M0001                       '��: Query
    strVal = strVal     & "&txtKeyStream="     & lgKeyStream                     '��: Query Key
    strVal = strVal     & "&txtPrevNext="      & "N"	                         '��: Direction

	Call RunMyBizASP(MyBizASP, strVal)										     '��: Run Biz 
    FncNext = True                                                               '��: Processing is OK
	
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
	Call Parent.FncFind(parent.C_SINGLE, True)
End Function

'========================================================================================================
' Name : FncExit
' Desc : developer describe this line Called by MainExit in Common.vbs
'========================================================================================================
Function FncExit()
	Dim IntRetCD

	FncExit = False
	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO,"x","x")			'��: Data is changed.  Do you want to exit? 
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If

	FncExit = True
End Function

'========================================================================================================
' Name : DbQuery
' Desc : This function is called by FncQuery
'========================================================================================================
Function DbQuery()
    Dim strVal
    Err.Clear                                                                    '��: Clear err status

    DbQuery = False                                                              '��: Processing is NG

	Call ggoOper.ClearField(Document, "2")
	Call InitVariables
	Call LayerShowHide(1)

    strVal = BIZ_PGM_ID & "?txtMode="          & parent.UID_M0001                       '��: Query
    strVal = strVal     & "&txtKeyStream="     & lgKeyStream                     '��: Query Key
    strVal = strVal     & "&txtPrevNext="      & ""	                             '��: Direction

    Call RunMyBizASP(MyBizASP, strVal)                                           '��:  Run biz logic
	
    DbQuery = True                                                               '��: Processing is NG
End Function
'========================================================================================================
' Name : DbSave
' Desc : This function is called by FncSave
'========================================================================================================
Function DbSave()
	Dim strVal
    Err.Clear                                                                    '��: Clear err status
		
	DbSave = False														         '��: Processing is NG
		
	Call LayerShowHide(1)
	With Frm1
		.txtMode.value        = parent.UID_M0002                                        '��: Delete
		.txtFlgMode.value     = lgIntFlgMode
        .txtKeyStream.Value   = lgKeyStream                                      '��: Save Key
	End With
	Call ExecMyBizASP(frm1, BIZ_PGM_ID)

    DbSave  = True                                                               '��: Processing is NG
End Function
'========================================================================================================
' Name : DbDelete
' Desc : This function is called by FncDelete
'========================================================================================================
Function DbDelete()
	Dim strVal
    Err.Clear                                                                    '��: Clear err status
		
	DbDelete = False			                                                 '��: Processing is NG
		
	Call LayerShowHide(1)
		
    strVal = BIZ_PGM_ID & "?txtMode="          & parent.UID_M0003                       '��: Query
    strVal = strVal     & "&txtKeyStream="     & lgKeyStream                     '��: Query Key
    strVal = strVal     & "&txtPrevNext="      & ""	                             '��: Direction

	Call RunMyBizASP(MyBizASP, strVal)                                           '��: Run Biz logic
	
	DbDelete = True                                                              '��: Processing is NG
End Function
'========================================================================================================
' Function Name : DbQueryOk
' Function Desc : Called by MB Area when query operation is successful
'========================================================================================================
Function DbQueryOk()

	lgIntFlgMode      = parent.OPMD_UMODE                                               '��: Indicates that current mode is Create mode

    Frm1.txtYear.focus 

	lgBlnFlgChgValue = False
	Call SetToolbar("1101100011011111")												'��: Set ToolBar

    Call ggoOper.LockField(Document, "Q")
    Set gActiveElement = document.ActiveElement   
	frm1.txtOther_insur_amt.focus
End Function
	
'========================================================================================================
' Function Name : DbSaveOk
' Function Desc : Called by MB Area when save operation is successful
'========================================================================================================
Function DbSaveOk()
    Call InitVariables
    Call FncQuery()
End Function
	
'========================================================================================================
' Function Name : DbDeleteOk
' Function Desc : Called by MB Area when delete operation is successful
'========================================================================================================
Function DbDeleteOk()
	Call InitVariables()
	Call FncQuery()	
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
	    arrParam(2) = lgUsrIntCd                    ' �ڷ���� Condition  
    Else 'spread
        frm1.vspdData.Col = C_EMP_NO
		arrParam(0) = frm1.vspdData.Text			' Code Condition
        frm1.vspdData.Col = C_NAME
	    arrParam(1) = ""'frm1.vspdData.Text			' Name Cindition
	    arrParam(2) = lgUsrIntCd                    ' �ڷ���� Condition  
	End If

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
'	Description : Item Popup���� Return�Ǵ� �� setting(grid�ܿ��� ���)
'=======================================================================================================
Sub SetEmpName(arrRet)
	With frm1
		.txtEmp_no.value    = arrRet(0)
		.txtName.value     = arrRet(1)
		.txtDept_nm.value  = arrRet(2)
		.txtRollPstn.value = arrRet(3)
		.txtPay_grd.value  = arrRet(4)
		.txtEntr_dt.text   = arrRet(5)
		Call ggoOper.ClearField(Document, "2")					 '��: Clear Contents  Field
		Set gActiveElement = document.ActiveElement
		.txtEmp_no.focus
		lgBlnFlgChgValue = False
	End With
End Sub

'========================================================================================================
'   Event Name : txtEmp_no_Onchange             '<==�λ縶���Ϳ� �ִ� ������� Ȯ�� 
'   Event Desc :
'========================================================================================================
Function txtEmp_no_Onchange()
    Dim iDx
    Dim IntRetCd
    Dim strName
    Dim strDept_nm
    Dim strRoll_pstn
    Dim strPay_grd1
    Dim strPay_grd2
    Dim strEntr_dt
    Dim strInternal_cd
    Dim RetStatus
    Dim strVal
    
    If frm1.txtEmp_no.value = "" Then
       frm1.txtName.value = ""
	   frm1.txtDept_nm.value = ""
	   frm1.txtRollPstn.value = ""
	   frm1.txtEntr_dt.text = ""
	   frm1.txtPay_grd.value = ""
	   Call ggoOper.ClearField(Document, "2")
    Else
	    IntRetCd = FuncGetEmpInf2(frm1.txtEmp_no.value,lgUsrIntCd,strName,strDept_nm,_
	                              strRoll_pstn, strPay_grd1, strPay_grd2, strEntr_dt, strInternal_cd)
	    If  IntRetCd < 0 then
	        If  IntRetCd = -1 then
    			Call DisplayMsgBox("800048","X","X","X")	'�ش����� �������� �ʽ��ϴ�.
            Else
                Call DisplayMsgBox("800454","X","X","X")	'�ڷῡ ���� ������ �����ϴ�.
            End if
            
			frm1.txtName.value = ""
			frm1.txtDept_nm.value  = ""
			frm1.txtRollPstn.value = ""
			frm1.txtPay_grd.value  = ""
			frm1.txtEntr_dt.text   = ""
            frm1.txtEmp_no.focus

		    Call ggoOper.ClearField(Document, "2")
            
            Set gActiveElement = document.ActiveElement
            txtEmp_no_Onchange = true
            Exit Function      
        Else
            frm1.txtName.value = strName
    		frm1.txtDept_nm.value  = strDept_nm
			frm1.txtRollPstn.value = strRoll_pstn
			frm1.txtPay_grd.value  = strPay_grd1 & "-" & strPay_grd2
			frm1.txtEntr_dt.text   = UNIDateClientFormat(strEntr_dt)
        End if 
    End if  

End Function 

'======================================================================================================
' Function Name : BundleCreate
' Function Desc : �ϰ����� 
'=======================================================================================================
Function BundleCreate() 
	Dim strVal
	Dim strYyyymm
	Dim IntRetCD

    If txtEmp_no_Onchange() Then         'ENTER KEY �� ��ȸ�� ����� ����� CHECK �Ѵ� 
        Exit Function
    End if

	IntRetCD = DisplayMsgBox("900018",parent.VB_YES_NO,"X","X")
	
	If IntRetCD = vbNo Then
		Exit Function
	End If

	Call LayerShowHide(1)

	BundleCreate = False                                                          '��: Processing is NG
    
	On Error Resume Next                                                   '��: Protect system from crashing

	strVal = BIZ_PGM_ID2 & "?txtMode=" & parent.UID_M0001
	strVal = strVal & "&txtYear_yy=" & Frm1.txtYear.Year

	Call RunMyBizASP(MyBizASP, strVal)										'��: �����Ͻ� ASP �� ���� 

	BundleCreate = True                                                           '��: Processing is NG

End Function

'======================================================================================================
' Function Name : BundleCreateOk
' Function Desc : BundleCreate�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncSave�� �ִ°��� �ű� 
'=======================================================================================================
Function BundleCreateOk()				            '��: ���� ������ ���� ���� 
	Dim IntRetCD 

	IntRetCD =DisplayMsgBox("990000","X","X","X")
	window.status = "�۾� �Ϸ�"

End Function

Function BundleCreateNo()				            '��: ���� ������ ���� ���� 
	Dim IntRetCD 

    Call DisplayMsgBox("800414","X","X","X")
	window.status = "�۾� ����"

End Function

'======================================================================================================
' Function Name : BasicDataCreate
' Function Desc : �����ڷ���� 
'=======================================================================================================
Function BasicDataCreate() 
	Dim strVal
	Dim strYyyymm
	Dim IntRetCD
    If txtEmp_no_Onchange() Then         'ENTER KEY �� ��ȸ�� ����� ����� CHECK �Ѵ� 
        Exit Function
    End if

	If Not chkField(Document, "1") Then
		Exit Function
	End If

	IntRetCD = DisplayMsgBox("900018",parent.VB_YES_NO,"X","X")
	If IntRetCD = vbNo Then
		Exit Function
	End If
	
    ' �̹� �ڷᰡ �ִ��� üũ 
    IntRetCd = CommonQueryRs(" EMP_NO "," HFA030T "," YY =  " & FilterVar(Frm1.txtYear.Year , "''", "S") & " AND EMP_NO =  " & FilterVar(Frm1.txtEmp_no.Value , "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

    If IntRetCd = True then
		IntRetCD = DisplayMsgBox("800397",parent.VB_YES_NO,"X","X")        '�ڷᰡ �����մϴ�. �ٽ� �����ϰڽ��ϱ�?	
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

	Call LayerShowHide(1)
	Call DisableToolBar(parent.TBC_QUERY)

	BasicDataCreate = False                                                          '��: Processing is NG

	strVal = BIZ_PGM_ID3 & "?txtMode=" & parent.UID_M0001
	strVal = strVal & "&txtYear_yy=" & Frm1.txtYear.Year
	strVal = strVal & "&txtEmp_no=" & Frm1.txtEmp_no.Value

	Call RunMyBizASP(MyBizASP, strVal)										'��: �����Ͻ� ASP �� ���� 

	BasicDataCreate = True                                                           '��: Processing is NG

End Function

'======================================================================================================
' Function Name : BasicDataCreateOk
' Function Desc : BasicDataCreate�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncSave�� �ִ°��� �ű� 
'=======================================================================================================
Function BasicDataCreateOk()				            '��: ���� ������ ���� ���� 
	Dim IntRetCD 

	IntRetCD =DisplayMsgBox("990000","X","X","X")
	window.status = "�۾� �Ϸ�"

	Call RestoreToolBar()
    Call MakeKeyStream("X")

    If DbQuery = False Then  
		Exit Function
	End If                                                                 '��: Query db data

End Function

Function BasicDataCreateNo()				            '��: ���� ������ ���� ���� 
	Dim IntRetCD 

    Call DisplayMsgBox("800414","X","X","X")
	window.status = "�۾� ����"

End Function


'======================================================================================================
' Function Name : ReCreate
' Function Desc : ����� 
'=======================================================================================================
Function ReCreate() 
	Dim strVal
	Dim strYyyymm
	Dim IntRetCD

    If txtEmp_no_Onchange() Then         'ENTER KEY �� ��ȸ�� ����� ����� CHECK �Ѵ� 
        Exit Function
    End if

	IntRetCD = DisplayMsgBox("900018",parent.VB_YES_NO,"X","X")
	
	If IntRetCD = vbNo Then
		Exit Function
	End If

	Call LayerShowHide(1)

	ReCreate = False                                                          '��: Processing is NG
    
	On Error Resume Next                                                   '��: Protect system from crashing

	strVal = BIZ_PGM_ID4 & "?txtMode=" & parent.UID_M0001
	strVal = strVal & "&txtYear_yy=" & Frm1.txtYear.Year

	Call RunMyBizASP(MyBizASP, strVal)										'��: �����Ͻ� ASP �� ���� 

	ReCreate = True                                                           '��: Processing is NG

End Function

'======================================================================================================
' Function Name : BundleCreateOk
' Function Desc : BundleCreate�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncSave�� �ִ°��� �ű� 
'=======================================================================================================
Function ReCreateOk()				            '��: ���� ������ ���� ���� 
	Dim IntRetCD 

	IntRetCD =DisplayMsgBox("990000","X","X","X")
	window.status = "�۾� �Ϸ�"

	Call RestoreToolBar()
    Call MakeKeyStream("X")

    If DbQuery = False Then  
		Exit Function
	End If  
End Function

Function ReCreateNo()				            '��: ���� ������ ���� ���� 
	Dim IntRetCD 

    Call DisplayMsgBox("800414","X","X","X")
	window.status = "�۾� ����"

End Function
'========================================================================================================
' Name : Change events
' Desc : Change events
'========================================================================================================
Sub txtOther_insur_amt_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtDisabled_insur_amt_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtMed_insur_amt_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtEmp_insur_amt_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtNational_pension_amt_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtPer_edu_amt_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtTot_med_amt_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtSpeci_med_amt_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtLegal_contr_amt_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtPoli_contr_amt1_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtOurstock_contr_amt_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtTaxLaw_contr_amt_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtTaxLaw_contr_amt2_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtApp_contr_amt_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtPriv_contr_amt_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtHouse_fund_amt_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtLong_house_loan_amt_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtIndiv_anu_amt_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtIndiv_anu2_amt_Change()
	lgBlnFlgChgValue = True
End Sub
 
Sub txtinvest2_sub_amt_Change()
	lgBlnFlgChgValue = True
End Sub 

Sub txtCard_use_amt_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtCard2_use_amt_Change() '����ī��(2003)
	lgBlnFlgChgValue = True
End Sub

Sub txtCard2_use_amt_Change() '�п������γ���(2005)
	lgBlnFlgChgValue = True
End Sub

Sub txtInstitution_giro_Change() '�ܱ��αٷ����Ǳ�����(2003)
	lgBlnFlgChgValue = True
End Sub

Sub txtOther_income_amt_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtFore_income_amt_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtAfter_bonus_amt_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtFore_edu_amt_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtHouse_repay_amt_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtStock_save_amt_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txOur_Stock_amt_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtRetire_pension_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtFore_pay_amt_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtIncome_redu_amt_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtTaxes_redu_amt_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtTax_Union_Ded_Change()	'2005
	lgBlnFlgChgValue = True
End Sub

Sub txtCeremony_cnt_Change()	'2004 ��ȥ/���/�̻��-��ȥ 
	lgBlnFlgChgValue = True
End Sub

Sub txtOld_cnt_t1_Change()		'2004 ��ο�����(65���̻�)
	lgBlnFlgChgValue = True
End Sub

Sub txtOld_cnt_t2_Change()		'2004 ��ο�����(70���̻�)
	lgBlnFlgChgValue = True
End Sub

Sub txtLong_house_loan_amt1_Change()		'2004 ��������������Աݻ�ȯ�Ⱓ 15���̻� 
	lgBlnFlgChgValue = True
End Sub
'========================================================================================================
' Name : txtYear_DblClick
' Desc : developer describe this line 
'========================================================================================================
Sub txtYear_DblClick(Button)
    If Button = 1 Then
		Call SetFocusToDocument("M")    
        frm1.txtYear.Action = 7 
        frm1.txtYear.focus
    End If
    lgBlnFlgChgValue = True
End Sub

</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
</HEAD>

<BODY SCROLL="AUTO" TABINDEX="-1">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR >
		<TD <%=HEIGHT_TYPE_00%>></TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>�����ڷ���</font></td>
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
		<TD WIDTH=100% CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS=TD5 NOWRAP>����⵵</TD>
									<TD CLASS=TD6 NOWRAP>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME="txtYear" style="HEIGHT: 20px; WIDTH: 50px" tag="12X1" Title="FPDATETIME" ALT="����⵵" id=fpDateTime1> </OBJECT>');</SCRIPT>
									</TD>	
									<TD CLASS=TD5 NOWRAP>���</TD>
			     					<TD CLASS=TD6 NOWRAP><INPUT NAME="txtEmp_no" ALT="���" TYPE="Text" SiZE=15 MAXLENGTH=13 tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnEmpNo" align=top TYPE="BUTTON" ONCLICK="VBScript: OpenEmpName('0')">
									                     <INPUT NAME="txtName" MAXLENGTH="30" SIZE="20" ALT ="����" tag="14XXXU"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>�μ���</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtDept_nm" MAXLENGTH="20" SIZE=20  ALT ="�μ���" tag="14">&nbsp;</TD>
									<TD CLASS=TD5 NOWRAP>��  ��</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtRollPstn" MAXLENGTH="20" SIZE=20 ALT ="����" tag="14">&nbsp;</TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>�Ի���</TD>
									<TD CLASS=TD6 NOWRAP>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT id=fpDateTime2 title=FPDATETIME style="LEFT: 0px; WIDTH: 80px; TOP: 0px; HEIGHT: 20px" name=txtEntr_dt CLASSID=<%=gCLSIDFPDT%> ALT="�Ի���" tag="14X1" VIEWASTEXT></OBJECT>');</SCRIPT>
									</TD>
									<TD CLASS=TD5 NOWRAP>��  ȣ</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtPay_grd" MAXLENGTH="20" SIZE=20 ALT ="��ȣ" tag="14">&nbsp;</TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD  WIDTH=100%></TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=* valign=top >
					    <TABLE <%=LR_SPACE_TYPE_20%> BORDER=0>
					        <TR>
					           <TD WIDTH=50% valign=top>
					                <TABLE <%=LR_SPACE_TYPE_20%>>
					                <TR>
					                    <TD VALIGN=TOP>
            								<FIELDSET CLASS="CLSFLD"><LEGEND ALIGN=LEFT>��������</LEGEND>					                    
                                            <TABLE <%=LR_SPACE_TYPE_20%>>
								                	<TR>
														<TD CLASS=TD5 NOWRAP><b>�⺻����</b></TD>
		                								<TD CLASS=TD6 NOWRAP>&nbsp;</TD>
								                	</TR>
								                    <TR>
            											<TD CLASS=TD5 NOWRAP>�����</TD>
			                                           	<TD CLASS=TD6 NOWRAP>
									                        <INPUT TYPE="CHECKBOX" NAME="rdoSpouse_t" ID="rdoPhantomType1" Value="Y" CLASS="RADIO" tag="24">
									                    </TD>
								                	</TR>
								                	<TR>
										                <TD CLASS=TD5 NOWRAP>�ξ���(��)</TD>
                										<TD CLASS=TD6 NOWRAP>
                										    <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%>fpDoubleSingle2="Object10" name=txtSupp_old_cnt_t style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 50px" title=FPDOUBLESINGLE tag="24X7Z" ALT="�ξ���(��)" id=OBJECT12></OBJECT>');</SCRIPT>
                										</TD>
								                	</TR>
								                    <TR>
														<TD CLASS=TD5 NOWRAP>�ξ���(��)</TD>
                										<TD CLASS=TD6 NOWRAP COLSPAN=3>
                										    <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle2 name=txtSupp_young_cnt_t style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 50px" title=FPDOUBLESINGLE tag="24X7Z" ALT="�ξ���(��)"></OBJECT>');</SCRIPT>
                										</TD>
								                	</TR>
								                	<TR>
														<TD CLASS=TD5 NOWRAP>&nbsp;</TD>
		                								<TD CLASS=TD6 NOWRAP>&nbsp;</TD>
								                	</TR>
								                	<TR>
														<TD CLASS=TD5 NOWRAP><b>�Ҽ�����</b></TD>
		                								<TD CLASS=TD6 NOWRAP>&nbsp;</TD>
								                	</TR>
								                    <TR>
            										    <TD CLASS=TD5 NOWRAP>�γ���</TD>
										                <TD CLASS=TD6 NOWRAP COLSPAN=3>
												            <INPUT TYPE="CHECKBOX" NAME="rdoLady_t" ID="rdoPhantomType2" Value="N" CLASS="RADIO" tag="24">
														</TD>
								                	</TR>	
								                    <TR>
										                <TD CLASS=TD5 NOWRAP>�����</TD>
				        								<TD CLASS=TD6 NOWRAP>
														    <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle2 name=txtParia_cnt_t style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 50px" title=FPDOUBLESINGLE tag="24X7Z" ALT="�����"></OBJECT>');</SCRIPT>
                										</TD>
										        	</TR>								                								                	
								                	<TR>
										                <TD CLASS=TD5 NOWRAP>�����(65���̻�)</TD>
                										<TD CLASS=TD6 NOWRAP>
                										    <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle2 name=txtOld_cnt_t1 style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 50px" title=FPDOUBLESINGLE tag="24X7Z" ALT="�����(65���̻�)"></OBJECT>');</SCRIPT>
                										</TD>
								                	</TR>
													<TR>
														<TD CLASS=TD5 NOWRAP>�����(70���̻�)</TD>
                										<TD CLASS=TD6 NOWRAP>
                										    <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle2 name=txtOld_cnt_t2 style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 50px" title=FPDOUBLESINGLE tag="24X7Z" ALT="�����(70���̻�)"></OBJECT>');</SCRIPT>
                										</TD>
								                	</TR>
								                    <TR>
														<TD CLASS=TD5 NOWRAP>�ڳ������</TD>
		                								<TD CLASS=TD6 NOWRAP>
				        								    <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle2 name=txtChl_rear_inwon_t style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 50px" title=FPDOUBLESINGLE tag="24X7Z" ALT="�ڳ������"></OBJECT>');</SCRIPT>
														</TD>
								                	</TR>
								                	<TR>
														<TD CLASS=TD5 NOWRAP>&nbsp;</TD>
		                								<TD CLASS=TD6 NOWRAP>&nbsp;</TD>
								                	</TR>
					                            </TABLE>
            			                    </TD>
            			                </TR>
					                    <TR>
					                        <TD VALIGN=TOP>
            										<FIELDSET CLASS="CLSFLD"><LEGEND ALIGN=LEFT>��Ÿ�ҵ����</LEGEND>							                        
                                                    <TABLE <%=LR_SPACE_TYPE_20%>>
								                        <TR>
								                            <TD CLASS=TD5 NOWRAP>���ο���(2000������)</TD>
								                            <TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle2 name=txtIndiv_anu_amt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 120px" title=FPDOUBLESINGLE tag="21X2Z" ALT="���ο���(2000������)"></OBJECT>');</SCRIPT></TD>
						                            	</TR>
								                        <TR>
								                            <TD CLASS=TD5 NOWRAP>��������(2001������)</TD>
								                            <TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle2 name=txtIndiv_anu2_amt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 120px" title=FPDOUBLESINGLE tag="21X2Z" ALT="��������(2001������)"></OBJECT>');</SCRIPT></TD>
						                            	</TR>
						                            	<TR>
								                            <TD CLASS=TD5 NOWRAP>�����������ھ�</TD>
								                            <TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="fpDoubleSingle2" name=txtinvest2_sub_amt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 120px" title=FPDOUBLESINGLE tag="21X2Z" ALT="2001.12.31�������ڱ�"></OBJECT>');</SCRIPT></TD>
						                            	</TR>
								                        <TR>
								                            <TD CLASS=TD5 NOWRAP>&nbsp;</TD>
								                            <TD CLASS=TD6 NOWRAP></TD>
						                            	</TR>						                            	
														<TR><TD CLASS=TD5 NOWRAP>�ſ�/����/����ī��</TD>
								                            <TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="Object6" name=txtCard_use_amt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 120px" title=FPDOUBLESINGLE tag="24X2Z" ALT="�ſ�/����/����ī��"></OBJECT>');</SCRIPT></TD>
						                            	</TR>
														<TR><TD CLASS=TD5 NOWRAP>���ݿ�����</TD>
								                            <TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="Object6" name=txtCard2_use_amt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 120px" title=FPDOUBLESINGLE tag="24X2Z" ALT="���ݿ�����"></OBJECT>');</SCRIPT></TD>
						                            	</TR>
														<TR><TD CLASS=TD5 NOWRAP>�п������γ���</TD>
								                            <TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="Object6" name=txtInstitution_giro style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 120px" title=FPDOUBLESINGLE tag="24X2Z" ALT="�п������γ���"></OBJECT>');</SCRIPT></TD>
						                            	</TR>
								                        <TR>
								                            <TD CLASS=TD5 NOWRAP>&nbsp;</TD>
								                            <TD CLASS=TD6 NOWRAP></TD>
						                            	</TR>						                            	
						                            	<TR>
								                            <TD CLASS=TD5 NOWRAP>�츮�����⿬��</TD>
								                            <TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="Object4" name=txOur_Stock_amt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 120px" title=FPDOUBLESINGLE tag="21X2Z" ALT="�츮�����⿬��"></OBJECT>');</SCRIPT></TD>
						                            	</TR>
						                            	<TR>
								                            <TD CLASS=TD5 NOWRAP>�������ݼҵ����</TD>
								                            <TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="Object4" name=txtRetire_pension style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 120px" title=FPDOUBLESINGLE tag="21X2Z" ALT="�츮�����⿬��"></OBJECT>');</SCRIPT></TD>
						                            	</TR> 							                            	 						                            		                                                    
						                            </TABLE>						                            
						                           </FIELDSET> 
            			                    </TD>
            			                </TR>     
										<tr>
											<TD WIDTH=* valign=top>
            									<FIELDSET CLASS="CLSFLD"><LEGEND ALIGN=LEFT>��Ÿ�ҵ�</LEGEND>
            									<TABLE <%=LR_SPACE_TYPE_20%> ID="Table1">
            										<TR>
														<TD CLASS=TD5 NOWRAP>��Ÿ�ҵ�</TD>
														<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="Object7" name=txtOther_income_amt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 120px" title=FPDOUBLESINGLE tag="21X2Z" ALT="��Ÿ�ҵ�"></OBJECT>');</SCRIPT></TD>
													</TR>
													<TR>
														<TD CLASS=TD5 NOWRAP>�ܱ��ҵ�</TD>
														<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="Object8" name=txtFore_income_amt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 120px" title=FPDOUBLESINGLE tag="21X2Z" ALT="�ܱ��ҵ�"></OBJECT>');</SCRIPT></TD>
													</TR>
													<TR>
														<TD CLASS=TD5 NOWRAP>������</TD>
														<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="Object9" name=txtAfter_bonus_amt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 120px" title=FPDOUBLESINGLE tag="21X2Z" ALT="������"></OBJECT>');</SCRIPT></TD>
													</TR>
						                            <TR>
								                        <TD CLASS=TD5 NOWRAP>�ܱ��α�����/������<br></TD>
								                        <TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="fpDoubleSingle2" name=txtFore_edu_amt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 100px" title=FPDOUBLESINGLE tag="21X2Z" ALT="�ܱ��α�����/������"></OBJECT>');</SCRIPT></TD>
						                            </TR>														
                								</table>   
												</FIELDSET> 
            								</TD>   
            							</tr> 
					                    <TR>
					                        <TD>
												<FIELDSET CLASS="CLSFLD"><LEGEND ALIGN=LEFT>���װ���</LEGEND>
                                                    <TABLE <%=LR_SPACE_TYPE_20%>>
								                        <TR>
								                            <TD CLASS=TD5 NOWRAP>�������Ա����ڻ�ȯ��</TD>
								                            <TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle2 name=txtHouse_repay_amt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 120px" title=FPDOUBLESINGLE tag="21X2Z" ALT="�������Ա����ڻ�ȯ��"></OBJECT>');</SCRIPT></TD>
						                            	</TR>
								                        <TR>
								                            <TD CLASS=TD5 NOWRAP>�ܱ����μ���</TD>
								                            <TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle2 name=txtFore_pay_amt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 120px" title=FPDOUBLESINGLE tag="21X2Z" ALT="�ܱ����μ���"></OBJECT>');</SCRIPT></TD>
						                            	</TR>
								                        <TR>
								                            <TD CLASS=TD5 NOWRAP>���ٹ����������</TD>
								                            <TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle2 name=txtSave_tax_sub_amt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 120px" title=FPDOUBLESINGLE tag="24X2Z" ALT="���ٹ����������"></OBJECT>');</SCRIPT></TD>
						                            	</TR>
								                        <TR>
								                            <TD CLASS=TD5 NOWRAP>�ҵ漼��</TD>
								                            <TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle2 name=txtIncome_redu_amt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 120px" title=FPDOUBLESINGLE tag="21X2Z" ALT="�ҵ漼��"></OBJECT>');</SCRIPT></TD>
						                            	</TR>
								                        <TR>
								                            <TD CLASS=TD5 NOWRAP>������</TD>
								                            <TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle2 name=txtTaxes_redu_amt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 120px" title=FPDOUBLESINGLE tag="21X2Z" ALT="������"></OBJECT>');</SCRIPT></TD>
						                            	</TR>	
								                        <TR>
								                            <TD CLASS=TD5 NOWRAP>���ٳ������հ���</TD>
								                            <TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle2 name=txtTax_Union_Ded style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 120px" title=FPDOUBLESINGLE tag="21X2Z" ALT="���ٳ������հ���"></OBJECT>');</SCRIPT></TD>
						                            	</TR>						                            		
						                            </TABLE>
												</FIELDSET> 
            			                    </TD>
            			                </TR>           			                
    
            			            </TABLE>            
            			        </TD>
					            <TD WIDTH=50% valign=top>
            									<FIELDSET CLASS="CLSFLD"><LEGEND ALIGN=LEFT>Ư������</LEGEND>					            
					                <TABLE <%=LR_SPACE_TYPE_20%>>
					                
					                    <TR>
					                        <TD VALIGN=TOP>
												<TABLE <%=LR_SPACE_TYPE_20%>>
								                    <TR>
								                         <TD CLASS=TD5 NOWRAP>&nbsp;<b>�����</b></TD>
								                         <TD CLASS=TD6 NOWRAP></TD>
						                         	</TR>
						                         	<TR>
												 	    <TD CLASS=TD5 NOWRAP>&nbsp;�� Ÿ �� ��</TD>
												 	    <TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle2 name=txtOther_insur_amt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 120px" title=FPDOUBLESINGLE tag="21X2Z" ALT="��Ÿ����"></OBJECT>');</SCRIPT></TD>
						                         	</TR>
						                         	<TR>
												 	    <TD CLASS=TD5 NOWRAP>&nbsp;��������뺸��</TD>
												 	    <TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle2 name=txtDisabled_insur_amt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 120px" title=FPDOUBLESINGLE tag="21X2Z" ALT="��������뺸��"></OBJECT>');</SCRIPT></TD>
						                         	</TR>
												 	<TR>
												 	    <TD CLASS=TD5 NOWRAP>�� �� �� ��</TD>
												 	    <TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle2 name=txtMed_insur_amt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 120px" title=FPDOUBLESINGLE tag="21X2Z" ALT="�ǰ�����"></OBJECT>');</SCRIPT></TD>
						                         	</TR>
												 	<TR>
												 	    <TD CLASS=TD5 NOWRAP>�� �� �� ��</TD>
												 	    <TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle2 name=txtEmp_insur_amt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 120px" title=FPDOUBLESINGLE tag="21X2Z" ALT="��뺸��"></OBJECT>');</SCRIPT></TD>
						                         	</TR>
												 	<TR>
												 	    <TD CLASS=TD5 NOWRAP>�� �� �� ��</TD>
												 	    <TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle2 name=txtNational_pension_amt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 120px" title=FPDOUBLESINGLE tag="21X2Z" ALT="���ο���"></OBJECT>');</SCRIPT></TD>
						                         	</TR>
												 	<TR>
												 	    <TD CLASS=TD5 NOWRAP>&nbsp;</TD>
												 	    <TD CLASS=TD6 NOWRAP></TD>
						                         	</TR>
								                     <TR>
								                         <TD CLASS=TD5 NOWRAP>&nbsp;<b>�Ƿ��</b></TD>
								                         <TD CLASS=TD6 NOWRAP></TD>
						                         	</TR>						                         	
  								                     <TR>
								                         <TD CLASS=TD5 NOWRAP>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;�Ϲ��Ƿ��</TD>
								                         <TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle2 name=txtTot_med_amt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 120px" title=FPDOUBLESINGLE tag="24X2Z" ALT="�Ϲ��Ƿ��"></OBJECT>');</SCRIPT></TD>
						                         	</TR>
								                     <TR>
								                         <TD CLASS=TD5 NOWRAP>����/�����/������Ƿ��</TD>
								                         <TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle2 name=txtSpeci_med_amt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 120px" title=FPDOUBLESINGLE tag="24X2Z" ALT="����/�����/������Ƿ��"></OBJECT>');</SCRIPT></TD>
						                         	</TR> 
												 	<TR>
												 	    <TD CLASS=TD5 NOWRAP>&nbsp;</TD>
												 	    <TD CLASS=TD6 NOWRAP></TD>
						                         	</TR>						                            	
						                         	                                                   
								                    <TR>
								                        <TD CLASS=TD5 NOWRAP>&nbsp;<b>������</b></TD>
								                        <TD CLASS=TD6 NOWRAP></TD>
						                            </TR>
								                    <TR>
								                        <TD CLASS=TD5 NOWRAP>���α�����</TD>
								                        <TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle2 name=txtPer_edu_amt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 100px" title=FPDOUBLESINGLE tag="24X2Z" ALT="���α�����"></OBJECT>');</SCRIPT></TD>
						                            </TR>
								                    <TR>
								                        <TD CLASS=TD5 NOWRAP>���߰�����/���(��)</TD>
								                        <TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle2 name=txtFam_edu_amt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 100px" title=FPDOUBLESINGLE tag="24X2Z" ALT="���߰�����"></OBJECT>');</SCRIPT>
								                        &nbsp;/&nbsp;<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="Object1" name=txtFam_edu_cnt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 30px" title=FPDOUBLESINGLE tag="24X7Z" ALT="���߰������ڳ��"></OBJECT>');</SCRIPT></TD>
						                            </TR>
								                    <TR>
								                        <TD CLASS=TD5 NOWRAP>��ġ��������/���(��)</TD>
								                        <TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle2 name=txtKind_edu_amt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 100px" title=FPDOUBLESINGLE tag="24X2Z" ALT="��ġ��������"></OBJECT>');</SCRIPT>
								                        &nbsp;/&nbsp;<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="Object2" name=txtKind_edu_cnt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 30px" title=FPDOUBLESINGLE tag="24X7Z" ALT="��ġ���������ڳ��"></OBJECT>');</SCRIPT></TD>
						                            </TR>
								                    <TR>
								                        <TD CLASS=TD5 NOWRAP>���б�����/���(��)</TD>
								                        <TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle2 name=txtUniv_edu_amt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 100px" title=FPDOUBLESINGLE tag="24X2Z" ALT="���б�����"></OBJECT>');</SCRIPT>
								                        &nbsp;/&nbsp;<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="Object3" name=txtUniv_edu_cnt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 30px" title=FPDOUBLESINGLE tag="24X7Z" ALT="���б������ڳ��"></OBJECT>');</SCRIPT></TD>
						                            </TR>
						                            <TR>
								                        <TD CLASS=TD5 NOWRAP>�����Ư��������/���(��)</TD>
								                        <TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="fpDoubleSingle2" name=txtDisabled_edu_amt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 100px" title=FPDOUBLESINGLE tag="24X2Z" ALT="�����Ư��������"></OBJECT>');</SCRIPT>
								                        &nbsp;/&nbsp;<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="Object5" name=txtDisabled_edu_cnt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 30px" title=FPDOUBLESINGLE tag="24X7Z" ALT="����μ�"></OBJECT>');</SCRIPT></TD>
						                            </TR>
								                    <TR>
								                        <TD CLASS=TD5 NOWRAP>&nbsp;</TD>
								                        <TD CLASS=TD6 NOWRAP></TD>
						                            </TR> 							                            	
								                    <TR>
								                        <TD CLASS=TD5 NOWRAP>&nbsp;<b>�����ڱ�</b></TD>
								                        <TD CLASS=TD6 NOWRAP></TD>
						                            </TR> 						                            							                            							                            	
								                    <TR>
								                        <TD CLASS=TD5 NOWRAP>��������/���Աݻ�ȯ��</TD>
								                        <TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle2 name=txtHouse_fund_amt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 120px" title=FPDOUBLESINGLE tag="21X2Z" ALT="��������/���Աݻ�ȯ��"></OBJECT>');</SCRIPT></TD>
						                            </TR>
								                    <TR>
								                        <TD CLASS=TD5 NOWRAP>��������������Ա����ڻ�ȯ��(15��̸�)</TD>
								                        <TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle2 name=txtLong_house_loan_amt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 120px" title=FPDOUBLESINGLE tag="21X2Z" ALT="��������������Ա����ڻ�ȯ��(15��̸�)">&nbsp;</OBJECT>');</SCRIPT></TD>
						                            </TR>
								                    <TR>
								                        <TD CLASS=TD5 NOWRAP>��������������Ա����ڻ�ȯ��(15���̻�)</TD>
								                        <TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle2 name=txtLong_house_loan_amt1 style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 120px" title=FPDOUBLESINGLE tag="21X2Z" ALT="��������������Ա����ڻ�ȯ��(15���̻�)">&nbsp;</OBJECT>');</SCRIPT></TD>
						                            </TR>
								                    <TR>
								                        <TD CLASS=TD5 NOWRAP>&nbsp;</TD>
								                        <TD CLASS=TD6 NOWRAP></TD>
						                            </TR> 							                            	
								                    <TR>
								                        <TD CLASS=TD5 NOWRAP>&nbsp;<b>��α�</b></TD>
								                        <TD CLASS=TD6 NOWRAP></TD>
						                            </TR> 						                            	
								                    <TR>
								                        <TD CLASS=TD5 NOWRAP>������α�</TD>
								                        <TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle2 name=txtLegal_contr_amt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 120px" title=FPDOUBLESINGLE tag="24X2Z" ALT="������α�"></OBJECT>');</SCRIPT></TD>
						                            </TR>
								                    <TR>
								                        <TD CLASS=TD5 NOWRAP>��ġ�ڱݱ�α�</TD>
								                        <TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle2 name=txtPoli_contr_amt1 style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 120px" title=FPDOUBLESINGLE tag="24X2Z" ALT="��ġ�ڱݱ�α�"></OBJECT>');</SCRIPT></TD>
						                            </TR>	
						                            <TR>
								                        <TD CLASS=TD5 NOWRAP>������(75%)</TD>
								                        <TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="Object11" name=txtTaxLaw_contr_amt2 style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 120px" title=FPDOUBLESINGLE tag="24X2Z" ALT="Ư�ʱ�α�(100%)"></OBJECT>');</SCRIPT></TD>
						                            </TR>	
					                            						                            	
						                            <TR>
								                        <TD CLASS=TD5 NOWRAP>Ư�ʱ�α�(50%)</TD>
								                        <TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="Object11" name=txtTaxLaw_contr_amt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 120px" title=FPDOUBLESINGLE tag="24X2Z" ALT="Ư�ʱ�α�(50%)"></OBJECT>');</SCRIPT></TD>
						                            </TR>
								                    <TR>
								                        <TD CLASS=TD5 NOWRAP>�츮�������ձ�α�(30%)</TD>
								                        <TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle2 name=txtOurstock_contr_amt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 120px" title=FPDOUBLESINGLE tag="24X2Z" ALT="�츮�������ձ�α�(30%)"></OBJECT>');</SCRIPT></TD>
						                            </TR>						                        
								                    <TR>
								                        <TD CLASS=TD5 NOWRAP>������α�(10%)</TD>
								                        <TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle2 name=txtApp_contr_amt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 120px" title=FPDOUBLESINGLE tag="24X2Z" ALT="������α�"></OBJECT>');</SCRIPT></TD>
						                            </TR>
								                    <TR>
								                        <TD CLASS=TD5 NOWRAP>�뵿���պ�</TD>
								                        <TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle2 name=txtPriv_contr_amt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 120px" title=FPDOUBLESINGLE tag="21X2Z" ALT="�뵿���պ�"></OBJECT>');</SCRIPT></TD>
						                            </TR>
							                        <TR>
								                        <TD CLASS=TD5 NOWRAP>&nbsp;</TD>
								                        <TD CLASS=TD6 NOWRAP></TD>
						                            </TR> 							                            	
								                    <TR>
								                        <TD CLASS=TD5 NOWRAP>&nbsp;<b>ȥ��/���/�̻��</b></TD>
								                        <TD CLASS=TD6 NOWRAP></TD>
						                            </TR> 						                            	
								                    <TR>
								                        <TD CLASS=TD5 NOWRAP>��ȥ/���/�̻��</TD>
								                        <TD CLASS=TD6 NOWRAP>
															Ƚ��<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle2 name=txtCeremony_cnt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 50px" title=FPDOUBLESINGLE tag="21X7Z" ALT="��ȥ���Ƚ��"></OBJECT>');</SCRIPT>
															&nbsp;<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle2 name=txtCeremony_amt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 120px" title=FPDOUBLESINGLE tag="24X2Z" ALT="��ȥ/���/�̻��"></OBJECT>');</SCRIPT></TD>
						                            </TR>																							                            							                            								                            	
						                            </TABLE>
            			                    </TD>
            			                </TR>
 
             			                             			                          			                
            			            </TABLE>            
               			        </TD>

            			    </TR>
            			</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR >
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</TR>
	<TR HEIGHT=20>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_30%>>
			    <TR>
				   <TD WIDTH=10>&nbsp;</TD>
				   <TD>
				        <BUTTON NAME="btnSplit" CLASS="CLSMBTN" onclick="BundleCreate()" Flag=1>�ϰ�����</BUTTON>&nbsp;
				        <BUTTON NAME="btnSplit" CLASS="CLSMBTN" onclick="BasicDataCreate()" Flag=1>�����ڷ����</BUTTON>&nbsp;
				        <BUTTON NAME="btnSplit" CLASS="CLSMBTN" onclick="ReCreate()" Flag=1>�����</BUTTON>
				        &nbsp;
				   </TD>
				   <TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR >
	<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP"  WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=no noresize framespacing=0></IFRAME>  

		</TD>
	</TR>
</TABLE>
<DIV ID="MousePT" NAME="MousePT">
	<IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm"></IFRAME>
</DIV>
<INPUT TYPE=HIDDEN NAME="txtMode"      TAG="24">
<INPUT TYPE=HIDDEN NAME="txtKeyStream"   TAG="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId"  TAG="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode"     TAG="24">
<INPUT TYPE=HIDDEN NAME="txtPrevNext"    TAG="24">
</FORM>

</BODY>
</HTML>
