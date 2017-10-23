<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : Human Resources
*  2. Function Name        : 보증인등록 
*  3. Program ID           : H2005ma1
*  4. Program Name         : H2005ma1
*  5. Program Desc         : 인사기본자료관리/보증인등록 
*  6. Comproxy List        :
*  7. Modified date(First) : 2001/05/10
*  8. Modified date(Last)  : 2003/06/10
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
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs">   </SCRIPT>

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incHRQuery.vbs"> </SCRIPT>

<Script Language="VBScript">
Option Explicit
'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
Const WARRANT_TYPE_MAJOR = "S0002"
Const DEL_TYPE_MAJOR     = "S0003"
Const BIZ_PGM_ID      = "h2005mb1.asp"                 '☆: Biz Logic ASP Name
Const BIZ_PGM_JUMP_ID = "H2001ma1"

'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" --> 
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================

Dim gSelframeFlg                                                       '현재 TAB의 위치를 나타내는 Flag %>
Dim gblnWinEvent                                                       'ShowModal Dialog(PopUp) Window가 여러 개 뜨는 것을 방지하기 위해 
Dim lgBlnFlawChgFlg 
Dim gtxtChargeType
Dim IsOpenPop                                          ' Popup

'========================================================================================================
' Name : InitVariables() 
' Desc : Initialize value
'========================================================================================================
Sub InitVariables()
	lgIntFlgMode      = parent.OPMD_CMODE              '⊙: Indicates that current mode is Create mode
	lgBlnFlgChgValue  = False            '⊙: Indicates that no value changed
	lgIntGrpCount     = 0          '⊙: Initializes Group View Size
	lgStrPrevKey      = ""                                      '⊙: initializes Previous Key
	lgSortKey         = 1                                       '⊙: initializes sort direction
	gblnWinEvent      = False
	lgBlnFlawChgFlg   = False
End Sub

'========================================================================================================
' Name : LoadInfTB19029() 
' Desc : Set System Number format
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A( "I", "H","NOCOOKIE","MA") %>
 End Sub
'========================================================================================================
' Name : CookiePage()
' Description : Item Popup에서 Return되는 값 setting
'========================================================================================================
Function CookiePage(ByVal flgs)
	On Error Resume Next

	Const CookieSplit = 4877      
	Dim strTemp

	If flgs = 1 Then
		 WriteCookie CookieSplit , frm1.txtEmp_no.Value
	ElseIf flgs = 0 Then

		strTemp = ReadCookie(CookieSplit)
		If strTemp = "" then Exit Function
			frm1.txtEmp_no.value =  strTemp
		If Err.number <> 0 Then
			Err.Clear
			WriteCookie CookieSplit , ""
			Exit Function 
		End If
		WriteCookie CookieSplit , ""
		Call FncQuery()
	End If
End Function

'========================================================================================================
' Function Name : MakeKeyStream
' Function Desc : This method set focus to pos of err
'========================================================================================================
Sub MakeKeyStream(pOpt)
   
   If pOpt = "Q" Then   ' Query
      lgKeyStream = Frm1.txtEmp_no.Value & parent.gColSep              'You Must append one character(parent.gColSep)
      lgKeyStream = lgKeyStream & Frm1.txtName.Value & parent.gColSep  'You Must append one character(parent.gColSep)
   Else
      lgKeyStream = Frm1.txtEmp_no.Value & parent.gColSep               'You Must append one character(parent.gColSep)
   End If   
End Sub        
 
'========================================================================================================
' Name : Form_Load
' Desc : developer describe this line Called by Window_OnLoad() evnt
'========================================================================================================
Sub Form_Load()

    Err.Clear  
                                                                         '☜: Clear err status
	Call LoadInfTB19029
	Call AppendNumberRange("0", "00x00", "13x440")
	Call ggoOper.FormatField(Document, "A",ggStrIntegeralPart, ggStrDeciPointPart,gDateFormat,parent.gComNum1000,parent.gComNumDec)
	Call ggoOper.LockField(Document, "N")           '⊙: Lock Field

    Call FuncGetAuth(gStrRequestMenuID, parent.gUsrID, lgUsrIntCd)     ' 자료권한:lgUsrIntCd ("%", "1%")
    Call SetToolbar("1100100000001111")                  '버튼 툴바 제어 
	Call InitVariables
	frm1.txtEmp_no.focus
	Call CookiePage (0)                                                             '☜: Check Cookie
   
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

    If  lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO,"x","x")      '☜: Data is changed.  Do you want to display it? 
		If IntRetCD = vbNo Then
			 Exit Function
		End If
    End If

    If Not chkField(Document, "1") Then                  '☜: This function check required field
       Exit Function
    End If

    FncQuery = False                '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

    Call ggoOper.ClearField(Document, "2")           '☜: Clear Contents  Field    
   
    if  frm1.txtEmp_no.value = "" AND frm1.txtName.value <> "" then
        OpenEmpName(0)
        exit function
    end if
    If  txtEmp_no_Onchange() then
        Exit Function
    End If    

    Call InitVariables                                                           '⊙: Initializes local global variables
    Call MakeKeyStream("Q")
    
	Call DisableToolBar(parent.TBC_QUERY)
    If DbQuery = False Then
		Call RestoreToolBar()
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
    
    FncNew = False                 '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
    
    If lgBlnFlgChgValue = True Then
       IntRetCD = DisplayMsgBox("900015", parent.VB_YES_NO,"x","x")      '☜: Data is changed.  Do you want to make it new? 
       If IntRetCD = vbNo Then
          Exit Function
       End If
    End If
    
    Call ggoOper.ClearField(Document, "A")                                       '☜: Clear Condition Field
    Call ggoOper.LockField(Document , "N")                                       '☜: Lock  Field

    Frm1.imgPhoto.src = ""
    Call SetToolbar("11101000000011")
    Call InitVariables                                                        '⊙: Initializes local global variables
    
    Set gActiveElement = document.ActiveElement   
    
    FncNew = True                 '☜: Processing is OK
End Function
 
'========================================================================================================
' Name : FncDelete
' Desc : developer describe this line Called by MainDelete in Common.vbs
'========================================================================================================
Function FncDelete()
    Dim intRetCD
    
    FncDelete = False                                                            '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
    
    If lgIntFlgMode <> parent.OPMD_UMODE Then                                           '☜: Please do Display first. 
        Call DisplayMsgBox("900002","x","x","x")                                
        Exit Function
    End If
    
    IntRetCD = DisplayMsgBox("900003", parent.VB_YES_NO,"x","x")                        '☜: Do you want to delete? 
	If IntRetCD = vbNo Then
		 Exit Function
	End If
    
    Call MakeKeyStream("D")
    
	Call DisableToolBar(parent.TBC_DELETE)
    If DbDelete = False Then
		Call RestoreToolBar()
        Exit Function
    End If
    
    Set gActiveElement = document.ActiveElement   
    
    FncDelete = True                                                            '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncSave
' Desc : developer describe this line Called by MainSave in Common.vbs
'========================================================================================================
Function FncSave()
    Dim IntRetCD 
	dim strNat_cd        
    FncSave = False                                                              '☜: Processing is NG
    
    Err.Clear                                                                    '☜: Clear err status
    
    If lgBlnFlgChgValue = False Then 
        IntRetCD = DisplayMsgBox("900001","x","x","x")                           '☜:There is no changed data. 
        Exit Function
    End If

    If Not chkField(Document, "1") Then                  '☜: This function check required field
       Exit Function
    End If    
    If Not chkField(Document, "2") Then                                          '☜: Check contents area
       Exit Function
    End If
    
    if  frm1.txtEmp_no.value = "" then
        Frm1.txtEmp_no.focus
        Set gActiveElement = document.ActiveElement   
        exit function
    end if
    If  txtEmp_no_Onchange() then
        Exit Function
    End If  
    IF frm1.txtwarnt_insur_nm.value <> "" THEN 
        IF frm1.txtwarnt_amt.text = "" THEN
            frm1.txtwarnt_amt.text = 0
        end if

        IF frm1.txtwarnt_insur_no.value = "" THEN
            Call DisplayMsgBox("970021","X","보험번호","X")
            frm1.txtwarnt_insur_no.focus
            Set gActiveElement = document.ActiveElement
            exit function
        ELSEIF frm1.txtwarnt_amt.text = 0 THEN
            Call DisplayMsgBox("970021","X","보험료","X")
            frm1.txtwarnt_amt.focus
            Set gActiveElement = document.ActiveElement
            exit function
        ELSEIF frm1.txtwarnt_start.text = "" THEN
            Call DisplayMsgBox("970021","X","보증기간","X")
            frm1.txtwarnt_start.focus
            Set gActiveElement = document.ActiveElement
            exit function
        ELSEIF frm1.txtwarnt_end.text = "" THEN
            Call DisplayMsgBox("970021","X","보증기간","X")
            frm1.txtwarnt_end.focus
            Set gActiveElement = document.ActiveElement
            exit function
        END IF
 
    END IF   ' 보증보험 이름입력시 필수 입력항목 체크 
 
	If ValidDateCheck(frm1.txtwarnt_start, frm1.txtwarnt_end)=False Then
	 Exit Function
	End if
 
	If ValidDateCheck(frm1.txtEntr_dt, frm1.txtWarnt_start)=False Then
	 Exit Function
	End if

    call CommonQueryRs(" nat_cd "," HAA010T "," EMP_NO =  " & FilterVar(Frm1.txtemp_no.Value , "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    strNat_cd = Replace(lgF0, Chr(11), "")  ' 주민번호 check를 위해서 

	if UCase(strNat_cd) = "KR" then
		frm1.txtwarnt1_res_no.value = replace(frm1.txtwarnt1_res_no.value,"-","")	
	end if
	if len(frm1.txtwarnt1_res_no.value) > 13  then
		frm1.txtwarnt1_res_no.value = mid(frm1.txtwarnt1_res_no.value,1,13)
	end if 								
	if UCase(strNat_cd) = "KR" then
		frm1.txtwarnt2_res_no.value = replace(frm1.txtwarnt2_res_no.value,"-","")	
	end if
	if len(frm1.txtwarnt2_res_no.value) > 13  then
		frm1.txtwarnt2_res_no.value = mid(frm1.txtwarnt2_res_no.value,1,13)
	end if 								
	    
    IF frm1.txtwarnt1_name.value <> "" THEN 
        IF frm1.txtwarnt1_incom_tax.text = "" THEN
            frm1.txtwarnt1_incom_tax.text = 0
		end if

		IF frm1.txtwarnt1_res_no.value = "" THEN
		    Call DisplayMsgBox("970021","X","주민번호","X")
		    frm1.txtwarnt1_res_no.focus
		    Set gActiveElement = document.ActiveElement
		    exit function
		ELSEIF frm1.txtwarnt1_start.text = "" THEN
		    Call DisplayMsgBox("970021","X","보증기간","X")
		    frm1.txtwarnt1_start.focus
		    Set gActiveElement = document.ActiveElement
		    exit function
		ELSEIF frm1.txtwarnt1_end.text = "" THEN
		    Call DisplayMsgBox("970021","X","보증기간","X")
		    frm1.txtwarnt1_end.focus
		    Set gActiveElement = document.ActiveElement
		    exit function
	    END IF
   END IF   ' 보증인1 이름입력시 필수 입력항목 체크 

	If ValidDateCheck(frm1.txtwarnt1_start, frm1.txtwarnt1_end)=False Then
	 Exit Function
	End if

	If ValidDateCheck(frm1.txtEntr_dt, frm1.txtWarnt1_start)=False Then
	 Exit Function
	End if
    
    IF frm1.txtwarnt2_name.value <> "" THEN 
        IF frm1.txtwarnt2_incom_tax.text = "" THEN
            frm1.txtwarnt2_incom_tax.text = 0
        end if

        IF frm1.txtwarnt2_res_no.value = "" THEN
            Call DisplayMsgBox("970021","X","주민번호","X")
            frm1.txtwarnt2_res_no.focus
            Set gActiveElement = document.ActiveElement
            exit function
        ELSEIF frm1.txtwarnt2_start.text = "" THEN
            Call DisplayMsgBox("970021","X","보증기간","X")
            frm1.txtwarnt2_start.focus
            Set gActiveElement = document.ActiveElement
            exit function
        ELSEIF frm1.txtwarnt2_end.text = "" THEN
            Call DisplayMsgBox("970021","X","보증기간","X")
            frm1.txtwarnt2_end.focus
            Set gActiveElement = document.ActiveElement
            exit function
        END IF
 
        
   END IF   ' 보증인2 이름입력시 필수 입력항목 체크 
   
	If ValidDateCheck(frm1.txtwarnt2_start, frm1.txtwarnt2_end)=False Then
		Exit Function
	End if
 
	If ValidDateCheck(frm1.txtEntr_dt, frm1.txtWarnt2_start)=False Then
		Exit Function
	End if

    Call MakeKeyStream("S")
    
	Call DisableToolBar(parent.TBC_SAVE)
    If DbSave = False Then
		Call RestoreToolBar()
        Exit Function
    End If    
    
    FncSave = True                                                              '☜: Processing is OK
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

    If  lgBlnFlgChgValue = True Then
  IntRetCD = DisplayMsgBox("900017", parent.VB_YES_NO,"x","x")         '☜: Data is changed.  Do you want to continue? 
  If IntRetCD = vbNo Then
   Exit Function
  End If
    End If
    
    lgIntFlgMode = parent.OPMD_CMODE                 '⊙: Indicates that current mode is Crate mode
    lgBlnFlgChgValue = true
    Call ggoOper.ClearField(Document, "1")                                       '⊙: Clear Condition Field
     
    Call ggoOper.LockField(Document, "N")              '⊙: This function lock the suitable field
    Call SetToolbar("11101000000011")
    Set gActiveElement = document.ActiveElement   
    FncCopy = True                                                            '☜: Processing is OK
    
End Function

'========================================================================================================
' Name : FncCancel
' Desc : developer describe this line Called by MainCancel in Common.vbs
'========================================================================================================
Function FncCancel() 
 On Error Resume Next                                                      '☜: Protect system from crashing
End Function

'========================================================================================================
' Name : FncInsertRow
' Desc : developer describe this line Called by MainInsertRow in Common.vbs
'========================================================================================================
Function FncInsertRow()
 On Error Resume Next                                                      '☜: Protect system from crashing
End Function

'========================================================================================================
' Name : FncDeleteRow
' Desc : developer describe this line Called by MainDeleteRow in Common.vbs
'========================================================================================================
Function FncDeleteRow()
 On Error Resume Next                                                      '☜: Protect system from crashing
End Function

'========================================================================================================
' Name : FncPrint
' Desc : developer describe this line Called by MainDeleteRow in Common.vbs
'========================================================================================================
Function FncPrint()
 Call Parent.FncPrint()                                                    '☜: Protect system from crashing
End Function

'========================================================================================================
' Name : FncPrev
' Desc : developer describe this line Called by MainPrev in Common.vbs
'========================================================================================================
Function FncPrev() 

    Dim strVal
    Dim IntRetCD

    FncPrev = False                                                              '☜: Processing is OK
    Err.Clear                                                                    '☜: Clear err status
    
    If lgIntFlgMode <> parent.OPMD_UMODE Then                                           '☜: Please do Display first. 
        Call DisplayMsgBox("900002","x","x","x")
        Exit Function
    End If
 
	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900017", parent.VB_YES_NO,"x","x")      '☜: Will you destory previous data
		If IntRetCD = vbNo Then
			 Exit Function
		End If
	End If
 
    Call MakeKeyStream("P")
    Call ggoOper.ClearField(Document, "2")           '⊙: Clear Contents Area
   
    Call InitVariables               '⊙: Initializes local global variables

	If   LayerShowHide(1) = False Then
	     Exit Function
	End If

    strVal = BIZ_PGM_ID & "?txtMode="          & parent.UID_M0001                       '☜: Query
    strVal = strVal     & "&txtKeyStream="     & lgKeyStream                     '☜: Query Key
    strVal = strVal     & "&txtPrevNext="      & "P"                          '☜: Direction

	Call RunMyBizASP(MyBizASP, strVal)               '☜: Run Biz 
    
    FncPrev = True                                                               '☜: Processing is OK

End Function
'========================================================================================================
' Name : FncNext
' Desc : developer describe this line Called by MainNext in Common.vbs
'========================================================================================================
Function FncNext() 
    Dim strVal
    Dim IntRetCD

    FncNext = False                                                              '☜: Processing is OK
    Err.Clear                                                                    '☜: Clear err status
    
    If lgIntFlgMode <> parent.OPMD_UMODE Then                                           '☜: Please do Display first. 
        Call DisplayMsgBox("900002","x","x","x")
        Exit Function
    End If
 
	If lgBlnFlgChgValue = True Then
		 IntRetCD = DisplayMsgBox("900017", parent.VB_YES_NO,"x","x")      '☜: Will you destory previous data
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If
 
	Call MakeKeyStream("N")

	Call ggoOper.ClearField(Document, "2")           '⊙: Clear Contents Area
 
	Call InitVariables                   '⊙: Initializes local global variables

	If   LayerShowHide(1) = False Then
	     Exit Function
	End If
	
    strVal = BIZ_PGM_ID & "?txtMode="          & parent.UID_M0001                       '☜: Query
    strVal = strVal     & "&txtKeyStream="     & lgKeyStream                     '☜: Query Key
    strVal = strVal     & "&txtPrevNext="      & "N"                          '☜: Direction

	Call RunMyBizASP(MyBizASP, strVal)               '☜: Run Biz 

   
    FncNext = True                                                               '☜: Processing is OK

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
  IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO,"x","x")   '⊙: Data is changed.  Do you want to exit? 
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
    Err.Clear                                                                    '☜: Clear err status

    DbQuery = False                                                              '☜: Processing is NG

 If   LayerShowHide(1) = False Then
      Exit Function
 End If

    strVal = BIZ_PGM_ID & "?txtMode="          & parent.UID_M0001                       '☜: Query
    strVal = strVal     & "&txtKeyStream="     & lgKeyStream                     '☜: Query Key
    strVal = strVal     & "&txtPrevNext="      & ""                              '☜: Direction
    Call RunMyBizASP(MyBizASP, strVal)                                           '☜:  Run biz logic
 
    DbQuery = True                                                               '☜: Processing is NG
End Function
'========================================================================================================
' Name : DbSave
' Desc : This function is called by FncSave
'========================================================================================================
Function DbSave()
 Dim strVal
    Err.Clear                                                                    '☜: Clear err status
  
	DbSave = False                       '☜: Processing is NG
	 
	If   LayerShowHide(1) = False Then
	     Exit Function
	End If

	With Frm1
	    .txtMode.value        = parent.UID_M0002                                        '☜: Delete
	    .txtFlgMode.value     = lgIntFlgMode
	    .txtKeyStream.Value   = lgKeyStream                                      '☜: Save Key
	End With
 
	Call ExecMyBizASP(frm1, BIZ_PGM_ID)
  
    DbSave  = True                                                               '☜: Processing is NG
End Function
'========================================================================================================
' Name : DbDelete
' Desc : This function is called by FncDelete
'========================================================================================================
Function DbDelete()
	Dim strVal
    Err.Clear                                                                    '☜: Clear err status
  
	DbDelete = False                                                    '☜: Processing is NG
  
	If   LayerShowHide(1) = False Then
	     Exit Function
	End If
	 
    strVal = BIZ_PGM_ID & "?txtMode="          & parent.UID_M0003                       '☜: Query
    strVal = strVal     & "&txtKeyStream="     & lgKeyStream                     '☜: Query Key
    strVal = strVal     & "&txtPrevNext="      & ""                              '☜: Direction

	Call RunMyBizASP(MyBizASP, strVal)                                           '☜: Run Biz logic
 
	DbDelete = True                                                              '⊙: Processing is NG
End Function
'========================================================================================================
' Function Name : DbQueryOk
' Function Desc : Called by MB Area when query operation is successful
'========================================================================================================
Function DbQueryOk()

	lgIntFlgMode      = parent.OPMD_UMODE                                               '⊙: Indicates that current mode is Create mode
    lgBlnFlgChgValue = false

	Call SetToolbar("11111000110111")

    Frm1.txtName.focus
    
    If  txtEmp_no_Onchange() then
        Exit Function
    End If
    
    Call ggoOper.LockField(Document, "Q")
    Set gActiveElement = document.ActiveElement   
	frm1.txtwarnt_insur_nm.focus
End Function
 
'========================================================================================================
' Function Name : DbSaveOk
' Function Desc : Called by MB Area when save operation is successful
'========================================================================================================
Function DbSaveOk()
    Call InitVariables
	Call MainQuery()
End Function
 
'========================================================================================================
' Function Name : DbDeleteOk
' Function Desc : Called by MB Area when delete operation is successful
'========================================================================================================
Function DbDeleteOk()
	Call InitVariables()
	Call ggoOper.ClearField(Document, "2")                                       '☜: Clear Contents  Field
	lgBlnFlgChgValue = false
End Function

'========================================================================================================
' Name : OpenEmpName()
' Desc : developer describe this line 
'========================================================================================================
Function OpenEmpName(iWhere)
	Dim arrRet
	Dim arrParam(2)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True
	   If  iWhere = 0 Then
	    arrParam(0) = ""   ' Code Condition
	    arrParam(1) = frm1.txtName.value   ' Name Cindition
	   Else
	    arrParam(0) = frm1.txtEmp_no.value   ' Code Condition
	    arrParam(1) = ""   ' Name Cindition
	End If
	   arrParam(2) = lgUsrIntCd
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
' Name : SetEmp()
' Description : Item Popup에서 Return되는 값 setting
'=======================================================================================================
Sub SetEmpName(arrRet)
	 With frm1
		.txtEmp_no.value = arrRet(0)
		.txtName.value = arrRet(1)
		Call ggoOper.ClearField(Document, "2")      '☜: Clear Contents  Field
		 
		Set gActiveElement = document.ActiveElement
        call txtEmp_no_Onchange()
		lgBlnFlgChgValue = False
		.txtEmp_no.focus
	 End With
End Sub

'========================================================================================================
' Name : SubOpenCollateralNoPop()
' Desc : developer describe this line Call Master L/C No PopUp
'========================================================================================================
Sub SubOpenCollateralNoPop()
	Dim strRet
	If gblnWinEvent = True Then Exit Sub
	gblnWinEvent = True
	 
	strRet = window.showModalDialog("s1413pa1.asp", "", _
	 "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	gblnWinEvent = False
	 
	If strRet = "" Then
	      Exit Sub
	Else
	      Call SetCollateralNo(strRet)
	End If 
End Sub

'========================================================================================================
' Name : SetCurrency
' Desc : developer describe this line 
'========================================================================================================
Function SetCurrency(arrRet)
 frm1.txtCurrency.Value = arrRet(0)
 lgBlnFlgChgValue = True
End Function

' 보증보험 보험료 
Sub txtWarnt_amt_Change()
 lgBlnFlgChgValue = True
End Sub
' 보증보험 보험기간 
Sub txtWarnt_start_Change()
 lgBlnFlgChgValue = True
End Sub

Sub txtWarnt_end_Change()
 lgBlnFlgChgValue = True
End Sub
' 보증인1 보증기간 
Sub txtWarnt1_start_Change()
 lgBlnFlgChgValue = True
End Sub
Sub txtWarnt1_end_Change()
 lgBlnFlgChgValue = True
End Sub
' 보증인1 갑근세 
Sub txtwarnt1_incom_tax_Change()
 lgBlnFlgChgValue = True
End Sub
' 보증인2 보증기간 
Sub txtWarnt2_start_Change()
 lgBlnFlgChgValue = True
End Sub
Sub txtWarnt2_end_Change()
 lgBlnFlgChgValue = True
End Sub
' 보증인2 갑근세 
Sub txtwarnt2_incom_tax_Change()
 lgBlnFlgChgValue = True
End Sub

'========================================================================================================
' Name : txtWarnt_start_DblClick
' Desc : developer describe this line 
'========================================================================================================
Sub txtWarnt_start_DblClick(Button)
    If Button = 1 Then
   		Call SetFocusToDocument("M")  
        frm1.txtWarnt_start.Action = 7
        frm1.txtWarnt_start.focus
    End If
End Sub

'========================================================================================================
' Name : txtWarnt_end_DblClick
' Desc : developer describe this line 
'========================================================================================================
Sub txtWarnt_end_DblClick(Button)
    If Button = 1 Then
   		Call SetFocusToDocument("M")      
        frm1.txtWarnt_end.Action = 7
        frm1.txtWarnt_end.focus
    End If
End Sub

'========================================================================================================
' Name : txtWarnt1_start_DblClick
' Desc : developer describe this line 
'========================================================================================================
Sub txtWarnt1_start_DblClick(Button)
    If Button = 1 Then
   		Call SetFocusToDocument("M")      
        frm1.txtWarnt1_start.Action = 7
        frm1.txtWarnt1_start.focus
    End If
End Sub

'========================================================================================================
' Name : txtWarnt1_end_DblClick
' Desc : developer describe this line 
'========================================================================================================
Sub txtWarnt1_end_DblClick(Button)
    If Button = 1 Then
   		Call SetFocusToDocument("M")      
        frm1.txtWarnt1_end.Action = 7
        frm1.txtWarnt1_end.focus
    End If
End Sub


'========================================================================================================
' Name : txtWarnt2_start_DblClick
' Desc : developer describe this line 
'========================================================================================================
Sub txtWarnt2_start_DblClick(Button)
    If Button = 1 Then
		Call SetFocusToDocument("M")  
        frm1.txtWarnt2_start.Action = 7
        frm1.txtWarnt2_start.focus
    End If
End Sub

'========================================================================================================
' Name : txtWarnt2_end_DblClick
' Desc : developer describe this line 
'========================================================================================================
Sub txtWarnt2_end_DblClick(Button)
    If Button = 1 Then
   		Call SetFocusToDocument("M")      
        frm1.txtWarnt2_end.Action = 7
        frm1.txtWarnt2_end.focus
    End If
End Sub

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
	frm1.txtName.value = ""
	frm1.txtDept_nm.value = ""
	frm1.txtRoll_pstn.value = ""
	frm1.txtEntr_dt.Text = ""
	frm1.txtPay_grd.value = ""
	Frm1.imgPhoto.src = ""
	
    strVal =""
    Frm1.imgPhoto.src = strVal   

    If  frm1.txtEmp_no.value = "" Then
		frm1.txtEmp_no.value = ""
    Else
     IntRetCd = FuncGetEmpInf2(frm1.txtEmp_no.value,lgUsrIntCd,strName,strDept_nm,_
                 strRoll_pstn, strPay_grd1, strPay_grd2, strEntr_dt, strInternal_cd)
     if  IntRetCd < 0 then
		strVal = "../../../CShared/image/default_picture.jpg"
    	Frm1.imgPhoto.src = strVal     
         if  IntRetCd = -1 then
			 Call DisplayMsgBox("800048","X","X","X") '해당사원은 존재하지 않습니다.
         else
                Call DisplayMsgBox("800454","X","X","X") '자료에 대한 권한이 없습니다.
         end if
        Call ggoOper.ClearField(Document, "2")
           
'        call InitVariables()
        frm1.txtEmp_no.focus
        Set gActiveElement = document.ActiveElement
        txtEmp_no_Onchange = true
    Else
        frm1.txtName.value = strName
        frm1.txtDept_nm.value = strDept_nm
        frm1.txtRoll_pstn.value = strRoll_pstn
        frm1.txtPay_grd.value = strPay_grd1 & "-" & strPay_grd2
        frm1.txtEntr_dt.Text = UNIDateClientFormat(strEntr_dt)
		'strEntr_dt는 Client Format(parent.gClientDateFormat) 그러므로 Client Format -->Company Format

		Call CommonQueryRs(" COUNT(*) "," HAA070T "," emp_no= " & FilterVar( Frm1.txtEmp_no.value, "''", "S") & "" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

			if   Replace(lgF0, Chr(11), "") > 0  then
			strVal = "../../ComASP/CPictRead.asp" & "?txtKeyValue=" & Frm1.txtEmp_no.value '☜: query key
			strVal = strVal     & "&txtDKeyValue=" & "default"                            '☜: default value
			strVal = strVal     & "&txtTable="     & "HAA070T"                            '☜: Table Name
			strVal = strVal     & "&txtField="     & "Photo"	                          '☜: Field
			strVal = strVal     & "&txtKey="       & "Emp_no"	                          '☜: Key
		else
			strVal = "../../../CShared/image/default_picture.jpg"
		end if

    	Frm1.imgPhoto.src = strVal
        End if 
    End if
    
End Function 

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>

<BODY SCROLL="NO" TABINDEX="-1">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_00%>>
 <TR >
  <TD <%=HEIGHT_TYPE_00%>></TD>
 </TR>
 <TR HEIGHT=23><% ' 탭위치 %>
  <TD WIDTH=100%>
   <TABLE <%=LR_SPACE_TYPE_10%>>
    <TR>
     <TD WIDTH=10>&nbsp;</TD>
     <TD CLASS="CLSMTABP">
      <TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
       <TR>
        <TD BACKGROUND"../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" WIDTH="10" HEIGHT="23"></td>
        <TD BACKGROUND="../../../CShared/image/table/seltab_up_bg.gif" CLASS="CLSMTAB" ALIGN="center"><FONT COLOR=white>보증인등록</font></td>
        <TD BACKGROUND="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_right.gif" WIDTH="10" HEIGHT="23"></td>
          </TR>
      </TABLE>
     </TD>
     <TD WIDTH=*>&nbsp;</TD>
    </TR>
   </TABLE>
  </TD>
 </TR><% ' 탭위치 종료 %>
 <TR HEIGHT=*>
  <TD WIDTH=100% CLASS="Tab11">
   <TABLE <%=LR_SPACE_TYPE_20%>>
    <TR>
     <TD <%=HEIGHT_TYPE_02%> WIDTH=100% COLSPAN=2></TD>
    </TR>
    <TR>
                 <TD HEIGHT=20 WIDTH=10%>
                        <img src="../../../CShared/image/default_picture.jpg" name="imgPhoto" WIDTH=80 HEIGHT=90 HSPACE=10 VSPACE=0 BORDER=1>
                 </TD>
                 <TD HEIGHT=20 WIDTH=90%>
                     <FIELDSET CLASS="CLSFLD">
               <TABLE <%=LR_SPACE_TYPE_40%>>
                <TR>
              <TD CLASS="TD5" NOWRAP>사원</TD>
              <TD CLASS="TD6"><INPUT NAME="txtEmp_no" ALT="사원" TYPE="Text" MAXLENGTH=13 SiZE=13 tag=12XXXU><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSItemDC" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenEmpName(1)"></TD>
                 <TD CLASS="TD5" NOWRAP>성명</TD>
              <TD CLASS="TD6"><INPUT NAME="txtName" ALT="성명" TYPE="Text" MAXLENGTH=30 SiZE=20 tag=14></TD>
                </TR>
                <TR>
                 <TD CLASS="TD5" NOWRAP>부서명</TD>
                 <TD CLASS="TD6"><INPUT NAME="txtDept_nm" ALT="부서명" TYPE="Text" SiZE=15 tag=14></TD>
                 <TD CLASS="TD5" NOWRAP>직위</TD>
                 <TD CLASS="TD6"><INPUT NAME="txtRoll_pstn" ALT="직위" TYPE="Text" SiZE=15 tag=14></TD>
                </TR>
                <TR>
                 <TD CLASS="TD5" NOWRAP>입사일</TD>
           <TD CLASS="TD6"><script language =javascript src='./js/h2005ma1_txtEntr_dt_txtEntr_dt.js'></script></TD>
                 <TD CLASS="TD5" NOWRAP>급호</TD>
                 <TD CLASS="TD6"><INPUT NAME="txtPay_grd" ALT="급호" TYPE="Text" SiZE=15 tag=14></TD>
                </TR>
               </TABLE>
            </FIELDSET>
     </TD>     
       </TR>
       <TR>
           <TD <%=LR_SPACE_TYPE_03%> WIDTH=100% COLSPAN=2></TD>
       </TR>
                <TR>
        <TD WIDTH=100% HEIGHT=* VALIGN="TOP" COLSPAN=2>
            <FIELDSET CLASS="CLSFLD"><LEGEND align=left>보증보험</LEGEND>
         <TABLE <%=LR_SPACE_TYPE_60%>>
          <TR>
           <TD CLASS="TD5" NOWRAP>보험명</TD>
           <TD CLASS="TD6"><INPUT NAME="txtwarnt_insur_nm" ALT="보험명" TYPE="Text" Maxlength=20 SiZE=20 tag=21></TD>
           <TD CLASS="TD5" NOWRAP>보험번호</TD>
           <TD CLASS="TD6"><INPUT NAME="txtWarnt_insur_no" ALT="보험번호" TYPE="Text" Maxlength=20 SiZE=20 tag=21></TD>
          </TR>
          <TR>
           <TD CLASS="TD5" NOWRAP>보험사명</TD>
           <TD CLASS="TD6"><INPUT NAME="txtWarnt_insur_comp" ALT="보험사명" TYPE="Text" Maxlength=20 SiZE=20 tag=21></TD>
           <TD CLASS="TD5" NOWRAP>보험료</TD>
           <TD CLASS="TD6"><script language =javascript src='./js/h2005ma1_txtWarnt_amt_txtWarnt_amt.js'></script></TD>
          </TR>
          <TR>
           <TD CLASS="TD5" NOWRAP>보증기간</TD>
           <TD CLASS="TD6">
               <script language =javascript src='./js/h2005ma1_txtWarnt_start_txtWarnt_start.js'></script>&nbsp; - &nbsp;
               <script language =javascript src='./js/h2005ma1_txtWarnt_end_txtWarnt_end.js'></script>
           </TD>
           <TD CLASS="TD5" NOWRAP></TD>
           <TD CLASS="TD6"></TD>
          </TR>
            </TABLE>
            </FIELDSET>
        </TD>
       </TR>
       <TR>
        <TD WIDTH=100% HEIGHT=* VALIGN="TOP" COLSPAN=2>
            <FIELDSET CLASS="CLSFLD"><LEGEND align=left>보증인1</LEGEND>
                        <TABLE <%=LR_SPACE_TYPE_60%>>
          <TR>
           <TD CLASS="TD5" NOWRAP>성명</TD>
           <TD CLASS="TD6"><INPUT NAME="txtWarnt1_name" ALT="성명" Maxlength=30 TYPE="Text" SiZE=30 tag=21></TD>
           <TD CLASS="TD5" NOWRAP>근무지</TD>
           <TD CLASS="TD6"><INPUT NAME="txtWarnt1_comp_nm" ALT="근무지" Maxlength=30 TYPE="Text" SiZE=30 tag=21></TD>
          </TR>
          <TR>
           <TD CLASS="TD5" NOWRAP>관계</TD>
           <TD CLASS="TD6"><INPUT NAME="txtWarnt1_rel" ALT="관계" TYPE="Text" Maxlength=20 SiZE=20 tag=21></TD>
           <TD CLASS="TD5" NOWRAP>직위</TD>
           <TD CLASS="TD6"><INPUT NAME="txtWarnt1_roll_pstn" ALT="직위" TYPE="Text" Maxlength=10 SiZE=10 tag=21></TD>
          </TR>
          <TR>
           <TD CLASS="TD5" NOWRAP>주민번호</TD>
           <TD CLASS="TD6"><INPUT NAME="txtWarnt1_res_no" ALT="주민번호" Maxlength=14 TYPE="Text" SiZE=15 tag=21></TD>
           <TD CLASS="TD5" NOWRAP>갑근세</TD>
           <TD CLASS="TD6"><script language =javascript src='./js/h2005ma1_txtwarnt1_incom_tax_txtwarnt1_incom_tax.js'></script></TD>
          </TR>
          <TR>
           <TD CLASS="TD5" NOWRAP>보증기간</TD>
           <TD CLASS="TD6"><script language =javascript src='./js/h2005ma1_txtWarnt1_start_txtWarnt1_start.js'></script>&nbsp; - &nbsp;
                           <script language =javascript src='./js/h2005ma1_txtWarnt1_end_txtWarnt1_end.js'></script></TD>
           <TD CLASS="TD5" NOWRAP>주소</TD>
           <TD CLASS="TD6"><INPUT NAME="txtWarnt1_addr" ALT="주소" TYPE="Text" MAXLENGTH=60 SiZE=45 tag=21></TD>
          </TR>
            </TABLE>
            </FIELDSET>
        </TD>
       </TR>
       <TR>
        <TD WIDTH=100% HEIGHT=* VALIGN="TOP" COLSPAN=2>
            <FIELDSET CLASS="CLSFLD"><LEGEND align=left>보증인2</LEGEND>
                        <TABLE <%=LR_SPACE_TYPE_60%>>
          <TR>
           <TD CLASS="TD5" NOWRAP>성명</TD>
           <TD CLASS="TD6"><INPUT NAME="txtWarnt2_name" ALT="성명" MAXLENGTH=30 TYPE="Text" SiZE=30 tag=21></TD>
           <TD CLASS="TD5" NOWRAP>근무지</TD>
           <TD CLASS="TD6"><INPUT NAME="txtWarnt2_comp_nm" ALT="근무지" MAXLENGTH=30 TYPE="Text" SiZE=30 tag=21></TD>
          </TR>
          <TR>
           <TD CLASS="TD5" NOWRAP>관계</TD>
           <TD CLASS="TD6"><INPUT NAME="txtWarnt2_rel" ALT="관계" MAXLENGTH=20 TYPE="Text" SiZE=20 tag=21></TD>
           <TD CLASS="TD5" NOWRAP>직위</TD>
           <TD CLASS="TD6"><INPUT NAME="txtWarnt2_roll_pstn" ALT="직위" TYPE="Text" MAXLENGTH=10 SiZE=10 tag=21></TD>
          </TR>
          <TR>
           <TD CLASS="TD5" NOWRAP>주민번호</TD>
           <TD CLASS="TD6"><INPUT NAME="txtWarnt2_res_no" ALT="주민번호" Maxlength=14 TYPE="Text" SiZE=15 tag=21></TD>
           <TD CLASS="TD5" NOWRAP>갑근세</TD>
           <TD CLASS="TD6"><script language =javascript src='./js/h2005ma1_txtwarnt2_incom_tax_txtwarnt2_incom_tax.js'></script></TD>
          </TR>
          <TR>
           <TD CLASS="TD5" NOWRAP>보증기간</TD>
           <TD CLASS="TD6"><script language =javascript src='./js/h2005ma1_txtWarnt2_start_txtWarnt2_start.js'></script>&nbsp; - &nbsp;
                           <script language =javascript src='./js/h2005ma1_txtWarnt2_end_txtWarnt2_end.js'></script></TD>
           <TD CLASS="TD5" NOWRAP>주소</TD>
           <TD CLASS="TD6"><INPUT NAME="txtWarnt2_addr" ALT="주소" TYPE="Text" MAXLENGTH=60 SiZE=45 tag=21></TD>
          </TR>
         </TABLE>
         </FIELDSET>
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
            <TD WIDTH="*" ALIGN=RIGHT><a href = "VBSCRIPT:PgmJump(BIZ_PGM_JUMP_ID)" ONCLICK="VBSCRIPT:CookiePage 1">인사마스타</a></TD>
                 <TD WIDTH=10>&nbsp;</TD>
             </TR>
         </TABLE>
     </TD>
 </TR>
 <TR>
  <TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=no noresize framespacing=0></IFRAME>
 </TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode"        TAG="24">
<INPUT TYPE=HIDDEN NAME="txtKeyStream"   TAG="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId"  TAG="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode"     TAG="24">
<INPUT TYPE=HIDDEN NAME="txtPrevNext"    TAG="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
 <IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm"></IFRAME>
</DIV>
</BODY>
</HTML>

