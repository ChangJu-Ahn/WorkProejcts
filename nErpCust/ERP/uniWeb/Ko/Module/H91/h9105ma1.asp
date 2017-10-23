<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : Human Resources
*  2. Function Name        : 
*  3. Program ID           : h9105ma1
*  4. Program Name         : h9105ma1
*  5. Program Desc         : 연말정산관리/연말정산/종전근무지등록 
*  6. Comproxy List        :
*  7. Modified date(First) : 2001/06/05
*  8. Modified date(Last)  : 2003/06/13
*  9. Modifier (First)     : mok young bin
* 10. Modifier (Last)      : Lee SiNa
* 11. Comment              :
=======================================================================================================-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/IncHRQuery.vbs">   </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/IncCliRdsQuery.vbs">   </SCRIPT>
<SCRIPT LANGUAGE="JavaScript"   SRC="../../inc/TabScript.js"></SCRIPT>
<Script Language="VBScript">
Option Explicit
'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
Const BIZ_PGM_ID      = "h9105mb1.asp"						           '☆: Biz Logic ASP Name
Const TAB1 = 1
Const TAB2 = 2
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

Dim IsOpenPop						                                    ' Popup
Dim lsInternal_cd

'========================================================================================================
' Name : InitVariables()	
' Desc : Initialize value
'========================================================================================================
Sub InitVariables()
	lgIntFlgMode      = parent.OPMD_CMODE						        '⊙: Indicates that current mode is Create mode
	lgBlnFlgChgValue  = False								    '⊙: Indicates that no value changed
	lgIntGrpCount     = 0										'⊙: Initializes Group View Size
    lgStrPrevKey      = ""                                      '⊙: initializes Previous Key
    lgSortKey         = 1                                       '⊙: initializes sort direction
		
	gblnWinEvent      = False
	lgBlnFlawChgFlg   = False
End Sub

'========================================================================================================
' Name : SetDefaultVal()	
' Desc : Set default value
'========================================================================================================
	
Sub SetDefaultVal()
    Dim strYear,strMonth,strDay
    
    frm1.txtYear_yy.focus	
    Call ggoOper.FormatDate(frm1.txtYear_yy, parent.gDateFormat, 3)
    Call ExtractDateFrom("<%=GetSvrDate%>",parent.gServerDateFormat,parent.gServerDateType,strYear,strMonth,strDay)    
    frm1.txtYear_yy.Year	= strYear
    frm1.txtYear_yy.Month	= strMonth
    frm1.txtYear_yy.Day		= strDay
End Sub
	
'========================================================================================================
' Name : LoadInfTB19029()	
' Desc : Set System Number format
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "H", "NOCOOKIE","MA") %>
End Sub

'========================================================================================================
'	Name : CookiePage()
'	Description : Item Popup에서 Return되는 값 setting
'========================================================================================================
Function CookiePage(ByVal flgs)
	On Error Resume Next

	Const CookieSplit = 4877						
	Dim strTemp

	If flgs = 1 Then
	ElseIf flgs = 0 Then

		strTemp = ReadCookie(CookieSplit)
		If strTemp = "" then 
		    Exit Function
		End if	
		If Err.number <> 0 Then
			Err.Clear
			WriteCookie CookieSplit , ""
			Exit Function 
		End If

		WriteCookie CookieSplit , ""
		Call MainQuery()
			
	End If

End Function


'========================================================================================================
' Function Name : MakeKeyStream
' Function Desc : This method set focus to pos of err
'========================================================================================================
Sub MakeKeyStream(pOpt)
   
    Dim strYYYYMM
    Dim strYear,strMonth,strDay
    
    lgKeyStream = Frm1.txtEmp_no.value & parent.gColSep                                           'You Must append one character(parent.gColSep)    
    lgKeyStream = lgKeyStream & frm1.txtYear_yy.year & parent.gColSep
    lgKeyStream = lgKeyStream & gSelframeFlg & parent.gColSep
    
    
    Call CommonQueryRs(" count(seq) "," hfa040t ","EMP_NO =  " & FilterVar(frm1.txtEmp_no.value, "''", "S") & " and year_yy= " & FilterVar(frm1.txtYear_yy.Year, "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
        lgKeyStream = lgKeyStream & Trim(Replace(lgF0,Chr(11),"")) & parent.gColSep
End Sub        

'========================================================================================================
' Name : Form_Load
' Desc : developer describe this line Called by Window_OnLoad() evnt
'========================================================================================================
Sub Form_Load()

    Err.Clear                                                                       '☜: Clear err status
	Call LoadInfTB19029                                                             '☜: Load table , B_numeric_format
		
	Call AppendNumberPlace("6", "3", "0")

    Call ggoOper.FormatField(Document, "A",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
        
    Call ggoOper.FormatNumber(frm1.txtA_comp_no2,9999999999,0,false)
    Call ggoOper.FormatNumber(frm1.txtA_comp_no,9999999999,0,false)
    
    frm1.txtA_comp_no.value=""
    frm1.txtA_comp_no2.value=""
   

	Call ggoOper.LockField(Document, "N")											'⊙: Lock Field
	
    Call SetDefaultVal()
	Call SetToolbar("1110100000001111")												'⊙: Set ToolBar
	
	Call InitVariables
	
    Call FuncGetAuth(gStrRequestMenuID, parent.gUsrID, lgUsrIntCd)                                ' 자료권한:lgUsrIntCd ("%", "1%")
	
    gSelframeFlg = TAB1
    Call changeTabs(TAB1)
    gIsTab     = "Y" ' <- "Yes"의 약자 Y(와이) 입니다.[V(브이)아닙니다]
    gTabMaxCnt = 2   ' Tab의 갯수를 적어 주세요    

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
    
    FncQuery = False															 '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgbox("900013", parent.VB_YES_NO,"x","x")					 '☜: Data is changed.  Do you want to display it? 
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    Call ggoOper.ClearField(Document, "2")										 '☜: Clear Contents  Field
    
    If Not chkField(Document, "1") Then									         '☜: This function check required field
       Exit Function
    End If
	
    If txtEmp_no_Onchange() Then        'enter key 로 조회시 사원을 check후 해당사항 없으면 query종료...
        Exit Function
    End if
   
    Call InitVariables                                                           '⊙: Initializes local global variables
    Call MakeKeyStream("Q")
    
    If DbQuery = False Then  
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
    
    If lgBlnFlgChgValue = True Then
       IntRetCD = DisplayMsgbox("900015", parent.VB_YES_NO,"x","x")					 '☜: Data is changed.  Do you want to make it new? 
       If IntRetCD = vbNo Then
          Exit Function
       End If
    End If
    
    Call ggoOper.ClearField(Document, "A")                                       '☜: Clear Condition Field
    Call ggoOper.LockField(Document , "N")                                       '☜: Lock  Field
   
    Call SetToolbar("11001000000011")
    Call InitVariables                                                        '⊙: Initializes local global variables
    
    Set gActiveElement = document.ActiveElement   
    
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
    
    If lgIntFlgMode <> parent.OPMD_UMODE Then                                           '☜: Please do Display first. 
        Call DisplayMsgbox("900002","x","x","x")                                
        Exit Function
    End If
    
    IntRetCD = DisplayMsgbox("900003", parent.VB_YES_NO,"x","x")                        '☜: Do you want to delete? 
	If IntRetCD = vbNo Then
		Exit Function
	End If
    
    Call MakeKeyStream("D")
    If DbDelete = False Then
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
    Dim iDx
    Dim arrACompNo
    Dim strACompNo
    Dim arrBuff(10)
    Dim sumAComp
    Dim intTenMinusSum
	
    FncSave = False                                                              '☜: Processing is NG
    
    Err.Clear                                                                    '☜: Clear err status
    
    If lgBlnFlgChgValue = False Then 
        IntRetCD = DisplayMsgbox("900001","x","x","x")                           '☜:There is no changed data. 
        Exit Function
    End If
    
    If Not ChkFieldLength(Document, "2") Then									         '☜: This function check required field
       Exit Function
    End If
    
    If gSelframeFlg = 1 Then         'tab1 일때 필수입력 check..
       If Trim(frm1.txtA_comp_nm.value) = "" Then
          Call DisplayMsgbox("970021","x","회사명","x")
          frm1.txtA_comp_nm.focus        
          Set gActiveElement = document.activeElement
          Exit Function 
       End If
       If Trim(frm1.txtA_comp_no.value) = "" Then
          Call DisplayMsgbox("970021","x","사업자등록번호","x")
          frm1.txtA_comp_no.focus        
          Set gActiveElement = document.activeElement
          Exit Function 
       End If
    Else                              'tab2 일때 필수입력 check..
       If Trim(frm1.txtA_comp_nm2.value) = "" Then
          Call DisplayMsgbox("970021","x","회사명","x")
          frm1.txtA_comp_nm2.focus        
          Set gActiveElement = document.activeElement
          Exit Function 
       End If
       If Trim(frm1.txtA_comp_no2.value) = "" Then
          Call DisplayMsgbox("970021","x","사업자등록번호","x")
          frm1.txtA_comp_no2.focus        
          Set gActiveElement = document.activeElement
          Exit Function 
       End If
    End if

    Call CommonQueryRs(" count(seq) "," hfa040t ","EMP_NO =  " & FilterVar(frm1.txtEmp_no.value, "''", "S") & " and year_yy= " & FilterVar(frm1.txtYear_yy.Year, "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    
    If lgIntFlgMode = parent.OPMD_CMODE Then                 'insert 시에...
        If gSelframeFlg = 1 Then                      'tab1 일때 
            If Trim(Replace(lgF0,Chr(11),"")) = 0 Then         '(사번과 정산년도로 조회)종전근무지가 없을때...
            ElseIf Trim(Replace(lgF0,Chr(11),"")) = 1 Then     '종전근무지가 하나일때 
                                                               '종전근무지1이 존재합니다. 종전근무지2에서 입력하시오'
                Call DisplayMsgbox("800247","x","X","x")         '조회작업을 선행한 후 저장하십시오.
                Exit Function
            Else                                               '종전근무지가 두개일때 
            '종전근무지1,2가 모두 존재합니다.조회를 선행하십시요 
                Call DisplayMsgbox("800247","x","X","x")         '조회작업을 선행한 후 저장하십시오.
                Exit Function
            End if
        Else                                          'tab2일때 
            If Trim(Replace(lgF0,Chr(11),"")) = 0 Then          '(사번과 정산년도로 조회)종전근무지가 없을때...
                Call DisplayMsgbox("800435","x","X","x")         '종전근무지1을 먼저 입력하십시오.
                Call ggoOper.ClearField(Document, "2")                                       '⊙: Clear Condition Field
                Exit Function 
            ElseIf Trim(Replace(lgF0,Chr(11),"")) = 1 Then      '종전근무지가 하나일때 
            Else                                                '종전근무지가 두개일때 
            '종전근무지1,2가 모두 존재합니다. 조회를 선행하십시요 
                Call DisplayMsgbox("800247","x","X","x")         '조회작업을 선행한 후 저장하십시오.
                Exit Function
            End if    
        End if
    ElseIf lgIntFlgMode = parent.OPMD_UMODE Then                                               'Update 시에 
        If gSelframeFlg = 1 Then                      'tab1 일때 
            If Trim(Replace(lgF0,Chr(11),"")) = 0 Then         '(사번과 정산년도로 조회)종전근무지가 없을때...
                lgIntFlgMode = parent.OPMD_CMODE 
            ElseIf Trim(Replace(lgF0,Chr(11),"")) = 1 Then     '종전근무지가 하나일때 
            Else                                               '종전근무지가 두개일때 
            End if
        Else                                          'tab2일때 
            If Trim(Replace(lgF0,Chr(11),"")) = 0 Then          '(사번과 정산년도로 조회)종전근무지가 없을때...
                Call DisplayMsgbox("800435","x","X","x")        '종전근무지1을 먼저 입력하십시오.
                Call ggoOper.ClearField(Document, "2")                                       '⊙: Clear Condition Field
                Exit Function 
            ElseIf Trim(Replace(lgF0,Chr(11),"")) = 1 Then      '종전근무지가 하나일때 
                lgIntFlgMode = parent.OPMD_CMODE             
            Else                                                '종전근무지가 두개일때 
            End if    
        End if    
    End if
    
    arrACompNo = Array(1, 3, 7, 1, 3, 7, 1, 3, 5)
    
    If gSelframeFlg = 1 Then         'tab1 일때 필수입력 check..

        If len(frm1.txtA_comp_no.value) <> 10 Then
            Call DisplayMsgbox("800436","X","X","X")	'잘못된 사업자등록번호를 입력하셨습니다.
            frm1.txtA_comp_no.value = ""
            frm1.txtA_comp_no.focus
            Set gActiveElement = document.ActiveElement
            Exit Function
        End if
        
        strACompNo = cdbl(frm1.txtA_comp_no.value)
        
        for iDx = 0 to 8
         	 arrBuff(iDx) = (mid(strACompNo,Cint(iDx)+1,1) * arrACompNo(iDx)) mod 10
         	 If iDx = 8 then
         	   arrBuff(iDx) =  (mid(strACompNo,Cint(iDx)+1,1) * arrACompNo(iDx))\10
         	   arrBuff(Cint(iDx) + 1) = (CInt(mid(strACompNo,Cint(iDx)+1,1)) * arrACompNo(iDx)) mod 10
         	 End if
        next
                
        for iDx = 0 to 9
         	 sumAComp = Cint(sumAComp) + Cint(arrBuff(iDx))
        next
                
        intTenMinusSum = 10 -(Cint(sumAComp) mod 10)
                
        If Cint(intTenMinusSum) = 10 then
            intTenMinusSum = 0
        End if
                
        If Cint(intTenMinusSum) = Cint(mid(strACompNo,10,1)) Then
        Else
            Call DisplayMsgbox("800436","X","X","X")	'잘못된 사업자등록번호를 입력하셨습니다.
            frm1.txtA_comp_no.value = ""
            frm1.txtA_comp_no.focus
            Set gActiveElement = document.ActiveElement
            Exit Function
        End if
    Else
        If len(frm1.txtA_comp_no2.value) <> 10 Then
            Call DisplayMsgbox("800436","X","X","X")	'잘못된 사업자등록번호를 입력하셨습니다.
            frm1.txtA_comp_no2.value = ""
            frm1.txtA_comp_no2.focus
            Set gActiveElement = document.ActiveElement
            Exit Function
        End if

        strACompNo = cdbl(frm1.txtA_comp_no2.value)
                
        for iDx = 0 to 8
         	 arrBuff(iDx) = (mid(strACompNo,Cint(iDx)+1,1) * arrACompNo(iDx)) mod 10
         	 If iDx = 8 then
         	   arrBuff(iDx) =  (mid(strACompNo,Cint(iDx)+1,1) * arrACompNo(iDx))\10
         	   arrBuff(Cint(iDx) + 1) = (CInt(mid(strACompNo,Cint(iDx)+1,1)) * arrACompNo(iDx)) mod 10
         	 End if
        next
                
        for iDx = 0 to 9
         	 sumAComp = Cint(sumAComp) + Cint(arrBuff(iDx))
        next
                
        intTenMinusSum = 10 -(Cint(sumAComp) mod 10)
                
        If Cint(intTenMinusSum) = 10 then
            intTenMinusSum = 0
        End if
                
        If Cint(intTenMinusSum) = Cint(mid(strACompNo,10,1)) Then
        Else
            Call DisplayMsgbox("800436","X","X","X")	'잘못된 사업자등록번호를 입력하셨습니다.
            frm1.txtA_comp_no2.value = ""
            frm1.txtA_comp_no2.focus
            Set gActiveElement = document.ActiveElement
            Exit Function
        End if
    End if 
    
    Call MakeKeyStream("S")
    If DbSave = False Then
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
	
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgbox("900017", parent.VB_YES_NO,"x","x")				     '☜: Data is changed.  Do you want to continue? 
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    lgIntFlgMode = parent.OPMD_CMODE												     '⊙: Indicates that current mode is Crate mode
    
    Call ggoOper.ClearField(Document, "1")                                       '⊙: Clear Condition Field
    Call ggoOper.LockField(Document, "N")									     '⊙: This function lock the suitable field
    Call SetToolbar("11011000000011")

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
        Call DisplayMsgbox("900002","x","x","x")
        Exit Function
    End If
	
	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgbox("900017", parent.VB_YES_NO,"x","x")					 '☜: Will you destory previous data
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If
	
    Call MakeKeyStream("P")
    Call ggoOper.ClearField(Document, "2")										 '⊙: Clear Contents Area
    
    Call InitVariables														 '⊙: Initializes local global variables

    if LayerShowHide(1) = false then
	    Exit Function
	end if

    strVal = BIZ_PGM_ID & "?txtMode="          & parent.UID_M0001                       '☜: Query
    strVal = strVal     & "&txtKeyStream="     & lgKeyStream                     '☜: Query Key
    strVal = strVal     & "&txtPrevNext="      & "P"	                         '☜: Direction

	Call RunMyBizASP(MyBizASP, strVal)										     '☜: Run Biz 

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
        Call DisplayMsgbox("900002","x","x","x")
        Exit Function
    End If
	
	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgbox("900017", parent.VB_YES_NO,"x","x")					 '☜: Will you destory previous data
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If
	
    Call MakeKeyStream("N")

    Call ggoOper.ClearField(Document, "2")										 '⊙: Clear Contents Area
    
    Call InitVariables														     '⊙: Initializes local global variables

     if LayerShowHide(1) = false then
	    Exit Function
	end if


    strVal = BIZ_PGM_ID & "?txtMode="          & parent.UID_M0001                       '☜: Query
    strVal = strVal     & "&txtKeyStream="     & lgKeyStream                     '☜: Query Key
    strVal = strVal     & "&txtPrevNext="      & "N"	                         '☜: Direction

	Call RunMyBizASP(MyBizASP, strVal)										     '☜: Run Biz 

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
		IntRetCD = DisplayMsgbox("900016", parent.VB_YES_NO,"x","x")			'⊙: Data is changed.  Do you want to exit? 
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

     if LayerShowHide(1) = false then
	    Exit Function
	end if

    strVal = BIZ_PGM_ID & "?txtMode="          & parent.UID_M0001                       '☜: Query
    strVal = strVal     & "&txtKeyStream="     & lgKeyStream                     '☜: Query Key
    strVal = strVal     & "&txtPrevNext="      & ""	                             '☜: Direction

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
		
	DbSave = False														         '☜: Processing is NG
	 if LayerShowHide(1) = false then
	    Exit Function
	end if
		
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
		
	DbDelete = False			                                                 '☜: Processing is NG
		
	 if LayerShowHide(1) = false then
	    Exit Function
	end if
		
    strVal = BIZ_PGM_ID & "?txtMode="          & parent.UID_M0003                       '☜: Query
    strVal = strVal     & "&txtKeyStream="     & lgKeyStream                     '☜: Query Key
    strVal = strVal     & "&txtPrevNext="      & ""	                             '☜: Direction

	Call RunMyBizASP(MyBizASP, strVal)                                           '☜: Run Biz logic
	
	DbDelete = True                                                              '⊙: Processing is NG
End Function
'========================================================================================================
' Function Name : DbQueryOk
' Function Desc : Called by MB Area when query operation is successful
'========================================================================================================
Function DbQueryOk()
    Dim strVal

	lgIntFlgMode      = parent.OPMD_UMODE                                               '⊙: Indicates that current mode is Create mode
    lgBlnFlgChgValue = false

	Call SetToolbar("1101100000011111")
    Call ggoOper.LockField(Document, "Q")
    Set gActiveElement = document.ActiveElement   
	If gSelframeFlg = TAB1 Then	
		frm1.txtA_comp_nm.focus
	else
'		frm1.txtA_comp_nm2.focus
	end if
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
    Call MainQuery()
End Function

'========================================================================================================
' Name : PgmJump1(PGM_JUMP_ID)
' Desc : developer describe this line 
'========================================================================================================

Function PgmJump1(PGM_JUMP_ID)
    Call CookiePage(1)  ' Write Cookie
    PgmJump(PGM_JUMP_ID)
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
			.txtDept_cd.value = arrRet(2)
			.txtRoll_pstn.value = arrRet(3)
			.txtEntr_dt.Text = arrRet(5)
			.txtPay_grd.value = arrRet(4)
			.txtEmp_no.focus
		End If
	End With
End Sub

Function ClickTab1()
    Dim IntRetCD

	If gSelframeFlg = TAB1 Then Exit Function
	
    If lgBlnFlgChgValue = True AND gSelframeFlg = TAB2 Then
		IntRetCD = DisplayMsgbox("800442", parent.VB_YES_NO,"x","x")	 '☜:(종전근무지2)데이터가 변경되었습니다.저장하시겠습니까?
		If IntRetCD = vbNo Then
		    lgBlnFlgChgValue = false
			Call changeTabs(TAB1)
			gSelframeFlg = TAB1
		Else
			gSelframeFlg = TAB2
		    Call FncSave()
		End If
    Else
		Call changeTabs(TAB1)
		gSelframeFlg = TAB1
    End If

End Function

Function ClickTab2()
    Dim IntRetCD
    
    Call CommonQueryRs(" count(seq) "," hfa040t ","EMP_NO =  " & FilterVar(frm1.txtEmp_no.value, "''", "S") & " and year_yy= " & FilterVar(frm1.txtYear_yy.Year, "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    
	If gSelframeFlg = TAB2 Then Exit Function
	
    If Trim(Replace(lgF0,Chr(11),"")) = 0 Then
        Call DisplayMsgbox("800435","x","X","x")                    '종전근무지1을 먼저 입력하시요.
        Exit Function 
    End if    
    
    If lgBlnFlgChgValue = True AND gSelframeFlg = TAB1  Then
		IntRetCD = DisplayMsgbox("800442", parent.VB_YES_NO,"X","x")	 '☜:(종전근무지1)데이터가 변경되었습니다.저장하시겠습니까?
		If IntRetCD = vbNo Then
		    lgBlnFlgChgValue = false
			Call changeTabs(TAB2)
			gSelframeFlg = TAB2
		Else
			gSelframeFlg = TAB1
		    Call FncSave()
		End If
    Else

		Call changeTabs(TAB2)
		gSelframeFlg = TAB2
    End If
	
End Function

'******************************  3.2.1 Object Tag 처리  *********************************************
'	Window에 발생 하는 모든 Even 처리	
'****************************************************************************************************

Sub txtA_comp_nm_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtA_comp_no_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtA_pay_tot_amt_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtA_bonus_tot_amt_Change()
	lgBlnFlgChgValue = True
End Sub
                					        
Sub txtA_after_bonus_amt_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtA_med_insur_amt_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtA_national_pension_amt_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtA_save_tax_sub_amt_Change()
	lgBlnFlgChgValue = True
End Sub
                					        
Sub txtA_indiv_anu_amt_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtA_indiv_anu2_amt_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtA_income_tax_amt_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtA_res_tax_amt_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtA_farm_tax_amt_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtA_non_tax1_amt_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtA_non_tax2_amt_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtA_non_tax3_amt_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtA_non_tax4_amt_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtA_non_tax5_amt_Change()
	lgBlnFlgChgValue = True
End Sub




Sub txtA_comp_nm2_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtA_comp_no2_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtA_pay_tot_amt2_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtA_bonus_tot_amt2_Change()
	lgBlnFlgChgValue = True
End Sub
                					        
Sub txtA_after_bonus_amt2_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtA_med_insur_amt2_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtA_national_pension_amt2_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtA_save_tax_sub_amt2_Change()
	lgBlnFlgChgValue = True
End Sub
                					        
Sub txtA_indiv_anu_amt2_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtA_indiv_anu2_amt2_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtA_income_tax_amt2_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtA_res_tax_amt2_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtA_farm_tax_amt2_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtA_non_tax1_amt2_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtA_non_tax2_amt2_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtA_non_tax3_amt2_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtA_non_tax4_amt2_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtA_non_tax5_amt2_Change()
	lgBlnFlgChgValue = True
End Sub
'========================================================================================================
'   Event Name : txtEmp_no_change             '<==인사마스터에 있는 사원인지 확인 
'   Event Desc :
'========================================================================================================
Function txtEmp_no_Onchange()
    Dim IntRetCd
    Dim strVal
    Dim RetStatus
    Dim strName
    Dim strDept_nm
    Dim strRoll_pstn
    Dim strPay_grd1
    Dim strPay_grd2
    Dim strEntr_dt
    Dim strInternal_cd

    If frm1.txtEmp_no.value = "" Then
		frm1.txtName.value = ""
		frm1.txtDept_cd.value = ""
		frm1.txtRoll_pstn.value = ""
		frm1.txtEntr_dt.text = ""
		frm1.txtPay_grd.value = ""
    Else
	    IntRetCd = FuncGetEmpInf2(frm1.txtEmp_no.value,lgUsrIntCd,strName,strDept_nm,_
	                strRoll_pstn, strPay_grd1, strPay_grd2, strEntr_dt, strInternal_cd)
	    if  IntRetCd < 0 then
	        if  IntRetCd = -1 then
    			Call DisplayMsgbox("800048","X","X","X")	'해당사원은 존재하지 않습니다.
            else
                Call DisplayMsgbox("800454","X","X","X")	'자료에 대한 권한이 없습니다.
            end if
			
			Call ggoOper.ClearField(Document, "2") 
			           
		    frm1.txtName.value = ""
		    frm1.txtDept_cd.value = ""
		    frm1.txtRoll_pstn.value = ""
		    frm1.txtEntr_dt.text = ""
		    frm1.txtPay_grd.value = ""
            frm1.txtEmp_no.focus
            Set gActiveElement = document.ActiveElement
            txtEmp_no_Onchange = true
            Exit Function      
        Else
            frm1.txtName.value = strName
            frm1.txtDept_cd.value = strDept_nm
            frm1.txtRoll_pstn.value = strRoll_pstn
            frm1.txtPay_grd.value = strPay_grd1 & "-" & strPay_grd2
            frm1.txtEntr_dt.text = UNIDateClientFormat(strEntr_dt)
        End if 
    End if  
    
End Function 


'=======================================
'   Event Name :txtYear_yy_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================

Sub txtYear_yy_DblClick(Button) 
    If Button = 1 Then
		Call SetFocusToDocument("M")    
        frm1.txtYear_yy.Action = 7
        frm1.txtYear_yy.focus
    End If
End Sub
'=======================================================================================================
'   Event Name : txtFr_year_Keypress(Key)
'   Event Desc : enter key down시에 조회한다.
'=======================================================================================================
Sub txtYear_yy_Keypress(Key)
    If Key = 13 Then
        FncQuery()
    End If
End Sub

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>

<BODY SCROLL="no" TABINDEX="-1">
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
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 ONCLICK="ClickTab1()">
							<TR>
								<td background="../../../Cshared/Image/table/seltab_up_bg.gif"><img src="../../../Cshared/Image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../Cshared/Image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>종전근무지1</font></td>
								<td background="../../../Cshared/Image/table/seltab_up_bg.gif" align="right"><img src="../../../Cshared/Image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 ONCLICK="ClickTab2()">
							<TR>
								<td background="../../../Cshared/Image/table/tab_up_bg.gif"><img src="../../../Cshared/Image/table/tab_up_left.gif" width="9" height="23"></td>
								<td background="../../../Cshared/Image/table/tab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>종전근무지2</font></td>
								<td background="../../../Cshared/Image/table/tab_up_bg.gif" align="right"><img src="../../../Cshared/Image/table/tab_up_right.gif" width="10" height="23"></td>
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
					<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
				</TR>
				<TR>
    	            <TD HEIGHT=20 WIDTH=100%>
    	                <FIELDSET CLASS="CLSFLD">
			            <TABLE <%=LR_SPACE_TYPE_40%>>
			            	<TR>
								<TD CLASS=TD5 NOWRAP>정산년도</TD>
								<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/h9105ma1_fpDateTime_txtYear_yy.js'></script></TD>
			    	    		<TD CLASS="TD5" NOWRAP>사원</TD>
			    	    		<TD CLASS="TD6"><INPUT NAME="txtEmp_no" ALT="사원" TYPE="Text" SiZE=13 MAXLENGTH=13  tag="12XXXU"><IMG SRC="../../../Cshared/Image/btnPopup.gif" NAME="btnSItemDC" align=top TYPE="BUTTON" ONCLICK="VBScript: OpenEmptName('0')">
			    	    		                <INPUT NAME="txtName" ALT="성명" TYPE="Text" SiZE=20 MAXLENGTH=30  tag="14XXXU"></TD>
			            	</TR>
			            	<TR>
			            		<TD CLASS="TD5" NOWRAP>부서명</TD>
			            		<TD CLASS="TD6"><INPUT NAME="txtDept_cd" ALT="부서명" TYPE="Text" SiZE=20  tag="14XXXU"></TD>
			            		<TD CLASS="TD5" NOWRAP>직위</TD>
			            		<TD CLASS="TD6"><INPUT NAME="txtRoll_pstn" ALT="직위" TYPE="Text" SiZE=15  tag="14XXXU"></TD>
			            	</TR>
			            	<TR>
			            		<TD CLASS="TD5" NOWRAP>입사일</TD>
			            		<TD CLASS="TD6"><script language =javascript src='./js/h9105ma1_fpDateTime2_txtEntr_dt.js'></script></TD>
			            		<TD CLASS="TD5" NOWRAP>급호</TD>
			            		<TD CLASS="TD6"><INPUT NAME="txtPay_grd" ALT="급호" TYPE="Text" SiZE=15  tag="14XXXU"></TD>
			            	</TR>
			            </TABLE>
			    	    </FIELDSET>
			        </TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
            
   	            <TR>
   	                <TD WIDTH=100% VALIGN="TOP" HEIGHT="*">
					    <DIV ID="TabDiv" STYLE="FLOAT: left; HEIGHT: 100%; WIDTH: 100%;" SCROLL="NO">
                        
						<FIELDSET CLASS="CLSFLD"><LEGEND ALIGN=LEFT>종전근무지1</LEGEND>
                        <TABLE <%=LR_SPACE_TYPE_60%>>
						    <TR>
							    <TD CLASS=TD5 NOWRAP>회사명</TD>
							    <TD CLASS=TD6 NOWRAP><INPUT NAME="txtA_comp_nm"  TYPE="Text" MAXLENGTH="30" SIZE=30  ALT ="회사명" tag="22XXXU"></TD>
							    <TD CLASS=TD5 NOWRAP>사업자등록번호</TD>
							    <TD CLASS=TD6 NOWRAP><script language =javascript src='./js/h9105ma1_fpDoubleSingle30_txtA_comp_no.js'></script></TD>
							</TR>
							<TR>
							    <TD CLASS=TD5 NOWRAP>급여총액</TD>
							    <TD CLASS=TD6 NOWRAP><script language =javascript src='./js/h9105ma1_fpDoubleSingle_txtA_pay_tot_amt.js'></script></TD>
							    <TD CLASS=TD5 NOWRAP>상여총액</TD>
							    <TD CLASS=TD6 NOWRAP><script language =javascript src='./js/h9105ma1_fpDoubleSingle1_txtA_bonus_tot_amt.js'></script></TD>
						    </TR>
							<TR>
							    <TD CLASS=TD5 NOWRAP>인정상여</TD>
							    <TD CLASS=TD6 NOWRAP><script language =javascript src='./js/h9105ma1_fpDoubleSingle2_txtA_after_bonus_amt.js'></script></TD>
							    <TD CLASS=TD5 NOWRAP>건강,고용보험료</TD>
							    <TD CLASS=TD6 NOWRAP><script language =javascript src='./js/h9105ma1_fpDoubleSingle3_txtA_med_insur_amt.js'></script></TD>
						    </TR>
							<TR>
							    <TD CLASS=TD5 NOWRAP>국민연금</TD>
							    <TD CLASS=TD6 NOWRAP><script language =javascript src='./js/h9105ma1_fpDoubleSingle4_txtA_national_pension_amt.js'></script></TD>
							    <TD CLASS=TD5 NOWRAP>저축기금</TD>
							    <TD CLASS=TD6 NOWRAP><script language =javascript src='./js/h9105ma1_fpDoubleSingle5_txtA_save_tax_sub_amt.js'></script></TD>
						    </TR>
							<TR>
							    <TD CLASS=TD5 NOWRAP>개인연금(2001년이전)</TD>
							    <TD CLASS=TD6 NOWRAP><script language =javascript src='./js/h9105ma1_fpDoubleSingle6_txtA_indiv_anu_amt.js'></script></TD>
							    <TD CLASS=TD5 NOWRAP>개인연금(2001년이후)</TD>
							    <TD CLASS=TD6 NOWRAP><script language =javascript src='./js/h9105ma1_fpDoubleSingle50_txtA_indiv_anu2_amt.js'></script></TD>
						    </TR>
							<TR>
							    <TD CLASS=TD5 NOWRAP>결정소득세</TD>
							    <TD CLASS=TD6 NOWRAP><script language =javascript src='./js/h9105ma1_fpDoubleSingle7_txtA_income_tax_amt.js'></script></TD>
							    <TD CLASS=TD5 NOWRAP>결정주민세</TD>
							    <TD CLASS=TD6 NOWRAP><script language =javascript src='./js/h9105ma1_fpDoubleSingle8_txtA_res_tax_amt.js'></script></TD>
						    </TR>
							<TR>
							    <TD CLASS=TD5 NOWRAP>농특세</TD>
							    <TD CLASS=TD6 NOWRAP><script language =javascript src='./js/h9105ma1_fpDoubleSingle9_txtA_farm_tax_amt.js'></script></TD>
							    <TD CLASS=TD5 NOWRAP>연장비과세</TD>
							    <TD CLASS=TD6 NOWRAP><script language =javascript src='./js/h9105ma1_fpDoubleSingle10_txtA_non_tax1_amt.js'></script></TD>
						    </TR>
							<TR>
							    <TD CLASS=TD5 NOWRAP>식대비과세</TD>
							    <TD CLASS=TD6 NOWRAP><script language =javascript src='./js/h9105ma1_fpDoubleSingle11_txtA_non_tax2_amt.js'></script></TD>
							    <TD CLASS=TD5 NOWRAP>기타비과세1</TD>
							    <TD CLASS=TD6 NOWRAP><script language =javascript src='./js/h9105ma1_fpDoubleSingle12_txtA_non_tax3_amt.js'></script></TD>
						    </TR>
							<TR>
							    <TD CLASS=TD5 NOWRAP>기타비과세2</TD>
							    <TD CLASS=TD6 NOWRAP><script language =javascript src='./js/h9105ma1_fpDoubleSingle13_txtA_non_tax4_amt.js'></script></TD>
							    <TD CLASS=TD5 NOWRAP>국외근로비과세</TD>
							    <TD CLASS=TD6 NOWRAP><script language =javascript src='./js/h9105ma1_fpDoubleSingle14_txtA_non_tax5_amt.js'></script></TD>
						    </TR>
						    <% Call SubFillRemBodyTD5656(9) %>
					    </TABLE>
					    </FIELDSET>
					    </DIV>
					    <DIV ID="TabDiv" STYLE="FLOAT: left; HEIGHT: 100%; WIDTH: 100%;" SCROLL="NO">
						<FIELDSET CLASS="CLSFLD"><LEGEND ALIGN=LEFT>종전근무지2</LEGEND>
                        <TABLE <%=LR_SPACE_TYPE_60%>>
						    <TR>
							    <TD CLASS=TD5 NOWRAP>회사명</TD>
							    <TD CLASS=TD6 NOWRAP><INPUT NAME="txtA_comp_nm2" MAXLENGTH="30" SIZE=30   ALT ="회사명" tag="22XXXU"></TD>
							    <TD CLASS=TD5 NOWRAP>사업자등록번호</TD>
							    <TD CLASS=TD6 NOWRAP><script language =javascript src='./js/h9105ma1_fpDoubleSingle31_txtA_comp_no2.js'></script></TD>
							</TR>
							<TR>
							    <TD CLASS=TD5 NOWRAP>급여총액</TD>
							    <TD CLASS=TD6 NOWRAP><script language =javascript src='./js/h9105ma1_fpDoubleSingle_txtA_pay_tot_amt2.js'></script></TD>
							    <TD CLASS=TD5 NOWRAP>상여총액</TD>
							    <TD CLASS=TD6 NOWRAP><script language =javascript src='./js/h9105ma1_fpDoubleSingle1_txtA_bonus_tot_amt2.js'></script></TD>
						    </TR>
							<TR>
							    <TD CLASS=TD5 NOWRAP>인정상여</TD>
							    <TD CLASS=TD6 NOWRAP><script language =javascript src='./js/h9105ma1_fpDoubleSingle2_txtA_after_bonus_amt2.js'></script></TD>
							    <TD CLASS=TD5 NOWRAP>건강,고용보험료</TD>
							    <TD CLASS=TD6 NOWRAP><script language =javascript src='./js/h9105ma1_fpDoubleSingle3_txtA_med_insur_amt2.js'></script></TD>
						    </TR>
							<TR>
							    <TD CLASS=TD5 NOWRAP>국민연금</TD>
							    <TD CLASS=TD6 NOWRAP><script language =javascript src='./js/h9105ma1_fpDoubleSingle4_txtA_national_pension_amt2.js'></script></TD>
							    <TD CLASS=TD5 NOWRAP>저축기금</TD>
							    <TD CLASS=TD6 NOWRAP><script language =javascript src='./js/h9105ma1_fpDoubleSingle5_txtA_save_tax_sub_amt2.js'></script></TD>
						    </TR>
							<TR>
							    <TD CLASS=TD5 NOWRAP>개인연금(2001년이전)</TD>
							    <TD CLASS=TD6 NOWRAP><script language =javascript src='./js/h9105ma1_fpDoubleSingle6_txtA_indiv_anu_amt2.js'></script></TD>
							    <TD CLASS=TD5 NOWRAP>개인연금(2001년이후)</TD>
							    <TD CLASS=TD6 NOWRAP><script language =javascript src='./js/h9105ma1_fpDoubleSingle55_txtA_indiv_anu2_amt2.js'></script></TD>
						    </TR>
							<TR>
							    <TD CLASS=TD5 NOWRAP>결정소득세</TD>
							    <TD CLASS=TD6 NOWRAP><script language =javascript src='./js/h9105ma1_fpDoubleSingle7_txtA_income_tax_amt2.js'></script></TD>
							    <TD CLASS=TD5 NOWRAP>결정주민세</TD>
							    <TD CLASS=TD6 NOWRAP><script language =javascript src='./js/h9105ma1_fpDoubleSingle8_txtA_res_tax_amt2.js'></script></TD>
						    </TR>
							<TR>
							    <TD CLASS=TD5 NOWRAP>농특세</TD>
							    <TD CLASS=TD6 NOWRAP><script language =javascript src='./js/h9105ma1_fpDoubleSingle9_txtA_farm_tax_amt2.js'></script></TD>
							    <TD CLASS=TD5 NOWRAP>연장비과세</TD>
							    <TD CLASS=TD6 NOWRAP><script language =javascript src='./js/h9105ma1_fpDoubleSingle10_txtA_non_tax1_amt2.js'></script></TD>
						    </TR>
							<TR>
							    <TD CLASS=TD5 NOWRAP>식대비과세</TD>
							    <TD CLASS=TD6 NOWRAP><script language =javascript src='./js/h9105ma1_fpDoubleSingle11_txtA_non_tax2_amt2.js'></script></TD>
							    <TD CLASS=TD5 NOWRAP>기타비과세1</TD>
							    <TD CLASS=TD6 NOWRAP><script language =javascript src='./js/h9105ma1_fpDoubleSingle12_txtA_non_tax3_amt2.js'></script></TD>
						    </TR>
							<TR>
							    <TD CLASS=TD5 NOWRAP>기타비과세2</TD>
							    <TD CLASS=TD6 NOWRAP><script language =javascript src='./js/h9105ma1_fpDoubleSingle13_txtA_non_tax4_amt2.js'></script></TD>
							    <TD CLASS=TD5 NOWRAP>국외근로비과세</TD>
							    <TD CLASS=TD6 NOWRAP><script language =javascript src='./js/h9105ma1_fpDoubleSingle14_txtA_non_tax5_amt2.js'></script></TD>
						    </TR>
			    		    <% Call SubFillRemBodyTD5656(9) %>
					    </TABLE>
					    </FIELDSET>
					    </DIV>
                    </TD>
                </TR>
            </TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=0 <%=BizSize%> FRAMEBORDER=0 SCROLLING=no noresize framespacing=0></IFRAME>
		</TD>
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
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
