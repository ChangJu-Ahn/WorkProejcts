<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% session.CodePage=949 %>

<!--
======================================================================================================
*  1. Module Name          : Human Resources
*  2. Function Name        : Single Sample
*  3. Program ID           : h9112ma1
*  4. Program Name         : h9112ma1
*  5. Program Desc         : 연말정산관리 
*  6. Comproxy List        :
*  7. Modified date(First) : 2001/05/18
*  8. Modified date(Last)  : 2003/06/13
*  9. Modifier (First)     : TGS 최용철 
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
Const BIZ_PGM_ID       = "h9112mb1.asp"						           '☆: Biz Logic ASP Name
'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim IsOpenPop
Dim lgOldRow
Dim gSelframeFlg                                                       '현재 TAB의 위치를 나타내는 Flag %>
Dim gblnWinEvent                                                       'ShowModal Dialog(PopUp) Window가 여러 개 뜨는 것을 방지하기 위해 
Dim lgBlnFlawChgFlg	
Dim gtxtChargeType
<%
   Enddate=GetSvrDate
%>
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
	lgOldRow = 0
		
    gblnWinEvent      = False
	lgBlnFlawChgFlg   = False
End Sub
	
'========================================================================================================
' Name : LoadInfTB19029()	
' Desc : Set System Number format
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "H","NOCOOKIE","MA") %>
End Sub

'------------------------------------------  CookiePage()  --------------------------------------------------
'	Name : CookiePage()
'	Description : Jump시 Condition에서 넘겨오는 값 setting
'---------------------------------------------------------------------------------------------------------
Function CookiePage(ByVal flgs)
End Function

'========================================================================================================
' Function Name : MakeKeyStream
' Function Desc : This method set focus to pos of err
'========================================================================================================
Sub MakeKeyStream(pOpt)
    lgKeyStream  = Trim(Frm1.txtpay_yymm_dt.Year)           & parent.gColSep
    lgKeyStream  = lgKeyStream & Trim(Frm1.txtemp_no.value) & parent.gColSep
    
    lgKeyStream  = lgKeyStream & lgUsrIntcd & parent.gColSep      
End Sub        

'========================================================================================================
' Name : Form_Load
' Desc : developer describe this line Called by Window_OnLoad() evnt
'========================================================================================================
Sub Form_Load()
	Dim strYear
    Dim strMonth
    Dim strDay
   
    Err.Clear                                                                       '☜: Clear err status
	Call LoadInfTB19029                                                             '☜: Load table , B_numeric_format

    Call ggoOper.FormatField(Document, "A",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
	Call ggoOper.LockField(Document, "N")											'⊙: Lock Field

    Call ggoOper.FormatNumber(frm1.txtsupp_old_cnt, "999", "0", False, 0)
    Call ggoOper.FormatNumber(frm1.txtsupp_young_cnt, "999", "0", False, 0)
    Call ggoOper.FormatNumber(frm1.txtparia_cnt, "999", "0", False, 0)
    Call ggoOper.FormatNumber(frm1.txtold_cnt1, "999", "0", False, 0)
    Call ggoOper.FormatNumber(frm1.txtold_cnt2, "999", "0", False, 0)

    Call InitVariables                                                              'Initializes local global variables
    '아래 연도 세팅해주는 부분을 setDefaultVal 함수에서 이리 옮겨옴. 안그러면 다른년도 조회가 안됨	
    Call ExtractDateFrom("<%=EndDate%>",parent.gServerDateFormat,parent.gServerDateType,strYear,strMonth,strDay)

    frm1.txtpay_yymm_dt.focus 
	frm1.txtpay_yymm_dt.Year = strYear
	
    Call ggoOper.FormatDate(frm1.txtpay_yymm_dt, parent.gDateFormat, 3)                    '싱글에서 년월말 입력하고 싶은경우 다음 함수를 콜한다.
    Call FuncGetAuth("H9112MA1", parent.gUsrID, lgUsrIntCd)     ' 자료권한:lgUsrIntCd ("%", "1%")
	Call SetToolbar("1100000000001111")												'⊙: Set ToolBar
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
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO,"x","x")					 '☜: Data is changed.  Do you want to display it? 
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    Call ggoOper.ClearField(Document, "2")										 '☜: Clear Contents  Field
    
    If  txtEmp_no_Onchange()  then
        Exit Function
    End If
    
    If Not chkField(Document, "1") Then									         '☜: This function check required field
       Exit Function
    End If
    
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
       IntRetCD = DisplayMsgBox("900015", parent.VB_YES_NO,"x","x")					 '☜: Data is changed.  Do you want to make it new? 
       If IntRetCD = vbNo Then
          Exit Function
       End If
    End If
    
    Call ggoOper.ClearField(Document, "1")                                       '☜: Clear Condition Field
    Call ggoOper.ClearField(Document, "2")                                       '☜: Clear Contents  Field
    Call ggoOper.LockField(Document , "N")                                       '☜: Lock  Field
    
    Call SetToolbar("1100000000001111")
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
        Call DisplayMsgBox("900002","x","x","x")                                
        Exit Function
    End If
    
    IntRetCD = DisplayMsgBox("900003", parent.VB_YES_NO,"x","x")                        '☜: Do you want to delete? 
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

    FncSave = False                                                              '☜: Processing is NG
    
    Err.Clear                                                                    '☜: Clear err status
    
    If lgBlnFlgChgValue = False Then 
        IntRetCD = DisplayMsgBox("900001","x","x","x")                           '☜:There is no changed data. 
        Exit Function
    End If
    
    If Not chkField(Document, "2") Then                                          '☜: Check contents area
       Exit Function
    End If

    Call MakeKeyStream("S")
    If DbSave = False Then
		Exit Function
	End If
    
    FncSave = True                                                              '☜: Processing is OK
End Function
'========================================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================================
Function FncCopy()
	Dim IntRetCD

    FncCopy = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
	
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900017", parent.VB_YES_NO,"x","x")				     '☜: Data is changed.  Do you want to continue? 
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    lgIntFlgMode = parent.OPMD_CMODE												     '⊙: Indicates that current mode is Crate mode
    
    Call ggoOper.ClearField(Document, "1")                                       '⊙: Clear Condition Field
    Call ggoOper.LockField(Document, "N")									     '⊙: This function lock the suitable field
    Call SetToolbar("1100000000001111")

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
		IntRetCD = DisplayMsgBox("900017", parent.VB_YES_NO,"x","x")					 '☜: Will you destory previous data
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If
	
    Call MakeKeyStream("P")
    Call ggoOper.ClearField(Document, "2")										 '⊙: Clear Contents Area
    
    Call InitVariables														 '⊙: Initializes local global variables

    Call LayerShowHide(1)


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
        Call DisplayMsgBox("900002","x","x","x")
        Exit Function
    End If
	
	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900017", parent.VB_YES_NO,"x","x")					 '☜: Will you destory previous data
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If
	
    Call MakeKeyStream("N")

    Call ggoOper.ClearField(Document, "2")										 '⊙: Clear Contents Area
    
    Call InitVariables														     '⊙: Initializes local global variables

    Call LayerShowHide(1)


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
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO,"x","x")			'⊙: Data is changed.  Do you want to exit? 
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

    Call LayerShowHide(1)

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
	Call LayerShowHide(1)
		
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
		
	Call LayerShowHide(1)
		
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

	Call SetToolbar("1100000000011111")

    Call ggoOper.LockField(Document, "Q")
    Set gActiveElement = document.ActiveElement   
	frm1.txtpay_yymm_dt.focus
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
	Call FncNew()	
End Function

'----------------------------------------  OpenEmptName()  ------------------------------------------
'	Name : OpenEmptName()                                                         <==== 성명/사번 팝업 
'	Description : Employee PopUp
'------------------------------------------------------------------------------------------------
Function OpenEmptName(iWhere)
	Dim arrRet
	Dim arrParam(2)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	If iWhere = 0 Then 'TextBox(Condition)
		arrParam(0) = frm1.txtEmp_no.value			' Code Condition
	Else 'spread
		arrParam(0) = frm1.vspdData.Text			' Code Condition
	End If
	    arrParam(1) = ""'frm1.txtName.value			' Name Cindition
	
	arrParam(2) = lgUsrIntcd
	
	arrRet = window.showModalDialog(HRAskPRAspName("EmpPopup"), Array(window.parent,arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		If iWhere = 0 Then
			frm1.txtEmp_no.focus
		Else
			frm1.vspdData.Col = C_EmpNo
			frm1.vspdData.action =0
		End If	
		Exit Function
	Else
		Call SetEmp(arrRet, iWhere)
	End If	
			
End Function

'------------------------------------------  SetEmp()  ------------------------------------------------
'	Name : SetEmp()
'	Description : Employee Popup에서 Return되는 값 setting
'------------------------------------------------------------------------------------------------------
Function SetEmp(Byval arrRet, Byval iWhere)
		
	With frm1
		If iWhere = 0 Then
			.txtEmp_no.value = arrRet(0)
			.txtName.value = arrRet(1)
			.txtEmp_no.focus
		Else
			.vspdData.Col = C_EmpNm
			.vspdData.Text = arrRet(1)
			.vspdData.Col = C_EmpNo
			.vspdData.Text = arrRet(0)
			.vspdData.action =0
		End If
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

    If  frm1.txtEmp_no.value = "" Then
		frm1.txtName.value = ""
    Else
	    IntRetCd = FuncGetEmpInf2(frm1.txtEmp_no.value,lgUsrIntCd,strName,strDept_nm,_
	                strRoll_pstn, strPay_grd1, strPay_grd2, strEntr_dt, strInternal_cd)
                
	    If  IntRetCd < 0 then
	        If  IntRetCd = -1 then
    			Call DisplayMsgBox("800048","X","X","X")	'해당사원은 존재하지 않습니다.
            Else
                Call DisplayMsgBox("800454","X","X","X")	'자료에 대한 권한이 없습니다.
            End if
            
            Call ggoOper.ClearField(Document, "2")	
			frm1.txtName.value = ""
            Frm1.txtEmp_no.focus 
            Set gActiveElement = document.ActiveElement
            txtEmp_no_Onchange = true
        Else
			frm1.txtName.value = strName
        End if 
    End if  
End Function

'=======================================================================================================
'   Event Name : txtpay_yymm_dt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtpay_yymm_dt_DblClick(Button)
    If Button = 1 Then
		Call SetFocusToDocument("M")    
        frm1.txtpay_yymm_dt.Action = 7
        frm1.txtpay_yymm_dt.focus
    End If
End Sub
'==========================================================================================
'   Event Name : txtpay_yymm_dt_KeyDown()
'   Event Desc : 조회조건부의 txtpay_yymm_dt_KeyDown시 EnterKey일 경우는 Query
'==========================================================================================
Sub txtpay_yymm_dt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call FncQuery()
End Sub


</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
</HEAD>

<BODY TABINDEX="-1" SCROLL="YES">
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
								<TD background="../../../CShared/image/table/seltab_up_bg.gif"><IMG src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23" ></TD>
								<TD background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>연말정산내역조회</font></TD>
								<TD background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><IMG src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23" ></TD>
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
	   	<TD WIDTH="100%" CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
			    <TR>
			        <TD <%=HEIGHT_TYPE_02%>></TD>
			    </TR>
				<TR>
					<TD HEIGHT="20" WIDTH="100%">
					    <FIELDSET CLASS="CLSFLD">
					        <TABLE <%=LR_SPACE_TYPE_40%>>
					 	        <TR>
							    	<TD CLASS=TD5 NOWRAP>급여년</TD>
							    	<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=txtpay_yymm_dt NAME="txtpay_yymm_dt" style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 50px; CENTER: 0px" title=FPDATETIME ALT="급여년월" tag="12X1" VIEWASTEXT> </OBJECT>');</SCRIPT></TD>		
							        <TD CLASS=TD5 NOWRAP>사원</TD>
							    	<TD CLASS=TD6 NOWRAP><INPUT NAME="txtEmp_no" MAXLENGTH="13" SIZE="13" ALT ="사번" tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" ALIGN=top TYPE="BUTTON" ONCLICK="VBScript: OpenEmptName(0)">
							    	                     <INPUT NAME="txtName" MAXLENGTH="30" SIZE="20" ALT ="성명" tag="14XXXU"></TD>
							    </TR>
					        </TABLE>
				        </FIELDSET>
				    </TD>
				</TR>
				<TR>                    <!-- Condition Area-->
				    <TD <%=HEIGHT_TYPE_03%>WIDTH="100%"></TD>
				</TR>
			    <TR>	                 <!-- space Area-->
				    <TD WIDTH="100%" HEIGHT=* valign=top>
                        <TABLE <%=LR_SPACE_TYPE_60%> bgcolor=#EEEEEC>
                            <TR>
                                <TD WIDTH="100%" VALIGN=TOP>
                                   <FIELDSET CLASS="Clsdiv"><LEGEND ALGIN="LEFT">소득사항</LEGEND>
                                   <table <%=LR_SPACE_TYPE_20%> border="1" width="100%">
                                      <TR>
                                      	  <TD BGCOLOR=#d1e8f9  width="48%" align="middle">구분</TD>
                                      	  <TD BGCOLOR=#d1e8f9  width="13%" align="middle">급여</TD>
                                          <TD BGCOLOR=#d1e8f9  width="13%" align="middle">상여</TD>
                                          <TD BGCOLOR=#d1e8f9  width="13%" align="middle">인정상여</TD>
                                          <TD BGCOLOR=#d1e8f9  width="13%" align="middle">합계</TD>
                                      </TR>
                                      <TR>
                                          <TD BGCOLOR=#d1e8f9 width="28%">1.현근무지근로소득수입금액</TD>
                                          <TD width="18%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=hfa050t_new_pay_tot_amt name=txtNew_pay_tot_amt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 180px" title=FPDOUBLESINGLE tag="24X2Z" ALT="급여"></OBJECT>');</SCRIPT></TD>
                                          <TD width="18%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=hfa050t_new_bonus_tot_amt name=txtNew_bonus_tot_amt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 180px" title=FPDOUBLESINGLE tag="24X2Z" ALT="상여"></OBJECT>');</SCRIPT></TD>
                                          <TD width="18%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=after_bonus_amt name=txtafter_bonus_amt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 180px" title=FPDOUBLESINGLE tag="24X2Z" ALT="상여"></OBJECT>');</SCRIPT></TD>
                                          <TD width="18%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=a_amt name=txta_amt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 180px" title=FPDOUBLESINGLE tag="24X2Z" ALT="합계"></OBJECT>');</SCRIPT></TD>
                                      </TR>
                                      <TR>
                                           <TD BGCOLOR=#d1e8f9 width="28%">2.전근무지근로소득수입금액</TD>
                                           <TD width="18%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=pay_tot_amt name=txtpay_tot_amt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 180px" title=FPDOUBLESINGLE tag="24X2Z" ALT="급여"></OBJECT>');</SCRIPT></TD>
                                           <TD width="18%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=bonus_tot_amt name=txtbonus_tot_amt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 180px" title=FPDOUBLESINGLE tag="24X2Z" ALT="상여"></OBJECT>');</SCRIPT></TD>
                                           <TD width="18%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=old_after_bonus_amt name=txtold_after_bonus_amt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 180px" title=FPDOUBLESINGLE tag="24X2Z" ALT="상여"></OBJECT>');</SCRIPT></TD>
                                           <TD width="18%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=b_amt name=txtb_amt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 180px" title=FPDOUBLESINGLE tag="24X2Z" ALT="합계"></OBJECT>');</SCRIPT></TD>
                                      </TR>
                                      <TR>
                                           <TD BGCOLOR=#d1e8f9 width="28%" COLSPAN="4">3.근로소득수입금액</TD>
                                           <TD><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=hfa050t_income_tot_amt name=txtincome_tot_amt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 180px" title=FPDOUBLESINGLE tag="24X2Z" ALT="합계"></OBJECT>');</SCRIPT></TD>
                                      </TR>
                                      <TR>
                                           <TD BGCOLOR=#d1e8f9 width="28%">4.근로소득공제</TD>
                                           <TD width="18%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=hfa050t_income_sub_amt name=txtincome_sub_amt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 180px" title=FPDOUBLESINGLE tag="24X2Z" ALT="급여"></OBJECT>');</SCRIPT></TD>
                                           <TD BGCOLOR=#d1e8f9 width="28%" COLSPAN="2">5.근로소득금액</TD>
                                           <TD width="18%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=hfa050t_income_amt name=txthfa050t_income_amt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 180px" title=FPDOUBLESINGLE tag="24X2Z" ALT="합계"></OBJECT>');</SCRIPT></TD>
                                      </TR>
                                   </TABLE>
                                   </FIELDSET>
                                   
                                   <BR>
                                   <FIELDSET CLASS="Clsdiv"><LEGEND ALGIN="LEFT">인적공제</LEGEND>
                                   <TABLE <%=LR_SPACE_TYPE_20%> border="1" width="100%">
                                      <TR>
                                           <TD BGCOLOR=#d1e8f9 width="30%" colspan="2" align="middle">공제항목</TD>
                                           <TD BGCOLOR=#d1e8f9 width="10%" align="middle">공제사항</TD>
                                           <TD BGCOLOR=#d1e8f9 width="10%" align="middle">정산결과</TD>
                                           <TD BGCOLOR=#d1e8f9 width="30%" colspan="2" align="middle">공제항목</TD>
                                           <TD BGCOLOR=#d1e8f9 width="10%" align="middle">공제사항</TD>
                                           <TD BGCOLOR=#d1e8f9 width="10%" align="middle">정산결과</TD>
                                      </TR>
                                      <TR>
                                           <TD BGCOLOR=#d1e8f9 width="6%" rowspan="4">6.기본공제</TD>    
                                           <TD BGCOLOR=#d1e8f9 width="24%">본인공제</TD>       
                                           <TD width="10%" colspan="2"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=hfa050t_per_sub_amt name=txtper_sub_amt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 254px" title=FPDOUBLESINGLE tag="24X2Z" ALT="본인공제정산결과"></OBJECT>');</SCRIPT></TD>
                                           <TD BGCOLOR=#d1e8f9 width="6%" rowspan="5">7.추가공제</TD>       
                                           <TD BGCOLOR=#d1e8f9 width="24%">장애인수</TD>       
                                           <TD width="10%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=hfa050t_paria_cnt name=txtparia_cnt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 125px" title=FPDOUBLESINGLE tag="24X3Z" ALT="장애인수공제사항"></OBJECT>');</SCRIPT></TD>
                                           <TD width="10%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=hfa050t_paria_sub_amt name=txtparia_sub_amt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 125px" title=FPDOUBLESINGLE tag="24X2Z" ALT="장애인수정산결과"></OBJECT>');</SCRIPT></TD>
                                      </TR>       
                                      <TR>
                                           <TD BGCOLOR=#d1e8f9 width="24%">배우자(Y/N)</TD>       
                                           <TD width="10%"><INPUT Name=txtspouse MAXLENGTH="10" SIZE=19 id=hfa050t_spouse Tag="24XXXU" ALT="배우자공제사항"></INPUT></TD>
                                           <TD width="10%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=hfa050t_spouse_sub_amt name=txtspouse_sub_amt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 125px" title=FPDOUBLESINGLE tag="24X2Z" ALT="배우자정산결과"></OBJECT>');</SCRIPT></TD>
                                           <TD BGCOLOR=#d1e8f9 width="24%">경로우대수(65세이상)</TD>      
                                           <TD width="10%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=hfa050t_old_cnt1 name=txtold_cnt1 style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 125px" title=FPDOUBLESINGLE tag="24X3Z" ALT="경로우대공제사항1"></OBJECT>');</SCRIPT></TD>
                                           <TD width="10%" rowspan=2 ><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=hfa050t_old_sub_amt1 name=txtold_sub_amt1 style="HEIGHT: 44px; LEFT: 0px; TOP: 0px; WIDTH: 125px" title=FPDOUBLESINGLE tag="24X2Z" ALT="경로우대정산결과1"></OBJECT>');</SCRIPT></TD>
                                      </TR> 
                                      <TR>   
                                           <TD BGCOLOR=#d1e8f9 width="24%">부양자(여55,남60세이상)</TD>     
                                           <TD width="10%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=hdf020t_supp_old_cnt name=txtsupp_old_cnt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 125px" title=FPDOUBLESINGLE tag="24X3Z" ALT="부양자공제사항"></OBJECT>');</SCRIPT></TD>
                                           <TD width="10%" rowspan="2"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=hfa050t_supp_sub_amt name=txtsupp_sub_amt style="HEIGHT: 44px; LEFT: 0px; TOP: 0px; WIDTH: 125px" title=FPDOUBLESINGLE tag="24X2Z" ALT="부양자정산결과"></OBJECT>');</SCRIPT></TD>
                                           <TD BGCOLOR=#d1e8f9 width="24%">경로우대수(70세이상)</TD>      
                                           <TD width="10%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=hfa050t_old_cnt2 name=txtold_cnt2 style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 125px" title=FPDOUBLESINGLE tag="24X3Z" ALT="경로우대공제사항2"></OBJECT>');</SCRIPT></TD>
                                      </TR>   
                                      <TR>   
                                           <TD BGCOLOR=#d1e8f9 width="24%">부양자(20세이하/초과장애인)</TD>      
                                           <TD width="10%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=hdf020t_supp_young_cnt name=txtsupp_young_cnt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 125px" title=FPDOUBLESINGLE tag="24X3Z" ALT="부양자정산결과"></OBJECT>');</SCRIPT></TD>
                                           <TD BGCOLOR=#d1e8f9 width="24%">부녀자세대주여부(Y/N)</TD>      
                                           <TD width="10%"><INPUT Name=txtlady MAXLENGTH="10" SIZE=19 id=hfa050t_lady tag="24XXXU" ALT="부녀자세대주공제사항"></INPUT></TD>
                                           <TD width="10%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=hfa050t_lady_sub_amt name=txtlady_sub_amt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 125px" title=FPDOUBLESINGLE tag="24X2Z" ALT="부녀자세대주정산결과"></OBJECT>');</SCRIPT></TD>
                                      </TR>   
                                      <TR>   
                                           <TD BGCOLOR=#d1e8f9 width="24%" colspan="2">다자녀추가공제</TD>
			<TD width="10%" colspan="2"> <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=hfa050t_small_sub_amt_txtsmall_sub_amt9 name=hfa050t_small_sub_amt_txtsmall_sub_amt9 style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 254px" title=FPDOUBLESINGLE tag="24X2Z" ALT="소수공제자추가공제"></OBJECT>');</SCRIPT></TD>

                                                <TD BGCOLOR=#d1e8f9 width="24%">자녀양육수(6세이하)</TD>
                                           <TD width="10%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=hfa050t_chl_rear name=txtchl_rear style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 125px" title=FPDOUBLESINGLE tag="24X2Z" ALT="자녀양육수공제사항"></OBJECT>');</SCRIPT></TD>
                                           <TD width="10%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=hfa050t_chl_rear_sub_amt name=txtchl_rear_sub_amt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 125px" title=FPDOUBLESINGLE tag="24X2Z" ALT="자녀양육수정산결과"></OBJECT>');</SCRIPT></TD>
                                      </TR>
 
			 <TR>   
          			  	<TD BGCOLOR=#d1e8f9 width="24%" colspan="2">8.소수공제자추가공제</TD>
                                           	<TD width="10%" colspan="2"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=hfa050t_small_sub_amt name=txtsmall_sub_amt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 254px" title=FPDOUBLESINGLE tag="24X2Z" ALT="소수공제자추가공제"></OBJECT>');</SCRIPT></TD>
                                         		  <TD BGCOLOR=#d1e8f9 width="24%" colspan=2>9.인적공제계</TD>
                                           	<TD width="10%" colspan=2><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=d_amt name=txtd_amt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 254px" title=FPDOUBLESINGLE tag="24X2Z" ALT="인적공제계"></OBJECT>');</SCRIPT></TD>
                                      </TR>                                       
                                   </TABLE>
                                   </FIELDSET>
            				       <BR>
                                   <FIELDSET CLASS="Clsdiv"><LEGEND ALGIN="LEFT">연금보험료공제</LEGEND>
                                   <TABLE <%=LR_SPACE_TYPE_20%> border="1" width="100%">
                                      <TR>
                                           <TD BGCOLOR=#d1e8f9 width="50%" align="middle">공제항목</TD>
                                           <TD BGCOLOR=#d1e8f9 width="25%" align="middle">공제사항</TD>
                                           <TD BGCOLOR=#d1e8f9 width="25%" align="middle">정산결과</TD>
                                      </TR>
                                      <TR> 
                                           <TD BGCOLOR=#d1e8f9 width="50%">10.연금보험료공제</TD>
                                           <TD width="25%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=hfa030t_National_pension_amt name=txthfa030t_National_pension_amt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 305px" title=FPDOUBLESINGLE tag="24X2Z" ALT="연금보험료공제사항"></OBJECT>');</SCRIPT></TD>
                                           <TD width="25%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=hfa050t_National_pension_sub_amt name=txthfa050t_National_pension_sub_amt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 305px" title=FPDOUBLESINGLE tag="24X2Z" ALT="연금보험료정산결과"></OBJECT>');</SCRIPT></TD>
                                      </TR>
                                   </TABLE>
                                   </FIELDSET>
            				       <BR>
            				       <FIELDSET CLASS="Clsdiv"><LEGEND ALGIN="LEFT">특별세액공제</LEGEND>
                                   <TABLE <%=LR_SPACE_TYPE_20%> border="1" ALGIN="TOP" width="100%">
                                      <TR>
                                           <TD BGCOLOR=#d1e8f9 width="30%" colspan="3" align="middle">공제항목</TD>
                                           <TD BGCOLOR=#d1e8f9 width="10%" align="middle">공제사항</TD>
                                           <TD BGCOLOR=#d1e8f9 width="10%" align="middle">정산결과</TD>
                                           <TD BGCOLOR=#d1e8f9 width="30%" colspan="2" align="middle">공제항목</TD>
                                           <TD BGCOLOR=#d1e8f9 width="10%" align="middle">공제사항</TD>
                                           <TD BGCOLOR=#d1e8f9 width="10%" align="middle">정산결과</TD>
                                      </TR>
                                      <TR>
                                           <TD BGCOLOR=#d1e8f9 width="2%" rowspan="10">특별공제</TD>
                                           <TD BGCOLOR=#d1e8f9 width="4%" rowspan="4">11.보험료</TD>
                                           <TD BGCOLOR=#d1e8f9 width="24%">건강보험료</TD>
                                           <TD width="10%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=insur_amt name=txtinsur_amt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 125px" title=FPDOUBLESINGLE tag="24X2Z" ALT="건강보험료공제사항"></OBJECT>');</SCRIPT></TD>
                                           <TD width="10%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=hfa050t_med_insur_amt name=txtmed_insur_amt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 125px" title=FPDOUBLESINGLE tag="24X2Z" ALT="건강보험료정산결과"></OBJECT>');</SCRIPT></TD>
                                           <TD BGCOLOR=#d1e8f9 width="6%" rowspan="3">14.주택자금</TD>
                                           <TD BGCOLOR=#d1e8f9 width="24%">주택저축/차입금상환액</TD>
                                           <TD width="10%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=hfa030t_house_fund_amt name=txthfa030t_house_fund_amt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 125px" title=FPDOUBLESINGLE tag="24X2Z" ALT="주택저축/차입금상환액공제사항"></OBJECT>');</SCRIPT></TD>
                                           <TD width="10%" rowspan="3"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=hfa050t_house_fund_amt name=txthfa050t_house_fund_amt style="HEIGHT: 68px; LEFT: 0px; TOP: 0px; WIDTH: 125px" title=FPDOUBLESINGLE tag="24X2Z" ALT="주택저축/차입금상환액정산결과"></OBJECT>');</SCRIPT></TD>
                                      </TR>
                                      <TR>
                                           <TD BGCOLOR=#d1e8f9 width="24%">고용보험료</TD>
                                           <TD width="10%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=hfa030t_emp_insur_amt name=txthfa030t_emp_insur_amt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 125px" title=FPDOUBLESINGLE tag="24X2Z" ALT="고용보험료공제사항"></OBJECT>');</SCRIPT></TD>
                                           <TD width="10%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=hfa050t_emp_insur_amt name=txthfa050t_emp_insur_amt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 125px" title=FPDOUBLESINGLE tag="24X2Z" ALT="고용보험료정산결과"></OBJECT>');</SCRIPT></TD>
                                           <TD BGCOLOR=#d1e8f9 width="24%">장기주택저당&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;(15년미만)</font></TD>  
                                           <TD width="10%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=hfa030t_long_house_loan_amt name=txtlong_house_loan_amt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 125px" title=FPDOUBLESINGLE tag="24X2Z" ALT="장기주택저당차입금이자상환액(15년미만)"></OBJECT>');</SCRIPT></TD>
                                      </TR>
                                      <TR>
                                           <TD BGCOLOR=#d1e8f9 width="24%">기타보장성보험료</TD>
                                           <TD width="10%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=hfa030t_other_insur_amt name=txthfa030t_other_insur_amt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 125px" title=FPDOUBLESINGLE tag="24X2Z" ALT="기타보장성보험료공제사항"></OBJECT>');</SCRIPT></TD>
                                           <TD width="10%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=hfa050t_other_insur_amt name=txthfa050t_other_insur_amt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 125px" title=FPDOUBLESINGLE tag="24X2Z" ALT="기타보장성보험료정산결과"></OBJECT>');</SCRIPT></TD>
                                           <TD BGCOLOR=#d1e8f9 width="24%">차입금이자상환액&nbsp;(15년이상)</font></TD>  
                                           <TD width="10%" ><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=hfa030t_long_house_loan_amt1 name=txtlong_house_loan_amt1 style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 125px" title=FPDOUBLESINGLE tag="24X2Z" ALT="장기주택저당차입금이자상환액(15년이상)"></OBJECT>');</SCRIPT></TD>
                                      </TR>
                                      <TR>
                                           <TD BGCOLOR=#d1e8f9 width="24%">장애인전용보험료</TD>
                                           <TD width="10%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=hfa030t_disabled_insur_amt name=txthfa030t_disabled_insur_amt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 125px" title=FPDOUBLESINGLE tag="24X2Z" ALT="장애인전용보험료공제사항"></OBJECT>');</SCRIPT></TD>
                                           <TD width="10%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=hfa050t_disabled_insur_amt name=txthfa050t_disabled_insur_amt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 125px" title=FPDOUBLESINGLE tag="24X2Z" ALT="장애인전용보험료정산결과"></OBJECT>');</SCRIPT></TD>
                                           <TD BGCOLOR=#d1e8f9 width="6%" rowspan="7">15.기부금</TD>
                                           <TD BGCOLOR=#d1e8f9 width="24%">법정기부금</TD>
                                           <TD width="10%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=hfa030t_legal_contr_amt name=txtlegal_contr_amt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 125px" title=FPDOUBLESINGLE tag="24X2Z" ALT="법정기부금공제사항"></OBJECT>');</SCRIPT></TD>
                                           <TD width="10%" rowspan="7"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=hfa050t_contr_sub_amt name=txtcontr_sub_amt style="HEIGHT: 100%; LEFT: 0px; TOP: 0px; WIDTH: 125px" title=FPDOUBLESINGLE tag="24X2Z" ALT="법정기부금정산결과"></OBJECT>');</SCRIPT></TD>
                                      </TR>
                                      <TR>
                                           <TD BGCOLOR=#d1e8f9 width="4%" rowspan="2">12.의료비</TD>
                                           <TD BGCOLOR=#d1e8f9 width="24%">일반의료비</TD>
                                           <TD width="10%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=hfa030t_tot_med_amt name=txttot_med_amt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 125px" title=FPDOUBLESINGLE tag="24X2Z" ALT="일반의료비공제사항"></OBJECT>');</SCRIPT></TD>
                                           <TD width="10%" rowspan="2"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=hfa050t_med_sub_amt name=txtmed_sub_amt style="HEIGHT: 44px; LEFT: 0px; TOP: 0px; WIDTH: 125px" title=FPDOUBLESINGLE tag="24X2Z" ALT="일반의료비정산결과"></OBJECT>');</SCRIPT></TD>
                                           <TD BGCOLOR=#d1e8f9 width="24%">정치자금기부금</TD>
                                           <TD width="10%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="Object1" name=txtPoli_contr_amt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 125px" title=FPDOUBLESINGLE tag="24X2Z" ALT="조특법 제73조 기부금"></OBJECT>');</SCRIPT></TD>
                                      </TR>
                                      <TR>
                                           <TD BGCOLOR=#d1e8f9 width="24%">본인/경로자/장애인의료비</TD>
                                           <TD width="10%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=hfa030t_speci_med_amt name=txtspeci_med_amt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 125px" title=FPDOUBLESINGLE tag="24X2Z" ALT="경로,장애의료비"></OBJECT>');</SCRIPT></TD>
                                           <TD BGCOLOR=#d1e8f9 width="24%">특례기부금(100%)</TD>
                                           <TD width="10%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=hfa030t_taxLaw_contr_amt2 name=txtTaxLaw_contr_amt2 style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 125px" title=FPDOUBLESINGLE tag="24X2Z" ALT="특례기부금"></OBJECT>');</SCRIPT></TD>
                                      </TR>
                                      <TR>
                                           <TD BGCOLOR=#d1e8f9 width="4%" rowspan="3">13.교육비</TD>
                                           <TD BGCOLOR=#d1e8f9 width="24%">본인교육비</TD>
                                           <TD width="10%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=hfa030t_per_edu_amt name=txtper_edu_amt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 125px" title=FPDOUBLESINGLE tag="24X2Z" ALT="본인교육비"></OBJECT>');</SCRIPT></TD>
                                           <TD width="10%" rowspan="3"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=hfa050t_edu_sub_amt name=txtedu_sub_amt style="HEIGHT: 100%; LEFT: 0px; TOP: 0px; WIDTH: 125px" title=FPDOUBLESINGLE tag="24X2Z" ALT="정산결과"></OBJECT>');</SCRIPT></TD>
                                           <TD BGCOLOR=#d1e8f9 width="24%">특례기부금(50%)</TD>
                                           <TD width="10%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=hfa030t_taxLaw_contr_amt name=txtTaxLaw_contr_amt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 125px" title=FPDOUBLESINGLE tag="24X2Z" ALT="특례기부금"></OBJECT>');</SCRIPT></TD>
                                      </TR>
                                      <TR>
									       <TD BGCOLOR=#d1e8f9 width="24%">가족교육비</TD>
									       <TD width="10%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=edu_sum_amt name=txtedu_sum_amt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 125px" title=FPDOUBLESINGLE tag="24X2Z" ALT="가족교육비공제사항"></OBJECT>');</SCRIPT></TD>
									       <TD BGCOLOR=#d1e8f9 width="24%">우리사주조합기부금</TD>
                                           <TD width="10%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=hfa030t_ourstock_contr_amt name=txtOurstock_contr_amt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 125px" title=FPDOUBLESINGLE tag="24X2Z" ALT="노동조합비"></OBJECT>');</SCRIPT></TD>
                                      </TR>
                                      <tr>
                                           <TD BGCOLOR=#d1e8f9 width="24%">장애인특수교육비</TD>
                                           <TD width="10%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=Disabled_edu_amt name=txtDisable_edu_amt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 125px" title=FPDOUBLESINGLE tag="24X2Z" ALT="장애인특수교육비"></OBJECT>');</SCRIPT></td> 
                                           <TD BGCOLOR=#d1e8f9 width="24%">지정기부금</TD>
                                           <TD width="10%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=hfa030t_app_contr_amt name=txtapp_contr_amt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 125px" title=FPDOUBLESINGLE tag="24X2Z" ALT="지정기부금"></OBJECT>');</SCRIPT></TD>
                                      </tr>
                                     <TR>
                                           <TD BGCOLOR=#d1e8f9 width="29%" colspan="2">결혼/장례/이사비</TD>
                                           <TD width="10%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=hfa030t_Ceremony_amt name=hfa030t_Ceremony_amt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 125px" title=FPDOUBLESINGLE tag="24X2Z" ALT="결혼/장례/이사비"></OBJECT>');</SCRIPT></TD>
                                           <TD width="10%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=hfa050t_Ceremony_amt name=hfa050t_Ceremony_amt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 125px" title=FPDOUBLESINGLE tag="24X2Z" ALT="결혼/장례/이사비정산결과"></OBJECT>');</SCRIPT></TD>
                                           <TD BGCOLOR=#d1e8f9 width="24%">노동조합비</TD>
                                           <TD width="10%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=hfa030t_priv_contr_amt name=txtpriv_contr_amt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 125px" title=FPDOUBLESINGLE tag="24X2Z" ALT="지정기부금"></OBJECT>');</SCRIPT></TD>
                                      </TR>
                                      <TR>
                                           <TD BGCOLOR=#d1e8f9 width="26%" colspan="7" >16.계 또는 표준공제</TD>
                                           <TD width="24%" colspan="2" ><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=hfa050t_std_sub_amt name=txtstd_sub_amt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 254px" title=FPDOUBLESINGLE tag="24X2Z" ALT="계 또는 표준공제"></OBJECT>');</SCRIPT></TD>
                                      </TR>
                                   </TABLE>
            				       </FIELDSET>
                                   <BR>
                                   <FIELDSET CLASS="Clsdiv"><LEGEND ALGIN="LEFT">기타공제</LEGEND>
                                   <table <%=LR_SPACE_TYPE_20%> border="1" width="100%">
                                      <TR>
                                           <TD BGCOLOR=#d1e8f9 width="30%" align="middle">공제항목</TD>
                                           <TD BGCOLOR=#d1e8f9 width="10%" align="middle">공제사항</TD>
                                           <TD BGCOLOR=#d1e8f9 width="10%" align="middle">정산결과</TD>
                                           <TD BGCOLOR=#d1e8f9 width="30%" align="middle">공제항목</TD>
                                           <TD BGCOLOR=#d1e8f9 width="10%" align="middle">공제사항</TD>
                                           <TD BGCOLOR=#d1e8f9 width="10%" align="middle">정산결과</TD>
                                      </TR>
                                      <TR>
                                           <TD BGCOLOR=#d1e8f9 width="30%">17.개인연금저축액</TD>
                                           <TD width="10%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=hfa030t_indiv_anu_amt name=txthfa030t_indiv_anu_amt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 125px" title=FPDOUBLESINGLE tag="24X2Z" ALT="개인연금저축액공제사항"></OBJECT>');</SCRIPT></TD>
                                           <TD width="10%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=hfa050t_indiv_anu_amt name=txthfa050t_indiv_anu_amt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 125px" title=FPDOUBLESINGLE tag="24X2Z" ALT="개인연금저축액정산결과"></OBJECT>');</SCRIPT></TD>
                                           <TD BGCOLOR=#d1e8f9 width="30%">18.투자조합등소득공제</TD>
                                           <TD width="20%" colspan="2"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=hfa050t_invest_sub_sum_amt name=txtinvest_sub_sum_amt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 254px" title=FPDOUBLESINGLE tag="24X2Z" ALT="투자소득공제"></OBJECT>');</SCRIPT></TD>
                                      </TR>
                                      <TR>
                                           <TD BGCOLOR=#d1e8f9 width="30%">19.신용카드소득공제</TD>
                                           <TD width="20%" colspan="2"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=hfa050t_card_sub_sum_amt name=txtcard_sub_sum_amt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 254px" title=FPDOUBLESINGLE tag="24X2Z" ALT="신용카드소득공제"></OBJECT>');</SCRIPT></TD>
                                           <TD BGCOLOR=#d1e8f9 width="30%">20.우리사주출연금</TD>
                                           <TD width="20%" colspan="2"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=hfa050t_Our_Stock_sub_amt name=txtOur_Stock_sub_amt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 254px" title=FPDOUBLESINGLE tag="24X2Z" ALT="우리사주출연금"></OBJECT>');</SCRIPT></TD>
                                      </TR>                                                                          
                                      <TR>
                                           <TD BGCOLOR=#d1e8f9 width="30%">21.외국인교육비/임차료</TD>
                                           <TD width="10%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=hfa030t_fore_edu_amt name=hfa030t_fore_edu_amt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 125px" title=FPDOUBLESINGLE tag="24X2Z" ALT="외국인근로자의교육비공제사항"></OBJECT>');</SCRIPT></TD>
                                           <TD width="10%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=hfa050t_fore_edu_sub_amt name=hfa050t_fore_edu_sub_amt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 125px" title=FPDOUBLESINGLE tag="24X2Z" ALT="외국인근로자의교육비정산결과"></OBJECT>');</SCRIPT></TD>
                                           <TD BGCOLOR=#d1e8f9 width="30%">퇴직연금소득공제</TD>
                                           <TD width="10%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=hfa030t_retire_pension name=hfa030t_retire_pension style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 125px" title=FPDOUBLESINGLE tag="24X2Z" ALT="퇴직연금소득공제사항"></OBJECT>');</SCRIPT></TD>
                                           <TD width="10%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=hfa050t_retire_pension name=hfa050t_retire_pension style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 125px" title=FPDOUBLESINGLE tag="24X2Z" ALT="퇴직연금소득정산결과"></OBJECT>');</SCRIPT></TD>
                                      </TR>
                                   </TABLE>
            				       </FIELDSET>
                                   <BR>
                                   <FIELDSET CLASS="Clsdiv"><LEGEND ALGIN="LEFT">소득공제계 , 과세표준 , 산출세액</LEGEND>
                                   <table <%=LR_SPACE_TYPE_20%> border="1" width="100%">
                                      <TR>
                                           <TD BGCOLOR=#d1e8f9 width="17%">22.소득공제계</TD>
                                           <TD width="13%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=sum_amt name=txtsum_amt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 175px" title=FPDOUBLESINGLE tag="24X2Z" ALT="소득공제계"></OBJECT>');</SCRIPT></TD>
                                           <TD BGCOLOR=#d1e8f9 width="17%">23.소득과세표준</TD>
                                           <TD width="13%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=hfa050t_tax_std_amt name=txttax_std_amt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 175px" title=FPDOUBLESINGLE tag="24X2Z" ALT="소득과세표준"></OBJECT>');</SCRIPT></TD>
                                           <TD BGCOLOR=#d1e8f9 width="17%">24.산출세액</TD>
                                           <TD width="13%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=hfa050t_calu_tax_amt name=txtcalu_tax_amt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 175px" title=FPDOUBLESINGLE tag="24X2Z" ALT="산출세액"></OBJECT>');</SCRIPT></TD>
                                      </TR>
                                   </TABLE>
            				       </FIELDSET>
                                   <BR>
                                   <FIELDSET CLASS="Clsdiv"><LEGEND ALGIN="LEFT">세액공제 및 세액감면</LEGEND>
                                   <table <%=LR_SPACE_TYPE_20%> border="1" width="100%">
                                      <TR>
                                           <TD BGCOLOR=#d1e8f9 width="30%">25.근로소득세액공제</TD>
                                           <TD width="20%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=hfa050t_income_tax_sub_amt name=txtincome_tax_sub_amt style="HEIGHT: 100%; LEFT: 0px; TOP: 0px; WIDTH: 254px" title=FPDOUBLESINGLE tag="24X2Z" ALT="근로소득"></OBJECT>');</SCRIPT></TD>
                                           <TD BGCOLOR=#d1e8f9 width="30%">26.주택자금차입금이자상환액</TD>
                                           <TD width="20%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=hfa050t_house_repay_amt name=txthouse_repay_amt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 254px" title=FPDOUBLESINGLE tag="24X2Z" ALT="주택자금차입금이자상환액"></OBJECT>');</SCRIPT></TD>
                                      </TR>
                                      
                                      <TR>
									       <TD BGCOLOR=#d1e8f9 width="30%">27.외국납부세액공제</TD>
                                           <TD width="20%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="Object2" name=txtFore_pay style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 254px" title=FPDOUBLESINGLE tag="24X2Z" ALT="외국납부세액공제"></OBJECT>');</SCRIPT></TD>
									       <TD BGCOLOR=#d1e8f9 width="30%">28.정치자금기부금세액공제</TD>
                                           <TD width="20%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="Object2" name=txtPolicontr_tax_sub_amt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 254px" title=FPDOUBLESINGLE tag="24X2Z" ALT="정치자금기부금세액공제"></OBJECT>');</SCRIPT></TD>
                                      </TR>
                                      <TR>
									       <TD BGCOLOR=#d1e8f9 width="30%">을근납세조합공제</TD>
                                           <TD width="20%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="Object2" name=txtTax_Union_Ded style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 254px" title=FPDOUBLESINGLE tag="24X2Z" ALT="을근납세조합공제"></OBJECT>');</SCRIPT></TD>
                                           <TD BGCOLOR=#d1e8f9 width="30%">29.세액공제계</TD>
                                           <TD width="20%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=hfa050t_tax_sub_sum_amt name=txttax_sub_sum_amt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 254px" title=FPDOUBLESINGLE tag="24X2Z" ALT="세액공제계"></OBJECT>');</SCRIPT></TD>
									  </TR>	

                                   </TABLE>
            				       </FIELDSET>
                                   <BR>
                                   <FIELDSET CLASS="Clsdiv"><LEGEND ALGIN="LEFT">결정세액/차감징수세액</LEGEND>
                                   <table <%=LR_SPACE_TYPE_20%> border="1" width="100%">
                                       <TR>
                                           <TD BGCOLOR=#d1e8f9 width="44%" align="middle">구분</TD>
                                           <TD BGCOLOR=#d1e8f9 width="14%" align="middle">소득세</TD>
                                           <TD BGCOLOR=#d1e8f9 width="14%" align="middle">주민세</TD>
                                           <TD BGCOLOR=#d1e8f9 width="14%" align="middle">농특세</TD>
                                           <TD BGCOLOR=#d1e8f9 width="14%" align="middle">계</TD>
                                      </TR>
                                      <TR>
                                           <TD BGCOLOR=#d1e8f9 width="44%">30.정산세액</TD>
                                           <TD width="14%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=hfa050t_dec_income_tax_amt name=txtdec_income_tax_amt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 180px" title=FPDOUBLESINGLE tag="24X2Z" ALT="소득세"></OBJECT>');</SCRIPT></TD>
                                           <TD width="14%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=hfa050t_dec_res_tax_amt name=txtdec_res_tax_amt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 180px" title=FPDOUBLESINGLE tag="24X2Z" ALT="주민세"></OBJECT>');</SCRIPT></TD>
                                           <TD width="14%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=hfa050t_dec_farm_tax_amt name=txtdec_farm_tax_amt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 180px" title=FPDOUBLESINGLE tag="24X2Z" ALT="농특세"></OBJECT>');</SCRIPT></TD>
                                           <TD width="14%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=dec_amt name=txtdec_amt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 180px" title=FPDOUBLESINGLE tag="24X2Z" ALT="계"></OBJECT>');</SCRIPT></TD>
                                      </TR>
                                      <TR>
                                           <TD BGCOLOR=#d1e8f9 width="44%">31.현근무지징수세액</TD>  
                                           <TD width="14%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=hfa050t_new_income_tax_amt name=txtnew_income_tax_amt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 180px" title=FPDOUBLESINGLE tag="24X2Z" ALT="소득세"></OBJECT>');</SCRIPT></TD>
                                           <TD width="14%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=hfa050t_new_res_tax_amt name=txtnew_res_tax_amt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 180px" title=FPDOUBLESINGLE tag="24X2Z" ALT="주민세"></OBJECT>');</SCRIPT></TD>
                                           <TD width="14%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=hfa050t_new_farm_tax_amt name=txtnew_farm_tax_amt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 180px" title=FPDOUBLESINGLE tag="24X2Z" ALT="농특세"></OBJECT>');</SCRIPT></TD>
                                           <TD width="14%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=new_amt name=txtincome_amt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 180px" title=FPDOUBLESINGLE tag="24X2Z" ALT="계"></OBJECT>');</SCRIPT></TD>
                                      </TR>
                                      <TR>
                                           <TD BGCOLOR=#d1e8f9 width="44%">32.종전근무지세액</TD>
                                           <TD width="14%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=hfa050t_old_income_tax_amt name=txtold_income_tax_amt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 180px" title=FPDOUBLESINGLE tag="24X2Z" ALT="소득세"></OBJECT>');</SCRIPT></TD>
                                           <TD width="14%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=hfa050t_old_res_tax_amt name=txtold_res_tax_amt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 180px" title=FPDOUBLESINGLE tag="24X2Z" ALT="주민세"></OBJECT>');</SCRIPT></TD>
                                           <TD width="14%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=hfa050t_old_farm_tax_amt name=txtold_farm_tax_amt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 180px" title=FPDOUBLESINGLE tag="24X2Z" ALT="농특세"></OBJECT>');</SCRIPT></TD>
                                           <TD width="14%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=old_amt name=txtold_amt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 180px" title=FPDOUBLESINGLE tag="24X2Z" ALT="계"></OBJECT>');</SCRIPT></TD>
                                      </TR>
                                      <TR>
                                           <TD BGCOLOR=#d1e8f9 width="44%">33.징수해야할세액</TD>
                                           <TD width="14%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=hfa050t_income_tax_amt name=txtincome_tax_amt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 180px" title=FPDOUBLESINGLE tag="24X2Z" ALT="소득세"></OBJECT>');</SCRIPT></TD>
                                           <TD width="14%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=hfa050t_res_tax_amt name=txtres_tax_amt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 180px" title=FPDOUBLESINGLE tag="24X2Z" ALT="주민세"></OBJECT>');</SCRIPT></TD>
                                           <TD width="14%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=hfa050t_farm_tax_amt name=txtfarm_tax_amt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 180px" title=FPDOUBLESINGLE tag="24X2Z" ALT="농특세"></OBJECT>');</SCRIPT></TD>
                                           <TD width="14%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=f_amt name=txtf_amt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 180px" title=FPDOUBLESINGLE tag="24X2Z" ALT="계"></OBJECT>');</SCRIPT></TD>
                                      </TR>
                                   </TABLE>
                                   </FIELDSET> <!-- Content Area **** Single **** -->
                                </TD>
                              </TR>
                        </TABLE>
                     </TD>
                 </TR>
<!-- Space Area -->
	
<!-- Button, Batch, Print, Jump Area -->
            </TABLE>
        </TD>
    </TR>
	<TR >
	    <TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP"  SRC = "../../blank.htm"  WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=no noresize framespacing=0></IFRAME>  
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






