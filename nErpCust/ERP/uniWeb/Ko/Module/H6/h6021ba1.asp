<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : Human Resources
*  2. Function Name        : 급여자동기표처리 
*  3. Program ID           : H6021ba1
*  4. Program Name         : H6021ba1
*  5. Program Desc         : 급여자동기표처리 
*  6. Comproxy List        :
*  7. Modified date(First) : 2001/06/07
*  8. Modified date(Last)  : 2003/06/13
*  9. Modifier (First)     : 송봉규 
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
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs"></SCRIPT>

<Script Language="VBScript">
Option Explicit

'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
Const BIZ_PGM_ID  = "H6021bb1.asp"
Const BIZ_PGM_ID2 = "H6021bb2.asp"
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
    
    lgStrPrevKey = ""                           'initializes Previous Key
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
	
	frm1.txtprov_dt.focus

	Call ExtractDateFrom("<%=GetsvrDate%>",Parent.gServerDateFormat , Parent.gServerDateType ,strYear,strMonth,strDay)
	
	frm1.txtprov_dt.Year = strYear 		 '년월일 default value setting
	frm1.txtprov_dt.Month = strMonth 
	frm1.txtprov_dt.Day = strDay
	frm1.txtacct_dt.Year = strYear 		 '년월일 default value setting
	frm1.txtacct_dt.Month = strMonth 
	frm1.txtacct_dt.Day = strDay
End Sub
	
'========================================================================================================
' Name : LoadInfTB19029()	
' Desc : Set System Number format
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "H", "NOCOOKIE", "BA") %>
End Sub

'========================================================================================================
'	Name : CookiePage()
'	Description : Item Popup에서 Return되는 값 setting
'========================================================================================================
Function CookiePage(ByVal flgs)
    Const CookieSplit = 4877	
    
	Dim strTemp

	If flgs = 0 Then                                       '☜: h6012ma1.asp 의 쿠기값을 받고 있음.
		strTemp = ReadCookie("PROV_DT")                      '         절대수정금지 요망........!    
		If strTemp = "" then Exit Function
		
        frm1.txtprov_dt.text = ReadCookie("PROV_DT")
		frm1.txtprov_type.value = ReadCookie("PROV_TYPE")
		frm1.txtprov_type_nm.value = ReadCookie("PROV_TYPE_NM")

        If ReadCookie("TRANS_DT") <> "" Then
        	frm1.txtacct_dt.text = ReadCookie("TRANS_DT")
    	End If
    	
		MainQuery()              
		WriteCookie "PROV_DT" , ""
	    WriteCookie "PROV_TYPE" , ""
        WriteCookie "TRANS_DT"  , ""
        
	ElseIf flgs = 1 Then 
        WriteCookie "PROV_DT" , frm1.txtprov_dt.text
	End IF
End Function

'========================================================================================================
' Function Name : MakeKeyStream
' Function Desc : This method set focus to pos of err
'========================================================================================================
Sub MakeKeyStream(pOpt)
   With frm1
	
	    lgKeyStream = .txtprov_dt.Text & Parent.gColSep
	    lgKeyStream = lgKeyStream & .txtProv_type.Text & Parent.gColSep
	    lgKeyStream = lgKeyStream & .txtacct_dt.Text & Parent.gColSep
	   
   End With   
End Sub        

'========================================================================================================
' Name : Form_Load
' Desc : developer describe this line Called by Window_OnLoad() evnt
'========================================================================================================
Sub Form_Load()

    Err.Clear                                                                       '☜: Clear err status
	Call LoadInfTB19029                                                             '☜: Load table , B_numeric_format
		
	Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
	Call InitVariables                                                     '⊙: Setup the Spread sheet
    Call ggoOper.FormatDate(frm1.txtprov_dt, Parent.gDateFormat, 1)
    Call ggoOper.FormatDate(frm1.txtacct_dt, Parent.gDateFormat, 1)

	Call SetDefaultVal
	Call SetToolbar("1000000000000111")										'⊙: 버튼 툴바 제어 
	Call CookiePage(0)                                                             '☜: Check Cookie
			
End Sub
	
'========================================================================================================
' Name : Form_QueryUnload
' Desc : developer describe this line Called by Window_OnUnLoad() evnt
'========================================================================================================
Sub Form_QueryUnload(Cancel, UnloadMode)

End Sub
	
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

End Function
	
'========================================================================================================
' Function Name : DbDeleteOk
' Function Desc : Called by MB Area when delete operation is successful
'========================================================================================================
Function DbDeleteOk()
	Call InitVariables()
End Function

Function OpenCondAreaPopup(Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True  Then  
	   Exit Function
	End If   

	IsOpenPop = True
	Select Case iWhere
	
        Case "1"
            arrParam(0) = "지급구분 팝업"			' 팝업 명칭 
	        arrParam(1) = "B_MINOR"				 		' TABLE 명칭 
	        arrParam(2) = frm1.txtProv_type.value		    ' Code Condition
	        arrParam(3) = ""'frm1.txtProv_type_nm.value		' Name Cindition
	        arrParam(4) = " MAJOR_CD = " & FilterVar("H0040", "''", "S") & " "    ' Where Condition	 'unicode
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
		frm1.txtProv_type.focus
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
	With Frm1
		Select Case iWhere
		    Case "1"
		        .txtProv_type.value = arrRet(0)
		        .txtProv_type_nm.value = arrRet(1)
		        .txtProv_type.focus
        End Select
	End With

End Sub
'========================================================================================================
'   Event Name : txtProv_type_Onchange()            '<==코드만 입력해도 앤터키,탭키를 치면 코드명을 불러준다 
'   Event Desc :
'========================================================================================================
Function txtProv_type_Onchange()
    Dim iDx
    Dim IntRetCd
    
    IF frm1.txtProv_type.value = "" THEN
        frm1.txtProv_type_nm.value = ""
    ELSE
        IntRetCd = CommonQueryRs(" minor_nm "," b_minor "," major_cd = " & FilterVar("H0040", "''", "S") & " and minor_cd =  " & FilterVar(frm1.txtProv_type.value , "''", "S") & "" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) 'unicode
        If IntRetCd = false then
			Call DisplayMsgBox("800140","X","X","X")	'지급내역코드에 등록되지 않은 코드입니다.
            frm1.txtProv_type_nm.value = ""
            frm1.txtProv_type.focus
            txtProv_type_Onchange = true
        ELSE    
            frm1.txtProv_type_nm.value = Trim(Replace(lgF0,Chr(11),""))   '수당코드 
        END IF
    END IF 
End Function 

'======================================================================================================
' Function Name : ExeReflect
' Function Desc : 
'=======================================================================================================
Function ExeReflect()
	Dim strVal
	Dim strprov_dt, stracct_dt
	Dim IntRetCD

	ExeReflect = False                                                          '⊙: Processing is NG
    
	On Error Resume Next                                                   '☜: Protect system from crashing

	If Not chkField(Document, "1") Then	
       Call BtnDisabled(0)
       Exit Function            								         '☜: This function check required field
    End If
    if txtProv_type_Onchange() then
		Exit Function
	end if

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
	   
    strprov_dt = UniConvDateToYYYYMMDD(frm1.txtprov_dt.text, Parent.gDateFormat, Parent.gComDateType)
    strprov_dt = left(strprov_dt,4) & mid(strprov_dt,6,2) & mid(strprov_dt, 9,2)
    stracct_dt = UniConvDateToYYYYMMDD(frm1.txtacct_dt.text, Parent.gDateFormat, Parent.gComDateType)
    stracct_dt = left(stracct_dt,4) & mid(stracct_dt,6,2) & mid(stracct_dt, 9,2)

	strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0006
	strVal = strVal & "&txtprov_dt=" & strprov_dt
	strVal = strVal & "&txtProv_type=" & frm1.txtProv_type.value
	strVal = strVal & "&txtacct_dt=" & stracct_dt

	Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 

	ExeReflect = True                                                           '⊙: Processing is NG
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
End Function

'======================================================================================================
' Function Name : ExeCancel
' Function Desc : 
'=======================================================================================================
Function ExeCancel()
	Dim strVal
	Dim strprov_dt, stracct_dt
	Dim IntRetCD

	ExeCancel = False                                                          '⊙: Processing is NG
	On Error Resume Next                                                   '☜: Protect system from crashing

	If Not chkField(Document, "1") Then	
       Call BtnDisabled(0)
       Exit Function            								         '☜: This function check required field
    End If
     if txtProv_type_Onchange() then
		Exit Function
	end if
   
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
	   
    strprov_dt = UniConvDateToYYYYMMDD(frm1.txtprov_dt.text, Parent.gDateFormat, Parent.gComDateType)
    strprov_dt = left(strprov_dt,4) & mid(strprov_dt,6,2) & mid(strprov_dt, 9,2)
    stracct_dt = UniConvDateToYYYYMMDD(frm1.txtacct_dt.text, Parent.gDateFormat, Parent.gComDateType)
    stracct_dt = left(stracct_dt,4) & mid(stracct_dt,6,2) & mid(stracct_dt, 9,2)

	strVal = BIZ_PGM_ID2 & "?txtMode=" & Parent.UID_M0006
	strVal = strVal & "&txtprov_dt=" & strprov_dt
	strVal = strVal & "&txtProv_type=" & frm1.txtProv_type.value
	strVal = strVal & "&txtacct_dt=" & stracct_dt

	Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 

	ExeCancel = True                                                           '⊙: Processing is NG
	Call BtnDisabled(0)
End Function

'======================================================================================================
' Function Name : ExeCancelOk
' Function Desc : ExeCancel가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncSave에 있는것을 옮김 
'=======================================================================================================
Function ExeCancelOk()				            '☆: 저장 성공후 실행 로직 
	Dim IntRetCD 

	IntRetCD =DisplayMsgBox("990000","X","X","X")

End Function

Sub txtprov_dt_DblClick(Button)
	If Button = 1 Then
		Call SetFocusToDocument("M")	
		frm1.txtprov_dt.Action = 7
		frm1.txtprov_dt.focus
	End If
End Sub

Sub txtacct_dt_DblClick(Button)
	If Button = 1 Then
		Call SetFocusToDocument("M")
		frm1.txtacct_dt.Action = 7
		frm1.txtacct_dt.focus
	End If
End Sub
'========================================================================================================
' Name : OpenProveDt
' Desc : 최근 지급일 POPUP
'========================================================================================================
Function OpenProveDt(iWhere)
	Dim arrRet
	Dim arrParam(4)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "지급일팝업"		
	arrParam(1) = "지급일"
	arrParam(2) = "hdf070t"
	arrParam(3) = "PROV_DT"
	arrParam(4) = frm1.txtprov_dt.text	

	arrRet = window.showModalDialog(HRAskPRAspName("StandardDtPopup"), Array(window.parent,arrParam), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	if arrRet(0) <> ""	then
		frm1.txtprov_dt.text = arrRet(0)
	end if
	frm1.txtprov_dt.focus
End Function

</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
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
					<TD >
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>급여자동기표처리</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
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
								<TD CLASS="TD5" NOWRAP>지급일자</TD>
								<TD CLASS="TD6" NOWRAP><script language =javascript src='./js/h6021ba1_txtprov_dt_txtprov_dt.js'></script>
								<IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnProveDt" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenProveDt(0)"></td>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>지급구분</TD>
								<TD CLASS="TD6" NOWRAP><INPUT NAME="txtProv_type" MAXLENGTH="1"  SIZE="10" ALT ="지급구분" TAG="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" ALIGN=top TYPE="BUTTON" ONCLICK="VBScript: OpenCondAreaPopup(1)">
								                       <INPUT NAME="txtProv_type_nm" MAXLENGTH="20" SIZE="20" ALT ="지급구분" tag="14"></TD>
	                        </TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>전표생성일자</TD>
								<TD CLASS="TD6" NOWRAP><script language =javascript src='./js/h6021ba1_txtacct_dt_txtacct_dt.js'></script>
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
  					     <BUTTON NAME="btnExe" CLASS="CLSSBTN" onclick="ExeReflect()" Flag=1>실행</BUTTON>&nbsp;
  					     <BUTTON NAME="btnExe" CLASS="CLSSBTN" onclick="ExeCancel()" Flag=1>취소</BUTTON>
		            </TD>
					<TD WIDTH=* ALIGN="right">&nbsp;</TD>
					<TD WIDTH=10>&nbsp;</TD>
		        </TR>
		    </TABLE>
		</TD>
	</TR>
	<TR>
		<TD HEIGHT=20><IFRAME NAME="MyBizASP" SRC = "../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=no noresize framespacing=0></IFRAME>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>


