
<%@ LANGUAGE="VBSCRIPT" %>
<%Response.Expires = -1%>
<!--'======================================================================================================
'*  1. Module Name          : EIS
'*  2. Function Name        : 
'*  3. Program ID           : 
'*  4. Program Name         : 
'*  5. Program Desc         : 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Cho Ig Sung
'* 10. Modifier (Last)      : 
'* 11. Comment              : 
'======================================================================================================= -->


<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!--
'##########################################################################################################
'												1. 선 언 부 
'##########################################################################################################

'******************************************  1.1 Inc 선언   **********************************************
'	기능: Inc. Include
'*********************************************************************************************************  -->
<!-- #Include file="../../inc/incSvrCcm.inc"  -->
<!-- #Include file="../../inc/incSvrHTML.inc"  -->
<!-- #Include file="../../inc/incEISComm.asp"  -->
<!--
'===============================================  1.1.1 Style Sheet  ======================================
'========================================================================================================== -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/button.css">

<!--'==========================================  1.1.2 공통 Include   ======================================
'========================================================================================================== -->
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAMain.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAEvent.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliVariables.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAOperation.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliRdsQuery.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incEB.vbs">				</SCRIPT>
<Script Language="JavaScript"	SRC="../../inc/incImage.js">			</SCRIPT>
<SCRIPT LANGUAGE="VBScript">

Option Explicit						'☜: indicates that All variables must be declared in advance

'----------------  공통 Global 변수값 정의  ----------------------------------------------------------- 
Dim IsOpenPop
Dim intRetCD

'++++++++++++++++  Insert Your Code for Global Variables Assign  ++++++++++++++++++++++++++++++++++++++ 

<!-- #Include file="../../inc/lgvariables.inc" --> 

Const BIZ_PGM_ID = "VS004MB1.asp"  


'==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'========================================================================================================= 
Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed

    '---- Coding part--------------------------------------------------------------------    
End Sub

'========================================  2.2.1 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
'========================================================================================================= 

Sub SetDefaultVal()
	txtQueryDate.Text	= UniConvDateAToB("<%=GetSvrDate%>" ,parent.gServerDateFormat,gDateFormat)
	Call ggoOper.FormatDate(txtQueryDate, gDateFormat, 2)
	cboUnit.value = "1000"
End Sub

'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== 

Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% Call LoadInfTB19029A("Q", "V", "NOCOOKIE", "OA") %>
End Sub

'******************************************  2.4 POP-UP 처리함수  ****************************************
'	기능: 기준 POP-UP
'   설명: 기준 POP-UP에 관한 Open은 Include한다. 
'	      하나의 ASP에서 Popup이 중복되면 하나는 링크시켜 사용하고 나머지는 재정의하여 사용한다.
'********************************************************************************************************* 

'+++++++++++++++++++  Insert Your Code for PopUp(Open)  ++++++++++++++++++++++++++++++++++++++++++++++++++ 
Function OpenPopup(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	Select Case iWhere
		Case 0
			arrParam(0) = "거래처팝업"
			arrParam(1) = "B_BIZ_PARTNER"
			arrParam(2) = strCode
			arrParam(3) = ""
			arrParam(4) = ""
			arrParam(5) = "거래처"			
	
			arrField(0) = "BP_CD"
			arrField(1) = "BP_NM"
			arrField(2) = "BP_FULL_NM"
			 
			arrHeader(0) = "거래처코드"
			arrHeader(1) = "거래처명"
			arrHeader(2) = "거래처전명"

		Case Else
			Exit Function

	End Select

	IsOpenPop = True

	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=540px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Select Case iWhere
			Case 0
				txtBpCd.focus

			Case Else
		End Select

		Exit Function
	Else
		Call SetPopup(arrRet, iWhere)
	End If	

End Function

'==========================================  2.4.3 Set???()  =============================================
'	Name : Set???()
'	Description : PopUp에서 전송된 값을 특정 Tag Object에 지정 
'========================================================================================================= 
Function SetPopup(Byval arrRet, Byval iWhere)
	
	Select Case iWhere
		Case 0
			txtBpCd.focus
			txtBpCd.value = arrRet(0)
			txtBpNm.value = arrRet(1)

		Case Else
			Exit Function

	End Select

End Function

'==============================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분 
'========================================================================================================= 
Sub Form_Load()

    Call LoadInfTB19029                           '⊙: Load table , B_numeric_format

    Call ggoOper.ClearField(Document, "1")        '⊙: Condition field clear
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")         '⊙: 조건에 맞는 Field locking
 
    Call InitVariables                            '⊙: Initializes local global Variables
                  
	Call InitComboBox
	
	Call SetDefaultVal
        
    Call SetToolbar("10000000000011")				'⊙: 버튼 툴바 제어 

End Sub

'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================

Sub Form_QueryUnload(Cancel , UnloadMode )
End Sub

'======================================================================================================
'   Event Name : txtQueryDate_DblClick
'   Event Desc : 달력 Popup을 호출 
'=======================================================================================================
Sub txtQueryDate_DblClick(Button)
    If Button = 1 Then
        fpDateTime1.Action = 7
    End If
End Sub

'==========================================  2.2.6 InitComboBox()  =======================================
'	Name : InitComboBox()
'	Description : Combo Display
'========================================================================================================= 
Sub InitComboBox()

	Err.clear
	
	Call CommonQueryRs("MINOR_CD, MINOR_NM", "B_MINOR", "MAJOR_CD = " & FilterVar("V0001", "''", "S") , lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)

	Call SetCombo2(cboUnit ,lgF0  ,lgF1  ,Chr(11))
	
End Sub

'======================================================================================================
' Function Name : SetPrintCond
' Function Desc : This function is related to Print/Preview Button
'=======================================================================================================

Function SetPrintCond(StrEbrFile, strUrl)

	Dim strSpId, strUnit,strPlantCd
		
	SetPrintCond = False

	strSpId		 = FilterVar(txtSpId.value,"","SNM")
	strUnit	  	 = FilterVar(cboUnit.value,"","SNM")
	strPlantCd	 = FilterVar(txtPlantCd.value,"","SNM")
	StrEbrFile	 = "VS004MA1"

	StrUrl = StrUrl & "strSpId|" & strSpId
	StrUrl = StrUrl & "|strUnit|" & strUnit
	StrUrl = StrUrl & "|PlantCd|" & strPlantCd

	SetPrintCond = True
	
End Function

'========================================================================================
' Function Name : FncBtnPreview
' Function Desc : This function is related to Preview Button
'========================================================================================

Function FncBtnPreview() 
	'On Error Resume Next                                                    '☜: Protect system from crashing
    
    Dim StrUrl, StrEbrFile, ObjName

    If Not chkField(Document, "1") Then									'⊙: This function check indispensable field
       Call LayerShowHide(0)
       Exit Function
    End If

	If Not SetPrintCond(StrEbrFile, strUrl) Then
		Exit Function
	End If
	
	ObjName = AskEBDocumentName(StrEbrFile,"ebr")
	lgEBProcessbarOut = "T"
	EBActionA.menu.value = 0
    Call FncEBR5RC2(ObjName, "view", StrUrl,EBActionA,"EBR")	
		
End Function

'======================================================================================================
' Function Name : FncBtnPrint
' Function Desc : This function is related to Print Button
'=======================================================================================================
Function FncBtnPrint() 
	'On Error Resume Next                                                    '☜: Protect system from crashing

    Dim StrUrl, StrEbrFile, ObjName
	
    If Not chkField(Document, "1") Then									'⊙: This function check indispensable field
       Call LayerShowHide(0)
       Exit Function
    End If

	If Not SetPrintCond(StrEbrFile, strUrl) Then
		Exit Function
	End If
	
	ObjName = AskEBDocumentName(StrEbrFile,"ebr")
	Call FncEBRPrint(EBAction,ObjName,StrUrl)	
		
End Function

'======================================================================================================
' Function Name :
' Function Desc :
'=======================================================================================================
Function DoSpExec(strPrintOpt)		'Preview, Print
	Dim strVal       
	Dim strYyyyMm, strBpCd,strPlantCd

    lgEBProcessbarOut = "F"

	If Not chkField(Document, "1") Then                             '⊙: Check contents area
		Exit Function
	End If

	' 화면 초기화 
	MyBizASP1.location.href = "../../blank.htm"

	strYyyyMm	= txtQueryDate.year & right("0" & txtQueryDate.month,2)
	strBpCd		= Trim(txtBpCd.value)
	strPlantCd		= Trim(txtPlantCd.value)
	
	If strBpCd = "" Then 	strBpCd		= "%"
	If strPlantCd = "" Then 	strPlantCd		= "%"

	' 미리보기시 작업중 메세지 처리 
	If strPrintOpt = "Preview" Then
		Call LayerShowHide(1) 
	End If
	    
	strVal = BIZ_PGM_ID & "?txtMode="	& Parent.UID_M0002
	strVal = strVal & "&strPrintOpt="	& strPrintOpt
	strVal = strVal & "&strYyyyMm="		& strYyyyMm
	strVal = strVal & "&strPlantCd="	& strPlantCd
	strVal = strVal & "&strBpCd="		& strBpCd

	Call RunMyBizASP(MyBizASP, strVal)

End Function

'========================================================================================
' Function Name : MyBizASP1_OnReadyStateChange
' Function Desc : 
'========================================================================================

Sub MyBizASP1_onreadystatechange()
	If lgEBProcessbarOut = "T" Then		
	   Call LayerShowHide(0)
	   lgEBProcessbarOut = "F"  
	End  If   
End Sub



'------------------------------------------  OpenCondPlant()  --------------------------------------------
'	Name : OpenCondPlant()
'	Description : Condition Plant PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenConPlant()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(txtPlantCd.className) = UCase(parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "공장팝업"				' 팝업 명칭 
	arrParam(1) = "B_PLANT"						' TABLE 명칭 
	arrParam(2) = Trim(txtPlantCd.Value)	' Code Condition
	arrParam(3) = ""							' Name Cindition
	arrParam(4) = ""							' Where Condition
	arrParam(5) = "공장"					' TextBox 명 
	
    arrField(0) = "PLANT_CD"					' Field명(0)
    arrField(1) = "PLANT_NM"					' Field명(1)
    
    arrHeader(0) = "공장"					' Header명(0)
    arrHeader(1) = "공장명"					' Header명(1)
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
			txtPlantCd.Value    = arrRet(0)		
			txtPlantNm.Value    = arrRet(1)		
	End If	
	
	Call SetFocusToDocument("M")
	txtPlantCd.focus
	
End Function


'========================================================================================================
'========================================================================================================
'   Event Name : txtBizAreaCd_onChange
'   Event Desc : 
'========================================================================================================
Sub txtBizAreaCd_onChange()
	Dim IntRetCD
	Dim arrVal

	If txtBizAreaCd.value = "" Then Exit Sub

	If CommonQueryRs("BIZ_AREA_NM", "B_BIZ_AREA ", " BIZ_AREA_CD=  " & FilterVar(txtBizAreaCd.value , "''", "S") & "", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) Then
		arrVal = Split(lgF0, Chr(11)) 
		txtBizAreaNm.value= Trim(arrVal(0)) 
	Else
		IntRetCD = DisplayMsgBox("124200","X","X","X")  	'%1 사업장 데이터가 존재하지 않습니다.
		txtBizAreaCd.focus
	End If
End Sub
'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================

Function FncExit()
    FncExit = True
End Function

'========================================================================================================
'   Event Name : txtBpCd_onChange
'   Event Desc : 
'========================================================================================================
Sub txtBpCd_onChange()
	Dim IntRetCD
	Dim arrVal

	If txtBpCd.value = "" Then 
		txtBpNm.value	= ""
		Exit Sub
	End If

	If CommonQueryRs("BP_NM", "B_BIZ_PARTNER ", " BP_CD=  " & FilterVar(txtBpCd.value , "''", "S") & "", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) Then
		arrVal = Split(lgF0, Chr(11)) 
		txtBpNm.value= Trim(arrVal(0)) 
	Else
		IntRetCD = DisplayMsgBox("971012","X","거래처","X")  	'%1 거래처 정보가 없습니다.
		txtBpNm.value=""
		txtBpCd.focus
	End If
End Sub

'========================================================================================
' Function Name : txtPlantCd_cd_OnChange
' Function Desc : 
'========================================================================================
Function txtPlantCd_OnChange()
    Dim IntRetCd

    If txtPlantCd.value = "" Then
        txtPlantnm.value = ""
    ELSE    
        IntRetCd =  CommonQueryRs(" PLANT_NM "," B_PLANT "," PLANT_CD="&filterVar(txtPlantCd.value,"''","S") & "" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
        If IntRetCd = false Then
			 call DisplayMsgBox("971012","X", "공장","X")
			 txtPlantnm.value=""
			 txtPlantCd.focus
			
        Else
            txtPlantnm.value=Trim(Replace(lgF0,Chr(11),""))
        End If
    End If
    
End Function


'☜: 아래 OBJECT Tag는 InterDev 사용자를 위한것으로 프로그램이 완성되면 아래 Include 코드로 대체되어야 한다 
</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->	

<SCRIPT LANGUAGE="JavaScript">
<!-- Hide script from old browsers
function setCookie(name, value, expire)
{
    //alert(value)
    //alert(escape(value))
	document.cookie = name + "=" + escape(value)
		+ ((expire == null) ? "" : ("; expires=" + expire.toGMTString()))
	document.cookie = name + "=" + escape(value)
		+ "; path=/EasyBaseWeb/bin"
		+ ((expire == null) ? "" : ("; expires=" + expire.toGMTString()))
	document.cookie = name + "=" + escape(value)
		+ "; path=/EasyBaseWeb/lib"
		+ ((expire == null) ? "" : ("; expires=" + expire.toGMTString()))
}

setCookie("client", "-1", null)
setCookie("owner", "admin", null)
setCookie("identity", "admin", null)
 -->
</SCRIPT>

</HEAD>
<!--
'#########################################################################################################
'       					6. Tag부 
'#########################################################################################################  -->

<BODY TABINDEX="-1" SCROLL="no">

    <%
        Call PrintTitle(Request("strASPMnuMnuNm"))
    %>

	<CENTER>
	<TABLE  <%=TABSTYLE01%> >
		<TR>
			<TD CLASS="TD5E" NOWRAP>조회일자</TD>
			<TD CLASS="TD6E" NOWRAP><script language =javascript src='./js/vs004ma1_fpDateTime1_txtQueryDate.js'></script>&nbsp;</TD>
			<TD CLASS="TD5E" NOWRAP>금액단위</TD>
			<TD CLASS="TD6E" NOWRAP><SELECT NAME="cboUnit" tag="12X" STYLE="WIDTH:120px:" ALT="금액단위"><OPTION VALUE="" selected></OPTION></SELECT></TD>
		</TR>
		
		<TR>
			<TD CLASS="TD5E" NOWRAP>공장</TD>
			<TD CLASS="TD6E" NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtPlantCd" SIZE=6 MAXLENGTH=4 tag="11xxxU" ALT="공장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenConPlant()">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=25 tag="14" ALT="공장명"></TD>
			<TD CLASS="TD5E" NOWRAP>거래처</TD>
			<TD CLASS="TD6E" NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtBpCd" SIZE=10 MAXLENGTH=10 tag="11XXXU" ALT="거래처코드"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBpCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPopup txtBpCd.value, 0">
										<INPUT TYPE=TEXT NAME="txtBpNm" SIZE=30 tag="14">
			</TD>
		</TR>
		
	</TABLE>
	
	<TABLE width=1016  height=476 cellspacing=0 cellpadding=0 border=0>
		<TR>
			<TD><IFRAME NAME="MyBizASP1"  WIDTH=100% HEIGHT=100% FRAMEBORDER=0 SCROLLING=auto framespacing=0 marginwidth=0 marginheight=0 ></IFRAME></TD>
		</TR>
		<TR>
			<TD WIDTH=100% HEIGHT=1><IFRAME NAME="MyBizASP" SRC = "../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=No noresize framespacing=0 TABINDEX="-1"></IFRAME>
			</TD>
		</TR>
	</table>
	</center>

	<INPUT TYPE="HIDDEN" NAME="txtSpId" tag="24" TABINDEX = "-1">

	<FORM NAME="EBActionA" ID="EBAction" TARGET="MyBizASP1" METHOD="POST"  scroll=yes> 
		<input type="hidden" name="menu" value=0 > 
		<input type="hidden" name="id" > 
		<input type="hidden" name="pw" >
		<input type="hidden" name="doc" > 
		<input type="hidden" name="form" > 
		<input type="hidden" name="runvar" > 
	</FORM>

	<DIV ID="MousePT" NAME="MousePT">
		<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
	</DIV>
	
	<FORM NAME=EBAction TARGET="MyBizASP" METHOD="POST">
		<INPUT TYPE="HIDDEN" NAME="uname"		TABINDEX = "-1" >
		<INPUT TYPE="HIDDEN" NAME="dbname"		TABINDEX = "-1" >
		<INPUT TYPE="HIDDEN" NAME="filename"	TABINDEX = "-1" >
		<INPUT TYPE="HIDDEN" NAME="condvar"		TABINDEX = "-1" >
		<INPUT TYPE="HIDDEN" NAME="date"		TABINDEX = "-1" >	
	</FORM>
</BODY>
</HTML>

