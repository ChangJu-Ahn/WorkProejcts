<%@ LANGUAGE="VBSCRIPT" %>
<% Response.Expires = -1%>

<HTML>
<HEAD>
<TITLE><%=Request("strTitle")%></TITLE>

<!-- #Include file="../ESSinc/IncServer.asp"  -->

<LINK REL="stylesheet" TYPE="Text/css" href="../ESSinc/common.css">

<SCRIPT LANGUAGE="VBScript"   SRC="../ESSinc/ccm.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../ESSinc/variables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../ESSinc/incCookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../ESSinc/operation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../ESSinc/incCommFunc.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../ESSinc/incEvent.vbs">   </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../ESSinc/adoQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../ESSinc/incEB.vbs"></SCRIPT>
<!-- #Include file="../ESSinc/incSvrVarSims.inc"  -->
<!-- #Include file="../ESSinc/incSvrFuncSims.inc" -->
<%
	dim lgEmpn_no
	dim lgPay_Yymm
	dim lgProv_Type
	
	lgEmp_no = request("Emp_no")
	lgPay_Yymm = request("Pay_Yymm")
	lgProv_Type = request("Prov_Type")

%>
<Script Language="VBScript">
Option Explicit 

Const BIZ_PGM_ID      = "e1202mb4.asp"						           '☆: Biz Logic ASP Name
Const C_SHEETMAXROWS = 6	                                      '☜: Visble row
Const C_SHEETMAXCOLS = 10

<!-- #Include file="../ESSinc/lgvariables.inc" --> 
<!-- #Include file="../ESSinc/incGrid.inc" -->

Dim Grid1

Dim gDecimal_day
Dim gDecimal_time

'========================================================================================================
' Function Name : MakeKeyStream
'========================================================================================================
Sub MakeKeyStream(pOpt)
    lgKeyStream = "<%=lgEmp_no%>" & gColSep
    lgKeyStream = lgKeyStream & "<%=lgPay_Yymm%>" & gColSep
    lgKeyStream = lgKeyStream & "<%=lgProv_Type%>" & gColSep
End Sub 
       
'========================================================================================================
' Name : InitComboBox()
'========================================================================================================
Sub InitComboBox()
    Dim strTemp
    Dim lgstrData
    strTemp = "<%=lgProv_Type%>"

     Call CommonQueryRs(" MINOR_NM "," B_MINOR ", " MAJOR_CD = " & FilterVar("H0040", "''", "S") & " AND MINOR_CD =  " & FilterVar(strTemp , "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    if  Replace(lgF0, Chr(11), "") = "" then
        lgstrData = ""
    else
        lgstrData = Replace(lgF0, Chr(11), "")
    end if
	document.frm1.Prov_nm.value = lgstrData
End Sub

'========================================================================================================
' Name : LoadInfTB19029()	
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../comasp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029(gCurrency, "Q", "H") %>
End Sub

'========================================================================================================
' Function Name : GridDsplay
'========================================================================================================
Function GridDsplay()
	Dim i 
	For i = 0 to C_SHEETMAXROWS
		document.writeln "<tr>"
		document.writeln "<td width=180 height=30 bgcolor=D4E5E8 class=ctrow03>"
		If i = 0 Then
		document.writeln "연차기준금액</td>" 
		document.writeln "<td width=174 align=right bgcolor=E1EEF1 class=ctrow04>"
		document.writeln "<INPUT class=base2 NAME=YEAR_SUM SiZE=20 style='text-align:right;' readonly>&nbsp;&nbsp;</td>"
		ElseIf i = 1 Then  
		document.writeln "월차기준금액</td>" 
		document.writeln "<td width=174 align=right bgcolor=E1EEF1 class=ctrow04>"
		document.writeln "<INPUT class=base2 NAME=MONTH_SUM SiZE=20 style='text-align:right;' readonly>&nbsp;&nbsp;</td>"
		ElseIf i = 2 Then  
		document.writeln "연차 1개당 기준금액</td>" 
		document.writeln "<td width=174 align=right bgcolor=E1EEF1 class=ctrow04>"
		document.writeln "<INPUT class=base2 NAME=YEAR_PER SiZE=20 style='text-align:right;' readonly>&nbsp;&nbsp;</td>"
		ElseIf i = 3 Then  
		document.writeln "월차 1개당 기준금액</td>" 
		document.writeln "<td width=174 align=right bgcolor=E1EEF1 class=ctrow04>"
		document.writeln "<INPUT class=base2 NAME=MONTH_PER SiZE=20 style='text-align:right;' readonly>&nbsp;&nbsp;</td>"
		Else
		document.writeln "&nbsp;</td>" 
		document.writeln "<td width=174 align=right bgcolor=E1EEF1 class=ctrow04></td>"
		End If
		document.writeln "<td width=180 height=30 bgcolor=D4E5E8 class=ctrow03>"
		document.writeln "<INPUT class=base1 NAME=sub_nm" & i & " SiZE=30 style='text-align:left;' readonly></td>"
		document.writeln "<td width=174 align=right bgcolor=E1EEF1 class=ctrow04>"
		document.writeln "<INPUT class=base2 NAME=sub_amt" & i & " SiZE=20 style='text-align:right;' readonly>&nbsp;&nbsp;</td>"
		document.writeln "</tr>"
	Next 
End Function

'========================================================================================================
' Name : Form_Load
'========================================================================================================
Sub Form_Load()

    Err.Clear                                                                       '☜: Clear err status
    parent.document.All("nextprev").style.VISIBILITY = "hidden"
	Call SetToolBar("0000")    
    Call LoadInfTB19029()
    call get_decimal()
    Call InitComboBox()

    Call DbQuery()
End Sub

'========================================================================================
' Function Name : Form_UnLoad
'========================================================================================
Sub Form_UnLoad()
	On Error Resume Next
    Set Grid1 = Nothing
End Sub

'========================================================================================
' Function Name : DbQuery
'========================================================================================
Function DbQuery()
    Dim strVal
    Err.Clear                                                                    '☜: Clear err status

    DbQuery = False                                                              '☜: Processing is NG
	Call LayerShowHide(1)
    Call MakeKeyStream("Q")

    strVal = BIZ_PGM_ID & "?txtMode="            & "UID_M0001"                     '☜: Query
    strVal = strVal     & "&txtKeyStream="       & lgKeyStream                   '☜: Query Key

    Call RunMyBizASP(MyBizASP, strVal)                                           '☜:  Run biz logic
 
    DbQuery = True                                                               '☜: Processing is NG
 
End Function

'========================================================================================
' Function Name : fncGoBack
'========================================================================================
Function fncGoBack()
	dim strVal
	strVal = "e1202ma1.asp?strTitle=월급여&year=<%=left(lgPay_Yymm,4)%>"
	
	document.location = strVal
End Function

'========================================================================================
' Function Name : DbQueryOk
'========================================================================================
Function DbQueryOk()
    Err.Clear                                                                    '☜: Clear err status

End Function

'========================================================================================
' Function Name : DbQueryFail
'========================================================================================
Function DbQueryFail()
    Err.Clear
   
End Function

'========================================================================================================
'                        5.5 Tag Event
'========================================================================================================
Sub Query_OnClick()
    Call DbQuery(1)
End Sub

'======================================================================================================
' Function Name : FncBtnPrint
'=======================================================================================================%>
Function FncBtnPrint() 

	Dim strUrl
	Dim lngPos
	Dim intCnt
	Dim StrEbrFile
	dim prov_tp
	
	prov_tp = "<%=lgProv_Type%>"

	StrEbrFile = "h6016oa1_2.ebr"	

	strUrl = strUrl & "Pay_Yymm|<%=lgPay_Yymm%>"
	strUrl = strUrl & "|Prov_Type|<%=lgProv_Type%>"
	strUrl = strUrl & "|Pay_cd|" & "%"
	strUrl = strUrl & "|Fr_Dept_cd|1" 
	strUrl = strUrl & "|To_Dept_cd|zzzzzzzzz"
	strUrl = strUrl & "|Emp_no|<%=lgEmp_no%>"
	strUrl = strUrl & "|gDecimal_day|" & gDecimal_day
	strUrl = strUrl & "|gDecimal_time|" & gDecimal_time

	call FncEBRPrint(EBAction , StrEbrFile , strUrl)

End Function

'======================================================================================================
' Function Name : get_decimal
'=======================================================================================================%>
Sub get_decimal()
    Dim intRetCd
    
	gDecimal_day = 0
	gDecimal_time = 0

	IntRetCd = CommonQueryRs(" DECI_PLACE "," HDA041T "," ATTEND_TYPE = " & FilterVar("1", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	If IntRetCd = True Then
	    gDecimal_day  = Trim(Replace(lgF0,Chr(11),""))
	End If

	IntRetCd = CommonQueryRs(" DECI_PLACE "," HDA041T "," ATTEND_TYPE = " & FilterVar("2", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	If IntRetCd = True Then
	    gDecimal_time  = Trim(Replace(lgF0,Chr(11),""))
	End If

End Sub

</SCRIPT>
 <!-- #Include file="../ESSinc/uniSimsClassID.inc" --> 

</HEAD>

<BODY topmargin=0 leftmargin=0 marginheight=0 marginwidth=0>
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
	<tr> 
		<td height="10"></td>
	</tr>
    <TR>
        <TD height=34 bgcolor='#EAEBD2' align = center>
            <table border="0" cellspacing="0" cellpadding="0">
              <tr>
				  <td class="ftgray"><font color="5F564D"><strong>
				    <%=left(lgPay_Yymm,4)%>&nbsp;년 &nbsp;<%=right(lgPay_Yymm,2)%>&nbsp;월</strong></font></td>
				  <td width="120" align="center">
					<input name="prov_nm" class="form" type="text" style="width:120px; text-align: right; FONT-WEIGHT: bold;">&nbsp;</td>
				  <td class="ftgray"><font color="5F564D"><strong>지급명세서</strong></font></td>
			  </tr>
			</table>
		</TD>
	</TR>
    <tr> 
      <td height="10"></td>
    </tr>
    <tr> 
      <td> 
        <!--------------- List S --------------->
        <table width="100%" border="0" cellpadding="0" cellspacing="1" bgcolor="DDDDDD">
          <tr> 
            <td width="180" height="30" bgcolor="D4E5E8" class="ctrow01">성명</td>
            <td width="174" bgcolor="E1EEF1" class="ctrow02">
				<INPUT NAME="name"  class=base2 SiZE=15 readonly></td>
            <td width="180" height="30" bgcolor="D4E5E8" class="ctrow01">사번</td>
            <td width="174" bgcolor="E1EEF1" class="ctrow02">
				<INPUT NAME="emp_no" class=base2 SiZE=20 readonly></td>
          </tr>
          <tr> 
            <td height="30" bgcolor="D4E5E8" class="ctrow01">직급</td>
            <td bgcolor="E1EEF1" class="ctrow02">
				<INPUT class=base2 NAME="grade" SiZE=15 readonly></td>
            <td height="30" bgcolor="D4E5E8" class="ctrow01">부서</td>
            <td bgcolor="E1EEF1" class="ctrow02">
				<INPUT class=base2 NAME="dept_cd" SiZE=20 readonly></td>
          </tr>
        </table>
        <!--------------- List E--------------->
      </td>
    </tr>
    <tr> 
      <td height="10"></td>
    </tr>
<!--    
    <tr> 
      <td> 
        <table width="100%" border="0" cellpadding="0" cellspacing="1" bgcolor="DDDDDD">
          <tr> 
            <td width="180" height="30" bgcolor="D4E5E8" class="ctrow01">연 차 발 생</td>
            <td width="174" bgcolor="E1EEF1" class="ctrow02">
				<INPUT class=Base2 NAME="YEAR_LEFT" style='text-align:right;' SiZE=15 readonly>&nbsp;</td>
            <td width="180" height="30" bgcolor="D4E5E8" class="ctrow01">월 차 발 생</td>
            <td width="174" bgcolor="E1EEF1" class="ctrow02">
				<INPUT class=base2  NAME="MONTH_LEFT" style='text-align:right;' SiZE=15 ID="Text7" readonly>&nbsp;</td>
          </tr>
		  <tr> 
			<td height="30" bgcolor="D4E5E8" class="ctrow01">근 속 가 산</td>
			<td bgcolor="E1EEF1" class="ctrow02">
				<INPUT class=Base2 NAME="YEAR_PART" style='text-align:right;' SiZE=15 readonly ID="Text1">&nbsp;</TD></td>
			<td height="30" bgcolor="D4E5E8" class="ctrow01">&nbsp;</td>
			<td bgcolor="E1EEF1" class="ctrow02">&nbsp;</td>
		  </tr>
          <tr> 
            <td width="180" height="30" bgcolor="D4E5E8" class="ctrow01">연 차 사 용</td>
            <td width="174" bgcolor="E1EEF1" class="ctrow02">
				<INPUT class=Base2 NAME="YEAR_USE" style='text-align:right;' SiZE=15 readonly>&nbsp;</td>
            <td width="180" height="30" bgcolor="D4E5E8" class="ctrow01">월 차 사 용</td>
            <td width="174" bgcolor="E1EEF1" class="ctrow02">
				<INPUT class=base2  NAME="MONTH_USE" style='text-align:right;' SiZE=15 ID="Text7" readonly>&nbsp;</td>
          </tr>
          <tr> 
            <td width="180" height="30" bgcolor="D4E5E8" class="ctrow01">연 차 지 급</td>
            <td width="174" bgcolor="E1EEF1" class="ctrow02">
				<INPUT class=Base2 NAME="YEAR_CNT" style='text-align:right;' SiZE=15 readonly>&nbsp;</td>
            <td width="180" height="30" bgcolor="D4E5E8" class="ctrow01">월 차 지 급</td>
            <td width="174" bgcolor="E1EEF1" class="ctrow02">
				<INPUT class=base2  NAME="MONTH_CNT" style='text-align:right;' SiZE=15 ID="Text7" readonly>&nbsp;</td>
          </tr>
        </table>
      </td>
    </tr>
    <tr> 
      <td height="10"></td>
    </tr>
-->    
    <tr> 
      <td> 
        <table width="100%" border="0" cellpadding="0" cellspacing="1" bgcolor="DDDDDD">
		  <tr> 
			<td height="30" colspan="2" bgcolor="D4E5E8" class="listitle01">기준금내역</td>
			<td height="30" colspan="2" bgcolor="D4E5E8" class="listitle01">공제내역</td>
		  </tr>
 			<script language=vbscript>    Call GridDsplay()  </script>
        </table>
      </td>
    </tr>
    <tr> 
      <td height="10"></td>
    </tr>
    <tr> 
      <td>
		<table width="100%" border="0" cellpadding="0" cellspacing="1" bgcolor="DDDDDD">
          <tr> 
            <td height="30" bgcolor="D4E5E8" class="ctrow03">지급총액</td>
            <td align="right" bgcolor="E1EEF1" class="ctrow02">
				<INPUT class=base2 NAME="pay_tot" SiZE=20 style='text-align:right;'  readonly>&nbsp;</td>
            <td width="180" height="30" bgcolor="D4E5E8" class="ctrow01">공제총액</td>
            <td width="174" align="right" bgcolor="E1EEF1" class="ctrow02">
				<INPUT NAME="sub_tot" class=base2 SiZE=20 style='text-align:right;'  readonly>&nbsp;</td>
          </tr>
          <tr> 
            <td height="30" bgcolor="D4E5E8" class="ctrow03" colspan=2>실지급액</td>
            <td align="right" bgcolor="E1EEF1" class="ctrow02" colspan=2 align=right>
				<INPUT class=base2 NAME="real" SiZE=30 style='text-align:right;'  readonly>&nbsp;</td>
          </tr>
        </table>
       </td>
    </tr>
    <tr> 
        <td height="10"></td>
    </tr>
    <TR align = center>
    <td>
		<IMG SRC="../ESSimage/button_04.gif" border="0" OnClick="vbscript: call FncBtnPrint()" name="printprev" alt='출력' onMouseOver="javascript:this.src='../ESSimage/button_r_04.gif';" onMouseOut="javascript:this.src='../ESSimage/button_04.gif';">
		<IMG SRC="../ESSimage/button_15.gif" border="0" OnClick="vbscript: call fncGoBack()" name="printprev" alt='돌아가기' onMouseOver="javascript:this.src='../ESSimage/button_r_15.gif';" onMouseOut="javascript:this.src='../ESSimage/button_15.gif';">
	</td>
	</tr>
</table>

<TABLE width="100%" height="0" border="0" cellpadding="0" cellspacing="0">
	<TR><TD WIDTH="100%" HEIGHT=0>
		<IFRAME NAME="MyBizASP" WIDTH="100%" HEIGHT=0 FRAMEBORDER=0 SCROLLING=no noresize framespacing=0></IFRAME></TD>
	</TR>
</TABLE>

    <INPUT TYPE=hidden NAME="txtMode">
    <INPUT TYPE=hidden NAME="txtKeyStream">
    <INPUT TYPE=hidden NAME="txtUpdtUserId">
    <INPUT TYPE=hidden NAME="txtInsrtUserId">
    <INPUT TYPE=hidden NAME="txtFlgMode">
    <INPUT TYPE=hidden NAME="txtPrevNext">
    
    <INPUT TYPE=hidden NAME=GRID_TOTPAGES >
    <INPUT TYPE=hidden NAME=GRID_PAGE value=1 >
</FORM>	

<FORM NAME="EBAction" TARGET = "MyBizASP" METHOD="POST">
	<INPUT TYPE="HIDDEN" NAME="uname">
	<INPUT TYPE="HIDDEN" NAME="dbname">
	<INPUT TYPE="HIDDEN" NAME="filename">
	<INPUT TYPE="HIDDEN" NAME="condvar">
	<INPUT TYPE="HIDDEN" NAME="date">
</FORM></BODY>
</HTML>
