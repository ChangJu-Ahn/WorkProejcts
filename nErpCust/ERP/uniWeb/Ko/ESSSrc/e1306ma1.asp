<%@ LANGUAGE="VBSCRIPT" %>
<% Response.Expires = -1%>

<HTML>
<HEAD>
<TITLE><%=Request("strTitle")%></TITLE>

<!-- #Include file="../ESSinc/incServer.asp"  -->

<LINK REL="stylesheet" TYPE="Text/css" href="../ESSinc/common.css">

<SCRIPT LANGUAGE="VBScript"   SRC="../ESSinc/ccm.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../ESSinc/variables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../ESSinc/incCookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../ESSinc/operation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../ESSinc/incCommFunc.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../ESSinc/incEvent.vbs">   </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../ESSinc/adoQuery.vbs"></SCRIPT>

<!-- #Include file="../ESSinc/incSvrVarSims.inc"  -->
<!-- #Include file="../ESSinc/incSvrFuncSims.inc" -->
<Script Language="VBScript">
Option Explicit  

Const BIZ_PGM_ID      = "e1306mb1.asp"						           '☆: Biz Logic ASP Name
Const BIZ_PGM_ID1     = "e1306ma2.asp"						           '☆: Biz Logic ASP Name
Const C_SHEETMAXCOLS = 13
Const C_SHEETMAXROWS = 8

<!-- #Include file="../ESSinc/lgvariables.inc" --> 
<!-- #Include file="../ESSinc/incGrid.inc" -->

Dim Grid1
dim fDiligAuth,fAuthCheck

<% EndDate   = GetSvrDate %>

'========================================================================================================
' Function Name : MakeKeyStream
'========================================================================================================
Sub MakeKeyStream(pOpt)
    if  pOpt = "Q" then
		if  Trim(parent.txtEmp_no2.Value) = "" then
		    lgKeyStream = Trim(parent.txtEmp_no.Value) & gColSep
		else
		    lgKeyStream = Trim(parent.txtEmp_no2.Value) & gColSep
		end if
    else     
            lgKeyStream = Trim(frm1.txtEmp_no.Value) & gColSep   
    end if
        lgKeyStream = lgKeyStream & Trim(parent.txtEmp_no.Value) & gColSep
        lgKeyStream = lgKeyStream & Trim(Replace(frm1.txtYear.Value,"-","")) & gColSep             
        lgKeyStream = lgKeyStream & Trim(fDiligAuth) & gColSep        
        lgKeyStream = lgKeyStream & Trim(fAuthCheck) & gColSep
End Sub   
     
'========================================================================================================
' Name : InitComboBox()
'========================================================================================================
Sub InitComboBox()
    Dim iCodeArr 
    Dim iNameArr
    Dim iDx
    dim strSQL, IntRetCD
    
	iCodeArr = ""
	iNameArr = ""

    strSQL = " org_cd = " & FilterVar("1", "''", "S") & " AND pay_gubun = " & FilterVar("Z", "''", "S") & " AND PAY_TYPE = " & FilterVar("*", "''", "S") & " "
    IntRetCD = CommonQueryRs(" year(close_dt) close_year "," hda270t ", strSQL,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    If  IntRetCd = true then
		iDx = Replace(lgF0, Chr(11), "") +1
	end if
	iCodeArr = cdbl(idx) & Chr(11) & iCodeArr
	iNameArr = cdbl(idx) & Chr(11) & iNameArr
	   
    Call SetCombo2(frm1.txtYear, iCodeArr, iNameArr, Chr(11))
End Sub

'========================================================================================================
' Function Name : InitSpreadSheet
'========================================================================================================
Sub InitGrid()
    Set Grid1 = New Grid
    Grid1.MaxCols = C_SHEETMAXCOLS 
    Grid1.SheetMaxrows = C_SHEETMAXROWS
    Set Grid1.Source = document.frm1
End Sub

'========================================================================================================
' Name : Form_Load
'========================================================================================================
Sub Form_Load()
    On Error Resume Next
    
    Err.Clear                                                                       '☜: Clear err status
    'frm1.txtUID.focus() 
    call FncGetDiligAuth(fDiligAuth,fAuthCheck)
    If Replace(fDiligAuth,Chr(11),"") = "" Then
        parent.document.All("nextprev").style.VISIBILITY = "hidden"
    Else
        parent.document.All("nextprev").style.VISIBILITY = "visible"
    End If
  
    Call InitComboBox()
    Call LayerShowHide(0)
    Call InitGrid()

    Call SetToolBar("11000")
	if parent.txtName2.value ="" then
		parent.txtEmp_no2.Value = parent.txtemp_no.value 
	end if

    Call LockField(Document)
    Call DbQuery(1)
   
End Sub
'========================================================================================
' Function Name : Form_unLoad
'========================================================================================
Sub Form_unLoad()
	On Error Resume Next
    Set Grid1 = Nothing
End Sub

'========================================================================================
' Function Name : DbQuery
'========================================================================================
Function DbQuery(ppage)

    Dim strDate
    Dim strVal
    Err.Clear                                                                    '☜: Clear err status

    If frm1.txtYear.value = "" then
		Call DisplayMsgBox("800094","X","X","X")
		Exit Function
    End if
    
    If len(frm1.txtYear.value)<>4 then
		Call DisplayMsgBox("800094","X","X","X")
		Exit Function
    End if 
    
    If Not chkFieldLength(Document, "1") Then									         '☜: This function check required field
		Exit Function
	end if
    DbQuery = False                                                              '☜: Processing is NG
    'If Grid1.ChkChange() Then Exit Function
    Call ClearField(document,2)
    Call LayerShowHide(1)
    Call MakeKeyStream("Q")

    strVal = BIZ_PGM_ID & "?txtMode="            & "UID_M0001"                     '☜: Query
    strVal = strVal     & "&txtKeyStream="       & lgKeyStream                   '☜: Query Key

    Call RunMyBizASP(MyBizASP, strVal)                                           '☜:  Run biz logic

    DbQuery = True                                                               '☜: Processing is NG
End Function

'========================================================================================
' Function Name : DbQueryOk
'========================================================================================
Function DbQueryOk()
    Err.Clear                                                                    '☜: Clear err status

    Call Grid1.ShowData(frm1,frm1.GRID_PAGE.VALUE)
End Function

'========================================================================================
' Function Name : DbQueryFail
'========================================================================================
Function DbQueryFail()
    Err.Clear
    Call ClearField(Document,2)                                                                    '☜: Clear err status
End Function

'========================================================================================================
' Name : FncNew
'========================================================================================================
Function FncNew()
    Dim IntRetCD 
    Dim strVal
   
    FncNew = False																 '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

    strVal = BIZ_PGM_ID1 & "?txtMode="            & "UID_M0001"                     '☜: Query
    strVal = strVal & "&txtMainInsertFlag=Y"
  
	Call RunMyBizASP(document, strVal)
    FncNew = True																 '☜: Processing is OK
    
End Function

'========================================================================================================
' Name : GetRow
'========================================================================================================
Function GetRow(pRow)
	GetRow=False
    Grid1.ActiveRow = pRow
    If Mid(document.activeElement.getAttribute("tag"),3,1) = "1" Then
	    arrRet = window.showModalDialog("../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
	    	"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	End If
	GetRow=True
End Function

'========================================================================================================
' Name : DoubleGetRow
'========================================================================================================
Function DoubleGetRow(pRow)
    On Error Resume Next
    Err.Clear
    Dim objList
    Dim elmCnt
	Dim strVal
    Dim txtMedDate, txtMed_nm, txtFamilyResNo,txtFamilyNm,txtflag
    Dim txtMed_Resno

	DoubleGetRow = False
	Grid1.ActiveRow = pRow

    txtFamilyName = ""
    txtRel_cd = ""
    txtType = ""
    txtAmt = ""
         
    with frm1
    	For elmCnt = 0 to .length - 1
    		Set objList = .elements(elmCnt)
    		If objList.name = "SPREADCELL_DATE" & pRow then
               txtMedDate = objList.value
    		End if

    		If objList.name = "SPREADCELL_MED_NM" & pRow then
               txtMed_nm = objList.value
    		End if
    		
			If objList.name = "SPREADCELL_MED_RGST_NO" & pRow then
               txtMed_Resno = objList.value
    		End if

    		If objList.name = "SPREADCELL_FAMILY_RES_NO" & pRow then
               txtFamilyResNo = objList.value
    		End if  

    		If objList.name = "SPREADCELL_FAMILY_NM" & pRow then
               txtFamilyNm = objList.value
    		End if   
    		  
    		If objList.name = "SPREADCELL_FLAG" & pRow then
               txtflag = objList.value
    		End if       		
    	Next

    End With

    If frm1.txtYear.value <> "" and frm1.txtEmp_no.value <> "" and txtMedDate<> ""  then

        strVal = BIZ_PGM_ID1 & "?txtYear=" & frm1.txtYear.value
        
        strVal = strVal & "&txtEmp_no=" & frm1.txtEmp_no.value
        strVal = strVal & "&txtMed_date=" & txtMedDate
        strVal = strVal & "&txtMed_Resno=" & txtMed_Resno
        strVal = strVal & "&txtFamily_res_no=" & txtFamilyResNo
        strVal = strVal & "&txtFamily_name=" & txtFamilyNm
        strVal = strVal & "&txtflag=" & txtflag 
        

        document.location = strVal
    end if

	DoubleGetRow = True
End Function

'========================================================================================================
' Name : MouseRow
'========================================================================================================
Sub MouseRow(pRow)
	If frm1.grid_totpages.value ="" Then Exit Sub
    Dim objList   

	Grid1.ActiveRow = pRow	
	Set objList = window.event.srcElement	
	
	If  Ucase(objList.getAttribute("flag")) = "SPREADCELL" then
        if objList.value = "" then            
             objList.style.cursor = "auto"
        else
             objList.style.cursor = "hand"
        end if
    End If        

End Sub

'========================================================================================================
'                        5.5 Tag Event
'========================================================================================================
Sub Query_OnClick()
    Call DbQuery(1)
End Sub

Sub GRID_PAGE_OnChange()
End Sub

Function txtEmp_no2_Onchange()
        Call DbQuery(1)	
End Function

'========================================================================================================
' Name : FncGetDiligAuth()
'========================================================================================================
Function FncGetDiligAuth(fDiligAuth,fAuthCheck)
    fDiligAuth = ""
    fAuthCheck = ""
    Call CommonQueryRs(" internal_cd,internal_auth "," e11090t "," emp_no = '" & Trim(parent.txtEmp_no.Value) & "'",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    fDiligAuth = replace(lgF0,chr(11),chr(12))
    fDiligAuth = replace(fDiligAuth," ","")    
    fAuthCheck = replace(lgF1,chr(11),chr(12))
    fAuthCheck = replace(fAuthCheck," ","")      
End Function

</SCRIPT>
<!-- #Include file="../ESSinc/uniSimsClassID.inc" --> 

</HEAD>
<BODY topmargin=0 leftmargin=0 marginheight=0 marginwidth=0>
<FORM NAME="frm1" TARGET="MyBizASP" METHOD="post">
    <TABLE cellSpacing=0 cellPadding=0 border=0>
        <TR>
            <TD valign="top">
                <TABLE width="100%" cellSpacing=0 cellPadding=0 border=0>
                    <TR>
                       <td height="10"></td>
                    </TR>
                    <TR>
                        <td><table width="732" border="0" cellspacing="1" cellpadding="0" bgcolor="DDDDDD">
                            <tr> 
								<td width="80" height="27" bgcolor="D4E5E8" class="base1">사번</td>
								<td width="85" bgcolor="FFFFFF"><INPUT class=base2 NAME="txtEmp_no" MAXLENGTH=13 SiZE=12 readonly></td>
								<td width="60" height="27" bgcolor="D4E5E8" class="base1">성명</td>
								<td width="80" bgcolor="FFFFFF"><INPUT class=base2 NAME="txtName" MAXLENGTH=20 SiZE=16  readonly></td>
								<td width="60" height="27" bgcolor="D4E5E8" class="base1">직위</td>
								<td width="80" bgcolor="FFFFFF"><INPUT class=base2 NAME="txtroll_pstn" MAXLENGTH=20 SiZE=16  readonly></td>
								<td width="60" height="27" bgcolor="D4E5E8" class="base1">부서</td>
								<td width="153" bgcolor="FFFFFF"><INPUT class=base2 NAME="txtDept_nm" MAXLENGTH=25 SiZE=22  readonly></td>
                            </tr>
                            <tr> 
								<td width="80" height="30" bgcolor="D4E5E8" class=base1 valign=middle>정산년도
								</td>
								<td width="85" bgcolor="FFFFF" align=center>
								    <SELECT Name="txtYear" tabindex=-1 class=base2>
								    </SELECT>
								</td>
								<td bgcolor="FFFFFF" class="base1">&nbsp;</td>
								<td bgcolor="FFFFFF" class="base2">&nbsp;</td>
								<td bgcolor="FFFFFF" class="base1">&nbsp;</td>
								<td bgcolor="FFFFFF" class="base2">&nbsp;</td>
								<td bgcolor="FFFFFF" class="base1">&nbsp;</td>
								<td bgcolor="FFFFFF" class="base2">&nbsp;</td>
                            </tr>
                            </table>
                        </td>
                    </TR>
                    <TR>
                       <td height="10"></td>
                    </TR>
                    <TR>
                        <td><table width="100%" border="0" cellspacing="1" cellpadding="0" bgcolor="DDDDDD">
								<TR> 
								    <TD class=TDFAMILY_TITLE1></TD>
	                        		<TD CLASS=TDFAMILY_TITLE1>지급일자</TD>
	                        		<TD CLASS=TDFAMILY_TITLE1>지급처상호</TD>
	                        		<TD CLASS=TDFAMILY_TITLE1>지급처사업자번호</TD>
	                        		<TD CLASS=TDFAMILY_TITLE1>가족성명</TD>
	                        		<TD CLASS=TDFAMILY_TITLE1>가족관계</TD>	   
	                        		<TD CLASS=TDFAMILY_TITLE1>주민번호</TD>
	                        		<TD CLASS=TDFAMILY_TITLE1>대상자구분</TD>	
	                        		<TD CLASS=TDFAMILY_TITLE1>지급금액</TD>
	                        		<TD CLASS=TDFAMILY_TITLE1>의료비내용</TD>	                        			                        			                        		                      		
                                </TR>
							    <% 
                                For i=1 To 8
                                    Response.Write "<TR bgcolor=#F8F8F8 height=24 onclick='vbscript: Call DoubleGetRow(" & i & ")' onMouseOver=" & chr(34) & "javascript: this.style.backgroundColor='E1EEF1'" & chr(34) & " onMouseOut=" & chr(34) & "javascript: this.style.backgroundColor=''" & chr(34) & ">"
                                    Response.Write "<TD><INPUT name='" & i & "'  class=listrow01 flag='SPREADCELL' style='WIDTH:  30px; text-align: center;' onMouseOver='vbscript: Call MouseRow(" & i & ")' readonly ></TD>"
                                    Response.Write "<TD><INPUT name='SPREADCELL_DATE" & i & "'			class=listrow01 flag='SPREADCELL' style='WIDTH: 80px; text-align: left;' onMouseOver='vbscript: Call MouseRow(" & i & ")' readonly></TD>"
                                	Response.Write "<TD><INPUT name='SPREADCELL_MED_NM" & i & "'		class=listrow01 flag='SPREADCELL' style='WIDTH: 120px; text-align: left;' onMouseOver='vbscript: Call MouseRow(" & i & ")' readonly></TD>"
                                    Response.Write "<TD><INPUT name='SPREADCELL_MED_RGST_NO" & i & "'	class=listrow01 flag='SPREADCELL' style='WIDTH: 110px; text-align: left;' onMouseOver='vbscript: Call MouseRow(" & i & ")' readonly></TD>"
                                    Response.Write "<TD><INPUT name='SPREADCELL_FAMILY_NM" & i & "'		class=listrow01 flag='SPREADCELL' style='WIDTH: 80px; text-align: left;' onMouseOver='vbscript: Call MouseRow(" & i & ")' readonly></TD>"
                                    Response.Write "<TD><INPUT name='SPREADCELL_FAMILY_REL_CD" & i & "' class=listrow01  type=hidden flag='SPREADCELL' style='WIDTH:   0px; text-align: center;'>"
                                	Response.Write "    <INPUT name='SPREADCELL_FAMILY_REL" & i & "'	class=listrow01 flag='SPREADCELL' style='WIDTH: 110px; text-align: left;'  onMouseOver='vbscript: Call MouseRow(" & i & ")' readonly></TD>"
                                    Response.Write "<TD><INPUT name='SPREADCELL_FAMILY_RES_NO" & i & "' class=listrow01 flag='SPREADCELL' style='WIDTH: 100px; text-align: left;' onMouseOver='vbscript: Call MouseRow(" & i & ")' readonly></TD>"
                                    Response.Write "<TD><INPUT name='SPREADCELL_FAMILY_TYPE" & i & "' class=listrow01  type=hidden flag='SPREADCELL' style='WIDTH:   0px; text-align: center;'>"
                                	Response.Write "    <INPUT name='SPREADCELL_FAMILY_TYPE_NM" & i & "'	class=listrow01 flag='SPREADCELL' style='WIDTH: 80px; text-align: left;'  onMouseOver='vbscript: Call MouseRow(" & i & ")' readonly></TD>"
 
                                    Response.Write "<TD><INPUT name='SPREADCELL_MED_AMT" & i & "' class=listrow01 flag='SPREADCELL' style='WIDTH: 80px; text-align: right;' onMouseOver='vbscript: Call MouseRow(" & i & ")' readonly></TD>"
                                    Response.Write "<TD><INPUT name='SPREADCELL_MED_TEXT" & i & "' class=listrow01 flag='SPREADCELL' style='WIDTH: 100px; text-align: left;' onMouseOver='vbscript: Call MouseRow(" & i & ")' readonly></TD>"
                                    Response.Write "<INPUT class=listrow01 name='SPREADCELL_FLAG" & i & "'	type=hidden flag='SPREADCELL' style='WIDTH: 40px; text-align: right;' onMouseOver='vbscript: Call MouseRow(" & i & ")' readonly>"
                                    Response.Write "</TR>"
                                Next
								%>
                           </table>
                        </td>
                    </TR>
                </TABLE>
            </TD>
        </TR>
        <TR>
            <TD height=5></TD>
        </TR>
        <TR height=20>
            <TD VALIGN=center ALIGN=center>
                <A onclick="VBSCRIPT:CALL GRID1.PREPAGES()" onMouseOver="javascript: this.style.cursor='hand'"><IMG alt="이전페이지" SRC=../ESSimage/button_07.gif border=0 ></A>&nbsp;
                <A onclick="VBSCRIPT: CALL GRID1.NEXTPAGES()" onMouseOver="javascript: this.style.cursor='hand'"><IMG alt="다음페이지" SRC=../ESSimage/button_08.gif border=0 ></A>&nbsp;&nbsp;
            </TD>
        </TR>
    </TABLE>
    <TABLE cellSpacing=0 cellPadding=0 border=0>
        <TR><TD WIDTH="100%" HEIGHT=0><IFRAME NAME="MyBizASP" WIDTH="100%" HEIGHT=0 FRAMEBORDER=0 SCROLLING=auto noresize framespacing=0></IFRAME></TD></TR>
    </TABLE>

    <INPUT TYPE=hidden NAME="txtMode">
    <INPUT TYPE=hidden NAME="txtKeyStream">
    <INPUT TYPE=hidden NAME="txtUpdtUserId">
    <INPUT TYPE=hidden NAME="txtInsrtUserId">
    <INPUT TYPE=hidden NAME="txtFlgMode">
    <INPUT TYPE=hidden NAME="txtPrevNext">
    <INPUT TYPE=hidden NAME=GRID_TOTPAGES>
    <INPUT TYPE=hidden NAME=txtMed_Resno value="www">
    
    <INPUT TYPE=hidden NAME=GRID_PAGE value=1>
 </FORM>	
</BODY>
</HTML>
