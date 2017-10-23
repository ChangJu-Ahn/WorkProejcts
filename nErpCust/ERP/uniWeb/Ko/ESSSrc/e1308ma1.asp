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

Const BIZ_PGM_ID      = "e1308mb1.asp"						           '☆: Biz Logic ASP Name
Const BIZ_PGM_ID1     = "e1308ma2.asp"						           '☆: Biz Logic ASP Name
Const C_SHEETMAXCOLS = 12
Const C_SHEETMAXROWS = 7
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
    Err.Clear
    Dim objList
    Dim elmCnt
	Dim strVal
    Dim txtFamilyName, txtFamilyRes_no
   
	DoubleGetRow = False
	Grid1.ActiveRow = pRow

    txtFamilyName = ""

    with frm1
   
    	For elmCnt = 0 to .length - 1
    		Set objList = .elements(elmCnt)
    		If objList.name = "SPREADCELL_NAME" & pRow then
               txtFamilyName = objList.value
    		End if

    		If objList.name = "SPREADCELL_RES_NO" & pRow then
               txtFamilyRes_no = objList.value
    		End if    		
    	Next

    End With

    If frm1.txtYear.value <> "" and frm1.txtEmp_no.value <> "" and txtFamilyName<> ""  then

        strVal = BIZ_PGM_ID1 & "?txtYear=" & frm1.txtYear.value
        
        strVal = strVal & "&txtEmp_no=" & frm1.txtEmp_no.value
        strVal = strVal & "&txtFamilyName=" & txtFamilyName 
        strVal = strVal & "&txtFamilyRes_no=" & txtFamilyRes_no 
 
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
'========================================================================================
' Function Name : Form_UnLoad
' Function Desc : 페이지 전환이나 화면이 닫힐 경우 실행해야 될 로직 처리 
'========================================================================================
Function dodata(strdo)
	dim IntRetCD, strEmpNo
	dim strVal 
	
	strEmpNo = frm1.txtEmp_no.value 
	Call CommonQueryRs(" COUNT(*) "," HAA020T  "," emp_no = " & FilterVar(strEmpNo, "''", "S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
	If Trim(Replace(lgF0,Chr(11),"")) = 0 then
        Call DisplayMsgBox("800065","X","X","X")	'자동 입력할 사원이 없습니다.
	    Call BtnDisabled(0)
	    Exit Function                                    '바로 return한다....자동입력을 멈춘다.
    End if
    
    Call CommonQueryRs(" COUNT(*) "," HFA150T "," emp_no = " & FilterVar(strEmpNo, "''", "S")  & " and year_yy = " & FilterVar(frm1.txtYear.value, "''", "S") ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

	If Trim(Replace(lgF0,Chr(11),"")) > 0 then
        intRetCD = DisplayMsgBox("800502",parent.VB_YES_NO,"X","X")	'이미생성된 자료가 있습니다.
        if intRetCD = vbNO then
			Call BtnDisabled(0)
			Exit Function                                    '바로 return한다....자동입력을 멈춘다.
		end if
    End if
    
    Call MakeKeyStream("Q")
    strVal = BIZ_PGM_ID & "?txtMode="            & "UID_M0003"                     '☜: Query
    strVal = strVal     & "&txtKeyStream="       & lgKeyStream                   '☜: Query Key
    
    Call RunMyBizASP(MyBizASP, strVal)                                           '☜:  Run biz logic
 		
 
End Function
</SCRIPT>
<!-- #Include file="../ESSinc/uniSimsClassID.inc" --> 

</HEAD>
<BODY topmargin=0 leftmargin=0 marginheight=0 marginwidth=0>
<FORM NAME="frm1" TARGET="MyBizASP" METHOD="post">
    <TABLE width=732 cellSpacing=0 cellPadding=0 border=0>
        <TR>
            <TD valign="top">
                <TABLE width="100%" cellSpacing=0 cellPadding=0 border=0>
                    <TR>
                       <td height="10"></td>
                    </TR>
                    <TR>
                        <td><table width="100%" border="0" cellspacing="1" cellpadding="0" bgcolor="DDDDDD">
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
	                        		<TD CLASS=TDFAMILY_TITLE1>가족성명</TD>
	                        		<TD CLASS=TDFAMILY_TITLE1>가족관계</TD>
	                        		<TD CLASS=TDFAMILY_TITLE1>주민번호</TD>
	                        		<TD CLASS=TDFAMILY_TITLE1>기본공제</TD>
	                        		<TD CLASS=TDFAMILY_TITLE1>장애인</TD>
	                        		<TD CLASS=TDFAMILY_TITLE1>자녀양육비</TD>
	                        		<TD CLASS=TDFAMILY_TITLE1>보험료</TD>
	                        		<TD CLASS=TDFAMILY_TITLE1>의료비</TD>
	                        		<TD CLASS=TDFAMILY_TITLE1>교육비</TD>
	                        		<TD CLASS=TDFAMILY_TITLE1>신용카드등</TD>
                                </TR>
							    <% 
                                For i=1 To 7
                                    Response.Write "<TR bgcolor=#F8F8F8 height=24 onclick='vbscript: Call DoubleGetRow(" & i & ")' onMouseOver=" & chr(34) & "javascript: this.style.backgroundColor='E1EEF1'" & chr(34) & " onMouseOut=" & chr(34) & "javascript: this.style.backgroundColor=''" & chr(34) & ">"
                                    Response.Write "<TD><INPUT class=listrow01 name='" & i & "'  flag='SPREADCELL' style='WIDTH:  30px; text-align: center;' onMouseOver='vbscript: Call MouseRow(" & i & ")' readonly ></TD>"
                                    Response.Write "<TD><INPUT class=listrow01 name='SPREADCELL_NAME" & i & "' flag='SPREADCELL' style='WIDTH: 80px; text-align: left;' onMouseOver='vbscript: Call MouseRow(" & i & ")' readonly></TD>"
                                    Response.Write "<TD><INPUT class=listrow01 name='SPREADCELL_REL_CD" & i & "' type=hidden flag='SPREADCELL' style='WIDTH:   0px; text-align: center;'>"
                                	Response.Write "    <INPUT class=listrow01 name='SPREADCELL_REL" & i & "' flag='SPREADCELL' style='WIDTH: 90px; text-align: left;'  onMouseOver='vbscript: Call MouseRow(" & i & ")'></TD>"
									Response.Write "<TD><INPUT class=listrow01 name='SPREADCELL_RES_NO" & i & "' flag='SPREADCELL' style='WIDTH: 90px; text-align: left;' onMouseOver='vbscript: Call MouseRow(" & i & ")' readonly></TD>"                                	
                                    Response.Write "<TD><INPUT class=listrow01 name='SPREADCELL_BASE_YN" & i & "' flag='SPREADCELL' style='WIDTH: 60px; text-align: left;' onMouseOver='vbscript: Call MouseRow(" & i & ")' readonly></TD>"
                                    Response.Write "<TD><INPUT class=listrow01 name='SPREADCELL_PARIA_YN" & i & "' flag='SPREADCELL' style='WIDTH: 60px; text-align: left;' onMouseOver='vbscript: Call MouseRow(" & i & ")' readonly></TD>"
                                    Response.Write "<TD><INPUT class=listrow01 name='SPREADCELL_CHILD_YN" & i & "' flag='SPREADCELL' style='WIDTH: 60px; text-align: left;' onMouseOver='vbscript: Call MouseRow(" & i & ")' readonly></TD>"
                                    Response.Write "<TD><INPUT class=listrow01 name='SPREADCELL_INSUR_YN" & i & "' flag='SPREADCELL' style='WIDTH: 60px; text-align: left;' onMouseOver='vbscript: Call MouseRow(" & i & ")' readonly></TD>"
                                    Response.Write "<TD><INPUT class=listrow01 name='SPREADCELL_MEDI_YN" & i & "' flag='SPREADCELL' style='WIDTH: 60px; text-align: left;' onMouseOver='vbscript: Call MouseRow(" & i & ")' readonly></TD>"
                                    Response.Write "<TD><INPUT class=listrow01 name='SPREADCELL_EDU_YN" & i & "' flag='SPREADCELL' style='WIDTH: 60px; text-align: left;' onMouseOver='vbscript: Call MouseRow(" & i & ")' readonly></TD>"
                                    Response.Write "<TD><INPUT class=listrow01 name='SPREADCELL_CARD_YN" & i & "' flag='SPREADCELL' style='WIDTH: 60px; text-align: left;' onMouseOver='vbscript: Call MouseRow(" & i & ")' readonly></TD>"
                                   
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
                <A onclick="VBSCRIPT:CALL GRID1.PREPAGES()" onMouseOver="javascript: this.style.cursor='hand'"><IMG alt="이전페이지" src=../ESSimage/button_07.gif border=0 ></A>&nbsp;
                <A onclick="VBSCRIPT: CALL GRID1.NEXTPAGES()" onMouseOver="javascript: this.style.cursor='hand'"><IMG alt="다음페이지" src=../ESSimage/button_08.gif border=0 ></A>&nbsp;&nbsp;
            </TD>
        </TR>
		<TR valign=middle height=50>
			<TD class=base2>작년정보가 존재할때는 작년정보를,작년정보가 없고 가족사항이 있으면 가족사항에서 정보를 가져옵니다.
				<img src="../ESSimage/button_16.gif" onclick = "javascript:dodata('1')" alt="기초자료생성"  onMouseOver="javascript:this.src='../ESSimage/button_r_16.gif';this.style.cursor='hand';" onMouseOut="javascript:this.src='../ESSimage/button_16.gif';"></a>
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
    <INPUT TYPE=hidden NAME=GRID_PAGE value=1>
 </FORM>	
</BODY>
</HTML>
