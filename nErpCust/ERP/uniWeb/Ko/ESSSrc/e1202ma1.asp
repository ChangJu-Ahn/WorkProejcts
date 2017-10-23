<%@ LANGUAGE="VBSCRIPT" %>
<% Response.Expires = -1%>
<HTML>
<HEAD>
<TITLE><%=Request("strTitle")%></TITLE>

<!-- #Include file="../ESSinc/IncServer.asp"  -->

<LINK REL="stylesheet" TYPE="Text/css" href="../ESSinc/Common.css">

<SCRIPT LANGUAGE="VBScript"   SRC="../ESSinc/ccm.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../ESSinc/variables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../ESSinc/incCookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../ESSinc/operation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../ESSinc/incCommFunc.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../ESSinc/incEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../ESSinc/adoQuery.vbs"></SCRIPT>
<!-- #Include file="../ESSinc/incSvrVarSims.inc"  -->
<!-- #Include file="../ESSinc/incSvrFuncSims.inc" -->
<Script Language="VBScript">
Option Explicit                                                   '☜: indicates that All variables must be declared in advance

Const BIZ_PGM_ID      = "e1202mb1.asp"						      '☆: Biz Logic ASP Name
'Const C_SHEETMAXROWS = 7	                                      '☜: Visble row
Const C_SHEETMAXROWS = 5	                                      '☜: Visble row / 2008.02.14
Const C_SHEETMAXCOLS = 10

<!-- #Include file="../ESSinc/lgvariables.inc" --> 
<!-- #Include file="../ESSinc/incGrid.inc" -->

Dim Grid1

'========================================================================================================
' Function Name : MakeKeyStream
'========================================================================================================
Sub MakeKeyStream(pOpt)
   
    if  pOpt = "Q" then
'        if  Trim(parent.txtEmp_no2.Value) = "" then
            lgKeyStream = Trim(parent.txtEmp_no.Value) & gColSep
'        else
'            lgKeyStream = Trim(parent.txtEmp_no2.Value) & gColSep
'        end if
        lgKeyStream = lgKeyStream & Trim(parent.txtinternal_cd.Value) & gColSep
        lgKeyStream = lgKeyStream & Trim(frm1.txtYear.Value) & gColSep
    else
        lgKeyStream = Trim(frm1.txtEmp_no.Value) & gColSep
        lgKeyStream = lgKeyStream & Trim(parent.txtinternal_cd.Value) & gColSep
        lgKeyStream = lgKeyStream & Trim(frm1.txtYear.Value) & gColSep
    end if
End Sub   
     
'========================================================================================================
' Name : InitComboBox()
'========================================================================================================
Sub InitComboBox()
    Dim iCodeArr 
    Dim iNameArr
    Dim iDx
    Dim lgYear,i,stYear
    
    If Err.number = 0 Then 	lgYear = Year(date)
    lgYear = "<%=request("year")%>"
    
    If lgyear = "" Then lgyear = Year(date)
	if Trim(parent.txtemp_no.value)="unierp" then
		stYear=lgyear-1
	else
		Call CommonQueryRs("entr_dt "," haa010t ","emp_no =  " & FilterVar(parent.txtemp_no.value, "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
		if lgF0="" then
			stYear = "1990"
		else
			stYear =Year(lgF0)
		end if
	end if	
    	For i=lgYear To cint(stYear) step -1
    		Call SetCombo(frm1.txtYear, i, i)
    	Next

        frm1.txtYear.value = CStr(lgYear)

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
' Name : LoadInfTB19029
'========================================================================================================
Private Sub LoadInfTB19029()

	<!--#Include file="../ComAsp/LoadInfTB19029.asp"-->

	<%Call loadInfTB19029(gCurrency,"Q","H")%>

End Sub

'========================================================================================================
' Name : Form_Load
'========================================================================================================
Sub Form_Load()

    Err.Clear                                                                       '☜: Clear err status
    parent.document.All("nextprev").style.VISIBILITY = "hidden"
    Call SetToolBar("1000")    

    Call InitComboBox()
    Call LayerShowHide(0)

    Call InitGrid()
    Call LockField(Document)
    Call LoadInfTB19029()

    Call DbQuery(1)
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
Function DbQuery(ppage)
    Dim strVal
    Err.Clear                                                                    '☜: Clear err status

    DbQuery = False                                                              '☜: Processing is NG

    Call ClearField(document,2)
    Call LayerShowHide(1)
    Call MakeKeyStream("Q")

    strVal = BIZ_PGM_ID & "?txtMode="            & "UID_M0001"                     '☜: Query
    strVal = strVal     & "&txtKeyStream="       & lgKeyStream                   '☜: Query Key
	
	frm1.grid_page.value = 1
    
    Call RunMyBizASP(MyBizASP, strVal)                                           '☜:  Run biz logic

    DbQuery = True                                                               '☜: Processing is NG
End Function

'========================================================================================
' Function Name : DbQueryOk
'========================================================================================
Function DbQueryOk()
    Err.Clear                                                                    '☜: Clear err status

    Call Grid1.ShowData(frm1,frm1.grid_page.value)

End Function

'========================================================================================
' Function Name : DbQueryFail
'========================================================================================
Function DbQueryFail()
    Err.Clear
    Call ClearField(Document,2)                                                                    '☜: Clear err status
	Call Grid1.Clear(frm1,frm1.grid_page.value)
End Function

'========================================================================================================
' Name : SubSum
'========================================================================================================
Sub SubSum()                      '// MB에서 데이터를 넘겨받는다.
    Err.Clear
    Set gArrData = Nothing
    Dim i,j,arrDataRow,arrDataCol

    If RetData="" Then Exit Sub
    arrDataRow = Split(RetData,Chr(12))

    MaxRows = Ubound(arrDataRow,1)-1

    MaxPages = Round(((MaxRows+1)/SheetMaxrows+0.5),0)
    Redim gArrData(MaxRows,MaxCols)

    For i=0 To MaxRows
        arrDataCol = Split(arrDataRow(i),Chr(11))
        For j=0 To MaxCols

            If j=0 Or j=MaxCols Then
                 gArrData(i,j)=i+1
            Else
                if arrDataCol(j) = "" then
                else
                    gArrData(i,j)=arrDataCol(j)
                end if
            End If
        Next
    Next
  
End Sub

'========================================================================================================
' Name : DoubleGetRow
'========================================================================================================
Function DoubleGetRow(pRow)
    If document.all(CStr(pRow)).value="" Then Exit Function

    Dim objList
    Dim elmCnt
    Dim emp_no
    Dim txtYear
    Dim txtType
    Dim txtTypeName
    Dim strUrl,StrEbrFile,arrParam, arrField, arrHeader

	DoubleGetRow = False
	Grid1.ActiveRow = pRow

    txtYear = ""
    txtType = ""

	emp_no = frm1.txtEmp_no.value

    with frm1
    	For elmCnt = 0 to .length - 1
    		Set objList = .elements(elmCnt)
    		If objList.name = "SPREADCELL_PROV_TYPE" & pRow then
               txtType = objList.value
    		End if
    		If objList.name = "SPREADCELL_PROV_NAME" & pRow then
               txtTypeName = objList.value
    		End if
    		If objList.name = "SPREADCELL_PROD_DT" & pRow then
               txtYear = objList.value
    		End if
    	Next
    End With

	if txtType = "1" then
		strUrl = "E1202ma2.asp?Prov_Type=" & txtType
	elseif  txtType = "Z" then
		strUrl = "E1202ma4.asp?Prov_Type=" & txtType
	else
		strUrl = "E1202ma3.asp?Prov_Type=" & txtType
	end if

	strUrl = strUrl & "&Pay_Yymm=" & txtYear
	strUrl = strUrl & "&Emp_no=" & emp_no

	document.location = strUrl

	DoubleGetRow = True
End Function

'========================================================================================================
' Name : MouseRow
'========================================================================================================
Sub MouseRow(pRow)
	If frm1.grid_totpages.value = "" Then Exit Sub
    Dim objList   

	Grid1.ActiveRow = pRow	
	Set objList = window.event.srcElement	
	
	If  UCase(objList.getAttribute("flag")) = "SPREADCELL" then
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

</SCRIPT>
<!-- #Include file="../ESSinc/uniSimsClassID.inc" --> 
</HEAD>
<BODY topmargin=0 leftmargin=0 marginheight=0 marginwidth=0>
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
    <TABLE cellSpacing=0 cellPadding=0 border=0>
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
		                        	<TD class=TDFAMILY_TITLE1>급여년월</TD>
		                        	<TD class=TDFAMILY_TITLE1 width=0></TD>
		                        	<TD class=TDFAMILY_TITLE1>급여구분</TD>
		                        	<TD class=TDFAMILY_TITLE1>급여액</TD>
		                        	<TD class=TDFAMILY_TITLE1>상여액</TD>
		                        	<TD class=TDFAMILY_TITLE1>공제액</TD>
		                        	<TD class=TDFAMILY_TITLE1>지급액</TD>
		                        	<TD class=TDFAMILY_TITLE1>소득세</TD>
		                        	<TD class=TDFAMILY_TITLE1>주민세</TD>
                                </TR>
							    <% 
                                For i=1 To 5 '7
                                    Response.Write "<TR bgcolor=#F8F8F8 height=24 onclick='vbscript: Call DoubleGetRow(" & i & ")' onMouseOver=" & chr(34) & "javascript: this.style.backgroundColor='E1EEF1'" & chr(34) & " onMouseOut=" & chr(34) & "javascript: this.style.backgroundColor=''" & chr(34) & ">"
                                    Response.Write "<TD><INPUT name='" & i & "'  class=listrow01 flag='SPREADCELL' style='WIDTH:  30px; text-align: center;' onMouseOver='vbscript: Call MouseRow(" & i & ")' readonly></TD>"
                                    Response.Write "<TD><INPUT name='SPREADCELL_PROD_DT" & i & "' class=listrow01 flag='SPREADCELL' style='WIDTH:  74px; text-align: center;' onMouseOver='vbscript: Call MouseRow(" & i & ")' readonly></TD>"
                                    Response.Write "<TD><INPUT name='SPREADCELL_PROV_TYPE" & i & "' class=listrow01 flag='SPREADCELL' style='WIDTH: 0px; text-align: center;' onMouseOver='vbscript: Call MouseRow(" & i & ")' readonly></TD>"
                                	Response.Write "<TD><INPUT name='SPREADCELL_PROV_NAME" & i & "' class=listrow01 flag='SPREADCELL' style='WIDTH: 102px; text-align: left;' onMouseOver='vbscript: Call MouseRow(" & i & ")' readonly></TD>"
                                	Response.Write "<TD><INPUT name='SPREADCELL' class=listrow01 flag='SPREADCELL' style='WIDTH:  94px; text-align: right;' onMouseOver='vbscript: Call MouseRow(" & i & ")' readonly></TD>"
                                	Response.Write "<TD><INPUT name='SPREADCELL' class=listrow01 flag='SPREADCELL' style='WIDTH:  94px; text-align: right;' onMouseOver='vbscript: Call MouseRow(" & i & ")' readonly></TD>"
                                	Response.Write "<TD><INPUT name='SPREADCELL' class=listrow01 flag='SPREADCELL' style='WIDTH:  90px; text-align: right;' onMouseOver='vbscript: Call MouseRow(" & i & ")' readonly></TD>"
                                	Response.Write "<TD><INPUT name='SPREADCELL' class=listrow01 flag='SPREADCELL' style='WIDTH:  78px; text-align: right;' onMouseOver='vbscript: Call MouseRow(" & i & ")' readonly></TD>"
                                	Response.Write "<TD><INPUT name='SPREADCELL' class=listrow01 flag='SPREADCELL' style='WIDTH:  80px; text-align: right;' onMouseOver='vbscript: Call MouseRow(" & i & ")' readonly></TD>"
                                	Response.Write "<TD><INPUT name='SPREADCELL' class=listrow01 flag='SPREADCELL' style='WIDTH:  78px; text-align: right;' onMouseOver='vbscript: Call MouseRow(" & i & ")' readonly></TD>"
                                    Response.Write "</TR>"
                                Next
								%>
                           </table>
                        </td>
                    </TR>
                    <TR>
                       <td height="10"></td>
                    </TR>
                    <tr>
                      <td><table width="100%" border="0" cellspacing="1" cellpadding="0">
                          <tr> 
                            <td width="110" height="27" bgcolor="D4E5E8" class="blrow01"><img src="../../CShared/ESSimage/icon_04.gif">급여총액</td>
                            <td width="110" bgcolor="E1EEF1" class="blrow02" align=left><input name="txtTotPayAmt" type="text" class="form02" style="width:120px; text-align:right" readonly></td>
                            <td width="110" height="27" bgcolor="D4E5E8" class="blrow01"><img src="../../CShared/ESSimage/icon_04.gif">상여총액</td>
                            <td width="110" bgcolor="E1EEF1" class="blrow02"><input name="txtTotBonusAmt" type="text" class="form02" style="width:120px; text-align:right" readonly></td>
						  </tr>
					      </table>
					  </td>
					</tr>    
                    <TR>
                       <td height="5"></td>
                    </TR>
                    <tr>
					    <td align=center>
					        <A onclick="VBSCRIPT:CALL GRID1.PREPAGES()" onMouseOver="javascript: this.style.cursor='hand'"><IMG alt="이전페이지" src=../ESSimage/button_07.gif border=0 ></A>&nbsp;
					        <A onclick="VBSCRIPT:CALL GRID1.NEXTPAGES()" onMouseOver="javascript: this.style.cursor='hand'"><IMG alt="다음페이지" src=../ESSimage/button_08.gif border=0 ></A>&nbsp;&nbsp;
					    </td>
					</tr>    
                </TABLE>
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
