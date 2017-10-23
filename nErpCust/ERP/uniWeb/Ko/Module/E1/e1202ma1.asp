<%@ LANGUAGE="VBSCRIPT" %>
<% Response.Expires = -1%>
<HTML>
<HEAD>
<TITLE><%=Request("strTitle")%></TITLE>

<!--
========================================================================================================
=                          3.1 Server Side Script
========================================================================================================-->
<!-- #Include file="../../inc/IncServer.asp"  -->
<!--
========================================================================================================
=                          3.2 Style Sheet
======================================================================================================== -->
<LINK REL="stylesheet" TYPE="Text/css" href="../../inc/CommStyleSheet.css">

<!--
========================================================================================================
=                          3.3 Client Side Script
======================================================================================================== -->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/ccm.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/variables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/operation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCommFunc.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incEvent.vbs">   </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/adoQuery.vbs"></SCRIPT>
<!-- #Include file="../../inc/incSvrVarSims.inc"  -->
<!-- #Include file="../../inc/incSvrFuncSims.inc" -->
<Script Language="VBScript">
Option Explicit                                                        '☜: indicates that All variables must be declared in advance

Const BIZ_PGM_ID      = "e1202mb1.asp"						           '☆: Biz Logic ASP Name
Const C_SHEETMAXROWS = 10	                                      '☜: Visble row
Const C_SHEETMAXCOLS = 10

'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" --> 
<!-- #Include file="../../inc/incGrid.inc" -->
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim Grid1

'========================================================================================================
' Function Name : MakeKeyStream
' Function Desc : This method set focus to pos of err
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
' Desc : Set ComboBox
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

    	'frm1.txtYear.remove 0
        frm1.txtYear.value = CStr(lgYear)

End Sub
'========================================================================================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================================
Sub InitGrid()
    Set Grid1 = New Grid
    Grid1.MaxCols = C_SHEETMAXCOLS
    Grid1.SheetMaxrows = C_SHEETMAXROWS
    Set Grid1.Source = document.frm1
End Sub

'========================================================================================================
' Name : LoadInfTB19029
' Desc : 
'========================================================================================================
Private Sub LoadInfTB19029()

<!--#Include file="../../ComAsp/LoadInfTB19029.asp"-->

<%Call loadInfTB19029(gCurrency,"Q","H")%>

End Sub
'========================================================================================================
' Name : Form_Load
' Desc : developer describe this line Called by Window_OnLoad() evnt
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
' Function Desc : 페이지 전환이나 화면이 닫힐 경우 실행해야 될 로직 처리 
'========================================================================================
Sub Form_UnLoad()
	On Error Resume Next
    Set Grid1 = Nothing
End Sub

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

Function DbQueryOk()
    Err.Clear                                                                    '☜: Clear err status

    Call Grid1.ShowData(frm1,frm1.grid_page.value)

End Function

Function DbQueryFail()
    Err.Clear
    Call ClearField(Document,2)                                                                    '☜: Clear err status
	Call Grid1.Clear(frm1,frm1.grid_page.value)
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
		.txtMode.value        = "UID_M0002"                                        '☜: Save
'		.txtFlgMode.value     = lgIntFlgMode
        .txtKeyStream.Value   = lgKeyStream                                      '☜: Save Key
	End With

	Call ExecMyBizASP(frm1, BIZ_PGM_ID)
		
    DbSave  = True                                                               '☜: Processing is NG
End Function

'========================================================================================================
' Function Name : DbSaveOk
' Function Desc : Called by MB Area when save operation is successful
'========================================================================================================
Function DbSaveOk()
    Call DbQuery(1)
End Function

Function GetRow(pRow)
	GetRow=False
    Grid1.ActiveRow = pRow
    If Mid(document.activeElement.getAttribute("tag"),3,1) = "1" Then
	    arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
	    	"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	End If
	GetRow=True
End Function

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
	'Call CommonQueryRs(" MENU_NAME "," E11000T "," Menu_id = 'E1202MA1' AND LANG_CD = '" & gLang &"'",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
		
	'	parent.txtTitle.value = replace(lgF0,chr(11),"")
  document.location = strUrl

'	window.showModalDialog strUrl, Array(arrParam, arrField, arrHeader), _
'	"dialogWidth=800px; dialogHeight=700px; center: Yes; help: No; resizable: Yes; status: No;"

	DoubleGetRow = True
End Function

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

Sub Print_onClick()
    Call SubPrint(MyBizASP)
End Sub

Sub GRID_PAGE_OnChange()
End Sub

Sub DELETE_OnClick()
    Call Grid1.DeleteClick()
End Sub

Sub CANCEL_OnClick()
    Call Grid1.CancelClick()
End Sub

</SCRIPT>
<!-- #Include file="../../inc/uniSimsClassID.inc" --> 
</HEAD>
<BODY topmargin=0 leftmargin=0 marginheight=0 marginwidth=0>
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
    <TABLE cellSpacing=0 cellPadding=0 width=749 border=0>
        <TR>
            <TD width=13></TD>
            <TD>
                <TABLE cellSpacing=0 cellPadding=0 width=705 border=0 bgcolor=#ffffff>
                    <TR height=26 valign=middle>
                        <TD class=base1>사번:<INPUT class=base1 NAME="txtEmp_no" MAXLENGTH=13 SiZE=12 tag=14></TD>
                        <TD class=base1>성명:<INPUT class=base1 NAME="txtName" MAXLENGTH=20 SiZE=10  tag=14></TD>
                        <TD class=base1>직위:<INPUT class=base1 NAME="txtroll_pstn" MAXLENGTH=20 SiZE=10  tag=14></TD>
                        <TD class=base1>부서:<INPUT class=base1 NAME="txtDept_nm" MAXLENGTH=25 SiZE=15  tag=14></TD>
                    </TR>
                    <TR height=25 valign=middle>
                        <TD class=base1 valign=middle>정산년도:
						    <SELECT Name="txtYear" tabindex=-1 STYLE="WIDTH: 100px">
						    </SELECT>
                        </TD>
                        <TD colspan=3 CLASS=base1></TD>
                    </TR>
                    <TR>
                        <TD colspan=4>
                            <TABLE cellSpacing=1 cellPadding=0 width="100%" border=0 bgcolor=#ffffff>
                                <TR bgcolor=#d0d6e4 height=20>
		                        	<TD></TD>
		                        	<TD class=TDFAMILY_TITLE1>급여년월</TD>
		                        	<TD class=TDFAMILY_TITLE1 width=0></TD>
		                        	<TD class=TDFAMILY_TITLE1>구분</TD>
		                        	<TD class=TDFAMILY_TITLE1>급여총액</TD>
		                        	<TD class=TDFAMILY_TITLE1>상여총액</TD>
		                        	<TD class=TDFAMILY_TITLE1>과세</TD>
		                        	<TD class=TDFAMILY_TITLE1>비과세</TD>
		                        	<TD class=TDFAMILY_TITLE1>소득세</TD>
		                        	<TD class=TDFAMILY_TITLE1>주민세</TD>
                                </TR>
							<%
						        For i = 1 To 10
						            Response.Write "<TR bgcolor=#E9EDF9 height=20 onclick='vbscript: Call DoubleGetRow(" & i & ")' onMouseOver=" & chr(34) & "javascript: this.style.backgroundColor='FEE2E3'" & chr(34) & " onMouseOut=" & chr(34) & "javascript: this.style.backgroundColor=''" & chr(34) & ">"
						            Response.Write "<TD><INPUT name='" & i & "'  tag='25X' flag='SPREADCELL' style='WIDTH:  30px;  TEXT-ALIGN: center' onMouseOver='vbscript: Call MouseRow(" & i & ")'></TD>"
						            Response.Write "<TD><INPUT name='SPREADCELL_PROD_DT" & i & "' tag='25X' flag='SPREADCELL' style='WIDTH:  90px; TEXT-ALIGN: center' onMouseOver='vbscript: Call MouseRow(" & i & ")'></TD>"
						            Response.Write "<TD width=0><INPUT name='SPREADCELL_PROV_TYPE" & i & "' TYPE=HIDDEN flag='SPREADCELL' style='WIDTH: 0px; TEXT-ALIGN: center' onMouseOver='vbscript: Call MouseRow(" & i & ")'></TD>"
						            Response.Write "<TD><INPUT name='SPREADCELL_PROV_NAME" & i & "' tag='25X' flag='SPREADCELL' style='WIDTH:  80px; TEXT-ALIGN: center' onMouseOver='vbscript: Call MouseRow(" & i & ")'></TD>"
						        	Response.Write "<TD><INPUT name='SPREADCELL' tag='25X' flag='SPREADCELL' style='WIDTH:  100px; TEXT-ALIGN: right' onMouseOver='vbscript: Call MouseRow(" & i & ")'></TD>"
						        	Response.Write "<TD><INPUT name='SPREADCELL' tag='25X' flag='SPREADCELL' style='WIDTH:  100px; TEXT-ALIGN: right' onMouseOver='vbscript: Call MouseRow(" & i & ")'></TD>"
						        	Response.Write "<TD><INPUT name='SPREADCELL' tag='25X' flag='SPREADCELL' style='WIDTH:  90px; TEXT-ALIGN: right' onMouseOver='vbscript: Call MouseRow(" & i & ")'></TD>"
						        	Response.Write "<TD><INPUT name='SPREADCELL' tag='25X' flag='SPREADCELL' style='WIDTH:  90px; TEXT-ALIGN: right' onMouseOver='vbscript: Call MouseRow(" & i & ")'></TD>"
						        	Response.Write "<TD><INPUT name='SPREADCELL' tag='25X' flag='SPREADCELL' style='WIDTH:  90px; TEXT-ALIGN: right' onMouseOver='vbscript: Call MouseRow(" & i & ")'></TD>"
						        	Response.Write "<TD><INPUT name='SPREADCELL' tag='25X' flag='SPREADCELL' style='WIDTH:  88px; TEXT-ALIGN: right' onMouseOver='vbscript: Call MouseRow(" & i & ")'></TD>"
						            Response.Write "</TR>"
						        Next
							%>
                            </TABLE>
                        </TD>
                    </TR>
                </TABLE>
            </TD>
            <TD width=14></TD>
        </TR>
        <TR height=20>
            <TD width=13></TD>
            <TD>
                <TABLE cellSpacing=0 cellPadding=0 border=0><TR>
                    <TD align=left width=680>
                        급여총액<INPUT TYPE=text NAME="txtTotPayAmt" size=14 tag=24 style='TEXT-ALIGN: right'>
                        상여총액<INPUT TYPE=text NAME="txtTotBonusAmt" size=14 tag=24 style='TEXT-ALIGN: right'>
                    </TD>
                    <TD align=right width=100>
                        <A onclick="VBSCRIPT:CALL GRID1.PREPAGES()" onMouseOver="javascript: this.style.cursor='hand'"><IMG alt="이전페이지" src=../../../Cshared/Image/uniSIMS/gprev.jpg border=0 ></A>&nbsp;
                        <A onclick="VBSCRIPT:CALL GRID1.NEXTPAGES()" onMouseOver="javascript: this.style.cursor='hand'"><IMG alt="다음페이지" src=../../../Cshared/Image/uniSIMS/gnext.jpg border=0 ></A>&nbsp;&nbsp;
                    </TD>
                    </TR>
                </TABLE>
            </TD>
            <TD width=14></TD>
        </TR>    
    </TABLE>
    <TABLE cellSpacing=0 cellPadding=0 width=700 border=0 bgcolor=#ffffff>
        <TR><TD WIDTH="100%" HEIGHT=0><IFRAME NAME="MyBizASP" WIDTH="100%" HEIGHT=0 FRAMEBORDER=0 SCROLLING=no noresize framespacing=0></IFRAME></TD></TR>
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

</BODY>
</HTML>
