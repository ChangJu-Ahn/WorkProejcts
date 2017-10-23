<%@ LANGUAGE="VBSCRIPT" %>
<!--
'======================================================================================================
'*  1. Module Name          : 
'*  2. Function Name        : 
'*  3. Program ID           : 
'*  4. Program Name         : 
'*  5. Program Desc         : 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : 
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'=======================================================================================================-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/incSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->
<!-- #include file="../../Inc/Common.asp"--> <%'추가파일 %>

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incHRQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incEB.vbs"></SCRIPT>

<Script Language="VBScript">
Option Explicit 
'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================

Sub SelectDB
    document.location.href = "ZK000MA6.asp"
End Sub

'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================
Function FncExit()
    FncExit = True
End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>
<%

	Call LoadBasisGlobalInf()

	MasterConnString = MakeConnString(GetGlobalInf("gDBServerIP") ,GetGlobalInf("gDBLoginID"),GetGlobalInf("gDBSAPwd"),"master")      

	Const ggWidth = 700
	Const gpWidth = 500
	Dim SDBList
	DIm TDBList
	    
	SDBList = Trim(UCase(Request.Form("SDBList")))
	TDBList = Trim(UCase(Request.Form("TDBList")))
	    
	Session("SDB") = EnCode(SDBList)
	Session("TDB") = EnCode(TDBList)
   
%>
<body scroll=auto>

<TABLE >
<TR HEIGHT=*>
<TD WIDTH=10>&nbsp;</TD>
<TD>

	<TABLE CLASS="BatchTB1" CELLSPACING=0 CELLPADDING=0>

		<TR>
			<TD <%=HEIGHT_TYPE_00%> ></TD>
		</TR>
		<TR HEIGHT=23>
			<TD >
				<TABLE <%=LR_SPACE_TYPE_10%>>
					<TR>
						<TD WIDTH=10>&nbsp;</TD>
						<TD CLASS="CLSMTABP">
							<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
								<TR>
									<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
									<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>테이블복사</font></td>
									<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
							    </TR>
							</TABLE>
						</TD>
						<TD WIDTH=* HEIGHT="right">&nbsp;</TD>
						<TD WIDTH=10>&nbsp;</TD>
					</TR>
				</TABLE>
			</TD>
		</TR>
		

		<TR HEIGHT=*>
			<TD>
				<TABLE <%=LR_SPACE_TYPE_60%> >   
				  <tr>
					<TR>
						<TD  STYLE="PADDING-RIGHT: 5px;BACKGROUND-COLOR: #d1e8f9;" ALIGN=LEFT NOWRAP COLSPAN=2>서버정보</TD>
					</TR>
					<TR>
						<TD STYLE="PADDING-RIGHT: 5px;BACKGROUND-COLOR: #d1e8f9;" ALIGN=LEFT WIDTH=10% NOWRAP>데이타베이스서버</TD>
						<TD CLASS=TD6 NOWRAP><%=GetGlobalInf("gDBServerIP")%></TD>
					</TR>
					<TR>
						<TD STYLE="PADDING-RIGHT: 5px;BACKGROUND-COLOR: #d1e8f9;" ALIGN=LEFT  WIDTH=10% NOWRAP>Source DataBase</TD>
						<TD CLASS=TD6 NOWRAP><%=SDBList%></TD>
					</TR>
					<TR>
						<TD STYLE="PADDING-RIGHT: 5px;BACKGROUND-COLOR: #d1e8f9;" ALIGN=LEFT  WIDTH=10% NOWRAP>Target DataBase</TD>
						<TD CLASS=TD6 NOWRAP><%=TDBList%></TD>
					</TR>
					<TR>
				  </tr>
				</table>
			</TD>	
		</TR>
		
		<TR HEIGHT=10><TD>&nbsp;</TD></TR>
		
		<table style="BACKGROUND-COLOR: #eeeeec;BORDER-RIGHT: buttonshadow 1px solid;PADDING-RIGHT: 0px;BORDER-TOP: buttonshadow 1px solid;PADDING-LEFT: 0px;PADDING-BOTTOM: 0px;BORDER-LEFT: buttonshadow 1px solid;PADDING-TOP: 0px;BORDER-BOTTOM: buttonshadow 1px solid;" border=0 cellspacing=1 cellpadding=1 width=<%=ggWidth%>>
		  <tr>
		    <TD STYLE="PADDING-RIGHT: 5px;BACKGROUND-COLOR: #d1e8f9;"  WIDTH=100% ALIGN=LEFT>전체 진척도 </TD>
		  </tr>
		</table>

		<table style="BACKGROUND-COLOR: #eeeeec;BORDER-RIGHT: buttonshadow 1px solid;PADDING-RIGHT: 0px;BORDER-TOP: buttonshadow 1px solid;PADDING-LEFT: 0px;PADDING-BOTTOM: 0px;BORDER-LEFT: buttonshadow 1px solid;PADDING-TOP: 0px;BORDER-BOTTOM: buttonshadow 1px solid;" border=0 cellspacing=1 cellpadding=1  width=<%=ggWidth%>>
		  <tr>
		    <TD>
		        <table border=0 cellspacing=1 cellpadding=1 bgcolor="#cccccc" width=100%>
		            <TR> <TD bgcolor=#E7F1D9>작업명           </TD><TD bgcolor=#F1F4EC colspan=3><SPAN NAME=txtState     ID=txtState>데이터베이스 생성중 <SPAN></TD>    </TR>
		            <TR> <TD bgcolor=#E7F1D9 width=15%>테이블명</TD><TD bgcolor=#F1F4EC    width=35%><SPAN NAME=txtObjectName ID=txtObjectName><SPAN></TD><TD bgcolor=#E7F1D9 width=15%>진행율</TD><TD bgcolor=#F1F4EC  width=35%><SPAN NAME=txtPercentage ID=txtPercentage ALIGN=RIGHT width=100><SPAN></TD>    </TR>
		            <TR> <TD bgcolor=#E7F1D9>진행율            </TD><TD bgcolor=#F1F4EC colspan=3><center>
		            <DIV align="left" STYLE="width:<%=gpWidth%>px;height:16px;border-width:1px;border-style:solid;border-color:silver">
		          <DIV ID="divProgress" STYLE="width:0px;height:15px;background-color:#FF6666"></DIV>
		          </DIV>
		          </center></TD>   </TR>
					<TR> <TD bgcolor=#E7F1D9>경과시간          </TD><TD bgcolor=#F1F4EC colspan=3><SPAN NAME=txtMDB     ID=txtMDB><SPAN></TD>    </TR>
		            <TR> <TD bgcolor=#E7F1D9>메시지            </TD><TD bgcolor=#F1F4EC colspan=3><SPAN NAME=txtMessage ID=txtMessage><SPAN></TD>    </TR>
		        </table>
		    </TD>
		  </tr>
		</table>
			

	</TABLE>
	
</TD>
</TR>
</TABLE>

<%

    Dim DBName,WorkingDir
    
    Dim gtotObjectCount
    Dim gstaTimer

    gtotObjectCount = 0 

    If Response.Buffer Then Response.Flush

    MasterConnString = MakeConnString(GetGlobalInf("gDBServerIP"),GetGlobalInf("gDBLoginID"),GetGlobalInf("gDBSAPwd"),"master")   
    MetaConnString   = MakeConnString(GetGlobalInf("gDBServerIP"),GetGlobalInf("gDBLoginID"),GetGlobalInf("gDBSAPwd"),gDataBase    )   
    SourceConnString = MakeConnString(GetGlobalInf("gDBServerIP"),GetGlobalInf("gDBLoginID"),GetGlobalInf("gDBSAPwd"),SDBList)   
    TargetConnString = MakeConnString(GetGlobalInf("gDBServerIP"),GetGlobalInf("gDBLoginID"),GetGlobalInf("gDBSAPwd"),TDBList)      
    
    gstaTimer = Timer

    Call Crete_usp_tableDependencies(SourceConnString)

    Call execSQL(MetaConnString,"update Z_TABLE_LIST set xBit = 0 ")
    
    Call MakeRelatedTable()

    Call Process()

Sub Process()

    Dim AdoConn
    Dim iLoop
    Dim iCommCount
    Dim iCommList
    Dim iLoopContents
    Dim intIncrement
    Dim iLoopCount
    Dim bTimer
    Dim iSTRSQL
    
    bTimer = Timer
    
    iCommCount = GetMetaTableCount(MetaConnString,Request.Form("RS_M"),Request.Form("RS_T"))
    
    iCommList  = GetMetaTableList(MetaConnString,Request.Form("RS_M"),Request.Form("RS_T"))
    
   intIncrement = gpWidth / iCommCount

    If Response.Buffer Then Response.Flush
    
    iLoopContents = Split(iCommList, ",")
    
    iLoopCount = UBound(iLoopContents)
    
    Set AdoConn = Server.CreateObject("ADODB.Connection")
    adoConn.Open MetaConnString
    iLoop =0 
    For iLoop = 0 To iLoopCount
        iSTRSQL = " dbo.CopyTABLE '" & GetGlobalInf("gDBServerIP") & "' , '" & SDBList & "','" & TDBList & "','" & GetGlobalInf("gDBLoginID") & "','" & GetGlobalInf("gDBSAPwd") & "','" & iLoopContents(iLoop) & "' "
     
        adoConn.Execute iSTRSQL
        Call RefreshProgressBar2(CInt(iLoop * intIncrement * 100 / gpWidth), CInt(iLoop * intIncrement), iLoopContents(iLoop),  CalElaspeTime(bTimer, Timer))
    Next
    
    Call RefreshProgressBar2(100, gpWidth, iLoopContents(iLoopCount),  CalElaspeTime(bTimer, Timer))
    adoConn.Close    
    Set adoConn = Nothing
    
End Sub

Sub MakeRelatedTable()
    Dim iTemp
    Dim iLoop
    Dim jLoop
    Dim iCommCount
    Dim iCommList
    Dim iLoopContents
    Dim intIncrement
    Dim iLoopCount
    Dim bTimer
    Dim iSTRSQL

    Call execSQL(MetaConnString,"update Z_TABLE_LIST set xBit = 0 ")

    iCommCount = GetMetaTableCount(MetaConnString,Request.Form("RS_M"),Request.Form("RS_T"))
    
    iCommList  = GetMetaTableList(MetaConnString,Request.Form("RS_M"),Request.Form("RS_T"))
    
   intIncrement = gpWidth / iCommCount

    If Response.Buffer Then Response.Flush
    
    iLoopContents = Split(iCommList, ",")
    
    iLoopCount = UBound(iLoopContents)

    For iLoop = 0 To iLoopCount
        iTemp = GetForeignTable(SourceConnString,iLoopContents(iLoop))
        If iTemp <> "" Then
           If Instr(iTemp,",") > 0 Then
              iTemp = Split(iTemp,",")
              For jLoop = 0 To UBound(iTemp)
                  Call execSQL(MetaConnString,"update Z_TABLE_LIST set xBit = 1 where table_id = '" & iTemp(jLoop) & "' ")
                  'Response.Write iTemp(jLoop) & "<br>"
              Next
           Else
              Call execSQL(MetaConnString,"update Z_TABLE_LIST set xBit = 0 where table_id = '" & iTemp & "' ")
           End If
        End If   
        Call RefreshProgressBar(CInt(iLoop * intIncrement * 100 / gpWidth), CInt(iLoop * intIncrement), iLoopContents(iLoop),  CalElaspeTime(bTimer, Timer))
    Next
       
End Sub


Sub RefreshProgressBar(ByVal ProgressBarPercentage,ByVal ProgressBarValue,ByVal TableName,ByVal pTime)
%>
    <SCRIPT LANGUAGE="VBS">
       Dim intWidth
       
       document.all("txtState").innerText   = "테이블 관계성 검사"
       document.all("txtMessage").innerText = "테이블 관계성 검사중"
       document.all("txtMDB").innerText   = "<%=pTime%>"
     
       document.all("divProgress").style.width =  <%= ProgressBarValue %>
       document.all("txtObjectName").innerText = "<%= TableName %>"
       document.all("txtPercentage").innerHTML = Right("     " & "<%= ProgressBarPercentage %>",6) & "%"
       
       If <%=ProgressBarPercentage%> = 100 Then
       document.all("txtMessage").innerText = "테이블 관계성 검사 완료"
       End If
    </SCRIPT>
<%
  If Response.Buffer Then Response.Flush
  
End Sub

Sub RefreshProgressBar2(ByVal ProgressBarPercentage,ByVal ProgressBarValue,ByVal TableName,ByVal pTime)
%>
    <SCRIPT LANGUAGE="VBS">
       Dim intWidth
       
       document.all("txtState").innerText   = "테이블 복사"
       document.all("txtMessage").innerText = "테이블 복사중"
       document.all("txtMDB").innerText   = "<%=pTime%>"
     
       document.all("divProgress").style.width =  <%= ProgressBarValue %>
       document.all("txtObjectName").innerText = "<%= TableName %>"
       document.all("txtPercentage").innerHTML = Right("     " & "<%= ProgressBarPercentage %>",6) & "%"
       
       If <%=ProgressBarPercentage%> = 100 Then
       document.all("txtMessage").innerText = "테이블 복사 완료"
       End If
    </SCRIPT>
<%
  If Response.Buffer Then Response.Flush
  
End Sub


Sub ShowMessage(ByVal pData)
%>
    <SCRIPT LANGUAGE="VBS">
       document.all("txtMessage").innerText  = "<%= pData %>"
    </SCRIPT>
<%
End Sub


%>
