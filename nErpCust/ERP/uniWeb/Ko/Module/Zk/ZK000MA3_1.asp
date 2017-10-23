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
'*  9. Modifier (First)     : Yim Young Ju
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'=======================================================================================================-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/incSvrHTML.inc" -->
<!-- #include file="../../Inc/Common.asp"--> <%'추가파일 %>
<!-- #Include file="../../inc/incSvrMain.asp"  -->
<%
	Server.ScriptTimeout = 6600
	
    Const ggWidth = 700
    Const gpWidth = 500
    Session("SDB") = EnCode(Trim(UCase(Request.Form("SDB"))))
    Session("TDB") = EnCode(Trim(UCase(Request.Form("TDB"))))
   
%>

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		

<Script Language="VBScript">
Option Explicit 

Function FncExit()
    FncExit = True
End Function


Sub SelectDB
    top.location.href = "ZK000MA2.asp"
End Sub

</SCRIPT>
<%
Call LoadBasisGlobalInf()

%>
</HEAD>

<BODY TABINDEX="-1" SCROLL="NO" >

<table border=0 cellspacing=1 cellpadding=1 CLASS=BatchTB1 width=<%=ggWidth%>>

	<TR>
		<TD <%=HEIGHT_TYPE_00%>></TD>
	</TR>
	<TR HEIGHT=23>
		<TD>
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>데이터베이스생성</font></td>
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
	
	<TR HEIGHT=30>
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
					<TD STYLE="PADDING-RIGHT: 5px;BACKGROUND-COLOR: #d1e8f9;" ALIGN=LEFT  WIDTH=10% NOWRAP>소스데이타베이스</TD>
					<TD CLASS=TD6 NOWRAP><%=DeCode(Session("SDB"))%></TD>
				</TR>
				<TR>
					<TD STYLE="PADDING-RIGHT: 5px;BACKGROUND-COLOR: #d1e8f9;" ALIGN=LEFT  WIDTH=10% NOWRAP>생성데이타베이스</TD>
					<TD CLASS=TD6 NOWRAP><%=DeCode(Session("TDB"))%></TD>
				</TR>
				<TR>
			  </tr>
			</table>
		</TD>	
	</TR>
	
	<TR HEIGHT=10><TD></TD></TR>
	
	<TR>
		<TD>
			<table style="BACKGROUND-COLOR: #eeeeec;BORDER-RIGHT: buttonshadow 1px solid;PADDING-RIGHT: 0px;BORDER-TOP: buttonshadow 1px solid;PADDING-LEFT: 0px;PADDING-BOTTOM: 0px;BORDER-LEFT: buttonshadow 1px solid;PADDING-TOP: 0px;BORDER-BOTTOM: buttonshadow 1px solid;" border=0 cellspacing=1 cellpadding=1 width=<%=ggWidth%>>
			  <tr>
			    <TD STYLE="PADDING-RIGHT: 5px;BACKGROUND-COLOR: #d1e8f9;"  WIDTH=100% ALIGN=LEFT>전체 진척도 </TD>
			  </tr>
			</table>
						  

			<table border=0 cellspacing=1 cellpadding=1 bgcolor="#cccccc" width=100% >
			    <tr><TD bgcolor=#E7F1D9 WIDTH=40%>&nbsp;                      </TD><TD bgcolor=#E7F1D9 WIDTH=40% ><CENTER>추출.생성                                                  </CENTER></TD><TD bgcolor=#E7F1D9 WIDTH=10% ALIGN=CENTER>소요시간                            </TD><TD bgcolor=#E7F1D9 WIDTH=10% ALIGN=CENTER>결과건수</TD></tr>
			    <tr><TD bgcolor=#ECF5EB          >[00]데이터베이스            </TD><TD bgcolor=#F8D2D2 ID=idCDB  ><CENTER><INPUT TYPE=CHECKBOX ID=chkEDB   name=chkEDB   CLASS=RADIO></CENTER></TD><TD bgcolor=#EFF1EF           ALIGN=CENTER><SPAN NAME=txtMDB  ID=txtMDB ><SPAN></TD><TD bgcolor=#EFF1EF           ALIGN=RIGHT><SPAN NAME=txtMCDB  ID=txtMCDB ><SPAN>&nbsp;&nbsp;</TD></tr>
			    <tr><TD bgcolor=#ECF5EB          >[01]테이블[주키,인덱스]     </TD><TD bgcolor=#EFF1EF ID=idETAB ><CENTER><INPUT TYPE=CHECKBOX id=chkETAB  name=chkETAB  CLASS=RADIO></CENTER></TD><TD bgcolor=#EFF1EF           ALIGN=CENTER><SPAN NAME=txtMTAB ID=txtMTAB><SPAN></TD><TD bgcolor=#EFF1EF           ALIGN=RIGHT><SPAN NAME=txtMCTAB ID=txtMCTAB><SPAN>&nbsp;&nbsp;</TD></tr>
			    <tr><TD bgcolor=#ECF5EB          >[02]테이블[외래키,제약조건] </TD><TD bgcolor=#EFF1EF ID=idEFK  ><CENTER><INPUT TYPE=CHECKBOX id=chkEFK   name=chkEFK   CLASS=RADIO></CENTER></TD><TD bgcolor=#EFF1EF           ALIGN=CENTER><SPAN NAME=txtMFK  ID=txtMFK ><SPAN></TD><TD bgcolor=#EFF1EF           ALIGN=RIGHT><SPAN NAME=txtMCFK  ID=txtMCFK ><SPAN>&nbsp;&nbsp;</TD></tr>
			    <tr><TD bgcolor=#ECF5EB          >[03]테이블[트리거]          </TD><TD bgcolor=#EFF1EF ID=idETRG ><CENTER><INPUT TYPE=CHECKBOX id=chkETRG  name=chkETRG  CLASS=RADIO></CENTER></TD><TD bgcolor=#EFF1EF           ALIGN=CENTER><SPAN NAME=txtMTRG ID=txtMTRG><SPAN></TD><TD bgcolor=#EFF1EF           ALIGN=RIGHT><SPAN NAME=txtMCTRG ID=txtMCTRG><SPAN>&nbsp;&nbsp;</TD></tr>
			    <tr><TD bgcolor=#ECF5EB          >[04]저장프로시져            </TD><TD bgcolor=#EFF1EF ID=idEPRC ><CENTER><INPUT TYPE=CHECKBOX id=chkEPRC  name=chkEPRC  CLASS=RADIO></CENTER></TD><TD bgcolor=#EFF1EF           ALIGN=CENTER><SPAN NAME=txtMPRC ID=txtMPRC><SPAN></TD><TD bgcolor=#EFF1EF           ALIGN=RIGHT><SPAN NAME=txtMCPRC ID=txtMCPRC><SPAN>&nbsp;&nbsp;</TD></tr>
			    <tr><TD bgcolor=#ECF5EB          >[05]사용자 정의함수         </TD><TD bgcolor=#EFF1EF ID=idEUDF ><CENTER><INPUT TYPE=CHECKBOX id=chkEUDF  name=chkEUDF  CLASS=RADIO></CENTER></TD><TD bgcolor=#EFF1EF           ALIGN=CENTER><SPAN NAME=txtMUDF ID=txtMUDF><SPAN></TD><TD bgcolor=#EFF1EF           ALIGN=RIGHT><SPAN NAME=txtMCUDF ID=txtMCUDF><SPAN>&nbsp;&nbsp;</TD></tr>
			    <tr><TD bgcolor=#ECF5EB          >[06]뷰                      </TD><TD bgcolor=#EFF1EF ID=idEVIW ><CENTER><INPUT TYPE=CHECKBOX id=chkEVIW  name=chkEVIW  CLASS=RADIO></CENTER></TD><TD bgcolor=#EFF1EF           ALIGN=CENTER><SPAN NAME=txtMVIW ID=txtMVIW><SPAN></TD><TD bgcolor=#EFF1EF           ALIGN=RIGHT><SPAN NAME=txtMCVIW ID=txtMCVIW><SPAN>&nbsp;&nbsp;</TD></tr>
			    <tr><TD bgcolor=#ECF5EB          >[07]집계                    </TD><TD bgcolor=#EFF1EF ID=idTOT ><CENTER>&nbsp;                                                      </CENTER></TD><TD bgcolor=#EFF1EF           ALIGN=CENTER><SPAN NAME=txtMTOT ID=txtMTOT><SPAN></TD><TD bgcolor=#EFF1EF           ALIGN=RIGHT><SPAN NAME=txtMCUDF ID=txtMCTOT><SPAN>&nbsp;&nbsp;</TD></tr>
			</table>

			  
		</TD>
	</TR>
	
	<TR HEIGHT=10><TD>&nbsp;</TD></TR>
	
	<table style="BACKGROUND-COLOR: #eeeeec;BORDER-RIGHT: buttonshadow 1px solid;PADDING-RIGHT: 0px;BORDER-TOP: buttonshadow 1px solid;PADDING-LEFT: 0px;PADDING-BOTTOM: 0px;BORDER-LEFT: buttonshadow 1px solid;PADDING-TOP: 0px;BORDER-BOTTOM: buttonshadow 1px solid;" border=0 cellspacing=1 cellpadding=1 width=<%=ggWidth%>>
	  <tr>
	    <TD STYLE="PADDING-RIGHT: 5px;BACKGROUND-COLOR: #d1e8f9;"  WIDTH=100% ALIGN=LEFT>현재 진척도 </TD>
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
	            <TR> <TD bgcolor=#E7F1D9>메시지            </TD><TD bgcolor=#F1F4EC colspan=3><SPAN NAME=txtMessage ID=txtMessage><SPAN></TD>    </TR>
	        </table>
	    </TD>
	  </tr>
	</table>



<%

    Dim DBName,WorkingDir
    
    Dim gtotObjectCount
    Dim gstaTimer

    gtotObjectCount = 0 

    If Response.Buffer Then Response.Flush

	MasterConnString = MakeConnString(GetGlobalInf("gDBServerIP"),GetGlobalInf("gDBLoginID"),GetGlobalInf("gDBSAPwd"),"master")   
	MetaConnString   = MakeConnString(GetGlobalInf("gDBServerIP"),GetGlobalInf("gDBLoginID"),GetGlobalInf("gDBSAPwd"),gDataBase)   
	SourceConnString = MakeConnString(GetGlobalInf("gDBServerIP"),GetGlobalInf("gDBLoginID"),GetGlobalInf("gDBSAPwd"),DeCode(Session("SDB")))   
	TargetConnString = MakeConnString(GetGlobalInf("gDBServerIP"),GetGlobalInf("gDBLoginID"),GetGlobalInf("gDBSAPwd"),DeCode(Session("TDB")))
			        
    
    DBName	   = DeCode(Session("SDB")) 
    WorkingDir = "C:\XCV"
    
    gstaTimer = Timer
    
    Call CreateDB()
   
    Call Process("ETAB")
    Call Process("EFK")
    Call Process("ETRG")   
    Call Process("EPRC")   
    Call Process("EUDF")   
    Call Process("EVIW")   
    Call RefreshProgressTotResultBar(gtotObjectCount,CalElaspeTime(gstaTimer, Timer))
    Call ShowMessage("작업이 완료되었습니다.")
    
Sub CreateDB()    
    Dim AdoConn
    Dim iDBName
    Dim bTimer

    bTimer = Timer
    
    iDBName = DeCode(Session("TDB"))
    
    If iDBName = "" Then
       Call ShowMessage("생성될 데이터베이스명이 지정되지 않았습니다.")
       Response.End 
    End If
    
    Set AdoConn = Server.CreateObject("ADODB.Connection")    
    AdoConn.Open SourceConnString

    If Err.number <> 0 Then
       Call ShowMessage(Err.Description)
       Response.End 
    End If
    
    Call RefreshProgressBar(0,gpWidth,iDBName,"EDB","00:00:00")

    AdoConn.Execute "dbo.dmoCreateDB '" & iDBName & "'"

    If Err.number = 0 Then
       Call RefreshProgressBar(100,gpWidth,iDBName,"EDB",CalElaspeTime(bTimer,Timer))
    Else   
       Call ShowMessage("데이터베이스 생성중 오류가 발생했습니다.")
       Response.End 
    End If

    AdoConn.Close
    
    Set AdoConn = Nothing
    
End Sub    


Sub Process(ByVal pOpt)

    Dim iLoopContents
    Dim iLoopCount
    Dim intIncrement
    Dim iLoop
    Dim iCommCount
    Dim iCommList
    Dim iSPName
    Dim bTimer
    Dim iSubFolder
    Dim iExtension
    Dim adoConn
    Dim iSQL
    
    bTimer = Timer  
    
    Select Case pOpt
         Case "ETAB":  iCommCount = GetTableCount(SourceConnString)
                       iSPName = "dbo.dmoScriptDatabase_TAB "
                       iSubFolder = "TAB"
                       iExtension = "TAB"
         Case "EFK":   iCommCount = GetTableCount(SourceConnString)
                       iSPName = "dbo.dmoScriptDatabase_FK  "
                       iSubFolder = "FK"
                       iExtension = "FK"
         Case "ETRG":  iCommCount = GetTableCount(SourceConnString)
                       iSPName = "dbo.dmoScriptDatabase_TRG "
                       iSubFolder = "TRG"
                       iExtension = "TRG"
         Case "EPRC":  iCommCount = GetSPCount(SourceConnString)                       
                       iSPName = "dbo.dmoScriptDatabase_PRC "
                       iSubFolder = "PRC"
                       iExtension = "PRC"
         Case "EVIW":  iCommCount = GetViewCount(SourceConnString)                       
                       iSPName = "dbo.dmoScriptDatabase_VIW "
                       iSubFolder = "VIW"
                       iExtension = "VIW"
         Case "EUDF":  iCommCount = GetUDFCount(SourceConnString)                       
                       iSPName = "dbo.dmoScriptDatabase_UDF "
                       iSubFolder = "UDF"
                       iExtension = "UDF"
    End Select
    
    If iCommCount = 0 Then
       Call ShowMessage("대상 객체가 한건도 없습니다.")
       'Call RefreshProgressBar(100,gpWidth,"N/A",pOpt)
       Call RefreshProgressBar(100,gpWidth,"N/A",pOpt,CalElaspeTime(0,0))
       'Response.End
    End If

    Select Case pOpt
         Case "ETAB":  iCommList = GetTableList(SourceConnString)
         Case "EFK":   iCommList = GetTableList(SourceConnString)
         Case "ETRG":  iCommList = GetTableList(SourceConnString)
         Case "EPRC":  iCommList = GetSPList   (SourceConnString)
         Case "EVIW":  iCommList = GetViewList (SourceConnString)
         Case "EUDF":  iCommList = GetUDFList  (SourceConnString)
    End Select    
    
    gtotObjectCount = gtotObjectCount + iCommCount
    
    IF iCommCount > 0 THEN
    intIncrement = gpWidth / iCommCount
    END IF

    If Response.Buffer Then Response.Flush
    
    iLoopContents = Split(iCommList, ",")
    
    iLoopCount = UBound(iLoopContents)
    
    Set AdoConn = Server.CreateObject("ADODB.Connection")
    adoConn.Open SourceConnString
    
    For iLoop = 0 To iLoopCount

        If iLoop = 0 Then
           iSQL = iSPName & " '" & DBName & "','" & WorkingDir & "','" & iLoopContents(iLoop) & "',1 "
        Else
           iSQL = iSPName & " '" & DBName & "','" & WorkingDir & "','" & iLoopContents(iLoop) & "',0 "
        End If
        adoConn.Execute iSQL
        
        iSQL = "dbo.dmoScriptGen '" & GetGlobalInf("gDBServerIP") & "','" & DeCode(Session("TDB")) & "','" & GetGlobalInf("gDBLoginID") & "','" & GetGlobalInf("gDBSAPwd") & "','" & WorkingDir & "\" & iSubFolder & "\" & iLoopContents(iLoop) & "." & iExtension & "'"
        If pOpt = "EVIW" Then
			adoConn.Execute iSQL
        End If
        adoConn.Execute iSQL

        Call RefreshProgressBar(CInt(iLoop * intIncrement * 100 / gpWidth), CInt(iLoop * intIncrement), iLoopContents(iLoop), pOpt, CalElaspeTime(bTimer, Timer))

    Next
    
    IF iCommCount > 0 THEN
    Call RefreshProgressBar(100, gpWidth, iLoopContents(iLoopCount), pOpt, CalElaspeTime(bTimer, Timer))
    END IF
    Call RefreshProgressResultBar(pOpt,iCommCount)
    
    
    adoConn.Close    
    Set adoConn = Nothing
    
End Sub


Sub RefreshProgressBar(ByVal ProgressBarPercentage,ByVal ProgressBarValue,ByVal TableName,ByVal pObject,ByVal pTime)
%>
    <SCRIPT LANGUAGE="VBS">
       Dim intWidth
       
       Select Case "<%=pObject%>"
           Case "EDB"  : document.all("txtState").innerText = "데이터베이스 생성중"
                         document.all("txtMDB").innerText   = "<%=pTime%>"
                         idCDB.style.backgroundColor        = "#F8D2D2"
           Case "ETAB" : document.all("txtState").innerText = "테이블 정보 추출중"
                         document.all("txtMTAB").innerText   = "<%=pTime%>"
                         idETAB.style.backgroundColor       = "#F8D2D2"
           Case "ETRG" : document.all("txtState").innerText = "트리거 정보 추출중"
                         document.all("txtMTRG").innerText   = "<%=pTime%>"
                         idETRG.style.backgroundColor       = "#F8D2D2"
           Case "EFK" : document.all("txtState").innerText  = "외부키 정보 추출중"
                         document.all("txtMFK").innerText   = "<%=pTime%>"
                         idEFK.style.backgroundColor        = "#F8D2D2"
           Case "EPRC" : document.all("txtState").innerText  = "저장 프로시져 정보 추출중"
                         document.all("txtMPRC").innerText   = "<%=pTime%>"
                         idEPRC.style.backgroundColor        = "#F8D2D2"
           Case "EVIW" : document.all("txtState").innerText  = "뷰 정보 추출중"
                         document.all("txtMVIW").innerText   = "<%=pTime%>"
                         idEVIW.style.backgroundColor        = "#F8D2D2"
           Case "EUDF" : document.all("txtState").innerText  = "사용자 정의 함수 정보 추출중"
                         document.all("txtMUDF").innerText   = "<%=pTime%>"
                         idEUDF.style.backgroundColor        = "#F8D2D2"
       End Select    
       
       document.all("divProgress").style.width =  <%= ProgressBarValue %>
       document.all("txtObjectName").innerText = "<%= TableName %>"
       document.all("txtPercentage").innerHTML = Right("     " & "<%= ProgressBarPercentage %>",6) & "%"
       
       if <%= ProgressBarPercentage %> = 100 Then
           document.all("chk<%=pObject%>").checked      = True
           document.all("divProgress").style.width =  0
           document.all("txtObjectName").innerText = ""
           document.all("txtPercentage").innerHTML = ""
           document.all("txtState").innerText  = ""
           Select Case "<%=pObject%>"
                 Case "EDB"  : idCDB.style.backgroundColor  = "#EFF1EF"
                 Case "ETAB" : idETAB.style.backgroundColor = "#EFF1EF"
                 Case "ETRG" : idETRG.style.backgroundColor = "#EFF1EF"
                 Case "EFK"  : idEFK.style.backgroundColor  = "#EFF1EF"
                 Case "EPRC" : idEPRC.style.backgroundColor = "#EFF1EF"
                 Case "EVIW" : idEVIW.style.backgroundColor = "#EFF1EF"
                 Case "EUDF" : idEUDF.style.backgroundColor = "#EFF1EF"
           End Select      
       End If
      
    </SCRIPT>
<%
  If Response.Buffer Then Response.Flush
  
End Sub

Sub RefreshProgressResultBar(ByVal pObject,ByVal pCount)
%>
    <SCRIPT LANGUAGE="VBS">
       Dim intWidth
       
       Select Case "<%=pObject%>"
           Case "EDB"  : document.all("txtMCDB").innerText   = "<%=pCount%>"
           Case "ETAB" : document.all("txtMCTAB").innerText  = "<%=pCount%>"
           Case "ETRG" : document.all("txtMCTRG").innerText  = "<%=pCount%>"
           Case "EFK"  : document.all("txtMCFK").innerText   = "<%=pCount%>"
           Case "EPRC" : document.all("txtMCPRC").innerText  = "<%=pCount%>"
           Case "EVIW" : document.all("txtMCVIW").innerText  = "<%=pCount%>"
           Case "EUDF" : document.all("txtMCUDF").innerText  = "<%=pCount%>"
       End Select    
       
    </SCRIPT>
<%
  If Response.Buffer Then Response.Flush
End Sub

Sub RefreshProgressTotResultBar(ByVal pCount, ByVal pTime)
%>
    <SCRIPT LANGUAGE="VBS">
       Dim intWidth
       document.all("txtMTOT").innerText   = "<%=pTime%>"
       document.all("txtMCTOT").innerText  = "<%=pCount%>"
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
		

<!--
<DIV ID="MousePT" NAME="MousePT" style='visibility:visible;    LEFT: expression((document.body.clientWidth-320)/2);    TOP: expression(document.body.clientHeight/2);' align=center>
	<table BORDER=1 width="320" border=1 cellpadding=1 cellspacing=1 bordercolor=#CCCCCC bordercolorlight=#CCCCCC bgcolor="buttonface" bordercolordark="#000000" vspace="0" hspace="0">
	<tr bgcolor="#CED3E7"> 
	<td bgcolor="#FFFFFF"><img src="../../image/net.gif" width="32" height="31" vspace="0" hspace="0" align="absmiddle">
	  <b>&nbsp;&nbsp;데이터 베이스를 생성하는 중입니다...</b></td>
	</tr>
	</table>
</DIV>

<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
-->
	
	
<body>
