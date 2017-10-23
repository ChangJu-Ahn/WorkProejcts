<%@LANGUAGE = VBScript%>
<%'********************************************************************************************
'*  1. Module Name			: DT
'*  2. Function Name		: 
'*  3. Program ID			: d1211PB1.asp
'*  4. Program Name			: Digital Tax (Query)
'*  5. Program Desc			:
'*  6. Comproxy List		: DB Agent
'*  7. Modified date(First)	: 2009/12/20
'*  8. Modified date(Last)	: 2009/12/22
'*  9. Modifier (First)		: Chen, Jae Hyun
'* 10. Modifier (Last)		: Chen, Jae Hyun
'* 11. Comment				:
'********************************************************************************************
%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%
Call LoadBasisGlobalInf
Call LoadInfTB19029B("Q", "M","NOCOOKIE", "RB")

Dim ADF										'ActiveX Data Factory 지정 변수선언 
Dim strRetMsg								'Record Set Return Message 변수선언 
Dim UNISqlId, UNIValue, UNILock, UNIFlag	'DBAgent Parameter 선언 
Dim rs1										'DBAgent Parameter 선언 
Dim strMode									'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 
Dim i
Dim strFlag
Dim strInvNo 

'=======================================================================================================
'	아래 선언되어 있는 변수들은 COOL:Gen 의 Record Return Count 의 제한에 따른 것이다.
'	따라서, ADO를 사용할 경우 그와같은 문제성이 없기 때문에 아래의 변수들은 사용하지 않지만 추후 
'	uniERP2000 에서 한번에 조회되는 Record Count 의 수를 30으로 제한하고 있는 만큼 그에 따른 
'	표준은 동시에 추가될 예정이므로 변수삭제는 하지 않고 그대로 놔둔다.
'=======================================================================================================

Call HideStatusWnd

strMode = Request("txtMode")												'☜ : 현재 상태를 받음 
strInvNo = Request("txtInvNo")

On Error Resume Next
Err.Clear
																	'☜: Protect system from crashing
	
	Redim UNISqlId(0)
	Redim UNIValue(0, 0)
	
	UNISqlId(0) = "D1211PA1"
	
	UNIValue(0, 0) = FilterVar(strInvNo, "''", "S")
	
	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs1)
	
	If (rs1.EOF And rs1.BOF) Then
		Call DisplayMsgBox("189200", vbOKOnly, "", "", I_MKSCRIPT)
		rs1.Close
		Set rs1 = Nothing
		Set ADF = Nothing
		Response.End
	End If
	
	'// QUERY REWORK ORDER HISTORY
	
	
%>
<Script Language=vbscript>
	Dim TmpBuffer1
    Dim iTotalStr
    Dim LngMaxRow
    Dim strData
	
    With parent												'☜: 화면 처리 ASP 를 지칭함 
		
	 	LngMaxRow = .vspdData.MaxRows							'Save previous Maxrow
		
<%  
		If Not(rs1.EOF And rs1.BOF) Then
%>	
		
			
			Redim TmpBuffer1(<%=rs1.RecordCount-1%>)
<%		
			For i=0 to rs1.RecordCount-1
%>
				strData = ""
				strData = strData & Chr(11) & "<%=ConvSPChars(rs1("inv_no"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs1("wrk_dtm"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs1("proc_flag"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs1("proc_flag_nm"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs1("proc_usr_id"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs1("proc_usr_name"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs1("attr02"))%>"
				strData = strData & Chr(11) & "<%=UniNumClientFormat(rs1("sup_tot_amt"),ggAmtOfMoney.DecPoint,0)%>"
				strData = strData & Chr(11) & "<%=UniNumClientFormat(rs1("sur_tax"),ggAmtOfMoney.DecPoint,0)%>"
				strData = strData & Chr(11) & LngMaxRow + <%=i%>
				strData = strData & Chr(11) & Chr(12)
				
				TmpBuffer1(<%=i%>) = strData
<%		
				rs1.MoveNext
				
			Next
%>
			
		iTotalStr = Join(TmpBuffer1,"") 
		.ggoSpread.Source = .vspdData
		.ggoSpread.SSShowDataByClip iTotalStr

<%	
		End If
		

		rs1.close

		Set rs1 = Nothing

%>	
		
		.DbQueryOk(LngMaxRow)
		
    End With
</Script>	
<%    
    Set ADF = Nothing
%>
