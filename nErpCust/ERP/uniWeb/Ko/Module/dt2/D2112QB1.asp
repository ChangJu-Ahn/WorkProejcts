<%'======================================================================================================
'*  1. Module Name          : Purchase
'*  2. Function Name        : 
'*  3. Program ID           : m5314mb1.asp
'*  4. Program Name         : 전자세금계산서
'*  5. Program Desc         : 
'*  6. Modified date(First) : 2009-07-07
'*  7. Modified date(Last)  : 2009-07-07
'*  8. Modifier (First)     : Lee Min Hyung
'*  9. Modifier (Last)      : Lee Min Hyung
'* 10. Comment              : 
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================
%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp"		-->
<!-- #Include file="../../inc/adovbs.inc"				-->
<!-- #Include file="../../inc/lgSvrVariables.inc"	-->
<!-- #Include file="../../inc/incServeradodb.asp"	-->
<!-- #Include file="../../inc/incSvrDate.inc"		-->
<!-- #Include file="../../inc/incSvrNumber.inc"		-->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%
Call LoadBasisGlobalInf
Call LoadInfTB19029B("Q", "M", "NOCOOKIE","MB")
Call LoadBNumericFormatB("Q","M", "NOCOOKIE", "MB")

Dim ADF												'ActiveX Data Factory 지정 변수선언 
Dim strRetMsg										'Record Set Return Message 변수선언 
Dim UNISqlId, UNIValue, UNILock, UNIFlag	'DBAgent Parameter 선언 
Dim rs0					                     'DBAgent Parameter 선언 
Dim strMode											'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 

Call HideStatusWnd

Dim StrSupplierCd
Dim StrcbobillStatus
Dim StrhdtxtRadio
Dim StrcboTransferStatus
Dim strIssuedFromDt
Dim strIssuedToDt
Dim i

	Call SubOpenDB(lgObjConn)

	lgStrSQL =	"SELECT a.inv_type, a.inv_no, a.dt_inv_no, a.process_date," & vbCrLf & _
					"		  a.success_flag, a.error_desc, a.re_flag, c.minor_nm," & vbCrLf & _
					"		  a.change_remark, a.change_remark2, a.change_remark3," & vbCrLf & _
					"		  a.insert_user_id, b.usr_nm user_name" & vbCrLf & _
					"  FROM dt_simple_issue_transmit_log a" & vbCrLf & _
					"  LEFT JOIN z_usr_mast_rec b" & vbCrLf & _
					"			 ON a.insert_user_id = b.usr_id" & vbCrLf & _
					"  LEFT JOIN b_minor c" & vbCrLf & _
					"			 ON c.minor_cd = a.change_reason" & vbCrLf & _
					"			AND c.major_cd = 'DT006'" & vbCrLf & _
					" WHERE a.inv_type LIKE " & FilterVar(Request("txtJobType"), "''", "S") & vbCrLf & _
					"	 AND convert(varchar(10), a.process_date, 21)  BETWEEN " & FilterVar(Request("txtFromDate"), "''", "S") & vbCrLf & _
					"									AND " & FilterVar(Request("txtToDate"), "''", "S") & vbCrLf & _
					"	 AND a.insert_user_id LIKE " & FilterVar(Request("txtERPUser"), "''", "S") & vbCrLf & _
					"	 AND a.inv_no LIKE " & FilterVar("%" & Request("txtINVNo") & "%", "''", "S") & vbCrLf & _
					"ORDER BY a.process_date DESC"

	If FncOpenRs("R", lgObjConn, lgObjRs, lgStrSQL, "X", "X") = False Then
		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
		Set lgObjRs = Nothing
		Call SubCloseDB(lgObjConn)
		Response.End													'☜: 비지니스 로직 처리를 종료함 
	End If
%>

<Script Language=vbscript>
	Dim LngMaxRow
	Dim strData
	Dim TmpBuffer
	Dim iTotalStr1

	With parent																	'☜: 화면 처리 ASP 를 지칭함 
		LngMaxRow = .frm1.vspdData.MaxRows								'Save previous Maxrow
		ReDim TmpBuffer(100)
<%		Dim iDx
		iDx = 0
		Do While Not lgObjRs.EOF %>
			strData = Chr(11) & "<%=ConvSPChars(lgObjRs("inv_type"))%>" & _
						 Chr(11) & "<%=ConvSPChars(lgObjRs("inv_no"))%>" & _
						 Chr(11) & "<%=ConvSPChars(lgObjRs("dt_inv_no"))%>" & _
						 Chr(11) & "<%=ConvSPChars(lgObjRs("process_date"))%>" & _
						 Chr(11) & "<%=ConvSPChars(lgObjRs("success_flag"))%>" & _
						 Chr(11) & "<%=ConvSPChars(lgObjRs("re_flag"))%>" & _
						 Chr(11) & "<%=ConvSPChars(lgObjRs("error_desc"))%>" & _
						 Chr(11) & "<%=ConvSPChars(lgObjRs("minor_nm"))%>" & _
						 Chr(11) & "<%=ConvSPChars(lgObjRs("change_remark"))%>" & _
						 Chr(11) & "<%=ConvSPChars(lgObjRs("change_remark2"))%>" & _
						 Chr(11) & "<%=ConvSPChars(lgObjRs("change_remark3"))%>" & _
						 Chr(11) & "<%=ConvSPChars(lgObjRs("insert_user_id"))%>" & _
						 Chr(11) & "<%=ConvSPChars(lgObjRs("user_name"))%>" & _
						 Chr(11) & LngMaxRow + <%=iDx%> & _
						 Chr(11) & Chr(12)

			TmpBuffer(<%=iDx%>) = strData
<%			lgObjRs.MoveNext
			iDx = iDx + 1
		Loop %>

		iTotalStr1 = Join(TmpBuffer, "")
		.ggoSpread.Source = .frm1.vspdData
		.ggoSpread.SSShowDataByClip iTotalStr1

<%		lgObjRs.Close
		Set lgObjRs = Nothing	%>

		.DbQueryOk()
	End With
</Script>	
<%	Call SubCloseDB(lgObjConn)	%>