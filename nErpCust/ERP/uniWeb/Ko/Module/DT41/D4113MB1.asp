<%@ LANGUAGE=VBSCript TRANSACTION=Required%>
<% Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgSvrVariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<% 											'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 
Call HideStatusWnd
'On Error Resume Next														'☜:
Err.Clear

Dim strMode																	'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 
Dim strSpread
Dim strBpCd

Dim LngMaxRow		' 현재 그리드의 최대Row
Dim LngRow

Call LoadBasisGlobalInf()

Call loadInfTB19029B("I", "B","NOCOOKIE","MB")

strMode = Request("txtMode")												'☜ : 현재 상태를 받음 
strBpCd = Request("txtBpCd")

Call SubOpenDB(lgObjConn)                                                        '☜: Make a DB Connection

Select Case strMode
	Case CStr(UID_M0001)
		Call SubBizQuery()
	Case CStr(UID_M0002)
		Call SubBizSave()
End Select

Call SubCloseDB(lgObjConn)

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
	On Error Resume Next                                                             '☜: Protect system from crashing
	Err.Clear                                                                        '☜: Clear Error status

	lgStrSQL =	"SELECT a.bp_cd, b.bp_nm, a.rev_flag, c.minor_nm, a.remarks"	& vbCrLf & _
					"  FROM dt_biz_partner_for_reverse a"					& vbCrLf & _
					"  LEFT JOIN b_biz_partner b"								& vbCrLf & _
					"			 ON a.bp_cd = b.bp_cd"							& vbCrLf & _
					"  LEFT JOIN b_minor c"										& vbCrLf & _
					"			 ON c.minor_cd = a.rev_flag"					& vbCrLf & _
					"			AND c.major_cd = 'DT005'"						& vbCrLf & _
					" WHERE a.bp_cd >= " & FilterVar(strBpCd, "''", "S")
	
	If FncOpenRs("R",lgObjConn, lgObjRs, lgStrSQL, "X", "X") = False Then
		Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
		Call SetErrorStatus()
	Else	
%>
<Script Language=vbscript>
	Dim LngMaxRow
	Dim strData
	Dim TmpBuffer
	Dim iTotalStr1
	Dim aaa

	With parent																	'☜: 화면 처리 ASP 를 지칭함 
		LngMaxRow = .frm1.vspdData.MaxRows										'Save previous Maxrow

		.ggoSpread.Source = .frm1.vspdData
		.ggoSpread.SSShowDataByClip <%Response.Write """"
		Dim iDx
		Do While Not lgObjRs.EOF
			Response.Write gColSep & ConvSPChars(lgObjRs("bp_cd")) & gColSep & gColSep & ConvSPChars(lgObjRs("bp_nm"))
			Response.Write gColSep & ConvSPChars(lgObjRs("rev_flag")) & gColSep & ConvSPChars(lgObjRs("minor_nm"))  
			Response.Write gColSep & ConvSPChars(lgObjRs("remarks")) & gColSep & iDx & gColSep & gRowSep

			lgObjRs.MoveNext
		Loop 
		Response.Write """"		%>

		.DbqueryOk
	End With
</Script>
<%
    End If
    
    Call SubCloseRs(lgObjRs)
End Sub	

'============================================================================================================
' Name : SubBizQuery
' Desc : Date data 
'============================================================================================================
Sub SubBizSave()
	Dim i
	Dim arrColumns
	
	On Error Resume Next
	Err.Clear

	strSpread = Split(Request("txtSpread"), gRowSep)

	For i = 0 To UBound(strSpread) - 1
		arrColumns = Split(strSpread(i), gColSep)
		
		Select Case arrColumns(0)
			Case "C"
				lgStrSQL = "INSERT dt_biz_partner_for_reverse" & vbCrLf & _
							  "		 (bp_cd, rev_flag, remarks, insert_user_id, insert_date, update_user_id,update_date)" & vbCrLf & _
							  "VALUES (" & Filtervar(arrColumns(2), "''", "S") & "," _
											 & Filtervar(arrColumns(3), "''", "S") & "," _
											 & Filtervar(arrColumns(4), "''", "S") & "," _
											 & Filtervar(trim(gUsrId),"''","S") & ", GETDATE()," _
											 & Filtervar(trim(gUsrId),"''","S") & ", GETDATE())"
			Case "U"
				lgStrSQL = "UPDATE dt_biz_partner_for_reverse" & vbCrLf & _
							  "	SET rev_flag = " & Filtervar(arrColumns(3), "''", "S") & "," & vbCrLf & _
							  "		 remarks = " & Filtervar(arrColumns(4), "''", "S") & "," & vbCrLf & _
							  "		 update_user_id =" & Filtervar(trim(gUsrId),"''","S") & "," & vbCrLf & _
							  "		 update_date = GETDATE()" & vbCrLf & _
							  " WHERE bp_cd = " & Filtervar(arrColumns(2), "''", "S") 
			Case "D"
				lgStrSQL = "DELETE dt_biz_partner_for_reverse" & vbCrLf & _
							  " WHERE bp_cd = " & Filtervar(arrColumns(2), "''", "S")
		End Select
		
		lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords

		If CheckSYSTEMError(Err, True) = True Then
			lgErrorStatus = "YES"
			ObjectContext.SetAbort
		End If
	Next
End Sub
	    
'============================================================================================================
' Name : CommonOnTransactionCommit
' Desc : This Sub is called by OnTransactionCommit Error handler
'============================================================================================================
Sub CommonOnTransactionCommit()
End Sub

'============================================================================================================
' Name : CommonOnTransactionAbort
' Desc : This Sub is called by OnTransactionAbort Error handler
'============================================================================================================
Sub CommonOnTransactionAbort()
    lgErrorStatus = "YES"
End Sub
%>

<Script Language="VBScript">
    Select Case "<%=strMode %>"
       Case "<%=UID_M0001%>"
          Parent.DBQueryOk
       Case "<%=UID_M0002%>"
<%	If lgErrorStatus <> "YES" Then%>
          Parent.DBSaveOk
<%	End If%>
    End Select
</Script>