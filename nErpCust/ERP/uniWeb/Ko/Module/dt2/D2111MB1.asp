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
Dim strUserId
Dim strSpread

Dim LngMaxRow		' 현재 그리드의 최대Row
Dim LngRow

Call LoadBasisGlobalInf()

Call loadInfTB19029B("I", "B","NOCOOKIE","MB")

strMode = Request("txtMode")												'☜ : 현재 상태를 받음 
strUserId = Request("txtUserId")

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

	lgStrSQL =  "SELECT a.user_id, b.usr_nm user_name, " & vbCrLf & _
                "       A.DT_ID, A.DT_PW, " & vbCrLf & _
                "       A.USER_DN, A.USER_INFO " & vbCrLf & _
                "  FROM DT_USER_INFO A (NOLOCK) " & vbCrLf & _
                "  LEFT JOIN Z_USR_MAST_REC B (NOLOCK) ON A.USER_ID = B.USR_ID " & vbCrLf & _
                " WHERE A.USER_ID >= " & FilterVar(strUserId, "''", "S")

    If FncOpenRs("R",lgObjConn, lgObjRs, lgStrSQL, "X", "X") = False Then
        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found.
        Call SetErrorStatus()
    Else    %>
<Script Language=vbscript>
    Dim LngMaxRow
    Dim strData
    Dim TmpBuffer
    Dim iTotalStr1
    Dim aaa

    With parent                                                                 '☜: 화면 처리 ASP 를 지칭함
        LngMaxRow = .frm1.vspdData.MaxRows                                      'Save previous Maxrow

        .ggoSpread.Source = .frm1.vspdData
        .ggoSpread.SSShowDataByClip <%Response.Write """"
        Dim iDx
        Do While Not lgObjRs.EOF
            Response.Write gColSep & ConvSPChars(lgObjRs("USER_ID")) & gColSep & gColSep & ConvSPChars(lgObjRs("USER_NAME"))
            Response.Write gColSep & ConvSPChars(lgObjRs("DT_ID")) & gColSep & ConvSPChars(lgObjRs("DT_PW"))
            Response.Write gColSep & iDx & gColSep & gRowSep

            lgObjRs.MoveNext
        Loop
        Response.Write """" %>

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
				lgStrSQL = "INSERT DT_USER_INFO " & vbCrLf & _
                       "      (USER_ID, DT_ID, DT_PW, insert_user_id, insert_date, update_user_id,update_date)" & vbCrLf & _
                       "VALUES (" & Filtervar(arrColumns(2), "''", "S") & "," _
                              	 & Filtervar(arrColumns(3), "''", "S") & "," _
                                	 & Filtervar(arrColumns(4), "''", "S") & "," _
                                  & Filtervar(trim(gUsrId),"''","S") & ", GETDATE()," _
                                	 & Filtervar(trim(gUsrId),"''","S") & ", GETDATE())"
			Case "U"
				lgStrSQL = "UPDATE DT_USER_INFO " & vbCrLf & _
                       "   SET DT_ID = " & Filtervar(arrColumns(3), "''", "S") & "," 		& vbCrLf & _
                       "       DT_PW = " & Filtervar(arrColumns(4), "''", "S") & "," 		& vbCrLf & _
                       "       update_user_id =" & Filtervar(trim(gUsrId),"''","S") & "," & vbCrLf & _
                       "       update_date = GETDATE()" 												& vbCrLf & _
                       " WHERE user_id = " & Filtervar(arrColumns(2), "''", "S")
			Case "D"
				lgStrSQL = "DELETE DT_USER_INFO " 															& vbCrLf & _
                       " WHERE user_id = " & Filtervar(arrColumns(2), "''", "S")
		End Select
		
		lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords

		If CheckSYSTEMError(Err, True) = True Then
			Response.Write "<Script LANGUAGE=VBScript>" & vbCrLf
			Response.Write "	parent.UNIMsgBox ""오류가 발생했습니다. 데이터를 확인해 보십시요."", 16, ""uniERPII""" & vbCrLf
			Response.Write "</Script>" & vbCrLf
			'Call DisplayMsgBox("205921", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
			Call SetErrorStatus()

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