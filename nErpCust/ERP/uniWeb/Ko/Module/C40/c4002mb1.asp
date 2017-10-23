<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% Option Explicit%>
<% session.CodePage=949 %>

<%'======================================================================================================
'*  1. Module Name          : Cost Accounting
'*  2. Function Name        : Distribution Factor
'*  3. Program ID           : c1701mb1
'*  4. Program Name         : 실제원가 배부요소 정보 등록 
'*  5. Program Desc         : 실제원가 배부요소 정보 
'*  6. Modified date(First) : 2000/11/08
'*  7. Modified date(Last)  : 2002/06/18
'*  8. Modifier (First)     : 강창구 
'*  9. Modifier (Last)      : cho Ig Sung/ Park, Joon-Won
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================

Response.Buffer = True								'☜ : ASP가 버퍼에 저장되지 않고 바로 Client에 내려간다.
%>

<!-- #Include file="../../inc/IncSvrMain.asp"  -->

<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->

<%													'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 

	Call LoadBasisGlobalInf()
	On Error Resume Next								'☜: 
	Err.Clear

	Call HideStatusWnd
'---------------------------------------Common-----------------------------------------------------------
    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '☜: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    
    'Multi SpreadSheet
	lgLngMaxRow       = Request("txtMaxRows")                                        '☜: Read Operation Mode (CRUD)

'------ Developer Coding part (Start ) ------------------------------------------------------------------

'------ Developer Coding part (End   ) ------------------------------------------------------------------ 
    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '☜: Query 
			Call SubBizQueryMulti()     				     
        Case CStr(UID_M0002)                                                         '☜: Save,Update
             Call SubBizSaveMulti()
        Case CStr(UID_M0003)                                                         '☜: Delete
             Call SubBizDelete()
    End Select

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

End Sub    
'============================================================================================================
' Name : SubBizSave
' Desc : Save Data 
'============================================================================================================
Sub SubBizSave()
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
End Sub
'============================================================================================================
' Name : SubBizDelete
' Desc : Delete DB data
'============================================================================================================
Sub SubBizDelete()
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

	Dim iPC4G002Q
	
	If Trim(Request("txtVER_CD")) = "" Then
		Call DisplayMsgBox("970021", vbInformation, "txtVER_CD", "", I_MKSCRIPT)      '☜ : No data is found. 
		Exit Sub
	End If
    
	Set iPC4G002Q = Server.CreateObject("PC4G002.cMngCostCenterRuleSvr")

    If CheckSYSTEMError(Err, True) = True Then
		Exit Sub
    End If    
	
	Call iPC4G002Q.C_ALL_DELETE(gStrGloBalCollection, Trim(Request("txtVER_CD")))
	
    
    If CheckSYSTEMError(Err, True) = True Then					
       Call SetErrorStatus  
       Set iPC4G002Q = Nothing
       Exit Sub
       
    End If    

    
    Set iPC4G002Q = Nothing
	    
End Sub


'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
	
    Dim iPC4G002Q
    Dim iStrData, iStrData2
 
    Dim arrRetData
    Dim iLngRow,iLngCol
    
   	Dim oRs, sTxt, arrRows, sNextKey, sRowSeq, iLngRowCnt, iLngColCnt', sGrpTxt
	Dim sVerCd
	Dim TmpBuffer()
	   
    Const C_VerCd = 0
    
    On Error Resume Next                                                                 '☜: Protect system from crashing
    Err.Clear                                                                            '☜: Clear Error status
	
	sVerCd=Trim(request("txtVer_Cd"))
	
	sNextKey	= Trim(Request("lgStrPrevKey"))

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
If Request("WhoQuery")="H" Then

Call SubCreateCommandObject(lgObjComm)    

 With lgObjComm
		.CommandTimeout = 0
	
		.CommandText = "dbo.usp_C_C42002MA1_HDR"		' --  변경해야할 SP 명 
	    .CommandType = adCmdStoredProc
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("RETURN_VALUE",  adInteger,adParamReturnValue)	' -- No 수정 
		' -- 변경해야할 조회조건 파라메타들 
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@VER_CD",	adVarXChar,	adParamInput, 3,sVerCd)	
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@SHEETCNT",  adSmallInt, adParamInput,, 100)	
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@NEXTKEY",	adVarXChar,	adParamInput, 15,sNextKey)
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@DEBUG",  adSmallInt, adParamInput,, 0)	' -- isqlw 에서만 사용하는 디버깅코드 
		    
        Set oRs = lgObjComm.Execute
        
    End With
    'Response.Write "Err=" & Err.Description
    If Instr( Err.Description , "B_MESSAGE") > 0 Then
		If HandleBMessageError(vbObjectError, Err.Description, "", "") = True Then
			Exit Sub
		End If
	Else
		If CheckSYSTEMError(Err, True) = True Then	
			Exit Sub
		End If
	End If
   'Response.End     
	If oRs.EoF and oRs.Bof then
		If sNextKey="" then
			Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
			oRs.Close
			Set oRs = Nothing
			Exit Sub
		Else
			Response.Write " <Script Language=vbscript>	                        " & vbCr
			Response.Write " With parent                                        " & vbCr
			Response.Write "	.lgStrPrevKey = """"" & vbCr 	
			Response.Write " End With                                        " & vbCr
			Response.Write  " </Script>                  " & vbCr
			oRs.Close
			Set oRs = Nothing
			Exit Sub		
		End If
	End If
	
		If Not oRs  is nothing Then
				
			arrRows = oRs.GetRows()
			iLngRowCnt = UBound(arrRows,2) 
			iLngColCnt	= UBound(arrRows, 1) 
			Redim TmpBuffer(iLngRowCnt)

			sRowSeq = arrRows(UBound(arrRows, 1), iLngRowCnt)
			
			
	
			
						
			For iLngRow = 0 To iLngRowCnt
				iStrData = ""
				For iLngCol = 0 To iLngColCnt
					If iLngCol = 3 Then
						iStrData = iStrData & Chr(11) & ConvSPChars(Trim(arrRows(iLngCol,iLngRow )))
						iStrData = iStrData & Chr(11) & ""
				    ElseIF iLngCol = 2 or iLngCol = 5 or iLngCol = 6 or iLngCol = 8 or iLngCol = 11 or iLngCol = 14 or iLngCol = 17 or iLngCol = 20 or iLngCol = 23 or iLngCol = 26 Then
						iStrData = iStrData & Chr(11) & ConvSPChars(arrRows(iLngCol,iLngRow))
						iStrData = iStrData & Chr(11) & ""		
				    ElseIF iLngCol = 13 or iLngCol = 16 or iLngCol = 19 or iLngCol = 22 or iLngCol = 25 Then
						iStrData = iStrData & Chr(11) & CDBL(arrRows(iLngCol,iLngRow))
					ELSEIf   iLngCol = 4 then 
						iStrData = iStrData & Chr(11) & ConvLang(ConvSPChars(arrRows(iLngCol,iLngRow))	)		
					ELSEIf   iLngCol = 10 then						
						iStrData = iStrData & Chr(11) & ConvLang(ConvSPChars(arrRows(iLngCol,iLngRow))	)
						
						IF ConvLang(ConvSPChars(arrRows(iLngCol,iLngRow))) = "*" Then
							iStrData = iStrData & Chr(11) & "*"
						ELSEIF ConvLang(ConvSPChars(arrRows(iLngCol,iLngRow))) = "D" Then
							iStrData = iStrData & Chr(11) & "직접"
						ELSE
							iStrData = iStrData & Chr(11) & "간접"
						END IF						

						
					Else					
						iStrData = iStrData & Chr(11) & ConvSPChars(arrRows(iLngCol,iLngRow))	
					END IF	
				Next

				'iStrData = iStrData & Chr(11) & ConvSPChars(arrRows(iLngCol,iLngRow))
				iStrData = iStrData & Chr(11) & Chr(12)

				TmpBuffer(iLngRow) = iStrData
			Next
			iStrData = Join(TmpBuffer, "")
		End IF
		Call SubCloseCommandObject(lgObjComm)	
	Else
		Set iPC4G002Q = Server.CreateObject("PC4G002.cListCostCenterRuleSvr")
		
		If CheckSYSTEMError(Err, True) = True Then
			Exit Sub
		End If    
	
		If Trim(Request("txtVER_CD")) = "" Then
			Call DisplayMsgBox("970021", vbInformation, "txtVER_CD", "", I_MKSCRIPT)      '☜ : No data is found. 
			Exit Sub
		End If
	
		Call iPC4G002Q.C_LIST_DSTB_RULE_BY_CC_S_SVR(gStrGloBalCollection,sVerCd, arrRetData, Request("WhoQuery"), Request("lgStrPrevKey"), Request("SeqNo"))
	    
		If CheckSYSTEMError(Err, True) = True Then					
		   Call SetErrorStatus  
		   Set iPC4G002Q = Nothing
		   Exit Sub       
		End If    
    
		Set iPC4G002Q = Nothing
    
		If Not isArray(arrRetData) Then Response.End    
 	
			iIntLoopCount = 0	
			If isArray(arrRetData) Then
				
				iLngRowCnt = UBound(arrRetData, 1) 
				Redim TmpBuffer(iLngRowCnt)
				
				For iLngRow = 0 To iLngRowCnt
					iStrData2 = ""
					
					For iLngCol = 0 To UBound(arrRetData, 2)
						If iLngCol = 3 Then
							iStrData2 = iStrData2 & Chr(11) & ConvSPChars(Trim(arrRetData(iLngRow, iLngCol)))
							iStrData2 = iStrData2 & Chr(11) & ""
					    ElseIF iLngCol = 2 Then
							iStrData2 = iStrData2 & Chr(11) & ConvSPChars(Trim(arrRetData(iLngRow, iLngCol)))
							iStrData2 = iStrData2 & Chr(11) & ""
						ELSEIF iLngCol = 4 Then 
							iStrData2 = iStrData2 & Chr(11) & ConvLang(ConvSPChars(Trim(arrRetData(iLngRow, iLngCol))))
						ELSE
							iStrData2 = iStrData2 & Chr(11) & ConvSPChars(Trim(arrRetData(iLngRow, iLngCol)))
						END IF	
					Next
					iStrData2 = iStrData2 & Chr(11) & iLngRow+1
					iStrData2 = iStrData2 & Chr(11) & Chr(12)
					
					TmpBuffer(iLngRow) = iStrData2
				Next
				
				iStrData2 = Join(TmpBuffer, "")
			Else
				iStrData2 = ""
			End If

	End If
			
	Response.Write " <Script Language=vbscript>	                        " & vbCr
	Response.Write " With parent                                        " & vbCr

	If Request("WhoQuery") = "H" Then

		Response.Write "	.frm1.vspdData.ReDraw = False					" & vbCr 			 
		Response.Write "	.frm1.hVerCd.value = 	""" & Trim(Request("txtVER_CD"))       & """" & vbCr
		Response.Write "	.lgStrPrevKey = 	""" & sRowSeq       & """" & vbCr

		Response.Write "	.ggoSpread.Source = .frm1.vspdData              " & vbCr 			 
		Response.Write "	.ggoSpread.SSShowData """ & iStrData       & """" & vbCr
		Response.Write "	.frm1.vspdData.ReDraw = True					" & vbCr 			 
	Else
	
		Response.Write "	.frm1.vspdData2.ReDraw = False					" & vbCr 			 
		Response.Write "	.ggoSpread.Source = .frm1.vspdData2              " & vbCr 			 
		Response.Write "	.ggoSpread.SSShowData """ & iStrData2       & """" & vbCr
		Response.Write "	.frm1.vspdData2.ReDraw = True					" & vbCr 			 
	End If    
    
    Response.Write " End With   " & vbCr
    Response.Write " </Script>  " & vbCr 

    If Request("WhoQuery") <> "H" Then
		Response.End 
    End If
    
End Sub    	 

Function ConvLang(Byval pLang)
	Dim pTmp
	
	pTmp = Replace(pLang , "%1", "비제조")
	pTmp = Replace(pTmp , "%2", "제조")
	ConvLang = pTmp
End Function


'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()

	On Error Resume Next                                                             '☜: Protect system from crashing
	Err.Clear                                                                        '☜: Clear Error status
  
    Dim PC4G002Data
    Dim sSQLI1, sSQLI2, sSQLU1, sSQLU2, sSQLD1, sSQLD2
    Dim iErrPosition, sErrMesg
    

    
    sSQLI1    = Trim(Request("txtSpreadI1"))
    sSQLU1    = Trim(Request("txtSpreadU1"))
    sSQLD1    = Trim(Request("txtSpreadD1"))
    sSQLI2    = Trim(Request("txtSpreadI2"))
    sSQLU2    = Trim(Request("txtSpreadU2"))
    sSQLD2    = Trim(Request("txtSpreadD2"))

    


	If sSQLI1 = "" And sSQLU1 = "" And sSQLD1 = "" And sSQLI2 = "" And sSQLU2 = "" And sSQLD2 = ""  Then
		Call DisplayMsgBox("970021", vbInformation, "txtSpreadI1~2", "", I_MKSCRIPT)      '☜ : No data is found. 
		Exit Sub
	End If
	
    
    Set PC4G002Data = Server.CreateObject("PC4G002.cMngCostCenterRuleSvr")

    If CheckSYSTEMError(Err, True) = True Then					
       Call SetErrorStatus
       Exit Sub
    End If    

    Call PC4G002Data.C_MANAGER_CC_RULE_SVR(gStrGlobalCollection, sSQLI1, sSQLU1, sSQLD1, sSQLI2, sSQLU2, sSQLD2, lgErrorPos)			
	
	sErrMesg = Err.Description

		
	If CheckSYSTEMError(Err, True ) = True Then
		If Instr(1, sErrMesg, "B_MESSAGE" & Chr(11) & "236307") > 0 Then
			Response.Write "<script language=vbscript>" & vbCr
			Response.Write "Call parent.OpenErrorPopup(""" & Replace(lgErrorPos, chr(34), chr(34) & chr(34)) & """)" & vbCr
			Response.Write "</script>" & vbCr
			Response.End 
		End If

		Call SetErrorStatus
		Set PC4G002Data = Nothing
		Exit Sub
	End If    
    
	Set PC4G002Data = Nothing

End Sub    

'============================================================================================================
' Name : SubBizSaveCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveMultiCreate(arrColVal)
   On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    
    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to create record
    '--------------------------------------------------------------------------------------------------------
    
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
    
End Sub

'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to update record
    '--------------------------------------------------------------------------------------------------------
    
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
End Sub

'============================================================================================================
' Name : SubBizSaveMultiDelete
' Desc : Delete Data from Db
'============================================================================================================
Sub SubBizSaveMultiDelete(arrColVal)

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to update record
    '--------------------------------------------------------------------------------------------------------

    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
End Sub

'============================================================================================================
' Name : CommonOnTransactionCommit
' Desc : This Sub is called by OnTransactionCommit Error handler
'============================================================================================================
Sub CommonOnTransactionCommit()
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

'============================================================================================================
' Name : CommonOnTransactionAbort
' Desc : This Sub is called by OnTransactionAbort Error handler
'============================================================================================================
Sub CommonOnTransactionAbort()
    lgErrorStatus    = "YES"
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

'============================================================================================================
' Name : SetErrorStatus
' Desc : This Sub set error status
'============================================================================================================
Sub SetErrorStatus()
    lgErrorStatus     = "YES"                                                         '☜: Set error status
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

%>

<Script Language="VBScript">
    
	With Parent
    Select Case "<%=lgOpModeCRUD %>"
       Case "<%=UID_M0001%>"                                                         '☜ : Query
          If Trim("<%=lgErrorStatus%>") = "NO" And "<%=Request("WhoQuery")%>" = "H" Then
            .DBQueryOk        
          ElseIf Trim("<%=lgErrorStatus%>") = "NO" And "<%=Request("WhoQuery")%>" <> "H" Then
            .DBQueryOk2        
          End If   
       Case "<%=UID_M0002%>"                                                         '☜ : Save
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DBSaveOk
          Else
            Parent.SubSetErrPos(Trim("<%=lgErrorPos%>"))
          End If   
       Case "<%=UID_M0003%>"                                                         '☜ : Delete
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DbDeleteOk
          Else   
          End If   
    End Select    
	End with
</Script>	
<%Response.End%>