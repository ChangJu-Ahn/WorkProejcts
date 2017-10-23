<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% Option Explicit%>
<% session.CodePage=949 %>

<%'======================================================================================================
'*  1. Module Name          : Cost Accounting
'*  2. Function Name        : Distribution Factor
'*  3. Program ID           : c4008mb1
'*  4. Program Name         : 직과항목등록 
'*  5. Program Desc         : 직과항목등록 
'*  6. Modified date(First) : 2000/11/08
'*  7. Modified date(Last)  : 2002/06/18
'*  8. Modifier (First)     : choe0tae 
'*  9. Modifier (Last)      : choe0tae
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================

Response.Buffer = True								'☜ : ASP가 버퍼에 저장되지 않고 바로 Client에 내려간다.
%>

<!-- #Include file="../../inc/IncSvrMain.asp"  -->

<%													'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 

	Call LoadBasisGlobalInf()
	
	Dim lgErrorStatus, lgErrorPos, lgOpModeCRUD 
    Dim lgLngMaxRow
	
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
	    
End Sub


'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
	
    Dim iPC4G008Q
    Dim iStrData, iStrData2
 
    Dim arrHeadData, arrDetailData
    Dim iLngRow,iLngCol
        
    Const C_VerCd = 0
    
    On Error Resume Next                                                                 '☜: Protect system from crashing
    Err.Clear                                                                            '☜: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------


	Set iPC4G008Q = Server.CreateObject("PC4G008.cListDirDstbFctrSvr")

    If CheckSYSTEMError(Err, True) = True Then
		Exit Sub
    End If    
	
	' -- minor_cd 는 필수입력이 아님.
	Call iPC4G008Q.C_LIST_DIR_DSTB_FCTR_S_SVR(gStrGloBalCollection, Trim(Request("txtMINOR_CD")), arrHeadData, arrDetailData)
    
    If CheckSYSTEMError(Err, True) = True Then					
       Call SetErrorStatus  
       Set iPC4G008Q = Nothing
       Exit Sub
       
    End If    

    
    Set iPC4G008Q = Nothing
	
	iStrData = ""
	iIntLoopCount = 0	
	For iLngRow = 0 To UBound(arrHeadData, 1) 		
		For iLngCol = 0 To UBound(arrHeadData, 2)
			iStrData = iStrData & Chr(11) & ConvSPChars(Trim(arrHeadData(iLngRow, iLngCol)))
		Next
		iStrData = iStrData & Chr(11) & iLngRow+1
		iStrData = iStrData & Chr(11) & Chr(12)
	Next

	iStrData2 = ""
	iIntLoopCount = 0	
	If isArray(arrDetailData) Then
		For iLngRow = 0 To UBound(arrDetailData, 1) 		
			For iLngCol = 0 To UBound(arrDetailData, 2)
				iStrData2 = iStrData2 & Chr(11) & ConvSPChars(Trim(arrDetailData(iLngRow, iLngCol)))
			Next
			iStrData2 = iStrData2 & Chr(11) & iLngRow+1
			iStrData2 = iStrData2 & Chr(11) & Chr(12)
		Next
	End If
		
	Response.Write " <Script Language=vbscript>	                        " & vbCr
	Response.Write " With parent                                        " & vbCr
	Response.Write "	.frm1.hMINOR_CD.value = 	""" & Trim(Request("txtMINOR_CD"))       & """" & vbCr
	Response.Write "	.frm1.vspdData.ReDraw = False					" & vbCr 			 
	Response.Write "	.frm1.vspdData2.ReDraw = False					" & vbCr 			 
    Response.Write "	.ggoSpread.Source = .frm1.vspdData              " & vbCr 			 
    Response.Write "	.ggoSpread.SSShowData """ & iStrData       & """" & vbCr
    Response.Write "	.ggoSpread.Source = .frm1.vspdData2              " & vbCr 			 
    If isArray(arrDetailData) Then
		Response.Write "	.ggoSpread.SSShowData """ & iStrData2       & """" & vbCr
	End If
	Response.Write "	.frm1.vspdData.ReDraw = True					" & vbCr 			 
	Response.Write "	.frm1.vspdData2.ReDraw = True					" & vbCr 			 
    
    Response.Write " End With   " & vbCr
    Response.Write " </Script>  " & vbCr 
    
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
  
    Dim PC4G008Data
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
	    
    Set PC4G008Data = Server.CreateObject("PC4G008.cMngDirDstbFctrSvr")

    If CheckSYSTEMError(Err, True) = True Then					
       Call SetErrorStatus
       Exit Sub
    End If    


    Call PC4G008Data.C_MNG_DIR_DSTB_FCTR_SVR(gStrGlobalCollection, sSQLI1, sSQLU1, sSQLD1, sSQLI2, sSQLU2, sSQLD2, lgErrorPos)			

	sErrMesg = Err.Description

	If Instr(1, sErrMesg, "B_MESSAGE" & Chr(11) & "970000") > 0 Then
		Response.Write "<script language=vbscript>" & vbCr
		Response.Write "Call parent.SubSetErrPos2(""" & Replace(lgErrorPos, chr(34), chr(34) & chr(34)) & """)" & vbCr
		Response.Write "</script>" & vbCr
		Response.End 
	End If
		
    If CheckSYSTEMError(Err, True ) = True Then
       Call SetErrorStatus
       Set PC4G008Data = Nothing
       Exit Sub
    End If    
    
    Set PC4G008Data = Nothing
	
   
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
    
    Select Case "<%=lgOpModeCRUD %>"
       Case "<%=UID_M0001%>"                                                         '☜ : Query
          If Trim("<%=lgErrorStatus%>") = "NO" Then
              With Parent
                .DBQueryOk        
	         End with
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
</Script>	
<%Response.End%>