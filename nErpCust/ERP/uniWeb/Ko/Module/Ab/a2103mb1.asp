<%@ LANGUAGE=VBSCript TRANSACTION=Required%>
<%Option Explicit%>
<!-- #Include file="../../inc/incSvrMain.asp"  -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<%


Call LoadBasisGlobalInf()

    On Error Resume Next
    Err.Clear

    Call HideStatusWnd                                                               '¢Ð: Hide Processing message
    '---------------------------------------Common-----------------------------------------------------------
    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '¢Ð: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '¢Ð: Read Operation Mode (CRUD)


    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '¢Ð: Query
             Call SubBizQuery()
        Case CStr(UID_M0002)                                                         '¢Ð: Save,Update
             Call SubBizSave()
        Case CStr(UID_M0003)                                                         '¢Ð: Delete
             Call SubBizDelete()
    End Select

    'Call SubCloseDB(lgObjConn)                                                       '¢Ð: Close DB Connection

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()

    Dim iPABG015
    Dim iStrData
    Dim ExportData(1)
    Dim ExportReturn
    Dim iLngRow,iLngCol
    Dim iStrPrevKey
    Dim iStrBank_cd
    Dim iIntMaxRows
    Dim iIntQueryCount
    Dim importArray
    Dim iIntLoopCount
    Dim txtClassType
    Const C_SHEETMAXROWS_D  = 100
    Const C_QueryConut		= 0
    Const C_MaxQueryReCord = 1
    Const C_TaxOffice_Cd = 2

    Const E1_class_type = 0
    Const E1_class_type_nm = 1

    On Error Resume Next
    Err.Clear

	txtClassType	= Request("txtClassType")
	iStrPrevKey		= Request("lgStrPrevKey")
    iIntMaxRows		= Request("txtMaxRows")
    iIntQueryCount	= Request("lgPageNo")
    If Trim(iStrPrevKey) = "" Then
		iStrBank_cd	= txtClassType
	Else
		iStrBank_cd	= iStrPrevKey
    End If

    If Len(Trim(iIntQueryCount))  Then                                        '¢Ð : Chnage Nextkey str into int value
       If Isnumeric(iIntQueryCount) Then
          iIntQueryCount = CInt(iIntQueryCount)
       End If
    Else
       iIntQueryCount = 0
    End If

    ReDim importArray(2)
    importArray(C_QueryConut)		= iIntQueryCount
    importArray(C_MaxQueryReCord)	= C_SHEETMAXROWS_D
    importArray(C_TaxOffice_Cd)		= iStrBank_cd

	Set iPABG015 = Server.CreateObject("PABG015.cAListAcctClssTpSvr")
    If CheckSYSTEMError(Err, True) = True Then
       Set iPABG015 = Nothing
       Exit Sub
    End If
	Call iPABG015.C_LIST_ACCT_CLASS_TYPE_SVR(gStrGloBalCollection, importArray, txtClassType, ExportData, ExportReturn)

    If CheckSYSTEMError(Err, True) = True Then
       Set iPABG015 = Nothing
       Exit Sub
    End If

    Set iPABG015 = Nothing
	iStrData = ""
	iIntLoopCount = 0
	For iLngRow = 0 To UBound(ExportReturn, 1)
		iIntLoopCount = iIntLoopCount + 1

	    If  iIntLoopCount < (C_SHEETMAXROWS_D + 1) Then
			For iLngCol = 0 To UBound(ExportReturn, 2)
					iStrData = iStrData & Chr(11) & ConvSPChars(ExportReturn(iLngRow, iLngCol))
			Next
				istrData = istrData & Chr(11) & iIntMaxRows + iIntLoopCount
				istrData = istrData & Chr(11) & Chr(12)
	    Else
			iStrPrevKey = ExportReturn(UBound(ExportReturn, 1), 0)
			iIntQueryCount = iIntQueryCount + 1
			Exit For
		End If
	Next

	If  iIntLoopCount < (C_SHEETMAXROWS_D + 1) Then
		iStrPrevKey = ""
	    iIntQueryCount = ""
	End If

	Response.Write " <Script Language=vbscript>	                           " & vbCr
	Response.Write " With parent                                           " & vbCr
    Response.Write "	.ggoSpread.Source			= .frm1.vspdData                 " & vbCr
    Response.Write "	.ggoSpread.SSShowData			""" & iStrData          & """" & vbCr
    Response.Write "	.frm1.hClassType.value		= """ & ConvSPChars(ExportData(E1_class_type))		      & """" & vbCr
    Response.Write "	.frm1.txtClassTypeNM.value	= """ & ConvSPChars(ExportData(E1_class_type_nm)) & """" & vbCr
    Response.Write "	.lgPageNo					= """ & iIntQueryCount		      & """" & vbCr
    Response.Write "	.lgStrPrevKey				= """ & ConvSPChars(iStrPrevKey)		      & """" & vbCr
    Response.Write "	.DbQueryOk " & vbCr
    Response.Write " End With   " & vbCr
    Response.Write " </Script>  " & vbCr

End Sub	
'============================================================================================================
' Name : SubBizQuery
' Desc : Date data 
'============================================================================================================
Sub SubBizSave()

    Dim iPABG015
    Dim txtSpread
    Dim iErrorPosition

    On Error Resume Next
    Err.Clear

    txtSpread = replace(Trim(Request("txtSpread")),",","")
    Set iPABG015 = Server.CreateObject("PABG015.cAMngAcctClssTpSvr")

    If CheckSYSTEMError(Err, True) = True Then
       Set iPABG015 = Nothing
       Exit Sub
    End If

    Call iPABG015.C_MANAGE_ACCT_CLASS_TYPE_SVR(gStrGloBalCollection, txtSpread, iErrorPosition)

  	If CheckSYSTEMError2(Err, True,iErrorPosition & "Çà","","","","") = True Then
       Set iPABG015 = Nothing
       Exit Sub
    End If

    Set iPABG015 = Nothing

	Response.Write " <Script Language=vbscript> " & vbCr
	Response.Write " parent.DbSaveOk            " & vbCr
    Response.Write "</Script>                   " & vbCr

End Sub	

'============================================================================================================
Sub SubBizDelete()
    On Error Resume Next
    Err.Clear
End Sub

'============================================================================================================
Sub SubBizSaveSingleCreate()
    On Error Resume Next
    Err.Clear
End Sub

'============================================================================================================
Sub SubBizSaveSingleUpdate()
    On Error Resume Next
    Err.Clear
End Sub

'============================================================================================================
Sub SubMakeSQLStatements(pMode)
    On Error Resume Next
End Sub

'============================================================================================================
Sub CommonOnTransactionCommit()
End Sub

'============================================================================================================
Sub CommonOnTransactionAbort()
    Call SetErrorStatus()
End Sub

'============================================================================================================
Sub SetErrorStatus()
    lgErrorStatus     = "YES"                                                         '¢Ð: Set error status
End Sub

'============================================================================================================
Sub SubHandleError(pOpCode,pConn,pRs,pErr)
    On Error Resume Next
    Err.Clear
End Sub

%>