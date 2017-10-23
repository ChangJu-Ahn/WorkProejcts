<%@ LANGUAGE=VBSCript TRANSACTION=Required%>
<%Option Explicit%>

<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../comasp/loadinftb19029.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->

<!-- #Include file="../../inc/incServeradodb.asp" -->
<% 
	Call LoadBasisGlobalInf() 
	Call LoadInfTB19029B("I","*","NOCOOKIE", "MB")

    Dim lgOpModeCRUD
    On Error Resume Next
    Err.Clear

    Call HideStatusWnd

    lgOpModeCRUD      = Request("txtMode")                                           '¢Ð: Read Operation Mode (CRUD)
    
    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '¢Ð: Query
          '  Call SubBizQuery()
             Call SubBizQueryMulti()
        Case CStr(UID_M0002)                                                         '¢Ð: Save,Update
          '  Call SubBizSave()
             Call SubBizSaveMulti()
        Case CStr(UID_M0003)                                                         '¢Ð: Delete
             Call SubBizDelete()
    End Select


'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
    On Error Resume Next
    Err.Clear
End Sub
'============================================================================================================
' Name : SubBizSave
' Desc : Save Data 
'============================================================================================================
Sub SubBizSave()
    On Error Resume Next
    Err.Clear
End Sub
'============================================================================================================
' Name : SubBizDelete
' Desc : Delete DB data
'============================================================================================================
Sub SubBizDelete()
    On Error Resume Next
    Err.Clear
End Sub

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()

	Const C_SHEETMAXROWS_D	= 100
    Const C_SMAXROWS	= 0
	Const C_QRYCNT		= 1

    Dim iPB2SA000
    Dim iStrData
    Dim exportData
    Dim exportData1
    Dim iLngRow,iLngCol
    Dim iStrPrevKey
    Dim iStrTaxCd
    Dim iIntMaxRows
    Dim iIntQueryCount
    Dim importArray
    Dim iSMaxRows_QryCnt
    Dim iIntLoopCount
    Dim txtTaxCd

    On Error Resume Next
    Err.Clear

	iStrPrevKey		= Request("lgStrPrevKey")
	txtTaxCd		= Request("txtTaxCd")
    iIntMaxRows		= Request("txtMaxRows")
    iIntQueryCount	= Request("lgPageNo")
    If Trim(iStrPrevKey) = "" Then
		iStrTaxCd	= txtTaxCd
	Else
		iStrTaxCd	= iStrPrevKey
    End If


    If Len(Trim(iIntQueryCount))  Then                                        '¢Ð : Chnage Nextkey str into int value
       If Isnumeric(iIntQueryCount) Then
          iIntQueryCount = CInt(iIntQueryCount)
       End If
    Else
       iIntQueryCount = 0
    End If

    ReDim iSMaxRows_QryCnt(C_QRYCNT)
    iSMaxRows_QryCnt(C_SMAXROWS)	= C_SHEETMAXROWS_D
    iSMaxRows_QryCnt(C_QRYCNT)		= iIntQueryCount

	Set iPB2SA000 = Server.CreateObject("PB2SA00.cBListTaxOfficeSvr")

    If CheckSYSTEMError(Err, True) = True Then
       Set iPB2SA000 = Nothing
       Exit Sub
    End If

	Redim exportData(1)
	Const A555_tax_office_cd = 0
	Const A555_tax_office_nm = 1

	Call iPB2SA000.B_LIST_TAX_OFFICE_SVR(gStrGlobalCollection, iSMaxRows_QryCnt, iStrTaxCd, exportData, exportData1)

    If CheckSYSTEMError(Err, True) = True Then
       Set iPB2SA000 = Nothing
       Exit Sub
    End If

    Set iPB2SA000 = Nothing

	iStrData = ""
	iIntLoopCount = 0

	For iLngRow = 0 To UBound(exportData1, 1)
		iIntLoopCount = iIntLoopCount + 1

	    If  iIntLoopCount < (C_SHEETMAXROWS_D + 1) Then
			For iLngCol = 0 To UBound(exportData1, 2)
			    iStrData = iStrData & Chr(11) & ConvSPChars(exportData1(iLngRow, iLngCol))
			Next
			    iStrData = iStrData & Chr(11) & iIntMaxRows + iIntLoopCount
			    iStrData = iStrData & Chr(11) & Chr(12)
	    Else
			iStrPrevKey = exportData1(UBound(exportData1, 1), 0)
			iIntQueryCount = iIntQueryCount + 1
			Exit For
		End If
	Next

	If  iIntLoopCount < (C_SHEETMAXROWS_D + 1) Then
		iStrPrevKey = ""
	    iIntQueryCount = ""
	End If

	Response.Write " <Script Language=vbscript>	                        " & vbCr
	Response.Write " With parent                                        " & vbCr
    Response.Write "	.ggoSpread.Source		= .frm1.vspdData        " & vbCr
    Response.Write "	.ggoSpread.SSShowData	  """ & iStrData       & """" & vbCr
    Response.Write "	.frm1.txtTaxNm.value	= """ & ConvSPChars(exportData(A555_tax_office_nm))     & """" & vbCr
    Response.Write "	.frm1.hTaxCd.value		= """ & ConvSPChars(exportData(A555_tax_office_cd))     & """" & vbCr
    Response.Write "	.lgPageNo				= """ & iIntQueryCount & """" & vbCr
    Response.Write "	.lgStrPrevKey			= """ & ConvSPChars(iStrPrevKey)    & """" & vbCr
    Response.Write "	.DbQueryOk " & vbCr
    Response.Write " End With   " & vbCr
    Response.Write " </Script>  " & vbCr

End Sub

'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()
    Dim iiPB2SA000
    Dim iErrorPosition

    On Error Resume Next
    Err.Clear

    Set iiPB2SA000 = Server.CreateObject("PB2SA00.cBManageTaxOfficeSvr")

    If CheckSYSTEMError(Err, True) = True Then
       Set iiPB2SA000 = Nothing
       Exit Sub
    End If

    Call iiPB2SA000.B_MANAGE_TAX_OFFICE_SVR(gStrGlobalCollection,Request("txtSpread"),iErrorPosition)

  	If CheckSYSTEMError2(Err, True,iErrorPosition & "Çà","","","","") = True Then				
       Set iiPB2SA000 = Nothing
       Exit Sub
    End If

    Set iiPB2SA000 = Nothing

	Response.Write " <Script Language=vbscript> " & vbCr
	Response.Write " parent.DbSaveOk            " & vbCr
    Response.Write "</Script>                   " & vbCr

End Sub

'============================================================================================================
' Name : SubBizSaveCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveMultiCreate(arrColVal)
    On Error Resume Next
    Err.Clear
End Sub

'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)
    On Error Resume Next
    Err.Clear
End Sub

'============================================================================================================
' Name : SubBizSaveMultiDelete
' Desc : Delete Data from Db
'============================================================================================================
Sub SubBizSaveMultiDelete(arrColVal)
    On Error Resume Next
    Err.Clear
End Sub

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pDataType,arrColVal)
    Dim iSelCount
    On Error Resume Next
    Err.Clear
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
End Sub

'============================================================================================================
' Name : SetErrorStatus
' Desc : This Sub set error status
'============================================================================================================
Sub SetErrorStatus()
End Sub
'============================================================================================================
' Name : SubHandleError
' Desc : This Sub handle error
'============================================================================================================
Sub SubHandleError(pOpCode,pConn,pRs,pErr)
    On Error Resume Next
    Err.Clear
End Sub
%>