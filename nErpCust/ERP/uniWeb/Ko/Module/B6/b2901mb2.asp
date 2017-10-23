<%@ LANGUAGE=VBSCript %>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<!-- #Include file="../../inc/IncSvrNumber.inc"  -->
<!-- #Include file="../../inc/IncSvrDate.inc"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->

<%
    On Error Resume Next
    Err.Clear
    Dim lgOpModeCRUD
    Call LoadBasisGlobalInf() 
	Call LoadInfTB19029B("I","*","NOCOOKIE", "MB")
    Call HideStatusWnd                                                               '¢Ð: Hide Processing message
    '---------------------------------------Common-----------------------------------------------------------
	'------ Developer Coding part (Start ) ------------------------------------------------------------------

	'------ Developer Coding part (End   ) ------------------------------------------------------------------ 
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

    Dim iPB6SA20BS
    Dim iStrData

    Dim exportData1
    Dim exportData2
    Dim exportData3
    Dim exportGroup

	Dim iLngRow,iLngCol
    Dim iStrPrevKey
       
    Dim iStrDeptCd
    Dim iStrDeptCd2
    Dim iStrOrgChgId

    Dim iIntMaxRows
    Dim iIntQueryCount  'iIntPrevKeyIndex
    Dim importArray
    Dim iIntLoopCount
    Dim i

    Const C_SHEETMAXROWS_D  = 100
    Const C_Org_Chd_Id = 0
    Const C_Dept_Cd = 1

	ReDim exportData1(1)
    Const A023_E1_hong_abs_orgid = 0
    Const A023_E1_hong_abs_orgnm = 1

	ReDim exportData2(1)
    Const A023_E2_b_acct_cdpt_dept_cd = 0
    Const A023_E2_b_acct_cdpt_dept_nm = 1

	On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

    iStrPrevKey		=	Request("lgStrPrevKey")
    iIntMaxRows		=	Request("txtMaxRows")
    iIntQueryCount	=	Request("lgPageNo")

    iStrOrgChgId	=	Request("txtOrgChgID")
	iStrDeptCd		=	Request("txtDeptCd")
	iStrDeptCd2		=	iStrDeptCd

    If Trim(iStrPrevKey) = "" Then
		iStrDeptCd	= iStrDeptCd
	Else
		iStrDeptCd	= iStrPrevKey
    End If

    
    If Len(Trim(iIntQueryCount))  Then                                        '¢Ð : Chnage Nextkey str into int value
       If Isnumeric(iIntQueryCount) Then
          iIntQueryCount = CInt(iIntQueryCount)
       End If
    Else
       iIntQueryCount = 0
    End If

	ReDim importArray(C_Dept_Cd)

    importArray(C_Org_Chd_Id)		= iStrOrgChgId
    importArray(C_Dept_Cd)			= iStrDeptCd

	Set iPB6SA20BS = Server.CreateObject("PB6SA20.cBListAcctDeptSvr")

    If CheckSYSTEMError(Err, True) = True Then
       Exit Sub
    End If 

	Call iPB6SA20BS.B_LIST_ACCT_DEPT_SVR(gStrGloBalCollection,C_SHEETMAXROWS_D,importArray, iStrDeptCd2, exportData1,exportData2,exportGroup)

    If CheckSYSTEMError(Err, True) = True Then
       Set iPB6SA20BS = Nothing
       Exit Sub
    End If

    Set iPB6SA20BS = Nothing

	iStrData = ""
	iIntLoopCount = 0

	If Not IsEmpty(exportGroup) Then
		For iLngRow = 0 To UBound(exportGroup, 1)
			iIntLoopCount = iIntLoopCount + 1

			If  iIntLoopCount < (C_SHEETMAXROWS_D + 1) Then
				istrData = iStrData & Chr(11) & ConvSPChars(exportGroup(iLngRow,0))
				iStrData = iStrData & Chr(11) & ConvSPChars(exportGroup(iLngRow,1))
			    iStrData = iStrData & Chr(11) & UNIDateClientFormat(Trim(exportGroup(iLngRow,2)))

				iStrData = iStrData & Chr(11) & ConvSPChars(exportGroup(iLngRow,3))
				iStrData = iStrData & Chr(11) & ConvSPChars(exportGroup(iLngRow,4))
				iStrData = iStrData & Chr(11) & ConvSPChars(exportGroup(iLngRow,5))
				iStrData = iStrData & Chr(11) & ConvSPChars(exportGroup(iLngRow,6))
				iStrData = iStrData & Chr(11) & ConvSPChars(exportGroup(iLngRow,7))
				iStrData = iStrData & Chr(11) & ""
				iStrData = iStrData & Chr(11) & ConvSPChars(exportGroup(iLngRow,8))
				iStrData = iStrData & Chr(11) & ConvSPChars(exportGroup(iLngRow,9))
				iStrData = iStrData & Chr(11) & ""
				iStrData = iStrData & Chr(11) & ConvSPChars(exportGroup(iLngRow,10))
				iStrData = iStrData & Chr(11) & exportGroup(iLngRow,11)
				iStrData = iStrData & Chr(11) & exportGroup(iLngRow,12)
				iStrData = iStrData & Chr(11) & exportGroup(iLngRow,13)
				iStrData = iStrData & Chr(11) & ConvSPChars(exportGroup(iLngRow,14))
				iStrData = iStrData & Chr(11) & ConvSPChars(exportGroup(iLngRow,15))
				iStrData = iStrData & Chr(11) & exportGroup(iLngRow,16)
				iStrData = iStrData & Chr(11) & exportGroup(iLngRow,18)
				iStrData = iStrData & Chr(11) & iIntMaxRows + iIntLoopCount

				iStrData = istrData & Chr(11) & Chr(12)
			Else
		  		iStrPrevKey = exportGroup(UBound(exportGroup,1), 0)			
				iIntQueryCount = iIntQueryCount + 1
				Exit For
			End If
		Next
	End If

	If iLngRow < C_SHEETMAXROWS_D Then
		iStrPrevKey = ""
		iIntQueryCount = ""  
	End If

	Response.Write " <Script Language=vbscript>	                        " & vbCr
	Response.Write " With parent                                        " & vbCr
    Response.Write "	.ggoSpread.Source = .frm1.vspdData              " & vbCr
    Response.Write "	.ggoSpread.SSShowData """ & iStrData       & """" & vbCr
    Response.Write "	.lgPageNo = """ & iIntQueryCount    & """" & vbCr
	Response.Write "	.frm1.hOrgChgId.value = """ &	ConvSPChars(exportData1(A023_E1_hong_abs_orgid))    & """" & vbCr
	Response.Write "	.frm1.txtOrgChgNm.value = """ & ConvSPChars(exportData1(A023_E1_hong_abs_orgnm))    & """" & vbCr
	Response.Write "	.frm1.txtDeptCd.value = """ &		ConvSPChars(Trim(exportData2(A023_E2_b_acct_cdpt_dept_cd)))    & """" & vbCr
	Response.Write "	.frm1.hDeptCd.value = """ &		ConvSPChars(exportData2(A023_E2_b_acct_cdpt_dept_cd))    & """" & vbCr
	Response.Write "	.frm1.txtDeptNM.value = """ &	ConvSPChars(exportData2(A023_E2_b_acct_cdpt_dept_nm))    & """" & vbCr
    Response.Write "	.lgStrPrevKey = """ & iStrPrevKey    & """" & vbCr
    Response.Write "	.DbQueryOk " & vbCr
    Response.Write " End With   " & vbCr
    Response.Write " </Script>  " & vbCr

End Sub    

'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()

    Dim iPB6SA20BS
    Dim iErrorPosition

    On Error Resume Next
    Err.Clear
    Set iPB6SA20BS = Server.CreateObject("PB6SA20.cBMngAcctDeptSvr")

    If CheckSYSTEMError(Err, True) = True Then
       Set iPB6SA20BS = Nothing
       Exit Sub
    End If


    Call iPB6SA20BS.B_MANAGE_ACCT_DEPT_SVR(gStrGloBalCollection, request("txtSpread"), iErrorPosition)
  	If CheckSYSTEMError2(Err, True,iErrorPosition & "Çà","","","","") = True Then
       Set iPB6SA20BS = Nothing
       Exit Sub
    End If


    Set iPB6SA20BS = Nothing

	Response.Write " <Script Language=vbscript> " & vbCr
	Response.Write " parent.DbSaveOk            " & vbCr
    Response.Write "</Script>                   " & vbCr

End Sub


'============================================================================================================
Sub SubBizSaveMultiCreate(arrColVal)
    On Error Resume Next
    Err.Clear
End Sub

'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)
    On Error Resume Next
    Err.Clear
End Sub

'============================================================================================================
Sub SubBizSaveMultiDelete(arrColVal)
    On Error Resume Next
    Err.Clear
End Sub

'============================================================================================================
Sub SubMakeSQLStatements(pDataType,arrColVal)
    Dim iSelCount
    On Error Resume Next
    Err.Clear
End Sub

'============================================================================================================
Sub CommonOnTransactionCommit()
End Sub

'============================================================================================================
Sub CommonOnTransactionAbort()
End Sub

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


