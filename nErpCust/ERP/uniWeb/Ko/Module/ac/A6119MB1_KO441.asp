<%@ LANGUAGE=VBSCript TRANSACTION=Required%>

<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/IncServerAdoDb.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%
	Dim lgStrPrevKey
	Const C_SHEETMAXROWS_D = 100

    Dim lgSvrDateTime
    Dim lgSubChkDataExist
   
    call LoadBasisGlobalInf()
    Call LoadInfTB19029B("I", "H","NOCOOKIE","MB")
    lgSvrDateTime = GetSvrDateTime

    Call HideStatusWnd                                                               '¢Ð: Hide Processing message

    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '¢Ð: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '¢Ð: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)

    lgLngMaxRow       = Request("txtMaxRows")                                        '¢Ð: Read Operation Mode (CRUD)
    lgStrPrevKey = UNICInt(Trim(Request("lgStrPrevKey")),0)                '¢Ð: "0"(First),"1"(Second),"2"(Third),"3"(...)

    Call SubOpenDB(lgObjConn)                                                        '¢Ð: Make a DB Connection
    
    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '¢Ð: Query
             Call SubBizQuery()
        Case CStr(UID_M0002)                                                         '¢Ð: Save,Update
             Call SubBizSaveMulti()
        Case CStr(UID_M0003)                                                         '¢Ð: Delete
             Call SubBizDelete()
    End Select
    
    Call SubCloseDB(lgObjConn)                                                       '¢Ð: Close DB Connection

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status
    Call SubBizQueryMulti()
End Sub    
'============================================================================================================
' Name : SubBizDelete
' Desc : Delete DB data
'============================================================================================================
Sub SubBizDelete()
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status
End Sub

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
    Dim iDx
    Dim iLoopMax
    Dim iKey1,iKey2,iKey3,iKey4
    Dim strRoll_pstn,strDate
    Dim strPay_grd1, strYear, strMonth, strDay

    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

	
		iKey1 = FilterVar(Replace(lgKeyStream(0),"-",""),"''", "S")
		iKey2 = FilterVar(Replace(lgKeyStream(1),"-",""),"''", "S")
	if lgKeyStream(2) <> "" then		'req_yn
		iKey3 = FilterVar(lgKeyStream(2),"''", "S")
	end if
	if lgKeyStream(3) <> "" then		'req_yn
		iKey4 = FilterVar(lgKeyStream(3),"''", "S")
	end if

    
    
    Call SubMakeSQLStatements("MR",iKey1,iKey2,iKey3,iKey4)                                 '¡Ù : Make sql statements
    
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        lgStrPrevKey = ""
        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '¢Ð : No data is found. 
        Call SetErrorStatus()
    Else

        Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKey)

        lgstrData = ""
        iDx       = 1
        
        Do While Not lgObjRs.EOF
			
			Call ExtractDateFrom(ConvSPChars(lgObjRs("deduction_dt")), "YYYYMM", "", strYear, strMonth, strDay)
			strDate = UniConvYYYYMMDDToDate(gAPDateFormat, strYear, strMonth, "01")
			strDate = left(strDate,7)
			lgstrData = lgstrData & Chr(11) & strDate
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("tax_biz_area_cd"))
			lgstrData = lgstrData & Chr(11) & ""
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("tax_biz_area_nm"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("deduction_type"))
			lgstrData = lgstrData & Chr(11) & ""
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("minor_nm"))
			lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("deduction_cnt"), 0, 0)	
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("deduction_amt"), 0, 0)
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("vat_amt"), 0, 0)
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("deduction_desc"))
            lgstrData = lgstrData & Chr(11) & lgLngMaxRow + iDx
            lgstrData = lgstrData & Chr(11) & Chr(12)

		    lgObjRs.MoveNext

            iDx =  iDx + 1
            If iDx > C_SHEETMAXROWS_D Then
               lgStrPrevKey = lgStrPrevKey + 1
               Exit Do
            End If   
               
        Loop 
    End If
    
    If iDx <= C_SHEETMAXROWS_D Then
       lgStrPrevKey = ""
    End If   

	Call SubHandleError("MR",lgObjConn,lgObjRs,Err)
    Call SubCloseRs(lgObjRs)                                                          '¢Ð: Release RecordSSet

End Sub    

'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()

    Dim arrRowVal
    Dim arrColVal
    Dim iDx

    On Error Resume Next                                                             '¢Ð: Protect system from crashing

    Err.Clear                                                                        '¢Ð: Clear Error status
    
	arrRowVal = Split(Request("txtSpread"), gRowSep)                                 '¢Ð: Split Row    data
	
    For iDx = 1 To lgLngMaxRow
        arrColVal = Split(arrRowVal(iDx-1), gColSep)                                 '¢Ð: Split Column data
        
        Select Case arrColVal(0)
            Case "C"
                    Call SubBizSaveMultiCreate(arrColVal)                            '¢Ð: Create
            Case "U"
                    Call SubBizSaveMultiUpdate(arrColVal)                            '¢Ð: Update
            Case "D"
                    Call SubBizSaveMultiDelete(arrColVal)                            '¢Ð: Delete
        End Select
        
        If lgErrorStatus    = "YES" Then
           lgErrorPos = lgErrorPos & arrColVal(1) & gColSep
           Exit For
        End If
        
    Next

End Sub    

'============================================================================================================
' Name : SubBizSaveCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveMultiCreate(ByVal arrColVal)
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

	Call SubChkDataExist(arrColVal)

	If lgSubChkDataExist = True Then
	    Call DisplayMsgBox("970001", vbInformation, replace(arrColVal(2),"-","")+","+Trim(UCase(arrColVal(3)))+","+Trim(UCase(arrColVal(4))), "", I_MKSCRIPT)      '¢Ð : No data is found. 
        Call SetErrorStatus()
	Else
		lgStrSQL = "INSERT INTO a_vat_deduction( deduction_dt, tax_biz_area_cd, deduction_type,deduction_cnt,deduction_amt,vat_amt,deduction_desc,INSRT_USER_ID,INSRT_DT,UPDT_USER_ID,UPDT_DT )" 
		lgStrSQL = lgStrSQL & " VALUES(" 
		lgStrSQL = lgStrSQL & FilterVar(replace(arrColVal(2),"-",""),"''","S")	& ","
		lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(3))),"","S")		& ","
		lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(4))),"","S")		& ","
		lgStrSQL = lgStrSQL & UNIConvNum(arrColVal(5),0)						& ","
		lgStrSQL = lgStrSQL & UNIConvNum(arrColVal(6),0)						& ","
		lgStrSQL = lgStrSQL & UNIConvNum(arrColVal(7),0)						& ","
		lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(8))),"","S")		& ","
		lgStrSQL = lgStrSQL & FilterVar(gUsrId,"","S")							& "," 
		lgStrSQL = lgStrSQL & FilterVar(lgSvrDateTime,"''","S")					& "," 
		lgStrSQL = lgStrSQL & FilterVar(gUsrId,"","S")							& "," 
		lgStrSQL = lgStrSQL & FilterVar(lgSvrDateTime,"''","S")
		lgStrSQL = lgStrSQL & ")"

		lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
		Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
	End If    
End Sub

'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate(ByVal arrColVal)

    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

	Call SubChkDataExist(arrColVal)

	If lgSubChkDataExist = True Then
		lgStrSQL = "UPDATE  a_vat_deduction"
		lgStrSQL = lgStrSQL & " SET "
		lgStrSQL = lgStrSQL & " deduction_cnt	= " &  UNIConvNum(arrColVal(5),0) & "," 
		lgStrSQL = lgStrSQL & " deduction_amt	= " &  UNIConvNum(arrColVal(6),0) & ","
		lgStrSQL = lgStrSQL & " vat_amt			= " &  UNIConvNum(arrColVal(7),0) & ","
		lgStrSQL = lgStrSQL & " deduction_desc  = " &  FilterVar(Trim(UCase(arrColVal(8))),"","S") & ","
		lgStrSQL = lgStrSQL & " UPDT_USER_ID	= " &  FilterVar(gUsrId,"","S")   & ","
		lgStrSQL = lgStrSQL & " UPDT_DT			= " &  FilterVar(lgSvrDateTime,"''","S")  
		lgStrSQL = lgStrSQL & " WHERE   deduction_dt		= " &  FilterVar(replace(arrColVal(2),"-",""),NULL,"S")
		lgStrSQL = lgStrSQL & "   and	tax_biz_area_cd		= " &  FilterVar(Trim(UCase(arrColVal(3))),"","S")
		lgStrSQL = lgStrSQL & "   and	deduction_type		= " &  FilterVar(Trim(UCase(arrColVal(4))),"","S")
    

		lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
		Call SubHandleError("MU",lgObjConn,lgObjRs,Err)
	Else
	    Call DisplayMsgBox("970010", vbInformation, replace(arrColVal(2),"-","")+","+Trim(UCase(arrColVal(3)))+","+Trim(UCase(arrColVal(4))), "", I_MKSCRIPT)      '¢Ð : No data is found. 
        Call SetErrorStatus()	
	End If	
	
End Sub


'============================================================================================================
' Name : SubBizSaveMultiDelete
' Desc : Delete Data from Db
'============================================================================================================
Sub SubBizSaveMultiDelete(ByVal arrColVal)

    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

	Call SubChkDataExist(arrColVal)

	If lgSubChkDataExist = True Then
		lgStrSQL = "DELETE  a_vat_deduction"
		lgStrSQL = lgStrSQL & " WHERE   deduction_dt			= " &  FilterVar(replace(arrColVal(2),"-",""),NULL,"S")
		lgStrSQL = lgStrSQL & "   and	tax_biz_area_cd		= " &  FilterVar(Trim(UCase(arrColVal(3))),"","S")
		lgStrSQL = lgStrSQL & "   and	deduction_type		= " &  FilterVar(Trim(UCase(arrColVal(4))),"","S")

		lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
		Call SubHandleError("MD",lgObjConn,lgObjRs,Err)
	Else
	    Call DisplayMsgBox("970010", vbInformation, replace(arrColVal(2),"-","")+","+Trim(UCase(arrColVal(3)))+","+Trim(UCase(arrColVal(4))), "", I_MKSCRIPT)      '¢Ð : No data is found. 
        Call SetErrorStatus()		
	End If
End Sub

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pDataType,pCode,pCode1,pCode2,pCode3)
    Dim iSelCount

    Select Case Mid(pDataType,1,1)
        Case "M"
        
           iSelCount = C_SHEETMAXROWS_D + C_SHEETMAXROWS_D *  lgStrPrevKey + 1
           
           Select Case Mid(pDataType,2,1)
               Case "R"

					lgStrSQL = "           Select TOP " & iSelCount  & "  deduction_dt, "
					lgStrSQL = lgStrSQL & "       a.tax_biz_area_cd, "
					lgStrSQL = lgStrSQL & "       TAX_BIZ_AREA_NM, "
					lgStrSQL = lgStrSQL & "       deduction_type, "
					lgStrSQL = lgStrSQL & "       minor_nm, "
					lgStrSQL = lgStrSQL & "       deduction_cnt, "
					lgStrSQL = lgStrSQL & "       deduction_amt, "
					lgStrSQL = lgStrSQL & "       vat_amt, "
					lgStrSQL = lgStrSQL & "       deduction_desc "
					lgStrSQL = lgStrSQL & "  from a_vat_deduction a, "
					lgStrSQL = lgStrSQL & "       b_tax_biz_area  b, "
					lgStrSQL = lgStrSQL & "       b_minor         c "
					lgStrSQL = lgStrSQL & " where a.tax_biz_area_cd = b.tax_biz_area_cd "
					lgStrSQL = lgStrSQL & "   and a.deduction_type = c.minor_cd "
					lgStrSQL = lgStrSQL & "   and c.major_cd = 'A3001' "
					lgStrSQL = lgStrSQL & "	  AND A.deduction_dt between " & pCode & " and " & pCode1
					if pCode2 <> "" then
					lgStrSQL = lgStrSQL & "	  AND A.tax_biz_area_cd = " & pCode2
					End if
					if pCode3 <> "" then
					lgStrSQL = lgStrSQL & "	  AND A.deduction_type = " & pCode3
                    End if
                    lgStrSQL = lgStrSQL & " order by deduction_dt,a.tax_biz_area_cd,deduction_type "
                    
           End Select             
    End Select
End Sub

Sub SubChkDataExist(ByVal arrColVal)
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear 

	lgStrSQL = "           Select deduction_type "
	lgStrSQL = lgStrSQL & "  from a_vat_deduction  "
	lgStrSQL = lgStrSQL & " where deduction_dt =  " & FilterVar(replace(arrColVal(2),"-",""),NULL,"S")
	lgStrSQL = lgStrSQL & "	 AND  tax_biz_area_cd = " & FilterVar(Trim(UCase(arrColVal(3))),"","S")
	lgStrSQL = lgStrSQL & "	 AND  deduction_type = " & FilterVar(Trim(UCase(arrColVal(4))),"","S")

    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = True Then
		lgSubChkDataExist = True  
    Else
		lgSubChkDataExist = False 
	End If
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
    lgErrorStatus    = "YES"
End Sub

'============================================================================================================
' Name : SetErrorStatus
' Desc : This Sub set error status
'============================================================================================================
Sub SetErrorStatus()
    lgErrorStatus     = "YES"                                                         '¢Ð: Set error status
End Sub
'============================================================================================================
' Name : SubHandleError
' Desc : This Sub handle error
'============================================================================================================
Sub SubHandleError(pOpCode,pConn,pRs,pErr)
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

    Select Case pOpCode
        Case "MC"
                 If CheckSYSTEMError(pErr,True) = True Then
                    ObjectContext.SetAbort
                    Call SetErrorStatus
                 Else
                    If CheckSQLError(pConn,True) = True Then
                       ObjectContext.SetAbort
                       Call SetErrorStatus
                    End If
                 End If
        Case "MD"
        Case "MR"
        Case "MU"
                 If CheckSYSTEMError(pErr,True) = True Then
                    ObjectContext.SetAbort
                    Call SetErrorStatus
                 Else
                    If CheckSQLError(pConn,True) = True Then
                       ObjectContext.SetAbort
                       Call SetErrorStatus
                    End If
                 End If
    End Select
End Sub

%>

<Script Language="VBScript">
    
    Select Case "<%=lgOpModeCRUD %>"
       Case "<%=UID_M0001%>"                                                         '¢Ð : Query
          If Trim("<%=lgErrorStatus%>") = "NO" Then
              With Parent
                .ggoSpread.Source     = .frm1.vspdData
                .lgStrPrevKey    = "<%=lgStrPrevKey%>"
                .ggoSpread.SSShowData "<%=lgstrData%>"
                .lgStrPrevKey         = "<%=lgStrPrevKey%>"
                .DBQueryOk        
	         End with
          End If   
       Case "<%=UID_M0002%>"                                                         '¢Ð : Save
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DBSaveOk
          Else
             Parent.SubSetErrPos(Trim("<%=lgErrorPos%>"))
          End If   
       Case "<%=UID_M0002%>"                                                         '¢Ð : Delete
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DbDeleteOk
          Else   
          End If   
    End Select    
    
       
</Script>	
