<%@ LANGUAGE=VBSCript TRANSACTION=Required%>
<% Option Explicit%>
<!-- #Include file="../../inc/IncSvrMAin.asp"  -->
<!-- #Include file="../../inc/IncSvrDate.inc"  -->
<!-- #Include file="../../inc/IncSvrNumber.inc"  -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/incServerAdoDb.asp"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<!-- #Include file="../../inc/lgsvrvariables.inc"  -->

<%
    On Error Resume Next
	Err.Clear

    Call HideStatusWnd

    Call LoadBasisGlobalInf()
    Call LoadInfTB19029B("I","*","NOCOOKIE","MB")

    Call SubOpenDB(lgObjConn)                                                        '¢Ð: Make a DB Connection
	Call SplitSpreadData()
    Call SubCloseDB(lgObjConn)                                                       '¢Ð: Close DB Connection

'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data To DB
'============================================================================================================
Sub SplitSpreadData()
	Dim spd_data_set
	Dim tmp_spd
	Dim arrColVal,arrRowVal,ii
	
	tmp_spd = Trim(request("txtSpread"))
	
	If tmp_spd <> "" Then
		arrRowVal = Split(tmp_spd,gRowSep)
		
        ReDim spd_data_set(UBound(arrRowVal) - 1, 1)

        For ii = 0 To UBound(arrRowVal) - 1
            arrColVal = Split(arrRowVal(ii), gColSep)
            
			If arrColVal(0) = "U" Then
				spd_data_set(ii,0) = arrColVal(2)
				spd_data_set(ii,1) = arrColVal(3)
			End If
        Next
        
        Call SubBizSaveMultiUpdate(spd_data_set)
	End If
End Sub

'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data To DB
'============================================================================================================
Sub SubBizSaveMultiUpdate(ByVal UpdateArr)
    Dim lgStrSQL
    Dim ii

    On Error Resume Next
    Err.Clear

	lgStrSQL = " Begin Tran "
	For ii = 0 To UBound(UpdateArr,1) 
	    lgStrSQL = lgStrSQL & " UPDATE  A_SUBLEDGER_CTRL SET "
		lgStrSQL = lgStrSQL & " GL_CTRL_NM = " & FilterVar(UpdateArr(ii,1), "''", "S")
		lgStrSQL = lgStrSQL & " WHERE GL_CTRL_FLD = " & FilterVar(UpdateArr(ii,0), "''", "S") & vbCr
	Next

    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords

    If CheckSYSTEMError(Err,True) = True Then
		lgStrSQL = " Rollback Tran "
		lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
    Else
		lgStrSQL = " Commit Tran "		
		lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords    
		Response.Write  " <Script Language=vbscript> " & vbCr
		Response.Write  "	Parent.DBSaveOk			" & vbCr
		Response.Write  " </Script>                  " & vbCr
    End If   
End Sub

'============================================================================================================
Sub CommonOnTransactionCommit()

End Sub

'============================================================================================================
Sub CommonOnTransactionAbort()

End Sub

%>
