<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!--'======================================================================================================
'*  1. Module Name          : Inventory
'*  2. Function Name        : 
'*  3. Program ID           : i1921qb1.asp
'*  4. Program Name         : 월별전표생성수불조회 
'*  5. Program Desc         : 
'*  6. Modified date(First) : 2003/05/23
'*  7. Modified date(Last)  : 2003/05/23
'*  8. Modifier (First)     : Ahn Jung Je
'*  9. Modifier (Last)      : 
'* 10. Comment              : 
'=======================================================================================================-->
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" --> 
<%
Call LoadBasisGlobalInf()
Call LoadInfTB19029B("Q", "I", "NOCOOKIE", "QB")

On Error Resume Next
Err.Clear																	

Call HideStatusWnd

Dim IntRetCD
Dim PvArr
Dim strYear, strMonth

lgErrorStatus   = "NO"
strYear		= Request("strYear")
strMonth	= Request("strMonth")

On Error Resume Next

Call SubOpenDB(lgObjConn)

Call SubCreateCommandObject(lgObjComm)

Call SubBizQuery()

Call SubCloseCommandObject(lgObjComm)    
Call SubCloseDB(lgObjConn)     

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
	
	Dim iDx
	
	On Error Resume Next             
    Err.Clear
    
	Call SubMakeSQLStatements        
    
	If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then  
		IntRetCD = -1
		Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT) 
		Call SetErrorStatus()

		Call SubCloseRs(lgObjRs)

		Response.End 
	Else
		IntRetCD = 1

        iDx = 0
        ReDim PvArr(0) 
        
        Do While Not lgObjRs.EOF
 
            iDx = iDx + 1
        
            lgstrData = Chr(11) & ConvSPChars(lgObjRs(1)) & _
						Chr(11) & UniConvNumberDBToCompany(lgObjRs(0), ggAmtOfMoney.DecPoint, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit, 0) & _
						Chr(11) & UniConvNumberDBToCompany(lgObjRs(3), ggAmtOfMoney.DecPoint, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit, 0) & _
						Chr(11) & UniConvNumberDBToCompany(lgObjRs(4), ggAmtOfMoney.DecPoint, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit, 0) & _
						Chr(11) & UniConvNumberDBToCompany(lgObjRs(5), ggAmtOfMoney.DecPoint, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit, 0) & _
						Chr(11) & UniConvNumberDBToCompany(lgObjRs(6), ggAmtOfMoney.DecPoint, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit, 0) & _
						Chr(11) & UniConvNumberDBToCompany(lgObjRs(7), ggAmtOfMoney.DecPoint, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit, 0) & _
						Chr(11) & UniConvNumberDBToCompany(lgObjRs(8), ggAmtOfMoney.DecPoint, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit, 0) & _
						Chr(11) & UniConvNumberDBToCompany(lgObjRs(9), ggAmtOfMoney.DecPoint, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit, 0) & _
						Chr(11) & UniConvNumberDBToCompany(lgObjRs(10), ggAmtOfMoney.DecPoint, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit, 0) & _
						Chr(11) & UniConvNumberDBToCompany(lgObjRs(11), ggAmtOfMoney.DecPoint, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit, 0) & _
						Chr(11) & ConvSPChars(lgObjRs(12)) & _
						Chr(11) & iDx & Chr(11) & Chr(12)
			
			ReDim Preserve PvArr(iDx - 1)
			 
			PvArr(iDx - 1) = lgstrData
		    lgObjRs.MoveNext
        Loop 
    End If

	lgstrData = Join(PvArr, "")

	Call SubHandleError("MR",lgObjConn,lgObjRs,Err)
    Call SubCloseRs(lgObjRs)                                             
		
End Sub

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements()

    On Error Resume Next      
    Err.Clear                 
	
	lgStrSQL = " SELECT isnull(d.INV_AMT,0), " & _
						" (select MINOR_NM from B_MINOR where MINOR_CD = isnull(d.ITEM_ACCT, z.ITEM_ACCT) and MAJOR_CD = " & FilterVar("P1001", "''", "S") & ") MINOR_NM, " &_ 
						" isnull(MR_AMT ,0), isnull(MR_AMT,0), isnull(PR_AMT,0), isnull(OR_AMT,0), isnull( ST_DEB_AMT,0),isnull(PI_AMT,0), isnull( DI_AMT,0),isnull( OI_AMT,0), isnull( ST_CRE_AMT,0), " & _
 						"isnull(d.INV_AMT,0) + (isnull(MR_AMT,0) + isnull(PR_AMT,0) +  isnull(OR_AMT,0) + isnull( ST_DEB_AMT,0)) - (isnull(PI_AMT,0) + isnull( DI_AMT,0) + isnull( OI_AMT,0) + isnull( ST_CRE_AMT,0)), " & _
						"isnull(d.ITEM_ACCT, z.ITEM_ACCT) " & _
				 " FROM ( Select c.ITEM_ACCT, " & _
							 " isnull(SUM(CASE WHEN b.TRNS_TYPE = " & FilterVar("MR", "''", "S") & " AND b.DEBIT_CREDIT_FLAG= " & FilterVar("D", "''", "S") & "  THEN b.AMOUNT " & _
									  " WHEN b.TRNS_TYPE = " & FilterVar("MR", "''", "S") & " AND b.DEBIT_CREDIT_FLAG= " & FilterVar("C", "''", "S") & "  THEN b.AMOUNT * (-1) ELSE 0 END),0) AS MR_AMT, " & _
							 " isnull(SUM(CASE WHEN b.TRNS_TYPE = " & FilterVar("PR", "''", "S") & " AND b.DEBIT_CREDIT_FLAG= " & FilterVar("D", "''", "S") & "  THEN b.AMOUNT " & _
							          " WHEN b.TRNS_TYPE = " & FilterVar("PR", "''", "S") & " AND b.DEBIT_CREDIT_FLAG= " & FilterVar("C", "''", "S") & "  THEN b.AMOUNT * (-1) ELSE 0 END),0) AS PR_AMT, " & _
							 " isnull(SUM(CASE WHEN b.TRNS_TYPE = " & FilterVar("OR", "''", "S") & " AND b.DEBIT_CREDIT_FLAG= " & FilterVar("D", "''", "S") & "  THEN b.AMOUNT " & _
							          " WHEN b.TRNS_TYPE = " & FilterVar("OR", "''", "S") & " AND b.DEBIT_CREDIT_FLAG= " & FilterVar("C", "''", "S") & "  THEN b.AMOUNT * (-1) ELSE 0 END),0) AS OR_AMT, " & _
							 " isnull(SUM(CASE WHEN b.TRNS_TYPE = " & FilterVar("ST", "''", "S") & " AND b.DEBIT_CREDIT_FLAG= " & FilterVar("D", "''", "S") & "  THEN b.AMOUNT ELSE 0 END ),0) AS ST_DEB_AMT, " & _
							 " isnull(SUM(CASE WHEN b.TRNS_TYPE = " & FilterVar("PI", "''", "S") & " AND b.DEBIT_CREDIT_FLAG= " & FilterVar("C", "''", "S") & "  THEN b.AMOUNT " & _
							          " WHEN b.TRNS_TYPE = " & FilterVar("PI", "''", "S") & " AND b.DEBIT_CREDIT_FLAG= " & FilterVar("D", "''", "S") & "  THEN b.AMOUNT * (-1) ELSE 0 END ),0) AS PI_AMT, " & _
							 " isnull(SUM(CASE WHEN b.TRNS_TYPE = " & FilterVar("DI", "''", "S") & " AND b.DEBIT_CREDIT_FLAG= " & FilterVar("C", "''", "S") & "  THEN b.AMOUNT " & _
							          " WHEN b.TRNS_TYPE = " & FilterVar("DI", "''", "S") & " AND b.DEBIT_CREDIT_FLAG= " & FilterVar("D", "''", "S") & "  THEN b.AMOUNT * (-1) ELSE 0 END ),0) AS DI_AMT, " & _
							 " isnull(SUM(CASE WHEN b.TRNS_TYPE = " & FilterVar("OI", "''", "S") & " AND b.DEBIT_CREDIT_FLAG= " & FilterVar("C", "''", "S") & "  THEN b.AMOUNT " & _
							          " WHEN b.TRNS_TYPE = " & FilterVar("OI", "''", "S") & " AND b.DEBIT_CREDIT_FLAG= " & FilterVar("D", "''", "S") & "  THEN b.AMOUNT * (-1) ELSE 0 END ),0) AS OI_AMT, " & _
							 " isnull(SUM(CASE WHEN b.TRNS_TYPE = " & FilterVar("ST", "''", "S") & " AND b.DEBIT_CREDIT_FLAG= " & FilterVar("C", "''", "S") & "  THEN b.AMOUNT ELSE 0 END ),0) AS ST_CRE_AMT " & _
						" From I_GOODS_MOVEMENT_HEADER a " & _
						     " inner join I_GOODS_MOVEMENT_DETAIL b " & _
						        " on a.ITEM_DOCUMENT_NO = b.ITEM_DOCUMENT_NO and a.DOCUMENT_YEAR = b.DOCUMENT_YEAR and b.DELETE_FLAG = " & FilterVar("N", "''", "S") & "  " & _
						     " left outer join B_ITEM_BY_PLANT c " & _
						        " on b.ITEM_CD = c.ITEM_CD and  b.PLANT_CD = c.PLANT_CD " & _
							 " left outer join (select MOV_TYPE, GUI_CONTROL_FLAG, GUI_CONTROL_FLAG3 from I_MOVETYPE_CONFIGURATION " & _
											   " where TRNS_TYPE = " & FilterVar("ST", "''", "S") & ")  g " & _
								" on g.MOV_TYPE = a.MOV_TYPE "
	
	lgStrSQL = lgStrSQL & " Where ((a.TRNS_TYPE = " & FilterVar("ST", "''", "S") & " and ((g.GUI_CONTROL_FLAG = " & FilterVar("Y", "''", "S") & "  and " & _
											     " b.BIZ_AREA_CD <> (select BIZ_AREA_CD from B_PLANT where PLANT_CD = b.TRNS_PLANT_CD)) " & _ 
											    " or g.GUI_CONTROL_FLAG3 = " & FilterVar("Y", "''", "S") & "  )) " & _
									" or a.TRNS_TYPE <> " & FilterVar("ST", "''", "S") & ") " & _
							" and a.POST_FLAG = " & FilterVar("Y", "''", "S") & "  " & _
							" and convert(char(6), a.DOCUMENT_DT, 112) = " & FilterVar(strYear + strMonth, "''", "S") & _
						  " Group by  c.ITEM_ACCT) Z " & _
				" FULL OUTER JOIN (select f.ITEM_ACCT, sum(e.INV_AMT) INV_AMT " & _ 
								   " from I_MONTHLY_INVENTORY e" & _
								     " inner join B_ITEM_BY_PLANT f " & _
										" on e.ITEM_CD = f.ITEM_CD and e.PLANT_CD = f.PLANT_CD " & _
								  " where e.MNTH_INV_YEAR = convert(char(4), dateadd(day,-1," & FilterVar(strYear + strMonth, "''", "S") & " +" & FilterVar("01", "''", "S") & "), 112) " & _
								    " and e.MNTH_INV_MONTH = convert(char(2), dateadd(day,-1," & FilterVar(strYear + strMonth, "''", "S") & " +" & FilterVar("01", "''", "S") & "), 110) " & _
								  " group by f.ITEM_ACCT) d" & _
					" ON D.ITEM_ACCT = Z.ITEM_ACCT " & _
				" ORDER BY isnull(d.ITEM_ACCT, z.ITEM_ACCT) "

  
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
    lgErrorStatus     = "YES"                                                   
End Sub

'============================================================================================================
' Name : SubHandleError
' Desc : This Sub handle error
'============================================================================================================
Sub SubHandleError(pOpCode,pConn,pRs,pErr)
    On Error Resume Next                                                           
    Err.Clear                                                                      

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
        Case "MB"
			ObjectContext.SetAbort
            Call SetErrorStatus        
    End Select
End Sub


Response.Write "<Script language=vbs> " & vbCr         
Response.Write " With Parent "      	& vbCr
Response.Write "	If """ & lgErrorStatus & """ = ""NO"" And """ & IntRetCd & """ <> -1 Then "	& vbCr
Response.Write "		.ggoSpread.Source	= .frm1.vspdData1 "				& vbCr
Response.Write "		.ggoSpread.SSShowData  """ & lgstrData  & """"        & vbCr
Response.Write "		.DbQueryOK						"				& vbCr
Response.Write "		.frm1.vspdData1.focus				"				& vbCr
Response.Write "    End If								"				& vbCr
Response.Write " End With "             & vbCr		
Response.Write "</Script> "             & vbCr 

%>
 

