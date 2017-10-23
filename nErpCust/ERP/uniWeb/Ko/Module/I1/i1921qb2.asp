<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!--'======================================================================================================
'*  1. Module Name          : Inventory
'*  2. Function Name        : 
'*  3. Program ID           : i1921qb2.asp
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
        
            lgstrData = Chr(11) & ConvSPChars(lgObjRs(0)) & _
						Chr(11) & ConvSPChars(lgObjRs(1)) & _
						Chr(11) & ConvSPChars(lgObjRs(2)) & _
						Chr(11) & UniConvNumberDBToCompany(lgObjRs(3), ggAmtOfMoney.DecPoint, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit, 0) & _
						Chr(11) & UniConvNumberDBToCompany(lgObjRs(4), ggAmtOfMoney.DecPoint, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit, 0) & _
						Chr(11) & UniConvNumberDBToCompany(lgObjRs(5), ggAmtOfMoney.DecPoint, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit, 0) & _
						Chr(11) & UniConvNumberDBToCompany(lgObjRs(6), ggAmtOfMoney.DecPoint, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit, 0) & _
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
	
	lgStrSQL = " SELECT (Select MINOR_NM From B_MINOR Where MINOR_CD = b.TRNS_TYPE and MAJOR_CD = " & FilterVar("I0002", "''", "S") & " ), b.MOV_TYPE, " & _
					  " (Select MINOR_NM From B_MINOR Where MINOR_CD = b.MOV_TYPE and MAJOR_CD = " & FilterVar("I0001", "''", "S") & "), " &_ 
            		  " SUM(CASE WHEN b.TRNS_TYPE in (" & FilterVar("MR", "''", "S") & "," & FilterVar("PR", "''", "S") & ", " & FilterVar("OR", "''", "S") & ", " & FilterVar("ST", "''", "S") & ")  AND b.DEBIT_CREDIT_FLAG= " & FilterVar("D", "''", "S") & "  THEN b.AMOUNT " & _
						   	   " WHEN b.TRNS_TYPE in (" & FilterVar("MR", "''", "S") & "," & FilterVar("PR", "''", "S") & ", " & FilterVar("OR", "''", "S") & ")  AND b.DEBIT_CREDIT_FLAG= " & FilterVar("C", "''", "S") & "  THEN b.AMOUNT*(-1) ELSE 0 END ) AS R_AMT, " & _
					  " SUM(CASE WHEN b.TRNS_TYPE in (" & FilterVar("PI", "''", "S") & ", " & FilterVar("DI", "''", "S") & ", " & FilterVar("OI", "''", "S") & ", " & FilterVar("ST", "''", "S") & ")  AND b.DEBIT_CREDIT_FLAG= " & FilterVar("C", "''", "S") & "  THEN b.AMOUNT " & _
						       " WHEN b.TRNS_TYPE in (" & FilterVar("PI", "''", "S") & ", " & FilterVar("DI", "''", "S") & ", " & FilterVar("OI", "''", "S") & ")  AND b.DEBIT_CREDIT_FLAG= " & FilterVar("D", "''", "S") & "  THEN b.AMOUNT*(-1) ELSE 0 END ) AS I_AMT, " & _
            		  " SUM(CASE WHEN b.TRNS_TYPE in (" & FilterVar("MR", "''", "S") & "," & FilterVar("PR", "''", "S") & ", " & FilterVar("OR", "''", "S") & ", " & FilterVar("ST", "''", "S") & ")  AND b.DEBIT_CREDIT_FLAG= " & FilterVar("D", "''", "S") & "  THEN b.COST_OF_DEVY " & _
						   	   " WHEN b.TRNS_TYPE in (" & FilterVar("MR", "''", "S") & "," & FilterVar("PR", "''", "S") & ", " & FilterVar("OR", "''", "S") & ")  AND b.DEBIT_CREDIT_FLAG= " & FilterVar("C", "''", "S") & "  THEN b.COST_OF_DEVY * (-1) ELSE 0 END ) AS COST_DEVY, " & _
            		  " SUM(CASE WHEN b.TRNS_TYPE in (" & FilterVar("MR", "''", "S") & "," & FilterVar("PR", "''", "S") & ", " & FilterVar("OR", "''", "S") & ", " & FilterVar("ST", "''", "S") & ")  AND b.DEBIT_CREDIT_FLAG= " & FilterVar("D", "''", "S") & "  THEN b.SUBCNTRCT_MFG_COST_AMOUNT " & _
						   	   " WHEN b.TRNS_TYPE in (" & FilterVar("MR", "''", "S") & "," & FilterVar("PR", "''", "S") & ", " & FilterVar("OR", "''", "S") & ")  AND b.DEBIT_CREDIT_FLAG= " & FilterVar("C", "''", "S") & "  THEN b.SUBCNTRCT_MFG_COST_AMOUNT * (-1) ELSE 0 END ) AS SUB_MFG " & _
				" FROM I_GOODS_MOVEMENT_HEADER a " & _
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
							" and c.ITEM_ACCT = " & FilterVar(Request("txtItemAcct"), "''", "S") & _
						  " Group by b.TRNS_TYPE, b.MOV_TYPE " & _
						  " ORDER BY b.TRNS_TYPE, b.MOV_TYPE "

  
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
Response.Write "		.ggoSpread.Source	= .frm1.vspdData2 "				& vbCr
Response.Write "		.ggoSpread.SSShowData  """ & lgstrData  & """"        & vbCr
Response.Write "		.DbDtlQueryOK						"				& vbCr
Response.Write "    End If								"				& vbCr
Response.Write " End With "             & vbCr		
Response.Write "</Script> "             & vbCr 

%>
 

