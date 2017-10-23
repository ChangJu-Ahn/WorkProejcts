<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!--'======================================================================================================
'*  1. Module Name          : Inventory
'*  2. Function Name        : 
'*  3. Program ID           : i1911qbb1.asp
'*  4. Program Name         : 전표차이발생수불조회 
'*  5. Program Desc         : 
'*  6. Modified date(First) : 2003/05/21
'*  7. Modified date(Last)  : 2003/05/21
'*  8. Modifier (First)     : Ahn Jung Je
'*  9. Modifier (Last)      : 
'* 10. Comment              : 
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================-->
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
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
Dim NextKey1, NextKey2
Dim strNextKey1, strNextKey2
Dim FromDate, ToDate
Dim SetComboList, ComboRow, ComboName

Const C_SHEETMAXROWS_D = 100

lgLngMaxRow     = Request("txtMaxRows") 
lgErrorStatus   = "NO"
FromDate		= UniConvDate(Request("txtTrnsFrDt"))
ToDate			= UniConvDate(Request("txtTrnsToDt"))
SetComboList    = SetComboSplit(Request("SetComboList"))

If Request("lgStrPrevKey1") <> "" Then
	strNextKey1 = Request("lgStrPrevKey")
	strNextKey2 = UniConvDate(Request("lgStrPrevKey1"))
End If

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
		lgStrPrevKeyIndex = ""		
		Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)
		Call SetErrorStatus()

		Call SubCloseRs(lgObjRs)

		Response.End 
	Else
		IntRetCD = 1

        iDx = 0
        ReDim PvArr(C_SHEETMAXROWS_D)
        
        Do While Not lgObjRs.EOF
 
            iDx = iDx + 1
        
            If iDx > C_SHEETMAXROWS_D Then
               NextKey1 = ConvSPChars(lgObjRs(1))
               NextKey2 = UNIDateClientFormat(lgObjRs(0))
               Exit Do
            End If   
	    
			For ComboRow = 0 To Ubound(SetComboList, 2)
				If UCase(Trim(SetComboList(0, ComboRow))) = UCase(Trim(lgObjRs(2)))  Then
					ComboName = Trim(SetComboList(1, ComboRow))
					Exit For
				End If
			Next
			
            lgstrData = Chr(11) & UNIDateClientFormat(lgObjRs(0)) & _
						Chr(11) & ConvSPChars(lgObjRs(1)) & _
						Chr(11) & ComboName & _
						Chr(11) & ConvSPChars(lgObjRs(3)) & _
						Chr(11) & ConvSPChars(lgObjRs(4)) & _
						Chr(11) & UniConvNumberDBToCompany(lgObjRs(5), ggAmtOfMoney.DecPoint, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit, 0) & _
						Chr(11) & UniConvNumberDBToCompany(lgObjRs(6), ggAmtOfMoney.DecPoint, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit, 0) & _
						Chr(11) & UniConvNumberDBToCompany(lgObjRs(7), ggAmtOfMoney.DecPoint, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit, 0) & _
						Chr(11) & UniConvNumberDBToCompany(lgObjRs(8), ggAmtOfMoney.DecPoint, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit, 0) & _
						Chr(11) & UniConvNumberDBToCompany(lgObjRs(9), ggAmtOfMoney.DecPoint, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit, 0) & _
						Chr(11) & UniConvNumberDBToCompany(lgObjRs(10), ggAmtOfMoney.DecPoint, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit, 0)
			
			If ConvSPChars(lgObjRs(11)) = "" Then
				lgstrData = lgstrData & Chr(11) & "N" & Chr(11) & "" & Chr(11) & "" & _
										Chr(11) & UNIDateClientFormat(lgObjRs(13)) & _
										Chr(11) & ConvSPChars(lgObjRs(14)) & _
										Chr(11) & lgLngMaxRow + iDx & Chr(11) & Chr(12)
			Else
				lgstrData = lgstrData & Chr(11) & "Y" & _
										Chr(11) & ConvSPChars(lgObjRs(11)) & _
										Chr(11) & ConvSPChars(lgObjRs(12)) & _
										Chr(11) & UNIDateClientFormat(lgObjRs(13)) & _
										Chr(11) & ConvSPChars(lgObjRs(14)) & _
										Chr(11) & lgLngMaxRow + iDx & Chr(11) & Chr(12)
			End If				
			
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
	
	lgStrSQL = " SELECT Top " & C_SHEETMAXROWS_D + 1 & _
					 " a.DOCUMENT_DT, a.ITEM_DOCUMENT_NO, a.TRNS_TYPE, a.MOV_TYPE, e.MINOR_NM," & _
				     " sum(b.AMOUNT) AMOUNT, sum(b.SUBCNTRCT_MFG_COST_AMOUNT) SUBCNTRCT_MFG_COST_AMOUNT, sum(abs(b.SALES_AMT)) SALES_AMT, " & _
				     " sum(b.AMOUNT + b.SUBCNTRCT_MFG_COST_AMOUNT + abs(b.SALES_AMT)) TOT_AMT, isnull(d.amt, 0) SLIP_AMT, " & _
				     " (sum(b.AMOUNT + b.SUBCNTRCT_MFG_COST_AMOUNT + abs(b.SALES_AMT)) - isnull(d.amt, 0)) TOT_DIF, " & _
				     " isnull(c.BATCH_NO, '') BATCH_NO,	isnull(c.GL_NO, '') GL_NO, a.POS_DT, a.BIZ_AREA_CD  " & _
				" FROM I_GOODS_MOVEMENT_HEADER a " & _
				     " inner join I_GOODS_MOVEMENT_DETAIL b " & _
				        " on a.ITEM_DOCUMENT_NO = b.ITEM_DOCUMENT_NO and a.DOCUMENT_YEAR = b.DOCUMENT_YEAR and b.DELETE_FLAG = " & FilterVar("N", "''", "S") & "  " & _
				     " left outer join A_BATCH c " & _
				        " on len(c.REF_NO) > 5 and a.ITEM_DOCUMENT_NO = left(c.REF_NO, len(c.REF_NO) - 5) and a.DOCUMENT_YEAR = right(c.REF_NO, 4) " & _
				     " left outer join (select BATCH_NO,sum(ITEM_LOC_AMT) amt  from A_BATCH_GL_ITEM where MAKE_ACCT_FG = " & FilterVar("N", "''", "S") & "  " & _ 
									   " group by BATCH_NO) d " & _
				        " on c.BATCH_NO = d.BATCH_NO" & _
				     " left outer join B_MINOR e " & _
				        " on a.MOV_TYPE = e.MINOR_CD and e.MAJOR_CD = " & FilterVar("I0001", "''", "S") & " " & _
					 " left outer join (select MOV_TYPE, GUI_CONTROL_FLAG, GUI_CONTROL_FLAG3 from I_MOVETYPE_CONFIGURATION " & _
									   " where TRNS_TYPE = " & FilterVar("ST", "''", "S") & ")  f " & _
						" on f.MOV_TYPE = a.MOV_TYPE "
	
	lgStrSQL = lgStrSQL & " WHERE ((a.TRNS_TYPE = " & FilterVar("ST", "''", "S") & " and ((f.GUI_CONTROL_FLAG = " & FilterVar("Y", "''", "S") & "  and " & _
											     " b.BIZ_AREA_CD <> (select BIZ_AREA_CD from B_PLANT where PLANT_CD = b.TRNS_PLANT_CD)) " & _ 
											    " or f.GUI_CONTROL_FLAG3 = " & FilterVar("Y", "''", "S") & "  )) " & _
									" or a.TRNS_TYPE <> " & FilterVar("ST", "''", "S") & ") " & _
							" and a.POST_FLAG = " & FilterVar("Y", "''", "S") & "  "
	
	If Trim(Request("cboTrnsType")) <> "" Then
		lgStrSQL = lgStrSQL & " and a.TRNS_TYPE = " & FilterVar(Request("cboTrnsType"), "''", "S")
	End If
	
	If strNextKey1 <> "" Then

		lgStrSQL = lgStrSQL & " and a.ITEM_DOCUMENT_NO >= " & FilterVar(strNextKey1, "''", "S") & _
							  "	and a.DOCUMENT_DT between " & FilterVar(strNextKey2, "''", "S") & " and " & FilterVar(ToDate, "''", "S") & _
							" Group by a.DOCUMENT_DT, a.ITEM_DOCUMENT_NO, a.TRNS_TYPE, a.MOV_TYPE, e.MINOR_NM,c.BATCH_NO, c.GL_NO, a.POS_DT, d.amt, a.BIZ_AREA_CD " & _
							" Having Sum(b.AMOUNT + b.SUBCNTRCT_MFG_COST_AMOUNT + abs(b.SALES_AMT)) - isnull(d.amt, 0) <> 0 " & _
							" Order by a.DOCUMENT_DT asc, a.ITEM_DOCUMENT_NO asc "    
	Else

		lgStrSQL = lgStrSQL & " and a.DOCUMENT_DT between " & FilterVar(FromDate, "''", "S") & " and " & FilterVar(ToDate, "''", "S") & _
							" GROUP BY a.DOCUMENT_DT, a.ITEM_DOCUMENT_NO, a.TRNS_TYPE, a.MOV_TYPE, e.MINOR_NM,c.BATCH_NO, c.GL_NO, a.POS_DT, d.amt, a.BIZ_AREA_CD " & _
							" HAVING Sum(b.AMOUNT + b.SUBCNTRCT_MFG_COST_AMOUNT + abs(b.SALES_AMT)) - isnull(d.amt, 0) <> 0 " & _
							" ORDER BY a.DOCUMENT_DT asc, a.ITEM_DOCUMENT_NO asc "

	End If

  
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


Function SetComboSplit(ByVal InitCombo)
	Dim ComboList
	Dim InitCode, InitName
	Dim iArrR
	
	ComboList = Split(Initcombo,Chr(12))
	InitCode  = Split(ComboList(0),Chr(11))
	InitName  = Split(ComboList(1),Chr(11))
	
	ReDim ComboList(1, Ubound(InitCode) - 1)
	
	For iArrR = 0 To Ubound(InitCode) - 1
		ComboList(0, iArrR) = InitCode(iArrR)
		ComboList(1, iArrR) = InitName(iArrR)
	Next
	SetComboSplit = ComboList
End Function
  

Response.Write "<Script language=vbs> " & vbCr         
Response.Write " With Parent "      	& vbCr
Response.Write "	If """ & lgErrorStatus & """ = ""NO"" And """ & IntRetCd & """ <> -1 Then "	& vbCr
Response.Write "    .lgStrPrevKey  = """ & NextKey1 & """" & vbCr  
Response.Write "    .lgStrPrevKey1  = """ & NextKey2 & """" & vbCr  
Response.Write "	.ggoSpread.Source	= .frm1.vspdData "				& vbCr
Response.Write "	.ggoSpread.SSShowData  """ & lgstrData  & """"        & vbCr
Response.Write "		If .frm1.vspdData.MaxRows < .parent.VisibleRowCnt(.frm1.vspdData, 0) And .lgStrPrevKey <> """" Then "	& vbCr
Response.Write "			.DbQuery						"				& vbCr
Response.Write "		Else								"				& vbCr
Response.Write "			.DbQueryOK						"				& vbCr
Response.Write "		End If								"				& vbCr
Response.Write "		.frm1.vspdData.focus				"				& vbCr
Response.Write "    End If								"				& vbCr
Response.Write " End With "             & vbCr		
Response.Write "</Script> "             & vbCr 

%>
 

