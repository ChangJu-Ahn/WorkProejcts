<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!--
'**********************************************************************************************
'*  1. Module Name            : Inventory
'*  2. Function Name          : 
'*  3. Program ID             : i1503qa1.asp
'*  4. Program Name           : 
'*  5. Program Desc           : 재고이동현황 조회(공장간)
'*  6. Comproxy List          :      
'*  7. Modified date(First)   : 2003/06/30
'*  8. Modified date(Last)    : 2003/06/30
'*  9. Modifier (First)       : Lee Seung Wook
'* 10. Modifier (Last)        : 
'* 11. Comment                :
'**********************************************************************************************
-->
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
Dim strPlantCd,FromDate,ToDate,strTrnsPlCd,strMovType,strItem
Dim ComboRow, ComboName


lgLngMaxRow       = Request("txtMaxRows")
lgMaxCount        = 100                  
lgStrPrevKeyIndex = UNICInt(Trim(Request("lgStrPrevKeyIndex")),0)   
lgErrorStatus     = "NO"
lgErrorPos        = ""                   
lgOpModeCRUD      = Request("txtMode") 

strPlantCd		= FilterVar(Request("txtPlantCd"), "''", "S")
FromDate		= FilterVar(UniConvDate(Request("txtMovFrDt")), "''", "S")
ToDate			= FilterVar(UniConvDate(Request("txtMovToDt")), "''", "S")
strTrnsPlCd		= FilterVar("%" & Trim(Request("txtTrnsPlantCd")) & "%", "''", "S")
strMovType		= FilterVar(Request("txtMovType"), "''", "S")
strItem			= FilterVar(Request("txtItemCd"), "''", "S")

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
		Call SubSkipRs(lgObjRs,lgMaxCount * lgStrPrevKeyIndex)
		
		lgstrData	= ""
		iDx			= 1

        ReDim PvArr(0)
        
        Do While Not lgObjRs.EOF
			
			ReDim Preserve PvArr(iDx-1)
        
            lgstrData = Chr(11) & ConvSPChars(lgObjRs(0)) & _
						Chr(11) & ConvSPChars(lgObjRs(1)) & _
						Chr(11) & ConvSPChars(lgObjRs(2)) & _
						Chr(11) & ConvSPChars(lgObjRs(3)) & _
						Chr(11) & ConvSPChars(lgObjRs(4)) & _
						Chr(11) & UniConvNumberDBToCompany(lgObjRs(5),ggQty.DecPoint,ggQty.RndPolicy, ggQty.RndUnit, 0) & _
						Chr(11) & ConvSPChars(lgObjRs(6)) & _
						Chr(11) & ConvSPChars(lgObjRs(7)) & _
						Chr(11) & UNIDateClientFormat(lgObjRs(8)) & _
						Chr(11) & UniConvNumberDBToCompany(lgObjRs(9),ggQty.DecPoint,ggQty.RndPolicy, ggQty.RndUnit, 0) & _
						Chr(11) & UniConvNumberDBToCompany(lgObjRs(10),ggUnitCost.DecPoint,ggUnitCost.RndPolicy, ggUnitCost.RndUnit, 0) & _
						Chr(11) & UniConvNumberDBToCompany(lgObjRs(11), ggAmtOfMoney.DecPoint, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit, 0) & _
						Chr(11) & ConvSPChars(lgObjRs(12)) & _
						Chr(11) & ConvSPChars(lgObjRs(13)) & _
						Chr(11) & ConvSPChars(lgObjRs(14)) & _
						Chr(11) & ConvSPChars(lgObjRs(15)) & _
						Chr(11) & ConvSPChars(lgObjRs(16)) & _
						Chr(11) & ConvSPChars(lgObjRs(17)) & _
						Chr(11) & lgLngMaxRow + iDx & Chr(11) & Chr(12) 
			
			PvArr(iDx - 1) = lgstrData
		    lgObjRs.MoveNext
		    
		    iDx =  iDx + 1
	        If iDx > lgMaxCount Then
	           lgStrPrevKeyIndex = lgStrPrevKeyIndex + 1
	           Exit Do
	        End If   
		    
        Loop
        
        lgstrData = Join(PvArr, "")

    End If
    
    If iDx <= lgMaxCount Then
	   lgStrPrevKeyIndex = ""
	End If 

	Call SubHandleError("MR",lgObjConn,lgObjRs,Err)
    Call SubCloseRs(lgObjRs)  
    
    lgStrSQL = ""                                           
		
End Sub

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements()

    On Error Resume Next
    Err.Clear
    
    lgStrSQL	=	" SELECT	A.item_cd,C.item_nm,C.spec,A.tracking_no,A.lot_no,A.lot_sub_no,A.base_unit,D.minor_nm,B.document_dt," & _
					"			A.qty,A.price,A.amount,A.sl_cd,A.trns_plant_cd,f.plant_nm,A.trns_sl_cd,A.item_document_no,A.seq_no " & _
					" FROM		I_GOODS_MOVEMENT_DETAIL A inner join I_GOODS_MOVEMENT_HEADER B" & _
					"		 	on a.ITEM_DOCUMENT_NO = b.ITEM_DOCUMENT_NO and a.DOCUMENT_YEAR = b.DOCUMENT_YEAR and a.DELETE_FLAG = " & FilterVar("N", "''", "S") & " " & _
					"			left outer join B_ITEM C on a.item_cd = c.item_cd" & _
					"			left outer join B_MINOR D on a.MOV_TYPE = d.MINOR_CD and d.MAJOR_CD = " & FilterVar("I0001", "''", "S") & "" & _
					"			left outer join (select MOV_TYPE, GUI_CONTROL_FLAG, GUI_CONTROL_FLAG2 from I_MOVETYPE_CONFIGURATION where TRNS_TYPE = " & FilterVar("ST", "''", "S") & ") E" & _
					"			on e.MOV_TYPE = a.MOV_TYPE " & _
					"			left outer join B_PLANT F on a.trns_plant_cd = f.plant_cd " & _
					" WHERE		A.AUTO_CRTD_FLAG = " & FilterVar("N", "''", "S") & " "		& _
					" AND		A.TRNS_TYPE = " & FilterVar("ST", "''", "S") & ""			& _	
					" AND		E.GUI_CONTROL_FLAG = " & FilterVar("Y", "''", "S") & " "	& _
					" AND		A.PLANT_CD = "				& strPlantCd & _
					" AND		(B.DOCUMENT_DT BETWEEN " & FromDate & " AND " & ToDate & " ) " & _
					" AND		A.MOV_TYPE >= "				& strMovType & _ 
					" AND		A.ITEM_CD >= "				& strItem & _
					" AND		A.TRNS_PLANT_CD like"		& strTrnsPlCd & _
					" ORDER BY	A.ITEM_CD ASC,B.DOCUMENT_DT ASC, A.SEQ_NO ASC "	 
	
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
  

Response.Write "<Script language=vbs> "																								& vbCr         
Response.Write " With Parent "      																								& vbCr
Response.Write "	If """ & lgErrorStatus & """ = ""NO"" And """ & IntRetCd & """ <> -1 Then "										& vbCr
Response.Write "	.ggoSpread.Source	= .frm1.vspdData "																			& vbCr
Response.Write "	.lgStrPrevKeyIndex	=  """ & lgStrPrevKeyIndex & """"															& vbCr	 
Response.Write "	.ggoSpread.SSShowData  """ & lgstrData  & """"																	& vbCr
Response.Write "		If .frm1.vspdData.MaxRows < .parent.VisibleRowCnt(.frm1.vspdData, 0) And .lgStrPrevKeyIndex <> """" Then "	& vbCr
Response.Write "			.DbQuery						"																		& vbCr
Response.Write "		Else								"																		& vbCr
Response.Write "			.DbQueryOK						"																		& vbCr
Response.Write "		End If								"																		& vbCr
Response.Write "		.frm1.vspdData.focus				"																		& vbCr
Response.Write "    End If								"																			& vbCr
Response.Write " End With "             & vbCr		
Response.Write "</Script> "             & vbCr 

%>
 

