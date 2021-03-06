<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../../ComAsp/LoadinfTB19029.asp" -->
<%
'**********************************************************************************************
'*  1. Module Name          : Quality Management
'*  2. Function Name        : 
'*  3. Program ID           : Q2111QB1
'*  4. Program Name         : 일보조회 
'*  5. Program Desc         : 
'*  6. Component List       : 
'*  7. Modified date(First) : 2002/05/14
'*  8. Modified date(Last)  : 2003/08/04
'*  9. Modifier (First)     : Koh Jae Woo
'* 10. Modifier (Last)      : Ahn Jung Je
'* 11. Comment
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change" 
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************

On Error Resume Next

Call HideStatusWnd															'☜: 모든 작업 완료후 작업진행중 표시창을 Hide
Err.Clear

Call LoadBasisGlobalInf
Call LoadinfTB19029B("Q", "Q", "NOCOOKIE", "QB")

Dim IntRetCD
Dim PvArr
Dim NextKey1
Dim strNextKey1

Const C_SHEETMAXROWS_D = 100

lgLngMaxRow     = Request("txtMaxRows") 
lgErrorStatus   = "NO"

Call HideStatusWnd 

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
 
            If iDx = C_SHEETMAXROWS_D Then
               NextKey1 = ConvSPChars(lgObjRs(0))
               Exit Do
            End If   
	    
            lgstrData = Chr(11) & ConvSPChars(lgObjRs(0)) & _
						Chr(11) & ConvSPChars(lgObjRs(1)) & _
						Chr(11) & UNIDateClientFormat(lgObjRs(2)) & _
						Chr(11) & UNIDateClientFormat(lgObjRs(3)) & _
						Chr(11) & ConvSPChars(lgObjRs(4)) & _
						Chr(11) & ConvSPChars(lgObjRs(5)) & _
						Chr(11) & ConvSPChars(lgObjRs(6)) & _
						Chr(11) & ConvSPChars(lgObjRs(7)) & _
						Chr(11) & ConvSPChars(lgObjRs(8)) & _
						Chr(11) & ConvSPChars(lgObjRs(9)) & _
						Chr(11) & ConvSPChars(lgObjRs(10)) & _
						Chr(11) & ConvSPChars(lgObjRs(11)) & _
						Chr(11) & UniConvNumberDBToCompany(lgObjRs(12), ggQty.DecPoint , ggQty.RndPolicy, ggQty.RndUnit, 0) & _
						Chr(11) & UniConvNumberDBToCompany(lgObjRs(13), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0) & _
						Chr(11) & UniConvNumberDBToCompany(lgObjRs(14), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0) & _
                        Chr(11) & UNINumClientFormat(lgObjRs(15), 2, 0) & _
						Chr(11) & lgLngMaxRow + iDx & Chr(11) & Chr(12)
			
			PvArr(iDx) = lgstrData
			iDx = iDx + 1
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
					 " a.INSP_REQ_NO, a.INSP_RESULT_NO, a.Release_DT, a.INSP_DT, a.ITEM_CD, b.ITEM_NM, a.BP_CD, c.BP_NM," & _
				     " a.LOT_NO, a.LOT_SUB_NO, a.DECISION, d.MINOR_NM, a.LOT_SIZE, a.INSP_QTY, a.DEFECT_QTY, " & _
				     " CASE WHEN A.INSP_QTY <> 0 THEN CAST((A.DEFECT_QTY/A.INSP_QTY) * 100 AS  NUMERIC(15,2)) END " & _
				" FROM Q_INSPECTION_RESULT a " & _
				     " left outer join B_ITEM b ON a.ITEM_CD = b.ITEM_CD " & _
				     " left outer join B_BIZ_PARTNER c ON a.BP_CD = c.BP_CD " & _
				     " left outer join B_MINOR d ON d.MAJOR_CD = 'Q0010' AND a.DECISION = d.MINOR_CD " & _
			   " WHERE a.INSP_CLASS_CD = 'R' AND a.PLANT_CD = " & FilterVar(Request("txtPlantCd"),"","S") & _
				 " AND a.Release_DT between " & FilterVar(UNIConvDate(Request("txtDtFr")),"","S") & " AND " & FilterVar(UNIConvDate(Request("txtDtTo")),"","S")
	
	
	If Trim(Request("txtItemCd")) <> "" Then
		lgStrSQL = lgStrSQL & " and a.ITEM_CD = " & FilterVar(Request("txtItemCd"),"","S")
	End If
	If Trim(Request("txtBpCd")) <> "" Then
		lgStrSQL = lgStrSQL & " and a.BP_CD = " & FilterVar(Request("txtBpCd"),"","S")
	End If
	If Trim(Request("cboDecision")) <> "" Then
		lgStrSQL = lgStrSQL & " and a.DECISION = " & FilterVar(Request("cboDecision"),"","S")
	End If
	
	If Trim(Request("txtLotNo")) <> "" Then
		lgStrSQL = lgStrSQL & " and a.LOT_NO >= " & FilterVar(Request("txtLotNo"),"","S")
	End If
	
	If Request("lgStrPrevKey") <> "" Then
		lgStrSQL = lgStrSQL & " and a.INSP_REQ_NO >= " & FilterVar(Request("lgStrPrevKey"),"","S")
	ElseIf Trim(Request("txtInspReqNo")) <> "" Then
		lgStrSQL = lgStrSQL & " and a.INSP_REQ_NO >= " & FilterVar(Request("txtInspReqNo"),"","S")
	End If
		
	lgStrSQL = lgStrSQL & " ORDER BY a.INSP_REQ_NO asc "

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
Response.Write "    .lgStrPrevKey  = """ & NextKey1 & """" & vbCr  
Response.Write "	.ggoSpread.Source	= .frm1.vspdData "				& vbCr
Response.Write "	.ggoSpread.SSShowDataByClip  """ & lgstrData  & """"        & vbCr
Response.Write "		If .frm1.vspdData.MaxRows < .parent.VisibleRowCnt(.frm1.vspdData, 0) And .lgStrPrevKey <> """" Then "	& vbCr
Response.Write "			.DbQuery						"				& vbCr
Response.Write "		Else								"				& vbCr
Response.Write "			.DbQueryOK						"				& vbCr
Response.Write "		End If								"				& vbCr
Response.Write "		.frm1.vspdData.focus				"				& vbCr
Response.Write "    End If								"				& vbCr
Response.Write " End With "             & vbCr		
Response.Write "</Script> "             & vbCr 
Response.End     

%>    
