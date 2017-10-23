<%@ LANGUAGE=VBSCript%>
<%Option Explicit%>
<!--'**********************************************************************************************
'*  1. Module Name          : Inventory
'*  2. Function Name        : 
'*  3. Program ID           : I1521pb1.asp
'*  4. Program Name         : 
'*  5. Program Desc         : VMI 수불번호 팝업 
'*  6. Comproxy List        : 
'                             
'*  7. Modified date(First) : 2003/01/11
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Lee Seung Wook
'* 10. Modifier (Last)      : Lee Seung Wook
'* 11. Comment              :
'**********************************************************************************************-->
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%		
Call LoadBasisGlobalInf
Call LoadInfTB19029B("Q", "I","NOCOOKIE","PB")   
Call HideStatusWnd 
											
On Error Resume Next
Err.Clear

lgLngMaxRow       = Cint(Request("txtMaxRows"))                                   
lgMaxCount        = 100                             
lgStrPrevKeyIndex = UNICInt(Trim(Request("lgStrPrevKeyIndex")),0)   
lgErrorStatus     = "NO"
lgErrorPos        = ""                                                          

Dim IntRetCD
Dim strPlantCd
Dim strItemDocumentNo
Dim strIDocumentYear
Dim strDocumentDt1
Dim strDocumentDt2
Dim strTrnsType

strPlantCd			= FilterVar(Request("PlantCd"), "''", "S")
strItemDocumentNo   = FilterVar(Request("txtItemDocumentNo"), "''", "S")
strIDocumentYear	= FilterVar(Request("txtDocumentYear"), "''", "S")

If Request("txtDocumentDt1") <> "" Then
	strDocumentDt1		= FilterVar(UNIConvDate(Request("txtDocumentDt1")), "''", "S")
End if

If Request("txtDocumentDt2") <> "" Then	
	strDocumentDt2		= FilterVar(Trim(UNIConvDate(Request("txtDocumentDt2")))	, "''", "S")
End if

strTrnsType			= FilterVar(Request("txtTrnsType"), "''", "S")

Call SubOpenDB(lgObjConn)
Call SubCreateCommandObject(lgObjComm)

Call SubBizQuery("AL") 

Call SubCloseCommandObject(lgObjComm)    
Call SubCloseDB(lgObjConn)      

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery(pType)
	Dim iDx
	Dim PvArr	
	
	On Error Resume Next                                                           
    Err.Clear
    '---------- Developer Coding part (Start) ---------------------------------------------------------------
	Call SubMakeSQLStatements("AL",strPlantCd,strItemDocumentNo,strIDocumentYear,strDocumentDt1,strDocumentDt2,strTrnsType)           '☜ : Make sql statements
		
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
        lgstrData = ""
        iDx       = 1
        ReDim PvArr(0)
        
        Do While Not lgObjRs.EOF
			
	        lgstrData = Chr(11) & ConvSPChars(lgObjRs(0)) & _
						Chr(11) & ConvSPChars(lgObjRs(1)) & _
						Chr(11) & UNIDateClientFormat(lgObjRs(2)) & _
						Chr(11) & ConvSPChars(lgObjRs(3)) & _
						Chr(11) & ConvSPChars(lgObjRs(4)) & _
						Chr(11) & ConvSPChars(lgObjRs(5)) & _
						Chr(11) & ConvSPChars(lgObjRs(6)) & _
						Chr(11) & ConvSPChars(lgObjRs(7)) & _
						Chr(11) & lgLngMaxRow + iDx & Chr(11) & Chr(12)

		    lgObjRs.MoveNext
	        ReDim Preserve PvArr(iDx-1)
			PvArr(iDx-1) = lgstrData
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
'============================================================================================================
Sub SubMakeSQLStatements(pDataType,pCode,pCode1,pCode2,pCode3,pCode4,pCode5)

    On Error Resume Next                                                           
    Err.Clear                                                                      
    
	Dim iSelCount
	
	Select Case pDataType
		
		Case "AL"
			lgStrSQL =	"SELECT 	A.item_document_no, A.Document_year, A.document_dt, A.plant_cd, B.plant_nm, A.bp_cd, C.bp_nm, A.document_text"	& _
						" FROM 		I_VMI_GOODS_MVMT_HDR A, B_PLANT B, B_BIZ_PARTNER C"	& _
						" WHERE		A.plant_cd 			=	B.plant_cd"					& _
						" AND		A.bp_cd				=	C.bp_cd"					& _
						" AND		A.delete_flag		=	"	 & "" & FilterVar("N", "''", "S") & " "				& _
						" AND		A.plant_cd			=	"	  & strPlantCd			& _
						" AND		A.item_document_no	>=		" & strItemDocumentNo	& _
						" AND		A.document_year		>=		" & strIDocumentYear
			If strDocumentDt1 <> "" Then
				lgStrSQL = lgStrSQL & " AND		A.document_dt		>=		" & strDocumentDt1
			End If
			If strDocumentDt2 <> "" Then
				lgStrSQL = lgStrSQL & " AND		A.document_dt		<=		" & strDocumentDt2
			End If
			lgStrSQL = lgStrSQL & " AND		A.trns_type			=		" & strTrnsType
			lgStrSQL = lgStrSQL & " ORDER BY A.item_document_no "
	End Select    
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
%>
<Script Language="VBScript">
	With parent
		
		If "<%=lgErrorStatus%>" = "NO" And "<%=IntRetCd%>" <> -1 Then
		    .ggoSpread.Source	= .frm1.vspdData
			.lgStrPrevKeyIndex	= "<%=lgStrPrevKeyIndex%>"
			.ggoSpread.SSShowData "<%=lgstrData%>"

        	If .frm1.vspdData.MaxRows < .VisibleRowCnt(.frm1.vspdData,0)  And .lgStrPrevKeyIndex <> "" Then	
				.DbQuery
			Else
				.DbQueryOk
			End If
			.frm1.vspdData.focus
		End If   

	End With	
       
</Script>	


	
