<%@ LANGUAGE=VBSCript%>
<!--
'**********************************************************************************************
'*  1. Module Name          : Child reservation Information
'*  2. Function Name        : 
'*  3. Program ID           : I1524rb1.asp
'*  4. Program Name         : 
'*  5. Program Desc         : VMI 입고현황 참조화면 
'*  6. Comproxy List        : 
'                             
'*  7. Modified date(First) : 2003/01/13
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Lee Seung Wook
'* 10. Modifier (Last)      : Lee Seung Wook
'* 11. Comment              :
'**********************************************************************************************
-->

<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->

<%		
Call LoadBasisGlobalInf()
Call LoadInfTB19029B("Q", "I","NOCOOKIE","RB")   
Call HideStatusWnd 
											
On Error Resume Next
Err.Clear

         
Dim strPlantCd
Dim strRcptNo
Dim strItemDocumentNo
Dim strTrnsType
Dim strDocumentDt

Dim IntRetCD

lgLngMaxRow       = Cint(Request("txtMaxRows"))         
lgMaxCount        = 100                             
lgStrPrevKeyIndex = UNICInt(Trim(Request("lgStrPrevKeyIndex")),0)   
lgErrorStatus     = "NO"
lgErrorPos        = ""                                               

Call SubOpenDB(lgObjConn)
Call SubCreateCommandObject(lgObjComm)


strPlantCd	= FilterVar(Request("txtPlantCd"), "''", "S")
strRcptNo   = FilterVar(Request("txtRcptNo"), "''", "S")

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
    
    
	If pType = "AL" Then
		Call SubMakeSQLStatements("AL",strPlantCd,strRcptNo)          
	End If
	
	'---------------------------
	' Header Single 조회 
	'---------------------------    
				
	If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then                   
		IntRetCD = -1
		lgStrPrevKeyIndex = ""		
		Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)    
		Call SetErrorStatus()
		
		Call SubCloseRs(lgObjRs)
%>
<Script Language="VBScript">
		parent.frm1.txtItemDocumentNo.value		= "<%=strItemDocumentNo%>"
		parent.frm1.txtTrnsType.value			= "<%=strTrnsType%>"
		parent.frm1.txtDocumentDt.value			= "<%=strDocumentDt%>"
</Script>	
<%
		Response.End 
	Else
		
		IntRetCD = 1
		 Call SubSkipRs(lgObjRs,lgMaxCount * lgStrPrevKeyIndex)
        lgstrData = ""
        iDx       = 1
        ReDim PvArr(0)
        
        strItemDocumentNo	= ConvSPChars(lgObjRs(12))
		strTrnsType			= ConvSPChars(lgObjRs(13))	
		strDocumentDt		= UNIDateClientFormat(lgObjRs(14))
        
        Do While Not lgObjRs.EOF
        
        	lgstrData =		Chr(11) & ConvSPChars(lgObjRs(0)) & _
							Chr(11) & ConvSPChars(lgObjRs(1)) & _
							Chr(11) & ConvSPChars(lgObjRs(2)) & _
							Chr(11) & ConvSPChars(lgObjRs(3)) & _
							Chr(11) & ConvSPChars(lgObjRs(4)) & _
							Chr(11) & UniNumClientFormat(lgObjRs(5),ggQty.DecPoint,0) & _
							Chr(11) & ConvSPChars(lgObjRs(6)) & _
							Chr(11) & ConvSPChars(lgObjRs(7)) & _
							Chr(11) & ConvSPChars(lgObjRs(8)) & _
							Chr(11) & ConvSPChars(lgObjRs(9)) & _
							Chr(11) & ConvSPChars(lgObjRs(10)) & _
							Chr(11) & ConvSPChars(lgObjRs(11)) & _
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
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pDataType,pCode,pCode1)

    On Error Resume Next                                                     
    Err.Clear                                                                
    
	Dim iSelCount
	
Select Case pDataType

		Case "AL"
			lgStrSQL = "SELECT	A.bp_cd,D.bp_nm,A.item_cd,C.item_nm,A.base_unit,A.qty,A.sl_cd,A.trns_sl_cd,A.tracking_no,A.lot_no,A.lot_sub_no,C.spec,B.item_document_no,E.minor_nm,B.document_dt"
			lgStrSQL = lgStrSQL & " FROM I_VMI_GOODS_MVMT_DTL A INNER JOIN  I_VMI_GOODS_MVMT_HDR B"
			lgStrSQL = lgStrSQL & "		on A.ITEM_DOCUMENT_NO = B.ITEM_DOCUMENT_NO and A.DOCUMENT_YEAR = B.DOCUMENT_YEAR "
			lgStrSQL = lgStrSQL & "		LEFT OUTER JOIN B_ITEM C on A.ITEM_CD = C.ITEM_CD "
			lgStrSQL = lgStrSQL & "		LEFT OUTER JOIN B_BIZ_PARTNER D on A.BP_CD = D.BP_CD "
			lgStrSQL = lgStrSQL & "		LEFT OUTER JOIN B_MINOR E on A.TRNS_TYPE = E.MINOR_CD and E.MAJOR_CD = " & FilterVar("I0006", "''", "S") & " "
			lgStrSQL = lgStrSQL & "	WHERE 	A.PLANT_CD = " & strPlantCd
			lgStrSQL = lgStrSQL & "	AND		B.MVMT_RCPT_NO = " & strRcptNo
			lgStrSQL = lgStrSQL & " ORDER BY A.item_cd "
			
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
		.frm1.txtItemDocumentNo.value	= "<%=strItemDocumentNo%>"
		.frm1.txtTrnsType.value			= "<%=strTrnsType%>"
		.frm1.txtDocumentDt.value		= "<%=strDocumentDt%>"
		
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


	
