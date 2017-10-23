<%@ LANGUAGE=VBSCript%>
<!--
'**********************************************************************************************
'*  1. Module Name          : Child reservation Information
'*  2. Function Name        : 
'*  3. Program ID           : I1321rb1.asp
'*  4. Program Name         : 
'*  5. Program Desc         : 사급품 출고 예정정보 
'*  6. Comproxy List        : 
'                             +B25018ListPlant
'                             +B25019LookUpPlant
'*  7. Modified date(First) : 2002/12/02
'*  8. Modified date(Last)  : 2003/06/03
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

<%		
Call LoadBasisGlobalInf()
Call LoadInfTB19029B("Q", "I","NOCOOKIE","RB")   
Call HideStatusWnd 
											
On Error Resume Next
Err.Clear

         
Dim strPlantCd
Dim strPlantNm
Dim strSlFrCd
Dim strSlFrNm
Dim strSlToCd
Dim strSlToNm
Dim strBpCd
Dim strBpNm
Dim strCondBpCd

Dim IntRetCD

Dim lgStrSQL2
Dim lgStrSQL3
Dim lgStrSQL4
Dim lgStrSQL5

lgLngMaxRow       = Cint(Request("txtMaxRows"))                                    
lgMaxCount        = 100                               
lgStrPrevKeyIndex = UNICInt(Trim(Request("lgStrPrevKeyIndex")),0)   
lgErrorStatus     = "NO"
lgErrorPos        = ""                                

Call SubOpenDB(lgObjConn)
Call SubCreateCommandObject(lgObjComm)


strPlantCd	= FilterVar(Request("txtPlantCd"), "''", "S")
strSlFrCd   = FilterVar(Request("txtSlFrCd"), "''", "S")
strSlToCd	= FilterVar(Request("txtSlToCd"), "''", "S")

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
    
lgStrSQL4 =		" SELECT Sl_Nm"					& _
				" FROM B_STORAGE_LOCATION"		& _
				" WHERE sl_type = " & "" & FilterVar("E", "''", "S") & " "		& _
				" AND sl_cd = "		& strSlToCd	
    
If FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL4,"X","X") = False Then
	intCondRet = -1
	lgStrPrevKeyIndex = ""
	Call DisplayMsgBox("169961",vbInformation, "", "",I_MKSCRIPT)
	Call SetErrorStatus()
	Response.End
End IF
	strSlToNm = ConvSPChars(lgObjRs(0))
Call SubCloseRs(lgObjRs)
	
lgStrSQL5 =		" SELECT A.Bp_cd,B.Bp_Nm"						& _
				" FROM B_STORAGE_LOCATION A, B_BIZ_PARTNER B"	& _
				" WHERE A.Bp_Cd = B.Bp_cd "						& _
				" AND A.sl_cd = " & strSlToCd
	
If FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL5,"X","X") = False Then
	IntRetCD = -1
	lgStrPrevKeyIndex = ""
	Call DisplayMsgBox("169960",vbInformation,"","",I_MKSCRIPT)
	Call SetErrorStatus()
	Response.End
End IF
	
strCondBpCd = FilterVar(ConvSPChars(lgObjRs(0)), "''", "S")
strBpCd = ConvSPChars(lgObjRs(0))
strBpNm = ConvSPChars(lgObjRs(1))
Call SubCloseRs(lgObjRs)
    
    
	If pType = "AL" Then
		Call SubMakeSQLStatements("AL",strPlantCd,strSlFrCd,strSlToCd,strBpCd)     
	End If
		
	If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL2,"X","X") = False Then   
		intCondRet = -1
		lgStrPrevKeyIndex = ""		
		Call DisplayMsgBox("125000", vbInformation, "", "", I_MKSCRIPT)   
		Call SetErrorStatus()
		Response.End 
	End If
	
	strPlantNm = ConvSPChars(lgObjRs(0))
	Call SubCloseRs(lgObjRs)
	
	If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL3,"X","X") = False Then    
		intCondRet = -1
		lgStrPrevKeyIndex = ""		
		Call DisplayMsgBox("169922", vbInformation, "", "", I_MKSCRIPT)   
		Call SetErrorStatus()
		Response.End 
	End If
	
	strSlFrNm = ConvSPChars(lgObjRs(0))
	Call SubCloseRs(lgObjRs)
	
				
	If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then    
		IntRetCD = -1
		lgStrPrevKeyIndex = ""		
		Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)  
		Call SetErrorStatus()

		Call SubCloseRs(lgObjRs)
%>
<Script Language="VBScript">
		parent.frm1.txtPlantNm.value	= "<%=strPlantNm%>"
		parent.frm1.txtSlFrNm.value		= "<%=strSlFrNm%>"
		parent.frm1.txtSlToNm.value		= "<%=strSlToNm%>"
		parent.frm1.txtBpCd.value		= "<%=strBpCd%>"
		parent.frm1.txtBpNm.value		= "<%=strBpNm%>"
</Script>	
<%
		Response.End 
	Else
		
		IntRetCD = 1
		 Call SubSkipRs(lgObjRs,lgMaxCount * lgStrPrevKeyIndex)
        lgstrData = ""
        iDx       = 1
        
        ReDim PvArr(0)
        
        Do While Not lgObjRs.EOF
			
			ReDim Preserve PvArr(iDx - 1)
        
            lgstrData = Chr(11) & ConvSPChars(lgObjRs(0)) & _
						Chr(11) & ConvSPChars(lgObjRs(1)) & _
						Chr(11) & ConvSPChars(lgObjRs(2)) & _
						Chr(11) & UniNumClientFormat(lgObjRs(3),ggQty.DecPoint,0) & _
						Chr(11) & ConvSPChars(lgObjRs(4)) & _
						Chr(11) & UniNumClientFormat(lgObjRs(5),ggQty.DecPoint,0) & _
						Chr(11) & UniNumClientFormat(lgObjRs(6),ggQty.DecPoint,0) & _
						Chr(11) & UniNumClientFormat(lgObjRs(7),ggQty.DecPoint,0) & _
						Chr(11) & UniNumClientFormat(lgObjRs(8),ggQty.DecPoint,0) & _
						Chr(11) & ConvSPChars(lgObjRs(9)) & _
						Chr(11) & ConvSPChars(lgObjRs(10)) & _
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
Sub SubMakeSQLStatements(pDataType,pCode,pCode1,pCode2,pCode3)

    On Error Resume Next                                                         
    Err.Clear                                                                    
    
	Dim iSelCount
	
	lgStrSQL2 =		" SELECT Plant_Nm"					& _
					" FROM B_PLANT"						& _
					" WHERE plant_cd = " & strPlantCd
	
	lgStrSQL3 =		" SELECT Sl_Nm"						& _
					" FROM B_STORAGE_LOCATION"			& _
					" WHERE sl_cd = " & strSlFrCd

Select Case pDataType
		
	Case "AL"
		lgStrSQL	= "	SELECT	T.ITEM_CD,E.ITEM_NM,T.TRACKING_NO,(SUM(T.REQMT_QTY)-SUM(T.ISSUE_QTY)) REQMT_QTY,T.REQMT_UNIT,SUM(T.ISSUE_QTY) ISSUE_QTY," _
					& "			ISNULL((D.GOOD_ON_HAND_QTY),0) SD_QTY,ISNULL((SUM(T.RESRV_QTY)-SUM(T.ISSUE_QTY)-ISNULL(D.GOOD_ON_HAND_QTY,0)),0) RESRV_QTY," _
					& "			ISNULL(C.GOOD_ON_HAND_QTY,0) ONHAND_QTY,E.BASIC_UNIT,E.SPEC" _
					& "	  FROM (SELECT	A.REQMT_QTY,B.ITEM_CD,A.RESRV_QTY,A.ISSUE_QTY,A.SL_CD AS MSL_CD,C.SL_CD AS TSL_CD, " _
					& "					B.TRACKING_NO,B.PLANT_CD,A.REQMT_UNIT,B.SPPL_CD " _
					& "			  FROM	M_CHILD_RESERV_HISTORY A(NOLOCK) INNER JOIN M_CHILD_RESERV B(NOLOCK) " _
					& "				ON	A.PR_NO = B.PR_NO AND A.RESVD_SEQ_NO = B.RESVD_SEQ_NO " _
					& "			  LEFT OUTER JOIN B_STORAGE_LOCATION C(NOLOCK) ON B.SPPL_CD = C.BP_CD " _
					& "			  INNER JOIN M_PUR_REQ D(NOLOCK) ON A.PR_NO = D.PR_NO " _
					& "			  INNER JOIN M_PUR_ORD_DTL E(NOLOCK) ON D.PR_NO = E.PR_NO " _	 	
 					& "			 WHERE	A.REQMT_QTY <> 0 AND B.SPPL_TYPE = " & FilterVar("F", "''", "S") & "  AND	B.PLANT_CD = " & strPlantCd _
 					& "		       AND	C.SL_CD = " & strSlToCd & " AND E.CLS_FLG <> " & FilterVar("Y","''","S") & ") T " _
 					& " LEFT OUTER JOIN	I_ONHAND_STOCK C(NOLOCK) ON (T.PLANT_CD = C.PLANT_CD AND T.MSL_CD = C.SL_CD AND T.ITEM_CD = C.ITEM_CD AND T.TRACKING_NO = C.TRACKING_NO AND T.MSL_CD = " & strSlFrCd & " ) " _
 					& " LEFT OUTER JOIN I_ONHAND_STOCK D(NOLOCK) ON (T.PLANT_CD = D.PLANT_CD AND T.TSL_CD = D.SL_CD AND T.ITEM_CD = D.ITEM_CD AND T.TRACKING_NO = D.TRACKING_NO AND T.TSL_CD = " & strSlToCd & ") " _
 					& " INNER JOIN B_ITEM E(NOLOCK) ON T.ITEM_CD = E.ITEM_CD " _
 					& " INNER JOIN B_ITEM_BY_PLANT F(NOLOCK) ON T.PLANT_CD = F.PLANT_CD AND T.ITEM_CD = F.ITEM_CD AND F.ISSUE_MTHD = " & FilterVar("A", "''", "S") & "  " _
 					& " WHERE 	T.SPPL_CD = " & strCondBpCd _
 					& " GROUP BY T.ITEM_CD,E.ITEM_NM,T.TRACKING_NO,T.REQMT_UNIT,D.good_on_hand_qty,C.good_on_hand_qty,E.basic_unit,E.spec " _
 					& " HAVING ISNULL((SUM(T.RESRV_QTY)-SUM(T.ISSUE_QTY)-ISNULL(D.GOOD_ON_HAND_QTY,0)),0) > 0 " _
 					& " ORDER BY T.item_cd,T.tracking_no "
			
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
		.frm1.txtPlantNm.value		= "<%=strPlantNm%>"
		.frm1.txtSlFrNm.value		= "<%=strSlFrNm%>"
		.frm1.txtSlToNm.value		= "<%=strSlToNm%>"
		.frm1.txtBpCd.value			= "<%=strBpCd%>"
		.frm1.txtBpNm.value			= "<%=strBpNm%>"
		
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


	
