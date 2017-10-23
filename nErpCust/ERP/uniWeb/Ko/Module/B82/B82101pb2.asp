f<%@ LANGUAGE=VBSCript TRANSACTION=Required%>

<!-- #Include file="../../inc/IncServer.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../B81/B81COMM.ASP" -->


<%
'On Error Resume Next                                                             '☜: Protect system from crashing
Err.Clear                                                                        '☜: Clear Error status

Call HideStatusWnd                                                               '☜: Hide Processing message
'---------------------------------------Common-----------------------------------------------------------
lgErrorStatus     = "NO"
lgErrorPos        = ""                                                           '☜: Set to space
lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
lgKeyStream       = Split(Request("txtKeyStream"),gColSep)

lgLngMaxRow       = Request("txtMaxRows")                                        '☜: Read Operation Mode (CRUD)
lgMaxCount        = CInt(Request("lgMaxCount"))                                  '☜: Fetch count at a time for VspdData
lgStrPrevKeyIndex = UNICInt(Trim(Request("lgStrPrevKeyIndex")),0)                '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
    
'------ Developer Coding part (Start ) ------------------------------------------------------------------

'------ Developer Coding part (End   ) ------------------------------------------------------------------ 

Call SubOpenDB(lgObjConn)                                                        '☜: Make a DB Connection
    
Select Case CStr(lgOpModeCRUD)
   Case CStr(UID_M0001)   
        Call SubBizQuery()
   Case CStr(UID_M0002)                                                         '☜: Save,Update
        'Call SubBizSaveMulti()
End Select
    
Call SubCloseDB(lgObjConn)                                                       '☜: Close DB Connection

Sub SubBizQuery()
    'On Error Resume Next                                                         '☜: Protect system from crashing
    Err.Clear                                                                    '☜: Clear Error status

    Dim iDx
    Dim iLoopMax
    Dim lgStrSQL
    Dim iSelCount
      'Call DisplayMsgBox("971001", vbOKOnly, lgKeyStream(4), "메시지", I_MKSCRIPT)
      'Call DisplayMsgBox("971001", vbOKOnly, lgKeyStream(5), "메시지", I_MKSCRIPT)
     'Call DisplayMsgBox("971001", vbOKOnly, lgKeyStream(6), "메시지", I_MKSCRIPT)
      
    '---------- set code name  --------------------------------------------------------------------
	call GetNameChk("minor_nm","b_minor","major_cd='Y1001' and minor_cd="&filterVar(lgKeyStream(3),"''","S"),	lgKeyStream(3),"txtItemKind","품목구분","N") '품목구분
	call GetNameChk("class_name","b_cis_item_class","item_acct='"&lgKeyStream(2)&"' and item_kind ='"&lgKeyStream(3)&"' and item_lvl='L1' and class_cd="&filterVar(lgKeyStream(4),"''","S"),	lgKeyStream(4),"txtItemLvl1","대분류","N") '대분류
	call GetNameChk("class_name","b_cis_item_class","item_acct='"&lgKeyStream(2)&"' and item_kind ='"&lgKeyStream(3)&"' and item_lvl='L2' and class_cd="&filterVar(lgKeyStream(5),"''","S"),	lgKeyStream(5),"txtItemLvl2","중분류","N") '중분류
	call GetNameChk("class_name","b_cis_item_class","item_acct='"&lgKeyStream(2)&"' and item_kind ='"&lgKeyStream(3)&"' and item_lvl='L3' and class_cd="&filterVar(lgKeyStream(6),"''","S"),	lgKeyStream(6),"txtItemLvl3","소분류","N") '소분류
    '---------- Developer Coding part (Start) ---------------------------------------------------------------
       
    iSelCount = lgMaxCount + lgMaxCount *  lgStrPrevKeyIndex + 1
        
    lgStrSQL = " SELECT TOP " & iSelCount  & " A.ITEM_CD,"
    lgStrSQL = lgStrSQL & "     A.ITEM_NM, "
    lgStrSQL = lgStrSQL & "     A.ITEM_SPEC,"
    lgStrSQL = lgStrSQL & "     A.ITEM_ACCT, "
    lgStrSQL = lgStrSQL & "     ITEM_ACCT_NM = dbo.ufn_GetCodeName('P1001',A.ITEM_ACCT), "    
    lgStrSQL = lgStrSQL & "     A.ITEM_KIND, "
    lgStrSQL = lgStrSQL & "     ITEM_KIND_NM = dbo.ufn_GetCodeName('Y1001',A.ITEM_KIND), "
    lgStrSQL = lgStrSQL & "     A.ITEM_LVL1,"
    lgStrSQL = lgStrSQL & "     dbo.ufn_s_CIS_GetParentNn(A.ITEM_ACCT,A.ITEM_KIND,A.ITEM_LVL1,'*','L1') ITEM_LVL1_NM , "
    lgStrSQL = lgStrSQL & "     A.ITEM_LVL2, "    
    lgStrSQL = lgStrSQL & "     dbo.ufn_s_CIS_GetParentNn(A.ITEM_ACCT,A.ITEM_KIND,A.ITEM_LVL2,A.ITEM_LVL1,'L2') ITEM_LVL2_NM , "
    lgStrSQL = lgStrSQL & "     A.ITEM_LVL3, "
    lgStrSQL = lgStrSQL & "     dbo.ufn_s_CIS_GetParentNn(A.ITEM_ACCT,A.ITEM_KIND,A.ITEM_LVL3,A.ITEM_LVL2,'L3') ITEM_LVL3_NM, "
    lgStrSQL = lgStrSQL & "     A.ITEM_SEQNO, "
    lgStrSQL = lgStrSQL & "     A.ITEM_VER, "
    lgStrSQL = lgStrSQL & "     ITEM_VER_NM  = dbo.ufn_GetCodeName('Y1004',A.ITEM_VER), "    
    lgStrSQL = lgStrSQL & "     A.ITEM_NM2, "
    lgStrSQL = lgStrSQL & "     A.ITEM_SPEC2,"
    lgStrSQL = lgStrSQL & "     A.ITEM_UNIT, "
    lgStrSQL = lgStrSQL & "     A.PUR_TYPE,  "    
    lgStrSQL = lgStrSQL & "     PUR_TYPE_NM  = dbo.ufn_GetCodeName('P1003',A.PUR_TYPE), "
    lgStrSQL = lgStrSQL & "     A.BASIC_CODE, "
    lgStrSQL = lgStrSQL & "     BASIC_CODE_NM = (SELECT S.ITEM_NM FROM B_CIS_ITEM_MASTER S WHERE S.ITEM_CD = A.BASIC_CODE),"
    lgStrSQL = lgStrSQL & "     A.PUR_GROUP, "
    lgStrSQL = lgStrSQL & "     PUR_GROUP_NM  = E.PUR_GRP_NM,"    
    lgStrSQL = lgStrSQL & "     A.PUR_VENDOR, "
    lgStrSQL = lgStrSQL & "     PUR_VENDOR_NM = F.BP_NM, "
    lgStrSQL = lgStrSQL & "     A.UNIFY_PUR_FLAG, "
    lgStrSQL = lgStrSQL & "     A.UNIT_WEIGHT, "
    lgStrSQL = lgStrSQL & "     A.UNIT_OF_WEIGHT, "    
    lgStrSQL = lgStrSQL & "     A.GROSS_WEIGHT,  "
    lgStrSQL = lgStrSQL & "     A.GROSS_UNIT,  "
    lgStrSQL = lgStrSQL & "     A.CBM, "
    lgStrSQL = lgStrSQL & "     A.CBM_DESCRIPTION,"
    lgStrSQL = lgStrSQL & "     A.HS_CODE, " 
    lgStrSQL = lgStrSQL & "     HS_CODE_NM = ( SELECT S.HS_NM FROM B_HS_CODE S WHERE S.HS_CD = A.HS_CODE ), "
    lgStrSQL = lgStrSQL & "     A.VALID_FROM_DT, "
    lgStrSQL = lgStrSQL & "     A.VALID_TO_DT, "
    lgStrSQL = lgStrSQL & "     A.DOC_NO, " 
    lgStrSQL = lgStrSQL & "     A.INTERNAL_CD, "
    lgStrSQL = lgStrSQL & "     B.R_GRADE, "
    lgStrSQL = lgStrSQL & "     B.T_GRADE, "
    lgStrSQL = lgStrSQL & "     B.P_GRADE ,"
    lgStrSQL = lgStrSQL & "     B.Q_GRADE "
    lgStrSQL = lgStrSQL & " FROM B_CIS_ITEM_MASTER  A(NOLOCK), "
    lgStrSQL = lgStrSQL & "      B_CIS_NEW_ITEM_REQ B(NOLOCK), "
    lgStrSQL = lgStrSQL & "      B_PUR_GRP          E(NOLOCK), "
    lgStrSQL = lgStrSQL & "      B_BIZ_PARTNER      F(NOLOCK) "

    lgStrSQL = lgStrSQL & "WHERE A.ITEM_CD     = B.ITEM_CD "
    lgStrSQL = lgStrSQL & "  AND A.INTERNAL_CD = B.INTERNAL_CD "
    
    If Trim(lgKeyStream(8)) = "NEW" Then
       lgStrSQL = lgStrSQL & "  AND B.STATUS     IN ('E','T') "  '신규의뢰일때 기준품목참조는 종료된것중에서도 참조가 가능하다.
    Else
       lgStrSQL = lgStrSQL & "  AND B.STATUS     = 'T'  "        '변경은 이관된 품목중에서 변경요청이 중이지 않는 품목만 가능하다.
       lgStrSQL = lgStrSQL & "  AND NOT EXISTS ( SELECT 1 FROM B_CIS_CHANGE_ITEM_REQ S WHERE S.ITEM_CD = A.ITEM_CD AND S.STATUS IN ('R','A','E')) "
       lgStrSQL = lgStrSQL & "  AND NOT EXISTS ( SELECT 1 FROM B_CIS_CHANGE_ITEM_NM_REQ S WHERE S.ITEM_CD = A.ITEM_CD AND S.STATUS IN ('R','A','E')) "
    End If
      
    lgStrSQL = lgStrSQL & "  AND A.PUR_GROUP *= E.PUR_GRP "
    lgStrSQL = lgStrSQL & "  AND A.PUR_VENDOR*= F.BP_CD "    

  
	If Trim(Request("txtNewChange")) = "NEW" Then
		lgStrSQL = lgStrSQL & " AND A.ITEM_CD LIKE " & FilterVar(lgKeyStream(0) & "%","","S")
	End If
	
	If Trim(lgKeyStream(1)) <> "" Then
		lgStrSQL = lgStrSQL & " AND A.ITEM_NM LIKE " & FilterVar(lgKeyStream(1) & "%","","S")
	End If
	
	If Trim(lgKeyStream(2)) <> "" Then
		lgStrSQL = lgStrSQL & " AND A.ITEM_ACCT = " & FilterVar(lgKeyStream(2),"","S")
	End If
	If Trim(lgKeyStream(3)) <> "" Then
		lgStrSQL = lgStrSQL & " AND A.ITEM_KIND = " & FilterVar(lgKeyStream(3),"","S")
	End If
	If Trim(lgKeyStream(4)) <> "" Then
		lgStrSQL = lgStrSQL & " AND A.ITEM_LVL1 = " & FilterVar(lgKeyStream(4),"","S")
	End If
	If Trim(lgKeyStream(5)) <> "" Then
		lgStrSQL = lgStrSQL & " AND A.ITEM_LVL2 = " & FilterVar(lgKeyStream(5),"","S")
	End If
	If Trim(lgKeyStream(6)) <> "" Then
		lgStrSQL = lgStrSQL & " AND A.ITEM_LVL3 = " & FilterVar(lgKeyStream(6),"","S")
	End If
	
	If Trim(lgKeyStream(7)) <> "" Then
		lgStrSQL = lgStrSQL & " AND A.ITEM_SPEC LIKE " & FilterVar(lgKeyStream(7) & "%","","S")
	End If
		
	lgStrSQL = lgStrSQL & " ORDER BY A.ITEM_CD ASC "
	
	If FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
       lgStrPrevKeyIndex = ""
       Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
       Call SetErrorStatus()
    Else

       Call SubSkipRs(lgObjRs,lgMaxCount * lgStrPrevKeyIndex)

       lgstrData = ""        
       iDx       = 1
       
       Do While Not lgObjRs.EOF

          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ITEM_CD"))
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ITEM_NM"))
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ITEM_SPEC"))
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ITEM_ACCT"))
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ITEM_ACCT_NM"))          
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ITEM_KIND"))
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ITEM_KIND_NM"))
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ITEM_LVL1"))
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ITEM_LVL1_NM"))
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ITEM_LVL2"))          
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ITEM_LVL2_NM"))
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ITEM_LVL3"))
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ITEM_LVL3_NM"))          
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ITEM_SEQNO"))          
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ITEM_VER"))
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ITEM_VER_NM"))          
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ITEM_NM2"))
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ITEM_SPEC2"))
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ITEM_UNIT"))          
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("PUR_TYPE"))          
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("PUR_TYPE_NM"))          
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("BASIC_CODE"))  
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("BASIC_CODE_NM")) 
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("PUR_GROUP"))
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("PUR_GROUP_NM"))
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("PUR_VENDOR"))
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("PUR_VENDOR_NM"))          
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("UNIFY_PUR_FLAG"))          
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("UNIT_WEIGHT"))
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("UNIT_OF_WEIGHT"))
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("GROSS_WEIGHT"))
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("GROSS_UNIT"))          
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("CBM"))          
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("CBM_DESCRIPTION"))
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("HS_CODE"))  
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("HS_CODE_NM"))
          lgstrData = lgstrData & Chr(11) & UniConvDateDbToCompany(lgObjRs("VALID_FROM_DT"),"")
          lgstrData = lgstrData & Chr(11) & UniConvDateDbToCompany(lgObjRs("VALID_TO_DT"),"")          
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("DOC_NO"))
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("INTERNAL_CD"))
          
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("R_GRADE"))
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("T_GRADE"))
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("P_GRADE"))
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("Q_GRADE"))
          
          lgstrData = lgstrData & Chr(11) & lgLngMaxRow + iDx
          lgstrData = lgstrData & Chr(11) & Chr(12)
       
		  lgObjRs.MoveNext

          iDx =  iDx + 1
          
          If iDx > lgMaxCount Then
             lgStrPrevKeyIndex = lgStrPrevKeyIndex + 1
             Exit Do
          End If   
               
       Loop 
       
    End If

    If iDx <= lgMaxCount Then
       lgStrPrevKeyIndex = ""
    End If   

	Call SubHandleError("MR",lgObjConn,lgObjRs,Err)
    Call SubCloseRs(lgObjRs)                                                      '☜: Release RecordSSet

End Sub    


'============================================================================================================
' Name : CommonOnTransactionCommit
' Desc : This Sub is called by OnTransactionCommit Error handler
'============================================================================================================
Sub CommonOnTransactionCommit()
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

'============================================================================================================
' Name : CommonOnTransactionAbort
' Desc : This Sub is called by OnTransactionAbort Error handler
'============================================================================================================
Sub CommonOnTransactionAbort()
    lgErrorStatus    = "YES"
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

'============================================================================================================
' Name : SetErrorStatus
' Desc : This Sub set error status
'============================================================================================================
Sub SetErrorStatus()
    lgErrorStatus     = "YES"                                                         '☜: Set error status
End Sub
'============================================================================================================
' Name : SubHandleError
' Desc : This Sub handle error
'============================================================================================================
Sub SubHandleError(pOpCode,pConn,pRs,pErr)
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

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
       Case "<%=UID_M0001%>"                                                         '☜ : Query
          If Trim("<%=lgErrorStatus%>") = "NO" Then
              With Parent
                .ggoSpread.Source     = .frm1.vspdData
                .lgStrPrevKeyIndex    = "<%=lgStrPrevKeyIndex%>"
                .ggoSpread.SSShowData "<%=lgstrData%>"
                .lgStrPrevKey         = "<%=lgStrPrevKey%>"
                .DBQueryOk        
	         End with
          End If       
    End Select    
       
</Script>	