<%@ LANGUAGE=VBSCript%>
<%Option Explicit%>
<!--'**********************************************************************************************
'*  1. Module Name          : Inventory
'*  2. Function Name        : 
'*  3. Program ID           : I1522pb1.asp
'*  4. Program Name         : 
'*  5. Program Desc         : VMI 공장별 품목 팝업 
'*  6. Comproxy List        : 
'                             
'*  7. Modified date(First) : 2003/01/10
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Lee Seung Wook
'* 10. Modifier (Last)      : Lee Seung Wook
'* 11. Comment              :
'**********************************************************************************************-->
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
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
'------ Developer Coding part (Start ) ------------------------------------------------------------------
Dim IntRetCD
        
Dim strPlantCd
Dim strItemCd
Dim strItemNm
Dim strItemAcct
Dim strSpec

strPlantCd	= FilterVar(Request("PlantCd"), "''", "S")
strItemCd   = FilterVar(Request("txtItemCd"), "''", "S")
strItemNm   = FilterVar(Request("txtItemNm"), "''", "S")
strItemAcct	= FilterVar("%" & Trim(Request("cboItemAccount")) & "%", "''", "S")
strSpec		= FilterVar(Trim(Request("txtSpec"))	, "''", "S")
'------ Developer Coding part (End   ) ------------------------------------------------------------------ 

Call SubOpenDB(lgObjConn)
Call SubCreateCommandObject(lgObjComm)

	If strItemCd <> "''" then 
		Call SubBizQuery("CD")
	Elseif strItemCd = "''" and strItemNm <> "''" then 
		Call SubBizQuery("NM") 
	Elseif strItemCd = "''" and strItemNm = "''" Then
		Call SubBizQuery("AL") 
	End if

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
    
	If pType = "CD" Then
    '---------- Developer Coding part (Start) ---------------------------------------------------------------
		 Call SubMakeSQLStatements("CD",strItemCd,strItemNm,strItemAcct,strSpec,strPlantCd)        
	Elseif pType = "NM" Then
		 Call SubMakeSQLStatements("NM",strItemCd,strItemNm,strItemAcct,strSpec,strPlantCd)         
	Else
		 Call SubMakeSQLStatements("AL",strItemCd,strItemNm,strItemAcct,strSpec,strPlantCd)         
	End If
		
	
				
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
        iDx       = 1
        ReDim PvArr(0)
        
        Do While Not lgObjRs.EOF
        
            lgstrData = Chr(11) & ConvSPChars(lgObjRs(0)) & _
			 			Chr(11) & ConvSPChars(lgObjRs(1)) & _
						Chr(11) & ConvSPChars(lgObjRs(2)) & _
						Chr(11) & ConvSPChars(lgObjRs(3)) & _
						Chr(11) & ConvSPChars(lgObjRs(4)) & _
						Chr(11) & ConvSPChars(lgObjRs(5)) & _
						Chr(11) & ConvSPChars(lgObjRs(6)) & _
						Chr(11) & ConvSPChars(lgObjRs(7)) & _
						Chr(11) & ConvSPChars(lgObjRs(8)) & _
						Chr(11) & ConvSPChars(lgObjRs(9)) & _
						Chr(11) & ConvSPChars(lgObjRs(10)) & _
						Chr(11) & ConvSPChars(lgObjRs(11)) & _
						Chr(11) & ConvSPChars(lgObjRs(12)) & _
						Chr(11) & ConvSPChars(lgObjRs(13)) & _
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
Sub SubMakeSQLStatements(pDataType,pCode,pCode1,pCode2,pCode3,pCode4)

    On Error Resume Next                                                          
    Err.Clear                                                                     
    
	Select Case pDataType
		Case "AL"
			lgStrSQL = "SELECT B.item_cd,A.item_nm,A.spec,A.basic_unit,B.item_acct,A.item_group_cd,B.procur_type,B.lot_flg,B.major_sl_cd,B.issued_sl_cd,B.valid_flg,B.recv_inspec_flg, B.tracking_flg,C.lot_gen_mthd"
			lgStrSQL = lgStrSQL & " FROM B_ITEM A, B_ITEM_BY_PLANT B,B_LOT_CONTROL C"
			lgStrSQL = lgStrSQL & " WHERE A.item_cd		=	B.item_cd "
			lgStrSQL = lgStrSQL & " AND B.plant_cd		*=	C.plant_cd "
			lgStrSQL = lgStrSQL & " AND B.item_cd		*=	C.item_cd "
			lgStrSQL = lgStrSQL & " AND B.material_type	=		" & "" & FilterVar("30", "''", "S") & ""
			lgStrSQL = lgStrSQL & " AND B.plant_cd		=		" & strPlantCd
			lgStrSQL = lgStrSQL & " AND B.item_cd		>=		" & strItemCd
			lgStrSQL = lgStrSQL & " AND A.item_nm		>=		" & strItemNm
			lgStrSQL = lgStrSQL & " AND B.item_acct		like	" & strItemAcct
			lgStrSQL = lgStrSQL & " AND A.spec			>=		" & strSpec
			lgStrSQL = lgStrSQL & " ORDER BY B.item_cd,A.item_nm "
		Case "CD"
			lgStrSQL = "SELECT B.item_cd,A.item_nm,A.spec,A.basic_unit,B.item_acct,A.item_group_cd,B.procur_type,B.lot_flg,B.major_sl_cd,B.issued_sl_cd,B.valid_flg,B.recv_inspec_flg, B.tracking_flg,C.lot_gen_mthd"
			lgStrSQL = lgStrSQL & " FROM B_ITEM A, B_ITEM_BY_PLANT B,B_LOT_CONTROL C"
			lgStrSQL = lgStrSQL & " WHERE A.item_cd		=	B.item_cd "
			lgStrSQL = lgStrSQL & " AND		B.plant_cd	*=	C.plant_cd "
			lgStrSQL = lgStrSQL & " AND B.item_cd		*=	C.item_cd "
			lgStrSQL = lgStrSQL & " AND B.material_type	=		" & "" & FilterVar("30", "''", "S") & ""
			lgStrSQL = lgStrSQL & " AND B.plant_cd		=		" & strPlantCd
			lgStrSQL = lgStrSQL & " AND B.item_cd		>=		" & strItemCd
			lgStrSQL = lgStrSQL & " AND A.item_nm		>=		" & strItemNm
			lgStrSQL = lgStrSQL & " AND B.item_acct		like	" & strItemAcct
			lgStrSQL = lgStrSQL & " AND A.spec			>=		" & strSpec
			lgStrSQL = lgStrSQL & " ORDER BY B.item_cd,A.item_nm "
		Case "NM"
			lgStrSQL = "SELECT B.item_cd,A.item_nm,A.spec,A.basic_unit,B.item_acct,A.item_group_cd,B.procur_type,B.lot_flg,B.major_sl_cd,B.issued_sl_cd,B.valid_flg,B.recv_inspec_flg, B.tracking_flg,C.lot_gen_mthd"
			lgStrSQL = lgStrSQL & " FROM B_ITEM A, B_ITEM_BY_PLANT B,B_LOT_CONTROL C"
			lgStrSQL = lgStrSQL & " WHERE A.item_cd		=	B.item_cd "
			lgStrSQL = lgStrSQL & " AND B.plant_cd		*=	C.plant_cd "
			lgStrSQL = lgStrSQL & " AND B.item_cd		*=	C.item_cd "
			lgStrSQL = lgStrSQL & " AND B.material_type	=		" & "" & FilterVar("30", "''", "S") & ""
			lgStrSQL = lgStrSQL & " AND B.plant_cd		=		" & strPlantCd
			lgStrSQL = lgStrSQL & " AND B.item_cd		>=		" & strItemCd
			lgStrSQL = lgStrSQL & " AND A.item_nm		>=		" & strItemNm
			lgStrSQL = lgStrSQL & " AND B.item_acct		like	" & strItemAcct
			lgStrSQL = lgStrSQL & " AND A.spec			>=		" & strSpec
			lgStrSQL = lgStrSQL & " ORDER BY A.item_nm,B.item_cd "
	End Select     
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


	
