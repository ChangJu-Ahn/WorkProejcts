<%@ LANGUAGE="VBSCRIPT" %>
<%Option Explicit    %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Prucurement
'*  2. Function Name        : 
'*  3. Program ID           : MC101PB1
'*  4. Program Name         : Delivery Item Popup Item by Plant 
'*  5. Program Desc         : Delivery Item Popup Item by Plant 
'*  6. Component List       : 
'*  7. Modified date(First) : 2003/02/22
'*  8. Modified date(Last)  : 2003/05/22
'*  9. Modifier (First)     : 2003/02/22
'* 10. Modifier (Last)      : Kang Su Hwan
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change" 
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************
-->

<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->

<%		
Call LoadBasisGlobalInf
call LoadInfTB19029B("I", "*","NOCOOKIE","MB")  
Call HideStatusWnd 
											
'On Error Resume Next
Err.Clear

'---------------------------------------Common-----------------------------------------------------------

lgLngMaxRow       = Cint(Request("txtMaxRows"))                                        '☜: Read Operation Mode (CRUD)
lgStrPrevKeyIndex = UNICInt(Trim(Request("lgStrPrevKeyIndex")),0)   
lgErrorStatus     = "NO"
lgErrorPos        = ""                                                           '☜: Set to space
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
	Const C_SHEETMAXROWS_D  = 100
	
	Dim iDx
	Dim PvArr
	
'	On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear

    
	If pType = "CD" Then
    '---------- Developer Coding part (Start) ---------------------------------------------------------------
		 Call SubMakeSQLStatements("CD",strItemCd,strItemNm,strItemAcct,strSpec,strPlantCd)           '☜ : Make sql statements
	Elseif pType = "NM" Then
		 Call SubMakeSQLStatements("NM",strItemCd,strItemNm,strItemAcct,strSpec,strPlantCd)           '☜ : Make sql statements
	Else
		 Call SubMakeSQLStatements("AL",strItemCd,strItemNm,strItemAcct,strSpec,strPlantCd)           '☜ : Make sql statements
	End If
		
	If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then                    'If data not exists	
		IntRetCD = -1
		lgStrPrevKeyIndex = ""		
		Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
		Call SetErrorStatus()

		Call SubCloseRs(lgObjRs)

		Response.End 
	Else
		
		IntRetCD = 1
		 Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKeyIndex)
        lgstrData = ""
        iDx       = 1
        
	    ReDim PvArr(C_SHEETMAXROWS_D - 1)
        Do While Not lgObjRs.EOF
        
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(0))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(1))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(2))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(3))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(4))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(5))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(6))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(7))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(8))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(9))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(10))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(11))
            
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
            lgstrData = lgstrData & Chr(11) & lgLngMaxRow + iDx
            lgstrData = lgstrData & Chr(11) & Chr(12)

		    lgObjRs.MoveNext

            iDx =  iDx + 1
            If iDx-1 > C_SHEETMAXROWS_D Then
               lgStrPrevKeyIndex = lgStrPrevKeyIndex + 1
               Exit Do
            End If   
               
            PvArr(iDx-1) = lgstrData	
			lgstrData = ""
        Loop 
        lgstrData = join(PvArr,"")
        
    End If
    If iDx <= C_SHEETMAXROWS_D Then
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

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    
	Dim iSelCount
	
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	Select Case pDataType
			
			Case "AL"
				'lgStrSQL = "SELECT B.item_cd,A.item_nm,A.spec,A.basic_unit,B.item_acct,A.item_group_cd,B.procur_type,B.lot_flg,B.major_sl_cd,B.issued_sl_cd,B.valid_flg,B.recv_inspec_flg"
				lgStrSQL = "SELECT B.item_cd,A.item_nm,A.spec,A.basic_unit,dbo.ufn_GetCodeName(" & FilterVar("P1001", "''", "S") & ",B.item_acct) minor_nm_item_acct,"
				lgStrSQL = lgStrSQL & " A.item_group_cd,dbo.ufn_GetCodeName(" & FilterVar("P1003", "''", "S") & ",B.procur_type) minor_nm_item_acct,B.lot_flg,"
				lgStrSQL = lgStrSQL & " B.major_sl_cd,B.issued_sl_cd,B.valid_flg,B.recv_inspec_flg"
				lgStrSQL = lgStrSQL & " FROM B_ITEM A, B_ITEM_BY_PLANT B"
				lgStrSQL = lgStrSQL & " WHERE A.item_cd		=	B.item_cd "
				lgStrSQL = lgStrSQL & " AND B.material_type	=		" & "" & FilterVar("20", "''", "S") & ""
				lgStrSQL = lgStrSQL & " AND B.plant_cd		=		" & strPlantCd
				lgStrSQL = lgStrSQL & " AND B.item_cd		>=		" & strItemCd
				lgStrSQL = lgStrSQL & " AND A.item_nm		>=		" & strItemNm
				lgStrSQL = lgStrSQL & " AND B.item_acct		like	" & strItemAcct
				lgStrSQL = lgStrSQL & " AND A.spec			>=		" & strSpec
				lgStrSQL = lgStrSQL & " ORDER BY B.item_cd,A.item_nm "
			Case "CD"
				'lgStrSQL = "SELECT B.item_cd,A.item_nm,A.spec,A.basic_unit,B.item_acct,A.item_group_cd,B.procur_type,B.lot_flg,B.major_sl_cd,B.issued_sl_cd,B.valid_flg,B.recv_inspec_flg"
				lgStrSQL = "SELECT B.item_cd,A.item_nm,A.spec,A.basic_unit,dbo.ufn_GetCodeName(" & FilterVar("P1001", "''", "S") & ",B.item_acct) minor_nm_item_acct,"
				lgStrSQL = lgStrSQL & " A.item_group_cd,dbo.ufn_GetCodeName(" & FilterVar("P1003", "''", "S") & ",B.procur_type) minor_nm_item_acct,B.lot_flg,"
				lgStrSQL = lgStrSQL & " B.major_sl_cd,B.issued_sl_cd,B.valid_flg,B.recv_inspec_flg"
				lgStrSQL = lgStrSQL & " FROM B_ITEM A, B_ITEM_BY_PLANT B"
				lgStrSQL = lgStrSQL & " WHERE A.item_cd		=		B.item_cd "
				lgStrSQL = lgStrSQL & " AND B.material_type	=		" & "" & FilterVar("20", "''", "S") & ""
				lgStrSQL = lgStrSQL & " AND B.plant_cd		=		" & strPlantCd
				lgStrSQL = lgStrSQL & " AND B.item_cd		>=		" & strItemCd
				lgStrSQL = lgStrSQL & " AND A.item_nm		>=		" & strItemNm
				lgStrSQL = lgStrSQL & " AND B.item_acct		like	" & strItemAcct
				lgStrSQL = lgStrSQL & " AND A.spec			>=		" & strSpec
				lgStrSQL = lgStrSQL & " ORDER BY B.item_cd,A.item_nm "	
			Case "NM"
				'lgStrSQL = "SELECT B.item_cd,A.item_nm,A.spec,A.basic_unit,B.item_acct,A.item_group_cd,B.procur_type,B.lot_flg,B.major_sl_cd,B.issued_sl_cd,B.valid_flg,B.recv_inspec_flg"
				lgStrSQL = "SELECT B.item_cd,A.item_nm,A.spec,A.basic_unit,dbo.ufn_GetCodeName(" & FilterVar("P1001", "''", "S") & ",B.item_acct) minor_nm_item_acct,"
				lgStrSQL = lgStrSQL & " A.item_group_cd,dbo.ufn_GetCodeName(" & FilterVar("P1003", "''", "S") & ",B.procur_type) minor_nm_item_acct,B.lot_flg,"
				lgStrSQL = lgStrSQL & " B.major_sl_cd,B.issued_sl_cd,B.valid_flg,B.recv_inspec_flg"
				lgStrSQL = lgStrSQL & " FROM B_ITEM A, B_ITEM_BY_PLANT B"
				lgStrSQL = lgStrSQL & " WHERE A.item_cd		=		B.item_cd "
				lgStrSQL = lgStrSQL & " AND B.material_type	=		" & "" & FilterVar("20", "''", "S") & ""
				lgStrSQL = lgStrSQL & " AND B.plant_cd		=		" & strPlantCd
				lgStrSQL = lgStrSQL & " AND B.item_cd		>=		" & strItemCd
				lgStrSQL = lgStrSQL & " AND A.item_nm		>=		" & strItemNm
				lgStrSQL = lgStrSQL & " AND B.item_acct		like	" & strItemAcct
				lgStrSQL = lgStrSQL & " AND A.spec			>=		" & strSpec
				lgStrSQL = lgStrSQL & " ORDER BY A.item_nm,B.item_cd "
	End Select    
   
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
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
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
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

        	If .frm1.vspdData.MaxRows < .VisibleRowCnt(.frm1.vspdData,0)  And .lgStrPrevKeyIndex <> "" Then	 ' GroupView 사이즈로 화면 Row수보다 쿼리가 작으면 다시 쿼리함 
				.DbQuery
				
			Else
				.DbQueryOk
				
			End If
			.frm1.vspdData.focus
		End If   

	End With	
       
</Script>	


	
