<%@ LANGUAGE=VBSCript%>
<%Option Explicit%>
<!--'**********************************************************************************************
'*  1. Module Name          : Inventory
'*  2. Function Name        : 
'*  3. Program ID           : I1523pb1.asp
'*  4. Program Name         : 
'*  5. Program Desc         : VMI 현재고현황팝업 
'*  6. Comproxy List        : 
'                             
'*  7. Modified date(First) : 2003/01/14
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
Call LoadBasisGlobalInf()
Call LoadInfTB19029B("Q", "I","NOCOOKIE","PB")   
Call HideStatusWnd 
											
On Error Resume Next
Err.Clear

Dim strPlantCd
Dim strPlantNm
Dim strSlCd
Dim strSlNm
Dim strBpCd
Dim strBpNm
Dim strItemCd
Dim strItemNm

Dim IntRetCD

Dim lgStrSQL2
Dim lgStrSQL3
Dim lgStrSQL4
'---------------------------------------Common-----------------------------------------------------------

lgLngMaxRow       = Cint(Request("txtMaxRows"))                                   
lgMaxCount        = 100                              
lgStrPrevKeyIndex = UNICInt(Trim(Request("lgStrPrevKeyIndex")),0)   
lgErrorStatus     = "NO"
lgErrorPos        = ""                                                        
'------ Developer Coding part (Start ) ------------------------------------------------------------------


strPlantCd	= FilterVar(Request("txtPlantCd"), "''", "S")
strSlCd		= FilterVar(Request("txtSlCd"), "''", "S")
strBpCd		= FilterVar(Request("txtBpCd"), "''", "S")
strItemCd	= FilterVar(Request("txtItemCd"), "''", "S")
strItemNm	= FilterVar(Request("txtItemNm"), "''", "S")

Call SubOpenDB(lgObjConn)
Call SubCreateCommandObject(lgObjComm)

If strItemCd = "''" and strItemNm <> "''" Then
	Call SubBizQuery("NM")
Else
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
    
lgStrSQL2 = "SELECT Plant_Nm"
lgStrSQL2 = lgStrSQL2 & " FROM B_PLANT"
lgStrSQL2 = lgStrSQL2 & " WHERE plant_cd	= "	& strPlantCd
    
If FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL2,"X","X") = False Then
	intCondRet = -1
	lgStrPrevKeyIndex = ""
	Call DisplayMsgBox("125000",vbInformation, "", "",I_MKSCRIPT)
	Call SetErrorStatus()
	Response.End
End IF
	strPlantNm = ConvSPChars(lgObjRs(0))
Call SubCloseRs(lgObjRs)
	
lgStrSQL3 = "SELECT Sl_Nm"
lgStrSQL3 = lgStrSQL3 & " FROM I_VMI_STORAGE_LOCATION"
lgStrSQL3 = lgStrSQL3 & " WHERE sl_cd =	"	& strSlCd
lgStrSQL3 = lgStrSQL3 & " AND plant_cd = "	& strPlantCd
	
If FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL3,"X","X") = False Then
	IntRetCD = -1
	lgStrPrevKeyIndex = ""
	Call DisplayMsgBox("162001",vbInformation,"","",I_MKSCRIPT)
	Call SetErrorStatus()
	Response.End
End IF
	
	strSlNm = ConvSPChars(lgObjRs(0))
Call SubCloseRs(lgObjRs)

lgStrSQL4 = "SELECT Bp_Nm"
lgStrSQL4 = lgStrSQL4 & " FROM B_BIZ_PARTNER"
lgStrSQL4 = lgStrSQL4 & " WHERE bp_cd = " & strBpCd

If FncOpenRs("R", lgObjConn,lgObjRs,lgStrSQL4,"X","X") = False Then
	IntRetCD = -1
	lgStrPrevKeyIndex = ""
	Call DisplayMsgBox("229927",vbInformation,"","",I_MKSCRIPT)
	Call SetErrorStatus()
	Response.End
End if

	strBpNm = ConvSPChars(lgObjRs(0))
Call SubCloseRs(lgObjRs)
    
    
	If pType = "AL" Then
    '---------- Developer Coding part (Start) ---------------------------------------------------------------
		Call SubMakeSQLStatements("AL",strPlantCd,strSlCd,strBpCd,strItemCd,strItemNm)         
	Else
		Call SubMakeSQLStatements("NM",strPlantCd,strSlCd,strBpCd,strItemCd,strItemNm)
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
		parent.frm1.txtPlantNm.value	= "<%=strPlantNm%>"
		parent.frm1.txtSlNm.value		= "<%=strSlNm%>"
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
        
            lgstrData = Chr(11) & ConvSPChars(lgObjRs(0)) & _
						Chr(11) & ConvSPChars(lgObjRs(1)) & _
						Chr(11) & UniNumClientFormat(lgObjRs(2),ggQty.DecPoint,0) & _
						Chr(11) & ConvSPChars(lgObjRs(3)) & _
						Chr(11) & ConvSPChars(lgObjRs(4)) & _
						Chr(11) & ConvSPChars(lgObjRs(5)) & _
						Chr(11) & ConvSPChars(lgObjRs(6)) & _
						Chr(11) & ConvSPChars(lgObjRs(7)) & _
						Chr(11) & ConvSPChars(lgObjRs(8)) & _
						Chr(11) & ConvSPChars(lgObjRs(9)) & _
						Chr(11) & ConvSPChars(lgObjRs(10)) & _
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
    
	Dim iSelCount
	
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
Select Case pDataType
		
		Case "AL"
			lgStrSQL = "SELECT 	A.item_cd,B.item_nm,A.good_onhand_qty,B.basic_unit,A.tracking_no,A.lot_no,A.lot_sub_no,B.spec,C.lot_flg, C.tracking_flg, C.recv_inspec_flg"
			lgStrSQL = lgStrSQL & " FROM	I_VMI_ONHAND_STOCK A,B_ITEM B, B_ITEM_BY_PLANT C "
			lgStrSQL = lgStrSQL & " WHERE 	A.item_cd	=	B.item_cd "
			lgStrSQL = lgStrSQL & " AND A.plant_cd		= C.plant_cd "
			lgStrSQL = lgStrSQL & " AND A.item_cd		= C.item_cd "
			lgStrSQL = lgStrSQL & " AND A.plant_cd		= "		& strPlantCd
			lgStrSQL = lgStrSQL & " AND A.sl_cd			= "		& strSlCd
			lgStrSQL = lgStrSQL & " AND A.bp_cd			= "		& strBpCd
			lgStrSQL = lgStrSQL & " AND A.item_cd		>= "	& strItemCd
			lgStrSQL = lgStrSQL & " AND B.item_nm		>= "	& strItemNm
			lgStrSQL = lgStrSQL & " ORDER BY A.item_cd,B.item_nm "
		Case "NM"
			lgStrSQL = "SELECT 	A.item_cd,B.item_nm,A.good_onhand_qty,B.basic_unit,A.tracking_no,A.lot_no,A.lot_sub_no,B.spec,C.lot_flg, C.tracking_flg, C.recv_inspec_flg"
			lgStrSQL = lgStrSQL & " FROM	I_VMI_ONHAND_STOCK A,B_ITEM B, B_ITEM_BY_PLANT C "
			lgStrSQL = lgStrSQL & " WHERE 	A.item_cd	=	B.item_cd "
			lgStrSQL = lgStrSQL & " AND A.plant_cd		= C.plant_cd "
			lgStrSQL = lgStrSQL & " AND A.item_cd		= C.item_cd "
			lgStrSQL = lgStrSQL & " AND A.plant_cd		= "		& strPlantCd
			lgStrSQL = lgStrSQL & " AND A.sl_cd			= "		& strSlCd
			lgStrSQL = lgStrSQL & " AND A.bp_cd			= "		& strBpCd
			lgStrSQL = lgStrSQL & " AND A.item_cd		>= "	& strItemCd
			lgStrSQL = lgStrSQL & " AND B.item_nm		>= "	& strItemNm
			lgStrSQL = lgStrSQL & " ORDER BY B.item_nm,A.item_cd "
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
		.frm1.txtPlantNm.value		= "<%=strPlantNm%>"
		.frm1.txtSlNm.value			= "<%=strSlNm%>"
		.frm1.txtBpNm.value			= "<%=strBpNm%>"
		
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


	
