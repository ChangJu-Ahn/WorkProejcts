<%@  LANGUAGE = VBSCript%>
<% Option Explicit%>
<!--'**********************************************************************************************
'*  1. Module Name          : Inventory List onhand stock detail
'*  2. Function Name        : 
'*  3. Program ID           : I1211pb1.asp
'*  4. Program Name         : 
'*  5. Program Desc         : 수불품목팝업 
'*  6. Comproxy List        : 
'                             +B25018ListPlant
'                             +B25019LookUpPlant
'*  7. Modified date(First) : 2002/04/03
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
On Error Resume Next
Err.Clear

Call LoadBasisGlobalInf()
Call LoadInfTB19029B("Q", "I","NOCOOKIE","MB")
Call HideStatusWnd
         
Dim strPlantCd
Dim strPlantNm
Dim strSlCd
Dim strSlNm
Dim strItemCd
Dim strItemNm
Dim strFlag
Dim IntRetCD
Dim lgStrSQL2
Dim lgStrSQL3

lgLngMaxRow       = Request("txtMaxRows")
lgMaxCount        = 100                  
lgStrPrevKeyIndex = UNICInt(Trim(Request("lgStrPrevKeyIndex")),0)   
lgErrorStatus     = "NO"
lgErrorPos        = ""                   
lgOpModeCRUD      = Request("txtMode")   

Call SubOpenDB(lgObjConn)

Call SubCreateCommandObject(lgObjComm)

strPlantCd = FilterVar(Request("txtPlantCd"), "''", "S")
strPlantNm = FilterVar(Request("txtPlantNm"), "''", "S")
strSlCd    = FilterVar(Request("txtSlCd"), "''", "S")
strSlNm    = FilterVar(Request("txtSlNm"), "''", "S")
strItemCd  = FilterVar(Request("txtItemCd1"), "''", "S")
strItemNm  = FilterVar("%" & Trim(Request("txtItemNm1")) & "%", "''", "S")
strFlag	   = trim(Request("txtFlag"))

If strItemCd <> "''" then 
	Call SubBizQuery("CD")
Elseif strItemNm <> "" & FilterVar("%%", "''", "S") & "" then 
	Call SubBizQuery("NM") 
Else
	Call SubBizQuery("AL") 
End if

 
Call SubCloseCommandObject(lgObjComm)    
Call SubCloseDB(lgObjConn)      


'============================================================================================================
' Name : SubBizQuery
'============================================================================================================
Sub SubBizQuery(pType)
 
 Dim iDx
 Dim PvArr

On Error Resume Next
Err.Clear

 If pType = "CD" Then
	Call SubMakeSQLStatements("CD",strItemCd,strItemNm,strSlCd,strPlantCd,strFlag)
 Elseif pType = "NM" Then
	Call SubMakeSQLStatements("NM",strItemCd,strItemNm,strSlCd,strPlantCd,strFlag)
 Elseif pType = "AL" Then
	Call SubMakeSQLStatements("AL","","",strSlCd,strPlantCd,strFlag)
 Else
 End If
 
  
 '---------------------------
 ' Header Single 조회 
 '---------------------------    
 If  FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL2,"X","X") = False Then
  intCondRet = -1
  lgStrPrevKeyIndex = ""  
  Call DisplayMsgBox("125000", vbInformation, "", "", I_MKSCRIPT)
  Call SetErrorStatus()
  Response.End 
 End If
 strPlantNm = lgObjRs(0)
 %>
 <Script Language="VBScript">
  parent.txtPlantNm.value = "<%=ConvSPChars(strPlantNm)%>"
 </Script> 
 <%
  
 Call SubCloseRs(lgObjRs)
 
 If  FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL3,"X","X") = False Then
  intCondRet = -1
  lgStrPrevKeyIndex = ""  
  Call DisplayMsgBox("169922", vbInformation, "", "", I_MKSCRIPT)
  Call SetErrorStatus()
  Response.End 
 End If
 strSlNm = lgObjRs(0)
 %>
 <Script Language="VBScript">
   parent.txtSlNm.value    = "<%=ConvSPChars(strSlNm)%>"
 </Script> 
 <%

	Call SubCloseRs(lgObjRs)
	 
	If  FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
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
		
		    ReDim Preserve PvArr(iDx-1)
	        
	        lgstrData = Chr(11) & ConvSPChars(lgObjRs(0)) & _
						Chr(11) & ConvSPChars(lgObjRs(1)) & _
						Chr(11) & ConvSPChars(lgObjRs(2)) & _
						Chr(11) & ConvSPChars(lgObjRs(3)) & _
						Chr(11) & ConvSPChars(lgObjRs(4)) & _
						Chr(11) & ConvSPChars(lgObjRs(5)) & _
						Chr(11) & ConvSPChars(lgObjRs(6)) & _
						Chr(11) & UniConvNumberDBToCompany(lgObjRs(7),ggQty.DecPoint,ggQty.RndPolicy, ggQty.RndUnit, 0) & _
						Chr(11) & UniConvNumberDBToCompany(lgObjRs(8),ggQty.DecPoint,ggQty.RndPolicy, ggQty.RndUnit, 0) & _
						Chr(11) & UniConvNumberDBToCompany(lgObjRs(9),ggQty.DecPoint,ggQty.RndPolicy, ggQty.RndUnit, 0) & _
						Chr(11) & UniConvNumberDBToCompany(lgObjRs(10),ggQty.DecPoint,ggQty.RndPolicy, ggQty.RndUnit, 0) & _
						Chr(11) & UniConvNumberDBToCompany(lgObjRs(11),ggQty.DecPoint,ggQty.RndPolicy, ggQty.RndUnit, 0) & _
						Chr(11) & ConvSPChars(lgObjRs(12)) & _
						Chr(11) & ConvSPChars(lgObjRs(13)) & _
						Chr(11) & ConvSPChars(lgObjRs(14)) & _
						Chr(11) & lgLngMaxRow + iDx & Chr(11) & Chr(12)
				
			PvArr(iDx-1) = lgstrData
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
'============================================================================================================
Sub SubMakeSQLStatements(pDataType,pCode,pCode1,pCode2,pCode3,pCode4)

    On Error Resume Next                                                     
    Err.Clear                                                                
    
 Dim iSelCount

 lgStrSQL2 = "SELECT Plant_Nm"
 lgStrSQL2 = lgStrSQL2 & " FROM B_PLANT"
 lgStrSQL2 = lgStrSQL2 & " WHERE plant_cd = " & strPlantCd
 
 lgStrSQL3 = "SELECT Sl_Nm"
 lgStrSQL3 = lgStrSQL3 & " FROM B_STORAGE_LOCATION"
 lgStrSQL3 = lgStrSQL3 & " WHERE sl_cd = " & strSlCd 

   lgStrSQL = "SELECT E.item_cd,E.item_nm,E.spec,E.basic_unit,B.tracking_no,A.lot_no,A.lot_sub_no,A.good_on_hand_qty,A.bad_on_hand_qty," _
			& "A.stk_on_insp_qty,A.stk_on_trns_qty,A.picking_qty, " _ 
			& " [DBO].[ufn_GetStockType](E.ITEM_CD, A.LOT_NO) stock_type ,  " _ 
			& " [DBO].[ufn_GetStockType_BPCD](E.ITEM_CD, A.LOT_NO) BP_CD ,  " _ 
			& " [DBO].[ufn_GetStockType_BPNM](E.ITEM_CD, A.LOT_NO) BP_NM  " _ 
			& " FROM I_ONHAND_STOCK_DETAIL A,I_ONHAND_STOCK B,I_MATERIAL_VALUATION C,B_ITEM_BY_PLANT D,B_ITEM E,B_STORAGE_LOCATION F" _
			& " WHERE A.plant_cd = B.plant_cd " _
			& " AND A.item_cd = B.item_cd " _
			& " AND A.sl_cd = B.sl_cd " _
			& " AND A.tracking_no = B.tracking_no " _
			& " AND B.plant_cd = C.plant_cd " _
			& " AND B.item_cd = C.item_cd " _
			& " AND B.tracking_no = C.tracking_no " _
			& " AND C.plant_cd = D.plant_cd " _
			& " AND C.item_cd = D.item_cd " _
			& " AND D.item_cd = E.item_cd " _
			& " AND B.sl_cd = F.sl_cd " _
			& " AND B.block_indicator = " & FilterVar("N", "''", "S")
   
   If pCode4 = "Y" Then
	lgStrSQL = lgStrSQL & " AND (A.good_on_hand_qty <> " & 0 _
						& " Or A.bad_on_hand_qty <> " & 0 & ")"
   End If

	Select Case pDataType

	  Case "CD"
	   lgStrSQL = lgStrSQL & " AND E.item_cd >= "   & pCode _
				& " AND E.item_nm LIKE " & pCode1 _
				& " AND A.sl_cd = "      & pCode2 _
				& " AND A.plant_cd = "   & pCode3 _
				& " ORDER BY E.item_cd,E.item_nm,B.tracking_no, A.lot_no, A.lot_sub_no " 

	  Case "NM"
	   lgStrSQL = lgStrSQL & " AND E.item_cd >= "   & pCode _
				& " AND E.item_nm LIKE " & pCode1 _
				& " AND A.sl_cd = "      & pCode2 _
				& " AND A.plant_cd = "   & pCode3 _
				& " ORDER BY E.item_nm,E.item_cd,B.tracking_no, A.lot_no, A.lot_sub_no "
	   
	  Case "AL"
	   lgStrSQL = lgStrSQL & " AND A.sl_cd = " & pCode2 _
				& " AND A.plant_cd = " & pCode3 _
				& " ORDER BY E.item_cd,B.tracking_no, A.lot_no, A.lot_sub_no "
	End Select    
End Sub    

'============================================================================================================
' Name : CommonOnTransactionAbort
'============================================================================================================
Sub CommonOnTransactionAbort()
    lgErrorStatus    = "YES"
End Sub

'============================================================================================================
' Name : SetErrorStatus
'============================================================================================================
Sub SetErrorStatus()
    lgErrorStatus     = "YES"     
End Sub

'============================================================================================================
' Name : SubHandleError
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
	.txtPlantNm.value = "<%=ConvSPChars(strPlantNm)%>"
	.txtSlNm.value    = "<%=ConvSPChars(strSlNm)%>"
	If "<%=lgErrorStatus%>" = "NO" And "<%=IntRetCd%>" <> -1 Then
		.ggoSpread.Source = .vspdData
		.lgStrPrevKeyIndex = "<%=lgStrPrevKeyIndex%>"
		.ggoSpread.SSShowData "<%=lgstrData%>"

		if .vspdData.MaxRows < .parent.VisibleRowCnt(.vspdData, 0) and .lgStrPrevKeyIndex <> "" Then
			.DbQuery
		Else
			.DbQueryOk
		End If
	End If   
End With 
</Script> 


 

