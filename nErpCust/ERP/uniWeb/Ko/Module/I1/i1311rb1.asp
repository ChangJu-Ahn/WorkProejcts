<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% Option Explicit%>
<% session.CodePage=949 %>

<!--'======================================================================================================
'*  1. Module Name          : Inventory
'*  2. Function Name        : ADO Template (Query)
'*  3. Program ID           : i1311rb1
'*  4. Program Name         : 사내재고이동정보 
'*  5. Program Desc         :
'*  6. Modified date(First) : 2001/11/08
'*  7. Modified date(Last)  : 2003/06/03
'*  8. Modifier (First)     : Lee Seung Wook
'*  9. Modifier (Last)      : Lee Seung Wook
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================-->
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->

<%                                                                       
Call LoadBasisGlobalInf()
Call LoadInfTB19029B("Q", "I","NOCOOKIE","MB")  
Call HideStatusWnd 

On Error Resume Next

Dim lgADF                                                              
Dim lgstrRetMsg                                                        
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0                          
Dim lgstrData                                                          
Dim lgStrPrevKey                                                       
Dim lgMaxCount                                                         
Dim strPlantCd	                                                          
Dim strSLCd1	                                                          
Dim strSLCd2                                                           

Dim strPlantNm
Dim strSLNm1
Dim strSLNm2

    lgStrPrevKey   = Request("lgStrPrevKey")                             
    lgMaxCount     = 100                        
    
    Call TrimData()
    Call HeaderData()
    Call FixUNISQLData()
    Call QueryData()
    

Sub TrimData()
     strPlantCd		= Trim(Request("txtPlantCd"))             
     strSlCd1       = Trim(Request("txtSlCd1"))               
     strSlCd2       = Trim(Request("txtSlCd2"))               
End Sub

 Sub HeaderData()
	Dim iStr
    Redim UNISqlId(0)                                             
	Redim UNIValue(0,0)                                           

	UNILock = DISCONNREAD :	UNIFlag = "1"                         

	Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")

	UNISqlId(0) = "160901saa"
	UNIValue(0,0)  = FilterVar(strPlantCd, "''", "S")	
	
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
  	iStr = Split(lgstrRetMsg,gColSep)
    
    If iStr(0) <> "0" Then
    		Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If    
        
    If  rs0.EOF And rs0.BOF Then
    		Call DisplayMsgBox("900014",vbOKOnly, "", "", I_MKSCRIPT)  
    		rs0.Close
    		Set rs0 = Nothing
    		Response.End												
    Else    
    		strPlantNm=rs0(0)
    		rs0.Close
    		Set rs0 = Nothing
    End If
    	
	If strSlCd1 <> "" Then
		UNISqlId(0) = "160903saa"
		UNIValue(0,0)  = FilterVar(strSlCd1, "''", "S")		
		
		lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
		iStr = Split(lgstrRetMsg,gColSep)
	
		If iStr(0) <> "0" Then
				Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
		End If    
	
		If  rs0.EOF And rs0.BOF Then
				Call DisplayMsgBox("900014",vbOKOnly, "", "", I_MKSCRIPT)   
				rs0.Close
				Set rs0 = Nothing
				Response.End											
		Else    
				strSlNm1=rs0(0)
				rs0.Close
				Set rs0 = Nothing
		End If
	End If
	
	If strSlCd2 <> "" Then
		UNISqlId(0) = "160903saa"
		UNIValue(0,0)  = FilterVar(strSlCd2, "''", "S")	
		
		lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
		iStr = Split(lgstrRetMsg,gColSep)
	
		If iStr(0) <> "0" Then
				Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
		End If    
	
		If  rs0.EOF And rs0.BOF Then
				Call DisplayMsgBox("900014",vbOKOnly, "", "", I_MKSCRIPT)   
				rs0.Close
				Set rs0 = Nothing
				Response.End												
		Else    
				strSlNm2=rs0(0)
				rs0.Close
				Set rs0 = Nothing
		End If
	End If
	
End Sub

Sub FixUNISQLData()

	Redim UNISqlId(0)
	Redim UNIValue(0,6)	

'   STATEMENTS TABLE UPDATE - i_statements_upd_I1311ra1_20070410_NYN.sql 
    UNISqlId(0) = "I1311ra1"

'   20070410 modify start
    UNIValue(0,0) = "a.item_cd,c.item_nm,b.tracking_no, sum((ISNULL(a.ss_qty,0) - ISNULL(d.good_on_hand_qty,0)) + b.qty) as qty ,c.basic_unit,1"

    UNIValue(0,1) = FilterVar(strPlantCd, "''", "S")
    UNIValue(0,2) = FilterVar(strSlCd1, "''", "S")
    UNIValue(0,3) = FilterVar(strSlCd2, "''", "S")

	UNIValue(0,4) = "GROUP BY  A.ITEM_CD ,  C.ITEM_NM ,  B.TRACKING_NO, C.BASIC_UNIT "
'   20070410  modify end

	UNIValue(0,5) = "HAVING SUM((ISNULL(a.ss_qty,0) - ISNULL(d.good_on_hand_qty,0)) + b.qty ) > 0"

	UNIValue(0,6) = "ORDER BY  A.ITEM_CD ASC ,  C.ITEM_NM ASC ,  B.TRACKING_NO ASC"

    UNILock = DISCONNREAD :	UNIFlag = "1"                                
End Sub
 
Sub QueryData()
    Dim iStr
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
        
    iStr = Split(lgstrRetMsg,gColSep)
    
    If iStr(0) <> "0" Then    	
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If    
        
    If  rs0.EOF And rs0.BOF Then
        Call DisplayMsgBox("900014",vbOKOnly, "", "", I_MKSCRIPT)  
        rs0.Close
        Set rs0 = Nothing
    Else        
        Call  MakeSpreadSheetData()
    End If
End Sub

Sub MakeSpreadSheetData()
    Dim  iCnt
    Dim  iRCnt
    Dim	PvArr
    
    iCnt = 0
    lgstrData = ""

    If Len(Trim(lgStrPrevKey)) Then                              
       If Isnumeric(lgStrPrevKey) Then
          iCnt = CInt(lgStrPrevKey)
       End If   
    End If   

    For iRCnt = 1 to iCnt  *  lgMaxCount                
        rs0.MoveNext
    Next

    iRCnt = -1
    
    ReDim PvArr(0)
    Do while Not (rs0.EOF Or rs0.BOF)
        iRCnt =  iRCnt + 1
	
	    ReDim Preserve PvArr(iRCnt)

		lgstrData = Chr(11) & rs0(0) & _ 
					Chr(11) & rs0(1) & _  
					Chr(11) & rs0(2) & _  
					Chr(11) & UniConvNumberDBToCompany(rs0(3), ggQty.DecPoint,ggQty.RndPolicy, ggQty.RndUnit, 0) & _
					Chr(11) & rs0(4) & _
					Chr(11) & CStr((iCnt * lgMaxCount) + iRCnt) & Chr(11) & Chr(12)
		
		PvArr(iRCnt) = lgstrData
		
        If  iRCnt >= lgMaxCount Then
            iCnt = iCnt + 1
            lgStrPrevKey = CStr(iCnt)
            Exit Do
        End If
        rs0.MoveNext
	Loop
	
	lgstrData = Join(PvArr, "")

    If  iRCnt < lgMaxCount Then                                        
        lgStrPrevKey = ""                                              
    End If
  	
	rs0.Close
    Set rs0 = Nothing 
    Set lgADF = Nothing                                                
End Sub

    Response.Write "<Script Language=vbscript> " & vbcr
    Response.Write "With parent "                & vbcr 
 
 	Response.Write "	.frm1.txtPlantNm.value  = """ & ConvSPChars(strPlantNm) & """" & vbCr  	   	  
	Response.Write "	.frm1.txtSlNm1.value  = """ & ConvSPChars(strSlNm1) & """" & vbCr  	   	  
	Response.Write "	.frm1.txtSlNm2.value  = """ & ConvSPChars(strSlNm2) & """" & vbCr  	   	  

    Response.Write " .ggoSpread.Source = .frm1.vspdData "             & vbcr
    Response.Write " .ggoSpread.SSShowData """ & ConvSPChars(lgstrData) & """ " & vbcr
    Response.Write " .lgStrPrevKey =    """ & ConvSPChars(lgStrPrevKey)    & """ " & vbcr

   	Response.Write "	If .frm1.vspdData.MaxRows < .parent.VisibleRowCnt(.frm1.vspdData, 0) And .lgStrPrevKey <> """" Then "	& vbCr
  	Response.Write "		.DbQuery								"				& vbCr
  	Response.Write "    Else								"				& vbCr
  	Response.Write "		.DbQueryOK								"				& vbCr
	Response.Write "    End If								"				& vbCr

    Response.Write "End With "       & vbcr
    Response.Write "</Script> "      & vbcr

%>

  
