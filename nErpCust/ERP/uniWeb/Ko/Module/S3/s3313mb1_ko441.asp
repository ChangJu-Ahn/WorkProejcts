<%@ LANGUAGE="VBSCRIPT"  TRANSACTION=Required%>
<%Option Explicit    %>
<!--
'**********************************************************************************************
'*  1. Module Name          : 영업
'*  2. Function Name        : 
'*  3. Program ID           : S4514MB1_KO441
'*  4. Program Name         : 일일수주실적조회
'*  5. Program Desc         : 
'*  6. Component List       : 
'*  7. Modified date(First) : 2007/12/26
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : 
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change" 
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************
-->

<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgSvrVariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->

<Script Language=vbscript>
	Dim strVar1
	Dim strVar2
	Dim TempstrPlantCd

	TempstrPlantCd = "<%=Request("txtPlantCd1")%>"
	'공장명 불러오기 
	Call parent.CommonQueryRs("PLANT_CD,PLANT_NM","B_PLANT","PLANT_CD =  " & parent.FilterVar(TempstrPlantCd , "''", "S") & "",strVar1,strVar2,strVar3,strVar4,strVar5,strVar6,strVar7)
	strVar1 = Replace(strVar1,Chr(11), "")
	strVar2 = Replace(strVar2,Chr(11), "")
	Parent.frm1.txtPlantNm1.Value = strVar2
</Script>

<%
	
    Const C_SHEETMAXROWS_D = 100
    
    Call HideStatusWnd    
    Call LoadBasisGlobalInf()
    Call LoadInfTB19029B("I", "*","NOCOOKIE", "MB")
    
    lgErrorStatus            = "NO"
    lgErrorPos               = ""                                                           '☜: Set to space
    lgOpModeCRUD             = Request("txtMode") 
    
    Dim PlantCd
    Dim documentDt
	

    lgLngMaxRow              = Request("txtMaxRows")     
    PlantCd 				 = Trim(UCase(Request("txtPlantCd1")))
    documentDt               = Trim(Ucase(Request("txtDocumentDt")))        
    
		
	
	
    Call SubOpenDB(lgObjConn)                                                        '☜: Make a DB Connection

    Call SubBizQueryMulti()
    
    Call SubCloseDB(lgObjConn)  
	
'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
    Dim iDx
    Dim iLoopMax
    Dim iKey1, baseDt

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
	
    Call SubMakeSQLStatements("")													 '☆ : Make sql statements

    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        lgStrPrevKeyIndex = ""
        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
        Call SetErrorStatus()
    Else

        Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKeyIndex)

        lgstrData = ""
        
        iDx       = 1
        
        Do While Not lgObjRs.EOF
            'lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("SHIP_TO_PARTY"))        
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("BP_NM"))      
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("GROUP_NM"))      
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("SO_QTY"))                
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("SO_QTY_01"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("SO_QTY_02"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("SO_QTY_03"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("SO_QTY_04"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("SO_QTY_05"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("SO_QTY_06"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("SO_QTY_07"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("SO_QTY_08"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("SO_QTY_09"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("SO_QTY_10"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("SO_QTY_11"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("SO_QTY_12"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("SO_QTY_13"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("SO_QTY_14"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("SO_QTY_15"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("SO_QTY_16"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("SO_QTY_17"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("SO_QTY_18"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("SO_QTY_19"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("SO_QTY_20"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("SO_QTY_21"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("SO_QTY_22"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("SO_QTY_23"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("SO_QTY_24"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("SO_QTY_25"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("SO_QTY_26"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("SO_QTY_27"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("SO_QTY_28"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("SO_QTY_29"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("SO_QTY_30"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("SO_QTY_31"))


            lgstrData = lgstrData & Chr(11) & lgLngMaxRow + iDx
            lgstrData = lgstrData & Chr(11) & Chr(12)
		    lgObjRs.MoveNext

            iDx =  iDx + 1
            If iDx > C_SHEETMAXROWS_D Then
               lgStrPrevKeyIndex = lgStrPrevKeyIndex + 1
            
   Exit Do
            End If   
               
        Loop 
    End If
    
    If iDx <= C_SHEETMAXROWS_D Then
       lgStrPrevKeyIndex = ""
    End If   

	 Call SubHandleError("MR",lgObjConn,lgObjRs,Err)
     Call SubCloseRs(lgObjRs)                                                          '☜: Release RecordSSet

End Sub    
 

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pComp)
   
	Dim iSelCount 
    Dim lgGroupIndex
	
	iSelCount = C_SHEETMAXROWS_D + C_SHEETMAXROWS_D * lgStrPrevKeyIndex + 1   
   
		
	SELECT CASE  Trim(Ucase(Request("cboApType")))  
		   CASE 0
				lgGroupIndex = 0 
		   CASE 1
				lgGroupIndex = 1 
		   CASE 2
				lgGroupIndex = 2 		  			
		   CASE 3
				lgGroupIndex = 3 		  			
		   CASE 4
				lgGroupIndex = 4 		  			
		   CASE 5
				lgGroupIndex = 5 			  			
	END SELECT


	lgStrSQL = "SELECT   TOT.PAYER, TOT.BP_NM, TOT.GROUP_NM, "	       
	lgStrSQL = lgStrSQL &  "          SUM(ISNULL(TOT.SO_QTY,0)) AS SO_QTY, " 
	lgStrSQL = lgStrSQL &  "          SUM(ISNULL(TOT.SO_QTY_01,0)) AS SO_QTY_01, "
	lgStrSQL = lgStrSQL &  "          SUM(ISNULL(TOT.SO_QTY_02,0)) AS SO_QTY_02, "
	lgStrSQL = lgStrSQL &  "          SUM(ISNULL(TOT.SO_QTY_03,0)) AS SO_QTY_03,"
	lgStrSQL = lgStrSQL &  "          SUM(ISNULL(TOT.SO_QTY_04,0)) AS SO_QTY_04,"
	lgStrSQL = lgStrSQL &  "          SUM(ISNULL(TOT.SO_QTY_05,0)) AS SO_QTY_05,"
	lgStrSQL = lgStrSQL &  "          SUM(ISNULL(TOT.SO_QTY_06,0)) AS SO_QTY_06,"
	lgStrSQL = lgStrSQL &  "          SUM(ISNULL(TOT.SO_QTY_07,0)) AS SO_QTY_07,"
	lgStrSQL = lgStrSQL &  "          SUM(ISNULL(TOT.SO_QTY_08,0)) AS SO_QTY_08,"
	lgStrSQL = lgStrSQL &  "          SUM(ISNULL(TOT.SO_QTY_09,0)) AS SO_QTY_09,"
	lgStrSQL = lgStrSQL &  "          SUM(ISNULL(TOT.SO_QTY_10,0)) AS SO_QTY_10,"
	lgStrSQL = lgStrSQL &  "          SUM(ISNULL(TOT.SO_QTY_11,0)) AS SO_QTY_11,"
	lgStrSQL = lgStrSQL &  "          SUM(ISNULL(TOT.SO_QTY_12,0)) AS SO_QTY_12,"
	lgStrSQL = lgStrSQL &  "          SUM(ISNULL(TOT.SO_QTY_13,0)) AS SO_QTY_13,"
	lgStrSQL = lgStrSQL &  "          SUM(ISNULL(TOT.SO_QTY_14,0)) AS SO_QTY_14,"
	lgStrSQL = lgStrSQL &  "          SUM(ISNULL(TOT.SO_QTY_15,0)) AS SO_QTY_15,"
	lgStrSQL = lgStrSQL &  "          SUM(ISNULL(TOT.SO_QTY_16,0)) AS SO_QTY_16,"
	lgStrSQL = lgStrSQL &  "          SUM(ISNULL(TOT.SO_QTY_17,0)) AS SO_QTY_17,"
	lgStrSQL = lgStrSQL &  "          SUM(ISNULL(TOT.SO_QTY_18,0)) AS SO_QTY_18,"
	lgStrSQL = lgStrSQL &  "          SUM(ISNULL(TOT.SO_QTY_19,0)) AS SO_QTY_19,"
	lgStrSQL = lgStrSQL &  "          SUM(ISNULL(TOT.SO_QTY_20,0)) AS SO_QTY_20,"
	lgStrSQL = lgStrSQL &  "          SUM(ISNULL(TOT.SO_QTY_21,0)) AS SO_QTY_21,"
	lgStrSQL = lgStrSQL &  "          SUM(ISNULL(TOT.SO_QTY_22,0)) AS SO_QTY_22,"
	lgStrSQL = lgStrSQL &  "          SUM(ISNULL(TOT.SO_QTY_23,0)) AS SO_QTY_23,"
	lgStrSQL = lgStrSQL &  "          SUM(ISNULL(TOT.SO_QTY_24,0)) AS SO_QTY_24,"
	lgStrSQL = lgStrSQL &  "          SUM(ISNULL(TOT.SO_QTY_25,0)) AS SO_QTY_25,"
	lgStrSQL = lgStrSQL &  "          SUM(ISNULL(TOT.SO_QTY_26,0)) AS SO_QTY_26,"
	lgStrSQL = lgStrSQL &  "          SUM(ISNULL(TOT.SO_QTY_27,0)) AS SO_QTY_27,"
	lgStrSQL = lgStrSQL &  "          SUM(ISNULL(TOT.SO_QTY_28,0)) AS SO_QTY_28,"
	lgStrSQL = lgStrSQL &  "          SUM(ISNULL(TOT.SO_QTY_29,0)) AS SO_QTY_29,"
	lgStrSQL = lgStrSQL &  "          SUM(ISNULL(TOT.SO_QTY_30,0)) AS SO_QTY_30,"
	lgStrSQL = lgStrSQL &  "          SUM(ISNULL(TOT.SO_QTY_31,0)) AS SO_QTY_31"
	lgStrSQL = lgStrSQL &  " FROM     (SELECT   A.PAYER,  C.BP_NM,   dbo.ufn_GetItemGroupNM(B.ITEM_CD, " & lgGroupIndex & ") AS GROUP_NM ,  "
	lgStrSQL = lgStrSQL &  "                   	SUM(ISNULL(CASE WHEN A.RET_ITEM_FLAG = 'N' THEN B.SO_QTY ELSE (-1)*B.SO_QTY END,0)) AS SO_QTY, "
	lgStrSQL = lgStrSQL &  "                   	SUM(ISNULL(CASE WHEN RIGHT(CONVERT(CHAR(10),A.SO_DT,120),2)='01' THEN "
	lgStrSQL = lgStrSQL &  " 							CASE WHEN A.RET_ITEM_FLAG = 'N'  THEN B.SO_QTY ELSE (-1)*B.SO_QTY END "		
	lgStrSQL = lgStrSQL &  " 				  ELSE 0  END,0 )) AS SO_QTY_01, "
	lgStrSQL = lgStrSQL &  "                             SUM(ISNULL(CASE WHEN RIGHT(CONVERT(CHAR(10),A.SO_DT,120),2)='02' THEN "
	lgStrSQL = lgStrSQL &  " 							CASE WHEN A.RET_ITEM_FLAG = 'N'  THEN B.SO_QTY ELSE (-1)*B.SO_QTY END "	
	lgStrSQL = lgStrSQL &  " 				  ELSE 0  END,0 )) AS SO_QTY_02,"
	lgStrSQL = lgStrSQL &  " 		SUM(ISNULL(CASE WHEN RIGHT(CONVERT(CHAR(10),A.SO_DT,120),2)='03' THEN "
	lgStrSQL = lgStrSQL &  " 							CASE WHEN A.RET_ITEM_FLAG = 'N'  THEN B.SO_QTY ELSE (-1)*B.SO_QTY END	"
	lgStrSQL = lgStrSQL &  " 				  ELSE 0  END,0 )) AS SO_QTY_03,"
	lgStrSQL = lgStrSQL &  " 		SUM(ISNULL(CASE WHEN RIGHT(CONVERT(CHAR(10),A.SO_DT,120),2)='04' THEN "
	lgStrSQL = lgStrSQL &  " 							CASE WHEN A.RET_ITEM_FLAG = 'N'  THEN B.SO_QTY ELSE (-1)*B.SO_QTY END	"
	lgStrSQL = lgStrSQL &  " 				  ELSE 0  END,0 )) AS SO_QTY_04,"
	lgStrSQL = lgStrSQL &  " 		SUM(ISNULL(CASE WHEN RIGHT(CONVERT(CHAR(10),A.SO_DT,120),2)='05' THEN "
	lgStrSQL = lgStrSQL &  " 							CASE WHEN A.RET_ITEM_FLAG = 'N'  THEN B.SO_QTY ELSE (-1)*B.SO_QTY END	"
	lgStrSQL = lgStrSQL &  " 				  ELSE 0  END,0 )) AS SO_QTY_05,"
	lgStrSQL = lgStrSQL &  " 		SUM(ISNULL(CASE WHEN RIGHT(CONVERT(CHAR(10),A.SO_DT,120),2)='06' THEN "
	lgStrSQL = lgStrSQL &  " 							CASE WHEN A.RET_ITEM_FLAG = 'N'  THEN B.SO_QTY ELSE (-1)*B.SO_QTY END "	
	lgStrSQL = lgStrSQL &  " 				  ELSE 0  END,0 )) AS SO_QTY_06,"
	lgStrSQL = lgStrSQL &  " 		SUM(ISNULL(CASE WHEN RIGHT(CONVERT(CHAR(10),A.SO_DT,120),2)='07' THEN "
	lgStrSQL = lgStrSQL &  " 							CASE WHEN A.RET_ITEM_FLAG = 'N'  THEN B.SO_QTY ELSE (-1)*B.SO_QTY END	"
	lgStrSQL = lgStrSQL &  " 				  ELSE 0  END,0 )) AS SO_QTY_07,"
	lgStrSQL = lgStrSQL &  " 		SUM(ISNULL(CASE WHEN RIGHT(CONVERT(CHAR(10),A.SO_DT,120),2)='08' THEN "
	lgStrSQL = lgStrSQL &  " 							CASE WHEN A.RET_ITEM_FLAG = 'N'  THEN B.SO_QTY ELSE (-1)*B.SO_QTY END "
	lgStrSQL = lgStrSQL &  " 				  ELSE 0  END,0 )) AS SO_QTY_08,"
	lgStrSQL = lgStrSQL &  " 		SUM(ISNULL(CASE WHEN RIGHT(CONVERT(CHAR(10),A.SO_DT,120),2)='09' THEN "
	lgStrSQL = lgStrSQL &  " 							CASE WHEN A.RET_ITEM_FLAG = 'N'  THEN B.SO_QTY ELSE (-1)*B.SO_QTY END"
	lgStrSQL = lgStrSQL &  " 				  ELSE 0  END,0 )) AS SO_QTY_09,"
	lgStrSQL = lgStrSQL &  " 		SUM(ISNULL(CASE WHEN RIGHT(CONVERT(CHAR(10),A.SO_DT,120),2)='10' THEN "
	lgStrSQL = lgStrSQL &  " 							CASE WHEN A.RET_ITEM_FLAG = 'N'  THEN B.SO_QTY ELSE (-1)*B.SO_QTY END "
	lgStrSQL = lgStrSQL &  " 				  ELSE 0  END,0 )) AS SO_QTY_10, "
	lgStrSQL = lgStrSQL &  " 		SUM(ISNULL(CASE WHEN RIGHT(CONVERT(CHAR(10),A.SO_DT,120),2)='11' THEN "
	lgStrSQL = lgStrSQL &  " 							CASE WHEN A.RET_ITEM_FLAG = 'N'  THEN B.SO_QTY ELSE (-1)*B.SO_QTY END"
	lgStrSQL = lgStrSQL &  " 				  ELSE 0  END,0 )) AS SO_QTY_11,"
	lgStrSQL = lgStrSQL &  " 		SUM(ISNULL(CASE WHEN RIGHT(CONVERT(CHAR(10),A.SO_DT,120),2)='12' THEN "
	lgStrSQL = lgStrSQL &  " 							CASE WHEN A.RET_ITEM_FLAG = 'N'  THEN B.SO_QTY ELSE (-1)*B.SO_QTY END"
	lgStrSQL = lgStrSQL &  " 				  ELSE 0  END,0 )) AS SO_QTY_12,"
	lgStrSQL = lgStrSQL &  " 		SUM(ISNULL(CASE WHEN RIGHT(CONVERT(CHAR(10),A.SO_DT,120),2)='13' THEN "
	lgStrSQL = lgStrSQL &  " 							CASE WHEN A.RET_ITEM_FLAG = 'N'  THEN B.SO_QTY ELSE (-1)*B.SO_QTY END"
	lgStrSQL = lgStrSQL &  " 				  ELSE 0  END,0 )) AS SO_QTY_13,"
	lgStrSQL = lgStrSQL &  " 		SUM(ISNULL(CASE WHEN RIGHT(CONVERT(CHAR(10),A.SO_DT,120),2)='14' THEN "
	lgStrSQL = lgStrSQL &  " 							CASE WHEN A.RET_ITEM_FLAG = 'N'  THEN B.SO_QTY ELSE (-1)*B.SO_QTY END"
	lgStrSQL = lgStrSQL &  " 				  ELSE 0  END,0 )) AS SO_QTY_14,"
	lgStrSQL = lgStrSQL &  " 		SUM(ISNULL(CASE WHEN RIGHT(CONVERT(CHAR(10),A.SO_DT,120),2)='15' THEN "
	lgStrSQL = lgStrSQL &  " 							CASE WHEN A.RET_ITEM_FLAG = 'N'  THEN B.SO_QTY ELSE (-1)*B.SO_QTY END"
	lgStrSQL = lgStrSQL &  " 				  ELSE 0  END,0 )) AS SO_QTY_15,"
	lgStrSQL = lgStrSQL &  " 		SUM(ISNULL(CASE WHEN RIGHT(CONVERT(CHAR(10),A.SO_DT,120),2)='16' THEN "
	lgStrSQL = lgStrSQL &  " 							CASE WHEN A.RET_ITEM_FLAG = 'N'  THEN B.SO_QTY ELSE (-1)*B.SO_QTY END"
	lgStrSQL = lgStrSQL &  " 				  ELSE 0  END,0 )) AS SO_QTY_16,"
	lgStrSQL = lgStrSQL &  " 		SUM(ISNULL(CASE WHEN RIGHT(CONVERT(CHAR(10),A.SO_DT,120),2)='17' THEN "
	lgStrSQL = lgStrSQL &  " 							CASE WHEN A.RET_ITEM_FLAG = 'N'  THEN B.SO_QTY ELSE (-1)*B.SO_QTY END"
	lgStrSQL = lgStrSQL &  " 				  ELSE 0  END,0 )) AS SO_QTY_17,"
	lgStrSQL = lgStrSQL &  " 		SUM(ISNULL(CASE WHEN RIGHT(CONVERT(CHAR(10),A.SO_DT,120),2)='18' THEN "
	lgStrSQL = lgStrSQL &  " 							CASE WHEN A.RET_ITEM_FLAG = 'N'  THEN B.SO_QTY ELSE (-1)*B.SO_QTY END"
	lgStrSQL = lgStrSQL &  " 				  ELSE 0  END,0 )) AS SO_QTY_18,"
	lgStrSQL = lgStrSQL &  " 		SUM(ISNULL(CASE WHEN RIGHT(CONVERT(CHAR(10),A.SO_DT,120),2)='19' THEN "
	lgStrSQL = lgStrSQL &  " 							CASE WHEN A.RET_ITEM_FLAG = 'N'  THEN B.SO_QTY ELSE (-1)*B.SO_QTY END"
	lgStrSQL = lgStrSQL &  " 				  ELSE 0  END,0 )) AS SO_QTY_19,"
	lgStrSQL = lgStrSQL &  " 		SUM(ISNULL(CASE WHEN RIGHT(CONVERT(CHAR(10),A.SO_DT,120),2)='20' THEN "
	lgStrSQL = lgStrSQL &  " 							CASE WHEN A.RET_ITEM_FLAG = 'N'  THEN B.SO_QTY ELSE (-1)*B.SO_QTY END"
	lgStrSQL = lgStrSQL &  " 				  ELSE 0  END,0 )) AS SO_QTY_20,"
	lgStrSQL = lgStrSQL &  " 		SUM(ISNULL(CASE WHEN RIGHT(CONVERT(CHAR(10),A.SO_DT,120),2)='21' THEN "
	lgStrSQL = lgStrSQL &  " 							CASE WHEN A.RET_ITEM_FLAG = 'N'  THEN B.SO_QTY ELSE (-1)*B.SO_QTY END"
	lgStrSQL = lgStrSQL &  " 				  ELSE 0  END,0 )) AS SO_QTY_21,"
	lgStrSQL = lgStrSQL &  " 		SUM(ISNULL(CASE WHEN RIGHT(CONVERT(CHAR(10),A.SO_DT,120),2)='22' THEN "
	lgStrSQL = lgStrSQL &  " 							CASE WHEN A.RET_ITEM_FLAG = 'N'  THEN B.SO_QTY ELSE (-1)*B.SO_QTY END"
	lgStrSQL = lgStrSQL &  " 				  ELSE 0  END,0 )) AS SO_QTY_22,"
	lgStrSQL = lgStrSQL &  " 		SUM(ISNULL(CASE WHEN RIGHT(CONVERT(CHAR(10),A.SO_DT,120),2)='23' THEN "
	lgStrSQL = lgStrSQL &  " 							CASE WHEN A.RET_ITEM_FLAG = 'N'  THEN B.SO_QTY ELSE (-1)*B.SO_QTY END"
	lgStrSQL = lgStrSQL &  " 				  ELSE 0  END,0 )) AS SO_QTY_23,"
	lgStrSQL = lgStrSQL &  " 		SUM(ISNULL(CASE WHEN RIGHT(CONVERT(CHAR(10),A.SO_DT,120),2)='24' THEN "
	lgStrSQL = lgStrSQL &  " 							CASE WHEN A.RET_ITEM_FLAG = 'N'  THEN B.SO_QTY ELSE (-1)*B.SO_QTY END"
	lgStrSQL = lgStrSQL &  " 				  ELSE 0  END,0 )) AS SO_QTY_24,"
	lgStrSQL = lgStrSQL &  " 		SUM(ISNULL(CASE WHEN RIGHT(CONVERT(CHAR(10),A.SO_DT,120),2)='25' THEN "
	lgStrSQL = lgStrSQL &  " 							CASE WHEN A.RET_ITEM_FLAG = 'N'  THEN B.SO_QTY ELSE (-1)*B.SO_QTY END"
	lgStrSQL = lgStrSQL &  " 				  ELSE 0  END,0 )) AS SO_QTY_25,"
	lgStrSQL = lgStrSQL &  " 		SUM(ISNULL(CASE WHEN RIGHT(CONVERT(CHAR(10),A.SO_DT,120),2)='26' THEN "
	lgStrSQL = lgStrSQL &  " 							CASE WHEN A.RET_ITEM_FLAG = 'N'  THEN B.SO_QTY ELSE (-1)*B.SO_QTY END"
	lgStrSQL = lgStrSQL &  " 				  ELSE 0  END,0 )) AS SO_QTY_26,"
	lgStrSQL = lgStrSQL &  " 		SUM(ISNULL(CASE WHEN RIGHT(CONVERT(CHAR(10),A.SO_DT,120),2)='27' THEN "
	lgStrSQL = lgStrSQL &  " 							CASE WHEN A.RET_ITEM_FLAG = 'N'  THEN B.SO_QTY ELSE (-1)*B.SO_QTY END"
	lgStrSQL = lgStrSQL &  " 				  ELSE 0  END,0 )) AS SO_QTY_27,"
	lgStrSQL = lgStrSQL &  " 		SUM(ISNULL(CASE WHEN RIGHT(CONVERT(CHAR(10),A.SO_DT,120),2)='28' THEN "
	lgStrSQL = lgStrSQL &  " 							CASE WHEN A.RET_ITEM_FLAG = 'N'  THEN B.SO_QTY ELSE (-1)*B.SO_QTY END"
	lgStrSQL = lgStrSQL &  " 				  ELSE 0  END,0 )) AS SO_QTY_28,"
	lgStrSQL = lgStrSQL &  " 		SUM(ISNULL(CASE WHEN RIGHT(CONVERT(CHAR(10),A.SO_DT,120),2)='29' THEN "
	lgStrSQL = lgStrSQL &  " 							CASE WHEN A.RET_ITEM_FLAG = 'N'  THEN B.SO_QTY ELSE (-1)*B.SO_QTY END"
	lgStrSQL = lgStrSQL &  " 				  ELSE 0  END,0 )) AS SO_QTY_29,"
	lgStrSQL = lgStrSQL &  " 		SUM(ISNULL(CASE WHEN RIGHT(CONVERT(CHAR(10),A.SO_DT,120),2)='30' THEN "
	lgStrSQL = lgStrSQL &  " 							CASE WHEN A.RET_ITEM_FLAG = 'N'  THEN B.SO_QTY ELSE (-1)*B.SO_QTY END"
	lgStrSQL = lgStrSQL &  " 				  ELSE 0  END,0 )) AS SO_QTY_30,"
	lgStrSQL = lgStrSQL &  " 		SUM(ISNULL(CASE WHEN RIGHT(CONVERT(CHAR(10),A.SO_DT,120),2)='31' THEN "
	lgStrSQL = lgStrSQL &  " 							CASE WHEN A.RET_ITEM_FLAG = 'N'  THEN B.SO_QTY ELSE (-1)*B.SO_QTY END	"
	lgStrSQL = lgStrSQL &  " 				  ELSE 0  END,0 )) AS SO_QTY_31  "                               
	lgStrSQL = lgStrSQL &  "           FROM     S_SO_HDR A ,  S_SO_DTL B,  B_BIZ_PARTNER C "                                                    
	lgStrSQL = lgStrSQL &  "           WHERE   A.SO_NO = B.SO_NO AND A.PAYER = C.BP_CD "
	lgStrSQL = lgStrSQL &  "           AND A.CFM_FLAG = 'Y'  AND B.PLANT_CD = " & FilterVar(Trim(PlantCd), "''", "S") & "  "
	lgStrSQL = lgStrSQL &  "           AND CONVERT(CHAR(7),A.SO_DT,120) = " & FilterVar(documentDt,"''","S") & " "
	lgStrSQL = lgStrSQL &  "           GROUP BY A.PAYER,  C.BP_NM,   dbo.ufn_GetItemGroupNM(B.ITEM_CD, " & lgGroupIndex & ") ,   A.SO_DT) TOT "
	lgStrSQL = lgStrSQL &  " GROUP BY TOT.PAYER, TOT.BP_NM, TOT.GROUP_NM " 
          
	
	
                                                
Response.Write lgStrSQL
'Response.End                       

End Sub

'============================================================================================================
' Name : CommonOnTransactionCommit
' Desc : This Sub is called by OnTransactionCommit Error handler
'============================================================================================================
Sub CommonOnTransactionCommit()
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
                .lgStrPrevKeyIndex = "<%=lgStrPrevKeyIndex%>"
                .ggoSpread.SSShowData "<%=lgstrData%>"
                .DBQueryOk        
	         End with
          End If   
       Case "<%=UID_M0002%>"                                                         '☜ : Save
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DBSaveOk
          Else
             Parent.SubSetErrPos(Trim("<%=lgErrorPos%>"))
          End If   
       Case "<%=UID_M0002%>"                                                         '☜ : Delete
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DbDeleteOk
          Else   
          End If   
    End Select    
    
       
</Script>	
 

    
    
