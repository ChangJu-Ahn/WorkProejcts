<!--
<%
'********************************************************************************************************
'*  1. Module Name          : Procuremant																*
'*  2. Function Name        :																			*
'*  3. Program ID           : m3212rb1.asp																*
'*  4. Program Name         :																			*
'*  5. Program Desc         : L/C내역참조 PopUp Transaction 처리용 ASP									*
'*  7. Modified date(First) : 2000/03/22																*
'*  8. Modified date(Last)  : 2001/12/10																*
'*  9. Modifier (First)     : Sun-jung Lee
'* 10. Modifier (Last)      : Jin-hyun Shin
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"									*
'*                            this mark(⊙) Means that "may  change"									*
'*                            this mark(☆) Means that "must change"									*
'* 13. History              : 1. 2000/03/22 : Coding Start												*
'********************************************************************************************************
%>
-->
<!-- #Include file="../../inc/IncServer.asp" -->
<%

   Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0,rs1               '☜ : DBAgent Parameter 선언 
   Dim lgStrData                                               '☜ : Spread sheet에 보여줄 데이타를 위한 변수 
   Dim iTotstrData
   Dim lgTailList
   Dim lgSelectList
   Dim lgSelectListDT
   Dim lgDataExist
   Dim lgPageNo
   
   Dim strItemCd												  '☜ : 품목코드 
   Dim strItemName												  '☜ : 품목명 
   Dim iFrPoint
   iFrPoint=0
      
   strItemCd	= ""
   strItemName	= ""		
   
    Call HideStatusWnd 
     
    lgPageNo         = UNICInt(Trim(Request("lgPageNo")),0)              '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
    lgSelectList     = Request("lgSelectList")
    lgTailList       = Request("lgTailList")
    lgSelectListDT   = Split(Request("lgSelectListDT"), gColSep)         '☜ : 각 필드의 데이타 타입 
    lgDataExist      = "No"

    Call  FixUNISQLData()                                                '☜ : DB-Agent로 보낼 parameter 데이타 set
    call  QueryData()                                                    '☜ : DB-Agent를 통한 ADO query

'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()

    Dim strVal
	
	Redim UNISqlId(1)                                                     '☜: SQL ID 저장을 위한 영역확보 
    Redim UNIValue(1,2)                                                 '⊙: DB-Agent로 전송될 parameter를 위한 변수 
                                                                        '    parameter의 수에 따라 변경함 
    UNISqlId(0) = "M3212RA300" 										' main query(spread sheet에 뿌려지는 query statement)
	UNISqlId(1) = "s0000qa016" 											' 품목명을 가져오기위한 Query
	
	'--- 2004-08-20 by Byun Jee Hyun for UNICODE	
    UNIValue(0,0) = Trim(lgSelectList)		                            '☜: Select 절에서 Summary    필드 
	
	strVal = " "

    If Len(Request("txtLCNo")) Then
		strVal = " AND A.LC_NO  = " & FilterVar(Trim(UCase(Request("txtLCNo"))), " " , "S") & " "
	Else
		strVal = ""
	End If	
	
	'2003.07 TrackingNo 추가 
    If Len(Request("txtTrackingNo")) Then
		strVal = strVal & " AND A.TRACKING_NO = " & FilterVar(Trim(UCase(Request("txtTrackingNo"))), " " , "S") & "  "		
	End If
	
	If Len(Request("txtItemCd")) Then
		strVal = strVal & " AND A.ITEM_CD = " & FilterVar(Trim(UCase(Request("txtItemCd"))), " " , "S") & " "
	End If

	If Len(Request("txtGrpCd")) Then
		strVal = strVal & " AND  B.PUR_GRP  = " & FilterVar(Trim(UCase(Request("txtGrpCd"))), " " , "S") & " "
	End If

	If Len(Request("txtBeneficiaryCd")) Then
		strVal = strVal & " AND  B.BENEFICIARY  = " & FilterVar(Trim(UCase(Request("txtBeneficiaryCd"))), " " , "S") & " "
	End If

    If Len(Trim(Request("txtCurrency"))) Then
		strVal = strVal & " AND B.Currency = " & FilterVar(Trim(UCase(Request("txtCurrency"))), " " , "S") & " "		
	End If		
	
	If Len(Trim(Request("txtPayMethCd"))) Then
		strVal = strVal & " AND B.PAY_METHOD = " & FilterVar(Trim(UCase(Request("txtPayMethCd"))), " " , "S") & " "	
	End If
	
	If Len(Trim(Request("txtIncoTermsCd"))) Then
		strVal = strVal & " AND B.incoterms = " & FilterVar(Trim(UCase(Request("txtIncoTermsCd"))), " " , "S") & " "	
	End If
	
    UNIValue(0,1) = strVal   
    

	'If Len(Trim(Request("txtItemCd"))) Then
		UNIValue(1,0) = FilterVar(Trim(UCase(Request("txtItemCd"))), " " , "S") 
	'End If

   
    UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList))			  '	UNISqlId(0)의 마지막 ?에 입력됨	
    
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '☜: set ADO read mode

End Sub

'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub QueryData()

    Dim lgstrRetMsg                                             '☜ : Record Set Return Message 변수선언 
    Dim lgADF                                                   '☜ : ActiveX Data Factory 지정 변수선언 
    Dim iStr
    
    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0,rs1)

	Set lgADF   = Nothing
	
    iStr = Split(lgstrRetMsg,gColSep)

	If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If 
    
    if SetConditionData = False then Exit sub
         
    If  rs0.EOF And rs0.BOF And FalsechkFlg =  False Then
        Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
        rs0.Close
        Set rs0 = Nothing
    Else    
        Call  MakeSpreadSheetData()
    End If  	
   
End Sub

'----------------------------------------------------------------------------------------------------------
' Make srpread sheet data
'----------------------------------------------------------------------------------------------------------
Sub MakeSpreadSheetData()

    Dim iLoopCount                                                                     
    Dim iRowStr
    Dim ColCnt
    Dim PvArr
    Const C_SHEETMAXROWS_D = 100 
    
    lgDataExist    = "Yes"
    lgstrData      = ""
  
    If CLng(lgPageNo) > 0 Then
       rs0.Move     = C_SHEETMAXROWS_D * CLng(lgPageNo)                  'C_SHEETMAXROWS_D:Max Fetched Count at once , lgStrPrevKeyIndex : Previous PageNo
       iFrPoint     = C_SHEETMAXROWS_D * CLng(lgPageNo)
    End If
    
    iLoopCount = -1
    ReDim PvArr(C_SHEETMAXROWS_D - 1)
    
   Do while Not (rs0.EOF Or rs0.BOF)
   
        iLoopCount =  iLoopCount + 1
        iRowStr = ""
        
		For ColCnt = 0 To UBound(lgSelectListDT) - 1 
            iRowStr = iRowStr & Chr(11) & FormatRsString(lgSelectListDT(ColCnt),rs0(ColCnt))
		Next
 
        If iLoopCount < C_SHEETMAXROWS_D Then
			lgstrData = lgstrData & iRowStr & Chr(11) & Chr(12)
			PvArr(iLoopCount) = lgstrData	
			lgstrData = ""
        Else
           lgPageNo = lgPageNo + 1
           Exit Do
        End If
        
        rs0.MoveNext
	Loop
	
	iTotstrData = Join(PvArr, "")
	
    If iLoopCount < C_SHEETMAXROWS_D Then                                      '☜: Check if next data exists
       lgPageNo = ""
    End If
    rs0.Close                                                       '☜: Close recordset object
    Set rs0 = Nothing	                                            '☜: Release ADF

End Sub

'----------------------------------------------------------------------------------------------------------
' Name : SetConditionData
' Desc : set value in condition area
'----------------------------------------------------------------------------------------------------------
Function SetConditionData()
    
    SetConditionData = False
    
    If Not(rs1.EOF Or rs1.BOF) Then
       strItemName =  rs1(0)
       Set rs1 = Nothing 
 	else
	    Set rs1 = Nothing
		If Len(Request("txtItemCd")) Then
			Call DisplayMsgBox("970000", vbInformation, "품목", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code	
		    Exit function
		End If   
    End If   
    
    SetConditionData = True
    
End Function

%>
<Script Language=vbscript>

	parent.frm1.txtItemNm.value = "<%=ConvSPChars(strItemName)%>"		
	
    If "<%=lgDataExist%>" = "Yes" Then
       parent.ggoSpread.Source  = parent.frm1.vspdData
       Parent.frm1.vspdData.Redraw = False
       parent.ggoSpread.SSShowData "<%=iTotstrData%>","F"          '☜ : Display data
		Call parent.ReFormatSpreadCellByCellByCurrency2(parent.frm1.vspdData,"<%=iFrPoint+1%>",parent.frm1.vspddata.maxrows,"<%=ConvSPChars(Request("txtCurrency"))%>",parent.GetKeyPos("A",7),"C","I","X","X")
       parent.lgPageNo      =  "<%=lgPageNo%>"               '☜ : Next next data tag
       parent.DbQueryOk
       Parent.frm1.vspdData.Redraw = True
    End If  
    
</Script>	

