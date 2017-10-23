<%@ LANGUAGE=VBSCript%>
<% Option Explicit%>

<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/lgSvrVariables.inc" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<%
	Const C_SHEETMAXROWS_D = 100
    Call LoadBasisGlobalInf()
    Call LoadInfTB19029B("Q", "H", "NOCOOKIE", "MB")

    Call HideStatusWnd                                                               '��: Hide Processing message

    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '��: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '��: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)
'    lgCurrentSpd      = Request("lgCurrentSpd")                                      '��: "M"(Spread #1) "S"(Spread #2)

    lgLngMaxRow       = Request("txtMaxRows")                                        '��: Read Operation Mode (CRUD)
    

     Call SubCreateCommandObject(lgObjComm)
     Call SubBizQuery()
     Call SubCloseCommandObject(lgObjComm)

 '============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()

	Dim oRs, sTxt, arrRows, iLngRow, iLngCol, iStrData, sNextKey, sRowSeq, iLngRowCnt, iLngColCnt, sGrpTxt
	Dim sDstbOrder, sSenderCostCd, sFromAcctCd, sToAcctCd, sType, sSendAmt, sRecvAmt, arrTmp
    Dim DFnm 

    Call SubCreateCommandObject(lgObjComm)	


    With lgObjComm
		.CommandTimeout = 0

		.CommandText = "dbo.usp_h_retire_hometax"  
	    .CommandType = adCmdStoredProc

		lgObjComm.Parameters.Append lgObjComm.CreateParameter("RETURN_VALUE",  adInteger,adParamReturnValue)	' -- No ���� 

		' -- �����ؾ��� ��ȸ���� �Ķ��Ÿ�� 
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@BIZ_AREA_CD",	adVarXChar,	adParamInput, 10, lgKeyStream(1) )	'�Ű����� 
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@SEND_DT",		adVarXChar,	adParamInput, 8,  lgKeyStream(2) )	'���⿬���� 
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@RETIRE_DT1",		adVarXChar,	adParamInput, 8,  lgKeyStream(3) )	'�Ⱓ`
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@RETIRE_DT2",		adVarXChar,	adParamInput, 8,  lgKeyStream(4) )	'�Ⱓ`
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@ALL_YN",		adVarXChar,	adParamInput, 10,  lgKeyStream(5) )	'���սŰ��� 
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@GUBUN",			adVarXChar,	adParamInput, 1,  lgKeyStream(6) )	'�����ڱ��� 
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@GIGAN",			adVarXChar,	adParamInput, 1,  lgKeyStream(7) )	'���Ⱓ 
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@SEMU",			adVarXChar,	adParamInput, 6,  lgKeyStream(8) )	'�����븮�ΰ�����ȣ 


	 Set oRs = lgObjComm.Execute
    End With
 dim li_biz_own_rgst_no
'-----------------------------------------------------------------
' SP���� A,B,C,D,E ���ڵ� ���ÿ� RETURN
 
    If Not oRs.EOF Then

' ------------- A ���ڵ� 
		Dim arrColRowA, i, j
	   
	    
	    
		arrColRowA = oRs.GetRows()
		iLngRowCnt	= UBound(arrColRowA, 2) 
		iLngColCnt	= UBound(arrColRowA, 1) 
        li_biz_own_rgst_no = split(arrColRowA(0,0),chr(11))(9)

					  
		For i = 0 To iLngRowCnt
					
			For j = 0 To  iLngColCnt 
				lgstrData = lgstrData  & arrColRowA(j, i) & Chr(11)
						
			Next
					
			lgstrData = lgstrData &  lgLngMaxRow + j
			lgstrData = lgstrData &  Chr(12)
		Next

		Set oRs = oRs.NextRecordSet()	' -- ����(����Ÿ) ���ڵ������ ���� 


' -------------B ���ڵ� 
		If oRs.EOF = TRUE Then
 			Set oRs = oRs.NextRecordSet()	' -- ����(����Ÿ) ���ڵ������ ����		
		Else
 			Dim arrColRowB
			arrColRowB		= oRs.GetRows()
		
			Set oRs = oRs.NextRecordSet()	' -- ����(����Ÿ) ���ڵ������ ���� 
			iLngRowCnt	= UBound(arrColRowB, 2) 
			iLngColCnt	= UBound(arrColRowB, 1)
		
 			For i = 0 To iLngRowCnt

				For j = 0 To  iLngColCnt 
					lgstrData1 = lgstrData1 & arrColRowB(j, i) & Chr(11)
						
				Next
					
				lgstrData1 = lgstrData1 & lgLngMaxRow + j
				lgstrData1 = lgstrData1 & Chr(12)
			Next
  		End If
  		
  		

' ------------- C ���ڵ� 
		If oRs.EOF = TRUE Then
 			Set oRs = oRs.NextRecordSet()	' -- ����(����Ÿ) ���ڵ������ ����		
		Else

 			Dim arrColRowC
			arrColRowC		= oRs.GetRows()
		
			Set oRs = oRs.NextRecordSet()	' -- ����(����Ÿ) ���ڵ������ ���� 
			iLngRowCnt	= UBound(arrColRowC, 2) 
			iLngColCnt	= UBound(arrColRowC, 1)
		
 			For i = 0 To iLngRowCnt

				For j = 0 To  iLngColCnt 
					lgstrData2 = lgstrData2 &  arrColRowC(j, i) & Chr(11)
						
				Next
					
				lgstrData2 = lgstrData2 & lgLngMaxRow + j
				lgstrData2 = lgstrData2 &  Chr(12)
			Next
 		End If
 ' ------------- D ���ڵ� 

		If oRs.EOF = TRUE Then
 			Set oRs = oRs.NextRecordSet()	' -- ����(����Ÿ) ���ڵ������ ����		
		Else

 			Dim arrColRowD
			arrColRowD		= oRs.GetRows()
		
			Set oRs = oRs.NextRecordSet()	' -- ����(����Ÿ) ���ڵ������ ���� 
			iLngRowCnt	= UBound(arrColRowD, 2) 
			iLngColCnt	= UBound(arrColRowD, 1)
		
 			For i = 0 To iLngRowCnt

				For j = 0 To  iLngColCnt 
					lgstrData3 = lgstrData3 & arrColRowD(j, i) & Chr(11)
						
				Next
					
				lgstrData3 = lgstrData3 & lgLngMaxRow + j
				lgstrData3 = lgstrData3 & Chr(12)
			Next
		End If
		
		
		

            li_biz_own_rgst_no = Left(li_biz_own_rgst_no,7) & "." & Right(li_biz_own_rgst_no,3)

			DFnm = "C:\e" & li_biz_own_rgst_no       
            
%>
<SCRIPT LANGUAGE=VBSCRIPT>
		parent.frm1.txtFile.value = "<%=DFnm%>"
</SCRIPT>
<%      
'		End If
	
    Else 
       Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '�� : No data is found. 
    End If       
End Sub	
%>


<Script Language="VBScript">

    Select Case "<%=lgOpModeCRUD %>"
       Case "<%=UID_M0001%>"                                                         '�� : Query
          If Trim("<%=lgErrorStatus%>") = "NO" Then
              With Parent
              
                .ggoSpread.Source     = .frm1.vspdData
				.ggoSpread.SSShowData "<%=lgstrData%>"
 
                .ggoSpread.Source     = .frm1.vspdData1
				.ggoSpread.SSShowData "<%=lgstrData1%>"

                .ggoSpread.Source     = .frm1.vspdData2
				.ggoSpread.SSShowData "<%=lgstrData2%>"

                .ggoSpread.Source     = .frm1.vspdData3
				.ggoSpread.SSShowData "<%=lgstrData3%>"

                End with
          End If   
    End Select    
       
</Script>	
