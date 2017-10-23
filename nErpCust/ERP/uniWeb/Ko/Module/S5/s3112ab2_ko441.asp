<% Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../comasp/loadinftb19029.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->

<% 
    On Error Resume Next                                                   '☜: Protect prorgram from crashing

    Call HideStatusWnd 
    Call LoadBasisGlobalInf()
    Call loadInfTB19029B("I", "*","NOCOOKIE","MB") 
   
    Dim iDx
    Dim iLngCnt
    Dim iArrExport
    Dim iStrData
    Dim TmpBuffer
    Dim iStrSlCd


    '---------- Developer Coding part (Start) ---------------------------------------------------------------

	lgStrSQL = lgStrSQL & " select 0 ,b.OUT_NO, " & vbcrlf _
		& FilterVar(Request("txtShipToParty"),"''","S") & " as BP_CD " & vbcrlf _
		& " ,c.BP_NM " & vbcrlf _
		& " ,a.ITEM_CD " & vbcrlf _
		& " ,d.ITEM_NM " & vbcrlf _
		& " ,d.SPEC " & vbcrlf _
		& " ,a.PLANT_CD " & vbcrlf _
		& " ,b.OUT_TYPE " & vbcrlf _
		& " ,e.UD_MINOR_NM " & vbcrlf _
		& " ,a.GOOD_ON_HAND_QTY " & vbcrlf _
		& " ,b.GI_QTY " & vbcrlf _
		& " ,b.GI_UNIT " & vbcrlf _
		& " ,a.LOT_NO " & vbcrlf _
		& " ,a.LOT_SUB_NO " & vbcrlf _
		& " ,b.ACTUAL_GI_DT,b.CUST_LOT_NO1 cust_lot_no ,a.SL_CD,b.TRANS_TIME,b.CREATE_TYPE,REPLACE(REPLACE(b.cust_lot_no, CHAR(13), ''), CHAR(10), '') rcpt_lot_no " & vbcrlf _
		& " , rtrim(b.PGM_NAME) as pgm_name " & vbcrlf _
		& " , ISNULL([DBO].[ufn_GetPrice]("&FilterVar(Request("txtShipToParty"),"''","S")&", a.ITEM_CD, " & FilterVar(Request("txtSoNo"),"''","S") & ", b.GI_UNIT, b.PGM_NAME),0) pgm_price " & vbcrlf _
		& " from i_onhand_stock_detail a (nolock) " & vbcrlf		
'원본
	'2009.08.05 MES 품목코드로 ERP품목코드 검색시 복수의 자료가 나오는 부분을 해결하기 위해 LOT번호로 제조오더를 검색하여, 제조오더의 품목코드를 검색하게 변경....시작
	'lgStrSQL = lgStrSQL & " inner join T_IF_RCV_VIRTURE_OUT_KO441 b (nolock) on (a.PLANT_CD=b.PLANT_CD and a.ITEM_CD=[DBO].[UFN_GETITEMCD](b.MES_ITEM_CD) " & vbcrlf _
	'& " AND a.TRACKING_NO='*' AND a.LOT_NO=b.LOT_NO AND a.LOT_SUB_NO=0) " & vbcrlf
'수정본
	lgStrSQL = lgStrSQL & " inner join T_IF_RCV_VIRTURE_OUT_KO441 b (nolock) on (a.PLANT_CD=b.PLANT_CD and a.ITEM_CD=[DBO].[ufn_GetProdOrdItemCd1](a.LOT_NO, a.PLANT_CD) " & vbcrlf _
	 & " AND a.TRACKING_NO='*' AND a.LOT_NO=b.LOT_NO AND a.LOT_SUB_NO=0) " & vbcrlf
	'2009.08.05 MES 품목코드로 ERP품목코드 검색시 복수의 자료가 나오는 부분을 해결하기 위해 LOT번호로 제조오더를 검색하여, 제조오더의 품목코드를 검색하게 변경....종료


	lgStrSQL = lgStrSQL & " inner join ( " & vbcrlf _
		& " 					select OUT_NO,TRANS_TIME  " & vbcrlf _
		& " 					from T_IF_RCV_VIRTURE_OUT_KO441  " & vbcrlf _
		& " 					group by OUT_NO,TRANS_TIME having count(*) <> 2 " & vbcrlf _
		& " 					) b2 on (b.OUT_NO=b2.OUT_NO and b.TRANS_TIME=b2.TRANS_TIME) " & vbcrlf _
		& " inner join ( " & vbcrlf _
		& " 			select OUT_NO, Max(Convert(varchar(10), ACTUAL_GI_DT,121) + TRANS_TIME) as ACTUAL_GI_DT " & vbcrlf _
		& " 			from   T_IF_RCV_VIRTURE_OUT_KO441  " & vbcrlf _
		& " 			where  mes_item_cd in (select item_nm from b_item where item_cd =  " & FilterVar(Request("txtItemCd"),"''","S") & ")" & vbcrlf _
		& " 			group by OUT_NO  " & vbcrlf _
		& " 		   ) b3 on (	b.OUT_NO	 = b3.OUT_NO " & vbcrlf _
		& " 			    and Convert(varchar(10), b.ACTUAL_GI_DT,121) + b.TRANS_TIME  = b3.ACTUAL_GI_DT " & vbcrlf _
		& " 			   )  " & vbcrlf 


	'2009.11.03 MES 거래처명으로 ERP 거래처정보의 ALIAS_NM으로 검색할때 1개만 하게 한 것을 복수개로 변경
	'원본
	lgStrSQL = lgStrSQL _
		& " inner join B_BIZ_PARTNER c (nolock) on ((SELECT top 1 BP_CD FROM B_BIZ_PARTNER WHERE BP_ALIAS_NM=b.SHIP_TO_PARTY and USAGE_FLAG = 'y')=c.BP_CD) " & vbcrlf 

	'수정본  (박영산과장 연락시점 부터 적용 예정)
	'lgStrSQL = lgStrSQL _
	'	& " inner join B_BIZ_PARTNER c (nolock) on ( c.BP_CD in (SELECT top 1 BP_CD FROM B_BIZ_PARTNER WHERE BP_ALIAS_NM=b.SHIP_TO_PARTY and USAGE_FLAG = 'y') ) " & vbcrlf 

	lgStrSQL = lgStrSQL _
		& " inner join B_ITEM d (nolock) on (a.ITEM_CD=d.ITEM_CD) " & vbcrlf _
		& " inner join B_USER_DEFINED_MINOR e (nolock)  on (e.UD_MAJOR_CD='ZZ002' and b.OUT_TYPE=e.UD_MINOR_CD) " & vbcrlf _
		& " where a.PLANT_CD=" & FilterVar(Request("txtPlantCd"),"''","S") & vbcrlf _
		& " AND a.ITEM_CD=" & FilterVar(Request("txtItemCd"),"''","S") & vbcrlf _
		& " AND c.bp_cd =" & FilterVar(Request("txtShipToParty"),"''","S") & vbcrlf _
		& " AND convert(varchar(10),b.ACTUAL_GI_DT,102) = " & FilterVar(Request("txtShipDate"),"''","S") & vbcrlf '20110725 added by songth
		'2009.09.04  거래처창고와 제품창고를 사용자가 선택 가능하게 변경
		If Ucase(Trim(Request("txtPlantCd"))) = "P01" Then
				lgStrSQL = lgStrSQL & " AND a.SL_CD='01100" & Trim(Request("txtSLRadio")) & "'" & vbcrlf
		ElseIf Ucase(Trim(Request("txtPlantCd"))) = "P02" Then
				lgStrSQL = lgStrSQL & " AND a.SL_CD='02100" & Trim(Request("txtSLRadio")) & "'" & vbcrlf
		Else 
				'T_IF_RCV_PART_OUT_KO441.SHIP_TO_PARTY_LINE 컬럼있음. 조인 테이블을 T_IF_RCV_VIRTURE_OUT_KO441 변경 후 관련컬럼이 없어 주석처리함.
				'lgStrSQL = lgStrSQL & " AND a.SL_CD in (select UD_REFERENCE from B_USER_DEFINED_MINOR where ud_major_cd='zz005' and UD_MINOR_CD=b.SHIP_TO_PARTY_LINE) " & vbcrlf	
		End If

		lgStrSQL = lgStrSQL & " AND a.GOOD_ON_HAND_QTY>0 " & vbcrlf _
				& " AND e.UD_REFERENCE = 'Y' " & vbcrlf _
				& " AND ISNULL(b.ERP_APPLY_FLAG,'N') <> 'Y' " & vbcrlf _
				& " 		order by b.OUT_NO " & vbcrlf 


    Call SubOpenDB(lgObjConn)                                                        '☜: Make a DB Connection

    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then                   'R(Read) X(CursorType) X(LockType) 
        Call SubCloseRs(lgObjRs)
	Call SubCloseDB(lgObjConn)
        Response.End
    End If
  
    iArrExport = lgObjRs.GetRows()
  
    Call SubCloseRs(lgObjRs)
    Call SubCloseDB(lgObjConn)

	iLngCnt = Ubound(iArrExport, 2)

	Redim TmpBuffer(iLngCnt)

	' 0: OUT_NO / 1: BP_CD / 2: BP_NM / 3: ITEM_CD
	' 4: ITEM_NM / 5: SPEC / 6: PLANT_CD / 7: OUT_TYPE
        ' 8: UD_MINOR_NM / 9: GOOD_ON_HAND_QTY / 10: GI_QTY / 
	' 11: GI_UNIT / 12: LOT_NO / 13: LOT_SUB_NO / 14: ACTUAL_GI_DT
	' 15: cust_lot_no / 16: SL_CD / 17: TRANS_TIME / 18: CREATE_TYPE
	' 19: rcpt_lot_no / 20: pgm_name / 21: pgm_price 

	For iDx = 0 To iLngCnt
	iStrData = Chr(11) & ConvSPChars(Trim(iArrExport(0, iDx))) _
			& Chr(11) & ConvSPChars(Trim(iArrExport(1, iDx))) _
			& Chr(11) & ConvSPChars(Trim(iArrExport(2, iDx))) _
			& Chr(11) & ConvSPChars(Trim(iArrExport(3, iDx))) _
			& Chr(11) & ConvSPChars(Trim(iArrExport(4, iDx))) _
			& Chr(11) & ConvSPChars(Trim(iArrExport(7, iDx))) _
			& Chr(11) & ConvSPChars(Trim(iArrExport(17, iDx))) _	
			& Chr(11) & ConvSPChars(Trim(iArrExport(8, iDx))) _	
			& Chr(11) & ConvSPChars(Trim(iArrExport(9, iDx))) _	
			& Chr(11) & UNINumClientFormat(iArrExport(10, iDx), ggQty.DecPoint, 0) _	
			& Chr(11) & UNINumClientFormat(iArrExport(11, iDx), ggQty.DecPoint, 0) _	
			& Chr(11) & ConvSPChars(Trim(iArrExport(12, iDx))) _	
			& Chr(11) & ConvSPChars(Trim(iArrExport(13, iDx))) _	
			& Chr(11) & ConvSPChars(Trim(iArrExport(14, iDx))) _	
			& Chr(11) & UNIDateClientFormat(iArrExport(15, iDx)) _	
			& Chr(11) & ConvSPChars(Trim(iArrExport(5, iDx))) _	
			& Chr(11) & ConvSPChars(Trim(iArrExport(6, iDx))) _	
			& Chr(11) & ConvSPChars(Trim(iArrExport(20, iDx))) _	
			& Chr(11) & ConvSPChars(Trim(iArrExport(16, iDx))) _	
			& Chr(11) & ConvSPChars(Trim(iArrExport(21, iDx))) _	
			& Chr(11) & UNINumClientFormat(iArrExport(22, iDx), ggQty.DecPoint, 0) _	
			& Chr(11) & ConvSPChars(Trim(iArrExport(18, iDx))) _	
			& Chr(11) & ConvSPChars(Trim(iArrExport(19, iDx))) _		
			& Chr(11) & iDx _
			& Chr(11) & Chr(12)
		TmpBuffer(iDx) = iStrData
	Next



	lgStrData = Join(TmpBuffer, "")
'       Response.Write  " <Script Language=vbscript>                                  " & vbCr
'       Response.Write  "    Parent.ggoSpread.Source     = Parent.vspdData2       " & vbCr       
'	Response.Write  "    Parent.ggoSpread.SSShowDataByClip   """ & lgStrData & """" & vbCr       
'       Response.Write  " </Script>             " & vbCr
Response.Write "<Script language=vbs> " & vbCr   
Response.Write "With parent " & vbCr   
		Response.Write " .ggoSpread.Source = .vspdData2" & vbCr
		Response.Write " .vspdData2.Redraw = False  "      & vbCr      
		Response.Write " .ggoSpread.SSShowDataByClip   """ & lgStrData & """" & vbCr       
		Response.Write " .lgStrPrevKey = """ & iStrNextKey  & """" & vbCr  
		Response.Write "End With " & vbCr   		
		Response.Write "</Script> " & vbCr   
%>


