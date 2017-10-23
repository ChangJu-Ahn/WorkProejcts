<!-- 이 페이지는 nepes에서 개발한 주식정보 입니다. -->
<!-- 이 페이지는 nepes에서 개발한 주식정보 입니다. -->
<!-- 이 페이지는 nepes에서 개발한 주식정보 입니다. -->
<!-- 이 페이지는 nepes에서 개발한 주식정보 입니다. -->
<!-- 이 페이지는 nepes에서 개발한 주식정보 입니다. -->

<%@ page language="java" contentType="text/html; charset=UTF-8"
	pageEncoding="UTF-8"%>
<%@ page import="java.beans.XMLEncoder"%>
<%@ page import="org.w3c.dom.*"%>
<%@ page import="org.xml.sax.*"%>
<%@ page import="java.util.*"%>
<%@ page import="java.io.*"%>
<%@ page import="java.net.*"%>
<%@ page import="javax.xml.parsers.*"%>
<%@ page import="javax.servlet.http.HttpServletResponse.*"%>
<%@ page import="java.text.*"%>
<%
	String geturl = "http://asp1.krx.co.kr/servlet/krx.asp.XMLSise?code=033640";
	String JongCd = geturl.substring(51, 57);
	String gettime = "";
	String janggubun = "";
	String DungRakrate_str = "";

	int timeconclude_length = 0;
	int dailystock_length = 0;
	int Askprice_length = 0;
	int Hoga_length = 0;

	int CurJuka = 0;
	int Debi = 0;
	float StandardPrice = 0;
	float DungRakrate = 0;

	String Bigup = "<img src=\"./img/updn15_1.gif\" align=\"absmidde\" alt=\"상승\" >";
	String Bigdown = "<img src=\"./img/updn15_2.gif\" align=\"absmidde\" alt=\"하락\" >";
	String up = "<img src=\"./img/updn09_1.gif\" align=\"absmidde\" alt=\"상승\" >";
	String down = "<img src=\"./img/updn09_2.gif\" align=\"absmidde\" alt=\"하락\" >";
	String bohab = "-";
	String line = "";
	String xml = "";

	String Stockinfo[] = new String[17];
	String Timeconclude[][] = new String[10][7];
	String Dailystock[][] = new String[10][9];
	String Askprice[][] = new String[5][4];
	String Hoga[] = new String[22];

	try {
		URL url = new URL(geturl);
		URLConnection conn = url.openConnection();
		HttpURLConnection httpConnection = (HttpURLConnection) conn;
		InputStream is = null;
		InputStreamReader isr = null;

		is = new URL(geturl).openStream();
		isr = new InputStreamReader(is, "UFT-8");

		BufferedReader rd = new BufferedReader(isr, 400);

		StringBuffer strbuf = new StringBuffer();

		//xml line1 공백제거
		while ((line = rd.readLine()) != null) {
			strbuf.append(line);
		}

		//System.out.println("주가정보");
		//System.out.println(strbuf.toString()); //xml파싱확인

		DocumentBuilderFactory docFact = DocumentBuilderFactory.newInstance();
		docFact.setNamespaceAware(true);
		DocumentBuilder docBuild = docFact.newDocumentBuilder();

		Document doc = docBuild.parse(new InputSource(new StringReader(strbuf.toString())));

		/*주가정보*/

		NodeList stockInfo = doc.getElementsByTagName("stockInfo");

		NamedNodeMap stockinfo = stockInfo.item(0).getAttributes();
		gettime = stockinfo.getNamedItem("myNowTime").getNodeValue();
		janggubun = stockinfo.getNamedItem("myJangGubun").getNodeValue();

		NodeList TBL_StockInfo = doc.getElementsByTagName("TBL_StockInfo");
		NamedNodeMap StockInfo = TBL_StockInfo.item(0).getAttributes();

		Stockinfo[0] = StockInfo.getNamedItem("JongName").getNodeValue(); //종목명 
		Stockinfo[1] = StockInfo.getNamedItem("CurJuka").getNodeValue(); //현재가 
		Stockinfo[2] = StockInfo.getNamedItem("DungRak").getNodeValue(); //전일대비코드
		Stockinfo[3] = StockInfo.getNamedItem("Debi").getNodeValue(); //전일대비
		Stockinfo[4] = StockInfo.getNamedItem("PrevJuka").getNodeValue(); //전일종가 
		Stockinfo[5] = StockInfo.getNamedItem("Volume").getNodeValue(); //거래량
		Stockinfo[6] = StockInfo.getNamedItem("Money").getNodeValue(); //거래대금  
		Stockinfo[7] = StockInfo.getNamedItem("StartJuka").getNodeValue(); //시가 
		Stockinfo[8] = StockInfo.getNamedItem("HighJuka").getNodeValue(); //고가
		Stockinfo[9] = StockInfo.getNamedItem("LowJuka").getNodeValue(); //저가 		
		Stockinfo[10] = StockInfo.getNamedItem("High52").getNodeValue(); //52주 최고 
		Stockinfo[11] = StockInfo.getNamedItem("Low52").getNodeValue(); //52주 최저  
		Stockinfo[12] = StockInfo.getNamedItem("UpJuka").getNodeValue(); //상한가 
		Stockinfo[13] = StockInfo.getNamedItem("DownJuka").getNodeValue(); //하한가 
		Stockinfo[14] = StockInfo.getNamedItem("Per").getNodeValue(); //PER            
		Stockinfo[15] = StockInfo.getNamedItem("Amount").getNodeValue(); //상장주식수    
		Stockinfo[16] = StockInfo.getNamedItem("FaceJuka").getNodeValue(); //액면가

		// 등락율 계산
		CurJuka = Integer.parseInt(Stockinfo[1].replaceAll(",", ""));
		Debi = Integer.parseInt(Stockinfo[3].replaceAll(",", ""));

		/*등락구분코드*/
		// 1 - 상한, 2 - 상승, 3 - 보합, 4 - 하한, 5 - 하락

		if (Stockinfo[2].equals("1") || Stockinfo[2].equals("2") || Stockinfo[2].equals("3")) {
			StandardPrice = CurJuka - Debi;
		} else if (Stockinfo[2].equals("4") || Stockinfo[2].equals("5")) {
			StandardPrice = CurJuka + Debi;
		}
		else{
			StandardPrice = CurJuka;
		}
		
		// 등락률 = (당일종가 - 기준가) / 기준가 * 100
		// 기준가 = 당일종가(현재가) - 전일대비
		DungRakrate = ((CurJuka - StandardPrice) / StandardPrice) * 100;
		DungRakrate_str = String.format("%.2f", DungRakrate);

		/*일자별시세*/

		NodeList TBL_Dailystock = doc.getElementsByTagName("DailyStock");

		dailystock_length = TBL_Dailystock.getLength();

		for (int j = 0; j < dailystock_length; j++) {

			NamedNodeMap DailyStock = TBL_Dailystock.item(j).getAttributes();

			Dailystock[j][0] = DailyStock.getNamedItem("day_Date").getNodeValue(); //일자
			Dailystock[j][1] = DailyStock.getNamedItem("day_EndPrice").getNodeValue(); //종가
			Dailystock[j][2] = DailyStock.getNamedItem("day_getDebi").getNodeValue(); //전일대비
			Dailystock[j][3] = DailyStock.getNamedItem("day_Start").getNodeValue(); //시가
			Dailystock[j][4] = DailyStock.getNamedItem("day_High").getNodeValue(); //고가
			Dailystock[j][5] = DailyStock.getNamedItem("day_Low").getNodeValue(); //저가
			Dailystock[j][6] = DailyStock.getNamedItem("day_Volume").getNodeValue(); //거래량
			Dailystock[j][7] = DailyStock.getNamedItem("day_getAmount").getNodeValue(); //거래대금
			Dailystock[j][8] = DailyStock.getNamedItem("day_Dungrak").getNodeValue(); //전일대비코드

		}

		/*시간대별 체결가*/

		NodeList TBL_TimeConclude = doc.getElementsByTagName("TBL_TimeConclude");

		timeconclude_length = TBL_TimeConclude.getLength() - 1;
		for (int i = 0; i < timeconclude_length; i++) {

			NamedNodeMap TimeConclude = TBL_TimeConclude.item(i + 1).getAttributes();

			Timeconclude[i][0] = TimeConclude.getNamedItem("time").getNodeValue(); //시간
			Timeconclude[i][1] = TimeConclude.getNamedItem("negoprice").getNodeValue(); //체결가
			Timeconclude[i][2] = TimeConclude.getNamedItem("Debi").getNodeValue(); //전일대비
			Timeconclude[i][3] = TimeConclude.getNamedItem("sellprice").getNodeValue(); //매도호가
			Timeconclude[i][4] = TimeConclude.getNamedItem("buyprice").getNodeValue(); //매수호가
			Timeconclude[i][5] = TimeConclude.getNamedItem("amount").getNodeValue(); //체결량
			Timeconclude[i][6] = TimeConclude.getNamedItem("Dungrak").getNodeValue(); //전일대비코드
		}

		/*증권사별거래*/

		NodeList TBL_AskPrice = doc.getElementsByTagName("AskPrice");

		Askprice_length = TBL_AskPrice.getLength();
		for (int i = 0; i < Askprice_length; i++) {

			NamedNodeMap AskPrice = TBL_AskPrice.item(i).getAttributes();

			Askprice[i][0] = AskPrice.getNamedItem("member_memdoMem").getNodeValue(); //매도증권사
			Askprice[i][1] = AskPrice.getNamedItem("member_memdoVol").getNodeValue(); //매도거래량
			Askprice[i][2] = AskPrice.getNamedItem("member_memsoMem").getNodeValue(); //매수증권사
			Askprice[i][3] = AskPrice.getNamedItem("member_mesuoVol").getNodeValue(); //매수거래량
		}

		/*호가*/

		NodeList TBL_Hoga = doc.getElementsByTagName("TBL_Hoga");

		Hoga_length = TBL_Hoga.getLength();

		NamedNodeMap hoga = TBL_Hoga.item(0).getAttributes();

		Hoga[0] = hoga.getNamedItem("mesuJan0").getNodeValue(); //매수잔량
		Hoga[1] = hoga.getNamedItem("mesuHoka0").getNodeValue(); //매수호가
		Hoga[2] = hoga.getNamedItem("mesuJan1").getNodeValue(); //매수잔량
		Hoga[3] = hoga.getNamedItem("mesuHoka1").getNodeValue(); //매수호가
		Hoga[4] = hoga.getNamedItem("mesuJan2").getNodeValue(); //매수잔량
		Hoga[5] = hoga.getNamedItem("mesuHoka2").getNodeValue(); //매수호가
		Hoga[6] = hoga.getNamedItem("mesuJan3").getNodeValue(); //매수잔량
		Hoga[7] = hoga.getNamedItem("mesuHoka3").getNodeValue(); //매수호가
		Hoga[8] = hoga.getNamedItem("mesuJan4").getNodeValue(); //매수잔량
		Hoga[9] = hoga.getNamedItem("mesuHoka4").getNodeValue(); //매수호가
		Hoga[10] = hoga.getNamedItem("medoHoka0").getNodeValue(); //매도잔량
		Hoga[11] = hoga.getNamedItem("medoJan0").getNodeValue(); //매도호가
		Hoga[12] = hoga.getNamedItem("medoHoka1").getNodeValue(); //매도잔량
		Hoga[13] = hoga.getNamedItem("medoJan1").getNodeValue(); //매도호가
		Hoga[14] = hoga.getNamedItem("medoHoka2").getNodeValue(); //매도잔량
		Hoga[15] = hoga.getNamedItem("medoJan2").getNodeValue(); //매도호가
		Hoga[16] = hoga.getNamedItem("medoHoka3").getNodeValue(); //매도잔량
		Hoga[17] = hoga.getNamedItem("medoJan3").getNodeValue(); //매도호가
		Hoga[18] = hoga.getNamedItem("medoHoka4").getNodeValue(); //매도잔량
		Hoga[19] = hoga.getNamedItem("medoJan4").getNodeValue(); //매도호가	

		DecimalFormat formatter = new DecimalFormat("###,###,###");

		Hoga[20] = formatter.format(Integer.parseInt(Hoga[0].replace(",", ""))
				+ Integer.parseInt(Hoga[2].replace(",", "")) + Integer.parseInt(Hoga[4].replace(",", ""))
				+ Integer.parseInt(Hoga[6].replace(",", "")) + Integer.parseInt(Hoga[8].replace(",", "")));
		Hoga[21] = formatter.format(Integer.parseInt(Hoga[11].replace(",", ""))
				+ Integer.parseInt(Hoga[13].replace(",", "")) + Integer.parseInt(Hoga[15].replace(",", ""))
				+ Integer.parseInt(Hoga[17].replace(",", "")) + Integer.parseInt(Hoga[19].replace(",", "")));

	} catch (Exception e) {

	}
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=UFT-8">
<!-- jQuery -->
<script src="//code.jquery.com/jquery-1.11.3.min.js"></script>
<script src="//code.jquery.com/ui/1.11.4/jquery-ui.js"></script>
<script src="<c:url value='/js/common.js'/>" charset="utf-8"></script>
<script src="<c:url value='/js/menu.js'/>" charset="utf-8"></script>
<script src="<c:url value='/js/lobibox.min.js'/>"></script>
<link rel="StyleSheet" type="text/css" href="./css/default.css" />
<title>nepes 주식정보</title>
<script type="text/javascript">

$(document).ready(function() {
	var stockInfo = "<%=Stockinfo[0]%>";

		if (stockInfo.length > 0) {
			setInterval(function() {
				window.location.reload(true);
			}, 30000);
		} else
			alert('페이지에 문제가 있습니다. 새로고침 후 확인바랍니다.');
	});
</script>
</head>
<body style="margin: 0px;">
	<form name="krx" method="post">
		<div style="width: 710px;">
			<div style="width: 700px; float: left; padding: 0 5 5 0;">
				<table width="100%" border="0" cellpadding="0" cellspacing="0"
					bgcolor="#CCCCCC">
					<caption class="dpn">현재가</caption>
					<tr>
						<td colspan="2" height="2" class="line_color"></td>
					</tr>
					<tr>
						<td width="45%" bgcolor="#FFFFFF" style="padding: 3 3 3 3">
							<table width="100%" border="0" cellpadding="0" cellspacing="0">
								<caption class="dpn">현재가, 전일대비 및 등락률(%)</caption>
								<tr>
									<td height="44" rowspan="2" align="center">
										<table width="100%" border="0" cellspacing="0" cellpadding="0">
											<caption class="dpn">현재가</caption>
											<tr>
												<th height="22" align="center" style="padding: 0 0 0 5"><img
													src="./img/wz_icon_coin.gif" align="absmidde" alt="현재가">&nbsp;<span
													style="font-family: 굴림, 굴림체, 돋움, 돋움체; font-size: 9pt; color: #000000;">현재가</span></th>
											</tr>
											<tr>
												<td align="center">
													<%
														if (Stockinfo[2].equals("1") || Stockinfo[2].equals("2")) {
													%> <span class="up"> <%=Bigup%>
												</span> <%
 	}
 %> <%
 	if (Stockinfo[2].equals("3")) {
 %> <span class="bohab"> <%=bohab%>
												</span> <%
 	}
 %> <%
 	if (Stockinfo[2].equals("4") || Stockinfo[2].equals("5")) {
 %> <span class="down"> <%=Bigdown%>
												</span> <%
 	}
 %> &nbsp;<span style="color: #000000;" class="updn15_1"><strong><%=Stockinfo[1]%></strong></span>
												</td>
											</tr>
										</table>
									</td>
									<td width="1" rowspan="2" bgcolor="#cccccc"></td>
									<th height="22" align="center"><span
										style="font-family: 굴림, 굴림체, 돋움, 돋움체; font-size: 9pt; color: #000000;">전일대비</span></th>
									<td width="1" rowspan="2" bgcolor="#cccccc"></td>
									<th align="center"><span
										style="font-family: 굴림, 굴림체, 돋움, 돋움체; font-size: 9pt; color: #000000;">등락률(%)</span></th>
								</tr>
								<tr>
									<td align="center">&nbsp; <%
 	if (Stockinfo[2].equals("1") || Stockinfo[2].equals("2")) {
 %> <span class="up"> <%=up%>
									</span> <%
 	}
 %> <%
 	if (Stockinfo[2].equals("3")) {
 %> <span class="bohab"> <%=bohab%>
									</span> <%
 	}
 %> <%
 	if (Stockinfo[2].equals("4") || Stockinfo[2].equals("5")) {
 %> <span class="down"> <%=down%>
									</span> <%
 	}
 %> &nbsp;<span class="updn09_1"><%=Stockinfo[3]%></span></td>
									<td align="center">&nbsp; <%
 	if (Stockinfo[2].equals("1") || Stockinfo[2].equals("2")) {
 %> <span class="up"> <%=up%>
									</span> <%
 	}
 %> <%
 	if (Stockinfo[2].equals("3")) {
 %> <span class="bohab"> <%=bohab%>
									</span> <%
 	}
 %> <%
 	if (Stockinfo[2].equals("4") || Stockinfo[2].equals("5")) {
 %> <span class="down"> <%=down%>
									</span> <%
 	}
 %> &nbsp;<span class="updn09_1"><%=DungRakrate_str%></span></td>
								</tr>
							</table>
						</td>
						<td width="55%" bgcolor="#FFFFFF" style="padding: 2 0 2 0">
							<table width="100%" border="0" cellspacing="0" cellpadding="0">
								<caption class="dpn">시가(원),고가(원),저가(원),개래량(주)</caption>
								<tr>
									<th height="24" align="center" class="hd_bgcolor"><span
										style="font-size: 9pt;" class="hd_cf">시가(원)</span></th>
									<td width="1" rowspan="2"></td>
									<th align="center" class="hd_bgcolor"><span
										style="font-size: 9pt;" class="hd_cf">고가(원)</span></th>
									<td width="1" rowspan="2"></td>
									<th align="center" class="hd_bgcolor"><span
										style="font-size: 9pt;" class="hd_cf">저가(원)</span></th>
									<td width="1" rowspan="2"></td>
									<th align="center" class="hd_bgcolor"><span
										style="font-size: 9pt;" class="hd_cf">거래량(주)</span></th>
								</tr>
								<tr>
									<td height="24" align="center"><span
										style="color: #000000; font-size: 9pt;" class="bd_f"><%=Stockinfo[7]%></span></td>
									<td align="center"><span
										style="color: #ff3c00; font-size: 9pt;" class="bd_f"><%=Stockinfo[8]%></span></td>
									<td align="center"><span
										style="color: #006ECC; font-size: 9pt;" class="bd_f"><%=Stockinfo[9]%></span></td>
									<td align="center"><span
										style="color: #000000; font-size: 9pt;" class="bd_f"><%=Stockinfo[5]%></span></td>
								</tr>
							</table>
						</td>
					</tr>
					<tr class="8pt">
						<td height="1" colspan="2" class="line_color"></td>
					</tr>
				</table>
			</div>
			<div style="width: 700px; float: left; padding: 0 5 5 0;">
				<table width="100%" border="0" cellspacing="0" cellpadding="0">
					<caption class="dpn">일자별시세 상세내용</caption>
					<tr>
						<td height="2" class="line_color"></td>
					</tr>
					<tr>
						<td>
							<table id="stg_byddsprcs_01" cellpadding="0" cellspacing="0"
								border="0" width="100%">
								<caption class="dpn">일자별시세 상세내용</caption>
								<thead>
									<tr>
										<th align="center" class="hd_bgcolor"><span
											class="hd_cfs">일자</span></th>
										<th align="center" class="hd_bgcolor"><span
											class="hd_cfs">종가</span></th>
										<th align="center" class="hd_bgcolor"><span
											class="hd_cfs">전일대비</span></th>
										<th align="center" class="hd_bgcolor"><span
											class="hd_cfs">시가</span></th>
										<th align="center" class="hd_bgcolor"><span
											class="hd_cfs">고가</span></th>
										<th align="center" class="hd_bgcolor"><span
											class="hd_cfs">저가</span></th>
										<th align="center" class="hd_bgcolor"><span
											class="hd_cfs">거래량</span></th>
										<th align="center" class="hd_bgcolor"><span
											class="hd_cfs">거래대금</span></th>
									</tr>
								</thead>
								<tfoot>
								</tfoot>
								<tbody id="contentByddsprcs01">
									<%
										if (dailystock_length > 0) {
									%>
									<%
										for (int j = 0; j < dailystock_length; j++) {
									%>
									<%
										if ((j % 2) == 0) {
									%>
									<tr>
										<td height="24" align="center" class="bd_bgcolor"><span
											class="bd_cfs"><%=Dailystock[j][0]%></span></td>
										<td height="24" align="center" class="bd_bgcolor"><span
											class="bd_cfs"><%=Dailystock[j][1]%></span></td>
										<td class="bd_bgcolor">
											<table width="100%" border="0" cellspacing="0"
												cellpadding="0">
												<tr>
													<%
														if (Dailystock[j][8].equals("1") || Dailystock[j][8].equals("2")) {
													%>
													<td width="20%" align="right"><%=up%> <%
 	}
 %></td>
													<%
														if (Dailystock[j][8].equals("3")) {
													%>
													<td width="20%" align="right"><%=bohab%> <%
 	}
 %></td>
													<%
														if (Dailystock[j][8].equals("4") || Dailystock[j][8].equals("5")) {
													%>
													<td width="20%" align="right"><%=down%> <%
 	}
 %></td>
													<td align="right"><span class="bd_cfs"><%=Dailystock[j][2]%></span>
													</td>
												</tr>
											</table>
										</td>
										<td height="24" align="center" class="bd_bgcolor"><span
											class="bd_cfs"><%=Dailystock[j][6]%></span></td>
										<td height="24" align="center" class="bd_bgcolor"><span
											class="bd_cfs"><%=Dailystock[j][4]%></span></td>
										<td height="24" align="center" class="bd_bgcolor"><span
											class="bd_cfs"><%=Dailystock[j][5]%></span></td>
										<td height="24" align="center" class="bd_bgcolor"><span
											class="bd_cfs"><%=Dailystock[j][7]%></span></td>
										<td height="24" align="center" class="bd_bgcolor"><span
											class="bd_cfs"><%=Dailystock[j][3]%></span></td>
									</tr>
									<%
										} else {
									%>
									<tr>
										<td height="24" align="center" bgcolor="#F5F5F5"><span
											class="bd_cfs"><%=Dailystock[j][0]%></span></td>
										<td height="24" align="center" bgcolor="#F5F5F5"><span
											class="bd_cfs"><%=Dailystock[j][1]%></span></td>
										<td bgcolor="#F5F5F5">
											<table width="100%" border="0" cellspacing="0"
												cellpadding="0">
												<tr>
													<%
														if (Dailystock[j][8].equals("1") || Dailystock[j][8].equals("2")) {
													%>
													<td width="20%" align="right"><%=up%> <%
 	}
 %></td>
													<%
														if (Dailystock[j][8].equals("3")) {
													%>
													<td width="20%" align="right"><%=bohab%> <%
 	}
 %></td>
													<%
														if (Dailystock[j][8].equals("4") || Dailystock[j][8].equals("5")) {
													%>
													<td width="20%" align="right"><%=down%> <%
 	}
 %></td>
													<td align="right"><span class="bd_cfs"><%=Dailystock[j][2]%></span>
													</td>
												</tr>
											</table>
										</td>
										<td height="24" align="center" bgcolor="#F5F5F5"><span
											class="bd_cfs"><%=Dailystock[j][6]%></span></td>
										<td height="24" align="center" bgcolor="#F5F5F5"><span
											class="bd_cfs"><%=Dailystock[j][4]%></span></td>
										<td height="24" align="center" bgcolor="#F5F5F5"><span
											class="bd_cfs"><%=Dailystock[j][5]%></span></td>
										<td height="24" align="center" bgcolor="#F5F5F5"><span
											class="bd_cfs"><%=Dailystock[j][7]%></span></td>
										<td height="24" align="center" bgcolor="#F5F5F5"><span
											class="bd_cfs"><%=Dailystock[j][3]%></span></td>
									</tr>
									<%
										}
									%>
									<%
										}
									%>
									<%
										} else {
									%>
									<tr>
										<td colspan="8">데이터가 없습니다.</td>
									</tr>
									<%
										}
									%>
								</tbody>
							</table>
						</td>
					</tr>
					<tr>
						<td height="1" class="line_color"></td>
					</tr>
					<tr>
						<td height="5" valign="bottom"></td>
					<tr>
						<td align="center">
							<table cellspacing="0" cellpadding="0" border="0">
								<tr>
									<td height="10"></td>
								</tr>
								<tr>
									<td><img src="./img/wz_ment_1.gif"
										alt="본 정보는 오류가 발생하거나 지연 될 수 있습니다. 제공된 정보에 의한 투자결과에 대한 법적인 책임을 지지 않습니다." /></td>
								</tr>

							</table>
						</td>
					</tr>
					<tr>
<!--						<td align="right" height="30"><a
								href="http://dart.fss.or.kr/html/search/SearchCompany_M2.html?textCrpNM=033640"
								style="color: Blue;" target="_blank">더 많은
									정보 보기 ></a></td> -->
					</tr> 
				</table>
			</div>
		</div>
	</form>
</body>
</html>