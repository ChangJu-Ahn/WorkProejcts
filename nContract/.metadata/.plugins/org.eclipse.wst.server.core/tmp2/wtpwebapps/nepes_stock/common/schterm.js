/**
 * Function : 湲곗��쇰줈 遺��� 議곌굔 ���낆뿉 �곕Ⅸ �좎쭨瑜� 怨꾩궛�섏뿬 諛섑솚�쒕떎.
 * @param   : baseDD   - 湲곗��쇱옄
 *          : type     - 怨꾩궛����(1:1��, 2:1二쇱씪, 3:1媛쒖썡, 4:3媛쒖썡, 5:6媛쒖썡, 6:1��, 7:2��, 8:3��)
 *          : frDDFdNm - Target Field Name
 *          : toDDFdNm - Source Field Name
 *          : formNm   - Form Name
 * @return  : String   - 怨꾩궛�� �좎쭨
 */
function calcDate(baseDD, type, frDDFdNm, toDDFdNm, formNm) {
    var relativeDD = null;
    
    try {
        if(baseDD == null || trimmed(baseDD).length == 0) {
            baseDD = today();
        }
        if(frDDFdNm == null) {
            frDDFdNm = "FR_TRD_DD";
        }
        if(toDDFdNm == null) {
            toDDFdNm = "TO_TRD_DD";
        }
        
        if(type == 1) { relativeDD = baseDD; }                                 // 1�� ��
        else if ( type == 2 ) { relativeDD = relativeDate(baseDD, -7); }       // 1二쇱씪 ��
        else if ( type == 3 ) { relativeDD = relativeMonth(baseDD, -1); }      // 1媛쒖썡 ��
        else if ( type == 4 ) { relativeDD = relativeMonth(baseDD, -3); }      // 3媛쒖썡 ��
        else if ( type == 5 ) { relativeDD = relativeMonth(baseDD, -6); }      // 6媛쒖썡 ��
        else if ( type == 6 ) { relativeDD = relativeYear(baseDD, -1); }       // 1�� ��
        else if ( type == 7 ) { relativeDD = relativeYear(baseDD, -2); }       // 2�� ��
        else if ( type == 8 ) { relativeDD = relativeYear(baseDD, -3); }       // 3�� ��
		else if ( type == 0) { relativeDD = "20040102"; }                      // �꾩껜
        else { throw "�좎쭨 怨꾩궛 ���� �몄닔媛� 遺��곸젅�⑸땲��.(" + type + ")"; }    // �덉쇅諛쒖깮
        
        if(formNm == null) {
            document.all[frDDFdNm].value = relativeDD;
            document.all[toDDFdNm].value = baseDD;
        }
        else {
            document.forms[formNm][frDDFdNm].value = relativeDD;
            document.forms[formNm][toDDFdNm].value = baseDD;
        }
    }
    catch(e) {
        dialogError(e);
    }
}

/**
 * Function : 臾몄옄�댁쓽 �욌뮘 space瑜� �쒓굅�쒕떎.
 * @param   : value
 * @return  : �욌뮘�� space媛� �쒓굅�� 臾몄옄��
 */
function trimmed(value) {
    value = value.replace(/^\s+/, "");     // remove leading white spaces
    value = value.replace(/\s+$/g, "");    // remove trailing while spaces
    
    return value;
}

/**
 * Function : �꾩옱�쇱쓣 諛섑솚�쒕떎.(�대씪�댁뼵�� �쒓컖)
 * @param   : delm   - 援щ텇��
 * @return  : String - �꾩옱�쇱옄
 */
function today(delm) {
    if(delm == null ) { delm = ""; }
    
    var now   = new Date();
    var year  = now.getFullYear();
    var month = now.getMonth() + 1;
    var date  = now.getDate();
    
    if(month < 10) { month = "0" + month; }
    if(date  < 10) { date  = "0" + date;  }
    
    return (year + delm + month + delm + date);
}

/**
 * Function : 湲곗��쇱쓽 �곷��곸씤 �쇱옄瑜� 怨꾩궛�� �좎쭨瑜� 援ы븳��.
 * @param   : baseDD - 湲곗���
 *          : n      - �곷��� �쇱옄��
 * @return  : String - 怨꾩궛�� �좎쭨
 */
function relativeDate(baseDD, n) {
    var oDestDate = null;
    
    if(typeof baseDD == "object") {    // 湲곗��� ���낆씠 Date 媛앹껜
        oDestDate = baseDD;
    }
    else {                             // 湲곗��� ���낆씠 String 媛앹껜
        oDestDate = castDateType(baseDD);
    }
    
    oDestDate.setDate(oDestDate.getDate() + n);
    
    return castStrType(oDestDate);
}

/**
 * Function : 湲곗��쇱쓽 �곷��곸씤 媛쒖썡瑜� 怨꾩궛�� �좎쭨瑜� 援ы븳��.
 * @param   : bastDD - 湲곗���
 *          : n      - �곷��� 媛쒖썡��
 * @return  : String - 怨꾩궛�� �쇱옄
 */
function relativeMonth(bastDD, n) {
    var oldDate, newLastDate;
    var oDestDate = null;
    
    if(typeof bastDD == "object") {    // 湲곗��� ���낆씠 Date 媛앹껜
        oDestDate = bastDD;
    }
    else {                             // 湲곗��� ���낆씠 String 媛앹껜
        oDestDate = castDateType(bastDD);
    }
    
    // �꾩옱 �쇱옄瑜� 諛깆뾽�� �먭퀬 1�쇰줈 �명똿�� �� �곷��곸씤 媛쒖썡 怨꾩궛�� �ㅼ떆 �꾩옱�쇱옄瑜� 蹂듭썝�쒕떎.
    // �댁쑀) 援ы븯�� �곷��쇱옄�� �꾩썡�� 留덉쭅留� �쇱옄媛� �꾩옱�� �쇱옄蹂대떎 �곸쓣 寃쎌슦,
    //       �곷��쇱옄�� 留덉�留됱씪�먮줈 �명똿�댁빞 ��
    //       ��> 20050731 �� �쒕떖�꾩쓣 �쇱옄�� 20050631(X) 媛� �꾨땲怨� 20050630 �대떎.
    oldDate = oDestDate.getDate();
    oDestDate.setDate(1);
    
    // �곷��곸씤 媛쒖썡 怨꾩궛
    oDestDate.setMonth(oDestDate.getMonth() + n);
    
    // �곷��곸씤 �꾩썡�� 留덉�留� �쇱옄�� �댁쟾 �꾩썡 �쇱옄瑜� 鍮꾧탳�� �� ���뱁븳 �쇱옄瑜� �명똿
    // ��> 20050731 �� �쒕떖�꾩쓣 �쇱옄�� 20050631(X) 媛� �꾨땲怨� 20050630 �대떎.
    var nTmp = oDestDate.getFullYear().toString();
    if(nTmp.length == 2) {
        nTmp = "19" + nTmp;
    }
    newLastDate = getDaysInMonth(eval(nTmp), oDestDate.getMonth() + 1);
    if(oldDate > newLastDate) {
        oDestDate.setDate(newLastDate);
    }
    else {
        oDestDate.setDate(oldDate);
    }
    
    if(bastDD.length == 6) {
        return castStrType(oDestDate, "yyyyMM");
    }
    else {
        return castStrType(oDestDate);
    }
}

/**
 * Function : 湲곗��쇱쓽 �곷��곸씤 �꾨룄瑜� 怨꾩궛�� �좎쭨瑜� 援ы븳��.
 * @param   : bastDD - 湲곗���
 *          : n      - �곷��� �꾨룄��
 * @return  : String - 怨꾩궛�� �쇱옄
 */
function relativeYear(bastDD, n) {
    return relativeMonth(bastDD, n * 12);
}

/**
 * Function : String �뺤떇�� Date �뺤떇�쇰줈 蹂���
 * @param   : strDate - String �뺤떇�� �좎쭨.
 * @return  : Date    - 蹂��섎맂 Date �뺤떇�� 媛앹껜
 */
function castDateType(strDate) {
    var dtRtn = null;
    
    if(strDate.length == 6) { strDate += "01"; }
    
    if(strDate.length == 10) {        // �щ㎎�� 媛�吏� �뺥깭濡� �꾨떖�섏뿀�� 寃쎌슦. (�� 2005.01.01)
        var aDate = strDate.split(strDate.substring(4, 5));
        dtRtn = new Date(aDate[0], eval(aDate[1]) - 1, aDate[2]);
    }
    else if(strDate.length == 8) {    // �щ㎎�� �녿뒗 �뺥깭濡� �꾨떖�섏뿀�� 寃쎌슦. (�� 20050101)
        var year  = eval(strDate.substring(0, 4));
        var month = eval(strDate.substring(4, 6));
        var date  = eval(strDate.substring(6, 8));
        
        dtRtn = new Date(year, month - 1, date);
    }
    else {
        throw "遺��곹빀�� �좎쭨 �뺤떇�낅땲��.(" + strDate + ")";
    }
    
    return dtRtn;
}

/**
 * Function : Date �뺤떇�� String �뺤떇�쇰줈 蹂���
 * @param   : dtDate - Date �뺤떇�� �좎쭨.
 *          : delm   - �좎쭨�� �щ㎎ 援щ텇�� (�� '.' -> 2005.08.01 )
 * @return  : String - 蹂��섎맂 String �뺤떇�� 媛앹껜
 */
function castStrType(dtDate, format) {
    var re = "";
    var delm = "";
    
    if(format == null)  { format = "yyyyMMdd"; }
    
    for(var i = 0; i < format.length; i++) {
        var vChr = format.charAt(i);
        if(vChr != 'y' && vChr != 'M' && vChr != 'd') { 
            delm = vChr;
            break;
        }
    }
    
    if(delm == "/") {
        re = eval("/\\" + delm + "/g");
    }
    else if(delm.length != 0) {
        re = eval("/" + delm + "/g");
    }
    
    format = format.replace(re, "");
    
    var year  = dtDate.getFullYear().toString().length == 2 ? "19" + dtDate.getFullYear() : dtDate.getFullYear();
    var month = dtDate.getMonth()+1;
    var date  = dtDate.getDate();
    
    if(month < 10) { month = "0" + month; }
    if(date  < 10) { date  = "0" + date;  }
    
    if(format == "yyyy") { return year; }
    else if (format == "yyyyMM") { return year + delm + month; }
    else if (format == "yyyyMMdd") { return year + delm + month + delm + date; }
}

/**
 * Function : �대떦�붿씠 紐뉗씪源뚯� �덈뒗吏� 怨꾩궛�쒕떎.
 *          : 13�� 15�� 異붽�.
 * @param   : year  - �꾨룄
 *          : month - ��
 * @return  : days  - �쇱닔
 */
function getDaysInMonth(year, month) {
    var days;
    if(month == 1 || month == 3 || month == 5 || month == 7 || month == 8 || month == 10 || month == 12) {
        days = 31;
    }
    else if(month == 4 || month == 6 || month == 9 || month == 11) {
        days = 30;
    }
    else if(month == 2) {
        if(leapYear(year) == 1) {
            days = 29;
        }
        else {
            days = 28;
        }
    }
    else if(month == 13) {
        days = 15;
    }
    
    return (days);
}

/**
 * Function : �대떦�꾩씠 �ㅻ뀈�몄� 寃��ы븳��.
 * @param   : year  - �꾨룄
 *          : month - ��
 * @return  : �ㅻ뀈�대㈃ 1, �꾨땲硫� 0
 */
function leapYear(Year) {
    if(((Year % 4) == 0) && ((Year % 100) != 0) || ((Year % 400) == 0))
        return (1);
    else
        return (0);
}

/**
 * Function : �щ젰瑜� �쒖떆�쒕떎.
 * @param   : oTgtFld   - �쇱옄媛� �명똿�� �낅젰 �꾨뱶 媛앹껜
 * @param   : vCategory - 遺꾨쪟 (* MS : �좉�利앷텒, �좊Ъ, 梨꾧텒 * MQ : 肄붿뒪��)
 * @param   : vSelMode  - [A]:�꾩껜�쇱옄, [W]:�됱씪留�, [P]:湲덉씪�쒖쇅 �μ슫�곸씪,
 *                        [0]:��, [1]:��, [2]:��, [3]:��, [4]:紐�, [5]:湲�, [6]:��, [-1]:�꾩껜
 *                        < 湲곕낯媛� : -1 (湲덉씪�ы븿 �μ슫�곸씪) >
 *                        [Y]:�꾨쭔, [M]:�붾쭔, (�꾩옱 �곸슜�섏� �딆� �듭뀡�낅땲�� /writer. 諛곗꽦��)
 * @return  : none
 */
function showCalendar(oTgtFld, vCategory, vSelMode) {
    if(!oTgtFld) {
        dialogError("�좏슚�섏� �딅뒗 �좎쭨 �낅젰 �꾨뱶�낅땲��.");
        return;
    }
    /*
    if(!vCategory || (vCategory != "MS" && vCategory != "MQ")) {
        dialogError("二쇱떇 遺꾨쪟 �몄닔 �꾨떖 �ㅻ쪟! (�щ젰 而댄룷�뚰듃)\n��) * MS(�좉�利앷텒, �좊Ъ, 梨꾧텒) * MQ(肄붿뒪��)");
        return;
    }
    */
    // �щ젰 �덉씠�� �쒓렇 �앹꽦.
    createCalLyrTag(170, 190);
    
    var oCalLyr = document.getElementById("calendarLyr");    // �щ젰 蹂댁씠湲�/�④린湲� 而⑦듃濡� �덉씠�� 媛앹껜
    var oCalIfm = document.getElementById("calendarIfm");    // �щ젰 JSP媛� 留곹겕�� iframe 媛앹껜
    
    // �쇱옄媛� �명똿�� input �쒓렇 value 媛� (�щ젰 �쒖떆�� 湲곕낯�꾩썡�� input �쒓렇�� �덊똿�� �꾩썡濡� �쒖떆�쒕떎.)
    var vYymm = oTgtFld.value;
    
    //if(!vYymm || vYymm.length < 6) { vYymm = ""; }    // 20061221 二쇱꽍泥섎━
    if(!vYymm || vYymm.length < 6) { vYymm = "0"; }    // 20061221 援먯껜
    else if(vYymm.length > 6) { vYymm = vYymm.substring(0, 6); }
    
    // �좏깮 湲곕낯媛믪쓣 �꾩껜濡� �쒕떎.
    if(!vSelMode || trimmed(vSelMode) == "") { vSelMode = "-1"; }
    
    var param = vYymm + ";" + vCategory + ";" + vSelMode + ";" + oTgtFld.name;
    oCalIfm.src = "/calendar/calendar.jsp?param=" + param;    // 20061221 援먯껜
    //"?yymm=" + vYymm +
    //"&category=" + vCategory +
    //"&sel_mode=" + vSelMode +
    //"&tgt_fld_nm=" + oTgtFld.name;    // 20061221 二쇱꽍泥섎━
    
    var x = event.clientX - event.offsetX + document.body.scrollLeft - 5;
    var y = event.clientY - event.offsetY + document.body.scrollTop + 19;
    
    oCalLyr.style.left = x;
    oCalLyr.style.top = y;    //oCalLyr.style.pixelTop = y;
    oCalLyr.style.visibility = "visible";
    
}

/**
 * Function : �щ젰 �덉씠�� �쒓렇瑜� �앹꽦�쒕떎.
 * @return  : none
 */
function createCalLyrTag(nW, nH) {
    // �대� �앹꽦�섏뼱 �덉쑝硫� 由ы꽩
    if(document.getElementById("calendarLyr")) { return; }
    
    if(!nW) { nW = 212; }
    if(!nH) { nH = 162; }
    
    // 釉뚮씪�곗� 泥댄겕
    var browser = chkBrowser();
    
    // �듭뒪 6,7,8 踰꾩졏�대㈃ 湲곗〈 �뚯뒪 洹몃�濡�
    // (湲곗〈�뚯뒪�� �듭뒪 9踰꾩졏 �댁긽怨� �щ＼�깆뿉�� �곸슜�섏� �덈뒗 臾몄젣�먯씠 �덉쓬)
    if ((browser.ie6) || (browser.ie7) || (browser.ie8)) {
    	
    	
        // �щ젰 �덉씠�� �쒓렇
        var vLyrElement = '<div id="calendarLyr" style="visibility:hidden; position:absolute;"></div>';
        // �щ젰 iframe �쒓렇
        var vIfmElement = '<iframe id="calendarIfm" width="' + nW + '" height="' + nH + '" scrolling="no" frameborder="0"></iframe>';
        
        // 1. body �섏쐞�� �щ젰 �덉씠�� �쒓렇 �쎌엯
        var oCalLyr = document.createElement(vLyrElement);
        document.body.insertAdjacentElement("afterBegin", oCalLyr);
        
        // 2. �щ젰 �덉씠�� �쒓렇 �섏쐞�� �щ젰 iframe �쒓렇 �쎌엯
        var oCalIfm = document.createElement(vIfmElement);
        oCalLyr.insertAdjacentElement("afterBegin", oCalIfm);
        
    	
    } else {
    // 洹몄쇅 踰꾩졏�대㈃ �덈줈�� �뚯뒪�곸슜
    	
	    var oCalLyr = document.createElement("div");
	    oCalLyr.setAttribute("id", "calendarLyr");
	    oCalLyr.setAttribute("style", "position:absolute; visibility:hidden;");
	    oCalLyr.style.position = "absolute";
	    oCalLyr.style.visibility = "hidden";
	    document.body.insertAdjacentElement("afterBegin", oCalLyr);
	    
	    var oCalIfm = document.createElement("iframe");
	    oCalIfm.setAttribute("id", "calendarIfm");
	    oCalIfm.setAttribute("width", nW);
	    oCalIfm.setAttribute("height", nH);
	    oCalIfm.setAttribute("scrolling", "no");
	    oCalIfm.setAttribute("frameborder", "0");
	    oCalLyr.style.border = "0";
	    oCalLyr.insertAdjacentElement("afterBegin", oCalIfm);
    }

}

/**
 * Function : 釉뚮씪�곗� 踰꾩졏�� �뚯븘�몃떎.
 * @return  : Browser  - 釉뚮씪�곗졇蹂� �쇱튂�щ�
 */
function chkBrowser() {
	var Browser = {
		chk : navigator.userAgent.toLowerCase()
	}
	
	Browser = {
		ie : Browser.chk.indexOf('msie') != -1,
		ie6 : Browser.chk.indexOf('msie 6') != -1,
		ie7 : Browser.chk.indexOf('msie 7') != -1,
		ie8 : Browser.chk.indexOf('msie 8') != -1,
		ie9 : Browser.chk.indexOf('msie 9') != -1,
		ie10 : Browser.chk.indexOf('msie 10') != -1,
		opera : !!window.opera,
		safari : Browser.chk.indexOf('safari') != -1,
		safari3 : Browser.chk.indexOf('applewebkir/5') != -1,
		mac : Browser.chk.indexOf('mac') != -1,
		chrome : Browser.chk.indexOf('chrome') != -1,
		firefox : Browser.chk.indexOf('firefox') != -1
	}
	return Browser;
}