function loading() {
    var ob = '';
    ob = '<object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" id="loading" width="120" height="72" codebase="http://fpdownload.macromedia.com/get/flashplayer/current/swflash.cab">'
        + '  <param name="movie" value="/images/loading.swf" />'
        + '  <param name="quality" value="high" />'
        + '  <param name="wmode" value="transparent" />'
        + '  <param name="bgcolor" value="#FFFFFF" />'
        + '  <param name="allowScriptAccess" value="sameDomain" />'
        + '  <embed src="/images/loading.swf" width="120" height="72" name="loading" align="middle" play="true" loop="false" quality="high" allowScriptAccess="sameDomain" type="application/x-shockwave-flash" pluginspage="http://www.adobe.com/go/getflashplayer"></embed>'
        + '</object>	';
    document.write(ob);
}

/**
 * dataCount - 그리드의 전체 크기
 * pageSize  - 한 페이지의 크기
 * pageNum   - 현재 페이지 번호
 */
function getPageNav(dataCount, pageSize, pageNum, gubun) {
    if(gubun == null) gubun = '';
    // 이미지 파일 정의
    var DEFAULT_FIRST_PAGE_BTN = "<img src=\"/images/wz_btn_before02.gif\" border=\"0\" align=\"absmiddle\" alt=\"처음으로\">" ;
    var DEFAULT_PRE_PAGE_BTN   = "<img src=\"/images/wz_btn_before01.gif\" border=\"0\" align=\"absmiddle\" alt=\"이전페이지\">" ;
    var DEFAULT_NEXT_PAGE_BTN  = "<img src=\"/images/wz_btn_next01.gif\" border=\"0\" align=\"absmiddle\" alt=\"다음페이지\">" ;
    var DEFAULT_LAST_PAGE_BTN  = "<img src=\"/images/wz_btn_next02.gif\" border=\"0\" align=\"absmiddle\" alt=\"마지막으로\">" ;
    var HID_FIRST_PAGE_BTN = "<img src=\"/images/wz_btn_before04.gif\" border=\"0\" align=\"absmiddle\" alt=\"처음으로(비활성)\" >" ;
    var HID_PRE_PAGE_BTN   = "<img src=\"/images/wz_btn_before03.gif\" border=\"0\" align=\"absmiddle\" alt=\"이전페이지(비활성)\" >" ;
    var HID_NEXT_PAGE_BTN  = "<img src=\"/images/wz_btn_next03.gif\" border=\"0\" align=\"absmiddle\" alt=\"다음페이지(비활성)\" >" ;
    var HID_LAST_PAGE_BTN  = "<img src=\"/images/wz_btn_next04.gif\" border=\"0\" align=\"absmiddle\" alt=\"마지막으로(비활성)\" >" ;
    
    var nav = "";
    var total_count = dataCount - 1;
    var total_page  = Math.ceil(total_count / pageSize);
    
    if(total_page <= 0){
        total_page = 1;
    }
    
    var endPage = total_page;
    var fromPage = (Math.ceil(pageNum / 10)-1) * 10 + 1;
    var toPage = fromPage + 10 - 1;
    
    if(toPage >= total_page) {
        toPage = total_page;
    }
    
    if(pageNum > 1) {
        nav += "<a href=\"javascript:searchData"+gubun+"(1)\" id=\"botton_1\">";
        nav += DEFAULT_FIRST_PAGE_BTN;
        nav += "</a>  ";
        nav += "<a href=\"javascript:searchData"+gubun+"(" + (pageNum - 1) + ")\" id=\"botton_" + (pageNum - 1) + "\">";
        nav += DEFAULT_PRE_PAGE_BTN;
        nav += "</a>";
    }
    else {
        nav += HID_FIRST_PAGE_BTN;
        nav += "&nbsp;";
        nav += HID_PRE_PAGE_BTN;
    }
    
    nav += "&nbsp;&nbsp;&nbsp;";
    
    for(var i = (fromPage - 1); i < toPage; i++) {
        if((i + 1) != pageNum) {
            nav += "<a href=\"javascript:searchData"+gubun+"(" + (i + 1) + ")\" id=\"botton_"+ (i + 1) +"\">";
            nav += i + 1;
            nav += "</a>";
        }
        else {
            nav += "<b>" + (i + 1) + "</b>";
        }
        if((i + 1) < toPage) {
            nav += " | ";
        }
    }
    
    nav += "&nbsp;&nbsp;&nbsp;";
    
    if(pageNum < endPage) {
        nav += "<a href=\"javascript:searchData"+gubun+"(" + (pageNum + 1) + ")\" id=\"botton_" + (pageNum + 1) + "\">";
        nav += DEFAULT_NEXT_PAGE_BTN;
        nav += "</a>  ";
        nav += "<a href=\"javascript:searchData"+gubun+"(" + endPage + ")\" id=\"botton_" + endPage + "\">";
        nav += DEFAULT_LAST_PAGE_BTN;
        nav += "</a>&nbsp;";
    }
    else {
        nav += HID_NEXT_PAGE_BTN;
        nav += "&nbsp;";
        nav += HID_LAST_PAGE_BTN;
    }
    
    $('bottonObj'+ gubun).innerHTML = nav;
}

function getPageNavTp1(dataCount, pageSize, pageNum, gubun, tabidx) {
    if(gubun == null) gubun = '';
    // 이미지 파일 정의
    var DEFAULT_PRE_PAGE_BTN   = "<img src=\"/images/mob/board_btn_pre.gif\" border=\"0\" align=\"absmiddle\" alt=\"이전페이지\">" ;
    var DEFAULT_NEXT_PAGE_BTN  = "<img src=\"/images/mob/board_btn_next.gif\" border=\"0\" align=\"absmiddle\" alt=\"다음페이지\">" ;
    
    var nav = "";
    var total_count = dataCount - 1;
    var total_page  = Math.ceil(total_count / pageSize);
    
    if(total_page <= 0){
        total_page = 1;
    }
    
    var endPage = total_page;
    var fromPage = (Math.ceil(pageNum / 3)-1) * 3 + 1;
    var toPage = fromPage + 3 - 1;
    
    if(toPage >= total_page) {
        toPage = total_page;
    }
    
    nav += "<p class=\"pg1\">";
    
    if(pageNum > 1) {
        nav += "<a href=\"javascript:searchData"+gubun+"(" + (pageNum - 1) + ",'"+ tabidx  + "')\" id=\"botton_" + (pageNum - 1) + "\" class=\"bt4 bt4pv\">";
        nav += DEFAULT_PRE_PAGE_BTN;
        nav += "</a>";
    }
    else {
        nav += "";
    }
    
    for(var i = (fromPage - 1); i < toPage; i++) {
        if((i + 1) != pageNum) {
            nav += "<a href=\"javascript:searchData"+gubun+"(" + (i + 1) + ",'"+ tabidx  + "')\" id=\"botton_"+ (i + 1) +"\">";
            nav += i + 1;
            nav += "</a>";
        }
        else {
            nav += "<a class=\"on\">" + (i + 1) + "</a>";
        }
    }
    
    if(pageNum < endPage) {
        nav += "<a href=\"javascript:searchData"+gubun+"(" + (pageNum + 1) + ",'"+ tabidx  + "')\" id=\"botton_" + (pageNum + 1) + "\" class=\"bt4 bt4nx\">";
        nav += DEFAULT_NEXT_PAGE_BTN;
        nav += "</a>  ";
    }
    else {
        nav += "";
    }
    nav += "</p>";
    
    $('bottonObj'+ gubun).innerHTML = nav;
}

function getPageNavTp2(dataCount, pageSize, pageNum, gubun) {
    if(gubun == null) gubun = '';
    // 이미지 파일 정의
    var DEFAULT_PRE_PAGE_BTN   = "<img src=\"/images/mob/board_btn_pre.gif\" border=\"0\" align=\"absmiddle\" alt=\"이전페이지\">" ;
    var DEFAULT_NEXT_PAGE_BTN  = "<img src=\"/images/mob/board_btn_next.gif\" border=\"0\" align=\"absmiddle\" alt=\"다음페이지\">" ;
    
    var nav = "";
    var total_count = dataCount - 1;
    var total_page  = Math.ceil(total_count / pageSize);
    
    if(total_page <= 0){
        total_page = 1;
    }
    
    var endPage = total_page;
    var fromPage = (Math.ceil(pageNum / 3)-1) * 3 + 1;
    var toPage = fromPage + 3 - 1;
    
    if(toPage >= total_page) {
        toPage = total_page;
    }
    
    nav += "<p class=\"pg1\">";
    
    if(pageNum > 1) {
        nav += "<a href=\"javascript:searchData"+gubun+"(" + (pageNum - 1) + ")\" id=\"botton_" + (pageNum - 1) + "\" class=\"bt4 bt4pv\">";
        nav += DEFAULT_PRE_PAGE_BTN;
        nav += "</a>";
    }
    else {
        nav += "";
    }
    
    for(var i = (fromPage - 1); i < toPage; i++) {
        if((i + 1) != pageNum) {
            nav += "<a href=\"javascript:searchData"+gubun+"(" + (i + 1) + ")\" id=\"botton_"+ (i + 1) +"\">";
            nav += i+1;
            nav += "</a>";
        }
        else {
            nav += "<a class=\"on\">" + (i + 1) + "</a>";
        }
    }
    
    
    if(pageNum < endPage) {
        nav += "<a href=\"javascript:searchData"+gubun+"(" + (pageNum + 1) + ")\" id=\"botton_" + (pageNum + 1) + "\" class=\"bt4 bt4nx\">";
        nav += DEFAULT_NEXT_PAGE_BTN;
        nav += "</a>  ";
    }
    else {
        nav += "";
    }
    nav += "</p>";
    
    $('bottonObj'+ gubun).innerHTML = nav;
}

function goSubmit(gubun) {
    if(gubun == null) gubun = '';
    
    if(gubun != '') {
        if($F('fr_' + gubun) > $F('to_' + gubun)) {
            alert("검색시작일이 검색종료일 보다 이전일 수 없습니다.");
            return;
        }
    }
    
    var f = document.krx;
    f.action = location.href;
    f.method = "post";
    f.target = "_self";
    
    f.submit();
}