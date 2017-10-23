/// <reference path="multicolumncombobox.js" />

//
//	  Copyright (C) 2003-2010 GrapeCity Inc.	All rights reserved.
//

//>>#99915912 - Jeff 2010/04/30
if (!Function.prototype.CreateDelegate)
  Function.prototype.CreateDelegate = function(instance, method) { return function() { return method.apply(instance, arguments); } }
//<<#99915912 - Jeff 2010/04/30
if (typeof (FarPoint) == "undefined")
  FarPoint = {}
if (typeof (FarPoint.Web) == "undefined")
  FarPoint.Web = {}
if (typeof (FarPoint.Web.Spread) == "undefined")
  FarPoint.Web.Spread = {}

FarPoint.Web.Spread.SpreadChart = function(chartSpread, container) {
  this.chartSpread = chartSpread;
  this.container = container;
  this.eventHandlers = {};
  this.tableChartData = null;
  this.chartID = null;
  this.chart = null;
  this.isActived = null;
};

// Methods.
FarPoint.Web.Spread.SpreadChart.prototype.init = function () {
  this.xmlData = typeof(this.chartSpread.getXmlDataObject)=="function"? this.chartSpread.getXmlDataObject():null; //lance - 2011/03/11 - IE9 Support
  if (this.xmlData == null) {
    this.tableChartData = document.getElementById(this.chartSpread.id + "_XMLDATA");
  }
  this.chartID = this.container.getAttribute("serverID");
  this.chart = this.getChart();
  
  var instance = this;
  this.container.getChartObj = function() { return instance;}
  
  this.chartStyleInfo = new FarPoint.Web.Spread.SpreadChartStyleInfo(this); //99915104
  this.registerChartItemEvent();
  //>>99917209 Justin 2010/09/10
  var node = this.getchartInfoByID(this.chartID);
  //lance - 2010/12/07 - #99918682
  //var left = parseInt(this.container.style.left);
  //var top = parseInt(this.container.style.top);
  //var width = parseInt(this.container.style.width);
  //var height = parseInt(this.container.style.height);
  //this.setChartLocation(node, left, top);
  var location = this.getChartLocation();
  this.setChartLocation(node, location.x, location.y);
  var chart = $$(this.chart);
  var width = chart.innerWidth();
  var height = chart.innerHeight();
  this.setChartSize(node, width, height);
  //<<99917209
  this.isActived = this.container.getAttribute("isActive") == "true";
  if (this.isActived) {
    this.activeChart();
  }

  if (this.chartSpread.virtualPaging == "true" && this.chartSpread.virtualTop > 0 && this.virtualUpdated != true) {
    this.container.style.top = parseInt(this.container.style.top) + this.chartSpread.virtualTop + "px";
    this.virtualUpdated = true;
  }

  if (this.chartSpread.preventScrollView) {
    this.container.removeAttribute('tabindex');
  }
};

FarPoint.Web.Spread.SpreadChart.prototype.getIsDrag = function(){
  return this.container._isMoved;
}

//lance - 2010/12/07 - #99918682
FarPoint.Web.Spread.SpreadChart.prototype.getChartLocation = function () {
  ///<summary>Gets the chart's location</summary>
  var location = $$(this.chart).clientLocation();
  location.x += parseInt(this.container.style.left);
  location.y += parseInt(this.container.style.top);
  if (this.isActived && !this.hasCustomHighlightCss()) {
    var container = $$(this.container);
    location.x += container.border("left");
    location.y += container.border("top");
  }
  return location;
}

FarPoint.Web.Spread.SpreadChart.prototype.addChartInfo = function () {
  var node = this.getchartInfoByID(this.chartID);
  if (node != null)
    return node;

  var root = this.getChartInfo();

  if (document.all != null) {
    if (this.xmlData != null) {  //lance - 2011/03/22 - 99920694
      var xmlData = this.xmlData;
      node = xmlData.createElement("chart");
      node.setAttribute("id", this.chartID);
    }
    else {
      node = this.tableChartData.createNode("element", "chart", "");
      attr = this.tableChartData.createNode("attribute", "id", "");
      attr.text = this.chartID;
      node.attributes.setNamedItem(attr);
    }
  }
  else {
    node = document.createElement("chart");
    node.setAttribute("id", this.chartID);
  }
  root.appendChild(node);

  return node;
};

FarPoint.Web.Spread.SpreadChart.prototype.getchartInfoByID = function () {
  var root = this.getChartInfo();
  var node;
  if (FpChartUtil.isIE()) {
    if (this.xmlData != null) {  //lance - 2011/03/22 - 99920694
      var xmlData = this.xmlData;
      xmlData.switchNodeContext(root);
      var node = xmlData.selectSingleNode("chart[id='" + this.chartID + "']");
      xmlData.resetNodeContext();
    }
    else {
      node = root.selectSingleNode("./chart[@id='" + this.chartID + "']");
    }
  }
  else {
    node = the_fpSpread.GetElementById(root, this.chartID);
  }
  return node;
};

FarPoint.Web.Spread.SpreadChart.prototype.getChart = function() {
  var chartImage = null;
  var items = this.container.getElementsByTagName("img");
  for (var i = 0; i < items.length; i++) {
    if (items[i].build == null) {
      chartImage = items[i];
      break;
    }
  }
  return chartImage;
};

FarPoint.Web.Spread.SpreadChart.prototype.getChartInfo = function () {

  var xmlData = this.xmlData;
  if (xmlData != null) {
    return xmlData.selectSingleNode("chartinfo");
  }
  else {
    if (typeof this.tableChartData.documentElement == "undefined") {
      return this.tableChartData.getElementsByTagName("chartinfo")[0];
    }
    else {
      return this.tableChartData.documentElement.selectSingleNode("//chartinfo");
    }
  }
};

FarPoint.Web.Spread.SpreadChart.prototype.getEventTarget = function(event) {
  if (event.target == document && event.currentTarget != null)
    return event.currentTarget;
  if (event.target != null)
    return event.target;
  return event.srcElement;
};

FarPoint.Web.Spread.SpreadChart.prototype.registerChartItemEvent = function() {
  this.attachEvent(this.container, "mousedown", this.containerMousedown, this, null);
  if (this.container.getAttribute("moveChart") != "false") {
    this.attachEvent(document, "mouseup", this.endDragChart, this, null);
    this.attachEvent(this.container, "mouseup", this.endDragChart, this, null);
    this.attachEvent(this.container, "keydown", this.chartKeyDown, this, null);
  }

  if (this.container.getAttribute("sizeChart") != "false") {
    this.buildResizeCorner();
  }
};

FarPoint.Web.Spread.SpreadChart.prototype.activeChart = function () {
  if (this.container.getAttribute("selectChart") != "false") {
    var root = this.getChartInfo();
    var attr = root.attributes.getNamedItem("activechart");
    var browser = $$.browser;
    if (browser.ie && this.xmlData == null) {  //lance - 2011/03/22 - 99920694
      attr = this.tableChartData.createNode("attribute", "activechart", "");
      attr.text = this.chartID;
      root.attributes.setNamedItem(attr);
    }
    else {
      root.setAttribute("activechart", this.chartID);
    }
    // 99915941 CoreyKou 2010.05.06
    this.container.setAttribute("isActive", "true");

    if (FpChartUtil.isIE()) {
      this.chartSpread.SetActiveChartObj(this);
    }
    else {
      the_fpSpread.SetActiveChartObj(this.chartSpread, this);
    }

    this.highlightChart();
    var chartObj = this;
    if (!this.chartSpread.preventScrollView)
      setTimeout("chartObj.RetainFocus();", 100); //99917301 Justin 2010/07/29
  }
};
 //>>99917301 Justin 2010/07/29
FarPoint.Web.Spread.SpreadChart.prototype.RetainFocus = function () {
  var chartObj = this;
  chartObj.container.focus();
}
//<<99917301

FarPoint.Web.Spread.SpreadChart.prototype.inactiveChart = function() {
  var activeChart;
  if (FpChartUtil.isIE()) {
    activeChart = this.chartSpread.GetActiveChartObj();
  }
  else {
    activeChart = the_fpSpread.GetActiveChartObj(this.chartSpread);
  }

  if (activeChart != null) {
    activeChart.normalChart();
    var root = this.getChartInfo();
    root.removeAttribute("activechart");
    root.removeAttribute("tabindex");
  }
};

//>>99916057 Justin 2010/05/27
//should consider more case while scrolling chart into view
FarPoint.Web.Spread.SpreadChart.prototype.scrollChartIntoView = function () {
    if (this.chartSpread.getAttribute("ClientAutoSize") == "True") return; //#99917627 - Jeff 2010/08/27
    var scrollLeft, scrollTop;
    var parentNode = this.container.parentNode;
    this.stopVirtualPaging(); //#99915912 - Jeff 2010/04/29
    var container = this.container;
    //>>#99917502 - Jeff 2010/08/25
    var scrollTop = parentNode.scrollTop;
    var scrollLeft = parentNode.scrollLeft;
    var divScroll = document.getElementById(this.chartSpread.id + "_scrollvp");
    var offsetWidth = parentNode.clientWidth;
    var offsetHeight = parentNode.clientHeight;

    var vpSize = null;
    if (this.chartSpread.getViewportSize)
        vpSize = this.chartSpread.getViewportSize();

    if (vpSize != null && vpSize.height < parentNode.offsetHeight) {
        scrollTop = 0;
    }
    else {
        var topMost = container.offsetTop < parentNode.scrollTop;
        var bottomMost = container.offsetTop + container.offsetHeight + 5 > parentNode.scrollTop + parentNode.clientHeight;
        if (container.moveUp) {
            if (topMost)
                scrollTop = container.offsetTop;
            else if (bottomMost || container.offsetTop > this.chartSpread.getViewport().offsetTop + this.chartSpread.getViewport().offsetHeight)//#99917960 - Jeff 2010/09/08
                scrollTop = container.offsetTop + container.offsetHeight - offsetHeight + 5;
        }
        else {
            if (bottomMost)
                scrollTop = container.offsetTop + container.offsetHeight - offsetHeight + 5;
            else if (topMost)
                scrollTop = container.offsetTop;
        }
    }

    if (vpSize != null && vpSize.width < parentNode.clientWidth) {
        scrollLeft = 0;
    }
    else {
        var leftMost = container.offsetLeft < parentNode.scrollLeft;
        var rightMost = container.offsetLeft + container.offsetWidth + 5 > parentNode.scrollLeft + parentNode.clientWidth;
        if (container.moveLeft) {
            if (leftMost)
                scrollLeft = container.offsetLeft;
            else if (rightMost)
                scrollLeft = container.offsetLeft + container.offsetWidth - offsetWidth + 5;
        }
        else {
            if (rightMost)
                scrollLeft = container.offsetLeft + container.offsetWidth - offsetWidth + 5;
            else if (leftMost)
                scrollLeft = container.offsetLeft;
        }
    }
    if (scrollTop < 0) scrollTop = 0;
    if (scrollLeft < 0) scrollLeft = 0;
    if (!FpChartUtil.isIE()) {
        parentNode.scrollLeft = scrollLeft;
        parentNode.scrollTop = scrollTop;
        if (!this.isActived)
            container.focus();
    }
    else {
        if (!this.isActived)
            container.focus();

        if (divScroll != null) {
            var scrollHandle = divScroll.onscroll;
            if (divScroll.scrollLeft != scrollLeft && divScroll.scrollTop != scrollTop)
                divScroll.onscroll = null;

            parentNode.scrollLeft = scrollLeft;
            parentNode.scrollTop = scrollTop;

            divScroll.scrollLeft = scrollLeft;
            divScroll.scrollLeft = scrollLeft;
            divScroll.onscroll = scrollHandle;
            divScroll.scrollTop = scrollTop;
            divScroll.scrollTop = scrollTop;
        } else {
            if (scrollLeft != null) {
                parentNode.scrollLeft = scrollLeft;
            }
            if (scrollTop != null) {
                parentNode.scrollTop = scrollTop;
            }
        }
    }
    //<<#99917502 - Jeff 2010/08/25
    this.startVirtualPaging(); //#99915912 - Jeff 2010/04/29
};
//<<99916057

//>>#99915912 - Jeff 2010/04/29  
FarPoint.Web.Spread.SpreadChart.prototype.stopVirtualPaging = function() {
  this.chartSpread.stopVirtualPaging(this.chartSpread);
}

FarPoint.Web.Spread.SpreadChart.prototype.startVirtualPaging = function() {
  var _handler = Function.CreateDelegate(this, this.startVirtualPaging2);
  setTimeout(_handler, 100);
}

FarPoint.Web.Spread.SpreadChart.prototype.startVirtualPaging2 = function() {
  this.chartSpread.startVirtualPaging(this.chartSpread);
}
//<<#99915912 - Jeff 2010/04/29

FarPoint.Web.Spread.SpreadChart.prototype.cancelDefault = function (event) {
  this.preventDefault(event);
  this.stopPropagation(event);
  return false;
};

FarPoint.Web.Spread.SpreadChart.prototype.preventDefault = function (event) {
  if (event.preventDefault != null) {
    event.preventDefault();
  }
  else {
    event.returnValue = false;
  }
  return false;
};

FarPoint.Web.Spread.SpreadChart.prototype.stopPropagation = function (event) {
  if (event.stopPropagation != null) {
    event.stopPropagation();
  }
  else {
    event.cancelBubble = true;
  }
  return false;
};

FarPoint.Web.Spread.SpreadChart.prototype.buildResizeCorner = function() {
  if (this.container.isBuildCorner != null && this.container.isBuildCorner == true)
    return;
  else
    this.container.isBuildCorner = true;
  // 99916904 dangjian  
  if (this.container.clientHeight==0)
  {
    var t = this.container;
    while (t!=null && t!=document.body) {
      t.displaySetting = t.style.display;
      t.style.display = "";
      t = t.parentNode;
    }
  }

  this.container.style.height = this.container.clientHeight + 'px';
  this.container.style.width = this.container.clientWidth + 'px';

  var containerBorderWidth = this.container.clientTop;
  var cornerWidth = 7;
  var offsetCorner = Math.floor(cornerWidth / 2);
  var cornerImageUrl = this.chartSpread.getAttribute("chartConerImg");

  var chartObj = this;
  buildOneCorner = function(id, left, top, url, cursor) {
    var corner = document.createElement('IMG');
    corner.build = true;
    corner.src = url;
    corner.style.position = 'absolute';
    corner.style.left = left;
    corner.style.top = top;
    corner.style.cursor = cursor;
    corner.id = id;
    corner.style.display = 'none';
    chartObj.attachEvent(corner, "mousedown", chartObj.startResizeChart, chartObj, null);
    chartObj.attachEvent(corner, "mouseup", chartObj.endResizeChart, chartObj, null);
    chartObj.attachEvent(corner, "click", chartObj.endResizeChart, chartObj, null);
    return corner;
  };

  var id = 'nw-resize';
  var left = (-offsetCorner - (containerBorderWidth * 2)) + 'px';
  var top = (-offsetCorner - (containerBorderWidth * 2)) + 'px';
  var cursor = 'nw-resize';
  var corner = buildOneCorner(id, left, top, cornerImageUrl, cursor);
  this.container.appendChild(corner);

  id = 'ne-resize';
  left = (parseInt(this.container.style.width) - offsetCorner) + 'px';
  top = (-offsetCorner - (containerBorderWidth * 2)) + 'px';
  cursor = 'ne-resize';
  corner = buildOneCorner(id, left, top, cornerImageUrl, cursor);
  this.container.appendChild(corner);

  id = 'sw-resize';
  left = (-offsetCorner - (containerBorderWidth * 2)) + 'px';
  top = (parseInt(this.container.style.height) - offsetCorner) + 'px';
  cursor = 'sw-resize';
  corner = buildOneCorner(id, left, top, cornerImageUrl, cursor);
  this.container.appendChild(corner);

  id = 'se-resize';
  left = (parseInt(this.container.style.width) - offsetCorner) + 'px';
  top = (parseInt(this.container.style.height) - offsetCorner) + 'px';
  cursor = 'se-resize';
  corner = buildOneCorner(id, left, top, cornerImageUrl, cursor);
  this.container.appendChild(corner);

  id = 'n-resize';
  left = (Math.floor(parseInt(this.container.style.width) / 2) - offsetCorner) + 'px';
  top = (-offsetCorner - (containerBorderWidth * 2)) + 'px';
  cursor = 's-resize';
  corner = buildOneCorner(id, left, top, cornerImageUrl, cursor);
  this.container.appendChild(corner);

  id = 's-resize';
  left = (Math.floor(parseInt(this.container.style.width) / 2) - offsetCorner) + 'px';
  top = (parseInt(this.container.style.height) - offsetCorner) + 'px';
  cursor = 's-resize';
  corner = buildOneCorner(id, left, top, cornerImageUrl, cursor);
  this.container.appendChild(corner);

  id = 'w-resize';
  left = (-offsetCorner - (containerBorderWidth * 2)) + 'px';
  top = (Math.floor(parseInt(this.container.style.height) / 2) - offsetCorner - (containerBorderWidth * 2)) + 'px';
  cursor = 'e-resize';
  corner = buildOneCorner(id, left, top, cornerImageUrl, cursor);
  this.container.appendChild(corner);

  id = 'e-resize';
  left = (parseInt(this.container.style.width) - offsetCorner) + 'px';
  top = (Math.floor(parseInt(this.container.style.height) / 2) - offsetCorner - (containerBorderWidth * 2)) + 'px';
  cursor = 'e-resize';
  corner = buildOneCorner(id, left, top, cornerImageUrl, cursor);
  this.container.appendChild(corner);
  
  // 99916904 dangjian  
  var t = this.container;
  while (t!=null && t!=document.body) {
    if (t.displaySetting != null) {
      t.style.display = t.displaySetting;
      t.displaySetting = null;
    } else {
      break;
    }
    t = t.parentNode;
  }

  this.container.isProcessingResize = false;
};

FarPoint.Web.Spread.SpreadChart.prototype.setResizeCornerPosition = function () {
  var nw_resize, sw_resize, w_resize, n_resize, s_resize, e_resize, ne_resize, se_resize;
  var children;
  if (typeof this.container.children == "undefined") {
    // mozilla
    children = this.container.childNodes;
  }
  else {
    // IE
    children = this.container.children;
  }
  for (var i = 0; i < children.length; i++) {
    switch (children[i].id) {
      case 'nw-resize':
        nw_resize = children[i];
        break;
      case 'sw-resize':
        sw_resize = children[i];
        break;
      case 'w-resize':
        w_resize = children[i];
        break;
      case 'n-resize':
        n_resize = children[i];
        break;
      case 's-resize':
        s_resize = children[i];
        break;
      case 'e-resize':
        e_resize = children[i];
        break;
      case 'ne-resize':
        ne_resize = children[i];
        break;
      case 'se-resize':
        se_resize = children[i];
        break;
    }
  }
  var containerBorderWidth = 1;
  var cornerWidth = 7;
  var offsetCorner = Math.floor(cornerWidth / 2);
  //lance - 2010/12/07 - #99918682
  //  n_resize.style.left = (Math.floor((parseInt(this.container.style.width) + (containerBorderWidth * 2)) / 2) - offsetCorner - (containerBorderWidth * 2)) + 'px';
  //  s_resize.style.left = (Math.floor((parseInt(this.container.style.width) + (containerBorderWidth * 2)) / 2) - offsetCorner - (containerBorderWidth * 2)) + 'px';
  //  ne_resize.style.left = (parseInt(this.container.style.width) + (containerBorderWidth * 2) - offsetCorner - (containerBorderWidth * 2)) + 'px';
  //  e_resize.style.left = (parseInt(this.container.style.width) + (containerBorderWidth * 2) - offsetCorner - (containerBorderWidth * 2)) + 'px';
  //  se_resize.style.left = (parseInt(this.container.style.width) + (containerBorderWidth * 2) - offsetCorner - (containerBorderWidth * 2)) + 'px';
  //  se_resize.style.top = (parseInt(this.container.style.height) + (containerBorderWidth * 2) - offsetCorner - (containerBorderWidth * 2)) + 'px';
  //  s_resize.style.top = (parseInt(this.container.style.height) + (containerBorderWidth * 2) - offsetCorner - (containerBorderWidth * 2)) + 'px';
  //  sw_resize.style.top = (parseInt(this.container.style.height) + (containerBorderWidth * 2) - offsetCorner - (containerBorderWidth * 2)) + 'px';
  //  w_resize.style.top = (Math.floor((parseInt(this.container.style.height) + containerBorderWidth) / 2) - offsetCorner - (containerBorderWidth * 2)) + 'px';
  //  e_resize.style.top = (Math.floor((parseInt(this.container.style.height) + containerBorderWidth) / 2) - offsetCorner - (containerBorderWidth * 2)) + 'px';

  var container = $$(this.container);
  var w = container.innerWidth();
  var h = container.innerHeight();
  var containerBorderHoz = containerBorderWidth * 2;
  var center = (Math.floor((w + containerBorderHoz) / 2) - offsetCorner - containerBorderHoz) + 'px';
  n_resize.style.left = center;
  s_resize.style.left = center;
  var right = (w - offsetCorner) + 'px';
  ne_resize.style.left = right;
  e_resize.style.left = right;
  se_resize.style.left = right;
  var bottom = (h - offsetCorner) + 'px';
  se_resize.style.top = bottom;
  s_resize.style.top = bottom;
  sw_resize.style.top = bottom;
  var middle = (Math.floor((h + containerBorderWidth) / 2) - offsetCorner - containerBorderHoz) + 'px';
  w_resize.style.top = middle;
  e_resize.style.top = middle;
};

//lance - 2010/12/07 - #99918682
FarPoint.Web.Spread.SpreadChart.prototype.hasCustomHighlightCss = function () {
  ///<summary>Gets the value which indicates whether the activated chart's container has custom CSS class.</summary>
  ///<returns>
  /// TRUE if the activated chart's container has custom CSS class.
  /// Otherwise, FALSE indicates that the activated chart's container is highlighted automatically by a 1px red border.
  ///</returns>
  var chartcssClass = this.container.getAttribute("selectedCssClass");
  return (chartcssClass != null && chartcssClass.length > 0);
}

FarPoint.Web.Spread.SpreadChart.prototype.highlightChart = function () {
  //lance - 2010/12/07 - #99918682
  if (this.hasCustomHighlightCss()) {
    this.container.className = this.container.getAttribute("selectedCssClass");
  }
  else {
    this.container.style.border = "double 1px red"
    this.container.style.left = (parseInt(this.container.style.left) - 1) + 'px';
    this.container.style.top = (parseInt(this.container.style.top) - 1) + 'px';
  }

  var children;
  if (typeof this.container.children == "undefined") {
    // mozilla
    children = this.container.childNodes;
  }
  else {
    // IE
    children = this.container.children;
  }
  for (var i = 0; i < children.length; i++) {
    if (children[i].id != null && children[i].id.indexOf("-resize") > 0) {
      children[i].style.display = '';
    }
  }

  //>>zesson - 2010/12/15 - #99919428
  if (this.container.getAttribute("sizeChart") != "false") {
  this.setResizeCornerPosition();  //lance - 2010/12/07 - #99918682
  }
  //<<

  chartObj = this;
  this.isActived = true;
};

FarPoint.Web.Spread.SpreadChart.prototype.normalChart = function () {
  //lance - 2010/12/07 - #99918682
  if (!this.isActived) return;

  if (this.hasCustomHighlightCss()) {
    this.container.className = "";
  }
  else {
    this.container.style.left = (parseInt(this.container.style.left) + 1) + 'px';
    this.container.style.top = (parseInt(this.container.style.top) + 1) + 'px';
    this.container.style.border = "none"
  }

  var children;
  if (typeof this.container.children == "undefined") {
    // mozilla
    children = this.container.childNodes;
  }
  else {
    // IE
    children = this.container.children;
  }
  for (var i = 0; i < children.length; i++) {
    if (children[i].id != null && children[i].id.indexOf("-resize") > 0) {
      children[i].style.display = 'none';
    }
  }

  //99915941 CoreyKou 2010.05.05
  this.container.removeAttribute("isActive");
  this.isActived = false;
};

FarPoint.Web.Spread.SpreadChart.prototype.setChartLocation = function (chartInfo, left, top) {
  if (chartInfo == null)
    return;
  //>>#99915912 - Jeff 2010/05/05
  if (this.chartSpread.virtualPaging == "true" && this.chartSpread.virtualTop > 0)
    top = top - this.chartSpread.virtualTop;
  //<<#99915912 - Jeff 2010/05/05
  var node = chartInfo.getElementsByTagName("location")[0];
  var browser = $$.browser;
  if (browser.ie) {
    if (this.xmlData == null) {  //lance - 2011/03/22 - 99920694
      if (node == null) {
        node = this.tableChartData.createNode("element", "location", "");
      }
      node.text = left + "," + top;
    }
    else {
      var xmlData = this.xmlData;
      if (node == null) {
        node = xmlData.createElement("location");
      }
      xmlData.textContent(node, left + "," + top);
    }
  }
  else {
    if (node == null) {
      node = document.createElement("location");
    }
    node.innerHTML = left + "," + top;
  }
  chartInfo.appendChild(node);
  if (FpChartUtil.isIE()) {
    this.chartSpread.UpdatePostbackData();
  }
  else {
    the_fpSpread.UpdatePostbackData(this.chartSpread);
  }
};

FarPoint.Web.Spread.SpreadChart.prototype.setChartSize = function (chartInfo, width, height) {
  if (chartInfo == null)
    return;
  var node = chartInfo.getElementsByTagName("size")[0];
  var browser = $$.browser;
  if (browser.ie) {
    if (this.xmlData == null) {  //lance - 2011/03/22 - 99920694
      if (node == null) {
        node = this.tableChartData.createNode("element", "size", "");
      }
      node.text = width + "," + height;
    }
    else {
      var xmlData = this.xmlData;
      if (node == null) {
        node = xmlData.createElement("size");
      }
      xmlData.textContent(node, width + "," + height);
    }
  }
  else {
    if (node == null) {
      node = document.createElement("size");  //Deal with namespace problem
    }
    if (browser.ie)
      node.textContent = width + "," + height;
    else
      node.innerHTML = width + "," + height;
  }
  chartInfo.appendChild(node);
};

FarPoint.Web.Spread.SpreadChart.prototype.attachEvent = function(target, eventName, handler, sender, argsObject) {
  if (target == null || eventName == null || handler == null)
    return;

  var eventHandler = function(e) {
    e.sender = sender;
    e.args = argsObject;

    handler.call(target, e);
  }

  var targetId = target.id;

  if (targetId == null && target.nodeName == "#document") {
    targetId = "document";
  }
  var handlerName = sender.chartID + ":" + targetId + ":" + eventName;
  var handlerList = this.eventHandlers[handlerName];

  if (handlerList == null) {
    handlerList = {};
    this.eventHandlers[handlerName] = handlerList;
  }

  if (handlerList[handler.toString()] == null) {
    handlerList[handler.toString()] = eventHandler;

    if (target.addEventListener != null) {
      target.addEventListener(eventName, eventHandler, false);
    }
    else
      if (target.attachEvent != null) {
      target.attachEvent("on" + eventName, eventHandler);
    }
  }
};

FarPoint.Web.Spread.SpreadChart.prototype.detachEvent = function(target, eventName, handler, sender) {
  if (target == null || eventName == null || handler == null || this.eventHandlers == null)
    return;

  var targetId = target.id;

  if (targetId == null && target.nodeName == "#document") {
    targetId = "document";
  }
  var handlerName = sender.chartID + ":" + targetId + ":" + eventName;
  var handlerList = this.eventHandlers[handlerName];

  if (handlerList == null)
    return;

  eventHandler = handlerList[handler.toString()];

  if (eventHandler != null) {
    if (target.removeEventListener != null) {
      target.removeEventListener(eventName, eventHandler, false);
    }
    else
      if (target.detachEvent != null) {
      target.detachEvent("on" + eventName, eventHandler);
    }
    handlerList[handler.toString()] = null;
  }
};

FarPoint.Web.Spread.SpreadChart.prototype.fireEvent = function(target, event, left, top, width, height) {
  if (target == null) {
    return;
  }
  if (document.createEvent) {// DOM event model
    var e = document.createEvent("Events");
    e.initEvent(event, true, false);
  }
  else
    if (document.createEventObject) { // IE event model
    var e = document.createEventObject();
  }
  else
    return;

  e.chartObj = this;
  e.top = top;
  e.left = left;
  e.width = width;
  e.height = height;

  if (target.dispatchEvent)
    target.dispatchEvent(e); // DOM
  else
    if (target.fireEvent)
    target.fireEvent("on" + event, e); // IE
};

FarPoint.Web.Spread.SpreadChart.prototype.updateChartImgitem = function() {
  if (this.container.getAttribute("imageDirty") == "true") {
    chartImg = this.getChart();

    if (this.chartSpread.inDesign)
      chartImg.src = chartImg.src;
    else
      chartImg.src = chartImg.src + "#" + Math.random();
  }
};

FarPoint.Web.Spread.SpreadChart.prototype.isLeftButtonClicked = function(event) {
  var chartObj = event.sender;
  if (FpChartUtil.isIE()) {
    return event.button < 2;
  }
  else {
    return event.button == 0;
  }
}

FarPoint.Web.Spread.SpreadChart.prototype.dispose = function() {
  this.detachEvent(this.container, "mousedown", this.containerMousedown, this);
  if (this.container.getAttribute("moveChart") != "false") {
    this.detachEvent(document, "mouseup", this.endDragChart, this);
    this.detachEvent(this.container, "mouseup", this.endDragChart, this);
    this.detachEvent(this.container, "keydown", this.chartKeyDown, this);
  }
};

// Raw Events.
FarPoint.Web.Spread.SpreadChart.prototype.containerMousedown = function(event) {
  var chartObj = event.sender;
  if (chartObj.container.getAttribute("selectChart") == "false") return; //#99917409, #99917483 - Jeff 2010/08/06
  chartObj.chartSpread.SetActiveChart(chartObj.chartID); //99915923 Justin 2010/05/11
  with (chartObj) {
    if (!chartObj.isLeftButtonClicked(event) || container.isResize == true) {
      return;
    }
    if (container.getAttribute("moveChart") != "false") {
      startDragChart(event);
    }
    chartObj.fireEvent(chartObj.container, "click", chartObj.chartID); //>>99917440 Justin 2010/08/06
  }
};

FarPoint.Web.Spread.SpreadChart.prototype.startResizeChart = function (event) {
  var chartObj = event.sender;
  with (chartObj) {
    var corner = this;
    container.isDrag == false;
    container.isResize = true;

    attachEvent(document, "mouseup", endResizeChart, chartObj, null);
    attachEvent(document, "click", endResizeChart, chartObj, null);
    attachEvent(document, "selectstart", cancelDefault, chartObj, null);

    corner.isResize = true;
    corner.startX = event.clientX;
    corner.startY = event.clientY;
    if (typeof corner.setCapture != "undefined") {
      corner.setCapture();
    }

    attachEvent(corner, "mousemove", resizeChart, chartObj, null);
    attachEvent(document, "mousemove", resizeChart, chartObj, null);
    currentResizeCorner = corner;
  }
  chartObj.preventDefault(event); //99917214 Justin 2010/07/26
};

FarPoint.Web.Spread.SpreadChart.prototype.resizeChart = function(event) {
  var chartObj = event.sender;

  with (chartObj) {
    if (!chartObj.isLeftButtonClicked(event)) {
      return;
    }

    var corner = currentResizeCorner;

    if (corner == null)
      return;

    var resizeType = corner.id;

    //99915104 Jusitn 2010/05/10: create object o manage style info of chart
    var chartStyleInfo = chartObj.chartStyleInfo;
    var minWidth = Math.max(chartStyleInfo.getStyleWidth(), 18);
    var minHeight = Math.max(chartStyleInfo.getStyleHeight(), 18);
    if (corner == null)
      return;
    if (container.isProcessingResize == true)
      return;
    container.isProcessingResize = true;
    //99916057 CoreyKou 20100508
    container.style.cursor = corner.style.cursor;
    var currentInfo = {
      left: parseInt(container.style.left),
      top: parseInt(container.style.top),
      width: container.clientWidth,
      height: container.clientHeight
    }
    var newInfo = {
      left: currentInfo.left,
      top: currentInfo.top,
      width: currentInfo.width,
      height: currentInfo.height
    }

    if (resizeType == 'e-resize' || resizeType == 'ne-resize' || resizeType == 'se-resize') { /* East resize */
      newInfo.width = Math.max(minWidth, (parseInt(container.style.width) + event.clientX - corner.startX));
    }
    if (resizeType == 's-resize' || resizeType == 'sw-resize' || resizeType == 'se-resize') {
      newInfo.height = Math.max(minHeight, (parseInt(container.style.height) + event.clientY - corner.startY));
    }
    if (resizeType == 'n-resize' || resizeType == 'nw-resize' || resizeType == 'ne-resize') {
      var newTop = Math.max(0, (parseInt(container.style.top) + event.clientY - corner.startY));
      newInfo.height += (currentInfo.top - newTop);
      newInfo.top = newTop;
    }
    if (resizeType == 'w-resize' || resizeType == 'sw-resize' || resizeType == 'nw-resize') {
      var newLeft = Math.max(0, (parseInt(container.style.left) + event.clientX - corner.startX));
      newInfo.width += (currentInfo.left - newLeft);
      newInfo.left = newLeft;
    }
    if (newInfo.height <= minHeight)
      newInfo.height = minHeight + 2; //add some for container border
    if (newInfo.width <= minWidth)
      newInfo.width = minWidth + 2;
    if (newInfo) {
      container.style.left = newInfo.left + 'px';
      container.style.top = newInfo.top + 'px';
      container.style.width = newInfo.width + 'px';
      container.style.height = newInfo.height + 'px';

      chartStyleInfo.setInnerSize(newInfo.width, newInfo.height); //99915104
    }

    setResizeCornerPosition();
    corner.startX = event.clientX;
    corner.startY = event.clientY;
    cancelDefault(event);
    // 99915958 Corey 20100504
    if (FpChartUtil.isIE()) {
      chartObj.chartSpread.ResumeLayout(true);
    }
    container.isProcessingResize = false;
  }
};

FarPoint.Web.Spread.SpreadChart.prototype.endResizeChart = function (event) {
  var chartObj = event.sender;
  with (chartObj) {
    detachEvent(document, "mousemove", resizeChart, chartObj);
    detachEvent(document, "mouseup", endResizeChart, chartObj);
    detachEvent(document, "click", endResizeChart, chartObj);
    detachEvent(document, "selectstart", cancelDefault, chartObj);
    container.isResize = false;
    //99916057 CoreyKou 20100508
    container.style.cursor = "default";
    var corner = currentResizeCorner;
    if (corner != null && corner.isResize != null && corner.isResize == true) {
      corner.isResize = false;
      detachEvent(corner, "mousemove", resizeChart, chartObj);
      if (typeof corner.releaseCapture != "undefined") {
        corner.releaseCapture();
      }
      var chartInfo = addChartInfo();
      //lance - 2010/12/07 - #99918682
      //var left = parseInt(container.style.left);
      //var top = parseInt(container.style.top);
      var location = chartObj.getChartLocation();
      var left = location.x, top = location.y;
      setChartLocation(chartInfo, left, top);

      var width = parseInt(chart.width);
      var height = parseInt(chart.height);
      setChartSize(chartInfo, width, height);
      if (FpChartUtil.isIE()) {
        chartSpread.RefreshChart();
      }
      else {
        the_fpSpread.RefreshChart(chartSpread);
      }

      fireEvent(container, "resize", left, top, width, height);
    }
    // 99915958 Corey 20100504
    if (FpChartUtil.isIE()) {
      chartObj.chartSpread.ResumeLayout(true);
    }
    currentResizeCorner = null;
    cancelDefault(event);
  }
};

FarPoint.Web.Spread.SpreadChart.prototype.startDragChart = function (event) {
    this.container.style.cursor = "move";
    this.container.isDrag = true;
    
    this.container._isMoved = false;
    this.container._oldLocation = chartObj.getChartLocation();
    
    this.container.x = this.container.startX = event.clientX;
    this.container.y = this.container.startY = event.clientY;
    
    this.attachEvent(this.container, "mousemove", this.dragChart, this, null);
    if (typeof this.container.setCapture != "undefined") {
       this.container.setCapture();
    }

    if (!FpChartUtil.isIE()) {
        this.container.setAttribute("tabindex", 0);
    }

    this.attachEvent(document, "mousemove", this.dragChart, this, null);
    this.attachEvent(document, "selectstart", this.cancelDefault, this, null);
    this.cancelDefault(event);  //99917214 Justin 2010/07/26
};

FarPoint.Web.Spread.SpreadChart.prototype.dragChart = function (event) {
    var chartObj = event.sender;
    with (chartObj) {
        if (!chartObj.isLeftButtonClicked(event)) {
            chartObj.endDragChart(event);
            return;
        }

        if (container.isDrag != null && container.isDrag == true) {
            container.style.left = (parseInt(container.style.left) + event.clientX - container.startX) + 'px';
            container.style.top = (parseInt(container.style.top) + event.clientY - container.startY) + 'px';
            container.startX = event.clientX;
            container.startY = event.clientY;
        }

        if (FpChartUtil.isIE()) {//#99917601 - Jeff 2010/09/08 - shoud try to find better solution in future
            chartObj.chartSpread.ResumeLayout(true);
        }
    }
};

FarPoint.Web.Spread.SpreadChart.prototype.chartKeyDown = function (event) {
  var chartObj = event.sender;
  with (chartObj) {
    if (event.keyCode == 38) {
      container.style.top = (parseInt(container.style.top + 0) - 1) + 'px';
    }
    else if (event.keyCode == 40) {
      container.style.top = (parseInt(container.style.top + 0) + 1) + 'px';
    }
    else if (event.keyCode == 37) {
      container.style.left = (parseInt(container.style.left + 0) - 1) + 'px';
    }
    else if (event.keyCode == 39) {
      container.style.left = (parseInt(container.style.left + 0) + 1) + 'px';
    }
    var chartInfo = addChartInfo();
    //lance - 2010/12/07 - #99918682
    //setChartLocation(chartInfo, parseInt(container.style.left), parseInt(container.style.top));
    var location = chartObj.getChartLocation();
    setChartLocation(chartInfo, location.x, location.y);
    cancelDefault(event);

    chartObj.stopVirtualPaging(); //#99915912 - Jeff 2010/04/29
    // Resize Spread
    if (FpChartUtil.isIE()) {
      chartObj.chartSpread.ResumeLayout(true);
    }
    else {
      chartObj.chartSpread.SizeSpread(chartObj.chartSpread);
    }
    chartObj.startVirtualPaging(); //#99915912 - Jeff 2010/04/29
  }
  return false;
};

FarPoint.Web.Spread.SpreadChart.prototype.endDragChart = function (event) {
  var chartObj = event.sender;
  with (chartObj) {  
    if (container.isDrag != null && container.isDrag == true) {
      container.style.cursor = "default";
      container.isDrag = false;
      if (typeof container.releaseCapture != "undefined") {
        container.releaseCapture();
      }

      detachEvent(container, "mousemove", dragChart, chartObj);
      detachEvent(document, "mousemove", dragChart, chartObj);
      detachEvent(document, "selectstart", cancelDefault, chartObj);           
      cancelDefault(event);      
      var chartInfo = addChartInfo();
      //lance - 2010/12/07 - #99918682
      //var left = parseInt(container.style.left);
      //var top = parseInt(container.style.top);
      var location = chartObj.getChartLocation();
      var left = location.x, top = location.y;
      setChartLocation(chartInfo, left, top);
      fireEvent(container, "move", left, top);

      if(container._oldLocation.x != location.x || container._oldLocation.y != location.y)
        container._isMoved = true;
      container._oldLocation = null;
      
      // Resize Spread
      chartObj.stopVirtualPaging(); //#99915912 - Jeff 2010/04/29
      if (FpChartUtil.isIE()) {
        chartObj.chartSpread.ResumeLayout(true);
      }
      else {
        chartObj.chartSpread.SizeSpread(chartObj.chartSpread);
      }
      chartObj.startVirtualPaging(); //#99915912 - Jeff 2010/04/29
      //>>#99917502 - Jeff 2010/08/25
      if (chartSpread.getViewportSize) {
        var needUpdate = false;
        var vpSize = chartSpread.getViewportSize();
        if (vpSize.height < container.parentNode.offsetHeight && container.parentNode.scrollTop > 0)
          needUpdate = true;

        if (vpSize.width < container.parentNode.offsetWidth && container.parentNode.scrollLeft > 0)
          needUpdate = true;
        if (needUpdate && FpChartUtil.isIE()) {
          chartObj.scrollChartIntoView();
          chartObj.chartSpread.ResumeLayout(true);
        }
      }

      if ((container.x == event.clientX) || (container.y == event.clientY)) {
        chartObj.scrollChartIntoView();
        if (container.x == event.clientX) container.moveLeft = null;
        if (container.y == event.clientY) container.moveUp = null;
      }
      else {
        container.moveLeft = (container.x > event.clientX);
        container.moveUp = (container.y > event.clientY);
        container.startX = event.clientX;
        container.startY = event.clientY;
      }
      //<<#99917502 - Jeff 2010/08/25
    }
  }
};

//
//Helper class to get some common values used in repositioning chart
//
FarPoint.Web.Spread.SpreadChartContainer = function(spreadChart) {
  this.container = spreadChart.container;
  this.init();
}

FarPoint.Web.Spread.SpreadChartContainer.prototype.init = function() {
  var c = this.container;
  var corner = this.getCornerWidth() / 2;
  var border = this.getBorderWidth();
  this.Left = c.offsetLeft;
  this.Right = c.offsetLeft + c.clientWidth;
  this.Top = c.offsetTop;
  this.Bottom = c.offsetTop + c.clientHeight;
  //>>#99917502 - Jeff 2010/08/25
  this.X = this.Left;
  this.Y = this.Top;
  //<<#99917502 - Jeff 2010/08/25
  this.Width = Math.max(this.Right - this.Left, 0);
  this.Height = Math.max(this.Bottom - this.Top, 0);
}

FarPoint.Web.Spread.SpreadChartContainer.prototype.getCornerWidth = function() {
  return 7;
}

FarPoint.Web.Spread.SpreadChartContainer.prototype.getBorderWidth = function() {
  return 1;
}

FarPoint.Web.Spread.SpreadChartContainer.prototype.focus = function() {
  if (this.container != null)
    this.container.focus();
}

FarPoint.Web.Spread.SpreadChartContainer.prototype.select = function() {
  if (this.container != null)
    this.container.select();
}

FarPoint.Web.Spread.SpreadChartContainer.prototype.isSelect = function() {
  return this.container.select;
}


//>>99915104 Justin 2010/05/10
//
//Helper class contains information about style of chart
//
FarPoint.Web.Spread.SpreadChartStyleInfo = function(spreadChart) {
  this.spreadChart = spreadChart;
  this.init();
}

FarPoint.Web.Spread.SpreadChartStyleInfo.prototype.init = function() {
  var chartImage = this.getChartImage();
  var style = chartImage.style;
  this.width = chartImage.width;
  this.height = chartImage.height;
  this.border = {
    top: style.borderTopWidth != '' ? parseInt(style.borderTopWidth) : 0,
    right: style.borderRightWidth != '' ? parseInt(style.borderRightWidth) : 0,
    bottom: style.borderBottomWidth != '' ? parseInt(style.borderBottomWidth) : 0,
    left: style.borderLeftWidth != '' ? parseInt(style.borderLeftWidth) : 0
  };
  this.margin = {
    top: style.marginTop != '' ? parseInt(style.marginTop) : 0,
    right: style.marginRight != '' ? parseInt(style.marginRight) : 0,
    bottom: style.marginBottom != '' ? parseInt(style.marginBottom) : 0,
    left: style.marginLeft != '' ? parseInt(style.marginLeft) : 0
  };
  this.padding = {
    top: style.paddingTop != '' ? parseInt(style.paddingTop) : 0,
    right: style.paddingRight != '' ? parseInt(style.paddingRight) : 0,
    bottom: style.paddingBottom != '' ? parseInt(style.paddingBottom) : 0,
    left: style.paddingLeft != '' ? parseInt(style.paddingLeft) : 0
  };
}

FarPoint.Web.Spread.SpreadChartStyleInfo.prototype.getChartImage = function() {
  return this.spreadChart.getChart();
}

FarPoint.Web.Spread.SpreadChartStyleInfo.prototype.getStyleWidth = function() {
  return (this.border.left + this.border.right) + (this.margin.left + this.margin.right) + (this.padding.left + this.padding.right);
}

FarPoint.Web.Spread.SpreadChartStyleInfo.prototype.getTotalWidth = function() {
  return this.width + this.getStyleWidth();
}

FarPoint.Web.Spread.SpreadChartStyleInfo.prototype.getStyleHeight = function() {
  return (this.border.top + this.border.bottom) + (this.margin.top + this.margin.bottom) + (this.padding.top + this.padding.bottom);
}

FarPoint.Web.Spread.SpreadChartStyleInfo.prototype.getTotalHeight = function() {
  return this.height + this.getStyleHeight();
}

FarPoint.Web.Spread.SpreadChartStyleInfo.prototype.setInnerSize = function(w, h) {
  var chartImage = this.getChartImage();
  chartImage.style.width = Math.max((w - this.getStyleWidth()), 1) + 'px';
  chartImage.style.height = Math.max((h - this.getStyleHeight()), 1) + 'px';
  this.init();
}
//<<99915104
//
//Helper class for performing common funtionality
//
FarPoint.Web.Spread.ChartUtil = function () {
  this.isIE = function () {
    return this.checkBrowserByName("IE");
  };

  this.checkBrowserByName = function (browsername, version) {
    // Check browser version
    var ua = window.navigator.userAgent;
    var rst = false;
    var bn = ("" + browsername).toLowerCase();
    if ((bn.indexOf("ms") >= 0) || (bn.indexOf("msie") >= 0) || (bn.indexOf("ie") >= 0))
      rst = (ua.indexOf("MSIE") >= 1);
    else
      if ((bn.indexOf("safari") >= 0) || (bn.indexOf("apple") >= 0))
        rst = (ua.indexOf("Safari") >= 1);
      else
        if ((bn.indexOf("ff") >= 0) || (bn.indexOf("firefox") >= 0))
          rst = (ua.indexOf("Firefox") >= 1);
    return rst;
  };
  //>>99916057 Justin 2010/05/27
  this.ChartPosition = {
    OUTSIDE: -1,
    INSIDE: 0,
    INTERSECT: 1
  };

  this.getViewportChartIntersection = function (x1, y1, w1, h1, x2, y2, h2, w2) {
    var pos = this.ChartPosition.OUTSIDE;

    if ((x2 < x1 && x2 + w2 > x1) || (x2 > x1 && x2 < x1 + w1)) {
      pos = this.ChartPosition.INTERSECT;
    } else if ((y1 > y2 && y1 < y2 + h2) || (y2 > y1 && y2 < y1 + h1)) {
      pos = this.ChartPosition.INTERSECT;
    }

    if (pos == 1) { //not inside
      if ((x1 > x2 && x1 + w1 < x2 + w2 && y1 > y2 && y1 + h1 < y2 + h2) || (x2 > x1 && x2 + h2 < x1 + h1 && y2 > y1 && y2 + h2 < y1 + h1))
        pos = this.ChartPosition.INSIDE;
    }
    return pos;
  }
  //<<99916057
  //>99917487 Justin 2010/08/11
  this.getChartsIdCollection = function (spread) {
    var result = [];
    var count = 0;
    var items = spread.getElementsByTagName("DIV");
    for (var i = 0; i < items.length; i++) {
      var item = items[i];
      if (item.id.indexOf("_chartContainer") > 0) {
        result[count] = item.id;
        count++;
      }
    }
    return result;
  }

  //#99918374 - Jeff 2011/04/01
  this.getChartElements = function (spread) {
    var result = [];
    var items = spread.getElementsByTagName("DIV");
    for (var i = 0; i < items.length; i++) {
      var item = items[i];
      if (item.id.indexOf("_chartContainer") > 0) {
        result.push(item);
      }
    }
    return result;
  }

  this.getUnionChartsBounds = function (spread) {
    var charts = this.getChartsIdCollection(spread);
    if (charts.length == 0) return false;
    var bounds = { top: 0, left: 0, right: 0, bottom: 0 };
    var first = true;
    for (var i = 0; i < charts.length; i++) {
      var id = charts[i];
      var chartContainer = document.getElementById(id);
      if (chartContainer) {
        var t = chartContainer.offsetTop;
        var l = chartContainer.offsetLeft;
        var r = l + chartContainer.offsetWidth;
        var b = t + chartContainer.offsetHeight;
        if (first) {
          bounds = { top: t, left: l, right: r, bottom: b };
          first = false;
        } else {
          bounds.top = Math.min(t, bounds.top);
          bounds.left = Math.min(l, bounds.left);
          bounds.right = Math.max(r, bounds.right);
          bounds.bottom = Math.max(b, bounds.bottom);
        }
      }
    }
    return bounds;
  }
  //<<99917487

  this.doMapAreaPostback = function (mapArea, id, args) {
    var c = document.getElementById(mapArea.id);
    if (c != null) {
      if (typeof (c.parentNode.parentNode.getChartObj) == "undefined") return; //#99923440 - Phillip - 2011-10-12
      var cobj = c.parentNode.parentNode.getChartObj();
      if (typeof (cobj) != "undefined" && cobj != null && !cobj.getIsDrag()) {
        __doPostBack(id, args);
      }
    }
  }
}


FpChartUtil = new FarPoint.Web.Spread.ChartUtil();
//#99923986 - Need to notify that the script has been loaded, otherwise an exception will be thrown!
if (typeof (Sys) !== 'undefined') Sys.Application.notifyScriptLoaded();