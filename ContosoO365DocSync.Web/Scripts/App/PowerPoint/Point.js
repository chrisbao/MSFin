$(function () {
    Office.initialize = function (reason) {
        $(document).ready(function () {
            point.init();
        });
    };
});

var point = (function () {
    var point = {
        ///The URL of current open powerpoint document.
        filePath: "",
        ///Define all UI controls.
        controls: {},
        ///The select excel document.
        selectDocument: null,
        ///Define the source points. 
        points: [],
        ///Define default search keywords.
        sourcePointKeyword: "",
        ///Define default page index 
        pagerIndex: 0,
        ///Define default page size.
        pagerSize: 30,
        ///Default default page total count.
        pagerCount: 0,
        ///Define all service endpoints.
        endpoints: {
            catalog: "/api/SourcePointCatalog?documentId=",
            token: "/api/GraphAccessToken",
            sharePointToken: "/api/SharePointAccessToken",
            graph: "https://graph.microsoft.com/v1.0"
        },
        ///Define the api host and token.
        api: {
            host: "",
            token: "",
            sharePointToken: ""
        }
    }, that = point;
    ///Initialize the event handlers.
    that.init = function () {
        that.filePath = window.location.href.indexOf("localhost") > -1 ? "https://cand3.sharepoint.com/Shared%20Documents/Presentation.pptx" : Office.context.document.url;
        that.controls = {
            body: $("body"),
            main: $(".main"),
            spinner: new fabric["Spinner"]($(".popups .ms-Spinner").get(0)),
            titleName: $("#lblSourcePointName"),
            sourcePointName: $("#txtSearchSourcePoint"),
            searchSourcePoint: $("#iSearchSourcePoint"),
            autoCompleteControl2: $("#autoCompleteWrap2"),
            list: $("#listPoints"),
            headerListPoints: $("#headerListPoints"),
            popupMain: $("#popupMain"),
            popupErrorOK: $("#btnErrorOK"),
            popupMessage: $("#popupMessage"),
            popupProcessing: $("#popupProcessing"),
            popupSuccessMessage: $("#lblSuccessMessage"),
            popupErrorMain: $("#popupErrorMain"),
            popupErrorTitle: $("#popupErrorMain #lblErrorTitle"),
            popupErrorMessage: $("#popupErrorMain #lblErrorMessage"),
            popupErrorRepair: $("#popupErrorMain #lblErrorRepair"),
            popupBrowseList: $("#browseList"),
            popupBrowseBack: $("#btnBrowseBack"),
            popupBrowseCancel: $("#btnBrowseCancel"),
            popupBrowseMessage: $("#txtBrowseMessage"),
            popupBrowseLoading: $("#popBrowseLoading"),
            pager: $("#pager"),
            pagerTotal: $("#pagerTotal"),
            pagerPages: $("#pagerPages"),
            pagerCurrent: $("#pagerCurrent"),
            pagerPrev: $("#pagerPrev"),
            pagerNext: $("#pagerNext"),
            pagerValue: $("#pagerValue"),
            pagerGo: $("#pagerGo")
        };

        that.controls.body.click(function () {
            that.action.body();
        });

        $(document).on("click", ".n-select, .ms-ContextualMenu-item:has(.ms-Icon--Add)", function () {
            that.browse.init();
        });
        $(document).on("click", ".n-refresh, .ms-ContextualMenu-item:has(.ms-Icon--Refresh)", function () {
            if (that.selectDocument != null) {
                that.list();
            }
            else {
                that.popup.message({ success: false, title: "Please select file first." });
            }
        });
        that.controls.sourcePointName.focus(function () {
            that.action.dft(this, true);
        });
        that.controls.sourcePointName.blur(function () {
            that.action.dft(this, false);
        });
        that.controls.sourcePointName.keydown(function (e) {
            if (e.keyCode == 13) {
                that.controls.searchSourcePoint.click();
            }
            else if (e.keyCode == 38 || e.keyCode == 40) {
                app.search.move({ result: that.controls.autoCompleteControl2, target: that.controls.sourcePointName, down: e.keyCode == 40 });
            }
        });
        that.controls.sourcePointName.bind("input", function (e) {
            that.action.autoComplete2();
        });
        that.controls.searchSourcePoint.click(function () {
            that.action.searchSourcePoint();
        });
        that.controls.popupErrorOK.click(function () {
            that.action.ok();
        });
        that.controls.popupBrowseList.on("click", "li", function () {
            that.browse.select($(this));
        });
        that.controls.popupBrowseCancel.click(function () {
            that.browse.popup.hide();
        });
        that.controls.popupBrowseBack.click(function () {
            that.browse.popup.back();
        });
        that.controls.list.on("click", ".i-history", function () {
            that.action.history($(this).closest(".point-item"));
            return false;
        });
        that.controls.main.on("click", ".search-tooltips li", function () {
            $(this).parent().parent().find("input").val($(this).text());
            $(this).parent().hide();
            that.controls.searchSourcePoint.click();
        });
        that.controls.main.on("mouseover", ".search-tooltips li", function () {
            $(this).parent().find("li.active").removeClass("active");
            $(this).addClass("active");
        });
        that.controls.main.on("mouseout", ".search-tooltips li", function () {
            $(this).removeClass("active");
        });
        that.controls.pagerPrev.on("click", function () {
            if (!$(this).hasClass("disabled")) {
                that.utility.pager.prev();
            }
        });
        that.controls.pagerNext.on("click", function () {
            if (!$(this).hasClass("disabled")) {
                that.utility.pager.next();
            }
        });
        that.controls.pagerValue.on("keydown", function (e) {
            if (e.keyCode == 13) {
                that.controls.pagerGo.click();
            }
        });
        that.controls.pagerGo.on("click", function () {
            var _v = $.trim(that.controls.pagerValue.val());
            if (isNaN(_v)) {
                that.popup.message({ success: false, title: "Only numbers are a valid input." });
            }
            else {
                _n = parseInt(_v);
                if (_n > 0 && _n <= that.pagerCount) {
                    that.utility.pager.init({ index: _n });
                }
                else {
                    that.popup.message({ success: false, title: "Invalid number." });
                }
            }
        });
        $(window).resize(function () {
            that.utility.height();
        });
        that.utility.height();
        that.action.dft(that.controls.sourcePointName, false);
        that.utility.pager.status({ length: 0 });

        that.fabric.init();
    };
    ///Initialize the fabric components
    that.fabric = {
        init: function () {
            ///Create fabric command bar
            new fabric['CommandBar']($(".nav-header").get(0));

            var SpinnerElements = document.querySelectorAll(".ms-Spinner");
            for (var i = 0; i < SpinnerElements.length; i++) {
                new fabric['Spinner'](SpinnerElements[i]);
            }
        }
    };
    ///Load the source point list.
    that.list = function () {
        that.popup.processing(true);
        that.service.catalog({ documentId: that.selectDocument.Id }, function (result) {
            if (result.status == app.status.succeeded) {
                that.popup.processing(false);
                that.controls.titleName.html("Source Points in " + that.selectDocument.Name);
                that.controls.titleName.prop("title", "Source Points in " + that.selectDocument.Name);
                that.points = result.data && result.data.SourcePoints && result.data.SourcePoints.length > 0 ? result.data.SourcePoints : [];
                if (that.points.length > 0) {
                    that.utility.pager.init({ index: 1 });
                }
                else {
                    that.utility.pager.status({ length: 0 });
                }
            }
            else {
                that.popup.message({ success: false, title: "Load source points failed." });
            }
        });
    };
    ///The utility methods.
    that.utility = {
        ///Display as 0n.
        format: function (n) {
            return n > 9 ? n : ("0" + n);
        },
        ///Display AM/PM and convert to PST.
        date: function (str) {
            var _v = new Date(str), _d = _v.getDate(), _m = _v.getMonth() + 1, _y = _v.getFullYear(), _h = _v.getHours(), _mm = _v.getMinutes(), _a = _h < 12 ? " AM" : " PM";
            return that.utility.format(_m) + "/" + that.utility.format(_d) + "/" + _y + " " + (_h < 12 ? (_h == 0 ? "12" : that.utility.format(_h)) : (_h == 12 ? _h : _h - 12)) + ":" + that.utility.format(_mm) + "" + _a + " PST";
        },
        ///Get sheet/cell name.
        position: function (p) {
            if (p != null && p != undefined) {
                var _i = p.lastIndexOf("!"), _s = p.substr(0, _i).replace(new RegExp(/('')/g), '\''), _c = p.substr(_i + 1, p.length);
                if (_s.indexOf("'") == 0) {
                    _s = _s.substr(1, _s.length);
                }
                if (_s.lastIndexOf("'") == _s.length - 1) {
                    _s = _s.substr(0, _s.length - 1);
                }
                return { sheet: _s, cell: _c };
            }
            else {
                return { sheet: "", cell: "" };
            }
        },
        ///Add source point to source point list.
        add: function (options) {
            that.points.push(options);
        },
        ///Get the file name.
        fileName: function (path) {
            path = decodeURI(path);
            return path.lastIndexOf("/") > -1 ? path.substr(path.lastIndexOf("/") + 1) : (path.lastIndexOf("\\") > -1 ? path.substr(path.lastIndexOf("\\") + 1) : path);
        },
        ///Paging feature for the source point list.
        pager: {
            ///Initialize the source point list UI.
            init: function (options) {
                that.controls.pagerValue.val("");
                that.pagerIndex = options.index ? options.index : 1;
                that.ui.list();
            },
            ///Go to prev page.
            prev: function () {
                that.controls.pagerValue.val("");
                that.pagerIndex--;
                that.ui.list();
            },
            ///Go to next page.
            next: function () {
                that.controls.pagerValue.val("");
                that.pagerIndex++;
                that.ui.list();
            },
            ///Get the status of paging.
            status: function (options) {
                that.pagerCount = Math.ceil(options.length / that.pagerSize);
                that.controls.pagerTotal.html(options.length);
                that.controls.pagerPages.html(that.pagerCount);
                that.controls.pagerCurrent.html(that.pagerCount > 0 ? that.pagerIndex : 0);
                that.pagerIndex == 1 || that.pagerCount == 0 ? that.controls.pagerPrev.addClass("disabled") : that.controls.pagerPrev.removeClass("disabled");
                that.pagerIndex == that.pagerCount || that.pagerCount == 0 ? that.controls.pagerNext.addClass("disabled") : that.controls.pagerNext.removeClass("disabled");
            }
        },
        ///Calculate the height source point list display area.
        height: function () {
            if (that.controls.main.hasClass("manage")) {
                var _h = that.controls.main.outerHeight();
                var _h1 = $("#pager").outerHeight();
                that.controls.list.css("maxHeight", (_h - 206 - 70 - _h1) + "px");
            }
        },
        ///Get an array of paths (server ralative URl splitted by '/' ).
        path: function () {
            var _a = that.filePath.split("//")[1], _b = _a.split("/"), _p = [_b[0]];
            _b.shift();
            _b.pop();
            for (var i = 1; i <= _b.length; i++) {
                var _c = [];
                for (var n = 0; n < i; n++) {
                    _c.push(_b[n]);
                }
                _p.push(_c.join("/"));
            }
            return _p;
        },
        ///Return 5 publish histories
        publishHistory: function (options) {
            var _ph = [], _td = options.data && options.data.length > 0 ? options.data.reverse() : [];
            $.each(_td, function (m, n) {
                var __c = _td[m].Value ? _td[m].Value : "",
                    __p = _td[m > 0 ? m - 1 : m].Value ? _td[m > 0 ? m - 1 : m].Value : "";
                if (m == 0 || __c != __p) {
                    _ph.push(n);
                }
            });
            return _ph.reverse().slice(0, 5);
        }
    };
    ///Define all the event handlers.
    that.action = {
        ///Body click event to hide the suggested search result.
        body: function () {
            $(".search-tooltips").hide();
        },
        ///Set the text box to default style and value.
        dft: function (elem, on) {
            var _k = $.trim($(elem).val()), _kd = $(elem).data("default");
            if (on) {
                if (_k == _kd) {
                    $(elem).val("");
                }
                $(elem).removeClass("input-default");
            }
            else {
                if (_k == "" || _k == _kd) {
                    $(elem).val(_kd).addClass("input-default");
                }
            }
        },
        ///Look up published hisotry of the source point.
        history: function (o) {
            o.hasClass("item-more") ? o.removeClass("item-more") : o.addClass("item-more");
        },
        ///Close the popup.
        ok: function () {
            that.controls.popupMain.removeClass("active message process confirm");
        },
        ///Display source point result layer after entering the source point keyword in search textbox in source point management page.
        autoComplete2: function () {
            var _e = $.trim(that.controls.sourcePointName.val()), _d = that.points;
            if (_e != "") {
                app.search.autoComplete({ keyword: _e, data: _d, result: that.controls.autoCompleteControl2, target: that.controls.sourcePointName });
            }
            else {
                that.controls.autoCompleteControl2.hide();
            }
        },
        ///Search the source points by source point name in management page.
        searchSourcePoint: function () {
            that.sourcePointKeyword = $.trim(that.controls.sourcePointName.val()) == that.controls.sourcePointName.data("default") ? "" : $.trim(that.controls.sourcePointName.val());
            that.utility.pager.init({});
        }
    };
    ///File explorer.
    that.browse = {
        path: [],
        ///Initialize the default UI.
        init: function () {
            that.api.token = "";
            that.browse.path = [];
            that.browse.popup.dft();
            that.browse.popup.show();
            that.browse.popup.processing(true);
            that.browse.token();
        },
        ///Get the graph and sharepoint access token.
        token: function () {
            that.service.token({ endpoint: that.endpoints.token }, function (result) {
                if (result.status == app.status.succeeded) {
                    that.api.token = result.data;
                    that.api.host = that.utility.path()[0].toLowerCase();
                    that.service.token({ endpoint: that.endpoints.sharePointToken }, function (result) {
                        if (result.status == app.status.succeeded) {
                            that.api.sharePointToken = result.data;
                            that.browse.siteCollection();
                        }
                        else {
                            that.document.error({ title: "Get sharepoint access token failed." });
                        }
                    });
                }
                else {
                    that.browse.popup.message("Get graph access token failed.");
                }
            });
        },
        ///Get available site colletion ID.
        siteCollection: function (options) {
            if (typeof (options) == "undefined") {
                options = {
                    path: that.utility.path().reverse(),
                    index: 0,
                    values: [],
                    webUrls: []
                };
            }
            if (options.index < options.path.length) {
                that.service.siteCollection({ path: options.path[options.index] }, function (result) {
                    if (result.status == app.status.succeeded) {
                        if (typeof (result.data.siteCollection) != "undefined") {
                            options.values.push(result.data.id);
                            options.webUrls.push(result.data.webUrl);
                        }
                        options.index++;
                        that.browse.siteCollection(options);
                    }
                    else {
                        if (result.error.status == 401) {
                            that.browse.popup.message("Access denied.");
                        }
                        else {
                            options.index++;
                            that.browse.siteCollection(options);
                        }
                    }
                });
            }
            else {
                if (options.values.length > 0) {
                    that.browse.sites({ siteId: options.values.shift(), siteUrl: options.webUrls.shift() });
                }
                else {
                    that.browse.popup.message("Get site collection ID failed.");
                }
            }
        },
        ///Get all subsites under current site collection.
        sites: function (options) {
            that.service.sites(options, function (result) {
                if (result.status == app.status.succeeded) {
                    var _s = [];
                    $.each(result.data.value, function (i, d) {
                        _s.push({ id: d.id, name: d.name, type: "site", siteUrl: d.webUrl });
                    });
                    that.browse.libraries({ siteId: options.siteId, siteUrl: options.siteUrl, sites: _s });
                }
                else {
                    that.browse.popup.message("Get sites failed.");
                }
            });
        },
        ///Get all document libraries under all webs.
        libraries: function (options) {
            that.service.libraries(options, function (result) {
                if (result.status == app.status.succeeded) {
                    var _l = options.sites ? options.sites : [];
                    $.each(result.data.value, function (i, d) {
                        if (d.driveType.toUpperCase() == "DocumentLibrary".toUpperCase()) {
                            _l.push({ id: d.id, name: decodeURI(d.name), type: "library", siteId: options.siteId, siteUrl: options.siteUrl, url: d.webUrl });
                        }
                    });
                    that.browse.display({ data: _l });
                }
                else {
                    that.browse.popup.message("Get libraries failed.");
                }
            });
        },
        ///Get all folders or excel files under the document library.
        items: function (options) {
            if (options.inFolder) {
                that.service.itemsInFolder(options, function (result) {
                    if (result.status == app.status.succeeded) {
                        var _fd = [], _fi = [];
                        $.each(result.data.value, function (i, d) {
                            var _u = d.webUrl, _n = d.name, _nu = decodeURI(_n);
                            if (d.folder) {
                                _fd.push({ id: d.id, name: _nu, type: "folder", url: _u, siteId: options.siteId, siteUrl: options.siteUrl, listId: options.listId, listName: options.listName });
                            }
                            else if (d.file) {
                                if (_n.toUpperCase().indexOf(".XLSX") > 0) {
                                    _fi.push({ id: d.id, name: _nu, type: "file", url: _u, siteId: options.siteId, siteUrl: options.siteUrl, listId: options.listId, listName: options.listName });
                                }
                            }
                        });
                        _fi.sort(function (_a, _b) {
                            return (_a.name.toUpperCase() > _b.name.toUpperCase()) ? 1 : (_a.name.toUpperCase() < _b.name.toUpperCase()) ? -1 : 0;
                        });
                        that.browse.display({ data: _fd.concat(_fi) });
                    }
                    else {
                        that.browse.popup.message("Get files failed.");
                    }
                });
            }
            else {
                that.service.items(options, function (result) {
                    if (result.status == app.status.succeeded) {
                        var _fd = [], _fi = [];
                        $.each(result.data.value, function (i, d) {
                            var _u = d.webUrl, _n = d.name, _nu = decodeURI(_n);
                            if (d.folder) {
                                _fd.push({ id: d.id, name: _nu, type: "folder", url: _u, siteId: options.siteId, siteUrl: options.siteUrl, listId: options.listId, listName: options.listName });
                            }
                            else if (d.file) {
                                if (_n.toUpperCase().indexOf(".XLSX") > 0) {
                                    _fi.push({ id: d.id, name: _nu, type: "file", url: _u, siteId: options.siteId, siteUrl: options.siteUrl, listId: options.listId, listName: options.listName });
                                }
                            }
                        });
                        _fi.sort(function (_a, _b) {
                            return (_a.name.toUpperCase() > _b.name.toUpperCase()) ? 1 : (_a.name.toUpperCase() < _b.name.toUpperCase()) ? -1 : 0;
                        });
                        that.browse.display({ data: _fd.concat(_fi) });
                    }
                    else {
                        that.browse.popup.message("Get files failed.");
                    }
                });
            }
        },
        ///Get document id by file name.
        file: function (options) {
            that.service.item(options, function (result) {
                if (result.status == app.status.succeeded) {
                    var _d = "";
                    $.each(result.data.value, function (i, d) {
                        if (decodeURI(options.url).toUpperCase() == decodeURI(d.EncodedAbsUrl).toUpperCase() && d.OData__dlc_DocId) {
                            _d = d.OData__dlc_DocId;
                            return false;
                        }
                    });
                    if (_d != "") {
                        that.browse.popup.hide();
                        that.selectDocument = { Id: _d, Name: options.fileName };
                        that.list();
                    }
                    else {
                        that.browse.popup.message("Get file Document ID failed.");
                    }
                }
                else {
                    that.browse.popup.message("Get file Document ID failed.");
                }
            });
        },
        ///Display site/library/folder/file in the file explorer popup.
        display: function (options) {
            that.controls.popupBrowseList.html("");
            $.each(options.data, function (i, d) {
                var _h = "";
                if (d.type == "site") {
                    _h = '<li class="i-site" data-id="' + d.id + '" data-type="site" data-siteurl="' + d.siteUrl + '">' + d.name + '</li>';
                }
                else if (d.type == "library") {
                    _h = '<li class="i-library" data-id="' + d.id + '" data-site="' + d.siteId + '" data-url="' + d.url + '" data-type="library" data-siteurl="' + d.siteUrl + '" data-listname="' + d.name + '">' + d.name + '</li>';
                }
                else if (d.type == "folder") {
                    _h = '<li class="i-folder" data-id="' + d.id + '" data-site="' + d.siteId + '" data-list="' + d.listId + '" data-url="' + d.url + '" data-type="folder" data-siteurl="' + d.siteUrl + '" data-listname="' + d.listName + '">' + d.name + '</li>';
                }
                else if (d.type == "file") {
                    _h = '<li class="i-file" data-id="' + d.id + '" data-site="' + d.siteId + '" data-list="' + d.listId + '" data-url="' + d.url + '" data-type="file" data-siteurl="' + d.siteUrl + '" data-listname="' + d.listName + '">' + d.name + '</li>';
                }
                that.controls.popupBrowseList.append(_h);
            });
            if (options.data.length == 0) {
                that.controls.popupBrowseList.html("No items found.");
            }
            that.browse.popup.processing(false);
        },
        ///Navigate the file explorer.
        select: function (elem) {
            var _t = $(elem).data("type");
            if (_t == "site") {
                that.browse.path.push({ type: "site", id: $(elem).data("id"), siteUrl: $(elem).data("siteurl") });
                that.browse.popup.nav();
                that.browse.popup.processing(true);
                that.browse.sites({ siteId: $(elem).data("id"), siteUrl: $(elem).data("siteurl") });
            }
            else if (_t == "library") {
                that.browse.path.push({ type: "library", id: $(elem).data("id"), site: $(elem).data("site"), url: $(elem).data("url"), siteUrl: $(elem).data("siteurl"), listName: $(elem).data("listname") });
                that.browse.popup.nav();
                that.browse.popup.processing(true);
                that.browse.items({ inFolder: false, siteId: $(elem).data("site"), siteUrl: $(elem).data("siteurl"), listId: $(elem).data("id"), listName: $(elem).data("listname") });
            }
            else if (_t == "folder") {
                that.browse.path.push({ type: "folder", id: $(elem).data("id"), site: $(elem).data("site"), siteUrl: $(elem).data("siteurl"), list: $(elem).data("list"), url: $(elem).data("url"), listName: $(elem).data("listname") });
                that.browse.popup.nav();
                that.browse.popup.processing(true);
                that.browse.items({ inFolder: true, siteId: $(elem).data("site"), siteUrl: $(elem).data("siteurl"), listId: $(elem).data("list"), listName: $(elem).data("listname"), itemId: $(elem).data("id") });
            }
            else {
                that.browse.popup.processing(true);
                that.browse.file({ siteUrl: $(elem).data("siteurl"), listName: $(elem).data("listname"), name: $(elem).text(), url: that.browse.path[that.browse.path.length - 1].url + "/" + encodeURI($(elem).text()), fileName: $.trim($(elem).text()) });
            }
        },
        ///Display the file explorer popup.
        popup: {
            ///Reset to popup default state (an empty popup)
            dft: function () {
                that.controls.popupBrowseList.html("");
                that.controls.popupBrowseBack.hide();
                that.controls.popupBrowseMessage.html("").hide();
                that.controls.popupBrowseLoading.hide();
            },
            ///Show the file explorer popup.
            show: function () {
                that.controls.popupMain.removeClass("message process confirm").addClass("active browse");
            },
            ///Hide the explorer popup
            hide: function () {
                that.controls.popupMain.removeClass("active message process confirm browse");
            },
            ///Display the loading before popup displays.
            processing: function (show) {
                if (show) {
                    that.controls.popupBrowseLoading.show();
                }
                else {
                    that.controls.popupBrowseLoading.hide();
                }
            },
            ///Display the prompted message in the popup.
            message: function (txt) {
                that.controls.popupBrowseMessage.html(txt).show();
                that.browse.popup.processing(false);
            },
            ///Go back
            back: function () {
                that.browse.path.pop();
                if (that.browse.path.length > 0) {
                    var _ip = that.browse.path[that.browse.path.length - 1];
                    if (_ip.type == "site") {
                        that.browse.popup.processing(true);
                        that.browse.sites({ siteId: _ip.id, siteUrl: _ip.siteUrl });
                    }
                    else if (_ip.type == "library") {
                        that.browse.popup.processing(true);
                        that.browse.items({ inFolder: false, siteId: _ip.site, listId: _ip.id, siteUrl: _ip.siteUrl, listName: _ip.listName });
                    }
                    else if (_ip.type == "folder") {
                        that.browse.popup.processing(true);
                        that.browse.items({ inFolder: true, siteId: _ip.site, listId: _ip.list, itemId: _ip.id, siteUrl: _ip.siteUrl, listName: _ip.listName });
                    }
                }
                else {
                    that.browse.popup.processing(true);
                    that.browse.siteCollection();
                }
                that.browse.popup.nav();
            },
            ///Display the back button or not.
            nav: function () {
                that.browse.path.length > 0 ? that.controls.popupBrowseBack.show() : that.controls.popupBrowseBack.hide();
            }
        }
    };
    ///Define all popup features.
    that.popup = {
        ///Display the message popup.
        message: function (options, callback) {
            if (options.success) {
                that.controls.popupMessage.removeClass("error").addClass("success");
                that.controls.popupSuccessMessage.html(options.title);
            }
            else {
                that.controls.popupMessage.removeClass("success").addClass("error");
                that.controls.popupErrorTitle.html(options.title ? options.title : "");
                var _s = "error-single";
                if (options.values) {
                    _s = "error-list";
                    var _h = "";
                    $.each(options.values, function (i, d) {
                        _h += "<li>";
                        $.each(d, function (m, n) {
                            _h += "<span>" + n + "</span>";
                        });
                        _h += "</li>";
                    });
                    that.controls.popupErrorMessage.html(_h);
                    if (options.repair) {
                        _s = "error-repair";
                        that.controls.popupErrorRepair.html(options.repair);
                    }
                }
                that.controls.popupErrorMain.removeClass("error-single error-list").addClass(_s);
            }
            that.controls.popupMain.removeClass("process confirm browse").addClass("active message");
            if (options.success) {
                callback();
            }
            else {
                that.controls.popupErrorOK.unbind("click").click(function () {
                    that.action.ok();
                    if (callback) {
                        callback();
                    }
                });
            }
        },
        ///Display the process popup.
        processing: function (show) {
            if (!show) {
                that.controls.popupMain.removeClass("active process");
            }
            else {
                that.controls.popupMain.removeClass("message confirm").addClass("active process");
            }
        },
        ///Diplay the file explorer.
        browse: function (show) {
            if (!show) {
                that.controls.popupMain.removeClass("active browse");
            }
            else {
                that.controls.popupMain.removeClass("message process confirm").addClass("active browse");
            }
        },
        ///Hide the popup.
        hide: function (millisecond) {
            if (millisecond) {
                setTimeout(function () {
                    that.controls.popupMain.removeClass("active message");
                }, millisecond);
            } else {
                that.controls.popupMain.removeClass("active message");
            }
        },
        ///Go to source management page.
        back: function (millisecond) {
            if (millisecond) {
                setTimeout(function () {
                    that.controls.popupMain.removeClass("active message");
                    that.controls.main.removeClass("manage add").addClass("manage");
                }, millisecond);
            }
            else {
                that.controls.popupMain.removeClass("active message");
                that.controls.main.removeClass("manage add").addClass("manage");
            }
        }
    };
    ///Default actions to interact with service.
    that.service = {
        ///Define a general request sent to service end and return the callback result.
        common: function (options, callback) {
            $.ajax({
                url: options.url,
                type: options.type,
                cache: false,
                data: options.data ? options.data : "",
                dataType: options.dataType,
                headers: options.headers ? options.headers : "",
                success: function (data) {
                    callback({ status: app.status.succeeded, data: data });
                },
                error: function (error) {
                    if (error.status == 410) {
                        that.popup.message({ success: false, title: "The current login gets expired and needs re-authenticate. You will be redirected to the login page by click OK." }, function () {
                            window.location = "../../Home/Index";
                        });
                    }
                    else {
                        callback({ status: app.status.failed, error: error });
                    }
                }
            });
        },
        ///Get source point by catalog document id.
        catalog: function (options, callback) {
            that.service.common({ url: that.endpoints.catalog + options.documentId, type: "GET", dataType: "json" }, callback);
        },
        ///Get graph or sharepoint token.
        token: function (options, callback) {
            that.service.common({ url: options.endpoint, type: "GET", dataType: "json" }, callback);
        },
        ///Get site collection id. 
        siteCollection: function (options, callback) {
            that.service.common({ url: that.endpoints.graph + "/sites/" + that.api.host + ":/" + options.path, type: "GET", dataType: "json", headers: { "authorization": "Bearer " + that.api.token } }, callback);
        },
        ///Get all subsites under the site collection.
        sites: function (options, callback) {
            that.service.common({ url: that.endpoints.graph + "/sites/" + options.siteId + "/sites", type: "GET", dataType: "json", headers: { "authorization": "Bearer " + that.api.token } }, callback);
        },
        ///Get libraries under the current site.
        libraries: function (options, callback) {
            that.service.common({ url: that.endpoints.graph + "/sites/" + options.siteId + "/drives", type: "GET", dataType: "json", headers: { "authorization": "Bearer " + that.api.token } }, callback);
        },
        ///Get items under the selected library.
        items: function (options, callback) {
            that.service.common({ url: that.endpoints.graph + "/sites/" + options.siteId + "/drives/" + options.listId + "/root/children", type: "GET", dataType: "json", headers: { "authorization": "Bearer " + that.api.token } }, callback);
        },
        ///Get items under the selected folder.
        itemsInFolder: function (options, callback) {
            that.service.common({ url: that.endpoints.graph + "/sites/" + options.siteId + "/drives/" + options.listId + "/items/" + options.itemId + "/children", type: "GET", dataType: "json", headers: { "authorization": "Bearer " + that.api.token } }, callback);
        },
        ///Search list item by file name.
        item: function (options, callback) {
            that.service.common({ url: options.siteUrl + "/_api/web/lists/getbytitle('" + options.listName + "')/items?$select=FileLeafRef,EncodedAbsUrl,OData__dlc_DocId&$filter=FileLeafRef eq '" + options.fileName + "'", type: "GET", dataType: "json", headers: { "authorization": "Bearer " + that.api.sharePointToken } }, callback);
        }
    };
    ///Build the HTML/UI.
    that.ui = {
        ///Set the default value for file & eyword textboxes.
        dft: function () {
            var _f = $.trim(that.controls.file.val()), _fd = that.controls.file.data("default"), _k = $.trim(that.controls.keyword.val()), _kd = that.controls.keyword.data("default");
            if (_f == "" || _f == _fd) {
                that.controls.file.val(_fd);
            }
            if (_k == "" || _k == _kd) {
                that.controls.keyword.val(_kd).addClass("input-default");
            }
        },
        ///Build the source points list in management page.
        list: function (options) {
            var _dt = $.extend([], that.points), _d = [], _ss = [];
            if (that.sourcePointKeyword != "") {
                var _sk = app.search.splitKeyword({ keyword: that.sourcePointKeyword });
                if (_sk.length > 26) {
                    that.popup.message({ success: false, title: "Only support less then 26 keywords." });
                }
                else {
                    $.each(_dt, function (i, d) {
                        if (app.search.weight({ keyword: _sk, source: d.Name }) > 0) {
                            _d.push(d);
                        }
                    });
                }
            }
            else {
                _d = _dt;
            }
            that.utility.pager.status({ length: _d.length });
            _d.sort(function (_a, _b) {
                return (_a.Name.toUpperCase() > _b.Name.toUpperCase()) ? 1 : (_a.Name.toUpperCase() < _b.Name.toUpperCase()) ? -1 : 0;
            });
            that.controls.list.find(".point-item").remove();
            that.ui.item({ index: 0, data: _d, selected: _ss });
        },
        ///Build each source point HTML in management page.
        item: function (options, callback) {
            if (options.index < options.data.length) {
                var _item = options.data[options.index];
                if (options.index >= that.pagerSize * (that.pagerIndex - 1) && options.index < that.pagerSize * that.pagerIndex) {
                    var _pn = that.utility.position(_item.NamePosition),
                        _pv = (_item.PublishedHistories && _item.PublishedHistories.length > 0 ? (_item.PublishedHistories[0].Value ? _item.PublishedHistories[0].Value : "") : ""),
                        _pht = that.utility.publishHistory({ data: _item.PublishedHistories });
                    var _h = '<li class="ms-ListItem point-item" data-id="' + _item.Id + '" data-range="' + _item.RangeId + '" data-position="' + _item.Position + '" data-namerange="' + _item.NameRangeId + '" data-nameposition="' + _item.NamePosition + '">';
                    _h += '<div class="point-item-line">';
                    _h += '<div class="i2"><span class="s-name" title="' + _item.Name + '">' + _item.Name + '</span>';
                    _h += '<span title="' + (_pn.sheet ? _pn.sheet : "") + ':[' + (_pn.cell ? _pn.cell : "") + ']"><strong>' + (_pn.sheet ? _pn.sheet : "") + ':</strong>[' + (_pn.cell ? _pn.cell : "") + ']</span>';
                    _h += '</div>';
                    _h += '<div class="i3" title="' + _pv + '">' + _pv + '</div>';
                    _h += '<div class="i5"><div class="i-line"><i title="History" class="ms-Icon ms-Icon--History ms-fontColor-themePrimary i-history"></i></div>';
                    _h += '<div class="i-menu"><a href="javascript:"><span title="Action">...</span><span><i title="History" class="ms-Icon ms-Icon--History ms-fontColor-themePrimary i-history"></i></span></a></div>';
                    _h += '</div>';
                    _h += '</div>';
                    _h += '<div class="item-history"><h6>Publish History</h6><ul class="ms-List history-list">';
                    _h += '<li class="ms-ListItem history-header"><div class="h1">Name</div><div class="h2">Value</div><div class="h3">Date</div></li>';
                    $.each(_pht, function (m, n) {
                        _h += '<li class="ms-ListItem history-item"><div class="h1" title="' + n.PublishedUser + '">' + n.PublishedUser + '</div><div class="h2" title="' + (n.Value ? n.Value : "") + '">' + (n.Value ? n.Value : "") + '</div><div class="h3" title="' + that.utility.date(n.PublishedDate) + '">' + that.utility.date(n.PublishedDate) + '</div></li>';
                    });
                    _h += '</ul>';
                    _h += '</div>';
                    _h += '</li>';
                    that.controls.list.append(_h);
                }
                options.index++;
                that.ui.item(options);
            }
            else {
                that.controls.main.get(0).scrollTop = 0;
            }
        }
    };

    return that;
})();