$(function () {
    Office.initialize = function (reason) {
        $(document).ready(function () {
            BigNumber.config({ EXPONENTIAL_AT: 1e+9 });
            point.init();
        });
    };
});

var point = (function () {
    var point = {
        ///The URL of current open excel document.
        filePath: "",
        ///The document id of the current open excel document.
        documentId: "",
        ///Define all UI controls.
        controls: {},
        ///The selected source point catalog.
        file: null,
        ///The selected source point.
        selected: null,
        ///The selected destination point.
        model: null,
        ///Define source point groups.
        groups: [],
        ///Define the destination points. 
        points: [],
        ///The keyword in add destination point page.
        keyword: "",
        ///The keyword in destination point list page.
        sourcePointKeyword: "",
        ///The background color for the highlighted destination ponit.
        highlightColor: "#66FF00",
        ///Highlight the destination point or not
        highlighted: false,
        ///Determine if it is first time load page or not.
        firstLoad: true,
        ///Define default page index 
        pagerIndex: 0,
        ///Define default page size.
        pagerSize: 30,
        ///Default default page total count.
        pagerCount: 0,
        ///Define all service endpoints.
        endpoints: {
            add: "/api/DestinationPoint",
            catalog: "/api/SourcePointCatalog?documentId=",
            groups: "/api/SourcePointGroup",
            list: "/api/DestinationPointCatalog?name=",
            del: "/api/DestinationPoint?id=",
            deleteSelected: "/api/DeleteSelectedDestinationPoint",
            token: "/api/GraphAccessToken",
            sharePointToken: "/api/SharePointAccessToken",
            graph: "https://graph.microsoft.com/v1.0",
            customFormat: "/api/CustomFormats",
            updateCustomFormat: "/api/UpdateDestinationPointCustomFormat"
        },
        ///Define the api host and token.
        api: {
            host: "",
            token: "",
            sharePointToken: ""
        }
    }, that = point;
    ///Initialize the event handlers & load the destination point list.
    that.init = function () {
        ///Get the document URL.
        that.filePath = window.location.href.indexOf("localhost") > -1 ? "https://cand3.sharepoint.com/Shared%20Documents/Test.docx" : Office.context.document.url;
        that.controls = {
            body: $("body"),
            main: $(".main"),
            back: $(".n-back"),
            add: $(".n-add"),
            highlight: $(".n-highlight"),
            refresh: $(".n-refresh"),
            del: $(".n-delete"),
            next: $("#btnNext"),
            cancel: $("#btnCancel"),
            save: $("#btnAdd"),
            update: $("#btnSave"),
            file: $("#txtFile"),
            fileTrigger: $("#btnOpenBrowse"),
            keyword: $("#txtKeyword"),
            search: $("#iSearch"),
            autoCompleteControl: $("#autoCompleteWrap"),
            filterMain: $(".point-filter"),
            filterTrigger: $(".filter-header span"),
            filterList: $("#filterList"),
            resultList: $("#resultList"),
            resultNotFound: $("#resultNotFound"),
            selectedName: $("#selectedName"),
            selectedFile: $("#selectedFile"),
            stepFirstMain: $(".add-point-first"),
            stepSecondMain: $(".add-point-second"),
            sourcePointName: $("#txtSearchSourcePoint"),
            searchSourcePoint: $("#iSearchSourcePoint"),
            autoCompleteControl2: $("#autoCompleteWrap2"),
            list: $("#listPoints"),
            documentIdError: $("#lblDocumentIDError"),
            documentIdReload: $("#btnDocumentIDReload"),
            headerListPoints: $("#headerListPoints"),
            moveUp: $("#btnMoveUp"),
            moveDown: $("#btnMoveDown"),
            tooltipMessage: $("#tlpMessage"),
            tooltipClose: $("#tlpClose"),
            formatBtn: $("#btnSelectFormat"),
            formatIcon: $("#iconSelectFormat"),
            formatList: $("#listFormats"),
            previewValue: $("#lbPreviewValue"),
            decimalIncrease: $(".i-increase"),
            decimalDecrease: $(".i-decrease"),
            addCustomFormat: $("#addCustomFormat"),
            popupMain: $("#popupMain"),
            popupErrorOK: $("#btnErrorOK"),
            popupMessage: $("#popupMessage"),
            popupProcessing: $("#popupProcessing"),
            popupSuccessMessage: $("#lblSuccessMessage"),
            popupErrorMain: $("#popupErrorMain"),
            popupErrorTitle: $("#lblErrorTitle"),
            popupErrorMessage: $("#lblErrorMessage"),
            popupErrorRepair: $("#lblErrorRepair"),
            popupConfirm: $("#popupConfirm"),
            popupConfirmTitle: $("#lblConfirmTitle"),
            popupConfirmYes: $("#btnYes"),
            popupConfirmNo: $("#btnNo"),
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
        ///Define the event handlers.
        that.highlighted = that.utility.highlight.get();
        that.controls.highlight.find("span").html(that.highlighted ? "Highlight Off" : "Highlight On");
        that.controls.highlight.prop("title", that.highlighted ? "Highlight Off" : "Highlight On");
        that.controls.body.click(function () {
            that.action.body();
        });
        that.controls.back.click(function () {
            that.action.back();
        });
        that.controls.add.click(function () {
            that.action.add();
        });
        that.controls.highlight.click(function () {
            that.action.highlightAll();
        });
        that.controls.refresh.click(function () {
            that.action.refresh();
        });
        that.controls.del.click(function () {
            that.action.deleteSelected();
        });
        that.controls.next.click(function () {
            that.action.next();
        });
        that.controls.cancel.click(function () {
            that.action.cancel();
        });
        that.controls.save.click(function () {
            that.controls.save.blur();
            that.action.save();
        });
        that.controls.update.click(function () {
            that.controls.update.blur();
            that.action.update();
        });
        that.controls.file.click(function () {
            that.browse.init();
        });
        that.controls.fileTrigger.click(function () {
            that.browse.init();
        });
        that.controls.keyword.focus(function () {
            that.action.dft(this, true);
        });
        that.controls.keyword.blur(function () {
            that.action.dft(this, false);
        });
        that.controls.keyword.keydown(function (e) {
            if (e.keyCode == 13) {
                that.controls.search.click();
            }
            else if (e.keyCode == 38 || e.keyCode == 40) {
                app.search.move({ result: that.controls.autoCompleteControl, target: that.controls.keyword, down: e.keyCode == 40 });
            }
        });
        that.controls.keyword.bind("input", function (e) {
            that.action.autoComplete();
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
        that.controls.search.click(function () {
            that.action.search();
        });
        that.controls.filterTrigger.click(function () {
            that.action.filter();
        });
        that.controls.resultList.on("click", "li", function () {
            that.action.choose($(this));
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
        that.controls.moveUp.click(function () {
            that.action.up();
        });
        that.controls.moveDown.click(function () {
            that.action.down();
        });
        that.controls.tooltipClose.click(function () {
            that.controls.tooltipMessage.removeClass("active");
        });
        that.controls.formatBtn.click(function () {
            that.utility.formatHeight();
            that.controls.formatList.hasClass("active") ? that.controls.formatList.removeClass("active") : that.controls.formatList.addClass("active");
            return false;
        });
        that.controls.formatIcon.click(function () {
            that.utility.formatHeight();
            that.controls.formatList.hasClass("active") ? that.controls.formatList.removeClass("active") : that.controls.formatList.addClass("active");
            return false;
        });
        that.controls.formatList.on("click", "ul > li", function () {
            var _ck = $(this).hasClass("checked"), _sg = $(this).closest(".drp-radio").length > 0, _cn = $(this).data("name");
            if (_sg) {
                $(this).closest("ul").find("li").removeClass("checked")
            }
            _ck ? $(this).removeClass("checked") : $(this).addClass("checked");
            if (_cn == "ConvertToThousands" || _cn == "ConvertToMillions" || _cn == "ConvertToBillions" || _cn == "ConvertToHundreds") {
                that.controls.formatList.removeClass("convert1 convert2 convert3 convert4");
                that.controls.formatList.find(".drp-descriptor li.checked").removeClass("checked");
                if (!_ck) {
                    var _tn = _cn == "ConvertToThousands" ? "IncludeThousandDescriptor" : (_cn == "ConvertToMillions" ? "IncludeMillionDescriptor" : (_cn == "ConvertToBillions" ? "IncludeBillionDescriptor" : (_cn == "ConvertToHundreds" ? "IncludeHundredDescriptor" : "")));
                    var _cl = _cn == "ConvertToThousands" ? "convert2" : (_cn == "ConvertToMillions" ? "convert3" : (_cn == "ConvertToBillions" ? "convert4" : (_cn == "ConvertToHundreds" ? "convert1" : "")));
                    that.controls.formatList.addClass(_cl);
                    that.controls.formatList.find("ul > li[data-name=" + _tn + "]").addClass("checked");
                }
            }
            that.action.selectedFormats($(this));
            return false;
        });
        that.controls.decimalIncrease.click(function () {
            var _p = that.controls.formatBtn.prop("place"), _v = that.format.remove(that.controls.previewValue.text());
            if (_p == "") {
                _p = that.format.getDecimalLength(_v);
            }
            _p = parseInt(_p);
            that.controls.formatBtn.prop("place", ++_p);
            that.format.preview();
        });
        that.controls.decimalDecrease.click(function () {
            var _p = that.controls.formatBtn.prop("place"), _v = that.format.remove(that.controls.previewValue.text());
            if (_p == "") {
                _p = that.format.getDecimalLength(_v);
            }
            _p = parseInt(_p);
            if (_p > 0) {
                that.controls.formatBtn.prop("place", --_p);
                that.format.preview();
            }
        });
        that.controls.documentIdReload.click(function () {
            window.location.reload();
        });

        that.controls.list.on("click", ".i-edit", function () {
            that.action.edit($(this).closest(".point-item"));
            return false;
        });
        that.controls.list.on("click", ".i-history", function () {
            that.action.history($(this).closest(".point-item"));
            return false;
        });
        that.controls.list.on("click", ".i-delete", function () {
            that.action.del($(this).closest(".point-item"));
            return false;
        });
        that.controls.list.on("click", ".point-item .i2, .point-item .i3, .point-item .i5, .point-item .error-info, .point-item .item-history, .point-item .item-format", function (e) {
            that.action.goto($(this).closest(".point-item"));
        });
        that.controls.main.on("change", ".ckb-wrapper input", function (e) {
            if ($(this).get(0).checked) {
                $(this).closest(".ckb-wrapper").addClass("checked");
                if ($(this).closest(".all").length > 0) {
                    that.controls.list.find("input").prop("checked", true);
                    that.controls.list.find(".ckb-wrapper").addClass("checked");
                }
            }
            else {
                $(this).closest(".ckb-wrapper").removeClass("checked");
                if ($(this).closest(".all").length > 0) {
                    that.controls.list.find("input").prop("checked", false);
                    that.controls.list.find(".ckb-wrapper").removeClass("checked");
                }
            }
        });
        that.controls.list.on("click", ".i-file", function () {
            that.action.open($(this));
            return false;
        });
        that.controls.main.on("click", "#filterList .ckb-wrapper input", function () {
            that.action.checked($(this));
        });
        that.controls.main.on("click", ".search-tooltips li", function () {
            $(this).parent().parent().find("input").val($(this).text());
            $(this).parent().hide();
            if ($(this).closest(".input-search").length > 0) {
                that.controls.search.click();
            }
            else {
                that.controls.searchSourcePoint.click();
            }
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
                    that.utility.pager.init({ index: _n, refresh: false });
                }
                else {
                    that.popup.message({ success: false, title: "Invalid number." });
                }
            }
        });
        $(window).resize(function () {
            that.utility.height();
            that.utility.formatHeight();
        });
        that.utility.height();
        that.action.dft(that.controls.sourcePointName, false);
        ///Retrieve the document ID via document URL
        that.document.init(function () {
            ///Load all destination points in management page.
            that.list({ refresh: false, index: 1 }, function (result) {
                if (result.status == app.status.failed) {
                    ///Dipslay the error message if failed to get destination point list.
                    that.popup.message({ success: false, title: result.error.statusText });
                }
            });
        });
    };
    ///Load the destination point list.
    that.list = function (options, callback) {
        that.popup.processing(true);
        that.service.list(function (result) {
            that.popup.processing(false);
            if (result.status == app.status.succeeded) {
                if (result.data) {
                    that.points = result.data.DestinationPoints;
                    that.utility.pager.init({ refresh: options.refresh, index: options.index });
                    callback({ status: result.status });
                }
                else {
                    that.utility.pager.status({ length: 0 });
                }
            }
            else {
                callback({ status: result.status, error: result.error });
            }
        });
    };
    ///Default display when add destination point.
    that.default = function () {
        if (that.groups.length == 0) {
            that.popup.processing(true);
            that.service.groups(function (result) {
                if (result.status == app.status.succeeded) {
                    that.popup.processing(false);
                    that.groups = result.data;
                    that.ui.groups(that.groups);
                    that.ui.reset();
                    that.controls.main.removeClass("manage").addClass("add");
                }
                else {
                    that.popup.message({ success: false, title: "Load groups failed." });
                }
            });
        }
        else {
            that.ui.groups(that.groups);
            that.ui.reset();
            that.controls.main.removeClass("manage").addClass("add");
        }
    };
    ///Load source points after selecting the file in the file explorer.
    that.files = function (options) {
        that.popup.processing(true);
        that.service.catalog({ documentId: options.documentId }, function (result) {
            if (result.status == app.status.succeeded) {
                that.popup.processing(false);
                that.file = result.data;
                that.ui.reset();
                that.action.select(options);
            }
            else {
                that.popup.message({ success: false, title: "Load source points failed." });
            }
        });
    };
    ///The utility methods.
    that.utility = {
        ///Get current destination point.
        model: function (id) {
            var _m = null;
            if (id && that.points) {
                var _t = [];
                $.each(that.points, function (i, d) {
                    _t.push(d.Id);
                });
                if ($.inArray(id, _t) > -1) {
                    _m = that.points[$.inArray(id, _t)];
                }
            }
            return _m;
        },
        ///Display as 0n.
        format: function (n) {
            return n > 9 ? n : ("0" + n);
        },
        ///Display AM/PM and convert to PST.
        date: function (str) {
            var _v = new Date(str), _d = _v.getDate(), _m = _v.getMonth() + 1, _y = _v.getFullYear(), _h = _v.getHours(), _mm = _v.getMinutes(), _a = _h < 12 ? " AM" : " PM";
            return that.utility.format(_m) + "/" + that.utility.format(_d) + "/" + _y + " " + (_h < 12 ? (_h == 0 ? "12" : that.utility.format(_h)) : (_h == 12 ? _h : _h - 12)) + ":" + that.utility.format(_mm) + "" + _a + " PST";
        },
        ///Determine if it is edit mode or not.
        mode: function (callback) {
            callback();
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
        ///Get the value of the fields (file, keyword, group)
        entered: function () {
            var file = $.trim(that.controls.file.val()), fd = that.controls.file.data("default"), keyword = $.trim(that.controls.keyword.val()), kd = that.controls.keyword.data("default"), groups = [];
            that.controls.filterList.find("input").each(function (i, d) {
                if (d.checked) {
                    groups.push(parseInt(d.value));
                }
            });
            return { file: file != fd ? file : "", keyword: keyword != kd ? keyword : "", groups: groups };
        },
        ///Get the index of current destination point.
        index: function (options) {
            var _i = -1;
            $.each(that.points, function (m, n) {
                if (n.Id == options.Id) {
                    _i = m;
                    return false;
                }
            });
            return _i;
        },
        ///Update the destination point in the array.
        update: function (options) {
            that.points[that.utility.index(options)] = options;
        },
        ///Add destination point to the points array.
        add: function (options) {
            that.points.push(options);
        },
        ///Remove destination point from the points array.
        remove: function (options) {
            that.points.splice(that.utility.index(options), 1);
        },
        ///If the one of the source point groups is in the selected groups, then return true. If there is no selected groups, then return true.
        contain: function (options) {
            var _f = false;
            if (options.selected.length == 0) {
                _f = true;
            }
            else {
                $.each(options.groups, function (i, d) {
                    if ($.inArray(d.Id, options.selected) > -1) {
                        _f = true;
                        return false;
                    }
                });
            }
            return _f;
        },
        ///Get the file name.
        fileName: function (path) {
            return path.lastIndexOf("/") > -1 ? path.substr(path.lastIndexOf("/") + 1) : (path.lastIndexOf("\\") > -1 ? path.substr(path.lastIndexOf("\\") + 1) : path);
        },
        ///Get/set the highlighted destination points.
        highlight: {
            get: function () {
                try {
                    var _d = Office.context.document.settings.get("HighLightPoint");
                    return _d && _d != null ? _d : false;
                } catch (e) {
                    return false;
                }
            },
            set: function (options, callback) {
                try {
                    Office.context.document.settings.set("HighLightPoint", options);
                    Office.context.document.settings.saveAsync({ overwriteIfStale: true }, function (asyncResult) {
                        callback({ status: asyncResult.status == Office.AsyncResultStatus.Succeeded ? app.status.succeeded : app.status.failed });
                    });
                } catch (e) {
                    callback({ status: app.status.failed });
                }
            }
        },
        //Convert the val tostring and remove the blank space at the start or end of the word.
        toString: function (val) {
            if (val != undefined && val != null) {
                return $.trim(val.toString());
            }
            return "";
        },
        ///Paging feature for the destination point list.
        pager: {
            ///Initialize the destination point list UI.
            init: function (options) {
                that.controls.pagerValue.val("");
                that.pagerIndex = options.index ? options.index : 1;
                that.ui.list({ refresh: options.refresh });
            },
            ///Go to prev page.
            prev: function () {
                that.controls.pagerValue.val("");
                that.pagerIndex--;
                that.ui.list({ refresh: false });
            },
            ///Go to next page.
            next: function () {
                that.controls.pagerValue.val("");
                that.pagerIndex++;
                that.ui.list({ refresh: false });
            },
            ///Get the status of paging.
            status: function (options) {
                that.pagerCount = Math.ceil(options.length / that.pagerSize);
                that.controls.pagerTotal.html(options.length);
                that.controls.pagerPages.html(that.pagerCount);
                that.controls.pagerCurrent.html(that.pagerIndex);
                that.pagerIndex == 1 || that.pagerCount == 0 ? that.controls.pagerPrev.addClass("disabled") : that.controls.pagerPrev.removeClass("disabled");
                that.pagerIndex == that.pagerCount || that.pagerCount == 0 ? that.controls.pagerNext.addClass("disabled") : that.controls.pagerNext.removeClass("disabled");
            }
        },
        ///Sort the destination point in managment page by the conntent controls in word document.
        order: function (callback) {
            var _d = $.extend([], that.points);
            that.range.all(function (result) {
                if (result.status == app.status.succeeded) {
                    $.each(_d, function (i, d) {
                        var __i = that.range.index({ data: result.data, tag: d.RangeId });
                        _d[i].orderBy = __i > -1 ? __i : 99999;
                    });
                }
                else {
                    $.each(_d, function (i, d) {
                        _d[i].orderBy = 99999;
                    });
                }

                _d.sort(function (_a, _b) {
                    return (_a.orderBy > _b.orderBy) ? 1 : (_a.orderBy < _b.orderBy) ? -1 : 0;
                });

                callback({ status: app.status.succeeded, data: _d });
            });
        },
        ///Return current selected destination points in management page.
        selected: function () {
            var _s = [];
            that.controls.list.find(".point-item .ckb-wrapper input").each(function (i, d) {
                if ($(d).prop("checked")) {
                    var _id = $(d).closest(".point-item").data("id"), _rid = $(d).closest(".point-item").data("range");
                    _s.push({ DestinationPointId: _id, RangeId: _rid });
                }
            });
            return _s;
        },
        ///Caculate the height for the destination point list in management page or add destination page dymamically and set the scroll bar accordingly.. 
        height: function () {
            if (that.controls.main.hasClass("add")) {
                var _h = that.controls.main.outerHeight();
                var _h1 = $("#word-addin .header").outerHeight(), _h2 = 0;
                $("#word-addin .add-point .add-point-first .height-item").each(function (i, d) {
                    _h2 += $(d).outerHeight();
                });
                that.controls.resultList.css("maxHeight", ((_h - _h1 - _h2 - 90) > 0 ? (_h - _h1 - _h2 - 90) : "150") + "px");
            }
            else if (that.controls.main.hasClass("manage")) {
                var _h = that.controls.main.outerHeight();
                var _h1 = $("#pager").outerHeight();
                that.controls.list.css("maxHeight", (_h - 192 - 70 - _h1) + "px");
            }
        },
        ///Uncheck all selected destination points.
        unSelectAll: function () {
            that.controls.headerListPoints.find(".point-header .ckb-wrapper input").prop("checked", false);
            that.controls.headerListPoints.find(".point-header .ckb-wrapper").removeClass("checked");
        },
        ///Get an array of paths (server ralative URl splitted by '/' )
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
        ///Calculate the height custom format display area.
        formatHeight: function () {
            if (that.controls.main.hasClass("add")) {
                var _a = that.controls.main.outerHeight();
                var _h = that.controls.formatBtn.offset().top;
                that.controls.formatList.css("maxHeight", (_a - _h - 96) + "px");
            }
        }
    };
    ///Define all the event handlers.
    that.action = {
        ///Hide the tooltip and hide the custom format dropdown list.
        body: function () {
            $(".search-tooltips").hide();
            that.controls.formatList.removeClass("active");
        },
        ///Add destination point.
        add: function () {
            that.utility.mode(function () {
                that.file = null;
                that.selected = null;
                that.default();
            });
        },
        ///Go to destination management page.
        back: function () {
            that.controls.main.removeClass("add edit step-first step-second").addClass("manage");
        },
        ///Set the default value for the control.
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
        ///Display toggle for the filter groups 
        filter: function () {
            that.controls.filterMain.hasClass("open-filter") ? that.controls.filterMain.removeClass("open-filter") : that.controls.filterMain.addClass("open-filter");
        },
        ///Set the selected excel file name in file textbox and display the source points
        select: function (options) {
            that.controls.file.val(options.name).removeClass("input-default");
            that.controls.selectedFile.html("<strong>Source file:</strong>" + options.name + "");
            that.ui.select();
            that.selected = null;
            that.ui.sources({ data: that.file, selected: [], keyword: "" });
            that.controls.stepFirstMain.addClass("selected-file");
            that.utility.height();
        },
        ///Search via the keyword and display the source points search result.
        search: function () {
            var _e = that.utility.entered();
            that.keyword = _e.keyword;
            that.selected = null;
            that.ui.sources({ data: that.file, selected: _e.groups, keyword: _e.keyword });
        },
        ///Select the checkbox.
        checked: function (o) {
            if (o.get(0).checked) {
                o.closest(".ckb-wrapper").addClass("checked");
            }
            else {
                o.closest(".ckb-wrapper").removeClass("checked");
            }
            that.selected = null;
            that.ui.sources({ data: that.file, selected: that.utility.entered().groups, keyword: that.keyword });
        },
        ///Enable the next button after selecting the source point search result.
        choose: function (o) {
            var _i = o.data("id");
            if (!that.selected || (that.selected && that.selected.Id != _i)) {
                that.selected = { Id: _i, File: o.data("file"), Name: o.data("name"), Value: o.data("value") };
                that.controls.resultList.find("li.selected").removeClass("selected");
                o.addClass("selected");
            }
            that.controls.selectedName.html(o.data("name"));
            that.controls.formatBtn.prop("original", o.data("value"));
            that.ui.status({ next: true });
        },
        ///Display the second step when add destination point.
        next: function () {
            if (!that.controls.next.hasClass("disabled")) {
                that.ui.status({ next: false, cancel: true, save: true });
                that.controls.main.removeClass("step-first").addClass("step-second");
                that.action.customFormat();
            }
        },
        ///Refresh all destination points.
        refresh: function () {
            that.list({ refresh: true, index: that.pagerIndex }, function (result) {
                if (result.status == app.status.failed) {
                    that.popup.message({ success: false, title: result.error.statusText });
                }
                else {
                    that.controls.tooltipMessage.removeClass("active");
                    that.popup.message({ success: true, title: "Refresh all destination points succeeded." }, function () { that.popup.hide(3000); });
                }
            });
        },
        ///Display the step 1.
        cancel: function () {
            if (that.controls.main.hasClass("edit")) {
                that.action.back();
            }
            else {
                that.ui.status({ next: true });
                that.controls.main.removeClass("step-second").addClass("step-first");
            }
        },
        ///Add new destination point.
        save: function () {
            that.utility.mode(function () {
                var _s = that.controls.formatBtn.prop("selected"), _n = that.controls.formatBtn.prop("name"),
                    _f = (typeof (_n) != "undefined" && _n != "") ? _n.split(",") : [],
                    _c = (typeof (_s) != "undefined" && _s != "") ? _s.split(",") : [],
                    _fa = [],
                    _x = that.controls.formatBtn.prop("place");
                $.each(_f, function (_a, _b) {
                    _fa.push({ Name: _b });
                });
                var _v = that.format.convert({ value: that.selected.Value, formats: _fa, decimal: _x });
                var _json = $.extend({}, that.selected, { RangeId: app.guid(), CatalogName: that.filePath, CustomFormatIds: _c, Value: _v, DecimalPlace: _x });

                that.range.create(_json, function (ret) {
                    if (ret.status == app.status.succeeded) {
                        that.popup.processing(true);
                        that.service.add({ data: { CatalogName: _json.CatalogName, DocumentId: that.documentId, RangeId: _json.RangeId, SourcePointId: _json.Id, CustomFormatIds: _json.CustomFormatIds, DecimalPlace: _json.DecimalPlace } }, function (result) {
                            if (result.status == app.status.succeeded) {
                                that.utility.add(result.data);
                                that.utility.pager.init({ refresh: false, index: that.pagerIndex });
                                that.popup.message({ success: true, title: "Add new destination point succeeded." }, function () { that.popup.back(3000); });
                            }
                            else {
                                that.popup.message({ success: false, title: "Add destination point failed." });
                            }
                        });
                    }
                    else {
                        that.popup.message({ success: false, title: "Create range in Word failed." });
                    }
                });
            });
        },
        ///Update destination point.
        update: function () {
            that.utility.mode(function () {
                var _s = that.controls.formatBtn.prop("selected"), _n = that.controls.formatBtn.prop("name"), _o = that.controls.formatBtn.prop("original"),
                    _f = (typeof (_n) != "undefined" && _n != "") ? _n.split(",") : [],
                    _c = (typeof (_s) != "undefined" && _s != "") ? _s.split(",") : [],
                    _fa = [],
                    _x = that.controls.formatBtn.prop("place");
                $.each(_f, function (_a, _b) {
                    _fa.push({ Name: _b });
                });
                var _v = that.format.convert({ value: _o, formats: _fa, decimal: _x });
                var _json = $.extend({}, {}, { Id: that.model.Id, RangeId: that.model.RangeId, CustomFormatIds: _c, Value: _v, DecimalPlace: _x });

                that.range.edit(_json, function (ret) {
                    if (ret.status == app.status.succeeded) {
                        that.popup.processing(true);
                        that.service.update({ data: { Id: _json.Id, CustomFormatIds: _json.CustomFormatIds, DecimalPlace: _json.DecimalPlace } }, function (result) {
                            if (result.status == app.status.succeeded) {
                                that.utility.update(result.data);
                                that.utility.pager.init({ refresh: false, index: that.pagerIndex });
                                that.popup.message({ success: true, title: "Update destination point custom format succeeded." }, function () { that.popup.back(3000); });
                            }
                            else {
                                that.popup.message({ success: false, title: "Update destination point custom format failed." });
                            }
                        });
                    }
                    else {
                        that.popup.message({ success: false, title: "Update range in Word failed." });
                    }
                });
            });
        },
        ///Highlight all destination point in word document.
        highlightAll: function () {
            that.popup.processing(true);
            that.utility.order(function (result) {
                options = { index: 0, data: result.data, errorAmount: 0, successAmount: 0 };
                that.action.highlight(options);
            });
        },
        ///Highligh one destination point in word document.
        highlight: function (options) {
            if (options.index < options.data.length) {
                that.range.highlight(options.data[options.index], function (result) {
                    if (result.status == app.status.succeeded) {
                        options.successAmount++;
                    }
                    else {
                        options.errorAmount++;
                    }
                    options.index++;
                    that.action.highlight(options);
                });
            }
            else {
                that.highlighted = !that.highlighted;
                that.utility.highlight.set(that.highlighted, function () {
                    that.controls.highlight.find("span").html(that.highlighted ? "Highlight Off" : "Highlight On");
                    that.controls.highlight.prop("title", that.highlighted ? "Highlight Off" : "Highlight On");
                    that.popup.processing(false);
                });
            }
        },
        ///Delete the destination point after clicking X icon.
        del: function (o) {
            that.utility.mode(function () {
                var _i = o.data("id"), _rid = o.data("range");
                that.popup.confirm({ title: "Do you want to delete the destination point?" }, function () {
                    that.popup.processing(true);
                    that.service.del({ Id: _i }, function (result) {
                        if (result.status == app.status.succeeded) {
                            that.range.del({ RangeId: _rid }, function (result) {
                                if (result.status == app.status.succeeded) {
                                    that.popup.message({ success: true, title: "Delete destination point succeeded." }, function () { that.popup.hide(3000); });
                                    that.utility.remove({ Id: _i });
                                    that.ui.remove({ Id: _i });
                                    that.utility.pager.init({ refresh: false, index: that.controls.list.find(".point-item").length > 0 ? that.pagerIndex : that.pagerIndex - 1 });
                                }
                                else {
                                    that.popup.message({ success: false, title: "Delete destination point in Word failed." });
                                }
                            });
                        }
                        else {
                            that.popup.message({ success: false, title: "Delete destination point failed." });
                        }
                    });
                }, function () {
                    that.controls.popupMain.removeClass("message process confirm active");
                });
            });
        },
        ///Delete selected destination point after clicking the delete button.
        deleteSelected: function () {
            var _s = that.utility.selected(), _ss = [], _sr = [];
            if (_s && _s.length > 0) {
                $.each(_s, function (_y, _z) {
                    _ss.push(_z.DestinationPointId);
                    _sr.push(_z.RangeId);
                });
                that.utility.mode(function () {
                    that.popup.confirm({
                        title: "Do you want to delete the selected destination point?"
                    }, function () {
                        that.popup.processing(true);
                        that.service.deleteSelected({ data: { "": _ss } }, function (result) {
                            if (result.status == app.status.succeeded) {
                                that.range.delSelected({ data: _sr, index: 0 }, function () {
                                    that.popup.message({ success: true, title: "Delete destination point succeeded." }, function () { that.popup.hide(3000); });
                                    $.each(_ss, function (_m, _n) {
                                        that.utility.remove({ Id: _n });
                                        that.ui.remove({ Id: _n });
                                    });
                                    that.utility.unSelectAll();
                                    that.utility.pager.init({ refresh: true, index: that.controls.list.find(".point-item").length > 0 ? that.pagerIndex : that.pagerIndex - 1 });
                                });
                            }
                            else {
                                that.popup.message({ success: false, title: "Delete destination point failed." });
                            }
                        });
                    }, function () {
                        that.controls.popupMain.removeClass("message process confirm active");
                    });
                });
            }
            else {
                that.popup.message({ success: false, title: "Please select destination point." });
            }
        },
        ///Edit the custom format.
        edit: function (o) {
            that.utility.mode(function () {
                var _i = $(o).data("id");
                that.model = that.utility.model(_i);
                if (that.model) {
                    that.range.goto({ RangeId: $(o).data("range") }, function (result) {
                        if (result.status == app.status.succeeded) {
                            that.controls.main.removeClass("manage add edit step-first step-second").addClass("add edit");
                            that.controls.formatBtn.prop("original", that.model.ReferencedSourcePoint.Value ? that.model.ReferencedSourcePoint.Value : "");
                            that.action.customFormat({ selected: that.model }, function () {
                                that.format.preview();
                            });
                        }
                        else {
                            that.popup.message({ success: false, title: "The point in the Word has been deleted." });
                        }
                    });
                }
                else {
                    that.popup.message({ success: false, title: "The destination point has been deleted." });
                }
            });
        },
        ///Toggle source point published history. 
        history: function (o) {
            o.hasClass("item-more") ? o.removeClass("item-more") : o.addClass("item-more");
        },
        ///Go to current seleccted destination point in word document.
        goto: function (o) {
            that.utility.mode(function () {
                that.controls.list.find(".point-item.item-selected").removeClass("item-selected");
                o.addClass("item-selected");
                var _rid = o.data("range");
                that.range.goto({ RangeId: _rid }, function (result) {
                    if (result.status == app.status.failed) {
                        that.popup.message({ success: false, title: "The point in the Word has been deleted." });
                    }
                });
            });
        },
        ///Support key up event to navigate up in the destination point management page.
        up: function () {
            var _i = that.controls.list.find(".point-item.item-selected").index();
            if (_i == -1) {
                _i = 0;
            }
            else {
                _i--;
            }
            if (_i >= 0) {
                that.action.goto(that.controls.list.find(">li").eq(_i));
            }
        },
        ///Support key down event to navigate down in the destination point management page.
        down: function () {
            var _i = that.controls.list.find(".point-item.item-selected").index(), _l = that.controls.list.find(">li").length;
            if (_i == -1) {
                _i = 0;
            }
            else {
                _i++;
            }
            if (_i < _l) {
                that.action.goto(that.controls.list.find(">li").eq(_i));
            }
        },
        ///Close the popup.
        ok: function () {
            that.controls.popupMain.removeClass("active message process confirm");
        },
        ///Open the source ponit file in a new window.
        open: function (o) {
            var _p = $(o).data("path");
            if (_p) {
                window.open(_p);
            }
        },
        ///Display source point result layer after entering the source point keyword in search textbox in add destination point page.
        autoComplete: function () {
            var _e = that.utility.entered(), _d = that.file.SourcePoints, _da = [];
            if ($.trim(_e.keyword) != "") {
                $.each(_d, function (i, d) {
                    if (that.utility.contain({ groups: d.Groups, selected: _e.groups })) {
                        _da.push(d);
                    }
                });
                app.search.autoComplete({ keyword: _e.keyword, data: _da, result: that.controls.autoCompleteControl, target: that.controls.keyword });
            }
            else {
                that.controls.autoCompleteControl.hide();
            }
        },
        ///Display source point result layer after entering the source point keyword in search textbox in destination point management page.
        autoComplete2: function () {
            var _e = $.trim(that.controls.sourcePointName.val()), _d = that.points, _da = [];
            if (_e != "") {
                $.each(_d, function (i, d) {
                    _da.push(d.ReferencedSourcePoint);
                });
                app.search.autoComplete({ keyword: _e, data: _da, result: that.controls.autoCompleteControl2, target: that.controls.sourcePointName });
            }
            else {
                that.controls.autoCompleteControl2.hide();
            }
        },
        ///Search the source points by source point name in management page.
        searchSourcePoint: function () {
            that.sourcePointKeyword = $.trim(that.controls.sourcePointName.val()) == that.controls.sourcePointName.data("default") ? "" : $.trim(that.controls.sourcePointName.val());
            that.utility.pager.init({ refresh: true });
        },
        ///Get the selected formats.
        selectedFormats: function (_that) {
            var _fi = [], _fd = [], _fn = [];
            that.controls.formatList.find("ul > li").each(function (i, d) {
                if ($(this).hasClass("checked")) {
                    _fi.push($(this).data("id"));
                    _fd.push($.trim($(this).text()));
                    _fn.push($.trim($(this).data("name")));
                    if ($.trim($(_that).data("name")).indexOf("ConvertTo") > -1) {
                        that.controls.formatBtn.prop("place", "");
                    }
                }
            });
            that.controls.formatBtn.html(_fd.length > 0 ? _fd.join(", ") : "None");
            that.controls.formatBtn.prop("title", _fd.length > 0 ? _fd.join(", ") : "None");
            that.controls.formatBtn.prop("selected", _fi.join(","));
            that.controls.formatBtn.prop("name", _fn.join(","));
            that.format.preview();
        },
        ///Get custom formats list and display in the add destination point 2nd step.
        customFormat: function (options, callback) {
            that.popup.processing(true);
            that.service.customFormat(function (result) {
                that.popup.processing(false);
                if (result.status == app.status.succeeded) {
                    if (result.data) {
                        that.ui.customFormat({ data: result.data, selected: options ? options.selected : null }, callback);
                    }
                }
                else {
                    that.ui.customFormat({ selected: options ? options.selected : null }, callback);
                    that.popup.message({ success: false, title: "Load custom format failed." });
                }
            });
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
                        that.files($.extend({}, { documentId: _d }, options));
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
                that.browse.file({ siteUrl: $(elem).data("siteurl"), listName: $(elem).data("listname"), name: $.trim($(elem).text()), url: that.browse.path[that.browse.path.length - 1].url + "/" + encodeURI($(elem).text()), fileName: $.trim($(elem).text()) });
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
    ///Define get document id function.
    that.document = {
        ///Initialize get graph and sharepoint access token.
        init: function (callback) {
            that.popup.processing(true);
            that.service.token({ endpoint: that.endpoints.token }, function (result) {
                if (result.status == app.status.succeeded) {
                    that.api.token = result.data;
                    that.api.host = that.utility.path()[0].toLowerCase();
                    that.service.token({ endpoint: that.endpoints.sharePointToken }, function (result) {
                        if (result.status == app.status.succeeded) {
                            that.api.sharePointToken = result.data;
                            that.document.site(null, callback);
                        }
                        else {
                            that.document.error({ title: "Get sharepoint access token failed." });
                        }
                    });
                }
                else {
                    that.document.error({ title: "Get graph access token failed." });
                }
            });
        },
        ///Get site id.
        site: function (options, callback) {
            if (options == null) {
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
                        options.values.push(result.data.id);
                        options.webUrls.push(result.data.webUrl);
                    }
                    options.index++;
                    that.document.site(options, callback);
                });
            }
            else {
                if (options.values.length > 0) {
                    that.document.library({ siteId: options.values.shift(), siteUrl: options.webUrls.shift() }, callback);
                }
                else {
                    that.document.error({ title: "Get site url failed." });
                }
            }
        },
        ///Get library id.
        library: function (options, callback) {
            that.service.libraries(options, function (result) {
                if (result.status == app.status.succeeded) {
                    var _l = "";
                    $.each(result.data.value, function (i, d) {
                        if (d.driveType.toUpperCase() == "DocumentLibrary".toUpperCase() && decodeURI(that.filePath).toUpperCase().indexOf(decodeURI(d.webUrl).toUpperCase()) > -1) {
                            _l = d.name;
                            return false;
                        }
                    });
                    if (_l != "") {
                        that.document.file({ siteId: options.siteId, siteUrl: options.siteUrl, listName: _l, fileName: that.utility.fileName(that.filePath) }, callback);
                    }
                    else {
                        that.document.error({ title: "Get library name failed." });
                    }
                }
                else {
                    that.document.error({ title: "Get library name failed." });
                }
            });
        },
        ///Get document id.
        file: function (options, callback) {
            that.service.item(options, function (result) {
                if (result.status == app.status.succeeded) {
                    var _d = "";
                    $.each(result.data.value, function (i, d) {
                        if (decodeURI(that.filePath).toUpperCase() == decodeURI(d.EncodedAbsUrl).toUpperCase() && d.OData__dlc_DocId) {
                            _d = d.OData__dlc_DocId;
                            return false;
                        }
                    });
                    if (_d != "") {
                        that.documentId = _d;
                        that.popup.processing(false);
                        callback();
                    }
                    else {
                        that.document.error({ title: "Get file Document ID failed." });
                    }
                }
                else {
                    that.document.error({ title: "Get file Document ID failed." });
                }
            });
        },
        ///Dispaly the error message.
        error: function (options) {
            that.controls.documentIdError.html("Error message: " + options.title);
            that.controls.main.addClass("error");
            that.popup.processing(false);
        }
    };
    ///Define custom format function.
    that.format = {
        ///Change the original value to formatted value.
        convert: function (options) {
            var _t = $.trim(options.value),
                _v = _t,
                _f = options.formats ? options.formats : [],
                _d = that.format.hasDollar(_v),
                _c = that.format.hasComma(_v),
                _p = that.format.hasPercent(_v),
                _m = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"],
                _x = options.decimal;
            $.each(_f, function (_a, _b) {
                if (!_b.IsDeleted) {
                    if (_b.Name == "ConvertToHundreds") {
                        if (that.format.isNumber(_v)) {
                            var _l = that.format.getDecimalLength(_v);
                            _v = new BigNumber(that.format.toNumber(_v)).div(100).toString();
                            _v = that.format.addDecimal(_v, _l);
                            if (_c) {
                                _v = that.format.addComma(_v);
                            }
                            if (_d) {
                                _v = that.format.addDollar(_v);
                            }
                        }
                    }
                    else if (_b.Name == "ConvertToThousands") {
                        if (that.format.isNumber(_v)) {
                            var _l = that.format.getDecimalLength(_v);
                            _v = new BigNumber(that.format.toNumber(_v)).div(1000).toString();
                            _v = that.format.addDecimal(_v, _l);
                            if (_c) {
                                _v = that.format.addComma(_v);
                            }
                            if (_d) {
                                _v = that.format.addDollar(_v);
                            }
                        }
                    }
                    else if (_b.Name == "ConvertToMillions") {
                        if (that.format.isNumber(_v)) {
                            var _l = that.format.getDecimalLength(_v);
                            _v = new BigNumber(that.format.toNumber(_v)).div(1000000).toString();
                            _v = that.format.addDecimal(_v, _l);
                            if (_c) {
                                _v = that.format.addComma(_v);
                            }
                            if (_d) {
                                _v = that.format.addDollar(_v);
                            }
                        }
                    }
                    else if (_b.Name == "ConvertToBillions") {
                        if (that.format.isNumber(_v)) {
                            var _l = that.format.getDecimalLength(_v);
                            _v = new BigNumber(that.format.toNumber(_v)).div(1000000000).toString();
                            _v = that.format.addDecimal(_v, _l);
                            if (_c) {
                                _v = that.format.addComma(_v);
                            }
                            if (_d) {
                                _v = that.format.addDollar(_v);
                            }
                        }
                    }
                    else if (_b.Name == "ShowNegativesAsPositives") {
                        var _h = that.format.hasDollar(_v),
                            _p = that.format.hasPercent(_v),
                            _hh = _v.toString().indexOf("hundred") > -1,
                            _ht = _v.toString().indexOf("thousand") > -1,
                            _hm = _v.toString().indexOf("million") > -1,
                            _hb = _v.toString().indexOf("billion") > -1;
                        var _tt = $.trim(_v.toString().replace(/\$/g, "").replace(/-/g, "").replace(/%/g, "").replace(/\(/g, "").replace(/\)/g, "").replace(/hundred/g, "").replace(/thousand/g, "").replace(/million/g, "").replace(/billion/g, ""));
                        if (that.format.isNumber(_tt)) {
                            _v = _tt;
                            if (_p) {
                                _v = _v + "%";
                            }
                            if (_h) {
                                _v = that.format.addDollar(_v);
                            }
                            if (_hh) {
                                _v = _v + " hundred";
                            }
                            else if (_ht) {
                                _v = _v + " thousand";
                            }
                            else if (_hm) {
                                _v = _v + " million";
                            }
                            else if (_hb) {
                                _v = _v + " billion";
                            }
                        }
                    }
                    else if (_b.Name == "IncludeHundredDescriptor") {
                        if (that.format.isNumber(_v)) {
                            _v = _v + " hundred";
                        }
                    }
                    else if (_b.Name == "IncludeThousandDescriptor") {
                        if (that.format.isNumber(_v)) {
                            _v = _v + " thousand";
                        }
                    }
                    else if (_b.Name == "IncludeMillionDescriptor") {
                        if (that.format.isNumber(_v)) {
                            _v = _v + " million";
                        }
                    }
                    else if (_b.Name == "IncludeBillionDescriptor") {
                        if (that.format.isNumber(_v)) {
                            _v = _v + " billion";
                        }
                    }
                    else if (_b.Name == "IncludeDollarSymbol") {
                        if (!that.format.hasDollar(_v)) {
                            _v = that.format.addDollar(_v);
                        }
                    }
                    else if (_b.Name == "ExcludeDollarSymbol") {
                        if (that.format.hasDollar(_v)) {
                            _v = that.format.removeDollar(_v);
                        }
                    }
                    else if (_b.Name == "DateShowLongDateFormat") {
                        if (that.format.isDate(_v)) {
                            var _tt = new Date(_v);
                            _v = _m[_tt.getMonth()] + " " + _tt.getDate() + ", " + _tt.getFullYear();
                        }
                    }
                    else if (_b.Name == "DateShowYearOnly") {
                        if (that.format.isDate(_v)) {
                            var _tt = new Date(_v);
                            _v = _tt.getFullYear();
                        }
                    }
                    else if (_b.Name == "ConvertNegativeSymbolToParenthesis") {
                        var _h = that.format.hasDollar(_v),
                            _hh = _v.toString().indexOf("hundred") > -1,
                            _ht = _v.toString().indexOf("thousand") > -1,
                            _hm = _v.toString().indexOf("million") > -1,
                            _hb = _v.toString().indexOf("billion") > -1;
                        if (_v.indexOf("-") > -1) {
                            var _tt = $.trim(_v.toString().replace(/\$/g, "").replace(/-/g, "").replace(/\(/g, "").replace(/\)/g, "").replace(/hundred/g, "").replace(/thousand/g, "").replace(/million/g, "").replace(/billion/g, ""));
                            _v = "(" + _tt + ")";
                            if (_h) {
                                _v = that.format.addDollar(_v);
                            }
                            if (_hh) {
                                _v = _v + " hundred";
                            }
                            else if (_ht) {
                                _v = _v + " thousand";
                            }
                            else if (_hm) {
                                _v = _v + " million";
                            }
                            else if (_hb) {
                                _v = _v + " billion";
                            }
                        }
                    }
                }
            });
            if (_x != null && _x.toString() != "") {
                var __a = that.format.hasComma(_v), __b = that.format.remove(_v);
                if (that.format.isNumber(__b)) {
                    var __c = __a ? that.format.addComma(__b) : __b, __d = "" + new BigNumber(__b).toFixed(_x) + "", __e = __a ? that.format.addComma(__d) : __d;
                    _v = _v.replace(__c, __e)
                }
            }
            return _v;
        },
        ///Remove the '$' & ',' at the start of the word.
        toNumber: function (_v) {
            return that.format.removeComma(that.format.removeDollar(_v));
        },
        ///Determine if this a number or not.
        isNumber: function (_v) {
            return !isNaN(that.format.toNumber(_v));
        },
        ///Determine if the input value can be converted to date or not.
        isDate: function (_v) {
            var _ff = false;
            try {
                _ff = _ff = (new Date(_v.toString().replace(/ /g, ""))).getFullYear() > 0;
            } catch (e) {
                _ff = false;
            }
            return _ff;
        },
        ///Determine if there is '$' or not.
        hasDollar: function (_v) {
            return _v.toString().indexOf("$") > -1;
        },
        ///Determine if there is ',' or not.
        hasComma: function (_v) {
            return _v.toString().indexOf(",") > -1;
        },
        ///Replace all '$' with empty in word.
        removeDollar: function (_v) {
            return _v.toString().replace("$", "");
        },
        ///Add '$' at the start of the word.
        addDollar: function (_v) {
            return "$" + _v;
        },
        ///Remove all ',' with empty in word.
        removeComma: function (_v) {
            return _v.toString().replace(/,/g, "");
        },
        ///add ',' when reach thousand. 
        addComma: function (_v) {
            var __s = _v.toString().split(".");
            __s[0] = __s[0].replace(new RegExp('(\\d)(?=(\\d{3})+$)', 'ig'), "$1,");
            return __s.join(".");
        },
        ///Determine if contain '%' or not in word.
        hasPercent: function (_v) {
            return _v.toString().indexOf("%") > -1;
        },
        ///Remove all '&', ',', '-', '%', '(', ')', 'hundred', 'thousand', 'million', and 'billion' with empty in word.
        remove: function (_v) {
            return $.trim(_v.toString().replace(/\$/g, "").replace(/,/g, "").replace(/-/g, "").replace(/%/g, "").replace(/\(/g, "").replace(/\)/g, "").replace(/hundred/g, "").replace(/thousand/g, "").replace(/million/g, "").replace(/billion/g, ""));
        },
        ///Get the length of decimal place 
        getDecimalLength: function (_v) {
            var _a = _v.toString().replace(/\$/g, "").replace(/,/g, "").replace(/-/g, "").replace(/\(/g, "").replace(/\)/g, "").split(".");
            if (_a.length == 2) {
                return _a[1].length;
            }
            else {
                return 0;
            }
        },
        ///Add decimal place.
        addDecimal: function (_v, _l) {
            var _dl = that.format.getDecimalLength(_v);
            if (_l > 0 && _dl == 0) {
                _v = "" + new BigNumber(_v).toFixed(_l) + "";
            }
            return _v;
        },
        ///Preview the formated value.
        preview: function () {
            var _v = that.controls.formatBtn.prop("original");
            var _n = that.controls.formatBtn.prop("name");
            var _f = (typeof (_n) != "undefined" && _n != "") ? _n.split(",") : [], _fa = [];
            var _x = that.controls.formatBtn.prop("place");
            $.each(_f, function (_a, _b) {
                _fa.push({ Name: _b });
            });
            var _fd = that.format.convert({ value: _v, formats: _fa, decimal: _x });
            that.controls.previewValue.html(_fd);
        }
    };
    ///Define all popup features.
    that.popup = {
        ///Dispay the message in the popup.
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
        ///Display the process popup or not
        processing: function (show) {
            if (!show) {
                that.controls.popupMain.removeClass("active process");
            }
            else {
                that.controls.popupMain.removeClass("message confirm browse").addClass("active process");
            }
        },
        ///Dipslay yes/no confirmation popup.
        confirm: function (options, yesCallback, noCallback) {
            that.controls.popupConfirmTitle.html(options.title);
            that.controls.popupMain.removeClass("message process browse").addClass("active confirm");
            that.controls.popupConfirmYes.unbind("click").click(function () {
                yesCallback();
            });
            that.controls.popupConfirmNo.unbind("click").click(function () {
                noCallback();
            });
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
        ///Go to destination management page.
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
    ///All actions in word for the content controls.
    that.range = {
        ///Create the content control.
        create: function (options, callback) {
            Word.run(function (ctx) {
                var _r = ctx.document.getSelection(), _c = _r.insertContentControl();
                _c.tag = options.RangeId;
                return ctx.sync().then(function () {
                    var _cc = ctx.document.contentControls.getByTag(options.RangeId);
                    ctx.load(_cc);
                    return ctx.sync().then(function () {
                        _cc.items[0].insertText(that.utility.toString(options.Value), Word.InsertLocation.replace);
                        return ctx.sync().then(function () {
                            callback({ status: app.status.succeeded });
                        });
                    });
                });
            }).catch(function (error) {
                callback({ status: app.status.failed, message: error.message });
            });
        },
        ///Go to content control by tag or range ID.
        goto: function (options, callback) {
            Word.run(function (ctx) {
                var r = ctx.document.contentControls.getByTag(options.RangeId);
                ctx.load(r, "items");
                return ctx.sync().then(function () {
                    r.items[0].select();
                    callback({ status: app.status.succeeded });
                });
            }).catch(function (error) {
                callback({ status: app.status.failed, message: error.message });
            });
        },
        ///Determine if the content control existed or not.
        exist: function (options, callback) {
            Word.run(function (ctx) {
                var r = ctx.document.contentControls.getByTag(options.RangeId);
                ctx.load(r);
                return ctx.sync().then(function () {
                    callback({ status: r.items.length > 0 ? app.status.succeeded : app.status.failed });
                });
            }).catch(function (error) {
                callback({ status: app.status.failed, message: error.message });
            });
        },
        ///Set the higtlight color to the selected content control.
        highlight: function (options, callback) {
            Word.run(function (ctx) {
                var r = ctx.document.contentControls.getByTag(options.RangeId);
                ctx.load(r);
                return ctx.sync().then(function () {
                    r.items[0].font.highlightColor = that.highlighted ? "" : that.highlightColor;
                    return ctx.sync().then(function () {
                        callback({ status: app.status.succeeded });
                    });
                });
            }).catch(function (error) {
                callback({ status: app.status.failed, message: error.message });
            });
        },
        ///Change the content control value to the destination point value from Azure storage. 
        edit: function (options, callback) {
            Word.run(function (ctx) {
                var r = ctx.document.contentControls.getByTag(options.RangeId);
                ctx.load(r);
                return ctx.sync().then(function () {
                    var _t = that.utility.toString(r.items[0].text), _v = that.utility.toString(options.Value);
                    if (_t != _v) {
                        r.items[0].insertText(_v, Word.InsertLocation.replace);
                        return ctx.sync().then(function () {
                            callback({ status: app.status.succeeded });
                        });
                    }
                    else {
                        callback({ status: app.status.succeeded });
                    }
                });
            }).catch(function (error) {
                callback({ status: app.status.failed, message: error.message });
            });
        },
        ///Delete the content control.
        del: function (options, callback) {
            Word.run(function (ctx) {
                var r = ctx.document.contentControls.getByTag(options.RangeId);
                ctx.load(r);
                return ctx.sync().then(function () {
                    if (r.items.length > 0) {
                        r.items[0].delete(false);
                        return ctx.sync().then(function () {
                            callback({ status: app.status.succeeded });
                        });
                    }
                    else {
                        callback({ status: app.status.succeeded });
                    }
                });
            }).catch(function (error) {
                callback({ status: app.status.failed, message: error.message });
            });
        },
        ///Delete the selected content control.
        delSelected: function (options, callback) {
            if (options.index < options.data.length) {
                var _rid = options.data[options.index];
                Word.run(function (ctx) {
                    var r = ctx.document.contentControls.getByTag(_rid);
                    ctx.load(r);
                    return ctx.sync().then(function () {
                        if (r.items.length > 0) {
                            r.items[0].delete(false);
                            return ctx.sync().then(function () {
                                options.index++;
                                that.range.delSelected(options, callback);
                            });
                        }
                        else {
                            options.index++;
                            that.range.delSelected(options, callback);
                        }
                    });
                }).catch(function (error) {
                    options.index++;
                    that.range.delSelected(options, callback);
                });
            }
            else {
                callback();
            }
        },
        ///Get all control controls.
        all: function (callback) {
            Word.run(function (ctx) {
                var _cc = ctx.document.contentControls, _ar = [];
                ctx.load(_cc);
                return ctx.sync().then(function () {
                    if (_cc.items.length > 0) {
                        for (var _i = 0; _i < _cc.items.length; _i++) {
                            _ar.push({ tag: _cc.items[_i].m_tag, text: _cc.items[_i].m_text });
                        }
                    }
                    callback({ status: app.status.succeeded, data: _ar });
                });
            }).catch(function (error) {
                callback({ status: app.status.failed, message: error.message });
            });
        },
        ///Return the index valoe of the array.
        index: function (options) {
            var _i = -1;
            $.each(options.data, function (i, d) {
                if (d.tag == options.tag) {
                    _i = i;
                    return false;
                }
            });
            return _i;
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
        ///Add destination point (POST) to Azure storage.
        add: function (options, callback) {
            that.service.common({ url: that.endpoints.add, type: "POST", data: options.data, dataType: "json" }, callback);
        },
        ///Update the destination point (PUT) in Azure storage.
        edit: function (options, callback) {
            that.service.common({ url: that.endpoints.edit, type: "PUT", data: options.data, dataType: "json" }, callback);
        },
        ///Get source point by catalog document id.
        catalog: function (options, callback) {
            that.service.common({ url: that.endpoints.catalog + options.documentId, type: "GET", dataType: "json" }, callback);
        },
        ///Get the source point groups.
        groups: function (callback) {
            that.service.common({ url: that.endpoints.groups, type: "GET", dataType: "json" }, callback);
        },
        ///Get a list of destination points from Azure storage.
        list: function (callback) {
            that.service.common({ url: that.endpoints.list + that.filePath + "&documentId=" + that.documentId, type: "GET", dataType: "json" }, callback);
        },
        ///Delete current destination point in Azure storage.
        del: function (options, callback) {
            that.service.common({ url: that.endpoints.del + options.Id, type: "DELETE" }, callback);
        },
        ///Delete the selected destination points in Azure storage.
        deleteSelected: function (options, callback) {
            that.service.common({ url: that.endpoints.deleteSelected, type: "POST", data: options.data }, callback);
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
        },
        ///Get custom formats from Azure storage.
        customFormat: function (callback) {
            that.service.common({ url: that.endpoints.customFormat, type: "GET", dataType: "json" }, callback);
        },
        ///Update the destination point custom formats(PUT) in Azure storage.
        update: function (options, callback) {
            that.service.common({ url: that.endpoints.updateCustomFormat, type: "PUT", data: options.data, dataType: "json" }, callback);
        }
    };
    ///Build the HTML/UI.
    that.ui = {
        ///Clear the file & keyword textboxes.
        clear: function () {
            that.controls.file.val("");
            that.controls.keyword.val("");
        },
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
        ///Set the status of next/cancel/save buttons.
        status: function (options) {
            options.next ? that.controls.next.removeClass("disabled") : that.controls.next.addClass("disabled");
            options.cancel ? that.controls.cancel.removeClass("disabled") : that.controls.cancel.addClass("disabled");
            options.save ? that.controls.save.removeClass("disabled") : that.controls.save.addClass("disabled");
        },
        ///Reset file/keyword textboxes and filter by groups checkboxes after selecting the excel file.
        select: function () {
            that.keyword = "";
            that.controls.keyword.val("");
            that.ui.dft();
            that.controls.filterList.find("input").removeAttr("checked");
        },
        ///Display the default screen when add destination point.
        reset: function () {
            that.ui.clear();
            that.ui.dft();
            that.ui.status({ next: false });
            that.controls.main.removeClass("step-second").addClass("step-first");
            that.controls.stepFirstMain.removeClass("selected-file");
            that.controls.filterMain.removeClass("open-filter");
        },
        ///Build filter by groups HTML
        groups: function (options) {
            that.controls.filterList.html("");
            $.each(options, function (i, d) {
                $('<li data-id="' + d.Id + '"><div><div class="ckb-wrapper"><input type="checkbox" id="cbkGroup_' + d.Id + '" value="' + d.Id + '" /><i></i></div></div><label for="cbkGroup_' + d.Id + '">' + d.Name + '</label></li>').appendTo(that.controls.filterList);
            });
        },
        ///Build the source points HTML after selecting the excel file.
        sources: function (options) {
            var _f = false, _d = options.data != null ? options.data.SourcePoints : [], _da = [];
            that.ui.status({ next: false });
            that.controls.resultList.html("");
            _d.sort(function (_a, _b) {
                return (_a.Name.toUpperCase() > _b.Name.toUpperCase()) ? 1 : (_a.Name.toUpperCase() < _b.Name.toUpperCase()) ? -1 : 0;
            });

            if (options.keyword != undefined && $.trim(options.keyword) != "") {
                var _sk = app.search.splitKeyword({ keyword: $.trim(options.keyword) });
                if (_sk.length > 26) {
                    that.popup.message({ success: false, title: "Only support less then 26 keywords." });
                }
                else {
                    var _dt = [];
                    $.each(_d, function (i, d) {
                        if (that.utility.contain({ groups: d.Groups, selected: options.selected })) {
                            var _wi = app.search.weight({ keyword: _sk, source: d.Name });
                            if (_wi > 0) {
                                _da.push(d);
                                _f = true;
                            }
                        }
                    });
                }
            }
            else {
                $.each(_d, function (i, d) {
                    if (that.utility.contain({ groups: d.Groups, selected: options.selected })) {
                        _da.push(d);
                        _f = true;
                    }
                });
            }

            $.each(_da, function (i, d) {
                var _p = that.utility.position(d.Position);
                $('<li data-id="' + d.Id + '" data-file="' + that.utility.fileName(options.data.Name) + '" data-name="' + d.Name + '" data-value="' + (d.Value ? d.Value : "") + '">' + d.Name + ' | <span>' + _p.sheet + ' [' + _p.cell + '] </span> | <span>' + (d.Value ? d.Value : "") + '</span></li>').appendTo(that.controls.resultList);
            });

            _f ? that.controls.resultNotFound.hide() : that.controls.resultNotFound.show();
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
                        if (app.search.weight({ keyword: _sk, source: d.ReferencedSourcePoint.Name }) > 0) {
                            _d.push(d);
                        }
                    });
                }
            }
            else {
                _d = _dt;
            }
            that.utility.pager.status({ length: _d.length });
            that.range.all(function (result) {
                if (result.status == app.status.succeeded) {
                    $.each(_d, function (i, d) {
                        var __i = that.range.index({ data: result.data, tag: d.RangeId });
                        _d[i].orderBy = __i > -1 ? __i : 99999;
                        _d[i].existed = __i > -1 ? true : false;
                        _d[i].changed = __i > -1 ? $.trim(result.data[__i].text) != that.format.convert({ value: that.utility.toString(d.ReferencedSourcePoint.Value), formats: d.CustomFormats, decimal: d.DecimalPlace }) : false;
                    });
                }
                else {
                    $.each(_d, function (i, d) {
                        _d[i].orderBy = 99999;
                        _d[i].existed = false;
                        _d[i].changed = false;
                    });
                }
                _d.sort(function (_a, _b) {
                    return (_a.orderBy > _b.orderBy) ? 1 : (_a.orderBy < _b.orderBy) ? -1 : 0;
                });

                if (options.refresh) {
                    var _c = that.controls.headerListPoints.find(".point-header .ckb-wrapper input").prop("checked"), _s = that.utility.selected();
                    $.each(_s, function (m, n) {
                        _ss.push(n.DestinationPointId);
                    });
                    that.controls.headerListPoints.find(".point-header .ckb-wrapper input").prop("checked", _c);
                    if (_c) {
                        that.controls.headerListPoints.find(".point-header .ckb-wrapper").addClass("checked");
                    }
                    else {
                        that.controls.headerListPoints.find(".point-header .ckb-wrapper").removeClass("checked");
                    }
                }
                else {
                    that.controls.headerListPoints.find(".point-header .ckb-wrapper input").prop("checked", false);
                    that.controls.headerListPoints.find(".point-header .ckb-wrapper").removeClass("checked");
                }
                that.controls.list.find(".point-item").remove();
                that.ui.item({ index: 0, data: _d, refresh: options.refresh, selected: _ss });
            });
        },
        ///Build each destination point HTML in management page.
        item: function (options, callback) {
            if (options.index < options.data.length) {
                var _dsp = options.data[options.index], _item = _dsp.ReferencedSourcePoint, _sourcePointCatalog = _item.Catalog, _s = _dsp.existed && _item.Status == 0;
                if (options.index >= that.pagerSize * (that.pagerIndex - 1) && options.index < that.pagerSize * that.pagerIndex) {
                    var _p = that.utility.position(_item.Position),
                        _fn = that.utility.fileName(_sourcePointCatalog.Name),
                        _sel = $.inArray(_dsp.Id, options.selected) > -1,
                        _fv = that.format.convert({ value: _item.Value ? _item.Value : "", formats: _dsp.CustomFormats, decimal: _dsp.DecimalPlace }),
                        _ff = [],
                        _pht = _item.PublishedHistories && _item.PublishedHistories.length > 0 ? _item.PublishedHistories : [],
                        _pi = 0;
                    $.each(_dsp.CustomFormats != null ? _dsp.CustomFormats : [], function (_x, _y) { _ff.push(_y.DisplayName); });
                    if (_dsp.DecimalPlace != null && _dsp.DecimalPlace != "") {
                        _ff.push("Displayed decimals");
                    }
                    var _cf = _ff.join("; ").replace(/"/g, "&quot;");
                    var _h = '<li class="point-item' + (_s ? "" : " item-error") + '" data-id="' + _dsp.Id + '" data-range="' + _dsp.RangeId + '">';
                    _h += '<div class="point-item-line">';
                    _h += '<div class="i1"><div class="ckb-wrapper' + (_sel ? " checked" : "") + '"><input type="checkbox" ' + (_sel ? 'checked="checked"' : '') + ' /><i></i></div></div>';
                    _h += '<div class="i2"><span class="s-name" title="' + _item.Name + '">' + _item.Name + '</span>';
                    _h += '<span><strong title="' + (_p.sheet ? _p.sheet : "") + ':[' + (_p.cell ? _p.cell : "") + ']">' + (_p.sheet ? _p.sheet : "") + ':</strong>[' + (_p.cell ? _p.cell : "") + ']</span>';
                    _h += '<span><strong class="i-file" title="' + _sourcePointCatalog.Name + '" data-path="' + _sourcePointCatalog.Name + '">' + _fn + '</strong></span>';
                    _h += '</div>';
                    _h += '<div class="i3" title="' + (_item.Value ? _item.Value : "") + '">' + (_item.Value ? _item.Value : "") + '</div>';
                    _h += '<div class="i5"><div class="i-line"><i class="i-history" title="History"></i><i class="i-delete" title="Delete"></i><i class="i-edit" title="Edit Custom Format"></i></div>';
                    _h += '<div class="i-menu"><a href="javascript:"><span title="Action">...</span><span><i class="i-history" title="History"></i><i class="i-delete" title="Delete"></i><i class="i-edit" title="Edit Custom Format"></i></span></a></div>';
                    _h += '</div>';
                    _h += '</div>';
                    _h += '<div class="item-format">';
                    _h += '<span class="item-formatted" title="' + _fv + '"><strong>' + (_ff.length > 0 ? "Formatted Value" : "Source Point Value") + ':</strong>' + _fv + '</span>';
                    _h += '<span class="item-formats" title="' + (_ff.length > 0 ? _cf : "No custom formatting applied") + '"><strong>Format:</strong>' + (_ff.length > 0 ? _cf : "No custom formatting applied") + '</span>';
                    _h += '</div>';
                    _h += '<div class="item-history"><h6>Publish History</h6><ul class="history-list">';
                    _h += '<li class="history-header"><div class="h1">Name</div><div class="h2">Value</div><div class="h3">Date</div></li>';
                    $.each(_pht, function (m, n) {
                        var __c = $.trim(_pht[m].Value ? _pht[m].Value : ""),
                            __p = $.trim(_pht[m > 0 ? m - 1 : m].Value ? _pht[m > 0 ? m - 1 : m].Value : "");
                        if (_pi < 5 && (m == 0 || __c != __p)) {
                            _h += '<li class="history-item"><div class="h1" title="' + n.PublishedUser + '">' + n.PublishedUser + '</div><div class="h2" title="' + (n.Value ? n.Value : "") + '">' + (n.Value ? n.Value : "") + '</div><div class="h3" title="' + that.utility.date(n.PublishedDate) + '">' + that.utility.date(n.PublishedDate) + '</div></li>';
                            _pi++;
                        }
                    });
                    _h += '</ul>';
                    _h += '</div>';
                    _h += '<div class="error-info">';
                    _h += '<span>Error</span>';
                    if (!_dsp.existed) {
                        _h += '<p>The content control that was bound to the source point for this destination point has been removed, please deleted destination point.</p>';
                    }
                    else if (_item.Status == 1) {
                        _h += '<p>The source point used for this destination point has been deleted from the catalog, please delete the destination point and use a valid source point.</p>';
                    }
                    _h += '</div>';
                    _h += '</li>';
                    that.controls.list.append(_h);
                }
                if (_dsp.changed) {
                    options.tooltip = true;
                }
                options.index++;
                if (options.refresh && _s && _dsp.changed) {
                    that.range.edit({ RangeId: _dsp.RangeId, Value: that.format.convert({ value: _item.Value ? _item.Value : "", formats: _dsp.CustomFormats, decimal: _dsp.DecimalPlace }) }, function (result) {
                        that.ui.item(options, callback);
                    });
                }
                else {
                    that.ui.item(options, callback);
                }
            }
            else {
                that.controls.main.get(0).scrollTop = 0;
                if (that.firstLoad && options.tooltip) {
                    that.controls.tooltipMessage.addClass("active");
                }
                that.firstLoad = false;
                if (callback) {
                    callback();
                }
            }
        },
        ///Remove the destination point from the management page.
        remove: function (options) {
            that.controls.list.find("[data-id=" + options.Id + "]").remove();
        },
        ///Build the custom format dropdown list HTML.
        customFormat: function (options, callback) {
            that.controls.formatList.html("");
            var _si = [], _sn = [], _sd = [];
            if (options.selected && options.selected != null) {
                $.each(options.selected.CustomFormats, function (_x, _y) {
                    _si.push(_y.Id);
                    _sn.push(_y.Name);
                    _sd.push(_y.DisplayName);
                });
                that.controls.formatBtn.html(_sd.length > 0 ? _sd.join(", ") : "None");
                that.controls.formatBtn.prop("title", _sd.length > 0 ? _sd.join(", ") : "None");
                that.controls.formatBtn.prop("selected", _si.join(","));
                that.controls.formatBtn.prop("name", _sn.join(","));
                that.controls.formatBtn.prop("place", options.selected.DecimalPlace && options.selected.DecimalPlace != null ? options.selected.DecimalPlace : "");
            }
            else {
                that.controls.formatBtn.html("None");
                that.controls.formatBtn.prop("title", "None");
                that.controls.formatBtn.prop("selected", "");
                that.controls.formatBtn.prop("name", "");
                that.controls.formatBtn.prop("place", "");
            }

            var _v = that.controls.formatBtn.prop("original");
            that.controls.formatList.removeClass("convert1 convert2 convert3 convert4");
            that.controls.previewValue.html(_v);
            that.controls.addCustomFormat.removeClass("selected-number selected-date");
            if (that.format.isNumber(_v)) {
                that.controls.addCustomFormat.addClass("selected-number");
            }
            else if (that.format.isDate(_v)) {
                that.controls.addCustomFormat.addClass("selected-date");
            }

            if (_sn.length > 0) {
                var _tn = $.inArray("ConvertToThousands", _sn) > -1 ? "IncludeThousandDescriptor" : ($.inArray("ConvertToMillions", _sn) > -1 ? "IncludeMillionDescriptor" : ($.inArray("ConvertToBillions", _sn) > -1 ? "IncludeBillionDescriptor" : ($.inArray("ConvertToHundreds", _sn) > -1 ? "IncludeHundredDescriptor" : "")));
                var _cl = $.inArray("ConvertToThousands", _sn) > -1 ? "convert2" : ($.inArray("ConvertToMillions", _sn) > -1 ? "convert3" : ($.inArray("ConvertToBillions", _sn) > -1 ? "convert4" : ($.inArray("ConvertToHundreds", _sn) > -1 ? "convert1" : "")));
                that.controls.formatList.addClass(_cl);
                that.controls.formatList.find("ul > li[data-name=" + _tn + "]").addClass("checked");
            }

            var _dd = [], _dt = [];
            $.each(options.data, function (_i, _e) {
                var _i = $.inArray(_e.GroupName, _dt);
                if (_i == -1) {
                    _dd.push({ Name: _e.GroupName, OrderBy: _e.GroupOrderBy, Formats: [{ Id: _e.Id, Name: _e.Name, DisplayName: _e.DisplayName, Description: _e.Description, OrderBy: _e.OrderBy }] });
                    _dt.push(_e.GroupName);
                }
                else {
                    _dd[_i].Formats.push({ Id: _e.Id, Name: _e.Name, DisplayName: _e.DisplayName, Description: _e.Description, OrderBy: _e.OrderBy });
                }
            });
            $.each(_dd, function (_i, _e) {
                _e.Formats.sort(function (_m, _n) {
                    return _m.OrderBy > _n.OrderBy ? 1 : _m.OrderBy < _n.OrderBy ? -1 : 0;
                });
            });
            _dd.sort(function (_m, _n) {
                return _m.OrderBy > _n.OrderBy ? 1 : _m.OrderBy < _n.OrderBy ? -1 : 0;
            });

            if (_dd) {
                $.each(_dd, function (m, n) {
                    var _h = '', _c = '';
                    _c = (n.Name == "Convert to" || n.Name == "Negative number" || n.Name == "Descriptor") ? "value-number" : (n.Name == "Symbol") ? "value-string" : "value-date";
                    _h += '<li class="' + (n.Name == "Descriptor" ? "drp-checkbox drp-descriptor " : "drp-radio ") + '' + _c + '">';
                    if (n.Name != "Descriptor") {
                        _h += '<label>' + n.Name + '</label>';
                    }
                    _h += '<ul>';
                    $.each(n.Formats, function (i, d) {
                        _h += '<li data-id="' + d.Id + '" data-name="' + d.Name + '" title="' + d.Description + '" class="' + ($.inArray(d.Id, _si) > -1 ? "checked" : "") + '">';
                        _h += '<div><i></i></div>';
                        _h += '<a href="javascript:">' + (n.Name == "Descriptor" ? "Descriptor" : d.DisplayName) + '</a>';
                        _h += '</li>';
                    });
                    _h += '</ul>';
                    _h += '</li>';
                    that.controls.formatList.append(_h);
                });
            }
            if (callback) {
                callback();
            }
        }
    };

    return that;
})();