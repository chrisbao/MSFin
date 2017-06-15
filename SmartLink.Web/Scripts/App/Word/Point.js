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
        /// Define all UI controls.
        controls: {},
        ///
        file: null,
        ///
        selected: null,
        ///Define source point groups.
        groups: [],
        ///Define the destination points. 
        points: [],
        ///todo
        keyword: "",
        ///todo
        sourcePointKeyword: "",
        //the background color for the highlighted destination ponit.
        highlightColor: "#66FF00",
        //highlight the destination point or not
        highlighted: false,
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
            catalog: "/api/SourcePointCatalog?name=",
            groups: "/api/SourcePointGroup",
            list: "/api/DestinationPointCatalog?name=",
            del: "/api/DestinationPoint?id=",
            deleteSelected: "/api/DeleteSelectedDestinationPoint",
            token: "/api/GraphAccessToken",
            graph: "https://graph.microsoft.com/beta",
            customFormat: "/api/CustomFormats"
        },
    }, that = point;

    that.init = function () {
        ///get the document URL.
        that.filePath = Office.context && Office.context.document && Office.context.document.url ? Office.context.document.url : "https://cand3.sharepoint.com/Shared%20Documents/MaxTest.docx";
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
            headerListPoints: $("#headerListPoints"),
            moveUp: $("#btnMoveUp"),
            moveDown: $("#btnMoveDown"),
            tooltipMessage: $("#tlpMessage"),
            tooltipClose: $("#tlpClose"),
            formatBtn: $("#btnSelectFormat"),
            formatIcon: $("#iconSelectFormat"),
            formatList: $("#listFormats"),
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
        ///define the event handlers.
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
            that.action.save();
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
            that.controls.formatList.hasClass("active") ? that.controls.formatList.removeClass("active") : that.controls.formatList.addClass("active");
            return false;
        });
        that.controls.formatIcon.click(function () {
            that.controls.formatList.hasClass("active") ? that.controls.formatList.removeClass("active") : that.controls.formatList.addClass("active");
            return false;
        });
        that.controls.formatList.on("click", "li", function () {
            var _ck = $(this).hasClass("checked");
            _ck ? $(this).removeClass("checked") : $(this).addClass("checked");
            if ($(this).hasClass("drp-header")) {
                if (_ck) {
                    that.controls.formatList.find("li").removeClass("checked");
                }
                else {
                    that.controls.formatList.find("li").addClass("checked");
                }
            }
            else {
                if (_ck) {
                    if (that.controls.formatList.find("li.drp-item.checked").length == 0) {
                        that.controls.formatList.find("li.drp-header").removeClass("checked");
                    }
                }
                else {
                    if (that.controls.formatList.find("li.drp-item.checked").length == that.controls.formatList.find("li.drp-item").length) {
                        that.controls.formatList.find("li.drp-header").addClass("checked");
                    }
                }
            }
            that.action.selectedFormats();
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
        that.controls.list.on("click", ".point-item .i2, .point-item .i3, .point-item .i5, .point-item .error-info, .point-item .item-history", function (e) {
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
        });
        that.utility.height();
        that.action.dft(that.controls.sourcePointName, false);
        that.list({ refresh: false, index: 1 }, function (result) {
            if (result.status == app.status.failed) {
                that.popup.message({ success: false, title: result.error.statusText });
            }
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
            }
            else {
                callback({ status: result.status, error: result.error });
            }
        });
    };
    ///default display when add destination point.
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
    ///load source points after selecting the file in the file explorer.
    that.files = function (options) {
        that.popup.processing(true);
        that.service.catalog({ path: options.path }, function (result) {
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
        /// get current destination point.
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
        //get sheet/cell name.
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
        ///get the index of current destination point.
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
        ///get the file name.
        fileName: function (path) {
            return path.lastIndexOf("/") > -1 ? path.substr(path.lastIndexOf("/") + 1) : (path.lastIndexOf("\\") > -1 ? path.substr(path.lastIndexOf("\\") + 1) : path);
        },
        /// get/set the highlighted destination points.
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
        //convert the val tostring and remove the blank space at the start or end of the word.
        toString: function (val) {
            if (val != undefined && val != null) {
                return $.trim(val.toString());
            }
            return "";
        },
        ///Paging feature for the destination point list.
        pager: {
            ///initialize the destination point list UI.
            init: function (options) {
                that.controls.pagerValue.val("");
                that.pagerIndex = options.index ? options.index : 1;
                that.ui.list({ refresh: options.refresh });
            },
            //Go to prev page.
            prev: function () {
                that.controls.pagerValue.val("");
                that.pagerIndex--;
                that.ui.list({ refresh: false });
            },
            //Go to next page.
            next: function () {
                that.controls.pagerValue.val("");
                that.pagerIndex++;
                that.ui.list({ refresh: false });
            },
            //Get the status of paging.
            status: function (options) {
                that.pagerCount = Math.ceil(options.length / that.pagerSize);
                that.controls.pagerTotal.html(options.length);
                that.controls.pagerPages.html(that.pagerCount);
                that.controls.pagerCurrent.html(that.pagerIndex);
                that.pagerIndex == 1 || that.pagerCount == 0 ? that.controls.pagerPrev.addClass("disabled") : that.controls.pagerPrev.removeClass("disabled");
                that.pagerIndex == that.pagerCount || that.pagerCount == 0 ? that.controls.pagerNext.addClass("disabled") : that.controls.pagerNext.removeClass("disabled");
            }
        },
        //Sort the destination point in managment page by the conntent controls in word document.
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
        //return current selected destination points in management page.
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
        //Caculate the height for the destination point list in management page or add destination page dymamically and set the scroll bar accordingly.. 
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
        //Uncheck all selected destination points.
        unSelectAll: function () {
            that.controls.headerListPoints.find(".point-header .ckb-wrapper input").prop("checked", false);
            that.controls.headerListPoints.find(".point-header .ckb-wrapper").removeClass("checked");
        },
            
        //Get an array of paths (server ralative URl splitted by '/' )
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
        }
    };

    that.action = {
        ///hide the tooltip and hide the custom format dropdown list.
        body: function () {
            $(".search-tooltips").hide();
            that.controls.formatList.removeClass("active");
        },
        ///add destination point.
        add: function () {
            that.utility.mode(function () {
                that.file = null;
                that.selected = null;
                that.default();
            });
        },
        ///go to destination management page.
        back: function () {
            that.controls.main.removeClass("add step-first step-second").addClass("manage");
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
        ///display toggle for the filter groups 
        filter: function () {
            that.controls.filterMain.hasClass("open-filter") ? that.controls.filterMain.removeClass("open-filter") : that.controls.filterMain.addClass("open-filter");
        },
        ///Set the selected excel file name in file textbox and display the source points
        select: function (options) {
            that.controls.file.val(options.name).removeClass("input-default");
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
        ///select the checkbox.
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
        ///enable the next button after selecting the source point search result.
        choose: function (o) {
            var _i = o.data("id");
            if (!that.selected || (that.selected && that.selected.Id != _i)) {
                that.selected = { Id: _i, File: o.data("file"), Name: o.data("name"), Value: o.data("value") };
                that.controls.resultList.find("li.selected").removeClass("selected");
                o.addClass("selected");
            }
            that.controls.selectedName.html(o.data("name"));
            that.controls.selectedFile.html("<strong>Source file:</strong>" + o.data("file") + "");
            that.ui.status({ next: true });
        },
        ///display the second step when add destination point.
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
        //Display the step 1.
        cancel: function () {
            that.ui.status({ next: true });
            that.controls.main.removeClass("step-second").addClass("step-first");
        },
        ///Add new destination point.
        save: function () {
            that.utility.mode(function () {
                var _s = that.controls.formatBtn.prop("selected"), _n = that.controls.formatBtn.prop("name"),
                    _f = (typeof (_n) != "undefined" && _n != "") ? _n.split(",") : [],
                    _c = (typeof (_s) != "undefined" && _s != "") ? _s.split(",") : [],
                    _fa = [];
                $.each(_f, function (_a, _b) {
                    _fa.push({ Name: _b });
                });
                var _v = that.format.convert({ value: that.selected.Value, formats: _fa });
                var _json = $.extend({}, that.selected, { RangeId: app.guid(), CatalogName: that.filePath, CustomFormatIds: _c, Value: _v });

                that.range.create(_json, function (ret) {
                    if (ret.status == app.status.succeeded) {
                        that.popup.processing(true);
                        that.service.add({ data: { CatalogName: _json.CatalogName, RangeId: _json.RangeId, SourcePointId: _json.Id, CustomFormatIds: _json.CustomFormatIds } }, function (result) {
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
        ///highlight all destination point in word document.
        highlightAll: function () {
            that.popup.processing(true);
            that.utility.order(function (result) {
                options = { index: 0, data: result.data, errorAmount: 0, successAmount: 0 };
                that.action.highlight(options);
            });
        },
        ///highligh one destination point in word document.
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
        ///delete the destination point after clicking X icon.
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
        ///delete selected destination point after clicking the delete button.
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
        ///toggle source point published history. 
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
        ///close the popup.
        ok: function () {
            that.controls.popupMain.removeClass("active message process confirm");
        },
        ///open the source ponit file in a new window.
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
        selectedFormats: function () {
            var _fi = [], _fd = [], _fn = [];
            that.controls.formatList.find("li").each(function (i, d) {
                if (!$(this).hasClass("drp-header") && $(this).hasClass("checked")) {
                    _fi.push($(this).data("id"));
                    _fd.push($.trim($(this).text()));
                    _fn.push($.trim($(this).data("name")));
                }
            });
            that.controls.formatBtn.html(_fd.length > 0 ? _fd.join(", ") : "None");
            that.controls.formatBtn.prop("selected", _fi.join(","));
            that.controls.formatBtn.prop("name", _fn.join(","));
        },
        ///get custom formats list and display in the add destination point 2nd step.
        customFormat: function () {
            that.popup.processing(true);
            that.service.customFormat(function (result) {
                that.popup.processing(false);
                if (result.status == app.status.succeeded) {
                    if (result.data) {
                        that.ui.customFormat({ data: result.data });
                    }
                }
                else {
                    that.ui.customFormat();
                    that.popup.message({ success: false, title: "Load custom format failed." });
                }
            });
        }
    };
    ///file explorer
    that.browse = {
        accessToken: "",
        host: "",
        path: [],
        init: function () {
            that.browse.accessToken = "";
            that.browse.path = [];
            that.browse.popup.dft();
            that.browse.popup.show();
            that.browse.popup.processing(true);
            that.browse.token();
        },
        //Get the graph access token
        token: function () {
            that.service.token(function (result) {
                if (result.status == app.status.succeeded) {
                    that.browse.accessToken = result.data;
                    that.browse.host = that.utility.path()[0].toLowerCase();
                    that.browse.siteCollection();
                }
                else {
                    that.browse.popup.message("Get graph access token failed.");
                }
            });
        },
        ///get available site colletion ID.
        siteCollection: function (options) {
            if (typeof (options) == "undefined") {
                options = {
                    path: that.utility.path().reverse(),
                    index: 0,
                    values: []
                };
            }
            if (options.index < options.path.length) {
                that.service.siteCollection({ path: options.path[options.index] }, function (result) {
                    if (result.status == app.status.succeeded) {
                        if (typeof (result.data.siteCollection) != "undefined") {
                            options.values.push(result.data.id);
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
                    that.browse.sites({ siteId: options.values.shift() });
                }
                else {
                    that.browse.popup.message("Get site collection ID failed.");
                }
            }
        },
        ///get all subsites under current site collection.
        sites: function (options) {
            that.service.sites(options, function (result) {
                if (result.status == app.status.succeeded) {
                    var _s = [];
                    $.each(result.data.value, function (i, d) {
                        _s.push({ id: d.id, name: d.name, type: "site" });
                    });
                    that.browse.libraries({ siteId: options.siteId, sites: _s });
                }
                else {
                    that.browse.popup.message("Get sites failed.");
                }
            });
        },
        ///get all document libraries under all webs.
        libraries: function (options) {
            that.service.libraries(options, function (result) {
                if (result.status == app.status.succeeded) {
                    var _l = options.sites ? options.sites : [];
                    $.each(result.data.value, function (i, d) {
                        if (d.list.template.toUpperCase() == "DocumentLibrary".toUpperCase()) {
                            _l.push({ id: d.id, name: decodeURI(d.name), type: "library", siteId: options.siteId, url: d.webUrl });
                        }
                    });
                    that.browse.display({ data: _l });
                }
                else {
                    that.browse.popup.message("Get libraries failed.");
                }
            });
        },

        /// get all folders or excel files under the document library.
        items: function (options) {
            if (options.inFolder) {
                that.service.itemsInFolder(options, function (result) {
                    if (result.status == app.status.succeeded) {
                        var _fd = [], _fi = [];
                        $.each(result.data.value, function (i, d) {
                            var _u = d.webUrl, _n = d.name, _nu = decodeURI(_n);
                            if (d.folder) {
                                _fd.push({ id: d.id, name: _nu, type: "folder", url: _u, siteId: options.siteId, listId: options.listId });
                            }
                            else if (d.file) {
                                if (_n.toUpperCase().indexOf(".XLSX") > 0) {
                                    _fi.push({ id: d.id, name: _nu, type: "file", url: _u, siteId: options.siteId, listId: options.listId });
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
                                _fd.push({ id: d.id, name: _nu, type: "folder", url: _u, siteId: options.siteId, listId: options.listId });
                            }
                            else if (d.file) {
                                if (_n.toUpperCase().indexOf(".XLSX") > 0) {
                                    _fi.push({ id: d.id, name: _nu, type: "file", url: _u, siteId: options.siteId, listId: options.listId });
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
        ///display site/library/folder/file in the file explorer popup.
        display: function (options) {
            that.controls.popupBrowseList.html("");
            $.each(options.data, function (i, d) {
                var _h = "";
                if (d.type == "site") {
                    _h = '<li class="i-site" data-id="' + d.id + '" data-type="site">' + d.name + '</li>';
                }
                else if (d.type == "library") {
                    _h = '<li class="i-library" data-id="' + d.id + '" data-site="' + d.siteId + '" data-url="' + d.url + '" data-type="library">' + d.name + '</li>';
                }
                else if (d.type == "folder") {
                    _h = '<li class="i-folder" data-id="' + d.id + '" data-site="' + d.siteId + '" data-list="' + d.listId + '" data-url="' + d.url + '" data-type="folder">' + d.name + '</li>';
                }
                else if (d.type == "file") {
                    _h = '<li class="i-file" data-id="' + d.id + '" data-site="' + d.siteId + '" data-list="' + d.listId + '" data-url="' + d.url + '" data-type="file">' + d.name + '</li>';
                }
                that.controls.popupBrowseList.append(_h);
            });
            that.browse.popup.processing(false);
        },
        ///Navigate the file explorer.
        select: function (elem) {
            var _t = $(elem).data("type");
            ///select Sharepoint site.
            if (_t == "site") {
                that.browse.path.push({ type: "site", id: $(elem).data("id") });
                that.browse.popup.nav();
                that.browse.popup.processing(true);
                that.browse.sites({ siteId: $(elem).data("id") });
            }
            ///select document library. 
            else if (_t == "library") {
                that.browse.path.push({ type: "library", id: $(elem).data("id"), site: $(elem).data("site"), url: $(elem).data("url") });
                that.browse.popup.nav();
                that.browse.popup.processing(true);
                that.browse.items({ inFolder: false, siteId: $(elem).data("site"), listId: $(elem).data("id") });
            }
            ///select folder.
            else if (_t == "folder") {
                that.browse.path.push({ type: "folder", id: $(elem).data("id"), site: $(elem).data("site"), list: $(elem).data("list"), url: $(elem).data("url") });
                that.browse.popup.nav();
                that.browse.popup.processing(true);
                that.browse.items({ inFolder: true, siteId: $(elem).data("site"), listId: $(elem).data("list"), itemId: $(elem).data("id") });
            }
            else {
                ///select the excel file.
                that.browse.popup.hide();
                that.files({ path: that.browse.path[that.browse.path.length - 1].url + "/" + encodeURI($(elem).text()), name: $(elem).text() });
            }
        },

        ///display the file explorer popup.
        popup: {
            ///Reset to popup default state (an empty popup)
            dft: function () {
                that.controls.popupBrowseList.html("");
                that.controls.popupBrowseBack.hide();
                that.controls.popupBrowseMessage.html("").hide();
                that.controls.popupBrowseLoading.hide();
            },
            ///show the file explorer popup.
            show: function () {
                that.controls.popupMain.removeClass("message process confirm").addClass("active browse");
            },
            ////hide the explorer popup
            hide: function () {
                that.controls.popupMain.removeClass("active message process confirm browse");
            },
            ///display the loading before popup displays.
            processing: function (show) {
                if (show) {
                    that.controls.popupBrowseLoading.show();
                }
                else {
                    that.controls.popupBrowseLoading.hide();
                }
            },
            ///display the prompted message in the popup.
            message: function (txt) {
                that.controls.popupBrowseMessage.html(txt).show();
                that.browse.popup.processing(false);
            },
            ///go back
            back: function () {
                that.browse.path.pop();
                if (that.browse.path.length > 0) {
                    var _ip = that.browse.path[that.browse.path.length - 1];
                    if (_ip.type == "site") {
                        that.browse.popup.processing(true);
                        that.browse.sites({ siteId: _ip.id });
                    }
                    else if (_ip.type == "library") {
                        that.browse.popup.processing(true);
                        that.browse.items({ inFolder: false, siteId: _ip.site, listId: _ip.id });
                    }
                    else if (_ip.type == "folder") {
                        that.browse.popup.processing(true);
                        that.browse.items({ inFolder: true, siteId: _ip.site, listId: _ip.list, itemId: _ip.id });
                    }
                }
                else {
                    that.browse.popup.processing(true);
                    that.browse.siteCollection();
                }
                that.browse.popup.nav();
            },
            ///display the back button or not.
            nav: function () {
                that.browse.path.length > 0 ? that.controls.popupBrowseBack.show() : that.controls.popupBrowseBack.hide();
            }
        }
    };

    that.format = {
        ///Change the original value to formatted value.
        convert: function (options) {
            var _t = $.trim(options.value),
                _v = _t,
                _f = options.formats ? options.formats : [],
                _d = that.format.hasDollar(_v),
                _c = that.format.hasComma(_v),
                _p = that.format.hasPercent(_v),
                _m = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
            $.each(_f, function (_a, _b) {
                if (_b.Name == "ConvertToHundreds") {
                    if (that.format.isNumber(_v)) {
                        var _l = that.format.getDecimalLength(_v);
                        _v = new BigNumber(that.format.toNumber(_v)).div(100).toString();
                        _v = that.format.AddDecimal(_v, _l);
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
                        _v = that.format.AddDecimal(_v, _l);
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
                        _v = that.format.AddDecimal(_v, _l);
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
                        _v = that.format.AddDecimal(_v, _l);
                        if (_c) {
                            _v = that.format.addComma(_v);
                        }
                        if (_d) {
                            _v = that.format.addDollar(_v);
                        }
                    }
                }
                else if (_b.Name == "AddDecimalPlace") {
                    if (that.format.isNumber(_v)) {
                        if (_v.indexOf(".") > -1) {
                            _v = _v + "0";
                        }
                        else {
                            _v = _v + ".0";
                        }
                    }
                }
                else if (_b.Name == "ShowNegativesAsPositives") {
                    var _tt = _v.toString().replace(/\$/g, "").replace(/-/g, "").replace(/%/g, "").replace(/\(/g, "").replace(/\)/g, "");
                    if (that.format.isNumber(_tt)) {
                        _v = _tt;
                        if (_p) {
                            _v = _v + "%";
                        }
                        if (_d) {
                            _v = that.format.addDollar(_v);
                        }
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
                    var _h = that.format.hasDollar(_v);
                    if (_v.indexOf("-") > -1) {
                        var _tt = _v.toString().replace(/\$/g, "").replace(/-/g, "").replace(/\(/g, "").replace(/\)/g, "");
                        _v = "(" + _tt + ")";
                        if (_h) {
                            _v = that.format.addDollar(_v);
                        }
                    }
                }
            });
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
            return _v && (_v.toString().indexOf('.') != -1 ? _v.toString().replace(/(\d)(?=(\d{3})+\.)/g, function ($0, $1) {
                return $1 + ",";
            }) : _v.toString().replace(/(\d)(?=(\d{3}))/g, function ($0, $1) {
                return $1 + ",";
            }));
        },
        ///Determine if contain '%' or not in word.
        hasPercent: function (_v) {
            return _v.toString().indexOf("%") > -1;
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
        AddDecimal: function (_v, _l) {
            var _dl = that.format.getDecimalLength(_v);
            if (_l > 0 && _dl == 0) {
                _v = new BigNumber(_v).toFixed(_l);
            }
            return _v;
        }
    };

    that.popup = {
        ///dispay the message in the popup.
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
        ///diplay the file explorer.
        browse: function (show) {
            if (!show) {
                that.controls.popupMain.removeClass("active browse");
            }
            else {
                that.controls.popupMain.removeClass("message process confirm").addClass("active browse");
            }
        },
        ///hide the popup.
        hide: function (millisecond) {
            if (millisecond) {
                setTimeout(function () {
                    that.controls.popupMain.removeClass("active message");
                }, millisecond);
            } else {
                that.controls.popupMain.removeClass("active message");
            }
        },
        ///go to destination management page.
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
        ///
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

    that.service = {
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

        add: function (options, callback) {
            that.service.common({ url: that.endpoints.add, type: "POST", data: options.data, dataType: "json" }, callback);
        },
        edit: function (options, callback) {
            that.service.common({ url: that.endpoints.edit, type: "PUT", data: options.data, dataType: "json" }, callback);
        },
        catalog: function (options, callback) {
            that.service.common({ url: that.endpoints.catalog + options.path, type: "GET", dataType: "json" }, callback);
        },
        groups: function (callback) {
            that.service.common({ url: that.endpoints.groups, type: "GET", dataType: "json" }, callback);
        },
        list: function (callback) {
            that.service.common({ url: that.endpoints.list + that.filePath, type: "GET", dataType: "json" }, callback);
        },
        del: function (options, callback) {
            that.service.common({ url: that.endpoints.del + options.Id, type: "DELETE" }, callback);
        },
        deleteSelected: function (options, callback) {
            that.service.common({ url: that.endpoints.deleteSelected, type: "POST", data: options.data }, callback);
        },
        token: function (callback) {
            that.service.common({ url: that.endpoints.token, type: "GET", dataType: "json" }, callback);
        },
        hostName: function (callback) {
            that.service.common({ url: that.endpoints.graph + "/sites/root", type: "GET", dataType: "json", headers: { "authorization": "Bearer " + that.browse.accessToken } }, callback);
        },
        siteCollection: function (options, callback) {
            that.service.common({ url: that.endpoints.graph + "/sites/" + that.browse.host + ":/" + options.path, type: "GET", dataType: "json", headers: { "authorization": "Bearer " + that.browse.accessToken } }, callback);
        },
        sites: function (options, callback) {
            that.service.common({ url: that.endpoints.graph + "/sites/" + options.siteId + "/sites", type: "GET", dataType: "json", headers: { "authorization": "Bearer " + that.browse.accessToken } }, callback);
        },
        libraries: function (options, callback) {
            that.service.common({ url: that.endpoints.graph + "/sites/" + options.siteId + "/lists", type: "GET", dataType: "json", headers: { "authorization": "Bearer " + that.browse.accessToken } }, callback);
        },
        items: function (options, callback) {
            that.service.common({ url: that.endpoints.graph + "/sites/" + options.siteId + "/lists/" + options.listId + "/drive/root/children", type: "GET", dataType: "json", headers: { "authorization": "Bearer " + that.browse.accessToken } }, callback);
        },
        itemsInFolder: function (options, callback) {
            that.service.common({ url: that.endpoints.graph + "/sites/" + options.siteId + "/lists/" + options.listId + "/drive/items/" + options.itemId + "/children", type: "GET", dataType: "json", headers: { "authorization": "Bearer " + that.browse.accessToken } }, callback);
        },
        customFormat: function (callback) {
            that.service.common({ url: that.endpoints.customFormat, type: "GET", dataType: "json" }, callback);
        }
    };

    that.ui = {
        ///Clear the file & keyword textboxes.
        clear: function () {
            that.controls.file.val("");
            that.controls.keyword.val("");
        },
        ///set the default value for file & eyword textboxes.
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
        ///display the default screen when add destination point.
        reset: function () {
            that.ui.clear();
            that.ui.dft();
            that.ui.status({ next: false });
            that.controls.main.removeClass("step-second").addClass("step-first");
            that.controls.stepFirstMain.removeClass("selected-file");
            that.controls.filterMain.removeClass("open-filter");
        },
        ///build filter by groups HTML
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
                        _d[i].changed = __i > -1 ? $.trim(result.data[__i].text) != that.format.convert({ value: that.utility.toString(d.ReferencedSourcePoint.Value), formats: d.CustomFormats }) : false;
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
        ///build each destination point HTML in management page.
        item: function (options, callback) {
            if (options.index < options.data.length) {
                var _dsp = options.data[options.index], _item = _dsp.ReferencedSourcePoint, _sourcePointCatalog = _item.Catalog, _s = _dsp.existed && _item.Status == 0;
                if (options.index >= that.pagerSize * (that.pagerIndex - 1) && options.index < that.pagerSize * that.pagerIndex) {
                    var _p = that.utility.position(_item.Position),
                        _fn = that.utility.fileName(_sourcePointCatalog.Name),
                        _sel = $.inArray(_dsp.Id, options.selected) > -1;
                    var _h = '<li class="point-item' + (_s ? "" : " item-error") + '" data-id="' + _dsp.Id + '" data-range="' + _dsp.RangeId + '">';
                    _h += '<div class="point-item-line">';
                    _h += '<div class="i1"><div class="ckb-wrapper' + (_sel ? " checked" : "") + '"><input type="checkbox" ' + (_sel ? 'checked="checked"' : '') + ' /><i></i></div></div>';
                    _h += '<div class="i2"><span class="s-name" title="' + _item.Name + '">' + _item.Name + '</span>';
                    _h += '<span><strong title="' + (_p.sheet ? _p.sheet : "") + ':[' + (_p.cell ? _p.cell : "") + ']">' + (_p.sheet ? _p.sheet : "") + ':</strong>[' + (_p.cell ? _p.cell : "") + ']</span>';
                    _h += '<span><strong class="i-file" title="' + _sourcePointCatalog.Name + '" data-path="' + _sourcePointCatalog.Name + '">' + _fn + '</strong></span>';
                    _h += '</div>';
                    _h += '<div class="i3" title="' + (_item.Value ? _item.Value : "") + '">' + (_item.Value ? _item.Value : "") + '</div>';
                    _h += '<div class="i5"><div class="i-line"><i class="i-history" title="History"></i><i class="i-delete" title="Delete"></i></div>';
                    _h += '<div class="i-menu"><a href="javascript:"><span title="Action">...</span><span><i class="i-history" title="History"></i><i class="i-delete" title="Delete"></i></span></a></div>';
                    _h += '</div>';
                    _h += '</div>';
                    _h += '<div class="item-history"><h6>Publish History</h6><ul class="history-list">';
                    _h += '<li class="history-header"><div class="h1">Name</div><div class="h2">Value</div><div class="h3">Date</div></li>';
                    $.each(_item.PublishedHistories, function (m, n) {
                        if (m < 5) {
                            _h += '<li class="history-item"><div class="h1" title="' + n.PublishedUser + '">' + n.PublishedUser + '</div><div class="h2" title="' + (n.Value ? n.Value : "") + '">' + (n.Value ? n.Value : "") + '</div><div class="h3" title="' + that.utility.date(n.PublishedDate) + '">' + that.utility.date(n.PublishedDate) + '</div></li>';
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
                    that.range.edit({ RangeId: _dsp.RangeId, Value: that.format.convert({ value: _item.Value ? _item.Value : "", formats: _dsp.CustomFormats }) }, function (result) {
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
        customFormat: function (options) {
            that.controls.formatList.html("");
            that.controls.formatBtn.html("None");
            that.controls.formatBtn.prop("selected", "");
            that.controls.formatBtn.prop("name", "");
            if (options.data) {
                that.controls.formatList.append('<li class="drp-header"><div class="custom-cbk"><i></i></div><a href="javascript:">Select all</a></li>');
                $.each(options.data, function (i, d) {
                    var _h = '';
                    _h += '<li class="drp-item" data-id="' + d.Id + '" data-name="' + d.Name + '" title="' + d.Description + '">';
                    _h += '<div class="custom-cbk"><i></i></div>';
                    _h += '<a href="javascript:">' + d.DisplayName + '</a>';
                    _h += '</li>';
                    that.controls.formatList.append(_h);
                });
            }
        }
    };

    return that;
})();