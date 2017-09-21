$(function () {
    //Office.initialize = function (reason) {
        $(document).ready(function () {
            point.init();
        });
    //};
});

var point = (function () {
    var point = {
        ///The URL of current open excel document.
        filePath: "",
        ///The document id of the current open excel document.
        documentId: "",
        ///Define all UI controls.
        controls: {},
        ///Define source point groups.
        groups: [],
        ///Define the source points. 
        points: [],
        ///Define the source points search results.
        filteredPoints: [],
        ///Define all service endpoints.
        endpoints: {
            add: "/api/SourcePoint",
            edit: "/api/SourcePoint",
            list: "/api/SourcePointCatalog?name=",
            groups: "/api/SourcePointGroup",
            del: "/api/SourcePoint?id=",
            publish: "/api/PublishSourcePoints",
            associated: "/api/DestinationPoint?sourcePointId=",
            deleteSelected: "/api/DeleteSelectedSourcePoint",
            token: "/api/GraphAccessToken",
            sharePointToken: "/api/SharePointAccessToken",
            graph: "https://graph.microsoft.com/v1.0"
        },
        ///The selected source point.
        model: null,
        ///Determine if it is bulk or single add.
        bulk: false,
        ///Define default search keywords.
        sourcePointKeyword: "",
        ///Define default page index 
        pagerIndex: 0,
        ///Define default page size.
        pagerSize: 30,
        ///Default default page total count.
        pagerCount: 0,
        ///Define the api host and token.
        api: {
            host: "",
            token: "",
            sharePointToken: ""
        }
    }, that = point;
    ///Initialize the event handlers & load the source point list.
    that.init = function () {
        ///Get the document URL.
        that.filePath = window.location.href.indexOf("localhost") > -1 ? "https://cand3.sharepoint.com/Shared%20Documents/Book.xlsx" : Office.context.document.url;
        that.controls = {
            body: $("body"),
            main: $(".main"),
            manageNavBar: new fabric['CommandBar']($(".nav-header").get(0)),
            addNavBar: new fabric['CommandBar']($(".nav-header").get(1)),
            back: $(".n-back"),
            add: $(".n-add"),
            publish: $(".n-publish, .ms-ContextualMenu-item:has(.ms-Icon--Upload)"),
            publishAll: $(".n-publishall, .ms-ContextualMenu-item:has(.ms-Icon--publishAll)"),
            refresh: $(".n-refresh, .ms-ContextualMenu-item:has(.ms-Icon--Refresh)"),
            del: $(".n-delete, .ms-ContextualMenu-item:has(.ms-Icon--Delete)"),
            bulk: $(".n-bulk, .ms-ContextualMenu-item:has(.ms-Icon--bulkAdd)"),
            cancel: $("#btnCancel"),
            save: $("#btnSave"),
            name: $("#txtName"),
            position: $("#txtLocation"),
            groups: $("#groupWrapper"),
            select: $("#btnLocation"),
            selectName: $("#btnSelectName"),
            list: $("#listPoints"),
            documentIdError: $("#lblDocumentIDError"),
            documentIdReload: $("#btnDocumentIDReload"),
            headerListPoints: $("#headerListPoints"),
            sourcePointName: $("#txtSearchSourcePoint"),
            searchSourcePoint: $("#iSearchSourcePoint"),
            autoCompleteControl2: $("#autoCompleteWrap2"),
            listAssociated: $("#listAssociated"),
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
        that.controls.add.click(function () {
            that.action.add();
        });
        that.controls.back.click(function () {
            that.action.back();
        });
        that.controls.publish.click(function () {
            that.action.publish();
        });
        that.controls.publishAll.click(function () {
            that.action.publishAll();
        });
        that.controls.refresh.click(function () {
            that.action.refresh();
        });
        that.controls.del.click(function () {
            that.action.deleteSelected();
        });
        that.controls.bulk.click(function () {
            that.action.bulk();
        });
        that.controls.selectName.click(function () {
            that.action.select({ input: that.controls.name });
        });
        that.controls.select.click(function () {
            that.action.select({ input: that.controls.position });
        });
        that.controls.save.click(function () {
            that.action.save();
        });
        that.controls.cancel.click(function () {
            that.action.back();
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
        that.controls.documentIdReload.click(function () {
            window.location.reload();
        });
        that.controls.list.on("click", ".i-history", function () {
            that.action.history($(this).closest(".point-item"));
            return false;
        });
        that.controls.list.on("click", ".i-delete", function () {
            that.action.del($(this).closest(".point-item").data("id"), $(this).closest(".point-item"));
            return false;
        });
        that.controls.list.on("click", ".i-edit", function () {
            that.action.edit($(this).closest(".point-item").data("id"), $(this).closest(".point-item"));
            return false;
        });
        that.controls.list.on("click", ".point-item .i2, .point-item .i3, .point-item .i4, .point-item .i5, .point-item .error-info, .point-item .item-history", function () {
            that.action.goto($(this).closest(".point-item").data("id"), $(this).closest(".point-item"));
        });
        that.controls.main.on("change", ".ckb-wrapper input", function () {
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
        ///Retrieve the document id via document url.
        that.document.init(function () {
            ///Load all source points in management page.
            that.list({ refresh: false, index: 1 }, function (result) {
                if (result.status == app.status.failed) {
                    ///Dipslay the error message if failed to get source point list.
                    that.popup.message({ success: false, title: result.error.statusText });
                }
            });
        });
    };

    ///Initialize the fabric components
    that.fabric = {
        init: function () {
            $(document).on("click", ".ms-ContextualMenu-item:has(.ms-Icon--Back)", function() {
                that.controls.back.click();
            });
            $(document).on("click", ".ms-ContextualMenu-item:has(.ms-Icon--Add)", function() {
                that.controls.add.click();
            });
            $(document).on("click", ".ms-ContextualMenu-item:has(.ms-Icon--Upload)", function() {
                that.controls.publish.click();
            });
            $(document).on("click", ".ms-ContextualMenu-item:has(.ms-Icon--publishAll)", function() {
                that.controls.publishAll.click();
            });
            $(document).on("click", ".ms-ContextualMenu-item:has(.ms-Icon--Refresh)", function() {
                that.controls.refresh.click();
            });
            $(document).on("click", ".ms-ContextualMenu-item:has(.ms-Icon--Delete)", function() {
                that.controls.del.click();
            });
            $(document).on("click", ".ms-ContextualMenu-item:has(.ms-Icon--bulkAdd)", function() {
                that.controls.bulk.click();
            });
        }
    };
    ///Load the source point list.
    that.list = function (options, callback) {
        ///Display the processing layer.
        that.popup.processing(true);
        that.service.list(function (result) {
            that.popup.processing(false);
            if (result.status == app.status.succeeded) {
                if (result.data) {
                    that.points = result.data.SourcePoints;
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
    ///Default display when add/edit/bulk add the source point.
    that.default = function (callback) {
        if (that.groups.length == 0) {
            that.popup.processing(true);
            that.service.groups(function (result) {
                if (result.status == app.status.succeeded) {
                    that.popup.processing(false);
                    that.groups = result.data;
                    that.controls.name.val(that.model ? that.model.NamePosition : "");
                    that.controls.position.val(that.model ? that.model.Position : "");
                    that.ui.groups({ data: that.groups, selected: that.model ? that.model.Groups : [] });
                    that.controls.main.removeClass("manage add edit bulk").addClass(that.model ? "add edit" : (that.bulk ? "add bulk" : "add"));
                    that.controls.addNavBar._doResize();
                    if (callback) {
                        callback();
                    }
                }
                else {
                    that.popup.message({ success: false, title: "Load groups failed." });
                }
            });
        }
        else {
            that.controls.name.val(that.model ? that.model.NamePosition : "");
            that.controls.position.val(that.model ? that.model.Position : "");
            that.ui.groups({ data: that.groups, selected: that.model ? that.model.Groups : [] });
            that.controls.main.removeClass("manage add edit bulk").addClass(that.model ? "add edit" : (that.bulk ? "add bulk" : "add"));
            that.controls.addNavBar._doResize();
            if (callback) {
                callback();
            }
        }
    };
    ///The utility methods.
    that.utility = {
        ///Get current source point model.
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
        ///Get current open excel document mode.
        mode: function (callback) {
            if (Office.context.document.mode == Office.DocumentMode.ReadOnly) {
                that.popup.message({ success: false, title: "Please click \"edit workbook\" button under the excel ribbon." });
            }
            else {
                callback();
            }
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
        ///Get the value of the required fields.
        entered: function () {
            var name = $.trim(that.controls.name.val()), position = $.trim(that.controls.position.val()), groups = [];
            that.controls.groups.find("input").each(function (i, d) {
                if ($(d).prop("checked")) {
                    groups.push(d.value);
                }
            });
            return { name: name, position: position, groups: groups };
        },
        ///Determine if there is any change when edit the source point.
        changed: function (callback) {
            var entered = that.utility.entered();
            if (that.model) {
                var _g = [];
                $.each(that.model.Groups, function (i, d) { _g.push(d.Id); });
                callback({ changed: !(that.model.NamePosition == entered.name && that.model.Position == entered.position && _g.sort().toString() == entered.groups.sort().toString()) });
            }
            else {
                callback({ changed: !(entered.name.length == 0 && entered.position.length == 0 && entered.groups.length == 0) });
            }
        },
        ///When add source point, it can only select a single cell
        ///When bulk add source points. it only supports selecting 2 adjacent columns 
        valid: function (options, callback) {
            var _p = that.utility.position(options.position);
            if ((!that.bulk && _p.cell.indexOf(":") == -1) || that.bulk) {
                Excel.run(function (ctx) {
                    var r = ctx.workbook.worksheets.getItem(_p.sheet).getRange(_p.cell);
                    r.load("address,text");
                    return ctx.sync().then(function () {
                        if (!that.bulk) {
                            callback({ status: app.status.succeeded, data: r.text[0][0] ? $.trim(r.text[0][0]) : "", address: r.address[0][0] });
                        }
                        else {
                            if (r.text.length > 0 && r.text[0].length == 2) {
                                callback({ status: app.status.succeeded, data: r.text, address: r.address });
                            }
                            else {
                                callback({ status: app.status.failed, message: "Only 2 adjacent columns can be selected." });
                            }
                        }
                    });
                }).catch(function (error) {
                    callback({ status: app.status.failed, message: "The selected " + options.title + " position is invalid." });
                });
            }
            else {
                callback({ status: app.status.failed, message: "Only 1 cell can be selected for " + options.title + "." });
            }
        },
        ///Check if the source point existed or not by source point name.
        exist: function (options, callback) {
            var a = [], existed = false;
            if (options.name) {
                $.each(that.points, function (i, d) {
                    a.push(d.Name);
                });
                if (that.model) {
                    existed = that.model.Name != options.name && $.inArray(options.name, a) > -1;
                }
                else {
                    existed = $.inArray(options.name, a) > -1;
                }
            }
            callback({ status: existed ? app.status.failed : app.status.succeeded, message: existed ? (options.name + " already exists, please input a different name.") : "" });
        },
        ///Get the entered value for the required fields and validate the required fields.
        ///The source point name could not exceed 255 characters.
        ///Determine if the position already existed in Azure storage or not.
        validation: function (callback) {
            var entered = that.utility.entered(),
                rangeId = that.model ? that.model.RangeId : app.guid(),
                nameRangeId = that.model ? that.model.NameRangeId : app.guid(),
                namePosition = entered.name,
                position = entered.position,
                groups = entered.groups,
                success = true,
                values = [];
            if (!that.bulk && namePosition.length == 0) {
                success = false;
                values.push(["Source Point Name"]);
            }
            if (position.length == 0) {
                success = false;
                values.push(["Select range"]);
            }
            if (!success) {
                callback({ status: success ? app.status.succeeded : app.status.failed, message: success ? "" : { success: success, title: "Please enter the following required fields:", values: values } });
            }
            else {
                if (!that.bulk) {
                    that.utility.valid({ position: namePosition, title: "source point name" }, function (result) {
                        if (result.status == app.status.succeeded) {
                            var name = result.data;
                            if (name.length > 0 && name.length <= 255) {
                                that.utility.exist({ name: name }, function (result) {
                                    if (result.status == app.status.succeeded) {
                                        that.utility.valid({ position: position, title: "range" }, function (result) {
                                            if (result.status == app.status.succeeded) {
                                                callback({
                                                    status: app.status.succeeded,
                                                    data: {
                                                        Id: that.model ? that.model.Id : "",
                                                        Name: name,
                                                        CatalogName: that.filePath,
                                                        DocumentId: that.documentId,
                                                        RangeId: rangeId,
                                                        NameRangeId: nameRangeId,
                                                        NamePosition: namePosition,
                                                        Position: position,
                                                        Value: result.data,
                                                        GroupIds: groups
                                                    }
                                                });
                                            }
                                            else {
                                                callback({ status: app.status.failed, message: { success: false, title: result.message } });
                                            }
                                        });
                                    }
                                    else {
                                        callback({ status: app.status.failed, message: { success: false, title: result.message } });
                                    }
                                });
                            }
                            else {
                                callback({ status: app.status.failed, message: { success: false, title: name.length > 0 ? "The source point name cannot exceed 255 characters." : "The source point name cannot be blank." } });
                            }
                        }
                        else {
                            callback({ status: app.status.failed, message: { success: false, title: result.message } });
                        }
                    });
                }
                else {
                    that.utility.valid({ position: position, title: "range" }, function (result) {
                        if (result.status == app.status.succeeded) {
                            callback({
                                status: app.status.succeeded,
                                data: {
                                    Id: that.model ? that.model.Id : "",
                                    Name: "",
                                    CatalogName: that.filePath,
                                    DocumentId: that.documentId,
                                    RangeId: rangeId,
                                    NameRangeId: "",
                                    NamePosition: "",
                                    Position: position,
                                    Value: result.data,
                                    GroupIds: groups
                                }
                            });
                        }
                        else {
                            callback({ status: app.status.failed, message: { success: false, title: result.message } });
                        }
                    });
                }
            }
        },
        ///Get the index of current source point.
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
        ///Add source point to source point list.
        add: function (options) {
            that.points.push(options);
        },
        ///Remove the source point from the array.
        remove: function (options) {
            that.points.splice(that.utility.index(options), 1);
        },
        ///Update the source point in the array.
        update: function (options) {
            that.points[that.utility.index(options)] = options;
        },
        ///Get an array of selected source points.
        selected: function () {
            var _s = [];
            that.controls.list.find(".point-item .ckb-wrapper input").each(function (i, d) {
                if ($(d).prop("checked")) {
                    _s.push({
                        SourcePointId: $(d).closest(".point-item").data("id"),
                        Name: $.trim($(d).closest(".point-item").find(".i2 .s-name").text()),
                        RangeId: $(d).closest(".point-item").data("range"),
                        CurrentValue: $(d).closest(".point-item").find(".i4 .s-value").text(),
                        Position: $(d).closest(".point-item").data("position"),
                        NameRangeId: $(d).closest(".point-item").data("namerange"),
                        NamePosition: $(d).closest(".point-item").data("nameposition")
                    });
                }
            });
            return _s;
        },
        ///Get all filtered source points.
        all: function () {
            var _s = [];
            $.each(that.filteredPoints, function (i, d) {
                _s.push({
                    SourcePointId: d.Id,
                    Name: d.Name,
                    RangeId: d.RangeId,
                    CurrentValue: d.Value,
                    Position: d.Position,
                    NameRangeId: d.NameRangeId,
                    NamePosition: d.NamePosition
                });
            });
            return _s;
        },
        ///Get the file name.
        fileName: function (path) {
            return path.lastIndexOf("/") > -1 ? path.substr(path.lastIndexOf("/") + 1) : (path.lastIndexOf("\\") > -1 ? path.substr(path.lastIndexOf("\\") + 1) : path);
        },
        ///Get latest the range value from Azure storage..
        value: function (options, callback) {
            if (options.index < options.data.length) {
                that.range.exist({ RangeId: options.data[options.index].NameRangeId }, function (ret) {
                    that.range.exist({ RangeId: options.data[options.index].RangeId }, function (result) {
                        if (result.status == app.status.succeeded && ret.status == app.status.succeeded) {
                            options.data[options.index].DocumentValue = result.data.text;
                            options.data[options.index].DocumentPosition = result.data.address;
                            options.data[options.index].DocumentNameValue = ret.data.text;
                            options.data[options.index].DocumentNamePosition = ret.data.address;
                            options.data[options.index].Existed = true;
                        }
                        else {
                            options.data[options.index].Existed = false;
                        }
                        options.index++;
                        that.utility.value(options, callback);
                    });
                });
            }
            else {
                callback(options);
            }
        },
        ///It is dirty if the current value or position is different with document value or position.
        dirty: function (options, callback) {
            var _f = false;
            $.each(options.data, function (i, d) {
                if (d.Existed && (d.CurrentValue != d.DocumentValue || d.Position != d.DocumentPosition || d.Name != d.DocumentNameValue || d.NamePosition != d.DocumentNamePosition)) {
                    _f = true;
                    return false;
                }
            });
            callback(_f);
        },
        ///Paging feature for the source point list.
        pager: {
            ///Initialize the source point list UI.
            init: function (options) {
                that.controls.pagerValue.val("");
                that.pagerIndex = options.index && options.index > 0 ? options.index : 1;
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
        ///Get all addresses for the selected ranges.
        addresses: function (options, callback) {
            if (options.index == undefined) {
                options.index = 0;
                options.position = that.utility.position(options.Position);
                options.cells = [];
                options.addresses = [];
                options.existed = [];
                options.texts = [];
                for (var i = 0; i < options.Value.length; i++) {
                    for (var m = 0; m < options.Value[i].length; m++) {
                        options.cells.push({ row: i, cell: m });
                    }
                }
            }
            if (options.index < options.cells.length / 2) {
                Excel.run(function (ctx) {
                    var r = ctx.workbook.worksheets.getItem(options.position.sheet).getRange(options.position.cell);
                    var c = r.getCell(options.cells[2 * options.index].row, options.cells[2 * options.index].cell);
                    c.load("text,address");
                    return ctx.sync().then(function () {
                        if (c.text && $.trim(c.text[0][0]) != "" && $.trim(c.text[0][0]).length <= 255) {
                            var _t = $.trim(c.text[0][0]), _a = c.address;
                            that.utility.exist({ name: _t }, function (result) {
                                if (result.status == app.status.succeeded) {
                                    var cc = r.getCell(options.cells[2 * options.index + 1].row, options.cells[2 * options.index + 1].cell);
                                    cc.load("text,address");
                                    return ctx.sync().then(function () {
                                        if ($.inArray(_t, options.texts) == -1) {
                                            options.addresses.push({ title: _t, text: $.trim(cc.text[0][0]), address: cc.address, nameAddress: _a });
                                            options.texts.push(_t);
                                        }
                                        else {
                                            options.existed.push(_t);
                                        }
                                        options.index++;
                                        that.utility.addresses(options, callback);
                                    });
                                }
                                else {
                                    options.existed.push(_t);
                                    options.index++;
                                    that.utility.addresses(options, callback);
                                }
                            });
                        }
                        else {
                            options.index++;
                            that.utility.addresses(options, callback);
                        }
                    });
                }).catch(function (error) {
                    options.index++;
                    that.utility.addresses(options, callback);
                });
            }
            else {
                callback(options);
            }
        },
        ///UnSelect all checkboxes.
        unSelectAll: function () {
            that.controls.headerListPoints.find(".point-header .ckb-wrapper input").prop("checked", false);
            that.controls.headerListPoints.find(".point-header .ckb-wrapper").removeClass("checked");
        },
        ///Calculate the height source point list display area.
        height: function () {
            if (that.controls.main.hasClass("manage")) {
                var _h = that.controls.main.outerHeight();
                var _h1 = $("#pager").outerHeight();
                that.controls.list.css("maxHeight", (_h - 205 - 70 - _h1) + "px");
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
        }
    };
    ///Define all the event handlers.
    that.action = {
        ///Body click event to hide the suggested search result.
        body: function () {
            $(".search-tooltips").hide();
        },
        ///Add source point.
        add: function () {
            that.utility.mode(function () {
                that.model = null;
                that.bulk = false;
                that.default();
                that.controls.name.focus();
            });
        },
        ///Bulk Add source point
        bulk: function () {
            that.utility.mode(function () {
                that.model = null;
                that.bulk = true;
                that.default();
                that.controls.name.focus();
            });
        },
        ///Back to management page.
        back: function () {
            that.utility.mode(function () {
                that.utility.changed(function (result) {
                    if (result.changed) {
                        that.popup.confirm({
                            title: "Do you want to save your changes?"
                        },
                        function () {
                            that.controls.popupMain.removeClass("message process confirm active");
                            that.controls.save.click();
                        }, function () {
                            that.controls.popupMain.removeClass("message process confirm active");
                            that.controls.main.removeClass("add edit bulk").addClass("manage");
                            that.controls.manageNavBar._doResize();
                        });
                    }
                    else {
                        that.controls.main.removeClass("add edit bulk").addClass("manage");
                        that.controls.manageNavBar._doResize();
                    }
                });
            });
        },
        ///Refresh the source points in management page.
        refresh: function () {
            that.utility.mode(function () {
                that.points = [];
                that.list({ refresh: true, index: that.pagerIndex }, function (result) {
                    if (result.status == app.status.failed) {
                        that.popup.message({ success: false, title: result.error.statusText });
                    }
                    else {
                        that.popup.message({ success: true, title: "Refresh all source points succeeded." }, function () { that.popup.hide(3000); });
                    }
                });
            });
        },
        ///Select the source points in management page.
        select: function (options) {
            that.utility.mode(function () {
                that.range.select(function (result) {
                    if (result.status == app.status.succeeded) {
                        options.input.val(result.data);
                    }
                    else {
                        that.popup.message({ success: false, title: result.message });
                    }
                });
            });
        },
        ///Save the source point in 'add/bulk add source point' pages.
        save: function () {
            that.utility.mode(function () {
                if (that.model) {
                    that.utility.validation(function (result) {
                        if (result.status == app.status.succeeded) {
                            that.range.del({ RangeId: result.data.NameRangeId }, function (ret) {
                                if (ret.status == app.status.succeeded) {
                                    that.range.create({ Position: result.data.NamePosition, RangeId: result.data.NameRangeId }, function (ret) {
                                        if (ret.status == app.status.succeeded) {
                                            that.range.del({ RangeId: result.data.RangeId }, function (ret) {
                                                if (ret.status == app.status.succeeded) {
                                                    that.range.create({ Position: result.data.Position, RangeId: result.data.RangeId }, function (ret) {
                                                        if (ret.status == app.status.succeeded) {
                                                            that.popup.processing(true);
                                                            that.service.edit({ data: result.data }, function (result) {
                                                                if (result.status == app.status.succeeded) {
                                                                    that.utility.update(result.data);
                                                                    that.utility.pager.init({ refresh: true, index: that.pagerIndex });
                                                                    that.popup.message({ success: true, title: "Source point update complete." }, function () { that.popup.back(3000); });
                                                                }
                                                                else {
                                                                    that.popup.message({ success: false, title: "Edit source point failed." });
                                                                }
                                                            });
                                                        }
                                                        else {
                                                            that.popup.message({ success: false, title: "Create range in Excel failed." });
                                                        }
                                                    });
                                                }
                                                else {
                                                    that.popup.message({ success: false, title: "Delete the previous range failed." });
                                                }
                                            });
                                        }
                                        else {
                                            that.popup.message({ success: false, title: "Create source point name range in Excel failed." });
                                        }
                                    });
                                }
                                else {
                                    that.popup.message({ success: false, title: "Delete the previous source point name range failed." });
                                }
                            });
                        }
                        else {
                            that.popup.message(result.message);
                        }
                    });
                }
                else {
                    that.utility.validation(function (result) {
                        if (result.status == app.status.succeeded) {
                            if (!that.bulk) {
                                that.range.create({ Position: result.data.NamePosition, RangeId: result.data.NameRangeId }, function (ret) {
                                    if (ret.status == app.status.succeeded) {
                                        that.range.create({ Position: result.data.Position, RangeId: result.data.RangeId }, function (ret) {
                                            if (ret.status == app.status.succeeded) {
                                                that.popup.processing(true);
                                                that.service.add({ data: result.data }, function (result) {
                                                    if (result.status == app.status.succeeded) {
                                                        that.utility.add(result.data);
                                                        that.utility.pager.init({ refresh: true, index: that.pagerIndex });
                                                        that.popup.message({ success: true, title: "Add new source point succeeded." }, function () { that.popup.back(3000); });
                                                    }
                                                    else {
                                                        that.popup.message({ success: false, title: "Add source point failed." });
                                                    }
                                                });
                                            }
                                            else {
                                                that.popup.message({ success: false, title: "Create range in Excel failed." });
                                            }
                                        });
                                    }
                                    else {
                                        that.popup.message({ success: false, title: "Create source point name range in Excel failed." });
                                    }
                                });
                            }
                            else {
                                that.utility.addresses(result.data, function (rst) {
                                    rst.index = undefined;
                                    that.action.bulkAdd(rst, function (rt) {
                                        that.utility.pager.init({ refresh: true, index: that.pagerIndex });
                                        if (rt.status == app.status.succeeded) {
                                            if (rst.existed.length > 0) {
                                                var _v = [];
                                                $.each(rst.existed, function (_i, _d) {
                                                    _v.push([_d]);
                                                });
                                                that.popup.message({ success: false, title: "The following Source Points already exist, please input a unique name:", values: _v }, function () { that.popup.back(); });
                                            }
                                            else {
                                                that.popup.message({ success: true, title: "Add bulk source points succeeded." }, function () { that.popup.back(3000); });
                                            }
                                        }
                                        else {
                                            that.popup.message({ success: false, title: "Add " + rt.success + " source points succeeded, add " + rt.error + "source point failed." }, function () { that.popup.back(); });
                                        }
                                    });
                                });
                            }
                        }
                        else {
                            that.popup.message(result.message);
                        }
                    });
                }
            });
        },
        ///Save an array of source points in 'bulk add source point' page.
        bulkAdd: function (options, callback) {
            if (options.index == undefined) {
                that.popup.processing(true);
                options.index = 0;
                options.error = 0;
            }
            if (options.index < options.addresses.length) {
                var _i = options.addresses[options.index], _d = {
                    Id: "",
                    Name: _i.title,
                    CatalogName: options.CatalogName,
                    DocumentId: options.DocumentId,
                    RangeId: app.guid(),
                    NameRangeId: app.guid(),
                    NamePosition: _i.nameAddress,
                    Position: _i.address,
                    Value: _i.text,
                    GroupIds: options.GroupIds
                };

                that.range.create({ Position: _d.NamePosition, RangeId: _d.NameRangeId }, function (ret) {
                    if (ret.status == app.status.succeeded) {
                        that.range.create({ Position: _d.Position, RangeId: _d.RangeId }, function (ret) {
                            if (ret.status == app.status.succeeded) {
                                that.service.add({ data: _d }, function (result) {
                                    if (result.status == app.status.succeeded) {
                                        that.utility.add(result.data);
                                        options.index++;
                                        that.action.bulkAdd(options, callback);
                                    }
                                    else {
                                        options.index++;
                                        options.error++;
                                        that.action.bulkAdd(options, callback);
                                    }
                                });
                            }
                            else {
                                options.index++;
                                options.error++;
                                that.action.bulkAdd(options, callback);
                            }
                        });
                    }
                    else {
                        options.index++;
                        options.error++;
                        that.action.bulkAdd(options, callback);
                    }
                });
            }
            else {
                callback({ status: options.error > 0 ? app.status.failed : app.status.succeeded, error: options.error, success: options.addresses.length - options.error });
            }
        },
        ///Delete current source point in management page after clicking X icon.
        del: function (i, o) {
            that.utility.mode(function () {
                that.popup.confirm({
                    title: "Do you want to delete the source point?"
                }, function () {
                    that.popup.processing(true);
                    that.service.del({
                        Id: i
                    }, function (result) {
                        if (result.status == app.status.succeeded) {
                            that.popup.message({ success: true, title: "Delete source point succeeded." }, function () { that.popup.hide(3000); });
                            that.utility.remove({ Id: i });
                            that.ui.remove({ Id: i });
                            that.utility.pager.init({ refresh: true, index: that.controls.list.find(".point-item").length > 0 ? that.pagerIndex : that.pagerIndex - 1 });
                        }
                        else {
                            that.popup.message({ success: false, title: "Delete source point failed." });
                        }
                    });
                }, function () {
                    that.controls.popupMain.removeClass("message process confirm active");
                });
            });
        },
        ///Delete the selected source points
        deleteSelected: function () {
            var _s = that.utility.selected(), _ss = [];
            if (_s && _s.length > 0) {
                $.each(_s, function (_y, _z) {
                    _ss.push(_z.SourcePointId);
                });
                that.utility.mode(function () {
                    that.popup.confirm({
                        title: "Do you want to delete the selected source point?"
                    }, function () {
                        that.popup.processing(true);
                        that.service.deleteSelected({ data: { "": _ss } }, function (result) {
                            if (result.status == app.status.succeeded) {
                                that.popup.message({ success: true, title: "Delete source point succeeded." }, function () { that.popup.hide(3000); });
                                $.each(_ss, function (_m, _n) {
                                    that.utility.remove({ Id: _n });
                                    that.ui.remove({ Id: _n });
                                });
                                that.utility.unSelectAll();
                                that.utility.pager.init({ refresh: true, index: that.controls.list.find(".point-item").length > 0 ? that.pagerIndex : that.pagerIndex - 1 });
                            }
                            else {
                                that.popup.message({ success: false, title: "Delete source point failed." });
                            }
                        });
                    }, function () {
                        that.controls.popupMain.removeClass("message process confirm active");
                    });
                });
            }
            else {
                that.popup.message({ success: false, title: "Please select source point." });
            }
        },
        ///Edit current source point after clicking pencil icon.
        edit: function (i, o) {
            that.utility.mode(function () {
                that.bulk = false;
                that.model = that.utility.model(i);
                if (that.model) {
                    that.model.NameRangeId = that.model.NameRangeId && that.model.NameRangeId != null ? that.model.NameRangeId : "";
                    that.model.NamePosition = that.model.NamePosition && that.model.NamePosition != null ? that.model.NamePosition : "";
                    that.range.goto({ RangeId: that.model.NameRangeId }, function (result) {
                        if (result.status == app.status.succeeded) {
                            that.model.NamePosition = result.data.address;
                            that.range.goto({ RangeId: that.model.RangeId }, function (result) {
                                if (result.status == app.status.succeeded) {
                                    that.model.Position = result.data.address;
                                    that.default(function () { that.action.associated(); });
                                }
                                else {
                                    that.popup.message({ success: false, title: "The range in the Excel has been deleted." });
                                }
                            });
                        }
                        else {
                            that.popup.message({ success: false, title: "The source point name range in the Excel has been deleted." });
                        }
                    });
                }
                else {
                    that.popup.message({ success: false, title: "The source point has been deleted." });
                }
            });
        },
        ///Look up published hisotry of the source point.
        history: function (o) {
            o.hasClass("item-more") ? o.removeClass("item-more") : o.addClass("item-more");
        },
        ///Publish current selected source point.
        publish: function () {
            var _s = that.utility.selected();
            if (_s && _s.length > 0) {
                that.popup.processing(true);
                that.utility.value({ index: 0, data: _s }, function (result) {
                    var _ss = [], _sf = false;
                    $.each(result.data, function (_m, _n) {
                        if (_n.Existed) {
                            _ss.push(_n);
                        }
                        else {
                            _sf = true;
                        }
                    });
                    if (_ss.length > 0) {
                        that.utility.dirty({ data: _ss }, function (f) {
                            if (f) {
                                that.popup.confirm({
                                    title: "The cell value or location is not the same with the one in Add-in. Please click 'No' to cancel it and click 'refresh' to manually sync the value or location from excel to Add-in. Are you sure to publish the selected source point still?"
                                },
                                function () {
                                    that.popup.processing(true);
                                    that.service.publish({ data: { "": _ss } }, function (result) {
                                        if (result.status == app.status.succeeded) {
                                            $.each(result.data.SourcePoints, function (_i, _d) {
                                                that.utility.update(_d);
                                            });
                                            that.ui.publish(result.data, function () {
                                                that.popup.message({ success: _sf ? false : true, title: _sf ? "Not all Source Points were published, please check Source Point list for errors." : "Publish source point succeeded." }, function () { that.popup.hide(3000); });
                                            });
                                        }
                                        else {
                                            that.popup.message({ success: false, title: "Publish source point failed." });
                                        }
                                    });
                                },
                                function () {
                                    that.controls.popupMain.removeClass("message process confirm active");
                                });
                            }
                            else {
                                that.service.publish({ data: { "": _ss } }, function (result) {
                                    if (result.status == app.status.succeeded) {
                                        $.each(result.data.SourcePoints, function (_i, _d) {
                                            that.utility.update(_d);
                                        });
                                        that.ui.publish(result.data, function () {
                                            that.popup.message({ success: _sf ? false : true, title: _sf ? "Not all Source Points were published, please check Source Point list for errors." : "Publish source point succeeded." }, function () {
                                                that.popup.hide(3000);
                                            });
                                        });
                                    }
                                    else {
                                        that.popup.message({ success: false, title: "Publish source point failed." });
                                    }
                                });
                            }
                        });
                    }
                    else {
                        that.popup.message({ success: false, title: "The source points you selected have been deleted." });
                    }
                });
            }
            else {
                that.popup.message({ success: false, title: "Please select source point." });
            }
        },
        ///Publish all source points
        publishAll: function () {
            var _s = that.utility.all();
            if (_s && _s.length > 0) {
                that.popup.processing(true);
                that.utility.value({ index: 0, data: _s }, function (result) {
                    var _ss = [], _sf = false;
                    $.each(result.data, function (_m, _n) {
                        if (_n.Existed) {
                            _ss.push(_n);
                        }
                        else {
                            _sf = true;
                        }
                    });
                    if (_ss.length > 0) {
                        that.utility.dirty({ data: _ss }, function (f) {
                            if (f) {
                                that.popup.confirm({
                                    title: "The cell value or location is not the same with the one in Add-in. Please click 'No' to cancel it and click 'refresh' to manually sync the value or location from excel to Add-in. Are you sure to publish the selected source point still?"
                                },
                                function () {
                                    that.popup.processing(true);
                                    that.service.publish({ data: { "": _ss } }, function (result) {
                                        if (result.status == app.status.succeeded) {
                                            $.each(result.data.SourcePoints, function (_i, _d) {
                                                that.utility.update(_d);
                                            });
                                            that.ui.publish(result.data, function () {
                                                that.popup.message({ success: _sf ? false : true, title: _sf ? "Not all Source Points were published, please check Source Point list for errors." : "Publish source point succeeded." }, function () {
                                                    that.popup.hide(3000);
                                                });
                                            });
                                        }
                                        else {
                                            that.popup.message({ success: false, title: "Publish source point failed." });
                                        }
                                    });
                                },
                                function () {
                                    that.controls.popupMain.removeClass("message process confirm active");
                                });
                            }
                            else {
                                that.service.publish({ data: { "": _ss } }, function (result) {
                                    if (result.status == app.status.succeeded) {
                                        $.each(result.data.SourcePoints, function (_i, _d) {
                                            that.utility.update(_d);
                                        });
                                        that.ui.publish(result.data, function () {
                                            that.popup.message({ success: _sf ? false : true, title: _sf ? "Not all Source Points were published, please check Source Point list for errors." : "Publish source point succeeded." }, function () {
                                                that.popup.hide(3000);
                                            });
                                        });
                                    }
                                    else {
                                        that.popup.message({ success: false, title: "Publish source point failed." });
                                    }
                                });
                            }
                        });
                    }
                    else {
                        that.popup.message({ success: false, title: "All source points have been deleted." });
                    }
                });
            }
            else {
                that.popup.message({ success: false, title: "There is no source point." });
            }
        },
        ///Go to the range of the source point.
        goto: function (i, o) {
            that.utility.mode(function () {
                var _m = that.utility.model(i);
                if (_m) {
                    that.range.exist({ RangeId: _m.NameRangeId }, function (result) {
                        if (result.status == app.status.failed) {
                            that.popup.message({ success: false, title: "The source point name range in the Excel has been deleted." });
                        }
                        else {
                            that.range.goto({ RangeId: _m.RangeId }, function (result) {
                                if (result.status == app.status.failed) {
                                    that.popup.message({ success: false, title: "The range in the Excel has been deleted." });
                                }
                            });
                        }
                    });
                }
                else {
                    that.popup.message({ success: false, title: "The point has been deleted." });
                }
            });
        },
        ///Close the popup.
        ok: function () {
            that.controls.popupMain.removeClass("active message process confirm");
        },
        ///Get the referenced word files.
        associated: function () {
            that.service.associated(that.model, function (result) {
                if (result.status == app.status.succeeded) {
                    that.ui.associated({ data: result.data });
                }
                else {
                    that.popup.message({ success: false, title: "Get associated files failed." });
                }
            });
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
            that.utility.pager.init({ refresh: true });
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
            that.controls.popupMain.removeClass("process confirm").addClass("active message");
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
        processing: function (hide) {
            if (!hide) {
                that.controls.popupMain.removeClass("active process");
            }
            else {
                that.controls.popupMain.removeClass("message confirm").addClass("active process");
            }
        },
        ///Dipslay yes/no confirmation popup.
        confirm: function (options, yesCallback, noCallback) {
            that.controls.popupConfirmTitle.html(options.title);
            that.controls.popupMain.removeClass("message process").addClass("active confirm");
            that.controls.popupConfirmYes.unbind("click").click(function () {
                yesCallback();
            });
            that.controls.popupConfirmNo.unbind("click").click(function () {
                noCallback();
            });
        },
        ///Hide the popup in millisecond.
        hide: function (millisecond) {
            if (millisecond) {
                setTimeout(function () {
                    that.controls.popupMain.removeClass("active message");
                }, millisecond);
            } else {
                that.controls.popupMain.removeClass("active message");
            }
        },
        ///Display management page and hide the add/bulk add page.
        back: function (millisecond) {
            if (millisecond) {
                setTimeout(function () {
                    that.controls.popupMain.removeClass("active message");
                    that.controls.main.removeClass("manage add edit bulk").addClass("manage");
                    that.controls.manageNavBar._doResize();
                }, millisecond);
            }
            else {
                that.controls.popupMain.removeClass("active message");
                that.controls.main.removeClass("manage add edit bulk").addClass("manage");
                that.controls.manageNavBar._doResize();
            }
        }
    };
    ///Define all range event handlers.
    that.range = {
        ///Create the range (source point).
        create: function (options, callback) {
            Excel.run(function (ctx) {
                var p = that.utility.position(options.Position), r = ctx.workbook.worksheets.getItem(p.sheet).getRange(p.cell);
                ctx.workbook.bindings.add(r, Excel.BindingType.range, options.RangeId);
                return ctx.sync().then(function () {
                    callback({ status: app.status.succeeded });
                });
            }).catch(function (error) {
                callback({ status: app.status.failed, message: error.message });
            });
        },
        ///Select the address of the range.
        select: function (callback) {
            Excel.run(function (ctx) {
                var r = ctx.workbook.getSelectedRange();
                r.load("address");
                return ctx.sync().then(function () {
                    callback({ status: app.status.succeeded, data: r.address });
                });
            }).catch(function (error) {
                callback({ status: app.status.failed, message: "Get selected cell failed." });
            });
        },
        ///Determine if the range existed or not.
        exist: function (options, callback) {
            Excel.run(function (ctx) {
                var r = ctx.workbook.bindings.getItem(options.RangeId).getRange();
                r.load("text,address");
                return ctx.sync().then(function () {
                    callback({ status: app.status.succeeded, data: { text: $.trim(r.text[0][0]), address: r.address } });
                });
            }).catch(function (error) {
                callback({ status: app.status.failed, message: error.message });
            });
        },
        ///Go to the range by range ID.
        goto: function (options, callback) {
            Excel.run(function (ctx) {
                var r = ctx.workbook.bindings.getItem(options.RangeId).getRange();
                r.select();
                r.load("text,address");
                return ctx.sync().then(function () {
                    callback({ status: app.status.succeeded, data: { text: $.trim(r.text[0][0]), address: r.address } });
                });
            }).catch(function (error) {
                callback({ status: app.status.failed, message: error.message });
            });
        },
        ///Delete the range by range ID.
        del: function (options, callback) {
            Excel.run(function (ctx) {
                ctx.workbook.bindings.getItem(options.RangeId).delete();
                return ctx.sync().then(function () {
                    callback({ status: app.status.succeeded });
                });
            }).catch(function (error) {
                callback({ status: app.status.failed, message: error.message });
            });
        },
        ///Get properties of the ranges.
        all: function (callback) {
            Excel.run(function (ctx) {
                var _b = ctx.workbook.bindings, _c = [];
                _b.load("items");
                return ctx.sync().then(function () {
                    for (var i = 0; i < _b.items.length; i++) {
                        var _r = _b.items[i].getRange();
                        _r.load("text,address");
                        return ctx.sync().then(function () {
                            _c.push({ id: _b.items[i].id, text: $.tirm(_r.text), address: _r.address });
                        });
                    }
                    callback({ status: app.status.succeeded, data: _c });
                });
            }).catch(function (error) {
                callback({ status: app.status.failed, message: "Get all bindings failed. " + error.message });
            });
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
        ///Add source point (POST) to Azure storage.
        add: function (options, callback) {
            that.service.common({ url: that.endpoints.add, type: "POST", data: options.data, dataType: "json" }, callback);
        },
        ///Update the source point (PUT) in Azure storage.
        edit: function (options, callback) {
            that.service.common({ url: that.endpoints.edit, type: "PUT", data: options.data, dataType: "json" }, callback);
        },
        ///Get the source point groups.
        groups: function (callback) {
            that.service.common({ url: that.endpoints.groups, type: "GET", dataType: "json" }, callback);
        },
        ///Get a list of source points from Azure storage.
        list: function (callback) {
            that.service.common({ url: that.endpoints.list + that.filePath + "&documentId=" + that.documentId, type: "GET", dataType: "json" }, callback);
        },
        ///Delete current source point in Azure storage.
        del: function (options, callback) {
            that.service.common({ url: that.endpoints.del + options.Id, type: "DELETE" }, callback);
        },
        ///Publish the source points in Azure storage.
        publish: function (options, callback) {
            that.service.common({ url: that.endpoints.publish, type: "POST", data: options.data, dataType: "json" }, callback);
        },
        ///Get referenced word document URLs from Azure storage.
        associated: function (options, callback) {
            that.service.common({ url: that.endpoints.associated + options.Id, type: "GET", dataType: "json" }, callback);
        },
        ///Delete the selected source points in Azure storage.
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
        ///Get libraries under the current site.
        libraries: function (options, callback) {
            that.service.common({ url: that.endpoints.graph + "/sites/" + options.siteId + "/drives", type: "GET", dataType: "json", headers: { "authorization": "Bearer " + that.api.token } }, callback);
        },
        ///Search list item by file name.
        item: function (options, callback) {
            that.service.common({ url: options.siteUrl + "/_api/web/lists/getbytitle('" + options.listName + "')/items?$select=FileLeafRef,EncodedAbsUrl,OData__dlc_DocId&$filter=FileLeafRef eq '" + options.fileName + "'", type: "GET", dataType: "json", headers: { "authorization": "Bearer " + that.api.sharePointToken } }, callback);
        }
    };
    ///Build the HTML/UI.
    that.ui = {
        ///Build the groups checkbox list.
        groups: function (options) {
            that.controls.groups.html("");
            var _s = [];
            $.each(options.selected, function (i, d) {
                _s.push(d.Id);
            });
            $.each(options.data, function (i, d) {
                $("<li><div><div class=\"ckb-wrapper" + ($.inArray(d.Id, _s) > -1 ? " checked" : "") + "\"><input type=\"checkbox\" id=\"group_" + d.Id + "\" value=\"" + d.Id + "\" " + (($.inArray(d.Id, _s) > -1 ? "checked=\"checked\"" : "")) + " /><i></i></div></div><label for=\"group_" + d.Id + "\">" + d.Name + "</label></li>").appendTo(that.controls.groups);
            });
        },
        ///Build the associated files list.
        associated: function (options) {
            var _s = [], _st = [];
            $.each(options.data, function (i, d) {
                if ($.inArray(d.Catalog.Name, _st) == -1) {
                    _s.push(that.utility.fileName(d.Catalog.Name));
                    _st.push(d.Catalog.Name);
                }
            });
            _s.sort(function (_a, _b) {
                return (_a.toUpperCase() > _b.toUpperCase()) ? 1 : (_a.toUpperCase() < _b.toUpperCase()) ? -1 : 0;
            });
            that.controls.listAssociated.html("");
            $.each(_s, function (i, d) {
                $("<li>" + d + "</li>").appendTo(that.controls.listAssociated);
            });
        },
        ///Build the source point management UI.
        list: function (options) {
            try {
                var _dt = $.extend([], that.points), _d = [], _ss = [];
                if (that.sourcePointKeyword != "") {
                    var _sk = app.search.splitKeyword({ keyword: that.sourcePointKeyword });
                    if (_sk.length > 26) {
                        that.popup.message({ success: false, title: "Only support less than 26 keywords." });
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
                if (options.refresh) {
                    var _c = that.controls.headerListPoints.find(".point-header .ckb-wrapper input").prop("checked"), _s = that.utility.selected();
                    $.each(_s, function (m, n) {
                        _ss.push(n.SourcePointId);
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
                that.ui.item({ index: 0, data: _d, selected: _ss });
            } catch (err) {
                that.popup.message({
                    success: false, title: "Error occurred: " + err.message
                });
            }
        },
        ///Build each source point html.
        item: function (options, callback) {
            if (options.index < options.data.length) {
                var _item = options.data[options.index];
                that.range.exist({ RangeId: _item.NameRangeId }, function (ret) {
                    that.range.exist({ RangeId: _item.RangeId }, function (result) {
                        var _s = result.status == app.status.succeeded, _st = ret.status == app.status.succeeded, _ss = _s && _st;
                        options.data[options.index].Value = _s ? result.data.text : "";
                        options.data[options.index].Position = _s ? result.data.address : "";
                        options.data[options.index].Name = _st ? ret.data.text : "";
                        options.data[options.index].NamePosition = _st ? ret.data.address : "";

                        if (options.index >= that.pagerSize * (that.pagerIndex - 1) && options.index < that.pagerSize * that.pagerIndex) {
                            var _v = _s ? result.data.text : "",
                                _n = _st ? ret.data.text : "",
                                _p = _s ? that.utility.position(result.data.address) : {},
                                _pn = _st ? that.utility.position(ret.data.address) : {},
                                _pv = (_item.PublishedHistories && _item.PublishedHistories.length > 0 ? (_item.PublishedHistories[0].Value ? _item.PublishedHistories[0].Value : "") : ""),
                                _sel = $.inArray(_item.Id, options.selected) > -1,
                                _pht = _item.PublishedHistories && _item.PublishedHistories.length > 0 ? _item.PublishedHistories : [],
                                _pi = 0;
                            var _h = '<li class="point-item' + (_ss ? "" : " item-error") + '" data-id="' + _item.Id + '" data-range="' + _item.RangeId + '" data-position="' + (_s ? result.data.address : "") + '" data-namerange="' + (_st ? _item.NameRangeId : "") + '" data-nameposition="' + (_st ? ret.data.address : "") + '">';
                            _h += '<div class="point-item-line">';
                            _h += '<div class="i1"><div class="ckb-wrapper' + (_sel ? " checked" : "") + '"><input type="checkbox" ' + (_sel ? 'checked="checked"' : '') + ' /><i></i></div></div>';
                            _h += '<div class="i2"><span class="s-name" title="' + _n + '">' + _n + '</span>';
                            if (_st) {
                                _h += '<span title="' + (_pn.sheet ? _pn.sheet : "") + ':[' + (_pn.cell ? _pn.cell : "") + ']"><strong>' + (_pn.sheet ? _pn.sheet : "") + ':</strong>[' + (_pn.cell ? _pn.cell : "") + ']</span>';
                            }
                            _h += '</div>';
                            _h += '<div class="i3" title="' + _pv + '">' + _pv + '</div>';
                            _h += '<div class="i4"><span class="s-value" title="' + _v + '">' + _v + '</span>';
                            if (_s) {
                                _h += '<span title="' + (_p.sheet ? _p.sheet : "") + ':[' + (_p.cell ? _p.cell : "") + ']"><strong>' + (_p.sheet ? _p.sheet : "") + ':</strong>[' + (_p.cell ? _p.cell : "") + ']</span>';
                            }
                            _h += '</div>';
                            _h += '<div class="i5"><div class="i-line"><i class="i-history" title="History"></i><i class="i-delete" title="Delete"></i><i class="i-edit" title="Edit"></i></div>';
                            _h += '<div class="i-menu"><a href="javascript:"><span title="Action">...</span><span><i class="i-history" title="History"></i><i class="i-delete" title="Delete"></i><i class="i-edit" title="Edit"></i></span></a></div>';
                            _h += '</div>';
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
                            _h += '<p>The source point is invalid. This could be caused by not saving the excel file after creating the source point. Please delete the source point and recreate it.</p>';
                            _h += '</div>';
                            _h += '</li>';
                            that.controls.list.append(_h);
                        }
                        options.index++;
                        that.ui.item(options, callback);
                    });
                });
            }
            else {
                that.filteredPoints = options.data;
                that.controls.main.get(0).scrollTop = 0;
                if (callback) {
                    callback();
                }
            }
        },
        ///Remove a row of source point in UI.
        remove: function (options) {
            that.controls.list.find("[data-id=" + options.Id + "]").remove();
        },
        ///Build the publish history of the source point.
        publish: function (options, callback) {
            $.each(options.SourcePoints, function (i, d) {
                var _e = that.controls.list.find("[data-id=" + d.Id + "]"),
                    _pv = (d.PublishedHistories && d.PublishedHistories.length > 0 ? (d.PublishedHistories[0].Value ? d.PublishedHistories[0].Value : "") : ""),
                    _pht = d.PublishedHistories && d.PublishedHistories.length > 0 ? d.PublishedHistories : [],
                    _pi = 0;
                _e.find(".history-list").find(".history-item").remove();
                $.each(_pht, function (m, n) {
                    var __c = $.trim(_pht[m].Value ? _pht[m].Value : ""),
                        __p = $.trim(_pht[m > 0 ? m - 1 : m].Value ? _pht[m > 0 ? m - 1 : m].Value : "");
                    if (_pi < 5 && (m == 0 || __c != __p)) {
                        _e.find(".history-list").append('<li class="history-item"><div class="h1" title="' + n.PublishedUser + '">' + n.PublishedUser + '</div><div class="h2" title="' + (n.Value ? n.Value : "") + '">' + (n.Value ? n.Value : "") + '</div><div class="h3" title="' + that.utility.date(n.PublishedDate) + '">' + that.utility.date(n.PublishedDate) + '</div></li>');
                        _pi++;
                    }
                });
                _e.find(".i3").prop("title", _pv).html(_pv);
                that.controls.list.find(".ckb-wrapper input").prop("checked", false);
                that.controls.list.find(".ckb-wrapper").removeClass("checked");
            });

            callback();
        }
    };

    return that;
})();