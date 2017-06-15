/* Common app functionality */

var app = (function () {
    "use strict";

    var app = {
        status: {
            failed: "failed",
            succeeded: "succeeded"
        },
        intelliSenseResults: 6
    };
    // Generate the GUID.
    app.guid = function () {
        var guid = "";
        for (var i = 1; i <= 32; i++) {
            var n = Math.floor(Math.random() * 16.0).toString(16);
            guid += n;
            if ((i == 8) || (i == 12) || (i == 16) || (i == 20))
                guid += "-";
        }
        return guid;
    };
    ///Search source point or destination point.
    app.search = {
        /// Split the keyword by ' ' and push the key words into an array.
        splitKeyword: function (options) {
            var _k = options.keyword.toLocaleLowerCase(), _ka = [];
            $.each(_k.split(" "), function (x, y) {
                if (y != "") {
                    _ka.push(y);
                }
            });
            return _ka;
        },

        /// Determine if the keyword fully matched or not.
        /// If it is matched, then return 1, otherwise return 0.
        /// For example: source point name: abc, the keywords are a b d.This will return 0. If the keywords are a b, then this will return 1.
        weight: function (options) {
            var _k = options.keyword, _s = $.trim(options.source).toLocaleLowerCase(), _w = 0, _f = 0;
            for (var i = 0; i < _k.length; i++) {
                if (_s.indexOf(_k[i]) > -1) {
                    _f++;
                }
            }
            return _f >= _k.length ? 1 : 0;
        },
        ///Get the keywords and splity the keywords then find the matched source point by splitted key words fianlly display the search results.
        autoComplete: function (options) {
            var _k = app.search.splitKeyword({ keyword: options.keyword }), _d = options.data, _r = [];
            $.each(_d, function (i, d) {
                var _w = app.search.weight({ keyword: _k, source: d.Name });
                if (_w > 0) {
                    d.weight = _w;
                    _r.push(d);
                }
            });
            if (_r.length > 0) {
                options.result.find("li").remove();
                $.each(_r, function (i, d) {
                    if (i < app.intelliSenseResults) {
                        $('<li>' + d.Name + '</li>').appendTo(options.result);
                    }
                });
                options.result.show();
            }
            else {
                options.result.hide();
            }
            options.target.data("keyword", options.keyword);
        },
        ///Support the cursor moving up & down in the search result area. 
        move: function (options) {
            if (!options.result.is(":hidden")) {
                var _i = options.result.find("li.active").index(), _l = options.result.find("li").length, _m = false;
                if (options.down) {
                    if (_i < _l - 1) {
                        _i++;
                        _m = true;
                    }
                }
                else {
                    if (_i == -1) {
                        _i = _l;
                    }
                    if (_i > 0) {
                        _i--;
                        _m = true;
                    }
                }
                if (_m) {
                    options.result.find("li.active").removeClass("active");
                    options.result.find("li").eq(_i).addClass("active");
                    options.target.val(options.result.find("li").eq(_i).text());
                }
                else {
                    options.result.find("li.active").removeClass("active");
                    options.target.val(options.target.data("keyword"));
                }
            }
        }
    };

    return app;
})();