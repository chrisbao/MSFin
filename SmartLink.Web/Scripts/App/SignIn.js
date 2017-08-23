$(function () {
    Office.initialize = function (reason) {
        $(document).ready(function () {
            var isDev = window.location.href.indexOf("localhost") > -1;
            var mode = { word: false, excel: false, powerPoint: false };
            ///If it is word add-in.
            if (Office.context.requirements.isSetSupported("WordApi")) {
                mode.word = true;
            }

                ///If it is excel add-in.
            else if (Office.context.requirements.isSetSupported("ExcelApi")) {
                mode.excel = true;
            }

                ///If it is powerpoint add-in.
            else if (Office.context.requirements.isSetSupported("ActiveView")) {
                mode.powerPoint = true;
            }

            if (mode.word || mode.excel || mode.powerPoint) {
                if (isDev || (Office.context.document.url && (Office.context.document.url.toUpperCase().indexOf("HTTP") > -1 || Office.context.document.url.toUpperCase().indexOf("HTTPS") > -1))) {
                    //If excel requirement set is 1.3 or word requirement set is 1.2 or it is powerpoint add-in.
                    if (Office.context.requirements.isSetSupported('ExcelApi', 1.3) || Office.context.requirements.isSetSupported("WordApi", 1.2) || mode.powerPoint) {
                        $(".sign-in").show();
                        $("#btnSignIn").click(function () {
                            if (mode.word) {
                                //go to word home page.
                                window.location = "/Word/Point";
                            }
                            else if (mode.excel) {
                                //go to excel home page.
                                window.location = "/Excel/Point";
                            }
                            else if (mode.powerPoint) {
                                //go to powerpoint home page.
                                window.location = "/PowerPoint/Point";
                            }
                        });
                    }
                    else {
                        $("#error-message").addClass(mode.word ? "word-version" : "excel-version");
                    }
                }
                else {
                    $("#error-message").addClass(mode.word ? "word-mode" : (mode.excel ? "excel-mode" : "powerpoint-mode"));
                }
            }
        });
    };
});