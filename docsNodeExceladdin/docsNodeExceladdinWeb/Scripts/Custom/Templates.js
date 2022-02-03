var SPURL = "";
var GraphAPIToken = '';
var SPToken = '';
var templateServerRelURL = "/sites/DocsNodeAdmin/";
var DocsNodePictureDisplayName = '';
var DocsNodeSlideDisplayName = '';
var DocsNodeConfigurationDisplayName = 'DocsNodeConfiguration';
var imageListSite = '';
var DocsNodeCategoryDisplayName = '';
var twoBreadCrumbArr = [];
var oneBreadCrumbArr = [];
var imageMaxID = 0;
var imageParentMaxID = 0;
var slideMaxID = 0;
var slideParentMaxID = 0;
var imageItem = 20;
var slideItem = 14;
var BRTemplatesJS = window.BRTemplatesJS || {};
var imageSearchQuery = "";
var slideSearchQuery = "";
var DocNodePowerPoint;
var TokenArray;

var snippetSearch = "";
var imgSearch = "";

(function () {
    "use strict";
    Office.initialize = function (reason) {
        $(document).ready(function () {
           
            $("body").on("click", ".clickme a", function () {
                $('#emptydataTwo').hide();
                $('#emptydataOne').hide();
                $('#ManageTabsContent').show();
                $("#WaitDialog").hide();
                $("#WaitDialog").hide();
                $('.clickme a').removeClass('activelink');
                $(this).addClass('activelink');
                var tabid = $(this).data('tag');
                $('.list').removeClass('active').addClass('hide');
                $('#' + tabid).addClass('active').removeClass('hide');
                if (tabid == 'two') {
                    $('#btnRefresh').show();
                    $('#secondMainContent').css('display', 'block');
                    $("#txtSearch").attr("placeholder", "Search Corporate Images");
                    $("#txtSearch").val(imgSearch);
                    $('#mainContent').css('display', 'none');
                    if (($("#two ul.categoryList").html() != "" || $("#two ul.categoryItems").html() != "") && ($("#hdnCategory").val() == "" || $("#hdnCategory").val() == "0" || $("#hdnCategory").val() == undefined)) {
                        $("#two ul.categoryList").show();
                        $("#two ul.categoryItems").show();
                        $("#two ul.childCategoryList").hide();
                        $("#two ul.childCategoryItems").hide();
                    } else {

                        $('#txtSearch').val("");
                        slideSearchQuery = "";
                        imageSearchQuery = "";

                        imageMaxID = 0;
                        var $_category = $("#hdnCategory").val();
                        var $_categoryName = $("#hdnCategoryName").val();
                        if ($_category != '' && $_categoryName != '') {
                            LoadCategoryandImagesOnFolderClick($_category, $_categoryName, false);
                        } else {
                            LoadImageCategoryandItemsOnPageLoad();
                        }
                    }
                }
                if (tabid == 'one') {
                    $('#btnRefresh').show();
                    $('#secondMainContent').css('display', 'block');
                    $("#txtSearch").attr("placeholder", "Search Slide");
                    $("#txtSearch").val(snippetSearch);
                    $('#mainContent').css('display', 'none');
                    if (($("#one ul.categoryList").html() != "" || $("#one ul.categoryItems").html() != "") && ($("#hdnTextCategory").val() == "" || $("#hdnTextCategory").val() == "0" || $("#hdnTextCategory").val() == undefined)) {
                        $("#one ul.categoryItems").show();
                        $("#one ul.categoryList").show();
                        $("#one ul.childCategoryList").hide();
                        $("#one ul.childCategoryItems").hide();
                    }
                    else {
                        $('#txtSearch').val("");
                        slideSearchQuery = "";
                        imageSearchQuery = "";
                        slideMaxID = 0;
                        var $_category = $("#hdnTextCategory").val();
                        var $_categoryName = $("#hdnTextCategoryName").val();
                        if ($_category != '' && $_categoryName != '') {
                            LoadCategoryandSlideOnFolderClick($_category, $_categoryName, false);
                        } else {
                            LoadCategoryandSlideOnFolderClick(null, '', false);
                        }
                    }
                }
                if (tabid == 'zero') {
                    $('#secondMainContent').css('display', 'none');
                    $('#mainContent').css('display', 'block');
                    $('#btnRefresh').hide();
                    $('#ManageTabsContent').hide();
                }
            });

            $("body").on("click", "#one .liTitle", function () {
                $("#txtSearch").val(snippetSearch);
                $('#emptydataOne').hide();
                var $_this = $(this).attr("attr-category");
                var catName = $(this).attr('attr-title');
                LoadCategoryandSlideOnFolderClick($_this, catName, true);
            });

            $("body").on("click", "#two .liTitle", function () {
                $('#emptydataTwo').hide();
                var $_this = $(this).attr("attr-category");
                var catName = $(this).attr('attr-title');
                LoadCategoryandImagesOnFolderClick($_this, catName, true);
            });

            var activeLi = localStorage.getItem("active");

            //on keydown, clear the countdown 
            $('#txtSearch').on('keypress', function (e) {
                if (e.which == 13) {
                    $('.list.SideContentBox.active').scrollTop(0);
                    setTimeout(function () {
                        if ($(".SideContentBox.active").attr("id") == "one") {
                            snippetSearch = $("#txtSearch").val();
                            var $_category = $("#hdnTextCategory").val();
                            if ($("#txtSearch").val().trim() != "") {
                                searchSlides($("#txtSearch").val().trim(), $_category);
                            }
                            else {
                                clearSearchText();
                            }
                        }
                        if ($(".SideContentBox.active").attr("id") == "two") {
                            var $_category = $("#hdnCategory").val();
                            imgSearch = $("#txtSearch").val();
                            if ($("#txtSearch").val().trim() != "") {
                                seachImages($("#txtSearch").val().trim(), $_category);
                            }
                            else {
                                clearSearchText();
                            }
                        }
                    }, 100);
                }
            });           

            $("body").on("click", "#clearSearch", function () {
                $('.list.SideContentBox.active').scrollTop(0);
                setTimeout(function () {
                    refresh();
                }, 100);
              
            });
            $('.list.SideContentBox').off('scroll');
            $('.list.SideContentBox').on('scroll', function () {
                if ($(this).scrollTop() + $(this).innerHeight() >= $(this)[0].scrollHeight) {
                    if ($('.list.SideContentBox.active').attr("id") == "two") // Image section
                    {
                        loadImagesOnScroll();
                    }
                    if ($('.list.SideContentBox.active').attr("id") == "one") // Image section
                    {
                        loadTextOnScroll();
                    }
                }
            });
            $('#emptydataTwo').hide();
            $("body").on("click", "#btnRefresh", function () {
                $('.list.SideContentBox.active').scrollTop(0);
                setTimeout(function () {
                    refresh();
                }, 100);
                
            });
        });
    };
})();
BRTemplatesJS.Config = function () {  
    var utility = new BRDocsNodeJS.postTokens();
    TokenArray = utility.callFunction();
    SPToken = TokenArray[0];
    GraphAPIToken = TokenArray[1];
    SPURL = TokenArray[2];

    setInterval(function () {
        TokenArray = utility.callFunction();
        SPToken = TokenArray[0];
        GraphAPIToken = TokenArray[1];
        SPURL = TokenArray[2];
        console.log("Refresh Tokens template");
    }, 540001);
    GetConfigurations();
};
    function GetConfigurations() {       

        DocsNodeCategoryDisplayName = localStorage.getItem('CategoryListGUID');  
        DocsNodePictureDisplayName = localStorage.getItem('ImageLibraryListGUID'); 
        DocsNodeSlideDisplayName = localStorage.getItem('SlidesLibraryListGUID');
        imageListSite = localStorage.getItem('ImageSourceListPath');
       LoadImageCategoryandItemsOnPageLoad();     
    }

    function refresh() {
        if ($(".list.SideContentBox.active").attr("id") == "one" || $(".list.SideContentBox.active").attr("id") == "two") {
            $('#txtSearch').val("");
            $('#emptydataTwo').hide();
            $('#emptydataOne').hide();
            slideSearchQuery = "";
            imageSearchQuery = "";
            clearSearchText();
        }
    }

    function LoadCategoryandSlidesOnPageLoad(filterCondition, catContainer, itemContainer, categoryId) {
        var $_ulList = catContainer;
        var $_ulItems = itemContainer;
        $('#emptydataOne').hide();
        new Promise(function (callback, reject) {
            var $_ul = catContainer;
            $('#WaitDialog').show();
            loadDataFromSharePoint(DocsNodeCategoryDisplayName, 'ID,Title,Category,ParentCategory,CategoryType,CategoryLevel', filterCondition).then(function (data) {
                var results = data.results;
                if (results.length > 0) {
                    $.each(data.results, function (index, value) {
                        var category = value.Category;
                        var str = "<li>" +
                            "<a href='#' class='liTitle' attr-title='" + category + "' attr-category='" + value.ID + "'>" +
                            "<span class='file-icon'><i class='ms-Icon ms-Icon--OpenFolderHorizontal' aria-hidden='true'></i></span>" +
                            "<span class='fileName' title='" + category + "'>" + (category.length > 20 ? category.substr(0, 20) + "..." : category) + "</span>" +
                            "<span class='fileEnter-icon'><i class='ms-Icon ms-Icon--ChevronRightSmall' aria-hidden='true'></i></span>" +
                            "</a >" +
                            "</li>"
                        $_ul.append(str);
                        $("#WaitDialog").hide();
                    });
                    $('#emptydataOne').hide();
                }
                callback(1);
            }, function (error) {
                var erroeMeg = JSON.parse(error.responseText);
                erroeMeg = erroeMeg["error"].message.value;
                docsTemplateList = "<div class='displayMessage'>" + erroeMeg + "\nNo Template Library found in DocsNode Admin Panel.</div>";
                $_ul.html(docsTemplateList);
                $("#WaitDialog").hide();
            });
        }).then(function (result) {
            var $_ul = itemContainer;
            $_ul.html("");
            loadDataFromSharePoint(DocsNodeSlideDisplayName, 'ID,Title,LinkFilename,EncodedAbsUrl,SlidesDiscription,SlidesCategory/ID,SlidesCategory/Title&$expand=SlidesCategory', "(SlidesCategory eq " + categoryId + ")", "slide").then(function (data) {
                var results = data.results;
                var catLi = '';
                if (results.length > 0) {
                    if (categoryId == null)
                        slideParentMaxID = _.max(_.pluck(results, "ID"));
                    else
                        slideMaxID = _.max(_.pluck(results, "ID"));
                    bindSlidesHTML(results, itemContainer);
                } else {
                    if ($_ulList[0].childElementCount == 0 && $_ulItems[0].childElementCount == 0) {
                        $('#emptydataOne').show();
                        $("#WaitDialog").hide();
                    } else {
                        $('#emptydataOne').hide();
                    }
                }
            });
        });
    };

    function clearSearchText() {
        if ($(".SideContentBox.active").attr("id") == "one") {
            snippetSearch = "";

            var $_category = $("#hdnTextCategory").val();
            var $_categoryName = $("#hdnTextCategoryName").val();
            $("#txtSearch").attr("placeholder", "Search Slides");
            $("#txtSearch").val('');
            if ($_category == "" || $_category == undefined || $_category == null) {
                slideParentMaxID = 0;
                $("#one ul.categoryItems").html("");
                $("#one ul.categoryList").html("");
                LoadCategoryandSlideOnFolderClick(null, "", false);
            }
            else {
                slideMaxID = 0;
                $("#one ul.childCategoryList").html("");
                $("#one ul.childCategoryItems").html("");
                LoadCategoryandSlideOnFolderClick($_category, $_categoryName, false);
            }
        }

        if ($(".SideContentBox.active").attr("id") == "two") {
            var $_category = $("#hdnCategory").val();
            var $_categoryName = $("#hdnCategoryName").val();
            imgSearch = '';
            $("#txtSearch").attr("placeholder", "Search Corporate Images");
            $("#txtSearch").val('');
            if ($_category == "" || $_category == undefined || $_category == null) {
                imageParentMaxID = 0;
                $("#two ul.categoryList").html("");
                $("#two ul.categoryItems").html("");
                LoadImageCategoryandItemsOnPageLoad(0);
            }
            else {
                imageMaxID = 0;
                $("#two ul.childCategoryList").html("");
                $("#two ul.childCategoryItems").html("");
                LoadCategoryandImagesOnFolderClick($_category, $_categoryName, false);
            }
        }
    }

    function searchSlides(searchParam, categoryId) {
        $("#WaitDialog").show();

        if (categoryId != "" && categoryId != "0") {

            slideMaxID = 0;
            $("#one ul.childCategoryList").html("");
            $("#one ul.childCategoryItems").html("");

            $("#one ul.categoryList").hide();
            $("#one ul.categoryItems").hide();
            $("#one ul.childCategoryList").hide();
            $("#one ul.childCategoryItems").show();

            var categoryList = [];
            categoryList.push(categoryId);

            loadDataFromSharePoint(DocsNodeCategoryDisplayName, 'ID,Title,Category,ParentCategory,CategoryType,CategoryLevel', "((CategoryType eq 'Slides') and (ParentCategory eq " + categoryId + "))").then(function (data) {
                $.each(data.results, function (index, value) {
                    categoryList.push(value.ID);
                });

                var searchString = "substringof('" + searchParam + "',Title) and ";
                var cateStr = "";
                if (categoryList.length > 1) {
                    cateStr += "("
                    for (var i = 0; i < categoryList.length; i++) {
                        if (i + 1 == categoryList.length)
                            cateStr += "(SlidesCategory/ID eq " + categoryList[i] + ")";
                        else
                            cateStr += "(SlidesCategory/ID eq " + categoryList[i] + ") or ";
                    }
                    cateStr += ")"

                    slideSearchQuery = "(" + searchString + cateStr + ")";
                }
                else {
                    slideSearchQuery = "(" + searchString + "(SlidesCategory/ID eq " + categoryId + "))";
                }
                slideSearchAjax($("#one ul.childCategoryItems"));
            });

        }
        else {
            slideSearchQuery = "(substringof('" + searchParam + "',Title))";
            slideParentMaxID = 0;
            $("#one ul.categoryList").html("");
            $("#one ul.categoryItems").html("");
            $("#one ul.categoryList").hide();
            $("#one ul.categoryItems").show();
            $("#one ul.childCategoryList").hide();
            $("#one ul.childCategoryItems").hide();
            slideSearchAjax($("#one ul.categoryItems"));
        }
    }

    function slideSearchAjax(container) {
        loadDataFromSharePoint(DocsNodeSlideDisplayName, 'ID,Title,LinkFilename,EncodedAbsUrl,SlidesDiscription,SlidesCategory/ID,SlidesCategory/Title&$expand=SlidesCategory', slideSearchQuery, 'slide').then(function (data) {
            var results = data.results;
            if (results.length > 0) {
                if ($("#hdnTextCategory").val() == "" || $("#hdnTextCategory").val() == "0" || $("#hdnTextCategory").val() == undefined) {
                    slideParentMaxID = _.max(_.pluck(data.results, "ID"));
                }
                else
                    slideMaxID = _.max(_.pluck(data.results, "ID"));
                bindSlidesHTML(results, container);
                $("#WaitDialog").hide();
            } else {
                if (container[0].childElementCount == 0) {
                    $('#emptydataOne').show();
                    $("#WaitDialog").hide();
                } else {
                    $('#emptydataOne').hide();
                }
            }
        });
    }

    function seachImages(searchParam, categoryId) {

        $("#WaitDialog").show();
        if (categoryId != "" && categoryId != "0") {
            imageMaxID = 0;
            var itemSeachQuery;

            $(".childCategoryList").html("");
            $(".childCategoryItems").html("");

            $(".childCategoryList").show();
            $(".childCategoryItems").show();

            $(".categoryList").hide();
            $("#two ul.categoryItems").hide();
            var categoryList = [];
            categoryList.push(categoryId);
            if (imageListSite.toLowerCase() == '/sites/docsnodeadmin/' || imageListSite.toLowerCase() == '/sites/docsnodeadmin') {
                loadDataFromSharePoint(DocsNodeCategoryDisplayName, 'ID,Title,Category,ParentCategory,CategoryType,CategoryLevel', "((CategoryType eq 'Images') and (ParentCategory eq " + categoryId + "))", "image", true).then(function (data) {
                    $.each(data.results, function (index, value) {
                        categoryList.push(value.ID);
                    });
                    var searchString = "substringof('" + searchParam + "',Title) and ";
                    var cateStr = "";
                    if (categoryList.length > 1) {
                        cateStr += "("
                        for (var i = 0; i < categoryList.length; i++) {
                            if (i + 1 == categoryList.length)
                                cateStr += "(ImageCategory/ID eq " + categoryList[i] + ")";
                            else
                                cateStr += "(ImageCategory/ID eq " + categoryList[i] + ") or ";
                        }
                        cateStr += ")"

                        itemSeachQuery = "(" + searchString + cateStr + ")";
                    }
                    else {
                        itemSeachQuery = "(" + searchString + "(ImageCategory/ID eq " + categoryId + "))";
                    }
                    //This is needed for lazy loading
                    imageSearchQuery = itemSeachQuery;
                    imageSearchAjax($("#two ul.childCategoryItems"));
                });
            }
            
        }
        else {
            imageSearchQuery = "(substringof('" + searchParam + "',Title))";
            //This is needed for lazy loading
            imageParentMaxID = 0;
            $("#two ul.categoryList").html("");
            $("#two ul.categoryItems").html("");
            $("#two ul.childCategoryList").hide();
            $("#two ul.childCategoryItems").hide();
            $("#two ul.categoryList").hide();
            $("#two ul.categoryItems").show();
            imageSearchAjax($("#two ul.categoryItems"));
        }
    }

    function imageSearchAjax(container) {
        loadDataFromSharePoint(DocsNodePictureDisplayName, 'ID,Title,LinkFilenameNoMenu,EncodedAbsThumbnailUrl,EncodedAbsUrl,ImageCategory/ID,ImageCategory/Title&$expand=ImageCategory', imageSearchQuery, "image",true).then(function (data) {
            var results = data.results;
            var $_ul = container;
            if (results.length > 0) {
                if ($("#hdnCategory").val() == "" || $("#hdnCategory").val() == null || $("#hdnCategory").val() == undefined) {
                    imageParentMaxID = _.max(_.pluck(results, "ID"));
                }
                else
                    imageMaxID = _.max(_.pluck(results, "ID"));

                $.each(results, function (index, image) {
                    toDataURL(image.EncodedAbsThumbnailUrl.toString(), function (dataUrl) {
                        var base64result = "data:image/png;base64, " + dataUrl;
                        $_ul.append("<li class='liIteamImg' onClick=\"insertImage('" + image.EncodedAbsUrl.toString() + "')\"><a href='#'><div class='liInnerImage'><img title='" + (image.Title != null ? image.Title : image.LinkFilenameNoMenu) + "' src='" + base64result + "'>" +
                            "</div></a></li>");
                        $("#WaitDialog").hide();
                    });
                });
                $('#emptydataTwo').hide();
            } else {
                if ($_ul[0].childElementCount == 0) {
                    $('#emptydataTwo').show();
                    $("#WaitDialog").hide();
                } else {
                    $('#emptydataTwo').hide();
                }
            }
        });
    }

    function LoadImageCategoryandItemsOnPageLoad() {
        twoBreadCrumbArr = [];
        var $_breaCrumbUL = $("#two ul.breadcrumb");
        $_breaCrumbUL.html("");
        $("#two ul.breadcrumb").hide();
        $("#two ul.categoryList").html("");
        $("#two ul.categoryItems").html("");
        $("#two ul.categoryList").show();
        $("#two ul.categoryItems").show();
        var $_ulList = $('#two ul.categoryList');
        var $_ulItems = $('#two ul.categoryItems');

        $("#two ul.childCategoryList").hide();
        $("#two ul.childCategoryItems").hide();

        if (imageListSite.toLowerCase() == "/sites/docsnodeadmin" || imageListSite.toLowerCase() == "/sites/docsnodeadmin/") {
            new Promise(function (callback, reject) {
                var $_ul = $("#two ul.categoryList");
                $("#WaitDialog").show();
                loadDataFromSharePoint(DocsNodeCategoryDisplayName, 'ID,Title,Category,ParentCategory,CategoryType,CategoryLevel', "((CategoryType eq 'Images') and (CategoryLevel eq 0))").then(function (response) {
                    var results = response.results
                    if (results.length > 0) {
                        $.each(results, function (index, value) {
                            var category = value.Category;
                            var str = "<li>" +
                                "<a href='#' class='liTitle' attr-title='" + category + "' attr-category='" + value.ID + "'>" +
                                "<span class='file-icon'><i class='ms-Icon ms-Icon--OpenFolderHorizontal' aria-hidden='true'></i></span>" +
                                "<span class='fileName ' title='" + category + "'>" + (category.length > 20 ? category.substr(0, 20) + "..." : category) + "</span>" +
                                "<span class='fileEnter-icon'><i class='ms-Icon ms-Icon--ChevronRightSmall' aria-hidden='true'></i></span>" +
                                "</a >" +
                                "</li>"
                            $_ul.addClass("root");
                            $_ul.append(str);
                        });
                    }
                    $('#emptydataTwo').hide();
                    $("#WaitDialog").hide();


                    callback(1);
                }, function (error) {
                    var erroeMeg = JSON.parse(error.responseText);
                    erroeMeg = erroeMeg["error"].message.value;
                    docsTemplateList = "<div class='displayMessage'>" + erroeMeg + "\nNo Template Library found in DocsNode Admin Panel.</div>";
                    $_ul.html(docsTemplateList);
                    $("#WaitDialog").hide();
                });
            }).then(function (result) {
                var $_ul = $("#two ul.categoryItems");
                $_ul.html("");

                loadDataFromSharePoint(DocsNodePictureDisplayName, 'ID,Title,LinkFilenameNoMenu,EncodedAbsThumbnailUrl,EncodedAbsUrl,ImageCategory/ID,ImageCategory/Title,*&$expand=ImageCategory', "(ImageCategory eq null)", "image").then(function (response) {
                    var results = response.results;
                    if (results.length > 0) {
                        imageParentMaxID = _.max(_.pluck(results, "ID"));
                        $.each(results, function (index, image) {
                            toDataURL(image.EncodedAbsThumbnailUrl.toString(), function (dataUrl) {
                                var base64result = "data:image/png;base64, " + dataUrl;
                                $_ul.append("<li class='liIteamImg' onClick=\"insertImage('" + image.EncodedAbsUrl.toString() + "')\"><a href='#'><div class='liInnerImage'><img title='" + (image.Title != null ? image.Title : image.LinkFilenameNoMenu) + "' src='" + base64result + "'>" +
                                    "</div></a></li>");
                                $("#WaitDialog").hide();
                            });
                        });
                        $('#emptydataTwo').hide();
                    } else {
                        if ($_ulList[0].childElementCount == 0 && $_ulItems[0].childElementCount == 0) {
                            $('#emptydataTwo').show();
                            $("#WaitDialog").hide();
                        } else {
                            $('#emptydataTwo').hide();
                        }
                    }
                });
            });
        }
        else {
            var $_ul = $("#two ul.categoryItems");
            $_ul.html("");
            loadDataFromSharePoint(DocsNodePictureDisplayName, 'ID,Title,LinkFilenameNoMenu,EncodedAbsThumbnailUrl,EncodedAbsUrl', "", "image").then(function (response) {
                var results = response.results;
                if (results.length > 0) {
                    imageParentMaxID = _.max(_.pluck(results, "ID"));
                    $.each(results, function (index, image) {
                        toDataURL(image.EncodedAbsThumbnailUrl.toString(), function (dataUrl) {
                            var base64result = "data:image/png;base64, " + dataUrl;
                            $_ul.append("<li class='liIteamImg' onClick=\"insertImage('" + image.EncodedAbsUrl.toString() + "')\"><a href='#'><div class='liInnerImage'><img title='" + (image.Title != null ? image.Title : image.LinkFilenameNoMenu) + "' src='" + base64result + "'>" +
                                "</div></a></li>");
                            $("#WaitDialog").hide();
                        });
                    });
                    $('#emptydataTwo').hide();
                } else {
                    if ($_ulList[0].childElementCount == 0 && $_ulItems[0].childElementCount == 0) {
                        $('#emptydataTwo').show();
                        $("#WaitDialog").hide();
                    } else {
                        $('#emptydataTwo').hide();
                    }
                }
            });
        }
    }
    function loadTextOnScroll() {

        var $_category = $("#hdnTextCategory").val();
        var $_ul = "";
        var filter = "";
        var cat = 0;

        if (slideSearchQuery != "") {
            filter = slideSearchQuery;
        }

        if ($_category == "" || $_category == undefined || $_category == null) {
            $_ul = $("#one ul.categoryItems");
        }
        else {
            cat = $_category;
            $_ul = $("#one ul.childCategoryItems");
        }

        getSlidestForLazyLoading(cat, filter).then(function (data) {
            if (cat == 0)
                slideParentMaxID = _.max(_.pluck(data.results, "ID"));
            else
                slideMaxID = _.max(_.pluck(data.results, "ID"));
            bindSlidesHTML(data.results, $_ul);
        });
    }
    function getSlidestForLazyLoading(categoryId, filter) {
        var dfd = $.Deferred();


        var cat = "";

        if (categoryId == 0)
            cat = "null";
        else
            cat = categoryId;

        var url = "";

        if (categoryId == 0)// It means parent level
        {
            if (filter == "")
                url = SPURL + templateServerRelURL + "_api/web/lists(guid'" + DocsNodeSlideDisplayName + "')/items?$select=ID,Title,LinkFilename,EncodedAbsUrl,SlidesDiscription,SlidesCategory/ID,SlidesCategory/Title&$expand=SlidesCategory&$filter=(SlidesCategory/ID eq " + cat + " and ID gt " + slideParentMaxID + ")&$top=" + slideItem;
            else
                url = SPURL + templateServerRelURL + "_api/web/lists(guid'" + DocsNodeSlideDisplayName + "')/items?$select=ID,Title,LinkFilename,EncodedAbsUrl,SlidesDiscription,SlidesCategory/ID,SlidesCategory/Title&$expand=SlidesCategory&$filter=(" + filter + " and ID gt " + slideParentMaxID + ")&$top=" + slideItem;
        }
        else {
            if (filter == "") {
                url = SPURL + templateServerRelURL + "_api/web/lists(guid'" + DocsNodeSlideDisplayName + "')/items?$select=ID,Title,LinkFilename,EncodedAbsUrl,SlidesDiscription,SlidesCategory/ID,SlidesCategory/Title&$expand=SlidesCategory&$filter=(SlidesCategory/ID eq " + cat + " and ID gt " + slideMaxID + ")&$top=" + slideItem;
            }
            else {
                url = SPURL + templateServerRelURL + "_api/web/lists(guid'" + DocsNodeSlideDisplayName + "')/items?$select=ID,Title,LinkFilename,EncodedAbsUrl,SlidesDiscription,SlidesCategory/ID,SlidesCategory/Title&$expand=SlidesCategory&$filter=(" + filter + " and ID gt " + slideMaxID + ")&$top=" + slideItem;
            }
        }


        try {
            callAjaxGet(url).done(function (data) {
                dfd.resolve(data.d);
            });
        } catch (err) {
            console.log(err);
            dfd.reject(err);
        }

        return dfd;
    }
    function loadImagesOnScroll() {
        var $_category = $("#hdnCategory").val();
        var $_ul = "";
        var filter = "";
        if (imageSearchQuery != "") {
            filter = imageSearchQuery;
            if ($_category == "" || $_category == undefined || $_category == null) {
                $_ul = $("#two ul.categoryItems");
            }
            else {
                $_ul = $("#two ul.childCategoryItems");
            }
        }
        else {
            if ($_category == "" || $_category == undefined || $_category == null) {
                filter = "(ImageCategory eq null)";
                $_ul = $("#two ul.categoryItems");
            }
            else {
                filter = "(ImageCategory/ID eq " + $_category + ")";
                $_ul = $("#two ul.childCategoryItems");
            }
        }

        loadDataFromSharePoint(DocsNodePictureDisplayName, 'ID,Title,LinkFilenameNoMenu,EncodedAbsThumbnailUrl,EncodedAbsUrl,ImageCategory/ID,ImageCategory/Title&$expand=ImageCategory', filter, "image").then(function (data) {
            var results = data.results;
            if (results.length > 0) {
                if ($_category == "" || $_category == undefined || $_category == null) {
                    imageParentMaxID = _.max(_.pluck(data.results, "ID"));
                }
                else {
                    imageMaxID = _.max(_.pluck(data.results, "ID"));
                }
                $.each(results, function (index, image) {
                    toDataURL(image.EncodedAbsThumbnailUrl.toString(), function (dataUrl) {
                        var base64result = "data:image/png;base64, " + dataUrl;
                        $_ul.append("<li class='liIteamImg' onClick=\"insertImage('" + image.EncodedAbsUrl.toString() + "')\"><a href='#'><div class='liInnerImage'><img title='" + (image.Title != null ? image.Title : image.LinkFilenameNoMenu) + "' src='" + base64result + "'>" +
                            "</div></a></li>");
                    });
                });
            }
        });
    }

    function LoadCategoryandImagesOnFolderClick($_this, $_categoryName, $_handleBreadcrumb) {
        $("#two ul.childCategoryList").html("");
        $("#two ul.childCategoryItems").html("");
        var $_category = $_this;
        //Added for Breadcrumb by Darshana
        if ($_handleBreadcrumb) {
            //console.log("Array:", twoBreadCrumbArr);
            if ($_category != 0) {
                twoBreadCrumbArr.push({ id: $_category, category: $_categoryName });
                if (twoBreadCrumbArr.length > 0) {
                    $("#two ul.breadcrumb").show();
                }
            }
            var $_breaCrumbUL = $("#two ul.breadcrumb");
            $_breaCrumbUL.html("");
            // Append to breadcrumb
            for (var key in twoBreadCrumbArr) {
                if (twoBreadCrumbArr.hasOwnProperty(key)) {
                    var catID = key == 0 ? 0 : twoBreadCrumbArr[key - 1].id;
                    $_breaCrumbUL.append('<li title="' + twoBreadCrumbArr[key].category + '"><a class="ms-Breadcrumb-itemLink" data-id="' + key + '" attr-category="' + catID + '"><i class="ms-Breadcrumb-chevron ms-Icon ms-Icon--ChevronLeft"></i><label class="ms-Breadcrumb-itemLink">' + twoBreadCrumbArr[key].category + '</label></li>');
                    $(document).off('click');
                    $(document).on("click", "#two ul.breadcrumb li", function (event) {
                        $('#emptydataTwo').hide();
                        $("#txtSearch").val("");
                        imageSearchQuery = "";
                        $("#two ul.childCategoryList").html("");
                        $("#two ul.childCategoryItems").html("");
                        var index = parseInt($(this).find("a").attr("data-id"));
                        if (index < twoBreadCrumbArr.length) {
                            twoBreadCrumbArr = twoBreadCrumbArr.splice(0, twoBreadCrumbArr.length - (twoBreadCrumbArr.length - index));
                            if (twoBreadCrumbArr.length == 0) {
                                $("#two ul.breadcrumb").hide();
                            }
                            var sli = $("#two ul.breadcrumb li a").filter(function () {
                                return $(this).attr("data-id") >= index;
                            });
                            if ($(sli).length > 0) $(sli).parent().remove();
                            if (parseInt($(this).find("a").attr("data-id")) - 1 == -1) {
                                $("#hdnCategory").val("");
                                $("#hdnCategoryName").val("");

                                $("#two ul.categoryList").show();
                                $("#two ul.categoryItems").show();
                                $("#two ul.childCategoryList").hide();
                                $("#two ul.childCategoryItems").hide();

                            } else {
                                imageMaxID = 0;
                                $("#hdnCategory").val($(this).find("a").attr("attr-category"));
                                $("#hdnCategoryName").val(twoBreadCrumbArr[index - 1].category);
                                loadImagesByCategory($(this).find("a").attr("attr-category"));
                            }
                        }


                    });
                }
            }
        }
        loadImagesByCategory($_category);
        $("#WaitDialog").show();
        $("#hdnCategory").val($_category);
        $("#hdnCategoryName").val($_categoryName);
    }

    function LoadCategoryandSlideOnFolderClick($_this, $_categoryName, $_handleBreadcrumb) {
        $("#one ul.childCategoryList").html("");
        $("#one ul.childCategoryItems").html("");
        var $_category = $_this;
        if ($_handleBreadcrumb) {
            if ($_category != 0) {
                oneBreadCrumbArr.push({ id: $_category, category: $_categoryName });
                if (oneBreadCrumbArr.length > 0) {
                    $("#one ul.breadcrumb").show();
                }
            }
            var $_breaCrumbUL = $("#one ul.breadcrumb");
            $_breaCrumbUL.html("");
            for (var key in oneBreadCrumbArr) {
                if (oneBreadCrumbArr.hasOwnProperty(key)) {
                    var catID = key == 0 ? 0 : oneBreadCrumbArr[key - 1].id;
                    var cateTitle = oneBreadCrumbArr[key]["category"];
                    $_breaCrumbUL.append('<li title="' + cateTitle + '"><a class="ms-Breadcrumb-itemLink" data-id="' + key + '" attr-category="' + catID + '"><i class="ms-Breadcrumb-chevron ms-Icon ms-Icon--ChevronLeft"></i><label class="ms-Breadcrumb-itemLink">' + oneBreadCrumbArr[key].category + '</label></li>');
                    $(document).on("click", "#one ul.breadcrumb li", function (event) {
                        $('#emptydataOne').hide();
                        $("#txtSearch").val("");
                        slideSearchQuery = "";
                        $("#one ul.childCategoryList").html("");
                        $("#one ul.childCategoryItems").html("");
                        var index = parseInt($(this).find("a").attr("data-id"));
                        if (index < oneBreadCrumbArr.length) {
                            oneBreadCrumbArr = oneBreadCrumbArr.splice(0, oneBreadCrumbArr.length - (oneBreadCrumbArr.length - index));
                            if (oneBreadCrumbArr.length == 0) {
                                $("#one ul.breadcrumb").hide();
                            }
                            var sli = $("#one ul.breadcrumb li a").filter(function () {
                                return $(this).attr("data-id") >= index;
                            });
                            if ($(sli).length > 0) $(sli).parent().remove();
                            if (parseInt($(this).find("a").attr("data-id")) - 1 == -1) {
                                oneBreadCrumbArr = [];
                                var $_breaCrumbUL = $("#one ul.breadcrumb");
                                $_breaCrumbUL.html("");
                                $("#hdnTextCategory").val("");
                                $("#hdnTextCategoryName").val("");

                                $("#one ul.categoryList").show();
                                $("#one ul.categoryItems").show();
                                $("#one ul.childCategoryList").hide();
                                $("#one ul.childCategoryItems").hide();
                            } else {
                                slideMaxID = 0;
                                $("#hdnTextCategory").val($(this).find("a").attr("attr-category"));
                                $("#hdnTextCategoryName").val(oneBreadCrumbArr[index - 1].category);
                                loadcategoryandSlides($(this).find("a").attr("attr-category"));
                            }
                        }
                    });
                }
            }
        }
        loadcategoryandSlides($_category);
        $("#WaitDialog").show();
        $("#hdnTextCategory").val($_category);
        $("#hdnTextCategoryName").val($_categoryName);
    }

    function loadcategoryandSlides(categoryId) {
        $('#emptydataOne').hide();
        $("#WaitDialog").show();

        var catContainer = "";
        var itemContainer = "";
        var filterCondition = "";

        if (categoryId == null) {

            catContainer = $("#one ul.categoryList");
            itemContainer = $("#one ul.categoryItems");

            $("#one ul.categoryItems").show();
            $("#one ul.categoryList").show();
            $("#one ul.childCategoryList").hide();
            $("#one ul.childCategoryItems").hide();

            filterCondition = "((CategoryType eq 'Slides') and (CategoryLevel eq 0))";
        }
        else {
            $("#one ul.childCategoryList").html("");
            $("#one ul.childCategoryItems").html("");

            catContainer = $("#one ul.childCategoryList");
            itemContainer = $("#one ul.childCategoryItems");

            $("#one ul.categoryItems").hide();
            $("#one ul.categoryList").hide();
            $("#one ul.childCategoryList").show();
            $("#one ul.childCategoryItems").show();

            filterCondition = "((CategoryType eq 'Slides') and (ParentCategory eq " + categoryId + "))";
        }
        LoadCategoryandSlidesOnPageLoad(filterCondition, catContainer, itemContainer, categoryId);
    }

    function loadImagesByCategory($_category) {
        imageMaxID = 0;

        $("#two ul.categoryList").hide();
        $("#two ul.categoryItems").hide();

        $("#two ul.childCategoryList").show();
        $("#two ul.childCategoryItems").show();
        var $_ulList = $("#two ul.childCategoryList");
        var $_ulItems = $("#two ul.childCategoryItems");

        var results = null;
        try {
            new Promise(function (callback, reject) {
                var $_ul = $("#two ul.childCategoryList");
                $_ul.html("");
                var isCategoryAvailable = false;
                if (imageListSite.toLowerCase() == '/sites/docsnodeadmin/' || imageListSite.toLowerCase() == '/sites/docsnodeadmin') {
                    loadDataFromSharePoint(DocsNodeCategoryDisplayName, 'ID,Title,Category,ParentCategory,CategoryType,CategoryLevel', "((CategoryType eq 'Images') and (ParentCategory eq " + $_category + "))").then(function (data) {
                        results = data.results;
                        if (results.length > 0) {
                            $.each(results, function (index, value) {
                                var cateTitle = value.Category;
                                var str = "<li>" +
                                    "<a href='#' class='liTitle' attr-title='" + cateTitle + "' attr-category='" + value.ID + "'>" +
                                    "<span class='file-icon'><i class='ms-Icon ms-Icon--OpenFolderHorizontal' aria-hidden='true'></i></span>" +
                                    "<span class='fileName' title='" + cateTitle + "'>" + (cateTitle.length > 20 ? cateTitle.substr(0, 20) + "..." : cateTitle) + "</span>" +
                                    "<span class='fileEnter-icon'><i class='ms-Icon ms-Icon--ChevronRightSmall' aria-hidden='true'></i></span>" +
                                    "</a >" +
                                    "</li>"
                                $_ul.append(str);
                            });
                            $("#WaitDialog").hide();
                            $('#emptydataTwo').hide();
                        }
                        callback(1);
                    });
                }
            }).then(function (result) {
                var $_ul = $("#two ul.childCategoryItems");
                $_ul.html("");
                if (imageListSite.toLowerCase() == '/sites/docsnodeadmin/' || imageListSite.toLowerCase() == '/sites/docsnodeadmin') {
                loadDataFromSharePoint(DocsNodePictureDisplayName, 'ID,Title,LinkFilenameNoMenu,EncodedAbsThumbnailUrl,EncodedAbsUrl,ImageCategory/ID,ImageCategory/Title&$expand=ImageCategory', "(ImageCategory/ID eq " + $_category + ")", "image").then(function (data) {
                    results = data.results;
                    if (results.length > 0) {
                        imageMaxID = _.max(_.pluck(results, "ID"));
                        $.each(results, function (index, image) {
                            toDataURL(image.EncodedAbsThumbnailUrl.toString(), function (dataUrl) {
                                var base64result = "data:image/png;base64, " + dataUrl;
                                $_ul.append("<li class='liIteamImg' onClick=\"insertImage('" + image.EncodedAbsUrl.toString() + "')\"><a href='#'><div class='liInnerImage'><img title='" + (image.Title != null ? image.Title : image.LinkFilenameNoMenu) + "' src='" + base64result + "'>" +
                                    "</div></a></li>");
                                $("#WaitDialog").hide();
                            });
                        });
                        $('#emptydataTwo').hide();
                    } else {
                        if ($_ulList[0].childElementCount == 0 && $_ulItems[0].childElementCount == 0) {
                            $('#emptydataTwo').show();
                            $("#WaitDialog").hide();
                        } else {
                            $('#emptydataTwo').hide();
                        }
                    }
                });
            }
            });
        } catch (error) {
            console.log('loadImagesByCategory: ' + error);
        }

    }

    function LoadCategories(category_name, columns, filters) {
        var dfd = $.Deferred();

        var url = SPURL + templateServerRelURL + "_api/web/lists(guid'" + category_name + "')/items?$select=" + columns + "&$filter=(" + filters + ")";

        try {
            callAjaxGet(url).done(function (data) {
                dfd.resolve(data.d);
            });
        } catch (err) {
            dfd.reject(err);
        }
        return dfd;
    }

    function bindSlidesHTML(data, itemContainer) {
        $.each(data, function (index, slides) {
            var filename = slides.LinkFilename;
            var title = slides['Title'];
            var insertSlideHtml = 'insertSlide(\'' + slides.EncodedAbsUrl + '\')';
            itemContainer.append("<li title='" + title + "' onclick=" + insertSlideHtml + " class='liIteamImg'><a href='#' class='liInnerImage'><label class='lblId' style='display:none'>" + slides.Id +
                "</label><span class='SubSicon'></span><span class='SubScontent'><img src='" + SPURL + templateServerRelURL + "_layouts/15/getpreview.ashx?path=" + slides.EncodedAbsUrl + "&resolution=0" + "'/></span></a></li>");
        });
        $("#WaitDialog").hide();
        $('#emptydataOne').hide();
    }

    function insertSlide(url) {
        Office.context.document.getSelectedDataAsync(Office.CoercionType.SlideRange,
            function (asyncResult) {
                var currentSlideIndex = asyncResult.value.slides[0].index;
                var presentation = new ActiveXObject("powerpoint.Application"); // Used to store a new presentation for each slide.
                presentation.ActivePresentation.Slides.InsertFromFile(url, currentSlideIndex, 1, 1);
            }
        );
    }

    function toDataURL(fileURL, callback) {
        try {
            fileURL = fileURL.replace(SPURL, "");
            var url = SPURL + imageListSite + "_api/web/GetFileByServerRelativeUrl('" + fileURL + "')/OpenBinaryStream";
            var xhr = new window.XMLHttpRequest();
            xhr.open("GET", url, true);
            xhr.setRequestHeader("Accept", "application/json; odata=verbose");
            xhr.setRequestHeader("Authorization", "Bearer " + SPToken);
            //Now set response type
            xhr.responseType = 'arraybuffer';
            xhr.addEventListener('load', function () {
                if (xhr.status === 200) {
                    filereader().then(function (imagedata) {
                        var dataUrl = imagedata;
                        var base64 = dataUrl.split(',')[1];
                        callback(base64);
                    });
                }
                function filereader() {
                    var dfdImg = $.Deferred();
                    var sampleBytes = new Uint8Array(xhr.response);
                    var blob = new Blob([sampleBytes], { type: "image/jpeg" });
                    var reader = new FileReader();
                    reader.onload = function () {
                        dfdImg.resolve(reader.result);
                    }
                    reader.readAsDataURL(blob);
                    return dfdImg.promise();
                }
            });
            xhr.send();
        } catch (error) {
            console.log('toDataUrl: ' + error);
        }
    }

function loadDataFromSharePoint(category_name, columns, filters, type, search) {
        var dfd = $.Deferred();
        var url = "";

        // type is not undefined means lazy loading should be done.
        if (type != undefined) {
            if (type == "image") {
                if (imageListSite.toLowerCase() == '/sites/docsnodeadmin/' || imageListSite.toLowerCase() == '/sites/docsnodeadmin') {
                    //columns = 'ID,Title,LinkFilenameNoMenu,EncodedAbsThumbnailUrl,EncodedAbsUrl,ImageCategory/ID,ImageCategory/Title&$expand=ImageCategory';
                    //filters = '(ImageCategory eq null)';
                    if (imageListSite.toLowerCase() == "/sites/docsnodeadmin") {
                        imageListSite = imageListSite + "/";
                    }
                }
                else {
                    columns = 'ID,Title,LinkFilenameNoMenu,EncodedAbsThumbnailUrl,EncodedAbsUrl';
                    if (!search)
                        filters = '';
                } 

                if ($("#hdnCategory").val() == "" || $("#hdnCategory").val() == null || $("#hdnCategory").val() == undefined)// It means parent level
                    url = SPURL + imageListSite + "_api/web/lists(guid'" + category_name + "')/items?$select=" + columns + "&$filter=(" + (filters ? (filters + " and") : '') + " ID gt " + imageParentMaxID + ")&$top=" + imageItem;
                else
                    url = SPURL + imageListSite + "_api/web/lists(guid'" + category_name + "')/items?$select=" + columns + "&$filter=(" + (filters ? (filters + " and") : '') + " ID gt " + imageMaxID + ")&$top=" + imageItem;
            }
            else if (type == "slide") {
                if ($("#hdnTextCategory").val() == "" || $("#hdnTextCategory").val() == null || $("#hdnTextCategory").val() == undefined)
                    url = SPURL + templateServerRelURL + "_api/web/lists(guid'" + category_name + "')/items?$select=" + columns + "&$filter=(" + filters + " and ID gt " + slideParentMaxID + ")&$top=" + slideItem;
                else
                    url = SPURL + templateServerRelURL + "_api/web/lists(guid'" + category_name + "')/items?$select=" + columns + "&$filter=(" + filters + " and ID gt " + slideMaxID + ")&$top" + slideItem;
            }
        }
        else {
            if (imageListSite.toLowerCase() != '/sites/docsnodeadmin/' && imageListSite.toLowerCase() != '/sites/docsnodeadmin') {
                url = SPURL + imageListSite + "_api/web/lists(guid'" + category_name + "')/items?$select=" + columns;
            }
            else {
                url = SPURL + templateServerRelURL + "_api/web/lists(guid'" + category_name + "')/items?$select=" + columns + "&$filter=(" + filters + ")";
            }
        }

        try {
            callAjaxGet(url).then(function (data) {
                dfd.resolve(data.d);
            }, function (error) {
                dfd.reject(error);
            });
        } catch (err) {
            dfd.reject(err);
        }

        return dfd;
    }

    function callAjaxGet(url) {
        var dfdGET = $.Deferred();
        try {
            if (SPToken) {
                $.ajax({
                    url: url,
                    method: "GET",
                    headers: { "Accept": "application/json; odata=verbose", 'Authorization': 'Bearer ' + SPToken },
                    success: function (data) {
                        dfdGET.resolve(data);
                    },
                    error: function (data) {
                        console.log(data);
                        dfdGET.reject(data);
                    }
                });
            }
        } catch (error) {
            console.log("callAjaxGet: " + error);
            dfdGET.reject(error);
        }
        return dfdGET.promise();
    }


    function insertImage(fileURL) {
        toDataURL(fileURL, function (base64result) {
            Office.context.document.setSelectedDataAsync(base64result, {
                coercionType: Office.CoercionType.Image
            },
                function (asyncResult) {
                    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                        console.log(asyncResult.error.message);
                    }
                });
        });
    }

