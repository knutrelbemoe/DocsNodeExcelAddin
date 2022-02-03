"use strict";
var DocsNodeTemplateDesigner = window.DocsNodeTemplateDesigner || {};
var isAdmin;
var isRoot = true;
var controlsID = [];
$(document).ready(function () {
    //if (localStorage.getItem("isOpen")) {
    //    $("#mainContent").removeClass("hide");
    //    $("#getStartedDiv").addClass("hide");
    //}
    //$("#startDiv").on("click", function () {
    //    $("#mainContent").removeClass("hide");
    //    $("#getStartedDiv").addClass("hide");
    //    localStorage.setItem("isOpen", true);
    //});
    
    var utility = new DocsNodeTemplateDesigner.Utility();
    var docsNode = new DocsNodeTemplateDesigner.load();
    var tenantName = config.endpoints.sharePointUrl.substr(8, config.endpoints.sharePointUrl.length);
    docsNode.init();
    $('#lblTenantUrl').text(config.endpoints.sharePointUrl);
    $('#SPLists').change(docsNode.getFieldsFromListBasedOnSelectedList);
    $("#listOfInputFields").on("click", ".fa-close", docsNode.onDeleteControlClick);
    $("#btnCreateUserInput").on("click", function () {
        if ($('#txtInputUserControl').val() != null && $('#txtInputUserControl').val() !== "") {
            docsNode.createUserInputControl();
        }
        else {
            utility.showErrorMessage('Please add input field name!');
        }
    });
    $("#divInputFields").on("click", "#btnCreateInput", function () {
        docsNode.createInputControle();
    });

    $("#listOfInputFields").on("click", ".fa-edit", function () {
        var controlId = $(this).parent().parent("li").attr("id");
        docsNode.editInputcontrol(controlId);
    });
    $('#dltInputCtrl').keypress(function (event) {
        var keycode = (event.keyCode ? event.keyCode : event.which);
        if (keycode == '13') {
            $(this).find("button.btnDeleteInputcontrol").trigger("click");
        }
        event.stopPropagation();
    });
    $(".btnDeleteInputcontrol").click(function () {
        docsNode.deleteInputControls($("#currentSelectedControl").val());
    });

    $("#btnSetPlaceholderValue").click(function () {
        if ($("#ddlInputFiedToSetValue").val() !== "0" && $("#SPListInputFieldValue").val() !== "0") {
            docsNode.setValuesofPlaceHolders($('#ddlInputFiedToSetValue option:selected'), $('#SPListInputFieldValue option:selected'));
        }
        else {
            utility.showErrorMessage("Please select input field and its value!");
        }

        $('#SPListInputFieldValue').val('0');
    });
    $("#listOfPlaceHolder").on("click", ".fa-plus", docsNode.createDuplicateContentControls);

    $("#inputLI").on("click", function () {
        $('#btnDiv').show();
    });
    $("#divPlaceholder").on("click", "#ddlInputFied", docsNode.getFieldsFromListBasedOnSelectedInputField);
    $("#divSetValue").on("change", "#ddlInputFiedToSetValue", docsNode.getItemsFromListBasedOnInuputField);

    $("#btnCreateInputList").on("click", function () {
        if ($(this).hasClass('active')) {
            $(this).children('.fa-plus').show();
            $(this).children('.fa-minus').hide();
            $('#inputListControl').hide();
            $(this).removeClass('active');
        }
        else {
            $('#inputUserControl').hide();
            $('#inputListControl').show();            
            $('.fa-plus').show();
            $('.fa-minus').hide()
            $(this).children('.fa-plus').hide();
            $(this).children('.fa-minus').show();
            $(this).addClass('active');
            $("#btnCreateInputUser").removeClass('active')
        }
    });
    $("#btnCreateInputUser").on("click", function () {
        if ($(this).hasClass('active')) {
            $(this).children('.fa-plus').show();
            $(this).children('.fa-minus').hide();
            $('#inputUserControl').hide();
            $(this).removeClass('active');
        }
        else {
            $('#inputListControl').hide();
            $('#inputUserControl').show();
            $('#txtInputUserControl').val('');            
            $('.fa-plus').show();
            $('.fa-minus').hide()
            $(this).children('.fa-plus').hide();
            $(this).children('.fa-minus').show();
            $(this).addClass('active');
            $("#btnCreateInputList").removeClass('active')
        }
    });
    $("#SPSiteCollections").change(function () {
        docsNode.clearDropDowns("siteCollection");
        if ($(this).val() != "0") {
            var selectedSiteCollection = $('#SPSiteCollections option:selected');
            var selectedSiteURL = selectedSiteCollection.attr('url').split("/sites/")[1];
            var rootSiteURL = selectedSiteCollection.attr('url').split(tenantName)[1];
            var selectedSiteURLForLib = selectedSiteCollection.attr('url').split("/sites/")[1] == undefined ? rootSiteURL : selectedSiteURL;
            isRoot = selectedSiteCollection.attr('url').split("/sites/")[1] == undefined ? true : false;
            docsNode.getAllsubSites(selectedSiteURL, true, "");
            docsNode.getAllListsFromSite(selectedSiteURLForLib);
        }
        else {
            utility.showErrorMessage('Please select site collection!');
        }
    });

    $("#SPSubsites").change(function () {
        docsNode.clearDropDowns("site");
        if ($(this).val() != "0") {
            var selectedSite = $('#SPSubsites option:selected');
            var selectedSiteURL = selectedSite.attr('url').split("/sites/")[1];
            var rootSiteURL = selectedSite.attr('url').split(tenantName)[1];
            var selectedSiteURLForLib = selectedSite.attr('url').split("/sites/")[1] == undefined ? rootSiteURL : selectedSiteURL;
            isRoot = selectedSite.attr('url').split("/sites/")[1] == undefined ? true : false;
            docsNode.getAllListsFromSite(selectedSiteURLForLib);
        }
        else {
            var selectedSiteCollection = $('#SPSiteCollections option:selected');
            var selectedSiteURL = selectedSiteCollection.attr('url').split("/sites/")[1];
            docsNode.getAllListsFromSite(selectedSiteURL);
        }
    });
});

DocsNodeTemplateDesigner.load = function () {
    var getAllProps = "";
    var WebRelativeUrl = "";
    var ListDisplayName = "";
    var utility = new DocsNodeTemplateDesigner.Utility();
    $('#adminProps').hide();
    isTenantAdmin().done(function (isAdmin) {
        if (isAdmin) {
            $('#adminProps').show();
            $('#chkAllProps').change(function () {
                if ($(this).prop('checked')) {
                    getAllProps = true;
                }
                else {
                    getAllProps = false;
                }
            });
        }
        else
            $('#adminProps').hide();
    });
 
    this.init = function () {
        utility.openWaitDialog();
        getAllSiteCollection(config.endpoints.sharePointUrl,
            function (query) {
                var resultsCount = query.PrimaryQueryResult.RelevantResults.RowCount;
                var listOfSiteCollections = "<option value='0'>Select</option>";
                for (var i = 0; i < resultsCount; i++) {
                    var row = query.PrimaryQueryResult.RelevantResults.Table.Rows.results[i];
                    var siteUrl = row.Cells.results[6].Value;
                    var siteTitle = row.Cells.results[3].Value;
                    listOfSiteCollections += "<option url= '" + siteUrl + "' value='" + siteTitle + "'>" + siteTitle + "</option>";
                }
                $('#SPSiteCollections').html(listOfSiteCollections);
                utility.sortData($('#SPSiteCollections'));
            },
            function (error) {
                console.log(JSON.stringify(error));
            }
        );
    }

    function isTenantAdmin() {
        var deferred = $.Deferred();
        utility.getAccessToken(true).done(function (token) {
            $.ajax({
                url: 'https://graph.microsoft.com/v1.0/me/memberOf',
                method: 'GET',
                headers: {
                    'Accept': 'application/json', 'Authorization': 'Bearer ' + token
                }
            }).success(function (data) {
                isAdmin = false;
                var obj = data.value[0];
                if (data.value.length > 0 && obj.displayName === 'Company Administrator') {
                    isAdmin = true;
                }
                deferred.resolve(isAdmin);
            }).error(function (err) {
                deferred.resolve(false);
            });
        });
        return deferred.promise();
    }



    function getAllSiteCollection(webUrl, success, failure) {
        utility.openWaitDialog();
        utility.getAccessToken(false).done(function (token) {
            // to create Personal site URL
            var tempArray = webUrl.split(".");
            var mySitePath = tempArray[0] + "-my." + tempArray[1] + "." + tempArray[2] + "/personal";
            var url = webUrl + "/_api/search/query?querytext='NOT Path:" + mySitePath + "/* contentclass:sts_site'&rowLimit=499&TrimDuplicates=false";
            $.ajax({
                url: url,
                method: "GET",
                headers: { "Accept": "application/json; odata=verbose", 'Authorization': 'Bearer ' + token },
                success: function (data) {
                    success(data.d.query);
                },
                error: function (data) {
                    failure(data);
                }
            });
        });

    };

    this.getAllsubSites = function (siteCollection, isCollection, customAttr) {
        utility.openWaitDialog();
        var docsNode = new DocsNodeTemplateDesigner.load();
        if (siteCollection === "0") {
            utility.showErrorMessage('Please select site collection!');
        }
        else {
            utility.getAccessToken(true).done(function (token) {
                var sharepointUrl = config.endpoints.sharePointUrl;
                var tenantName = sharepointUrl.substr(8, sharepointUrl.length);
                if (token) {
                    if (isCollection) {
                        if (siteCollection != undefined) {
                            var GraphAPI = "https://graph.microsoft.com/v1.0/sites/" + tenantName + ":/sites/" + siteCollection + ":/sites";
                        }
                        else {
                            var GraphAPI = "https://graph.microsoft.com/v1.0/sites/" + tenantName + "/sites/";
                        }                        
                    }
                    else {
                        if (siteCollection.split(":")[1] == undefined) {
                            var GraphAPI = "https://graph.microsoft.com/v1.0/sites/" + tenantName + ":/sites/" + siteCollection + ":/sites";
                        }
                        else {
                            var GraphAPI = "https://graph.microsoft.com/v1.0/sites/" + tenantName + ":/" + siteCollection + "/sites";
                        }                        
                    }

                    $.ajax({
                        beforeSend: function (request) {
                            request.setRequestHeader("Accept", "application/json");
                        },
                        type: "GET",
                        url: GraphAPI,
                        dataType: "json",
                        headers: {
                            'Authorization': 'Bearer ' + token
                        }
                    }).done(function (response) {
                        var result = response.value;
                        if (result && result.length > 0) {
                            var listOfSubSites = "";
                            for (var i = 0; i < result.length; i++) {
                                if (result[i].webUrl.indexOf(tenantName) > -1) {
                                    listOfSubSites += "<option siteHierarchy='" + customAttr + (i + 1) + "' url= '" + result[i].webUrl + "' value='" + result[i].name + "'>" + result[i].name + "</option>";
                                    var rootsubSite = result[i].webUrl.split(tenantName)[1] + ":";
                                    var subsitesVar = result[i].webUrl.split("/sites/")[1] == undefined ? rootsubSite : result[i].webUrl.split("/sites/")[1];                                    
                                    docsNode.getAllsubSites(subsitesVar, false, customAttr + (i + 1) + ".");
                                }
                            }
                            $("#SPSubsites").append(listOfSubSites);
                        }
                        else {
                            utility.closeWaitDialog();
                        }
                    }).fail(function (response) {
                        console.log('error:- ' + response.responseText);
                        utility.closeWaitDialog();
                    });
                }
            });
        }
    }

    this.getAllListsFromSite = function (site) {
        if (site === "0") {
            utility.showErrorMessage('Please select site collection!');
        }
        else {
            utility.openWaitDialog();
            utility.getAccessToken(true).done(function (token) {
                var sharepointUrl = config.endpoints.sharePointUrl;
                var tenantName = sharepointUrl.substr(8, sharepointUrl.length);
                if (token) {
                    if (site == undefined) {
                        var GraphAPI = "https://graph.microsoft.com/v1.0/sites/" + tenantName + "/lists";
                    }
                    else if (site.split("/sites/")[1] == undefined & (isRoot)) {
                        if (site == "") {
                            var GraphAPI = "https://graph.microsoft.com/v1.0/sites/" + tenantName + "/lists";
                        }
                        else {                           
                            var GraphAPI = "https://graph.microsoft.com/v1.0/sites/" + tenantName + ":" + site + ":/lists";
                        }
                    }
                    else {
                        var GraphAPI = "https://graph.microsoft.com/v1.0/sites/" + tenantName + ":/sites/" + site + ":/lists";
                    }                   

                    $.ajax({
                        beforeSend: function (request) {
                            request.setRequestHeader("Accept", "application/json");
                        },
                        type: "GET",
                        url: GraphAPI,
                        dataType: "json",
                        headers: {
                            'Authorization': 'Bearer ' + token
                        }
                    }).done(function (response) {
                        var result = response.value;
                        var listOfLibraries = "<option value='0'>Select</option>";
                        for (var i = 0; i < result.length; i++) {
                            if (result[i].list.template === "genericList" && !result[i].list.hidden) {
                                listOfLibraries += "<option guid='" + result[i].id +"'   internalname= '" + result[i].name + "' siteURL='" + site + "' value='" + result[i].displayName + "'>" + result[i].displayName + "</option>";
                            }
                        }
                        $("#SPLists").html(listOfLibraries);
                        utility.closeWaitDialog();
                    }).fail(function (response) {
                        console.log('error:- ' + response.responseText);
                        utility.closeWaitDialog();
                    });
                }
            });
        }

    }

    function setValuesOfUserPlaceholders(inputFieldDetail, item, isPageLoad) {
        utility.openWaitDialog();
        var GraphAPI;
        utility.getAccessToken(true).done(function (token) {
            if (token) {
                if (inputFieldDetail.checkCurrentUser && $('#SPListInputFieldValue').val() === "0") {
                    GraphAPI = "https://graph.microsoft.com/v1.0/me/";
                }
                else {                    
                    if (item != null && item != '' && item != undefined)
                        GraphAPI = "https://graph.microsoft.com/v1.0/users/" + item.attr('data-Text') + "?$select=aboutMe,birthday,businessPhones,city,companyName,country,department,id,interests,jobTitle,mail,mySite,officeLocation,pastProjects,postalCode,preferredLanguage,preferredName,responsibilities,schools,skills,state,streetAddress,surname,userPrincipalName,userType,displayName,givenName";
                }
                $.ajax({
                    beforeSend: function (request) {
                        request.setRequestHeader("Accept", "application/json");
                    },
                    type: "GET",
                    url: GraphAPI,
                    dataType: "json",
                    headers: {
                        'Authorization': 'Bearer ' + token,
                        'Prefer': 'HonorNonIndexedQueriesWarningMayFailRandomly',
                    }
                }).done(function (response) {
                    var itemVal = response.value === undefined ? response : response.value[0];
                    if (response) {
                        Word.run(function (context) {
                            var contentControls = context.document.contentControls;
                            context.load(contentControls, ["tag", "title", "id", "text"]);
                            return context.sync().then(function () {
                                var numberOfItem = contentControls.items.length;
                                if (numberOfItem && numberOfItem > 0) {
                                    for (var i = 0; i < numberOfItem; i++) {
                                        var tagName = getInfromationFromTage(contentControls.items[i].tag);
                                        if (tagName[0] == inputFieldDetail.InuptControleName) {
                                            var selectedField = itemVal[tagName[2]];
                                            if (tagName[2] == "skills" || tagName[2] == "businessPhones" || tagName[2] == "schools" || tagName[2] == "responsibilities" || tagName[2] == "pastProjects" || tagName[2] == "interests") {
                                                contentControls.items[i].insertText(selectedField.join(), 'replace');
                                            }
                                            else if (selectedField != null && selectedField != "" & selectedField != undefined) {
                                                if ((isPageLoad) && (contentControls.items[i].text)) {
                                                }
                                                else {
                                                    contentControls.items[i].insertText(selectedField, 'replace');
                                                }
                                            }
                                            else {
                                                contentControls.items[i].clear();
                                            }
                                        }
                                    }                                   
                                }
                            })
                        })
                            .catch(function (error) {
                                console.log('Error: ' + JSON.stringify(error));
                                if (error instanceof OfficeExtension.Error) {
                                    console.log('Debug info: ' + JSON.stringify(error.debugInfo));
                                }
                            });
                    }
                    else {
                        utility.closeWaitDialog();
                    }
                    utility.closeWaitDialog();

                }).fail(function (response) {                   
                    utility.closeWaitDialog();
                });              
            }
            else {
                utility.closeWaitDialog();
            }

        }).fail(function (error) {
            utility.showErrorMessage("There was some issue to get Token");
            utility.closeWaitDialog();
        });
    }

    this.setValuesofPlaceHolders = function (inputField, item) {
        utility.openWaitDialog();
        var inputFieldDetail = JSON.parse(inputField.attr('Value'));
        WebRelativeUrl = inputFieldDetail.WebRelativeURL;
        ListDisplayName = inputFieldDetail.ListDisplayName;
        if (ListDisplayName == "" || ListDisplayName == null) {
            setValuesOfUserPlaceholders(inputFieldDetail, item, false);
        }
        else {
            var fieldName = inputFieldDetail.FieldInternalName;
            var tenantName = utility.getTenantName();
            var WebRelativeGraphAPI = utility.CreateGraphAPIToWeb(tenantName, WebRelativeUrl);
            utility.getAccessToken(true).done(function (token) {
                if (token) {                   
                    //var GraphAPI = WebRelativeGraphAPI + "/lists/" + ListDisplayName + "/items/" + item.attr("itemID");
                    var GraphAPI = WebRelativeGraphAPI + "/lists/" + inputFieldDetail.GUID + "/items/" + item.attr("itemID");
                    $.ajax({
                        beforeSend: function (request) {
                            request.setRequestHeader("Accept", "application/json");
                        },
                        type: "GET",
                        url: GraphAPI,
                        dataType: "json",
                        headers: {
                            'Authorization': 'Bearer ' + token,
                            'Prefer': 'HonorNonIndexedQueriesWarningMayFailRandomly',
                        }
                    }).done(function (response) {
                        if (response) {
                            Word.run(function (context) {
                                var contentControls = context.document.contentControls;
                                context.load(contentControls, ["tag", "title", "id"]);
                                return context.sync().then(function () {
                                    var numberOfItem = contentControls.items.length;
                                    if (numberOfItem && numberOfItem > 0) {
                                        for (var i = 0; i < numberOfItem; i++) {
                                            var tagName = getInfromationFromTage(contentControls.items[i].tag);
                                            if (tagName[0] == inputFieldDetail.InuptControleName && tagName[1] == ListDisplayName) {
                                                //var selectedField = response.value[0].fields[tagName[2]];
                                                var selectedField = response.fields[tagName[2]];
                                                if (selectedField === true || selectedField === false) {
                                                    selectedField = selectedField == true ? "Yes" : "No";
                                                }
                                                else if (selectedField.Label !== undefined) {
                                                    selectedField = selectedField.Label;
                                                }
                                                else if (selectedField.Url !== undefined) {
                                                    selectedField = selectedField.Url;
                                                }
                                                var selectedField = selectedField.replace == undefined ? selectedField : selectedField.replace(/<[^>]+>/g, '');
                                                contentControls.items[i].insertText(selectedField.toString(), 'replace');
                                            }
                                        }
                                        return context.sync()
                                            .then(function () {
                                            });
                                    }
                                })
                            })
                                .catch(function (error) {
                                    if (error instanceof OfficeExtension.Error) {
                                        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
                                    }
                                });
                        }
                        else {
                            utility.closeWaitDialog();
                        }
                        utility.closeWaitDialog();

                    }).fail(function (response) {
                        console.log('error:- ' + response.responseText);
                        utility.closeWaitDialog();
                    });

                    return context.sync()
                        .then(function () {
                        });
                }
                else {
                    utility.closeWaitDialog();
                }

            }).fail(function (error) {
                console.log('error:- ' + response.responseText);
                utility.showErrorMessage("There was some issue to get Token");
                utility.closeWaitDialog();
            });
        }
        utility.closeWaitDialog();
    };
    this.getInputControles = function () {
        getInputControles();
    };

    this.listInputControls = function (inputcontrols) {
        var unique = inputcontrols.filter(function (itm, i, inputcontrols) {
            return i == inputcontrols.indexOf(itm);
        });

    }

    this.onDeletePlacehoderClick = function () {
        var tagName = getTagFromLI(this);
        $("#currentSelectedTag").val(tagName);
    }

    this.onDeleteControlClick = function () {
        var id = $(this).parent().parent('li').attr('id');
        $("#currentSelectedControl").val(id);
        var iscontrolAvailable = false;
        Word.run(function (context) {
            var contentControls = context.document.contentControls;
            context.load(contentControls, ["tag", "title", "id"]);
            return context.sync().then(function () {
                var numberOfItem = contentControls.items.length;
                var tags = [];
                if (numberOfItem && numberOfItem > 0) {
                    for (var i = 0; i < numberOfItem; i++) {
                        var item = contentControls.items[i];
                        if (item.tag != null && item.tag != undefined) {
                            var currentControlId = item.tag.split("¤")[3];
                            if (currentControlId === id) {
                                iscontrolAvailable = true;
                                break;
                            }
                            else {
                                iscontrolAvailable = false;
                            }
                        }
                    }
                    if (iscontrolAvailable) {
                        $("#dltInputCtrl").modal("hide");
                        utility.showErrorMessage('please remove placeholders created by using this input field!');
                    }
                    else {
                        $("#dltInputCtrl").modal("show");
                    }
                    return context.sync()
                        .then(function () {
                            utility.closeWaitDialog();
                        });
                }
                else {

                    $("#dltInputCtrl").modal("show");
                    utility.closeWaitDialog();
                }
            });
        })
            .catch(function (error) {
                console.log('Error: ' + JSON.stringify(error));
                if (error instanceof OfficeExtension.Error) {
                    console.log('Debug info: ' + JSON.stringify(error.debugInfo));
                }
                utility.closeWaitDialog();
            });
    }
    this.getItemsFromListBasedOnInuputField = function () {
        var inputFieldInformation = "";
        if ($(this).val() === "0") {
            $('#SPListInputFieldValue').find('option:not(:first)').remove();
            return;
        }
        else {
            inputFieldInformation = $(this).val();
        }
        utility.openWaitDialog();
        var columnInfromationArray = JSON.parse(inputFieldInformation);
        utility.getAccessToken(true).done(function (token) {
            if (token) {
                var GraphAPI = "";
                if (columnInfromationArray.WebRelativeURL != undefined) {
                    var WebRelativeGraphAPI = utility.CreateGraphAPIToWeb(columnInfromationArray.TenatName, columnInfromationArray.WebRelativeURL);
                    //var GraphAPI = WebRelativeGraphAPI + "/lists/" + columnInfromationArray.ListDisplayName + "/items?expand=fields(select=" + columnInfromationArray.FieldInternalName + ",ID)";
                    var GraphAPI = WebRelativeGraphAPI + "/lists/" + columnInfromationArray.GUID + "/items?expand=fields(select=" + columnInfromationArray.FieldInternalName + ",ID)";
                }
                else {
                    var GraphAPI = "https://graph.microsoft.com/v1.0/users?$filter=accountEnabled eq true";
                }


                $.ajax({
                    beforeSend: function (request) {
                        request.setRequestHeader("Accept", "application/json");
                    },
                    type: "GET",
                    url: GraphAPI,
                    dataType: "json",
                    headers: {
                        'Authorization': 'Bearer ' + token,
                    }
                }).done(function (response) {
                    var listOfItems = "<option selected value='0'>Select</option>";
                    var isDataFound = false;
                    if (response) {
                        if (response.value.length > 0) {
                            for (var i = 0; i < response.value.length; i++) {
                                if (columnInfromationArray.FieldInternalName != undefined) {
                                    var selectedField = response.value[i].fields[columnInfromationArray.FieldInternalName];
                                    var selectedItemID = response.value[i].fields["id"];
                                    if (selectedField !== undefined) {
                                        isDataFound = true;
                                        if (selectedField.Label !== undefined) {
                                            listOfItems += "<option internalName='" + selectedField.Label + "' data-text= '" + inputFieldInformation + "' itemId='" + selectedItemID + "' value='" + selectedField.Label + "'>" + selectedField.Label + "</option>";
                                        }
                                        else if (selectedField.Url !== undefined) {
                                            listOfItems += "<option internalName='" + selectedField.Url + "' data-text= '" + inputFieldInformation + "' itemId='" + selectedItemID + "' value='" + selectedField.Url + "'>" + selectedField.Url + "</option>";
                                        }
                                        //else if (isNaN(selectedField) && !isNaN(Date.parse(selectedField))) {
                                        //    var selectedDate = moment(selectedField).format('MM/DD/YYYY');
                                        //    listOfItems += "<option internalName='" + selectedDate + "' data-text= '" + inputFieldInformation + "' value='" + selectedDate + "'>" + selectedDate + "</option>";
                                        //}
                                        else {
                                            listOfItems += "<option internalName='" + selectedField + "' data-text= '" + inputFieldInformation + "' itemId='" + selectedItemID + "' value='" + selectedField + "'>" + selectedField + "</option>";
                                        }
                                    }
                                }
                                else {
                                    isDataFound = true;
                                    listOfItems += "<option data-text= '" + response.value[i].userPrincipalName + "'>" + response.value[i].displayName + "</option>";
                                }
                            }
                        }
                        else {
                            utility.showErrorMessage("selected list does not have values!");
                        }
                    }
                    else {
                        utility.showErrorMessage("selected list does not have values!");
                        utility.closeWaitDialog();
                    }
                    if (!isDataFound) {
                        $("#SPListItemNotification").css("display", "block");
                        $('#SPListInputFieldValue').prop("disabled", true)
                    }
                    else {
                        $("#SPListItemNotification").css("display", "none");
                        $('#SPListInputFieldValue').prop("disabled", false)
                    }
                    $('#SPListInputFieldValue').html(listOfItems);
                    utility.sortData($('#SPListInputFieldValue'));
                    utility.removeDuplicateValues();
                    utility.closeWaitDialog();
                }).fail(function (response) {
                    console.log('error:- ' + response.responseText);
                    utility.showErrorMessage("There was some issue. This site is doesn't exist or you don't have permission to access this site");
                    utility.closeWaitDialog();
                });
            }
            else {
                utility.closeWaitDialog();
            }
        }).fail(function (error) {
            console.log('error:- ' + response.responseText);
            utility.showErrorMessage("There was some issue to get Token");
            utility.closeWaitDialog();
        });
    }

    //this is old function that used when user Enter URL
    this.getCustomListFromWeb = function () {
        $('#SPListFields').find('option:not(:first)').remove();
        $('#SPLists').find('option:not(:first)').remove();
        utility.getAccessToken(true).done(function (token) {
            WebRelativeUrl = $('#txtSPRelativeURL').val();
            if (!utility.isValidWebURL(WebRelativeUrl)) {
                utility.closeWaitDialog();
                return false;
            }
            if (token) {
                utility.openWaitDialog();
                var GraphAPI = "";
                var tenantName = utility.getTenantName();
                var WebRelativeGraphAPI = utility.CreateGraphAPIToWeb(tenantName, WebRelativeUrl);
                var GraphAPI = WebRelativeGraphAPI + "/lists";
                $.ajax({
                    beforeSend: function (request) {
                        request.setRequestHeader("Accept", "application/json");
                    },
                    type: "GET",
                    url: GraphAPI,
                    dataType: "json",
                    headers: {
                        'Authorization': 'Bearer ' + token,
                    }
                }).done(function (response) {
                    var result = response.value;
                    var listOfLibrary = "<option value='0'>Select</option>";
                    if (response) {
                        for (var i = 0; i < result.length; i++) {
                            if (result[i].list.template === "genericList" && !result[i].list.hidden) {
                                listOfLibrary += "<option internalName='" + result[i].name + "' value='" + result[i].displayName + "'>" + result[i].displayName + "</option>";
                            }
                        }
                        utility.closeWaitDialog();
                    }
                    else {
                        utility.closeWaitDialog();
                    }
                    $('#SPLists').html(listOfLibrary);
                }).fail(function (response) {
                    console.log('error:- ' + response.responseText);
                    utility.showErrorMessage("There was some issue. This site is doesn't exist or you don't have permission to access this site");
                    utility.closeWaitDialog();
                });
            }
        }).fail(function (error) {
            console.log('error:- ' + response.responseText);
            utility.showErrorMessage("There was some issue to get Token");
            utility.closeWaitDialog();
        });
    };

    //Added by Chandni
    this.getFieldsFromListBasedOnSelectedList = function () {
        utility.openWaitDialog();
        var selectedList = $("#SPLists").val();
        clearDropDowns("list");
        if (selectedList === "0") {
            utility.showErrorMessage("please select List");
            utility.closeWaitDialog();
            return;
        }
        else {
            ListDisplayName = selectedList;
            $("#ListDisplayName").val(ListDisplayName);
        }
        var selectedSiteRelativeURL = "";
        if ($("#SPSubsites").val() === "0") {
            selectedSiteRelativeURL = $('option:selected', "#SPSiteCollections").attr('url');
            selectedSiteRelativeURL = selectedSiteRelativeURL ? selectedSiteRelativeURL.replace(config.endpoints.sharePointUrl, "") : "";
        }
        else {
            selectedSiteRelativeURL = $('option:selected', '#SPSubsites').attr('url');
            selectedSiteRelativeURL = selectedSiteRelativeURL ? selectedSiteRelativeURL.replace(config.endpoints.sharePointUrl, "") : "";
        }
        var item = {
            TenatName: utility.getTenantName(),
            WebRelativeURL: selectedSiteRelativeURL,
            ListDisplayName: ListDisplayName,
            GUID: $('option:selected', '#SPLists').attr('guid')
        };
        $("#WebRelativeUrl").val(selectedSiteRelativeURL);
        getFieldsFromList(item, "#SPListFields");

    }

    this.getFieldsFromListBasedOnSelectedInputField = function () {

        var selectedInputField = $("#ddlInputFied").val();
        if (selectedInputField !== "0") {
            utility.openWaitDialog();
            selectedInputField = JSON.parse(selectedInputField);
            if (selectedInputField.ListDisplayName != undefined) {
                getFieldsFromList(selectedInputField, "#SPListFieldForPlaceHolder");
            }
            else {
                getUserDetails(selectedInputField, "#SPListFieldForPlaceHolder");
            }

        }
        else {
            $('#SPListFieldForPlaceHolder').find('option:not(:first)').remove();
        }
    }

    function getUserDetails(item, ddlID) {
        if (isAdmin != undefined) {
            isTenantAdmin().done(function (isAdmin) {
                if ((isAdmin) && (getAllProps) && (item.checkAllProps)) {
                    $(ddlID).html('<option value="0">Select</option><option value="aboutMe">aboutMe</option><option value="birthday">birthday</option><option value="businessPhones">businessPhones</option><option value="city">city</option><option value="companyName">companyName</option><option value="country">country</option><option value="department">department</option><option value="displayName">displayName</option><option value="givenName">givenName</option><option value="hireDate">hireDate</option><option value="id">id</option><option value="interests">interests</option><option value="jobTitle">jobTitle</option><option value="mail">mail</option><option value="mobilePhone">mobilePhone</option><option value="mySite">mySite</option><option value="officeLocation">officeLocation</option><option value="pastProjects">pastProjects</option><option value="postalCode">postalCode</option><option value="preferredLanguage">preferredLanguage</option><option value="preferredName">preferredName</option><option value="responsibilities">responsibilities</option><option value="schools">schools</option><option value="skills">skills</option><option value="state">state</option><option value="streetAddress">streetAddress</option><option value="surname">surname</option><option value="userPrincipalName">userPrincipalName</option><option value="userType">userType</option>');
                }
                else {
                    $(ddlID).html("<option value='0'>Select</option><option value='displayName'>displayName</option><option value='givenName'>givenName</option><option value='mail'>mail</option><option value='surname'>surname</option><option value='userPrincipalName'>userPrincipalName</option>");
                }
                utility.closeWaitDialog();
            }).fail(function (isAdmin) {
                $(ddlID).html("<option value='0'>Select</option><option value='displayName'>displayName</option><option value='givenName'>givenName</option><option value='mail'>mail</option><option value='surname'>surname</option><option value='userPrincipalName'>userPrincipalName</option>");
            });
        }
        else {
            $(ddlID).html("<option value='0'>Select</option><option value='displayName'>displayName</option><option value='givenName'>givenName</option><option value='mail'>mail</option><option value='surname'>surname</option><option value='userPrincipalName'>userPrincipalName</option>");
        }
        utility.closeWaitDialog();


    }
    
    //Added by Bharat
    function getFieldsFromList(item, ddlID) {
        utility.getAccessToken(true).done(function (token) {
            if (token) {
                var WebRelativeGraphAPI = utility.CreateGraphAPIToWeb(item.TenatName, item.WebRelativeURL);
                //var GraphAPI = WebRelativeGraphAPI + "/lists/" + item.ListDisplayName + "?expand=columns&filter=readOnly eq false";
                var GraphAPI = WebRelativeGraphAPI + "/lists/" + item.GUID + "?expand=columns&filter=readOnly eq false";
                $.ajax({
                    beforeSend: function (request) {
                        request.setRequestHeader("Accept", "application/json");
                    },
                    type: "GET",
                    url: GraphAPI,
                    dataType: "json",
                    headers: {
                        'Authorization': 'Bearer ' + token,
                    }
                }).done(function (response) {
                    var result = response.columns;
                    var listOfFields = "<option value='0'>Select</option>";
                    if (result) {
                        for (var i = 0; i < result.length; i++) {
                            if (result[i].readOnly === true && response.columns[i].calculated != undefined) {
                                listOfFields += "<option internalName='" + result[i].name + "' value='" + result[i].displayName + "'>" + result[i].displayName + "</option>";
                            }
                            if (result[i].readOnly === false && result[i].displayName !== "Content Type" && result[i].displayName !== "Attachments" && result[i].personOrGroup === undefined && result[i].dateTime === undefined && result[i].lookup === undefined) {
                                listOfFields += "<option internalName='" + result[i].name + "' value='" + result[i].displayName + "'>" + result[i].displayName + "</option>";
                            }
                        }
                        utility.closeWaitDialog();
                    }
                    else {
                        utility.closeWaitDialog();
                    }
                    $(ddlID).html(listOfFields);
                }).fail(function (response) {
                    console.log('error:- ' + response.responseText);
                    utility.showErrorMessage("There was some issue. This site is doesn't exist or you don't have permission to access this site");
                    utility.closeWaitDialog();
                });
            }
        }).fail(function (error) {
            console.log('error:- ' + response.responseText);
            utility.showErrorMessage("There was some issue to get Token");
            utility.closeWaitDialog();
        });

    }
    

    this.deleteInputControls = function (selectedID) {
        Office.context.document.customXmlParts.getByIdAsync(selectedID, function (result) {
            var xmlPart = result.value;
            xmlPart.deleteAsync(function (eventArgs) {
                getInputControles();
            });
        });
    }

    this.deleteContentControls = function (selectedTag) {
        utility.openWaitDialog();
        Word.run(function (context) {
            var contentControls = context.document.contentControls.getByTag(selectedTag);
            context.load(contentControls, ["tag", "title", "id"]);//["id","tag"]
            return context.sync().then(function () {
                var numberOfItem = contentControls.items.length;
                if (numberOfItem && numberOfItem > 0) {
                    for (var i = 0; i < numberOfItem; i++) {
                        contentControls.items[i].clear();
                        contentControls.items[i].delete(true);
                    }
                    return context.sync()
                        .then(function () {
                            DocsNodeOfficeFn.loadContentControls();
                        });
                }
                else {
                    utility.closeWaitDialog();
                }
            });
        }).catch(function (error) {
            console.log('Error: ' + JSON.stringify(error));
            if (error instanceof OfficeExtension.Error) {
                console.log('Debug info: ' + JSON.stringify(error.debugInfo));
            }
            utility.closeWaitDialog();
        });
    }

    this.editInputcontrol = function (currentElementID) {
        var docsNode = new DocsNodeTemplateDesigner.load();
        Office.context.document.customXmlParts.getByIdAsync(currentElementID, function (result) {
            var xmlPart = result.value;
            xmlPart.getXmlAsync({}, function (eventArgs) {
                var inputControlXML = $.parseXML(eventArgs.value);
                var inputControlDetails = $(inputControlXML).find("parameters")[0].textContent;
                var details = JSON.parse(inputControlDetails);
                if (details.ListDisplayName != undefined) {
                    //$('#SPSiteCollections').val(details.siteCollectionURL);
                    //var selectedSiteCollection = $('#SPSiteCollections option:selected');
                    //var selectedSiteURL = selectedSiteCollection.attr('url').split("/sites/")[1];
                    //docsNode.getAllsubSites(selectedSiteURL, true, "");
                    //docsNode.getAllListsFromSite(selectedSiteURL);
                    //docsNode.getFieldsFromListBasedOnSelectedList();
                    //if ($("#SPSubsites").val() != "0") {
                    //    var selectedSite = $('#SPSubsites option:selected');
                    //    var selectedSiteURL = selectedSite.attr('url').split("/sites/")[1];
                    //    docsNode.getAllListsFromSite(selectedSiteURL);
                    //}
                    //setTimeout(function () {
                    //    $('#SPLists').val(details.ListDisplayName);                       
                    //}, 3000);
                    //setTimeout(function () {                        
                    //    $('#SPListFields').val(details.FieldDisplayName);
                    //    $('#SPSubsites').val(details.subsiteURL);
                    //}, 3000);
                    //$('#txtInputControle').val(details.InuptControleName);                   
                    //$('#inputListControl').show();
                    //$('#inputUserControl').hide();
                    //$('#btnDiv').hide();
                    //$("#divInputFields").on("click", "#btnCreateInput", function () {
                    //    var inputControleName = $('#txtInputControle').val().trim();
                    //    if ($('#SPListFields').val() !== '0' && inputControleName && inputControleName !== "") {
                    //        WebRelativeUrl = $("#WebRelativeUrl").val();
                    //        ListDisplayName = $("#ListDisplayName").val();
                    //        debugger
                    //        var tenantName = utility.getTenantName();
                    //        var selectedListField = $('option:selected', $("#SPListFields")).attr('internalName');
                    //        var updatedControlDetails = {
                    //            TenatName: tenantName,
                    //            WebRelativeURL: WebRelativeUrl,
                    //            ListDisplayName: ListDisplayName,
                    //            FieldInternalName: selectedListField,
                    //            InuptControleName: inputControleName,
                    //            FieldDisplayName: $("#SPListFields").val(),
                    //            siteCollectionURL: $("#SPSiteCollections").val(),
                    //            subsiteURL: $("#SPSubsites").val()
                    //        };
                    //    }
                    //    updateInputControl(currentElementID, updatedControlDetails);
                    //    $('#SPSiteCollections').val(0);
                    //    $('#SPSubsites').val(0);
                    //    $('#SPLists').val(0);
                    //    $('#SPListFields').val(0);
                    //    $('#txtInputControle').val('');
                    //    $('#inputListControl').hide();
                    //    $('#inputUserControl').hide();
                    //    $('#btnDiv').show();
                    //    getInputControles();
                    //});
                }
                else {
                    $("#btnCreateUserInput").addClass("hide");
                    $("#btnUpdateUserInput").removeClass("hide");
                    $('#inputUserControl').show();
                    $('#txtInputUserControl').val(details.InuptControleName);
                    $("#divInputFields").on("click", "#btnUpdateUserInput", function () {
                        var inputControleName = $('#txtInputUserControl').val().trim();
                        var tenantName = utility.getTenantName();
                        var checkCurrentUser = $('#chkCurrentUser').prop('checked');
                        var updatedDetails = {
                            TenatName: tenantName,
                            InuptControleName: inputControleName,
                            checkCurrentUser: checkCurrentUser
                        };
                        updateInputControl(currentElementID, updatedDetails);
                        $('#inputListControl').hide();
                        $('#inputUserControl').hide();
                        $('#btnDiv').show();
                        getInputControles();
                    });
                    $('#btnDiv').hide();
                    isTenantAdmin().done(function (isAdmin) {
                        if (isAdmin) {
                            $('#adminProps').show();
                            $('#chkAllProps').change(function () {
                                if ($(this).prop('checked')) {
                                    getAllProps = true;
                                }
                                else {
                                    getAllProps = false;
                                }
                            });
                        }
                        else
                            $('#adminProps').hide();
                    });
                }
            });
            //xmlPart.getNodesAsync('*', function (nodeResults) {
            //    for (i = 0; i < nodeResults.value.length; i++) {
            //        var node = nodeResults.value[i];
            //        node.setXmlAsync("<childNode>" + i + "</childNode>");
            //    }
            //});
        });
    }

    function updateInputControl(currentElementId, updatedDetails) {
        Office.context.document.customXmlParts.getByIdAsync(currentElementId, function (result) {
            var xmlPart = result.value;
            xmlPart.getNodesAsync('*', function (nodeResults) {
                var node = nodeResults.value[0];
                node.setXmlAsync("<parameters>" + JSON.stringify(updatedDetails) + "</parameters>");
            });
        });
    }
    this.createDuplicateContentControls = function (currentElement) {
        utility.openWaitDialog();
        var tag = getTagFromLI(currentElement.currentTarget);//.parent().parent("li").attr("tag");
        Word.run(function (context) {
            var contentControls = context.document.contentControls.getByTag(tag);
            context.load(contentControls, ["tag", "title", "id", "style"]);//["id","tag"]
            return context.sync().then(function () {
                var numberOfItem = contentControls.items.length;
                if (numberOfItem && numberOfItem > 0) {
                    var range = context.document.getSelection();
                    var myContentControl = range.insertContentControl();
                    myContentControl.tag = contentControls.items[0].tag;
                    myContentControl.title = contentControls.items[0].title;
                    myContentControl.style = contentControls.items[0].style;
                    return context.sync()
                        .then(function () {
                            DocsNodeOfficeFn.loadContentControls();
                        });
                }
                else {
                    utility.closeWaitDialog();
                }
            });
        }).catch(function (error) {
            console.log('Error: ' + JSON.stringify(error));
            if (error instanceof OfficeExtension.Error) {
                console.log('Debug info: ' + JSON.stringify(error.debugInfo));
            }
            utility.closeWaitDialog();
        });
    }

    this.createUserInputControl = function () {
        var inputControleName = $('#txtInputUserControl').val().trim();
        if ($('#txtInputUserControl').val() != null && inputControleName !== null && inputControleName && inputControleName !== "") {
            $('#txtInputUserControl').val('');
            var tenantName = utility.getTenantName();
            var checkCurrentUser = $('#chkCurrentUser').prop('checked');
            var checkAllProps = $('#chkAllProps').prop('checked');
            var inputControleDetails = {
                TenatName: tenantName,
                InuptControleName: inputControleName,
                checkCurrentUser: checkCurrentUser,
                checkAllProps: checkAllProps
            };
            var customXML = "<dataConnections xmlns='http://schema.binaryrepublik.com/2018/inputFieldsDetails'>" +
                " <parameters>" + JSON.stringify(inputControleDetails) + "</parameters> </dataConnections>";
            Word.run(function (context) {
                Office.context.document.customXmlParts.addAsync(customXML,
                    function (r) {
                        r.value.addHandlerAsync(Office.EventType.DataNodeInserted,
                            function (a) {
                            },
                            function (s) {
                                getInputControles();
                            });
                    });
                return context.sync()
                    .then(function () {
                        console.log("control added");
                    });
            }).catch(function (error) {
                if (error instanceof OfficeExtension.Error) {
                    console.log('Debug info: ' + JSON.stringify(error.debugInfo));
                }
                utility.closeWaitDialog();
            });
            $('#chkCurrentUser').prop('checked', false);
            $('#chkAllProps').prop('checked', false);
        }
        else {
            utility.showErrorMessage("Please add input control name!");
        }
    }


    this.createInputControle = function () {
        var inputControleName = $('#txtInputControle').val().trim();
        if ($('#SPListFields').val() !== '0' && inputControleName && inputControleName !== "") {
            WebRelativeUrl = $("#WebRelativeUrl").val();
            ListDisplayName = $("#ListDisplayName").val();
            var tenantName = utility.getTenantName();
            var selectedListField = $('option:selected', $("#SPListFields")).attr('internalName');
            var inputControleDetails = {
                TenatName: tenantName,
                WebRelativeURL: WebRelativeUrl,
                ListDisplayName: ListDisplayName,
                FieldInternalName: selectedListField,
                InuptControleName: inputControleName,
                FieldDisplayName: $("#SPListFields").val(),
                siteCollectionURL: $("#SPSiteCollections").val(),
                subsiteURL: $("#SPSubsites").val(),
                GUID: $('option:selected', '#SPLists').attr('guid')
            };

            var customXML = "<dataConnections xmlns='http://schema.binaryrepublik.com/2018/inputFieldsDetails'>" +
                " <parameters>" + JSON.stringify(inputControleDetails) + "</parameters> </dataConnections>";
            Word.run(function (context) {
                Office.context.document.customXmlParts.addAsync(customXML,
                    function (r) {
                        r.value.addHandlerAsync(Office.EventType.DataNodeInserted,
                            function (a) {
                            },
                            function (s) {
                                getInputControles();
                            });
                    });
                return context.sync()
                    .then(function () {
                        console.log("control added");
                    });
            }).catch(function (error) {
                if (error instanceof OfficeExtension.Error) {
                    console.log('Debug info: ' + JSON.stringify(error.debugInfo));
                }
                utility.closeWaitDialog();
            });
            $('#SPSiteCollections').val(0);
            $('#SPSubsites').val(0);
            $('#SPLists').val(0);
            $('#SPListFields').val(0);
            $('#txtInputControle').val('');
        }
        else {
            utility.showErrorMessage("Please fill all necessary fields!");
        }
    }


    function getInputControles() {
        var list = [];
        //var d = $.Deferred();
        Office.context.document.customXmlParts.getByNamespaceAsync('http://schema.binaryrepublik.com/2018/inputFieldsDetails', function (asyncResult) {
            if (asyncResult.value.length > 0) {
                $("#thisInputNotification").css('display', 'none');
                $("#thisInputNotification").html('');
                GetData(asyncResult).done(function (asyncResultData) {
                    var uniqueIds = utility.GetUnique(controlsID);
                    controlsID = [];
                    bindaInputFields(asyncResultData, uniqueIds);
                    getUserPlaceolderValues(asyncResultData);
                }).fail(function (error) {
                    console.log(error.message)
                });
            }
            else {
                $("#thisInputNotification").css('display', 'block');
                $("#thisInputNotification").html('Currently no input controls are created.');
                $("#listOfInputFields").html('');
                $("#ddlInputFied").find('option:not(:first)').remove();
                $("#ddlInputFiedToSetValue").find('option:not(:first)').remove();
                $("#SPListFieldForPlaceHolder").find('option:not(:first)').remove();
                $("#SPListInputFieldValue").find('option:not(:first)').remove();
            }
        });

    }

    function GetData(asyncResult) {
        var def = $.Deferred();
        var list = [];
        var temCount = 1;
        for (var i = 0; i < asyncResult.value.length; i++) {
            Office.context.document.customXmlParts.getByIdAsync(asyncResult.value[i].id, function (result) {
                var xmlPart = result.value;
                controlsID.push(xmlPart.id);
                xmlPart.getXmlAsync({}, function (eventArgs) {
                    var inputControlXML = $.parseXML(eventArgs.value);
                    var xmlDoc = inputControlXML;
                    var inputControlDetails = $(inputControlXML).find("parameters")[0].textContent;
                    list.push(inputControlDetails);
                    if (temCount >= asyncResult.value.length) {
                        def.resolve(list);
                    }
                    else {
                        temCount++;
                    }
                });
            });
        }
        return def.promise();
    }

    function getUserPlaceolderValues(data) {
        for (var i = 0; i < data.length; i++) {
            var item = JSON.parse(data[i]);
            if (item.checkCurrentUser != undefined && item.checkCurrentUser != '' && (item.checkCurrentUser)) {
                setValuesOfUserPlaceholders(item, '', true);
            }
        }
    }

    function bindaInputFields(data, ids) {
        var listOfInputFields = "";
        $("#thisInputNotification").html('');
        var bindSPFiedsValue = "<option value='0'>Select</option>";
        for (var i = 0; i < data.length; i++) {
            var item = JSON.parse(data[i]);
            if (item) {
                listOfInputFields += generateInputFieldsLI(item, ids[i]);
                bindSPFiedsValue += "<option id='" + ids[i] + "' value='" + JSON.stringify(item) + "'>" + item.InuptControleName + "</option>";
            }
        }
        $("#listOfInputFields").html(listOfInputFields);
        $("#ddlInputFied").html(bindSPFiedsValue);
    }

    function generateInputFieldsLI(item, id) {
        var controlName = (item.InuptControleName).length > 14 ? (item.InuptControleName).substring(0, 13) + '...' : item.InuptControleName;        
        var liElement = '<li  title="' + item.InuptControleName + '" id="' + id + '" TenatName="' + item.TenatName + '" WebRelativeURL="' + item.WebRelativeURL + '" ListDisplayName="' + item.ListDisplayName + '" FieldInternalName="' + item.FieldInternalName + '" class="list-group-item"><div class="inline icons">' + controlName + '</div><div class="inline icons" style="float: right;"><i data-target="#dltInputCtrl" class="fa fa-close"></i></div></li>';        
        return liElement;
    }

    // Function that writes to a div with id='message' on the page.
    //function write(message) {
    //    document.getElementById('pmsg').innerText += message;
    //}


    this.clearDropDowns = function (onChangeValue) {
        clearDropDowns(onChangeValue);
    }

    //To clear All dropdowns based on selection
    function clearDropDowns(onChangeValue) {
        switch (onChangeValue) {
            case "siteCollection":
                $('#SPListFields').find('option:not(:first)').remove();
                $('#SPLists').find('option:not(:first)').remove();
                $('#SPSubsites').find('option:not(:first)').remove();
                break;
            case "site":
                $('#SPLists').find('option:not(:first)').remove();
                $('#SPListFields').find('option:not(:first)').remove();
                break;
            case "list":
                $('#SPListFields').find('option:not(:first)').remove();
                break;
        }
    }

    function getInfromationFromTage(tagName) {
        var clmnInformation = tagName.split('¤');
        return clmnInformation;
    }

    function getTagFromLI(data) {
        return $(data).parent().parent("li").attr("tag");
    }

}