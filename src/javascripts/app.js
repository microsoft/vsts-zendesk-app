import BaseApp from "base_app";
import helpers from "helpers";
import Base64 from "base64";
window.helpers = helpers;
window.Base64 = Base64;
import _ from "lodash";

String.prototype.fmt = function() {
    return helpers.fmt.apply(this, [this, ...arguments]);
};

const invokeMethods = /^(hide|show|preloadPane|popover|enableSave|disableSave|setIconState|notify)$/;
const wrapZafClient = async (client, apiPath, ...rest) => {
    let method = "get";
    const isInvoke = invokeMethods.test(apiPath);
    if (!isInvoke && rest.length) {
        if (/^(ticket|user|organization)(Fields|\.customField)$/.test(apiPath)) {
            apiPath = `${apiPath}:${rest.shift()}`;
        }
        if (rest.length) method = "set";
    } else if (isInvoke) {
        method = "invoke";
    }
    try {
        let result, errors;
        // Use destructuring to get the value from path on result object
        ({ [apiPath]: result, errors } = await client[method](apiPath, ...rest));
        if (errors && Object.keys(errors).length) {
            console.warn(`Some errors were encountered in request ${apiPath}`, errors);
        }
        return result;
    } catch ({ message }) {
        console.error(message);
    }
};

const objGet = function(obj, path) {
    const npath = path.replace(/\]/g, "");
    const pieces = npath.split("[");
    let result = obj;
    for (const piece of pieces) {
        result = result[piece];
    }
    return result;
};

const tempArgKey = "VSTS_ZENDESK_TEMP_ARG";
const getMessageArg = function() {
    const argVal = window.localStorage.getItem(tempArgKey);
    window.localStorage.removeItem(tempArgKey);
    return JSON.parse(argVal);
};
const setMessageArg = function(data) {
    window.localStorage.setItem(tempArgKey, JSON.stringify(data));
};

const sharedDataKey = "VSTS_ZENDESK_SHARED_DATA";
const replaceVm = function(replacement) {
    window.localStorage.setItem(sharedDataKey, JSON.stringify(replacement));
};
const mergeVm = function(toMerge) {
    const storedVm = JSON.parse(window.localStorage.getItem(sharedDataKey));
    _.merge(storedVm, toMerge);
    window.localStorage.setItem(sharedDataKey, JSON.stringify(storedVm));
};
const assignVm = function(toAssign) {
    const storedVm = JSON.parse(window.localStorage.getItem(sharedDataKey));
    Object.assign(storedVm, toAssign);
    window.localStorage.setItem(sharedDataKey, JSON.stringify(storedVm));
};
const getVm = function(path) {
    const storedVm = JSON.parse(window.localStorage.getItem(sharedDataKey));
    if (path) {
        return objGet(storedVm, path);
    }
    return storedVm;
};

const App = (function() {
    "use strict"; //#region Constants

    var INSTALLATION_ID = 0,
        //For dev purposes, when using Zat, set this to your current installation id
        VSO_URL_FORMAT = "https://%@.visualstudio.com/DefaultCollection",
        VSO_API_DEFAULT_VERSION = "1.0",
        VSO_API_RESOURCE_VERSION = {},
        TAG_PREFIX = "vso_wi_",
        DEFAULT_FIELD_SETTINGS = JSON.stringify({
            "System.WorkItemType": {
                summary: true,
                details: true,
            },
            "System.Title": {
                summary: false,
                details: true,
            },
            "System.Description": {
                summary: true,
                details: true,
            },
        }),
        VSO_ZENDESK_LINK_TO_TICKET_PREFIX = "ZendeskLinkTo_Ticket_",
        VSO_ZENDESK_LINK_TO_TICKET_ATTACHMENT_PREFIX = "ZendeskLinkTo_Attachment_Ticket_",
        VSO_WI_TYPES_WHITE_LISTS = ["Bug", "Task", "Improvement", "User Story", "Feature"],
        VSO_PROJECTS_PAGE_SIZE = 100; //#endregion

    return {
        defaultState: "loading",

        //#region Events Declaration
        events: {
            // App
            "app.activated": "onAppActivated",
            // Requests
            "getVsoProjects.done": "onGetVsoProjectsDone",
            "getVsoFields.done": "onGetVsoFieldsDone",
            //New workitem dialog
            "click .newWorkItem": "onNewWorkItemClick",
            //Admin side pane
            "click .cog": "onCogClick",
            "click .closeAdmin": "onCloseAdminClick",
            "change .settings .summary, .settings .details": "onSettingChange",
            //Details dialog
            "click .showDetails": "onShowDetailsClick",
            //Link work item dialog
            "click .link": "onLinkClick",
            //Unlink click
            "click .unlink": "onUnlinkClick",
            //Notify dialog
            "click .notify": "onNotifyClick",
            //Refresh work items
            "click .refreshWorkItemsLink": "onRefreshWorkItemClick",
            //Login
            "click .user,.user-link": "onUserIconClick",
            "click .closeLogin": "onCloseLoginClick",
            "click .login-button": "onLoginClick",
        },
        //#endregion
        //#region Requests
        requests: {
            getComments: async function() {
                const ticket = await wrapZafClient(this.zafClient, "ticket");

                return {
                    url: helpers.fmt("/api/v2/tickets/%@/comments.json", ticket.id),
                    type: "GET",
                };
            },
            addTagToTicket: async function(tag) {
                const ticket = await wrapZafClient(this.zafClient, "ticket");
                return {
                    url: helpers.fmt("/api/v2/tickets/%@/tags.json", ticket.id),
                    type: "PUT",
                    dataType: "json",
                    data: {
                        tags: [tag],
                    },
                };
            },
            removeTagFromTicket: async function(tag) {
                const ticket = await wrapZafClient(this.zafClient, "ticket");
                return {
                    url: helpers.fmt("/api/v2/tickets/%@/tags.json", ticket.id),
                    type: "DELETE",
                    dataType: "json",
                    data: {
                        tags: [tag],
                    },
                };
            },
            addPrivateCommentToTicket: async function(text) {
                const ticket = await wrapZafClient(this.zafClient, "ticket");
                return {
                    url: helpers.fmt("/api/v2/tickets/%@.json", ticket.id),
                    type: "PUT",
                    dataType: "json",
                    data: {
                        ticket: {
                            comment: {
                                public: false,
                                body: text,
                            },
                        },
                    },
                };
            },
            saveSettings: function(data) {
                return {
                    type: "PUT",
                    url: helpers.fmt("/api/v2/apps/installations/%@.json", this.installationId() || INSTALLATION_ID),
                    dataType: "json",
                    data: {
                        enabled: true,
                        settings: data,
                    },
                };
            },
            getVsoProjects: function(skip) {
                return this.vsoRequest("/_apis/projects", {
                    $top: VSO_PROJECTS_PAGE_SIZE,
                    $skip: skip || 0,
                });
            },
            getVsoProjectWorkItemTypes: function(projectId) {
                return this.vsoRequest(helpers.fmt("/%@/_apis/wit/workitemtypes", projectId));
            },
            getVsoProjectAreas: function(projectId) {
                return this.vsoRequest(helpers.fmt("/%@/_apis/wit/classificationnodes/areas", projectId), {
                    $depth: 9999,
                });
            },
            getVsoProjectWorkItemQueries: function(projectName) {
                return this.vsoRequest(helpers.fmt("/%@/_apis/wit/queries", projectName), {
                    $depth: 2,
                });
            },
            getVsoFields: function() {
                return this.vsoRequest("/_apis/wit/fields");
            },
            getVsoWorkItems: function(ids) {
                return this.vsoRequest("/_apis/wit/workItems", {
                    ids: ids,
                    $expand: "relations",
                });
            },
            getVsoWorkItem: function(workItemId) {
                return this.vsoRequest(helpers.fmt("/_apis/wit/workItems/%@", workItemId), {
                    $expand: "relations",
                });
            },
            getVsoWorkItemQueryResult: function(projectName, queryId) {
                return this.vsoRequest(helpers.fmt("/%@/_apis/wit/wiql/%@", projectName, queryId));
            },
            createVsoWorkItem: function(projectId, witName, data) {
                return this.vsoRequest(helpers.fmt("/%@/_apis/wit/workitems/$%@", projectId, witName), undefined, {
                    type: "PUT",
                    contentType: "application/json-patch+json",
                    data: JSON.stringify(data),
                    headers: {
                        "X-HTTP-Method-Override": "PATCH",
                    },
                });
            },
            updateVsoWorkItem: function(workItemId, data) {
                return this.vsoRequest(helpers.fmt("/_apis/wit/workItems/%@", workItemId), undefined, {
                    type: "PUT",
                    contentType: "application/json-patch+json",
                    data: JSON.stringify(data),
                    headers: {
                        "X-HTTP-Method-Override": "PATCH",
                    },
                });
            },
            updateMultipleVsoWorkItem: function(data) {
                return this.vsoRequest("/_apis/wit/workItems", undefined, {
                    type: "PUT",
                    contentType: "application/json",
                    data: JSON.stringify(data),
                    headers: {
                        "X-HTTP-Method-Override": "PATCH",
                    },
                });
            },
        },

        action_ajax: async endpoint => {
            try {
                return await window.appThis.ajax.apply(window.appThis, endpoint);
            } catch (e) {
                // e is jqXHR
                throw new Error(e.responseJSON.message);
            }
        },

        action_linkTicket: async workItemId => {
            await window.appThis.linkTicket(workItemId);
        },

        action_unlinkTicket: async workItemId => {
            await window.appThis.unlinkTicket(workItemId);
        },

        action_getLinkedWorkItemIds: async () => {
            return await window.appThis.getLinkedWorkItemIds();
        },

        action_setDirty: function() {
            window.appThis.isDirty = true;
        },

        action_fetchLinkedVsoWorkItems: async () => {
            return await window.appThis.fetchLinkedVsoWorkItems();
        },

        //When we create the modal we send the current client's guid as a hash parameter to the modal by adding it to the url

        createModal: async function(context, template) {
            const parentGuid = context.instanceGuid;
            const options = {
                location: "modal",
                url: "assets/modal.html#parentGuid=" + parentGuid,
            };

            const registeredModalClient = await new Promise(async resolve => {
                const modalContext = await this.zafClient.invoke("instances.create", options);
                const newModalGuid = modalContext["instances.create"][0].instanceGuid;
                const modalClient = this.zafClient.instance(newModalGuid);
                this._nextModalRegistrationResolver = resolve.bind(null, modalClient);
                return registeredModalClient;
            });
            this._currentModalClient = registeredModalClient;

            if (template) {
                registeredModalClient.trigger("load_template", template);
            }

            registeredModalClient.on("modal.close", () => {
                this.onModalClosed();
            });

            return registeredModalClient;
        },

        onModalClosed: async function() {
            if (this.isDirty) {
                this.getLinkedVsoWorkItems();
                this.isDirty = false;
            }
        },

        execQueryOnModal: async function(taskName) {
            console.log(`Executing ${taskName} on modal.`);
            setMessageArg(taskName);
            this._currentModalClient.trigger("execute.query");
            const response = await new Promise(resolve => {
                this._nextModalQueryResponseResolver = resolve;
            });
            return response;
        },

        execActionOnModal: function(actionName) {
            console.log(`Executing ${actionName} action on modal.`);
            setMessageArg(actionName);
            this._currentModalClient.trigger("execute.action");
        },

        onModalRegistered: function(modalGuid) {
            // resolve the promise that was set up when the instance was created.
            this._nextModalRegistrationResolver();
        },

        onModalResponse: function(response) {
            this._nextModalQueryResponseResolver(response);
        },

        onGetVsoProjectsDone: function(projects) {
            assignVm({
                projects: _.sortBy(
                    getVm("projects").concat(
                        _.map(projects.value, function(project) {
                            return {
                                id: project.id,
                                name: project.name,
                                workItemTypes: [],
                            };
                        }),
                    ),
                    function(project) {
                        return project.name.toLowerCase();
                    },
                ),
            });
        },
        onGetVsoFieldsDone: function(data) {
            assignVm({
                fields: _.map(data.value, function(field) {
                    return {
                        refName: field.referenceName,
                        name: field.name,
                        type: field.type,
                    };
                }),
            });
        },

        fetchLinkedVsoWorkItems: async function() {
            const vsoLinkedIds = await this.getLinkedWorkItemIds();
            if (!vsoLinkedIds || vsoLinkedIds.length === 0) {
                return [];
            }
            try {
                return await Promise.all(vsoLinkedIds.map(i => this.ajax("getVsoWorkItem", i)));
            } catch (e) {
                await this.displayMain(e.message);
            }
        },

        getLinkedVsoWorkItems: async function(func) {
            var vsoLinkedIds = await this.getLinkedWorkItemIds();

            var finish = async function(workItems) {
                if (func && _.isFunction(func)) {
                    func(workItems);
                } else {
                    this.onGetLinkedVsoWorkItemsDone(workItems);
                    await this.displayMain();
                }
            }.bind(this);

            if (!vsoLinkedIds || vsoLinkedIds.length === 0) {
                finish([]);
                return;
            }

            //make a call for each linked wi to get the data we need (web URL is not returned from the getVsoWorkItems)
            try {
                const requests = vsoLinkedIds.map(i => this.ajax("getVsoWorkItem", i));
                const linkedWorkItems = await Promise.all(requests);
                finish(linkedWorkItems);
            } catch (e) {
                await this.displayMain(e.message);
            }

            // this.ajax('getVsoWorkItems', vsoLinkedIds.join(','))
            //    .then(function (data) { finish(data.value); })
            //    .catch(function (jqXHR) { this.displayMain(this.getAjaxErrorMessage(jqXHR)); }.bind(this));
        },
        onGetLinkedVsoWorkItemsDone: function(data) {
            this.vmLocal.workItems = data;

            _.each(
                this.vmLocal.workItems,
                function(workItem) {
                    workItem.title = helpers.fmt("%@: %@", workItem.id, this.getWorkItemFieldValue(workItem, "System.Title"));
                }.bind(this),
            );

            this.drawWorkItems();
        },
        //#endregion
        //#region Events Implementation
        // App
        onAppActivated: async function(data) {
            window.appThis = this;

            //Global view model shared by all instances
            replaceVm({
                accountUrl: null,
                projects: [],
                fields: [],
                fieldSettings: {},
                userProfile: {},
                isAppLoadedOk: false,
                settings: { "vso_wi_description_template": this.setting("vso_wi_description_template") }
            });

            if (data.firstLoad) {
                //Modal registration
                this.zafClient.on("registered.done", (...params) => {
                    this.onModalRegistered(getMessageArg());
                });
                this.zafClient.on("execute.response", () => {
                    this.onModalResponse(getMessageArg());
                });
                this.zafClient.on(
                    "execute.query",
                    async function() {
                        let args = getMessageArg();
                        if (typeof args === "string") {
                            args = [args];
                        }
                        let result;
                        try {
                            result = await this["action_" + args[0]].call(this, args.slice(1));
                        } catch (e) {
                            result = { err: e.message };
                        }
                        setMessageArg(result);
                        this._currentModalClient.trigger("execute.response");
                    }.bind(this),
                );

                //Check if everything is ok to continue
                if (!this.setting("vso_account")) {
                    return this.switchTo("finish_setup");
                } //set account url

                assignVm({ accountUrl: this.buildAccountUrl() });

                if (!this.store("auth_token_for_" + this.setting("vso_account"))) {
                    return this.switchTo("login");
                } //Private instance view model

                this.vmLocal = {
                    workItems: [],
                };

                if (!getVm("isAppLoadedOk")) {
                    //Initialize global data
                    try {
                        assignVm({ fieldSettings: JSON.parse(this.setting("vso_field_settings") || DEFAULT_FIELD_SETTINGS) });
                    } catch (ex) {
                        this.zafClient.invoke("notify", this.I18n.t("errorReadingFieldSettings"), "alert");
                        assignVm({ fieldSettings: JSON.parse(DEFAULT_FIELD_SETTINGS) });
                    } // Function to get all VSTS projects paginated if needed

                    var getAllVsoProjects = function() {
                        return this.promise(
                            function(done, fail) {
                                var getPage = function(page) {
                                    var skip = page * VSO_PROJECTS_PAGE_SIZE;
                                    this.ajax("getVsoProjects", skip)
                                        .then(function(data) {
                                            // If the page is full, get a new page
                                            if (data.count === VSO_PROJECTS_PAGE_SIZE) {
                                                getPage(page + 1);
                                            } else {
                                                done();
                                            }
                                        })
                                        .catch(function(xhr, status, err) {
                                            fail(xhr, status, err);
                                        });
                                }.bind(this); // Get First page

                                getPage(0);
                            }.bind(this),
                        );
                    }.bind(this);

                    this.when(getAllVsoProjects(), this.ajax("getVsoFields"))
                        .then(
                            async function() {
                                assignVm({ isAppLoadedOk: true });
                                await this.getLinkedVsoWorkItems();
                            }.bind(this),
                        )
                        .fail(
                            function(jqXHR, textStatus, err) {
                                this.switchTo("error_loading_app", {
                                    invalidAccount: jqXHR.status === 404,
                                    accountName: this.setting("vso_account"),
                                });
                            }.bind(this),
                        );
                } else {
                    await this.getLinkedVsoWorkItems();
                }
            }
        },
        resize: function() {
            this.zafClient.invoke("resize", { height: this.$("html").outerHeight(true) + 15, width: "100%" });
        },

        // UI
        onNewWorkItemClick: async function() {
            assignVm({ temp: { ticket: await wrapZafClient(this.zafClient, "ticket") } });
            const modalClient = await this.createModal(this._context, "newWorkItemModal");
            modalClient.on("modal.close", () => {
                this.onModalClosed();
            });
            this.execActionOnModal("initNewWorkItem");
        },
        onNewVsoProjectChange: function() {
            var $modal = this.$(".newWorkItemModal");
            var projId = $modal.find(".project").val();
            this.showSpinnerInModal($modal);
            this.loadProjectMetadata(projId)
                .then(
                    function() {
                        this.drawAreasList($modal.find(".area"), projId);
                        this.drawTypesList($modal.find(".type"), projId);
                        $modal.find(".type").change();
                        this.hideSpinnerInModal($modal);
                    }.bind(this),
                )
                .catch(
                    function(jqXHR) {
                        this.showErrorInModal($modal, this.getAjaxErrorMessage(jqXHR));
                    }.bind(this),
                );
        },
        onNewVsoWorkItemTypeChange: function() {
            var $modal = this.$(".newWorkItemModal");
            var project = this.getProjectById($modal.find(".project").val());
            var workItemType = this.getWorkItemTypeByName(project, $modal.find(".type").val()); //Check if we have severity

            if (this.hasFieldDefined(workItemType, "Microsoft.VSTS.Common.Severity")) {
                $modal.find(".severityInput").show();
            } else {
                $modal.find(".severityInput").hide();
            }
        },
        onNewCopyDescriptionClick: async function(event) {
            const _ticket2 = await wrapZafClient(this.zafClient, "ticket");

            event.preventDefault();
            this.$(".newWorkItemModal .description").val(_ticket2.description);
        },
        onCogClick: function() {
            this.switchTo("admin");
            this.drawSettings();
            this.resize();
        },
        onCloseAdminClick: async function() {
            await this.displayMain();
        },
        onSettingChange: function() {
            var self = this;
            var fieldSettings = {};
            this.$("tr").each(function() {
                var line = self.$(this);
                var fieldName = line.attr("data-refName");
                if (!fieldName) return true; //continue

                var inSummary = line.find(".summary").is(":checked");
                var inDetails = line.find(".details").is(":checked");

                if (inSummary || inDetails) {
                    fieldSettings[fieldName] = {
                        summary: inSummary,
                        details: inDetails,
                    };
                } else if (fieldSettings[fieldName]) {
                    delete fieldName[fieldName];
                }
            });
            assignVm({ fieldSettings: fieldSettings });

            this.ajax("saveSettings", {
                vso_field_settings: JSON.stringify(fieldSettings),
            }).then(
                function() {
                    this.zafClient.invoke("notify", this.I18n.t("admin.settingsSaved"));
                }.bind(this),
            );
        },
        onShowDetailsClick: async function(event) {
            const modalClient = await this.createModal(this._context, "detailsModal");
            var id = this.$(event.target)
                .closest(".workItem")
                .attr("data-id");
            this.execActionOnModal(["initWorkItemDetails", this.getWorkItemById(id)]);
        },
        onLinkClick: async function() {
            assignVm({ temp: { ticket: await wrapZafClient(this.zafClient, "ticket") } });
            const modalClient = await this.createModal(this._context, "linkModal");
            modalClient.on("modal.close", () => {
                this.onModalClosed();
            });
            this.execActionOnModal(["initLinkWorkItem"]);
        },
        onUnlinkClick: async function(event) {
            assignVm({ temp: { ticket: await wrapZafClient(this.zafClient, "ticket") } });
            var id = this.$(event.target)
                .closest(".workItem")
                .attr("data-id");
            var workItem = this.getWorkItemById(id);
            const modalClient = await this.createModal(this._context, "unlinkModal");
            modalClient.on("modal.close", () => {
                this.onModalClosed();
            });
            this.execActionOnModal(["initUnlinkWorkItem", workItem]);
        },
        onNotifyClick: async function() {
            const modalClient = await this.createModal(this._context, "notifyModal");
            modalClient.on("modal.close", () => {
                this.onModalClosed();
            });
            this.execActionOnModal("initNotify");
        },
        onRefreshWorkItemClick: async function(event) {
            event.preventDefault();
            this.$(".workItemsError").hide();
            this.switchTo("loading");
            await this.getLinkedVsoWorkItems();
        },
        onLoginClick: async function(event) {
            event.preventDefault();
            var vso_username = this.$(".vso_username").val();
            var vso_password = this.$(".vso_password").val();

            if (!vso_password) {
                this.$(".login-form")
                    .find(".errors")
                    .text(this.I18n.t("login.errRequiredFields"))
                    .show();
                return;
            }

            this.authString(vso_username, vso_password);
            this.zafClient.invoke("notify", this.I18n.t("notify.credentialsSaved"));
            this.switchTo("loading");

            if (!getVm("isAppLoadedOk")) {
                await this.onAppActivated({
                    firstLoad: true,
                });
            } else {
                await this.getLinkedVsoWorkItems();
            }
        },
        onCloseLoginClick: async function() {
            await this.displayMain();
        },
        onUserIconClick: function() {
            this.switchTo("login");
        },
        //#endregion
        //#region Drawing
        displayMain: async function(err) {
            if (getVm("isAppLoadedOk")) {
                this.$(".cog").toggle(await this.isAdmin());
                this.switchTo("main");

                if (!err) {
                    this.drawWorkItems();
                } else {
                    this.$(".workItemsError").show();
                }
            } else {
                this.$(".cog").toggle(false);
                this.switchTo("error_loading_app");
            }
        },
        drawWorkItems: function(data) {
            var workItems = _.map(
                data || this.vmLocal.workItems,
                function(workItem) {
                    var tmp = this.attachRestrictedFieldsToWorkItem(workItem, "summary");
                    return tmp;
                }.bind(this),
            );

            this.$(".workItems").html(
                this.renderTemplate("workItems", {
                    workItems: workItems,
                }),
            );
            this.$(".buttons .notify").prop("disabled", !workItems.length);
            this.resize();
        },
        drawTypesList: function(select, projectId) {
            var project = this.getProjectById(projectId);
            select.html(
                this.renderTemplate("types", {
                    types: project.workItemTypes,
                }),
            );
        },
        drawAreasList: function(select, projectId) {
            var project = this.getProjectById(projectId);
            select.html(
                this.renderTemplate("areas", {
                    areas: project.areas,
                }),
            );
        },
        drawSettings: function() {
            const fields = getVm("fields");
            var settings = _.sortBy(
                _.map(
                    fields,
                    function(field) {
                        var current = getVm("fieldSettings[" + field.refName + "]");

                        if (current) {
                            field = _.extend(field, current);
                        }

                        return field;
                    }.bind(this),
                ),
                function(f) {
                    return f.name;
                },
            );
            assignVm({ fields: fields });

            var html = this.renderTemplate("settings", {
                settings: settings,
            });
            this.$(".content").html(html);
        },
        showSpinnerInModal: function($modal) {
            if ($modal.find(".modal-body form")) {
                $modal.find(".modal-body form").hide();
            }

            if ($modal.find(".modal-body .loading")) {
                $modal.find(".modal-body .loading").show();
            }

            if ($modal.find(".modal-footer button")) {
                $modal.find(".modal-footer button").attr("disabled", "disabled");
            }
        },
        hideSpinnerInModal: function($modal) {
            if ($modal.find(".modal-body form")) {
                $modal.find(".modal-body form").show();
            }

            if ($modal.find(".modal-body .loading")) {
                $modal.find(".modal-body .loading").hide();
            }

            if ($modal.find(".modal-footer button")) {
                $modal.find(".modal-footer button").prop("disabled", false);
            }
        },
        showErrorInModal: function($modal, err) {
            this.hideSpinnerInModal($modal);

            if ($modal.find(".modal-body .errors")) {
                $modal
                    .find(".modal-body .errors")
                    .text(err)
                    .show();
            }
        },
        closeModal: function($modal) {
            $modal.find("#loading").hide();
            $modal
                .modal("hide")
                .find(".modal-footer button")
                .attr("disabled", "");
        },
        fillComboWithProjects: function(el) {
            el.html(
                _.reduce(
                    getVm("projects"),
                    function(options, project) {
                        return "%@<option value='%@'>%@</option>".fmt(options, project.id, project.name);
                    },
                    "",
                ),
            );
        },
        //#endregion
        //#region Helpers
        isAdmin: async function() {
            const _currentUser2 = await wrapZafClient(this.zafClient, "currentUser");

            return _currentUser2.role === "admin";
        },
        vsoUrl: function(url, parameters) {
            url = url[0] === "/" ? url.slice(1) : url;
            var full = [getVm("accountUrl"), url].join("/");

            if (parameters) {
                full +=
                    "?" +
                    _.map(parameters, function(value, key) {
                        return [key, value].join("=");
                    }).join("&");
            }

            return full;
        },
        authString: function(vso_username, vso_password) {
            if (vso_password) {
                var b64 = Base64.encode([vso_username, vso_password].join(":"));
                this.store("auth_token_for_" + this.setting("vso_account"), b64);
            }

            return helpers.fmt("Basic %@", this.store("auth_token_for_" + this.setting("vso_account")));
        },
        vsoRequest: function(url, parameters, options) {
            var requestOptions = _.extend(
                {
                    url: this.vsoUrl(url, parameters),
                    dataType: "json",
                },
                options,
            );

            var fixedHeaders = {
                Authorization: this.authString(),
                Accept: helpers.fmt("application/json;api-version=%@", this.getVsoResourceVersion(url)),
            };
            requestOptions.headers = _.extend(fixedHeaders, options ? options.headers : {});
            return requestOptions;
        },
        getVsoResourceVersion: function(url) {
            var resource = url.split("/_apis/")[1].split("/")[0];
            return VSO_API_RESOURCE_VERSION[resource] || VSO_API_DEFAULT_VERSION;
        },
        attachRestrictedFieldsToWorkItem: function(workItem, type) {
            const fieldSettings = getVm("fieldSettings");
            var fields = _.compact(
                _.map(
                    fieldSettings,
                    function(value, key) {
                        if (value[type]) {
                            if (_.has(workItem.fields, key)) {
                                return {
                                    refName: key,
                                    name: _.find(getVm("fields"), function(f) {
                                        return f.refName == key;
                                    }).name,
                                    value: workItem.fields[key],
                                    isHtml: this.isHtmlContentField(key),
                                };
                            }
                        }
                    }.bind(this),
                ),
            );
            assignVm({ fieldSettings: fieldSettings });

            return _.extend(workItem, {
                restricted_fields: fields,
            });
        },
        getWorkItemById: function(id) {
            return _.find(this.vmLocal.workItems, function(workItem) {
                return workItem.id == id;
            });
        },
        getProjectById: function(id) {
            const projects = getVm("projects");
            return _.find(projects, function(proj) {
                return proj.id == id;
            });
        },
        getWorkItemTypeByName: function(project, name) {
            return _.find(project.workItemTypes, function(wit) {
                return wit.name == name;
            });
        },
        getFieldByFieldRefName: function(fieldRefName) {
            const fields = getVm("fields");
            return _.find(fields, function(f) {
                return f.refName == fieldRefName;
            });
        },
        getWorkItemFieldValue: function(workItem, fieldRefName) {
            var field = workItem.fields[fieldRefName];
            return field || "";
        },
        hasFieldDefined: function(workItemType, fieldRefName) {
            return _.some(workItemType.fieldInstances, function(fieldInstance) {
                return fieldInstance.referenceName === fieldRefName;
            });
        },
        linkTicket: async function(workItemId) {
            var linkVsoTag = TAG_PREFIX + workItemId;
            await this.zafClient.invoke("ticket.tags.add", linkVsoTag);
            this.ajax("addTagToTicket", linkVsoTag);
        },
        unlinkTicket: async function(workItemId) {
            var linkVsoTag = TAG_PREFIX + workItemId;
            await this.zafClient.invoke("ticket.tags.remove", linkVsoTag);
            this.ajax("removeTagFromTicket", linkVsoTag);
        },
        buildTicketLinkUrl: async function() {
            const _currentAccount = await wrapZafClient(this.zafClient, "currentAccount"),
                _ticket6 = await wrapZafClient(this.zafClient, "ticket");

            return helpers.fmt("https://%@.zendesk.com/agent/#/tickets/%@", _currentAccount.subdomain, _ticket6.id);
        },
        getLinkedWorkItemIds: async function() {
            const tags = (await this.zafClient.get("ticket.tags"))["ticket.tags"];

            return _.compact(
                tags.map(t => {
                    var p = t.indexOf(TAG_PREFIX);

                    if (p === 0) {
                        return t.slice(TAG_PREFIX.length);
                    }
                }),
            );
        },
        isAlreadyLinkedToWorkItem: async function(id) {
            return _.contains(await this.getLinkedWorkItemIds(), id);
        },
        loadProjectMetadata: function(projectId) {
            var project = this.getProjectById(projectId);

            if (project.metadataLoaded === true) {
                return this.promise(function(done) {
                    done();
                });
            }

            var loadWorkItemTypes = this.ajax("getVsoProjectWorkItemTypes", project.id).then(
                function(data) {
                    project.workItemTypes = this.restrictToAllowedWorkItems(data.value);
                }.bind(this),
            );
            var loadAreas = this.ajax("getVsoProjectAreas", project.id).then(
                function(rootArea) {
                    var areas = []; // Flatten areas to format \Area 1\Area 1.1

                    var visitArea = function(area, currentPath) {
                        currentPath = currentPath ? currentPath + "\\" : "";
                        currentPath = currentPath + area.name;
                        areas.push({
                            id: area.id,
                            name: currentPath,
                        });

                        if (area.children && area.children.length > 0) {
                            _.forEach(area.children, function(child) {
                                visitArea(child, currentPath);
                            });
                        }
                    };

                    visitArea(rootArea);
                    project.areas = _.sortBy(areas, function(area) {
                        return area.name;
                    });
                }.bind(this),
            );
            return this.when(loadWorkItemTypes, loadAreas).then(function() {
                project.metadataLoaded = true;
            });
        },
        loadProjectWorkItemQueries: function(projectId, reload) {
            var project = this.getProjectById(projectId);

            if (project.queries && !reload) {
                return this.promise(function(done) {
                    done();
                });
            } //Let's load project queries

            return this.ajax("getVsoProjectWorkItemQueries", project.name).then(
                function(data) {
                    project.queries = data.value;
                }.bind(this),
            );
        },
        restrictToAllowedWorkItems: function(wits) {
            return _.filter(wits, function(wit) {
                return _.contains(VSO_WI_TYPES_WHITE_LISTS, wit.name);
            });
        },
        isHtmlContentField: function(fieldName) {
            var field = this.getFieldByFieldRefName(fieldName);

            if (field && field.type) {
                var fieldType = field.type.toLowerCase();
                return fieldType === "html" || fieldType === "history";
            } else {
                return false;
            }
        },
        getAjaxErrorMessage: function(jqXHR, errMsg) {
            errMsg = errMsg || this.I18n.t("errorAjax"); //Let's try get a friendly message based on some cases

            var serverErrMsg;

            if (jqXHR.responseJSON) {
                serverErrMsg = jqXHR.responseJSON.message || jqXHR.responseJSON.value.Message;
            } else {
                serverErrMsg = jqXHR.responseText.substring(0, 50) + "...";
            }

            var detail = this.I18n.t("errorServer").fmt(jqXHR.status, jqXHR.statusText, serverErrMsg);
            return errMsg + " " + detail;
        },
        buildAccountUrl: function() {
            var baseUrl;
            var setting = this.setting("vso_account");
            var loweredSetting = setting.toLowerCase();

            if (loweredSetting.indexOf("http://") === 0 || loweredSetting.indexOf("https://") === 0) {
                baseUrl = setting;
            } else {
                baseUrl = helpers.fmt(VSO_URL_FORMAT, setting);
            }

            baseUrl = baseUrl[baseUrl.length - 1] === "/" ? baseUrl.slice(0, -1) : baseUrl; //check if collection defined

            if (baseUrl.lastIndexOf("/") <= "https://".length) {
                baseUrl = baseUrl + "/DefaultCollection";
            }

            return baseUrl;
        }, //#endregion
    };
})();
const extendedApp = BaseApp.extend(App);
export default extendedApp;
