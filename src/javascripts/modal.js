import ZAFClient from "zendesk_app_framework_sdk";
import I18n from "i18n";
import View from "view";
import BaseApp from "base_app";
import helpers from "helpers";
window.helpers = helpers;
import _ from "lodash";

String.prototype.fmt = function() {
    return helpers.fmt.apply(this, [this, ...arguments]);
};

// matches polyfill
if (!Element.prototype.matches) {
    Element.prototype.matches =
        Element.prototype.matchesSelector ||
        Element.prototype.mozMatchesSelector ||
        Element.prototype.msMatchesSelector ||
        Element.prototype.oMatchesSelector ||
        Element.prototype.webkitMatchesSelector ||
        function(s) {
            var matches = (this.document || this.ownerDocument).querySelectorAll(s),
                i = matches.length;
            while (--i >= 0 && matches.item(i) !== this);
            return i > -1;
        };
}

// closest polyfill
if (!Element.prototype.closest)
    Element.prototype.closest = function(s) {
        var el = this;
        if (!document.documentElement.contains(el)) return null;
        do {
            if (el.matches(s)) return el;
            el = el.parentElement;
        } while (el !== null);
        return null;
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

// @TODO: for some reason trigger is not passing along additional data. Store it in local storage for now.
const tempArgKey = "VSTS_ZENDESK_TEMP_ARG";
const getMessageArg = function() {
    const argVal = window.localStorage.getItem(tempArgKey);
    window.localStorage.removeItem(tempArgKey);
    let result;
    try {
        result = JSON.parse(argVal);
    } catch (e) {
        result = null;
    }
    return result;
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

// Create a new ZAFClient
var client = ZAFClient.init();

// add an event listener to detect once your app is registered with the framework
client.on("app.registered", function(appData) {
    client.get("currentUser.locale").then(userData => {
        // load translations based on the account's current locale
        I18n.loadTranslations(userData["currentUser.locale"]);
        new ModalApp(client, appData);
    });
});

const ModalApp = BaseApp.extend({
    ajax: function(endpoint) {
        this.execQueryOnSidebar(["ajax", endpoint]);
    },
    onAppActivated: function(data) {
        const parentGuid = /(?:parentGuid=)(.*?)(?:$|&)/.exec(window.location.hash)[1];
        const parentClient = this.zafClient.instance(parentGuid);
        this._parentClient = parentClient;

        setMessageArg(this._context.instanceGuid);
        parentClient.trigger("registered.done");

        this.zafClient.on("load_template", data => {
            console.log("loading template: " + data);
            this.switchTo(data);
        });
        this.zafClient.on("execute.action", () => {
            let args = getMessageArg();
            console.log("executing: " + JSON.stringify(args));
            if (typeof args === "string") {
                args = [args];
            }
            this["action_" + args[0]].apply(this, args.slice(1));
        });
        this.zafClient.on("execute.query", async () => {
            let args = getMessageArg();
            if (typeof args === "string") {
                args = [args];
            }
            setMessageArg(await this["action_" + args[0]](args.slice(1)));
            parentClient.trigger("execute.response");
        });
        this.zafClient.on("execute.response", () => {
            this.onSidebarResponse(getMessageArg());
        });

        this.$("[data-main]").on("click", e => {
            if (e.target.matches("[data-dismiss=modal]")) {
                this.zafClient.invoke("destroy");
            }
        });
    },
    onSidebarResponse: function(response) {
        if (response && response.err) {
            this._nextSidebarQueryResponseResolver.reject({ message: response.err });
        } else {
            this._nextSidebarQueryResponseResolver.resolve(response);
        }
    },
    execQueryOnSidebar: async function(taskName) {
        this.showBusy();
        setMessageArg(taskName);
        this._parentClient.trigger("execute.query");
        let response;
        try {
            response = await new Promise((resolve, reject) => {
                this._nextSidebarQueryResponseResolver = { resolve, reject };
            });
        } finally {
            this.hideBusy();
        }
        return response;
    },

    action_initNotify: async function() {
        const $modal = this.$("[data-main]");
        $modal.find(".modal-body").html(this.renderTemplate("loading"));
        const data = await this.execQueryOnSidebar(["ajax", "getComments"]);
        this.lastComment = data.comments[data.comments.length - 1].body;
        const attachments = _.flatten(
            _.map(data.comments, function(comment) {
                return comment.attachments || [];
            }),
            true,
        );

        $modal.find(".modal-body").html(
            this.renderTemplate("notify", {
                attachments: attachments,
            }),
        );

        $modal.find(".modal-footer button").prop("disabled", false);

        $modal.find(".accept").on("click", e => {
            this.onNotifyAcceptClick(e);
        });
        $modal.find(".copyLastComment").on("click", e => {
            this.onCopyLastCommentClick(e);
        });
        this.resize({ width: "550px", height: "300px" });
    },

    action_initUnlinkWorkItem: function(workItem) {
        var $modal = this.$("[data-main]");
        $modal.find(".modal-body").html(this.renderTemplate("unlink"));
        $modal.find(".modal-footer button").removeAttr("disabled");
        $modal.find(".modal-body .confirm").html(
            this.I18n.t("modals.unlink.text", {
                name: workItem.title,
            }),
        );
        $modal.attr("data-id", workItem.id);
        $modal.find(".accept").on("click", e => {
            this.onUnlinkAcceptClick(e);
        });
        this.resize({ width: "580px", height: "200px" });
    },

    action_initLinkWorkItem: function() {
        const $modal = this.$("[data-main]");
        $modal.find(".modal-footer button").removeAttr("disabled");
        $modal.find(".modal-body").html(this.renderTemplate("link"));
        $modal.find("button.search").show();
        const projectCombo = $modal.find(".project");
        this.fillComboWithProjects(projectCombo);
        this.resize({ width: "580px", height: "280px" });

        $modal.find(".search").on("click", () => {
            this.onLinkSearchClick();
        });
        $modal.find(".project").on("change", () => {
            this.onLinkVsoProjectChange();
        });
        $modal.find(".reloadQueriesBtn").on("click", () => {
            this.onLinkReloadQueriesButtonClick();
        });
        $modal.find(".queryBtn").on("click", () => {
            this.onLinkQueryButtonClick();
        });
        $modal.on("click", e => {
            if (e.target.closest("a.workItemResult") !== null) {
                this.onLinkResultClick(e);
            }
        });
        $modal.find(".accept").on("click", e => {
            this.onLinkAcceptClick(e);
        });
        projectCombo.change();
    },

    action_initWorkItemDetails: async function(workItem) {
        var $modal = this.$("[data-main]");
        $modal.find(".modal-header h3").html(this.I18n.t("modals.details.loading"));
        $modal.find(".modal-body").html(this.renderTemplate("loading"));

        const workItemWithFields = this.attachRestrictedFieldsToWorkItem(workItem, "details");
        $modal.find(".modal-header h3").html(
            this.I18n.t("modals.details.title", {
                name: workItem.title,
            }),
        );
        $modal.find(".modal-body").html(this.renderTemplate("details", workItem));
        this.resize({ width: "770px" });
    },

    action_initNewWorkItem: async function() {
        const $modal = this.$("[data-main]");
        $modal.find(".modal-body").html(this.renderTemplate("loading"));
        const data = await this.execQueryOnSidebar(["ajax", "getComments"]);
        var attachments = _.flatten(
            _.map(data.comments, function(comment) {
                return comment.attachments || [];
            }),
            true,
        ); // Check if we have a template for decription

        var templateDefined = !!this.setting("vso_wi_description_template");
        $modal.find(".modal-body").html(
            this.renderTemplate("new", {
                attachments: attachments,
                templateDefined: templateDefined,
            }),
        );
        $modal.find(".summary").val(getVm("temp[ticket]").subject);
        var projectCombo = $modal.find(".project");
        this.fillComboWithProjects(projectCombo);
        $modal.find(".inputVsoProject").on("change", this.onNewVsoProjectChange.bind(this));
        $modal.find(".copyDescription").on("click", () => {
            $modal.find(".description").val(getVm("temp[ticket]").description);
        });
        $modal.find(".accept").on("click", () => {
            this.onNewWorkItemAcceptClick();
        });
        $modal.find(".copyTemplate").on("click", (e) => {
            this.onNewCopyTemplateClick(e);
        });
        projectCombo.change();
        this.resize({ height: "520px", width: "780px" });
    },

    showBusy: function() {
        this.$("[data-main] .busySpinner").show();
    },

    hideBusy: function() {
        this.$("[data-main] .busySpinner").hide();
    },

    onNewCopyTemplateClick: function(event) {
        event.preventDefault();
        this.$("[data-main] .description").val(getVm("settings[vso_wi_description_template]"));
    },

    onCopyLastCommentClick: function(event) {
        event.preventDefault();
        this.$(".notifyModal")
            .find("textarea")
            .val(this.lastComment);
    },

    onNotifyAcceptClick: async function(event) {
        const _currentUser = await this.zafClient.get("currentUser");

        var $modal = this.$("[data-main]");
        var text = $modal.find("textarea").val();

        if (!text) {
            return this.showErrorInModal($modal, this.I18n.t("modals.notify.errCommentRequired"));
        }

        const workItems = await this.execQueryOnSidebar(["fetchLinkedVsoWorkItems"]);

        // Must do these serially because my execQueryOnSidebar isn't threadsafe
        try {
            for (const workItem of workItems) {
                await this.execQueryOnSidebar([
                    "ajax",
                    "updateVsoWorkItem",
                    workItem.id,
                    [this.buildPatchToAddWorkItemField("System.History", text)],
                ]);
            }

            const ticketMsg = [this.I18n.t("notify.message", { name: _currentUser.name }), text].join("\r\n\r\n");
            await this.execQueryOnSidebar(["ajax", "addPrivateCommentToTicket", ticketMsg]);
            this.zafClient.invoke("notify", this.I18n.t("notify.notification"));
            this.zafClient.invoke("destroy");
        } catch (e) {
            this.showErrorInModal($modal, e.message);
        }
    },

    onUnlinkAcceptClick: async function(event) {
        const ticket = getVm("temp[ticket]");
        event.preventDefault();

        const $modal = this.$("[data-main]");
        const workItemId = $modal.attr("data-id");

        const updateWorkItem = async function(workItem) {
            // Calculate the positions of links to remove
            const posOfLinksToRemove = [];

            _.each(
                workItem.relations,
                function(link, idx) {
                    if (
                        link.rel.toLowerCase() === "hyperlink" &&
                        (link.attributes.name === VSO_ZENDESK_LINK_TO_TICKET_PREFIX + ticket.id ||
                            link.attributes.name === VSO_ZENDESK_LINK_TO_TICKET_ATTACHMENT_PREFIX + ticket.id)
                    ) {
                        posOfLinksToRemove.push(idx - posOfLinksToRemove.length);
                    }
                }.bind(this),
            );

            const finish = async function() {
                await this.unlinkTicket(workItem.id);
                this.zafClient.invoke("notify", this.I18n.t("notify.workItemUnlinked").fmt(workItem.id));

                // close the modal.
                await this.setDirty();
                this.zafClient.invoke("destroy");
            }.bind(this);

            if (posOfLinksToRemove.length === 0) {
                finish();
            } else {
                const operations = [
                    {
                        op: "test",
                        path: "/rev",
                        value: workItem.rev,
                    },
                ].concat(
                    _.map(
                        posOfLinksToRemove,
                        function(pos) {
                            return this.buildPatchToRemoveWorkItemHyperlink(pos);
                        }.bind(this),
                    ),
                );
                try {
                    await this.execQueryOnSidebar(["ajax", "updateVsoWorkItem", workItemId, operations]);
                    finish();
                } catch (e) {
                    this.showErrorInModal($modal, this.I18n.t("modals.unlink.errUnlink") + " - " + e.message);
                }
            }
        }.bind(this); //Get work item to get the last revision and then update

        try {
            const workItem = await this.execQueryOnSidebar(["ajax", "getVsoWorkItem", workItemId]);
            updateWorkItem(workItem);
        } catch (e) {
            this.showErrorInModal($modal, e.message);
        }
    },

    onLinkAcceptClick: async function(event) {
        const ticket = getVm("temp[ticket]");
        const $modal = this.$("[data-main]");
        const workItemId = $modal.find(".inputVsoWorkItemId").val();

        if (!/^([0-9]+)$/.test(workItemId)) {
            return this.showErrorInModal($modal, this.I18n.t("modals.link.errWorkItemIdNaN"));
        }

        if (await this.isAlreadyLinkedToWorkItem(workItemId)) {
            return this.showErrorInModal($modal, this.I18n.t("modals.link.errAlreadyLinked"));
        }

        const updateWorkItem = async function(workItem) {
            //Let's check if there is already a link in the WI returned data
            const currentLink = _.find(
                workItem.relations || [],
                async function(link) {
                    if (
                        link.rel.toLowerCase() === "hyperlink" &&
                        link.attributes.name === VSO_ZENDESK_LINK_TO_TICKET_PREFIX + ticket.id
                    ) {
                        return link;
                    }
                }.bind(this),
            );

            const finish = async function() {
                await this.linkTicket(workItemId);
                this.zafClient.invoke("notify", this.I18n.t("notify.workItemLinked").fmt(workItemId));

                // close the modal.
                await this.setDirty();
                this.zafClient.invoke("destroy");
            }.bind(this);

            if (currentLink) {
                finish();
            } else {
                const addLinkOperation = this.buildPatchToAddWorkItemHyperlink(
                    await this.buildTicketLinkUrl(),
                    VSO_ZENDESK_LINK_TO_TICKET_PREFIX + ticket.id,
                );
                try {
                    await this.execQueryOnSidebar(["ajax", "updateVsoWorkItem", workItemId, [addLinkOperation]]);
                    finish();
                } catch (e) {
                    this.showErrorInModal($modal, this.I18n.t("modals.link.errCannotUpdateWorkItem") + " - " + e.message);
                }
            }
        }.bind(this);

        // Get work item and then update
        try {
            const data = await this.execQueryOnSidebar(["ajax", "getVsoWorkItem", workItemId]);
            await updateWorkItem(data);
        } catch (e) {
            this.showErrorInModal($modal, this.I18n.t("modals.link.errCannotGetWorkItem") + " - " + e.message);
        }
    },

    onLinkResultClick: function(event) {
        event.preventDefault();
        var $modal = this.$("[data-main]");
        var id = this.$(event.target)
            .closest(".workItemResult")
            .attr("data-id");
        $modal.find(".inputVsoWorkItemId").val(id);
        $modal.find(".search-section").hide();
        this.resize();
    },

    onLinkQueryButtonClick: async function() {
        const $modal = this.$("[data-main]");
        const projId = $modal.find(".project").val();
        const queryId = $modal.find(".query").val();

        const drawQueryResults = function(results, countQueryItemsResult) {
            const workItems = _.map(
                results,
                function(workItem) {
                    return {
                        id: workItem.id,
                        type: this.getWorkItemFieldValue(workItem, "System.WorkItemType"),
                        title: this.getWorkItemFieldValue(workItem, "System.Title"),
                    };
                }.bind(this),
            );

            $modal.find(".results").html(
                this.renderTemplate("query_results", {
                    workItems: workItems,
                }),
            );
            $modal.find(".alert-success").html(
                this.I18n.t("queryResults.returnedWorkItems", {
                    count: countQueryItemsResult,
                }),
            );
            this.resize();
        }.bind(this);

        const [done, proj] = this.getProjectById(projId);

        try {
            const data = await this.execQueryOnSidebar(["ajax", "getVsoWorkItemQueryResult", proj.name, queryId]);
            const getWorkItemsIdsFromQueryResult = function(result) {
                if (result.queryType === "oneHop" || result.queryType === "tree") {
                    return _.map(result.workItemRelations, function(rel) {
                        return rel.target.id;
                    });
                } else {
                    return _.pluck(result.workItems, "id");
                }
            };
            const ids = getWorkItemsIdsFromQueryResult(data);
            if (!ids || ids.length === 0) {
                return drawQueryResults([], 0);
            }

            const results = await this.execQueryOnSidebar(["ajax", "getVsoWorkItems", _.first(ids, 200).join(",")]);
            drawQueryResults(results.value, ids.length);
        } catch (e) {
            this.showErrorInModal($modal, this.I18n.t("modals.link.errCannotGetWorkItem. " + e.message));
        }
    },

    onLinkVsoProjectChange: function() {
        this.loadQueriesList();
    },

    onLinkReloadQueriesButtonClick: function() {
        this.loadQueriesList(true);
    },

    onLinkSearchClick: function() {
        const $modal = this.$("[data-main]");
        $modal.find(".search-section").show();
        this.resize({ width: "580px" });
    },

    onNewWorkItemAcceptClick: async function() {
        const ticket = getVm("temp[ticket]");

        const $modal = this.$("[data-main]");

        const [proj, done] = this.getProjectById($modal.find(".project").val());

        if (!proj) {
            return this.showErrorInModal($modal, this.I18n.t("modals.new.errProjRequired"));
        }

        // read area id
        const areaId = $modal.find(".area").val(); //check work item type

        const workItemType = this.getWorkItemTypeByName(proj, $modal.find(".type").val());
        if (!workItemType) {
            return this.showErrorInModal($modal, this.I18n.t("modals.new.errWorkItemTypeRequired"));
        }

        //check summary
        const summary = $modal.find(".summary").val();
        if (!summary) {
            return this.showErrorInModal($modal, this.I18n.t("modals.new.errSummaryRequired"));
        }

        const description = $modal.find(".description").val();
        let operations = [].concat(
            this.buildPatchToAddWorkItemField("System.Title", summary),
            this.buildPatchToAddWorkItemField("System.Description", description),
        );

        if (areaId) {
            operations.push(this.buildPatchToAddWorkItemField("System.AreaId", areaId));
        }

        if (this.hasFieldDefined(workItemType, "Microsoft.VSTS.Common.Severity") && $modal.find(".severity").val()) {
            operations.push(this.buildPatchToAddWorkItemField("Microsoft.VSTS.Common.Severity", $modal.find(".severity").val()));
        }

        if (this.hasFieldDefined(workItemType, "Microsoft.VSTS.TCM.ReproSteps")) {
            operations.push(this.buildPatchToAddWorkItemField("Microsoft.VSTS.TCM.ReproSteps", description));
        } 
        
        //Set tag
        if (this.setting("vso_tag")) {
            operations.push(this.buildPatchToAddWorkItemField("System.Tags", this.setting("vso_tag")));
        }

        //Add hyperlink to ticket url
        operations.push(
            this.buildPatchToAddWorkItemHyperlink(await this.buildTicketLinkUrl(), VSO_ZENDESK_LINK_TO_TICKET_PREFIX + ticket.id),
        );

        //Add hyperlinks to attachments
        const attachments = this.getSelectedAttachments($modal);
        if (attachments.length > 0) {
            operations = operations.concat(this.buildPatchToAddWorkItemAttachments(attachments, ticket));
        }

        try {
            const data = await this.execQueryOnSidebar(["ajax", "createVsoWorkItem", proj.id, workItemType.name, operations]);
            const newWorkItemId = data.id; //sanity check due tfs returning 200 ok  but with exception

            if (newWorkItemId > 0) {
                await this.linkTicket(newWorkItemId); // @TODO
            }

            this.zafClient.invoke("notify", this.I18n.t("notify.workItemCreated").fmt(newWorkItemId));
        } catch (exception) {
            this.showErrorInModal($modal, exception.message);
        }
        done();

        await this.setDirty();

        // close the modal.
        this.zafClient.invoke("destroy");
    },

    isAlreadyLinkedToWorkItem: async function(id) {
        return _.contains(await this.getLinkedWorkItemIds(), id);
    },

    getLinkedWorkItemIds: async function() {
        return await this.execQueryOnSidebar("getLinkedWorkItemIds");
    },

    setDirty: async function() {
        this.execQueryOnSidebar("setDirty");
    },

    loadQueriesList: async function(reload) {
        const $modal = this.$("[data-main]");
        const projId = $modal.find(".project").val();
        try {
            await this.loadProjectWorkItemQueries(projId, reload);
        } catch (e) {
            this.showErrorInModal($modal, e.message);
        }
        this.drawQueriesList($modal.find(".query"), projId);
    },

    loadProjectWorkItemQueries: async function(projectId, reload) {
        const [project, doneWithProj] = this.getProjectById(projectId);

        if (project.queries && !reload) {
            return;
        }

        // Load project queries
        const data = await this.execQueryOnSidebar(["ajax", "getVsoProjectWorkItemQueries", project.name]);
        project.queries = data.value;
        doneWithProj();
        return data;
    },

    drawQueriesList: function(select, projId) {
        const [project, done] = this.getProjectById(projId);

        const drawNode = function(node, prefix) {
            //It's a folder
            if (node.isFolder) {
                return "<optgroup label='%@ %@'>%@</optgroup>".fmt(
                    prefix,
                    node.name,
                    _.reduce(
                        node.children,
                        function(options, childNode, ix) {
                            return "%@%@".fmt(options, drawNode(childNode, prefix + (ix + 1) + "."));
                        },
                        "",
                    ),
                );
            }

            //It's a query
            return "<option value='%@'>%@ %@</option>".fmt(node.id, prefix, node.name);
        }.bind(this);

        select.html(
            _.reduce(
                project.queries,
                function(options, query, ix) {
                    return "%@%@".fmt(options, drawNode(query, "" + (ix + 1) + "."));
                },
                "",
            ),
        );

        done();
    },

    buildTicketLinkUrl: async function() {
        const ticket = getVm("temp[ticket]");
        const _currentAccount = await wrapZafClient(this.zafClient, "currentAccount");

        return helpers.fmt("https://%@.zendesk.com/agent/#/tickets/%@", _currentAccount.subdomain, ticket.id);
    },
    linkTicket: async function(workItemId) {
        await this.execQueryOnSidebar(["linkTicket", workItemId]);
    },
    unlinkTicket: async function(workItemId) {
        await this.execQueryOnSidebar(["unlinkTicket", workItemId]);
    },
    buildPatchToAddWorkItemField: function(fieldName, value) {
        // Check if the field type is html to replace newlines by br
        if (this.isHtmlContentField(fieldName)) {
            value = value.replace(/\n/g, "<br>");
        }

        return {
            op: "add",
            path: helpers.fmt("/fields/%@", fieldName),
            value: value,
        };
    },
    buildPatchToAddWorkItemHyperlink: function(url, name, comment) {
        return {
            op: "add",
            path: "/relations/-",
            value: {
                rel: "Hyperlink",
                url: url,
                attributes: {
                    name: name,
                    comment: comment,
                },
            },
        };
    },
    buildPatchToAddWorkItemAttachments: function(attachments, ticket) {
        return _.map(
            attachments,
            function(att) {
                return this.buildPatchToAddWorkItemHyperlink(
                    att.url,
                    VSO_ZENDESK_LINK_TO_TICKET_ATTACHMENT_PREFIX + ticket.id,
                    att.name,
                );
            }.bind(this),
        );
    },
    getSelectedAttachments: function($modal) {
        var attachments = [];
        $modal.find(".attachments input").each(
            function(ix, el) {
                var $el = this.$(el);
                if ($el.is(":checked")) {
                    attachments.push({
                        url: $el.val(),
                        name: $el.attr("data-file-name"),
                    });
                }
            }.bind(this),
        );
        return attachments;
    },    
    buildPatchToRemoveWorkItemHyperlink: function(pos) {
        return {
            op: "remove",
            path: helpers.fmt("/relations/%@", pos),
        };
    },
    getFieldByFieldRefName: function(fieldRefName) {
        const fields = getVm("fields");
        return _.find(fields, function(f) {
            return f.refName == fieldRefName;
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
    onNewVsoProjectChange: function() {
        var $modal = this.$("[data-main]");
        var projId = $modal.find(".project").val();

        this.showBusy();
        this.loadProjectMetadata(projId)
            .then(
                function() {
                    this.drawAreasList($modal.find(".area"), projId);
                    this.drawTypesList($modal.find(".type"), projId);
                    $modal.find(".type").change();
                    this.hideBusy();
                }.bind(this),
            )
            .catch(
                function(jqXHR) {
                    this.hideBusy();
                    this.showErrorInModal($modal, this.getAjaxErrorMessage(jqXHR));
                }.bind(this),
            );
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
    loadProjectMetadata: async function(projectId) {
        var [project, done] = this.getProjectById(projectId);

        if (project.metadataLoaded === true) {
            return;
        }

        const workItemData = await this.execQueryOnSidebar(["ajax", "getVsoProjectWorkItemTypes", project.id]);
        project.workItemTypes = this.restrictToAllowedWorkItems(workItemData.value);

        const areaData = await this.execQueryOnSidebar(["ajax", "getVsoProjectAreas", project.id]);
        var areas = []; // Flatten areas to format \Area 1\Area 1.1

        const visitArea = function(area, currentPath) {
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

        visitArea(areaData);
        project.areas = _.sortBy(areas, function(area) {
            return area.name;
        });

        project.metadataLoaded = true;
        done(); // set project back to localstorage
    },

    /**
     * @return [proj: Project, done: Function]
     * Make sure you call done() when you are done editing
     * the project, or it will not get saved back to storage.
     */
    getProjectById: function(id) {
        const projects = getVm("projects");
        return [
            _.find(projects, function(proj) {
                return proj.id == id;
            }),
            function() {
                assignVm({ projects: projects });
            },
        ];
    },
    setProject: function(proj) {
        const projects = getVm("projects");
        _.find(projects, function(p) {
            return p.id === proj.id;
        });
    },
    getWorkItemTypeByName: function(project, name) {
        return _.find(project.workItemTypes, function(wit) {
            return wit.name == name;
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
    drawTypesList: function(select, projectId) {
        var [project, done] = this.getProjectById(projectId);
        select.html(
            this.renderTemplate("types", {
                types: project.workItemTypes,
            }),
        );
        done();
    },
    drawAreasList: function(select, projectId) {
        var [project, done] = this.getProjectById(projectId);
        select.html(
            this.renderTemplate("areas", {
                areas: project.areas,
            }),
        );
        done();
    },
    showErrorInModal: function($modal, err) {
        if ($modal.find(".modal-body .errors")) {
            $modal
                .find(".modal-body .errors")
                .text(err)
                .show();
            this.resize();
        }
    },
    resize: function(size = {}) {
        // Automatically resize the iframe based on document height, if it's not in the "nav_bar" location
        if (this._context.location !== "nav_bar") {
            this.zafClient.invoke("resize", { height: size.height || this.$("html").height() + 40, width: size.width || this.$("html").outerWidth(true) });
        }
    },
    restrictToAllowedWorkItems: function(wits) {
        return _.filter(wits, function(wit) {
            return _.contains(VSO_WI_TYPES_WHITE_LISTS, wit.name);
        });
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
    getAjaxErrorMessage: function(jqXHR, errMsg) {
        errMsg = errMsg || this.I18n.t("errorAjax"); //Let's try get a friendly message based on some cases

        var serverErrMsg;

        if (jqXHR.responseJSON) {
            serverErrMsg = jqXHR.responseJSON.message || jqXHR.responseJSON.value.Message;
        } else if (jqXHR.responseText) {
            serverErrMsg = jqXHR.responseText.substring(0, 50) + "...";
        }

        var detail = this.I18n.t("errorServer").fmt(jqXHR.status, jqXHR.statusText, serverErrMsg);
        return errMsg + " " + detail;
    },
    events: {
        "app.activated": "onAppActivated",
    },
});
