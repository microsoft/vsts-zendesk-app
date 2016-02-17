//********************************************************* 
// 
// Copyright (c) Microsoft. All rights reserved. 
// THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF 
// ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY 
// IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR 
// PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT. 
// 
//********************************************************* 
/* global helpers, services, _, Base64 */
(function () {
  'use strict';

  //#region Constants

  var
      INSTALLATION_ID = 0,    //For dev purposes, when using Zat, set this to your current installation id
      VSO_URL_FORMAT = "https://%@.visualstudio.com/DefaultCollection",
      VSO_API_DEFAULT_VERSION = "1.0",
      VSO_API_RESOURCE_VERSION = {},
      TAG_PREFIX = "vso_wi_",
      DEFAULT_FIELD_SETTINGS = JSON.stringify({
        "System.WorkItemType": { summary: true, details: true },
        "System.Title": { summary: false, details: true },
        "System.Description": { summary: true, details: true }
      }),
      VSO_ZENDESK_LINK_TO_TICKET_PREFIX = "ZendeskLinkTo_Ticket_",
      VSO_ZENDESK_LINK_TO_TICKET_ATTACHMENT_PREFIX = "ZendeskLinkTo_Attachment_Ticket_",
      VSO_WI_TYPES_WHITE_LISTS = ["Bug", "Product Backlog Item", "User Story", "Requirement", "Issue"];

  //#endregion

  return {
    defaultState: 'loading',

    //Global view model shared by all instances
    vm: {
      accountUrl: null,
      projects: [],
      fields: [],
      fieldSettings: {},
      userProfile: {},
      isAppLoadedOk: false
    },

    //#region Events Declaration
    events: {
      // App
      'app.activated': 'onAppActivated',

      // Requests
      'getVsoProjects.done': 'onGetVsoProjectsDone',
      'getVsoFields.done': 'onGetVsoFieldsDone',

      //New workitem dialog
      'click .newWorkItem': 'onNewWorkItemClick',
      'change .newWorkItemModal .inputVsoProject': 'onNewVsoProjectChange',
      'change .newWorkItemModal .type': 'onNewVsoWorkItemTypeChange',
      'click .newWorkItemModal .copyDescription': 'onNewCopyDescriptionClick',
      'click .newWorkItemModal .accept': 'onNewWorkItemAcceptClick',

      //Admin side pane
      'click .cog': 'onCogClick',
      'click .closeAdmin': 'onCloseAdminClick',
      'change .summary,.details': 'onSettingChange',

      //Details dialog
      'click .showDetails': 'onShowDetailsClick',

      //Link work item dialog
      'click .link': 'onLinkClick',
      'change .linkModal .project': 'onLinkVsoProjectChange',
      'click .linkModal button.queryBtn': 'onLinkQueryButtonClick',
      'click .linkModal button.reloadQueriesBtn': 'onLinkReloadQueriesButtonClick',
      'click .linkModal button.accept': 'onLinkAcceptClick',
      'click .linkModal button.search': 'onLinkSearchClick',
      'click .linkModal a.workItemResult': 'onLinkResultClick',

      //Unlink click
      'click .unlink': 'onUnlinkClick',
      'click .unlinkModal .accept': 'onUnlinkAcceptClick',

      //Notify dialog
      'click .notify': 'onNotifyClick',
      'click .notifyModal .accept': 'onNotifyAcceptClick',
      'click .notifyModal .copyLastComment': 'onCopyLastCommentClick',

      //Refresh work items
      'click .refreshWorkItemsLink': 'onRefreshWorkItemClick',

      //Login
      'click .user,.user-link': 'onUserIconClick',
      'click .closeLogin': 'onCloseLoginClick',
      'click .login-button': 'onLoginClick'
    },

    //#endregion

    //#region Requests
    requests: {
      getComments: function () {
        return {
          url: helpers.fmt('/api/v2/tickets/%@/comments.json', this.ticket().id()),
          type: 'GET'
        };
      },

      addTagToTicket: function (tag) {
        return {
          url: helpers.fmt('/api/v2/tickets/%@/tags.json', this.ticket().id()),
          type: 'PUT',
          dataType: 'json',
          data: {
            "tags": [tag]
          }
        };
      },

      removeTagFromTicket: function (tag) {
        return {
          url: helpers.fmt('/api/v2/tickets/%@/tags.json', this.ticket().id()),
          type: 'DELETE',
          dataType: 'json',
          data: {
            "tags": [tag]
          }
        };
      },

      addPrivateCommentToTicket: function (text) {
        return {
          url: helpers.fmt('/api/v2/tickets/%@.json', this.ticket().id()),
          type: 'PUT',
          dataType: 'json',
          data: {
            "ticket": {
              "comment": {
                "public": false,
                "body": text
              }
            }
          }
        };
      },

      saveSettings: function (data) {
        return {
          type: 'PUT',
          url: helpers.fmt("/api/v2/apps/installations/%@.json", this.installationId() || INSTALLATION_ID),
          dataType: 'json',
          data: {
            enabled: true,
            settings: data
          }
        };
      },

      getVsoProjects: function () { return this.vsoRequest('/_apis/projects'); },
      getVsoProjectWorkItemTypes: function (projectId) { return this.vsoRequest(helpers.fmt('/%@/_apis/wit/workitemtypes', projectId)); },
      getVsoProjectWorkItemQueries: function (projectName) { return this.vsoRequest(helpers.fmt('/%@/_apis/wit/queries', projectName), { $depth: 2 }); },
      getVsoFields: function () { return this.vsoRequest('/_apis/wit/fields'); },
      getVsoWorkItems: function (ids) { return this.vsoRequest('/_apis/wit/workItems', { ids: ids, '$expand': 'relations' }); },
      getVsoWorkItem: function (workItemId) { return this.vsoRequest(helpers.fmt('/_apis/wit/workItems/%@', workItemId), { '$expand': 'relations' }); },
      getVsoWorkItemQueryResult: function (projectName, queryId) { return this.vsoRequest(helpers.fmt('/%@/_apis/wit/wiql/%@', projectName, queryId)); },
      createVsoWorkItem: function (projectId, witName, data) {
        return this.vsoRequest(helpers.fmt('/%@/_apis/wit/workitems/$%@', projectId, witName), undefined, {
          type: 'PUT',
          contentType: 'application/json-patch+json',
          data: JSON.stringify(data),
          headers: {
            'X-HTTP-Method-Override': 'PATCH',
          }
        });
      },

      updateVsoWorkItem: function (workItemId, data) {
        return this.vsoRequest(helpers.fmt('/_apis/wit/workItems/%@', workItemId), undefined, {
          type: 'PUT',
          contentType: 'application/json-patch+json',
          data: JSON.stringify(data),
          headers: {
            'X-HTTP-Method-Override': 'PATCH',
          },
        });
      },

      updateMultipleVsoWorkItem: function (data) {
        return this.vsoRequest('/_apis/wit/workItems', undefined, {
          type: 'PUT',
          contentType: 'application/json',
          data: JSON.stringify(data),
          headers: {
            'X-HTTP-Method-Override': 'PATCH',
          },
        });
      }
    },

    onGetVsoProjectsDone: function (projects) {
      this.vm.projects = _.sortBy(_.map(projects.value, function (project) {
        return {
          id: project.id,
          name: project.name,
          workItemTypes: []
        };
      }), function (project) {
        return project.name.toLowerCase();
      });
    },

    onGetVsoFieldsDone: function (data) {
      this.vm.fields = _.map(data.value, function (field) {
        return {
          refName: field.referenceName,
          name: field.name,
        };
      });
    },

    getLinkedVsoWorkItems: function (func) {
      var vsoLinkedIds = this.getLinkedWorkItemIds();

      var finish = function (workItems) {
        if (func && _.isFunction(func)) { func(workItems); } else { this.displayMain(); }
        this.onGetLinkedVsoWorkItemsDone(workItems);
      }.bind(this);

      if (!vsoLinkedIds || vsoLinkedIds.length === 0) {
        finish([]);
        return;
      }

      //make a call for each linked wi to get the data we need (web URL is not returned from the getVsoWorkItems)
      var requests = _.map(vsoLinkedIds, function (workItemId) {
        return this.ajax('getVsoWorkItem', workItemId);
      }.bind(this));

      //wait for all requests to complete
      this.when.apply(this, requests)
      .done(function () {
        var linkedWorkItems = [];
        if (vsoLinkedIds.length === 1) {
          //just one wi: arguments is [data, status, jqXhr]
          linkedWorkItems.push(arguments[0]);
        } else {
          //more than 1 wi: arguments is [[data1, status1, jqXhr1],...]
          for (var i = 0; i < arguments.length; i++) {
            linkedWorkItems.push(arguments[i][0]);
          }
        }
        finish(linkedWorkItems);
      }.bind(this))
      .fail(function (jqXHR) {
        this.displayMain(this.getAjaxErrorMessage(jqXHR));
      }.bind(this));

      //this.ajax('getVsoWorkItems', vsoLinkedIds.join(','))
      //    .done(function (data) { finish(data.value); })
      //    .fail(function (jqXHR) { this.displayMain(this.getAjaxErrorMessage(jqXHR)); }.bind(this));
    },

    onGetLinkedVsoWorkItemsDone: function (data) {
      this.vmLocal.workItems = data;
      _.each(this.vmLocal.workItems, function (workItem) {
        workItem.title = helpers.fmt("%@: %@", workItem.id, this.getWorkItemFieldValue(workItem, "System.Title"));
      }.bind(this));
      this.drawWorkItems();
    },

    //#endregion

    //#region Events Implementation

    // App
    onAppActivated: function (data) {

      if (data.firstLoad) {
        //Check if everything is ok to continue
        if (!(this.setting('vso_account'))) {
          return this.switchTo('finish_setup');
        }

        var ticket = this.ticket(),
          requester = ticket.requester(),
          organization = ticket.organization(),
          org_field_vso_account = this.setting('org_field_vso_account'),
          user_field_vso_account = this.setting('user_field_vso_account');

        var current_vso_account =
          (user_field_vso_account && requester && requester.customField(user_field_vso_account)) ||
          (org_field_vso_account && organization && organization.customField(org_field_vso_account)) ||
          this.setting('vso_account');

        //Private instance view model 
        this.vmLocal = {
          vso_account: current_vso_account,
          workItems: [] 
        };

        //set account url
        this.vm.accountUrl = this.buildAccountUrl();

        if (!this.store("auth_token_for_" + this.vmLocal.vso_account)) {
          return this.switchTo('login', this.vmLocal);
        }

        if (!this.vm.isAppLoadedOk) {
          //Initialize global data
          try {
            this.vm.fieldSettings = JSON.parse(this.setting('vso_field_settings') || DEFAULT_FIELD_SETTINGS);
          } catch (ex) {
            services.notify(this.I18n.t('errorReadingFieldSettings'), 'alert');
            this.vm.fieldSettings = JSON.parse(DEFAULT_FIELD_SETTINGS);
          }
          this.when(
              this.ajax('getVsoProjects'),
              this.ajax('getVsoFields')
          ).done(function () {
            this.vm.isAppLoadedOk = true;
            this.getLinkedVsoWorkItems();
          }.bind(this))
          .fail(function (jqXHR, textStatus, err) {
            this.switchTo('error_loading_app', {
              invalidAccount: jqXHR.status === 404,
              accountName: this.vmLocal.vso_account
          });
          }.bind(this));
        } else {
          this.getLinkedVsoWorkItems();
        }
      }
    },

    // UI
    onNewWorkItemClick: function () {

      var $modal = this.$('.newWorkItemModal').modal();
      $modal.find('.modal-body').html(this.renderTemplate('loading'));
      this.ajax('getComments').done(function (data) {
        var attachments = _.flatten(_.map(data.comments, function (comment) {
          return comment.attachments || [];
        }), true);
        $modal.find('.modal-body').html(this.renderTemplate('new', { attachments: attachments }));
        $modal.find('.summary').val(this.ticket().subject());

        var projectCombo = $modal.find('.project');
        this.fillComboWithProjects(projectCombo);
        projectCombo.change();

      }.bind(this));
    },

    onNewVsoProjectChange: function () {
      var $modal = this.$('.newWorkItemModal');
      var projId = $modal.find('.project').val();

      this.showSpinnerInModal($modal);

      this.loadProjectWorkItemTypes(projId)
      .done(function () {
        this.drawTypesList($modal.find('.type'), projId);
        $modal.find('.type').change();
        this.hideSpinnerInModal($modal);
      }.bind(this))
      .fail(function (jqXHR) {
        this.showErrorInModal($modal, this.getAjaxErrorMessage(jqXHR));
      }.bind(this));
    },

    onNewVsoWorkItemTypeChange: function () {
      var $modal = this.$('.newWorkItemModal');
      var project = this.getProjectById($modal.find('.project').val());
      var workItemType = this.getWorkItemTypeByName(project, $modal.find('.type').val());

      //Check if we have severity
      if (this.hasFieldDefined(workItemType, "Microsoft.VSTS.Common.Severity")) {
        $modal.find('.severityInput').show();
      } else {
        $modal.find('.severityInput').hide();
      }
    },

    onNewCopyDescriptionClick: function (event) {
      event.preventDefault();
      this.$('.newWorkItemModal .description').val(this.ticket().description());
    },

    onNewWorkItemAcceptClick: function () {
      var $modal = this.$('.newWorkItemModal').modal();

      //check project
      var proj = this.getProjectById($modal.find('.project').val());
      if (!proj) { return this.showErrorInModal($modal, this.I18n.t("modals.new.errProjRequired")); }

      //check work item type
      var workItemType = this.getWorkItemTypeByName(proj, $modal.find('.type').val());
      if (!workItemType) { return this.showErrorInModal($modal, this.I18n.t("modals.new.errWorkItemTypeRequired")); }

      //check summary
      var summary = $modal.find(".summary").val();
      if (!summary) { return this.showErrorInModal($modal, this.I18n.t("modals.new.errSummaryRequired")); }

      var description = $modal.find(".description").val();
      var attachments = this.getSelectedAttachments($modal);

      var operations = [].concat(
          this.buildPatchToAddWorkItemField("System.Title", summary),
          this.buildPatchToAddWorkItemField("System.Description", description));

      if (this.hasFieldDefined(workItemType, "Microsoft.VSTS.Common.Severity") && $modal.find('.severity').val()) {
        operations.push(this.buildPatchToAddWorkItemField("Microsoft.VSTS.Common.Severity", $modal.find('.severity').val()));
      }

      if (this.hasFieldDefined(workItemType, "Microsoft.VSTS.TCM.ReproSteps")) {
        operations.push(this.buildPatchToAddWorkItemField("Microsoft.VSTS.TCM.ReproSteps", description));
      }

      //Set tag
      if (this.setting("vso_tag")) {
        operations.push(this.buildPatchToAddWorkItemField("System.Tags", this.setting("vso_tag")));
      }

      //Add hyperlink to ticket url
      operations.push(this.buildPatchToAddWorkItemHyperlink(
        this.buildTicketLinkUrl(),
        VSO_ZENDESK_LINK_TO_TICKET_PREFIX + this.ticket().id()));

      //Add hyperlinks to attachments
      operations = operations.concat(this.buildPatchToAddWorkItemAttachments(attachments));

      this.showSpinnerInModal($modal);

      this.ajax('createVsoWorkItem', proj.id, workItemType.name, operations)
        .done(function (data) {
          var newWorkItemId = data.id;
          //sanity check due tfs returning 200 ok  but with exception
          if (newWorkItemId > 0) { this.linkTicket(newWorkItemId); }

          services.notify(this.I18n.t('notify.workItemCreated').fmt(newWorkItemId));
          this.getLinkedVsoWorkItems(function () { this.closeModal($modal); }.bind(this));
        }.bind(this))
        .fail(function (jqXHR) {
          this.showErrorInModal($modal, this.getAjaxErrorMessage(jqXHR));
        }.bind(this));
    },

    onCogClick: function () {
      this.switchTo('admin');
      this.drawSettings();
    },

    onCloseAdminClick: function () {
      this.displayMain();
    },

    onSettingChange: function () {
      var self = this;
      var fieldSettings = {};
      this.$('tr').each(function () {
        var line = self.$(this);
        var fieldName = line.attr('data-refName');
        if (!fieldName) return true; //continue

        var inSummary = line.find('.summary').is(':checked');
        var inDetails = line.find('.details').is(':checked');

        if (inSummary || inDetails) {
          fieldSettings[fieldName] = {
            summary: inSummary,
            details: inDetails
          };
        } else if (fieldSettings[fieldName]) {
          delete fieldName[fieldName];
        }
      });
      this.vm.fieldSettings = fieldSettings;
      this.ajax('saveSettings', { vso_field_settings: JSON.stringify(fieldSettings) })
          .done(function () {
            services.notify(this.I18n.t('admin.settingsSaved'));
          }.bind(this));
    },

    onShowDetailsClick: function (event) {
      var $modal = this.$('.detailsModal').modal();
      $modal.find('.modal-header h3').html(this.I18n.t('modals.details.loading'));
      $modal.find('.modal-body').html(this.renderTemplate('loading'));
      var id = this.$(event.target).closest('.workItem').attr('data-id');
      var workItem = this.getWorkItemById(id);
      workItem = this.attachRestrictedFieldsToWorkItem(workItem, 'details');
      $modal.find('.modal-header h3').html(this.I18n.t('modals.details.title', { name: workItem.title }));
      $modal.find('.modal-body').html(this.renderTemplate('details', workItem));
    },

    onLinkClick: function () {
      var $modal = this.$('.linkModal').modal();
      $modal.find('.modal-footer button').removeAttr('disabled');
      $modal.find('.modal-body').html(this.renderTemplate('link'));
      $modal.find("button.search").show();

      var projectCombo = $modal.find('.project');
      this.fillComboWithProjects(projectCombo);
      projectCombo.change();
    },

    onLinkSearchClick: function () {
      var $modal = this.$('.linkModal');
      $modal.find(".search-section").show();
    },

    onLinkResultClick: function (event) {
      event.preventDefault();
      var $modal = this.$('.linkModal');
      var id = this.$(event.target).closest('.workItemResult').attr('data-id');
      $modal.find('.inputVsoWorkItemId').val(id);
      $modal.find('.search-section').hide();
    },

    onLinkVsoProjectChange: function () {
      this.loadQueriesList();
    },

    onLinkReloadQueriesButtonClick: function () {
      this.loadQueriesList(true);
    },

    loadQueriesList: function (reload) {
      var $modal = this.$('.linkModal');
      var projId = $modal.find('.project').val();

      this.showSpinnerInModal($modal);

      this.loadProjectWorkItemQueries(projId, reload)
      .done(function () {
        this.drawQueriesList($modal.find('.query'), projId);
        this.hideSpinnerInModal($modal);
      }.bind(this))
      .fail(function (jqXHR) {
        this.showErrorInModal($modal, this.getAjaxErrorMessage(jqXHR));
      }.bind(this));      
    },

    onLinkQueryButtonClick: function () {
      var $modal = this.$('.linkModal');
      var projId = $modal.find('.project').val();
      var queryId = $modal.find('.query').val();

      var drawQueryResults = function (results, countQueryItemsResult) {
        var workItems = _.map(results, function (workItem) {
          return {
            id: workItem.id,
            type: this.getWorkItemFieldValue(workItem, "System.WorkItemType"),
            title: this.getWorkItemFieldValue(workItem, "System.Title")
          };
        }.bind(this));

        $modal.find('.results').html(this.renderTemplate('query_results', { workItems: workItems }));
        $modal.find('.alert-success').html(this.I18n.t('queryResults.returnedWorkItems', { count: countQueryItemsResult }));
        this.hideSpinnerInModal($modal);

      }.bind(this);

      this.showSpinnerInModal($modal);

      this.ajax('getVsoWorkItemQueryResult', this.getProjectById(projId).name, queryId)
          .done(function (data) {

            var getWorkItemsIdsFromQueryResult = function (result) {
              if (result.queryType === 'oneHop' || result.queryType === 'tree') {
                return _.map(result.workItemRelations, function (rel) { return rel.target.id; });
              } else {
                return _.pluck(result.workItems, 'id');
              }
            };

            var ids = getWorkItemsIdsFromQueryResult(data);
            if (!ids || ids.length === 0) {
              return drawQueryResults([], 0);
            }

            this.ajax('getVsoWorkItems', _.first(ids, 200).join(',')).done(function (results) {
              drawQueryResults(results.value, ids.length);
            });
          }.bind(this))
          .fail(function (jqXHR, textStatus, errorThrown) {
            this.showErrorInModal($modal, this.getAjaxErrorMessage(jqXHR, this.I18n.t('modals.link.errCannotGetWorkItem')));
          }.bind(this));

    },

    onLinkAcceptClick: function (event) {
      var $modal = this.$('.linkModal');
      var workItemId = $modal.find('.inputVsoWorkItemId').val();

      if (!/^([0-9]+)$/.test(workItemId)) {
        return this.showErrorInModal($modal, this.I18n.t('modals.link.errWorkItemIdNaN'));
      }

      if (this.isAlreadyLinkedToWorkItem(workItemId)) {
        return this.showErrorInModal($modal, this.I18n.t('modals.link.errAlreadyLinked'));
      }

      this.showSpinnerInModal($modal);
      var updateWorkItem = function (workItem) {

        //Let's check if there is already a link in the WI returned data
        var currentLink = _.find(workItem.relations || [], function (link) {
          if (link.rel.toLowerCase() === "hyperlink" && link.attributes.name === (VSO_ZENDESK_LINK_TO_TICKET_PREFIX + this.ticket().id())) {
            return link;
          }
        }.bind(this));

        var finish = function () {
          this.linkTicket(workItemId);
          services.notify(this.I18n.t('notify.workItemLinked').fmt(workItemId));
          this.getLinkedVsoWorkItems(function () { this.closeModal($modal); }.bind(this));
        }.bind(this);

        if (currentLink) {
          finish();
        } else {

          var addLinkOperation = this.buildPatchToAddWorkItemHyperlink(
                  this.buildTicketLinkUrl(),
                  VSO_ZENDESK_LINK_TO_TICKET_PREFIX + this.ticket().id());

          this.ajax('updateVsoWorkItem', workItemId, [addLinkOperation])
              .done(function () {
                finish();
              }.bind(this))
              .fail(function (jqXHR) {
                this.showErrorInModal($modal, this.getAjaxErrorMessage(jqXHR, this.I18n.t('modals.link.errCannotUpdateWorkItem')));
              }.bind(this));
        }
      }.bind(this);

      //Get work item and then update
      this.ajax('getVsoWorkItem', workItemId)
          .done(function (data) {
            updateWorkItem(data);
          }.bind(this))
          .fail(function (jqXHR) {
            this.showErrorInModal($modal, this.getAjaxErrorMessage(jqXHR, this.I18n.t('modals.link.errCannotGetWorkItem')));
          }.bind(this));
    },

    onUnlinkClick: function (event) {
      var id = this.$(event.target).closest('.workItem').attr('data-id');
      var workItem = this.getWorkItemById(id);
      var $modal = this.$('.unlinkModal').modal();
      $modal.find('.modal-body').html(this.renderTemplate('unlink'));
      $modal.find('.modal-footer button').removeAttr('disabled');
      $modal.find('.modal-body .confirm').html(this.I18n.t('modals.unlink.text', { name: workItem.title }));
      $modal.attr('data-id', id);
    },

    onUnlinkAcceptClick: function (event) {
      event.preventDefault();
      var $modal = this.$(event.target).closest('.unlinkModal');

      this.showSpinnerInModal($modal);
      var workItemId = $modal.attr('data-id');

      var updateWorkItem = function (workItem) {
        //Calculate the positions of links to remove
        var posOfLinksToRemove = [];

        _.each(workItem.relations, function (link, idx) {
          if (link.rel.toLowerCase() === 'hyperlink' &&
              (link.attributes.name === VSO_ZENDESK_LINK_TO_TICKET_PREFIX + this.ticket().id() ||
              link.attributes.name === VSO_ZENDESK_LINK_TO_TICKET_ATTACHMENT_PREFIX + this.ticket().id())) {
            posOfLinksToRemove.push(idx - posOfLinksToRemove.length);
          }
        }.bind(this));

        var finish = function () {
          this.unlinkTicket(workItem.id);
          services.notify(this.I18n.t('notify.workItemUnlinked').fmt(workItem.id));
          this.getLinkedVsoWorkItems(function () { this.closeModal($modal); }.bind(this));
        }.bind(this);

        if (posOfLinksToRemove.length === 0) {
          finish();
        } else {
          var operations = [{ op: "test", path: "/rev", value: workItem.rev }]
              .concat(_.map(posOfLinksToRemove, function (pos) {
                return this.buildPatchToRemoveWorkItemHyperlink(pos);
              }.bind(this)));

          this.ajax('updateVsoWorkItem', workItemId, operations)
              .done(function () { finish(); })
              .fail(function (jqXHR) {
                this.showErrorInModal($modal, this.getAjaxErrorMessage(jqXHR, this.I18n.t('modals.unlink.errUnlink')));
              }.bind(this));
        }
      }.bind(this);

      //Get work item to get the last revision and then update
      this.ajax('getVsoWorkItem', workItemId)
          .done(function (workItem) { updateWorkItem(workItem); }.bind(this))
          .fail(function (jqXHR) { this.showErrorInModal($modal, this.getAjaxErrorMessage(jqXHR)); }.bind(this));
    },

    onNotifyClick: function () {
      var $modal = this.$('.notifyModal');
      $modal.find('.modal-body').html(this.renderTemplate('loading'));
      $modal.modal();

      this.ajax('getComments').done(function (data) {
        this.lastComment = data.comments[data.comments.length - 1].body;
        var attachments = _.flatten(_.map(data.comments, function (comment) {
          return comment.attachments || [];
        }), true);
        $modal.find('.modal-body').html(this.renderTemplate('notify', { attachments: attachments }));
        $modal.find('.modal-footer button').prop('disabled', false);
      }.bind(this));
    },

    onNotifyAcceptClick: function () {
      var $modal = this.$('.notifyModal');
      var text = $modal.find('textarea').val();

      if (!text) { return this.showErrorInModal($modal, this.I18n.t("modals.notify.errCommentRequired")); }

      var attachments = this.getSelectedAttachments($modal);

      this.showSpinnerInModal($modal);

      //Refresh linked VSO work items
      this.getLinkedVsoWorkItems(function (workItems) {

        //create an array of promises with individual request
        var requests = _.map(workItems, function (workItem) {

          //exclude selected attachments that are already in the work item
          var newAttachments = _.reject(attachments, function (att) {
            return _.some(workItem.relations || [], function (rel) {
              return rel.url === att.url;
            });
          });

          var operations = [this.buildPatchToAddWorkItemField("System.History", text)].concat(
            this.buildPatchToAddWorkItemAttachments(newAttachments));
          return this.ajax('updateVsoWorkItem', workItem.id, operations);
        }.bind(this));

        //wait for all requests to complete
        this.when.apply(this, requests)
        //this.ajax('updateMultipleVsoWorkItem', updatePayload)
        .done(function () {
          var ticketMsg = [this.I18n.t('notify.message', { name: this.currentUser().name() }), text].join("\n\r\n\r");
          this.ajax('addPrivateCommentToTicket', ticketMsg);
          services.notify(this.I18n.t('notify.notification'));
          this.closeModal($modal);
        }.bind(this))
        .fail(function (jqXHR) {
          this.showErrorInModal($modal, this.getAjaxErrorMessage(jqXHR));
        }.bind(this));
      }.bind(this));
    },

    onCopyLastCommentClick: function (event) {
      event.preventDefault();
      this.$('.notifyModal').find('textarea').val(this.lastComment);
    },

    onRefreshWorkItemClick: function (event) {
      event.preventDefault();
      this.$('.workItemsError').hide();
      this.switchTo('loading');
      this.getLinkedVsoWorkItems();
    },

    onLoginClick: function (event) {
      event.preventDefault();
      var vso_username = this.$('.vso_username').val();
      var vso_password = this.$('.vso_password').val();

      if (!vso_username || !vso_password) {
        this.$(".login-form").find('.errors').text(this.I18n.t("login.errRequiredFields")).show();
        return;
      }

      this.authString(vso_username, vso_password);
      services.notify(this.I18n.t('notify.credentialsSaved'));

      this.switchTo('loading');
      if (!this.vm.isAppLoadedOk) {
        this.onAppActivated({ firstLoad: true });
      } else {
        this.getLinkedVsoWorkItems();
      }
    },

    onCloseLoginClick: function () {
      this.displayMain();
    },

    onUserIconClick: function () {
      this.switchTo('login', this.vmLocal);
    },

    //#endregion

    //#region Drawing

    displayMain: function (err) {
      if (this.vm.isAppLoadedOk) {
        this.$('.cog').toggle(this.isAdmin());
        this.switchTo('main');
        if (!err) {
          this.drawWorkItems();
        } else {
          this.$('.workItemsError').show();
        }
      } else {
        this.$('.cog').toggle(false);
        this.switchTo('error_loading_app');
      }
    },

    drawWorkItems: function (data) {

      var workItems = _.map(data || this.vmLocal.workItems, function (workItem) {
        var tmp = this.attachRestrictedFieldsToWorkItem(workItem, 'summary');
        return tmp;
      }.bind(this));

      this.$('.workItems').html(this.renderTemplate('workItems', { workItems: workItems }));
      this.$('.buttons .notify').prop('disabled', !workItems.length);
    },

    drawTypesList: function (select, projectId) {
      var project = this.getProjectById(projectId);
      select.html(this.renderTemplate('types', { types: project.workItemTypes }));
    },

    drawQueriesList: function (select, projectId) {
      var project = this.getProjectById(projectId);

      var drawNode = function (node, prefix) {
        //It's a folder
        if (node.isFolder) {
          return "<optgroup label='%@ %@'>%@</optgroup>".fmt(
             prefix,
              node.name,
              _.reduce(node.children, function (options, childNode, ix) {
                return "%@%@".fmt(options, drawNode(childNode, prefix + (ix + 1) + "."));
              }, ""));
        }

        //It's a query
        return "<option value='%@'>%@ %@</option>".fmt(node.id, prefix, node.name);

      }.bind(this);

      select.html(_.reduce(project.queries, function (options, query, ix) {
        return "%@%@".fmt(options, drawNode(query, "" + (ix + 1) + "."));
      }, ""));

    },

    drawSettings: function () {
      var settings = _.sortBy(
          _.map(this.vm.fields, function (field) {
            var current = this.vm.fieldSettings[field.refName];
            if (current) { field = _.extend(field, current); }
            return field;
          }.bind(this)), function (f) { return f.name; });

      var html = this.renderTemplate('settings', { settings: settings });
      this.$('.content').html(html);
    },

    showSpinnerInModal: function ($modal) {
      if ($modal.find('.modal-body form')) { $modal.find('.modal-body form').hide(); }
      if ($modal.find('.modal-body .loading')) { $modal.find('.modal-body .loading').show(); }
      if ($modal.find('.modal-footer button')) { $modal.find('.modal-footer button').attr('disabled', 'disabled'); }
    },

    hideSpinnerInModal: function ($modal) {
      if ($modal.find('.modal-body form')) { $modal.find('.modal-body form').show(); }
      if ($modal.find('.modal-body .loading')) { $modal.find('.modal-body .loading').hide(); }
      if ($modal.find('.modal-footer button')) { $modal.find('.modal-footer button').prop('disabled', false); }
    },

    showErrorInModal: function ($modal, err) {
      this.hideSpinnerInModal($modal);
      if ($modal.find('.modal-body .errors')) { $modal.find('.modal-body .errors').text(err).show(); }
    },

    closeModal: function ($modal) {
      $modal.find('#loading').hide();
      $modal.modal('hide').find('.modal-footer button').attr('disabled', '');
    },

    fillComboWithProjects: function (el) {

      el.html(_.reduce(this.vm.projects, function (options, project) {
        return "%@<option value='%@'>%@</option>".fmt(options, project.id, project.name);
      }, ""));
    },

    //#endregion

    //#region Helpers

    isAdmin: function () {
      return this.currentUser().role() === 'admin';
    },

    vsoUrl: function (url, parameters) {
      url = (url[0] === '/') ? url.slice(1) : url;
      var full = [this.vm.accountUrl, url].join('/');
      if (parameters) {
        full += '?' + _.map(parameters, function (value, key) {
          return [key, value].join('=');
        }).join('&');
      }
      return full;
    },

    authString: function (vso_username, vso_password) {

      if (vso_username && vso_password) {
        var b64 = Base64.encode([vso_username, vso_password].join(':'));
        this.store('auth_token_for_' + this.vmLocal.vso_account, b64);
      }

      return helpers.fmt("Basic %@", this.store('auth_token_for_' + this.vmLocal.vso_account));
    },

    vsoRequest: function (url, parameters, options) {
      var requestOptions = _.extend({
        url: this.vsoUrl(url, parameters),
        dataType: 'json',
      }, options);

      var fixedHeaders = {
        'Authorization': this.authString(),
        'Accept': helpers.fmt("application/json;api-version=%@", this.getVsoResourceVersion(url))
      };

      requestOptions.headers = _.extend(fixedHeaders, options ? options.headers : {});
      return requestOptions;
    },

    getVsoResourceVersion: function (url) {
      var resource = url.split("/_apis/")[1].split("/")[0];
      return VSO_API_RESOURCE_VERSION[resource] || VSO_API_DEFAULT_VERSION;

    },

    attachRestrictedFieldsToWorkItem: function (workItem, type) {
      var fields = _.compact(_.map(this.vm.fieldSettings, function (value, key) {
        if (value[type]) {
          if (_.has(workItem.fields, key)) {
            return {
              refName: key,
              name: _.find(this.vm.fields, function (f) { return f.refName == key; }).name,
              value: workItem.fields[key]
            };
          }
        }
      }.bind(this)));
      return _.extend(workItem, { restricted_fields: fields });
    },

    getWorkItemById: function (id) {
      return _.find(this.vmLocal.workItems, function (workItem) { return workItem.id == id; });
    },

    getProjectById: function (id) {
      return _.find(this.vm.projects, function (proj) { return proj.id == id; });
    },

    getWorkItemTypeByName: function (project, name) {
      return _.find(project.workItemTypes, function (wit) { return wit.name == name; });
    },

    getFieldByFieldRefName: function (fieldRefName) {
      return _.find(this.vm.fields, function (f) { return f.refName == fieldRefName; });
    },

    getWorkItemFieldValue: function (workItem, fieldRefName) {
      var field = workItem.fields[fieldRefName];

      return field || "";
    },

    hasFieldDefined: function (workItemType, fieldRefName) {
      return _.some(workItemType.fieldInstances, function (fieldInstance) {
        return fieldInstance.referenceName === fieldRefName;
      });
    },

    linkTicket: function (workItemId) {
      var linkVsoTag = TAG_PREFIX + workItemId;
      this.ticket().tags().add(linkVsoTag);

      this.ajax('addTagToTicket', linkVsoTag);
    },

    unlinkTicket: function (workItemId) {
      var linkVsoTag = TAG_PREFIX + workItemId;
      this.ticket().tags().remove(linkVsoTag);

      this.ajax('removeTagFromTicket', linkVsoTag);
    },

    buildTicketLinkUrl: function () {
      return helpers.fmt("https://%@.zendesk.com/agent/#/tickets/%@", this.currentAccount().subdomain(), this.ticket().id());
    },

    getLinkedWorkItemIds: function () {
      return _.compact(this.ticket().tags().map(function (t) {
        var p = t.indexOf(TAG_PREFIX);
        if (p === 0) { return t.slice(TAG_PREFIX.length); }
      }));
    },

    isAlreadyLinkedToWorkItem: function (id) { return _.contains(this.getLinkedWorkItemIds(), id); },

    loadProjectWorkItemTypes: function (projectId) {
      var project = this.getProjectById(projectId);
      if (project.metadataLoaded === true) { return this.promise(function (done) { done(); }); }

      //Let's load project metadata
      return this.ajax('getVsoProjectWorkItemTypes', project.id).done(function (data) {
        project.workItemTypes = this.restrictToAllowedWorkItems(data.value);
        project.metadataLoaded = true;
      }.bind(this));
    },

    loadProjectWorkItemQueries: function (projectId, reload) {
      var project = this.getProjectById(projectId);
      if (project.queries && !reload) { return this.promise(function (done) { done(); }); }

      //Let's load project queries
      return this.ajax('getVsoProjectWorkItemQueries', project.name).done(function (data) {
        project.queries = data.value;
      }.bind(this));
    },

    restrictToAllowedWorkItems: function (wits) {
      return _.filter(wits, function (wit) { return _.contains(VSO_WI_TYPES_WHITE_LISTS, wit.name); });
    },

    buildPatchToAddWorkItemField: function (fieldName, value) {
      return {
        op: "add",
        path: helpers.fmt("/fields/%@", fieldName),
        value: value
      };
    },

    buildPatchToAddWorkItemHyperlink: function (url, name, comment) {
      return {
        op: "add",
        path: "/relations/-",
        value: {
          rel: "Hyperlink",
          url: url,
          attributes: { "name": name, "comment": comment }
        }
      };
    },

    buildPatchToRemoveWorkItemHyperlink: function (pos) {
      return {
        op: "remove",
        path: helpers.fmt("/relations/%@", pos)
      };
    },

    getAjaxErrorMessage: function (jqXHR, errMsg) {
      errMsg = errMsg || this.I18n.t("errorAjax");

      //Let's try get a friendly message based on some cases
      var serverErrMsg;
      if (jqXHR.responseJSON) {
        serverErrMsg = jqXHR.responseJSON.message || jqXHR.responseJSON.value.Message;
      } else {
        serverErrMsg = jqXHR.responseText.substring(0, 50) + "...";
      }

      var detail = this.I18n.t("errorServer").fmt(jqXHR.status, jqXHR.statusText, serverErrMsg);
      return errMsg + " " + detail;
    },

    buildPatchToAddWorkItemAttachments: function (attachments) {
      return _.map(attachments, function (att) {
        return this.buildPatchToAddWorkItemHyperlink(
          att.url,
          VSO_ZENDESK_LINK_TO_TICKET_ATTACHMENT_PREFIX + this.ticket().id(),
          att.name);
      }.bind(this));
    },

    getSelectedAttachments: function ($modal) {
      var attachments = [];
      $modal.find('.attachments input').each(function (ix, el) {
        var $el = this.$(el);
        if ($el.is(':checked')) {
          attachments.push({
            url: $el.val(),
            name: $el.data('fileName')
          });
        }
      }.bind(this));

      return attachments;
    },

    buildAccountUrl: function () {
      var baseUrl;
      var setting = this.vmLocal.vso_account;
      var loweredSetting = setting.toLowerCase();

      if (loweredSetting.indexOf('http://') === 0 || loweredSetting.indexOf('https://') === 0) {
        baseUrl = setting;
      } else {
        baseUrl = helpers.fmt(VSO_URL_FORMAT, setting);
      }

      baseUrl = (baseUrl[baseUrl.length - 1] === '/') ? baseUrl.slice(0, -1) : baseUrl;

      //check if collection defined
      if (baseUrl.lastIndexOf('/') <= 'https://'.length) {
        baseUrl = baseUrl + '/DefaultCollection';
      }

      return baseUrl;
    }

    //#endregion
  };
}());