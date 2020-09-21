// MPL/GPL
// Opera.Wang 2013/06/04
'use strict';
var EXPORTED_SYMBOLS = ['autoArchivePrefDialog'];

Cu.import('resource://gre/modules/Services.jsm');
try {
  Cu.import('resource:///modules/MailServices.jsm');
} catch (err) {
  Cu.import('resource:///modules/mailServices.js');
}
Cu.import('resource:///modules/gloda/utils.js');
// Cu.import("resource://gre/modules/FileUtils.jsm");
Cu.import('resource:///modules/iteratorUtils.jsm');
Cu.import('resource:///modules/folderUtils.jsm');
try {
  Cu.import('resource:///modules/MailUtils.jsm');
} catch (err) {
  Cu.import('resource:///modules/MailUtils.js');
}
Cu.import('chrome://awsomeAutoArchive/content/aop.jsm');
Cu.import('chrome://awsomeAutoArchive/content/autoArchiveService.jsm');
Cu.import('chrome://awsomeAutoArchive/content/autoArchivePref.jsm');
Cu.import('chrome://awsomeAutoArchive/content/autoArchiveUtil.jsm');
Cu.import('chrome://awsomeAutoArchive/content/log.jsm');
const perfDialogTooltipID = 'awsome_auto_archive-perfDialogTooltip';
const perfDialogAgeTooltipID = 'awsome_auto_archive-perfDialogAgeTooltip';
const XUL = 'http://www.mozilla.org/keymaster/gatekeeper/there.is.only.xul';
const ruleClass = 'awsome_auto_archive-rule';
const ruleHeaderContextMenuID = 'awsome_auto_archive-rule-header-context';

const autoArchivePrefDialog = {
  strBundle: Services.strings.createBundle('chrome://awsomeAutoArchive/locale/awsome_auto_archive.properties'),
  hookedFunctions: [],
  _doc: null,
  _win: null,
  _savedRules: '',
  cleanup: function () {
    autoArchiveLog.info('autoArchivePrefDialog cleanup');
    if (this._win && !this._win.closed) this._win.close();
    autoArchiveLog.info('autoArchivePrefDialog cleanup done');
  },

  showPrettyTooltip: function (URI, pretty) {
    return decodeURIComponent(URI).replace(/(.*\/)[^/]*/, '$1') + pretty;
  },
  getPrettyName: function (msgFolder) {
    // access msgFolder.prettyName for non-existing local folder may cause creating wrong folder
    if (!(msgFolder instanceof Ci.nsIMsgFolder) || msgFolder.server.type != 'none' || autoArchiveUtil.folderExists(msgFolder)) return msgFolder.prettyName;
    return msgFolder.URI.replace(/^.*\/([^\/]+)/, '$1');
  },
  getFolderAndSetLabel: function (folderPicker, setLabel) {
    let msgFolder;
    try {
      msgFolder = MailUtils.getExistingFolder(folderPicker.value);
    } catch (err) {}
    if (!msgFolder) msgFolder = { value: '', prettyName: 'N/A', server: {} };
    if (!this._doc || !setLabel) return msgFolder;
    const showFolderAs = Preferences.get('extensions.awsome_auto_archive.show_folder_as');
    let label = '';
    switch (showFolderAs.value) {
      case 0:
        label = self.getPrettyName(msgFolder);
        break;
      case 1:
        label = '[' + msgFolder.server.prettyName + '] ' + (msgFolder == msgFolder.rootFolder ? '/' : self.getPrettyName(msgFolder));
        break;
      case 2:
      default:
        label = self.showPrettyTooltip(msgFolder.ValueUTF8 || msgFolder.value, self.getPrettyName(msgFolder));
        break;
    }
    folderPicker.setAttribute('label', label);
    folderPicker.setAttribute('folderStyle', showFolderAs.value); // for css to set correct length
    return msgFolder;
  },
  changeShowFolderAs: function () {
    if (!this._doc) return;
    const container = this._doc.getElementById('awsome_auto_archive-rules');
    if (!container) return;
    for (const row of container.childNodes) {
      if (row.classList.contains(ruleClass)) {
        for (const item of row.childNodes) {
          const key = item.getAttribute('rule');
          if (['src', 'dest'].indexOf(key) >= 0 /* && item.style.visibility != 'hidden' */) { this.getFolderAndSetLabel(item, true); }
        }
      }
    }
  },
  updateFolderStyle: function (folderPicker, folderPopup, init) {
    const msgFolder = this.getFolderAndSetLabel(folderPicker, false);
    const updateStyle = function () {
      let hasError = !autoArchiveUtil.folderExists(msgFolder);
      try {
        if (typeof (folderPopup.selectFolder) !== 'undefined') folderPopup.selectFolder(msgFolder); // false alarm by addon validator
        else return;
        if (!hasError) folderPopup._setCssSelectors(msgFolder, folderPicker); // _setCssSelectors may also create wrong local folders
      } catch (err) {
        hasError = true;
        // autoArchiveLog.logException(err);
      }
      if (hasError) {
        autoArchiveLog.info("Error: folder '" + self.getPrettyName(msgFolder) + "' can't be selected");
        folderPicker.classList.add('awsome_auto_archive-folderError');
        folderPicker.classList.remove('folderMenuItem');
      } else {
        folderPicker.classList.remove('awsome_auto_archive-folderError');
        folderPicker.classList.add('folderMenuItem');
      }
      if (msgFolder.noSelect) folderPicker.setAttribute('NoSelect', 'true');
      else folderPicker.removeAttribute('NoSelect');
      self.getFolderAndSetLabel(folderPicker, true);
    };
    if (msgFolder.rootFolder) {
      if (init) this._win.setTimeout(updateStyle, 0); // use timer to wait for the XBL bindings add SelectFolder / _setCssSelectors to popup
      else updateStyle();
    }
    folderPicker.setAttribute('tooltiptext', self.showPrettyTooltip(msgFolder.ValueUTF8 || msgFolder.value, self.getPrettyName(msgFolder)));
  },
  onFolderPick: function (folderPicker, aEvent, folderPopup) {
    const folder = aEvent.target._folder;
    if (!folder) return;
    const value = folder.URI || folder.folderURL;
    folderPicker.value = value; // must set value before set label, or next line may fail when previous value is empty
    self.updateFolderStyle(folderPicker, folderPopup, false);
  },
  initFolderPick: function (folderPicker, folderPopup, isSrc) {
    folderPicker.addEventListener('command', function (aEvent) { return self.onFolderPick(folderPicker, aEvent, folderPopup); }, false);
    /* Folders like [Gmail] are disabled by default, enable them, also true for IMAP servers like mercury/32, http://kb.mozillazine.org/Grey_italic_folders */
    const nsResolver = function (prefix) {
      const ns = { xul: 'http://www.mozilla.org/keymaster/gatekeeper/there.is.only.xul' };
      return ns[prefix] || null;
    };
    folderPicker.addEventListener('popupshown', function (aEvent) {
      try {
        // nsIDOMXPathResult was removed in TB60, so have to use self._doc.defaultView.XPathResult
        const menuitems = self._doc.evaluate(".//xul:menuitem[@disabled='true' and @generated='true']", folderPicker, nsResolver, self._doc.defaultView.XPathResult.UNORDERED_NODE_SNAPSHOT_TYPE, null);
        for (let i = 0; i < menuitems.snapshotLength; i++) {
          const menuitem = menuitems.snapshotItem(i);
          if (menuitem._folder && menuitem._folder.noSelect) {
            menuitem.removeAttribute('disabled');
            menuitem.setAttribute('NoSelect', 'true'); // so it will show as in folder pane
          }
        }
      } catch (err) { autoArchiveLog.logException(err); }
    }, false);
    folderPicker.classList.add('folderMenuItem');
    folderPicker.setAttribute('crop', 'center');

    folderPopup.setAttribute('mode', 'search');
    folderPopup.setAttribute('showAccountsFileHere', 'true');
    if (!isSrc) {
      folderPopup.setAttribute('mode', 'filing');
      folderPopup.setAttribute('showFileHereLabel', 'true');
    }
    folderPopup.classList.add('menulist-menupopup');
    self.updateFolderStyle(folderPicker, folderPopup, true);
  },
  createRuleHeader: function () {
    try {
      const doc = this._doc;
      const container = doc.getElementById('awsome_auto_archive-rules');
      if (!container) return;
      while (container.firstChild) container.removeChild(container.firstChild);
      // container.style.height="100 px";
      // container.style.backgroundColor = "red";
      const row = doc.createElementNS(XUL, 'row');
      ['', 'action', 'source', 'scope', 'dest', 'from', 'recipient', 'subject', 'size', 'tags', 'age', '', '', 'picker'].forEach(function (label) {
        let item;
        if (label == 'picker') {
          item = doc.createElementNS(XUL, 'image');
          item.classList.add('tree-columnpicker-icon');
          item.addEventListener('click', function (event) { return doc.getElementById(ruleHeaderContextMenuID).openPopup(item, 'after_start', 0, 0, true, false, event); }, false);
          item.setAttribute('tooltiptext', self.strBundle.GetStringFromName('perfdialog.tooltip.picker'));
        } else {
          item = doc.createElementNS(XUL, 'label');
          item.setAttribute('value', label ? self.strBundle.GetStringFromName('perfdialog.' + label) : '');
          item.setAttribute('rule', label); // header does not have class ruleClass
        }
        const preference = Preferences.get('extensions.awsome_auto_archive.show_' + label);
        if (preference) {
          const actualValue = preference.value !== undefined ? preference.value : preference.defaultValue;
          item.style.display = actualValue ? '-moz-box' : 'none';
        }
        row.insertBefore(item, null);
      });
      row.id = 'awsome_auto_archive-rules-header';
      row.setAttribute('context', ruleHeaderContextMenuID);
      container.insertBefore(row, null);
    } catch (err) {
      autoArchiveLog.logException(err);
    }
  },
  creatOneRule: function (rule, ref) {
    try {
      const doc = this._doc;
      const container = doc.getElementById('awsome_auto_archive-rules');
      if (!container) return;
      const row = doc.createElementNS(XUL, 'row');

      const enable = doc.createElementNS(XUL, 'checkbox');
      enable.setAttribute('checked', rule.enable);
      enable.setAttribute('rule', 'enable');

      const menulistAction = doc.createElementNS(XUL, 'menulist');
      const menupopupAction = doc.createElementNS(XUL, 'menupopup');
      ['archive', 'copy', 'delete', 'move'].forEach(function (action) {
        const menuitem = doc.createElementNS(XUL, 'menuitem');
        menuitem.setAttribute('label', self.strBundle.GetStringFromName('perfdialog.action.' + action));
        menuitem.setAttribute('value', action);
        menupopupAction.insertBefore(menuitem, null);
      });
      menulistAction.insertBefore(menupopupAction, null);
      menulistAction.setAttribute('value', rule.action || 'archive');
      menulistAction.setAttribute('rule', 'action');

      const menulistSrc = doc.createElementNS(XUL, 'menulist');
      const menupopupSrc = doc.createElementNS(XUL, 'menupopup', { is: 'folder-menupopup' });
      menulistSrc.insertBefore(menupopupSrc, null);
      menulistSrc.value = rule.src || '';
      menulistSrc.setAttribute('rule', 'src');

      const menulistSub = doc.createElementNS(XUL, 'menulist');
      const menupopupSub = doc.createElementNS(XUL, 'menupopup');
      const types = [{ key: self.strBundle.GetStringFromName('perfdialog.type.only'), value: 0 }, { key: self.strBundle.GetStringFromName('perfdialog.type.sub'), value: 1 }, { key: self.strBundle.GetStringFromName('perfdialog.type.sub_keep'), value: 2 }];
      types.forEach(function (type) {
        const menuitem = doc.createElementNS(XUL, 'menuitem');
        menuitem.setAttribute('label', type.key);
        menuitem.setAttribute('value', type.value);
        menupopupSub.insertBefore(menuitem, null);
      });
      menulistSub.insertBefore(menupopupSub, null);
      menulistSub.setAttribute('value', rule.sub || 0);
      menulistSub.setAttribute('rule', 'sub');
      menulistSub.setAttribute('tooltiptext', self.strBundle.GetStringFromName('perfdialog.tooltip.scope'));

      const menulistDest = doc.createElementNS(XUL, 'menulist');
      const menupopupDest = doc.createElementNS(XUL, 'menupopup', { is: 'folder-menupopup' });
      menulistDest.insertBefore(menupopupDest, null);
      menulistDest.value = rule.dest || '';
      menulistDest.setAttribute('rule', 'dest');

      const [from, recipient, subject, size, tags, age] = [
        // filter, size, default, tooltip, type, min
        ['from', '10', '', perfDialogTooltipID],
        ['recipient', '10', '', perfDialogTooltipID],
        ['subject', '', '', perfDialogTooltipID],
        ['size', '5', '', perfDialogTooltipID],
        ['tags', '10', '', perfDialogTooltipID],
        ['age', '4', autoArchivePref.options.default_days, perfDialogAgeTooltipID, 'number', '0']].map(function (attributes) {
        const element = doc.createElementNS(XUL, 'html:input');
        const [filter, size, defaultValue, tooltip, type, min] = attributes;
        element.setAttribute('rule', filter);
        if (size) element.setAttribute('size', size);
        element.setAttribute('value', typeof (rule[filter]) !== 'undefined' ? rule[filter] : defaultValue);
        if (tooltip) element.tooltip = tooltip;
        if (type) element.setAttribute('type', type);
        if (typeof (min) !== 'undefined') element.setAttribute('min', '0');
        const preference = Preferences.get('extensions.awsome_auto_archive.show_' + filter);
        const actualValue = preference.value !== undefined ? preference.value : preference.defaultValue;
        element.style.display = actualValue ? '-moz-box' : 'none';
        return element;
      });

      const [up, down, remove] = [
        ['\u2191', function (aEvent) { self.upDownRule(row, true); }, ''],
        ['\u2193', function (aEvent) { self.upDownRule(row, false); }, ''],
        ['x', function (aEvent) { self.removeRule(row); }, 'awsome_auto_archive-delete-rule']].map(function (attributes) {
        const element = doc.createElementNS(XUL, 'toolbarbutton');
        element.setAttribute('label', attributes[0]);
        element.addEventListener('command', attributes[1], false);
        if (attributes[2]) element.classList.add(attributes[2]);
        return element;
      });

      row.classList.add(ruleClass);
      row.tooltip = perfDialogTooltipID;
      [enable, menulistAction, menulistSrc, menulistSub, menulistDest, from, recipient, subject, size, tags, age, up, down, remove].forEach(function (item) {
        row.insertBefore(item, null);
      });
      container.insertBefore(row, ref);
      self.initFolderPick(menulistSrc, menupopupSrc, true);
      self.initFolderPick(menulistDest, menupopupDest, false);
      self.checkAction(menulistAction, menulistDest, menulistSub);
      self.checkEnable(enable, row);
      menulistAction.addEventListener('command', function (aEvent) { self.checkAction(menulistAction, menulistDest, menulistSub); }, false);
      enable.addEventListener('command', function (aEvent) { self.checkEnable(enable, row); }, false);
      row.addEventListener('focus', function (aEvent) { self.checkFocus(row); }, true);
      row.addEventListener('click', function (aEvent) { self.checkFocus(row); }, true);
      return row;
    } catch (err) {
      autoArchiveLog.logException(err);
    }
  },

  focusRow: null,
  checkFocus: function (row) {
    if (this.focusRow && this.focusRow != row) this.focusRow.removeAttribute('awsome_auto_archive-focused');
    row.setAttribute('awsome_auto_archive-focused', true);
    this.focusRow = row;
  },

  upDownRule: function (row, isUp) {
    try {
      const ref = isUp ? row.previousSibling : row;
      const remove = isUp ? row : row.nextSibling;
      if (ref && remove && ref.classList.contains(ruleClass) && remove.classList.contains(ruleClass)) {
        const rule = this.getOneRule(remove);
        remove.parentNode.removeChild(remove);
        // remove.parentNode.insertBefore(remove, ref); // lost all unsaved values
        const newBox = this.creatOneRule(rule, ref);
        this.checkFocus(isUp ? newBox : row);
        this.syncToPerf(true);
      }
    } catch (err) {
      autoArchiveLog.logException(err);
    }
  },

  removeRule: function (row) {
    row.parentNode.removeChild(row);
    this.syncToPerf(true);
  },

  revertRules: function () {
    if (!this._doc) return;
    this.syncToPerf(true);
    const preference = Preferences.get('extensions.awsome_auto_archive.rules');
    autoArchiveLog.info('Revert rules from\n' + preference.value + '\nto\n' + this._savedRules);
    preference.value = this._savedRules;
  },

  checkEnable: function (enable, row) {
    if (enable.checked) {
      row.classList.remove('awsome_auto_archive-disable');
    } else {
      row.classList.add('awsome_auto_archive-disable');
    }
  },

  checkAction: function (menulistAction, menulistDest, menulistSub) {
    const limit = ['archive', 'delete'].indexOf(menulistAction.value) >= 0;
    if (limit && menulistSub.value == 2) menulistSub.value = 1;
    menulistDest.style.visibility = limit ? 'hidden' : 'visible';
    menulistSub.firstChild.lastChild.style.display = limit ? 'none' : '-moz-box';
  },

  starStopNow: function (dry_run) {
    autoArchiveService.starStopNow(this.getRules(), dry_run);
  },

  statusCallback: function (status, detail) {
    const run_button = self._doc.getElementById('awsome_auto_archive-action');
    const dry_button = self._doc.getElementById('awsome_auto_archive-dry-run');
    if (!run_button || !dry_button) return;
    if ([autoArchiveService.STATUS_SLEEP, autoArchiveService.STATUS_WAITIDLE, autoArchiveService.STATUS_FINISH, autoArchiveService.STATUS_HIBERNATE].indexOf(status) >= 0) {
      // change run_button to "Run"
      run_button.setAttribute('label', self.strBundle.GetStringFromName('perfdialog.action.button.run'));
      dry_button.setAttribute('label', self.strBundle.GetStringFromName('perfdialog.action.button.dryrun'));
    } else if (status == autoArchiveService.STATUS_RUN) {
      // change run_button to "Stop"
      run_button.setAttribute('label', self.strBundle.GetStringFromName('perfdialog.action.button.stop'));
      dry_button.setAttribute('label', self.strBundle.GetStringFromName('perfdialog.action.button.stop'));
    }
    run_button.setAttribute('tooltiptext', detail);
    dry_button.setAttribute('tooltiptext', detail);
  },

  creatNewRule: function (rule) {
    if (!rule) rule = { action: 'archive', enable: true, sub: 0, age: autoArchivePref.options.default_days };
    this.checkFocus(this.creatOneRule(rule, null));
    this.syncToPerf(true);
  },
  changeRule: function (how) {
    if (!this.focusRow) return;
    if (how == 'up') this.upDownRule(this.focusRow, true);
    else if (how == 'down') this.upDownRule(this.focusRow, false);
    else if (how == 'remove') this.removeRule(this.focusRow);
  },

  createRulesBasedOnString: function (value, emptyRule) {
    if (value === self.oldvalue) return;
    self.createRuleHeader();
    const rules = JSON.parse(value);
    if (rules.length) {
      rules.forEach(function (rule) {
        self.creatOneRule(rule, null);
      });
      self.oldvalue = value;
    } else if (typeof (emptyRule) === 'undefined' || emptyRule) self.creatNewRule();
  },
  syncFromPerf: function (win) { // this need 0.5s for 8 rules
    // autoArchiveLog.info('syncFromPerf');
    // if not modal, user can open 2nd pref window, we will close the old one, and close/unLoadPerfWindow seems a sync call, so we are fine
    if (this._win && this._win != win && !this._win.closed) this._win.close();
    this._win = win;
    this._doc = win.document;
    const preference = Preferences.get('extensions.awsome_auto_archive.rules');
    const actualValue = preference.value !== undefined ? preference.value : preference.defaultValue;
    this.createRulesBasedOnString(actualValue, !win.arguments || !win.arguments[0]); // don't create empty rule if loadPerfWindow will create new rule based on selected email
    // autoArchiveLog.info('syncFromPerf done');
  },

  syncToPerf: function (store2pref) { // this need 0.005s for 8 rules
    autoArchiveLog.info('syncToPerf');
    const value = JSON.stringify(this.getRules());
    this.oldvalue = value; // need before set preference.value, which will cause syncFromPref
    autoArchiveLog.info('syncToPerf:' + value);
    if (store2pref) {
      const preference = Preferences.get('extensions.awsome_auto_archive.rules');
      preference.value = value;
      autoArchiveLog.info('preference:' + preference.value);
    }
    // autoArchiveLog.info('syncToPerf done');
    return value;
  },

  bindPerfed: false,
  bindPerf: function () {
    if (this.bindPerfed) return;
    this.bindPerfed = true;
    // Preferences.forceEnableInstantApply();
    Preferences.addAll([
      { id: 'extensions.awsome_auto_archive.rules', type: 'string' },
      { id: 'extensions.awsome_auto_archive.show_folder_as', type: 'int', onchange: 'autoArchivePrefDialog.changeShowFolderAs();' }, // TODO
      { id: 'extensions.awsome_auto_archive.update_statusbartext', type: 'bool' },
      { id: 'extensions.awsome_auto_archive.enable_verbose_info', type: 'bool' },
      { id: 'extensions.awsome_auto_archive.dry_run', type: 'bool' },
      { id: 'extensions.awsome_auto_archive.enable_tag', type: 'bool' },
      { id: 'extensions.awsome_auto_archive.enable_flag', type: 'bool' },
      { id: 'extensions.awsome_auto_archive.enable_unread', type: 'bool' },
      { id: 'extensions.awsome_auto_archive.age_tag', type: 'int' },
      { id: 'extensions.awsome_auto_archive.age_flag', type: 'int' },
      { id: 'extensions.awsome_auto_archive.age_unread', type: 'int' },
      { id: 'extensions.awsome_auto_archive.startup_delay', type: 'int' },
      { id: 'extensions.awsome_auto_archive.idle_delay', type: 'int' },
      { id: 'extensions.awsome_auto_archive.start_next_delay', type: 'int' },
      { id: 'extensions.awsome_auto_archive.rule_timeout', type: 'int' },
      { id: 'extensions.awsome_auto_archive.default_days', type: 'int' },
      { id: 'extensions.awsome_auto_archive.messages_number_limit', type: 'int' },
      { id: 'extensions.awsome_auto_archive.messages_size_limit', type: 'int' },
      { id: 'extensions.awsome_auto_archive.start_exceed_delay', type: 'int' },
      { id: 'alerts.disableSlidingEffect', type: 'bool' },
      { id: 'extensions.awsome_auto_archive.alert_show_time', type: 'int' },
      { id: 'extensions.awsome_auto_archive.delete_duplicate_in_src', type: 'bool' },
      { id: 'extensions.awsome_auto_archive.ignore_spam_folders', type: 'bool' },
      { id: 'extensions.awsome_auto_archive.generate_rule_use', type: 'int' },
      { id: 'extensions.awsome_auto_archive.show_from', type: 'bool' },
      { id: 'extensions.awsome_auto_archive.show_recipient', type: 'bool' },
      { id: 'extensions.awsome_auto_archive.show_subject', type: 'bool' },
      { id: 'extensions.awsome_auto_archive.show_size', type: 'bool' },
      { id: 'extensions.awsome_auto_archive.show_tags', type: 'bool' },
      { id: 'extensions.awsome_auto_archive.show_age', type: 'bool' }
    ]);
    // hack to enable isElementEditable, https://bugzilla.mozilla.org/show_bug.cgi?id=1557989
    const alerts = Preferences.get('alerts.disableSlidingEffect');
    self.hookedFunctions.push(autoArchiveaop.around({ target: alerts.__proto__, method: 'isElementEditable' }, function (invocation) {
      const aElement = invocation.arguments[0];
      return invocation.proceed() || aElement.getAttribute('preference-editable');
    })[0]);
  },

  PopupShowing: function (event) {
    try {
      const doc = event.view.document;
      const tooltip = doc.getElementById(perfDialogTooltipID);
      const line1 = tooltip.firstChild.nextSibling.firstChild;
      const line2 = line1.nextSibling;
      const line3 = line2.nextSibling;
      const line4 = line3.nextSibling;
      let triggerNode = event.target.triggerNode;
      const rule = triggerNode.getAttribute('rule');
      if (rule == '' || !autoArchiveService.advancedTerms[rule]) { // tooltip for row
        const container = doc.getElementById('awsome_auto_archive-rules');
        if (!container) return;
        let i = 0;
        for (const row of container.childNodes) {
          if (row.contains(triggerNode)) {
            triggerNode = row;
            break;
          }
          i++;
        }
        line1.value = 'Rule ' + i;
        line3.value = JSON.stringify(self.getOneRule(triggerNode));
        line2.value = line4.value = '';
        return true;
      }
      // tooltip for From/To etc
      const supportRE = autoArchiveService.advancedTerms[rule].some(function (term) {
        return MailServices.filters.getCustomTerm(term);
      });
      const str = function (label) { return self.strBundle.GetStringFromName('perfdialog.tooltip.' + label); };
      line1.value = (triggerNode.value == '') ? str('emptyFilter') : triggerNode.value;
      line2.value = supportRE ? str('hasRE') : (['size', 'tags'].indexOf(rule) >= 0 ? str('line2.' + rule) : str('noRE'));
      line3.value = str('line3.' + rule);
      line4.value = str('negativeSearch') + str('negative.' + rule + 'Example');
    } catch (err) {
      autoArchiveLog.logException(err);
    }
    return true;
  },

  syncFromPerf4Filter: function (obj) {
    autoArchiveLog.info('syncFromPerf4Filter:' + obj.getAttribute('preference'));
    const doc = obj.ownerDocument;
    const perfID = obj.getAttribute('preference');
    const preference = Preferences.get(perfID);
    const actualValue = preference.value !== undefined ? preference.value : preference.defaultValue;
    const oldValue = obj.oldValue;
    obj.setAttribute('checked', actualValue);
    if (oldValue != actualValue) {
      obj.setAttribute('checked', actualValue);
      const container = doc.getElementById('awsome_auto_archive-rules');
      if (!container) return;
      for (const row of container.childNodes) {
        for (const item of row.childNodes) {
          const key = item.getAttribute('rule');
          if ('extensions.awsome_auto_archive.show_' + key == perfID) { item.style.display = actualValue ? '-moz-box' : 'none'; }
        }
      }
    }
    obj.oldValue = actualValue;
    return actualValue;
  },

  syncToPerf4Filter: function (obj) {
    // autoArchiveLog.info('syncToPerf4Filter:' + obj.getAttribute("preference"));
    const preference = Preferences.get(obj.getAttribute('preference'));
    preference.value = !!obj.getAttribute('checked');
    return preference.value;
  },

  loadPerfWindow: function (win) {
    try {
      autoArchiveLog.info('loadPerfWindow');
      if (!this._win) this.syncFromPerf(win); // SeaMonkey may have one dialog open from addon manager, and then open another one from icon or context menu
      this.instantApply = Preferences.instantApply || false;
      autoArchivePref.setInstantApply(this.instantApply);
      if (this.instantApply) { // only use synctopreference for instantApply, else use acceptPerfWindow
        // must be a onsynctopreference attribute, not a event handler, ref preferences.xml
        this._doc.getElementById('awsome_auto_archive-rules').setAttribute('onsynctopreference', 'return autoArchivePrefDialog.syncToPerf(true);');
        // no need to show 'Apply' button for instantApply
        const extra1 = this._doc.documentElement.getButton('extra1');
        if (extra1 && extra1.parentNode) extra1.parentNode.removeChild(extra1);
      }
      autoArchiveService.addStatusListener(this.statusCallback);
      this.fillIdentities(false);
      const tooltip = this._doc.getElementById(perfDialogTooltipID);
      if (tooltip) tooltip.addEventListener('popupshowing', this.PopupShowing, true);
      this._savedRules = autoArchivePref.options.rules;
      if (win.arguments && win.arguments[0]) { // new rule based on message selected, not including in the revert all
        const msgHdr = win.arguments[0];
        this.creatNewRule({
          action: 'archive',
          enable: true,
          src: msgHdr.folder.URI,
          sub: 0,
          from: this.getSearchStringFromAddress(msgHdr.mime2DecodedAuthor),
          recipient: this.getSearchStringFromAddress(msgHdr.mime2DecodedTo || msgHdr.mime2DecodedRecipients),
          subject: msgHdr.mime2DecodedSubject,
          age: autoArchivePref.options.default_days
        });
      }

      ['extensions.awsome_auto_archive.show_from',
        'extensions.awsome_auto_archive.show_recipient',
        'extensions.awsome_auto_archive.show_subject',
        'extensions.awsome_auto_archive.show_size',
        'extensions.awsome_auto_archive.show_tags',
        'extensions.awsome_auto_archive.show_age'].forEach(
        function (prefId) {
          Preferences.addSyncFromPrefListener(
            document.getElementById(prefId),
            (obj) => syncFromPerf4Filter(obj));

          Preferences.addSyncToPrefListener(
            document.getElementById(prefId),
            (obj) => syncToPerf4Filter(obj));
        });
    } catch (err) { autoArchiveLog.logException(err); }
    return true;
  },
  getSearchStringFromAddress: function (mails) {
    // GetDisplayNameInAddressBook() in http://mxr.mozilla.org/comm-central/source/mailnews/base/src/nsMsgDBView.cpp
    try {
      const parsedMails = GlodaUtils.parseMailAddresses(mails);
      const returnMails = [];
      for (let i = 0; i < parsedMails.count; i++) {
        let email = parsedMails.addresses[i]; let card; let displayName;
        if (!autoArchivePref.options.generate_rule_use) {
          if (Services.prefs.getBoolPref('mail.showCondensedAddresses')) { // the usage of getSearchStringFromAddress might be few, so won't add Observer
            const allAddressBooks = MailServices.ab.directories;
            while (!card && allAddressBooks.hasMoreElements()) {
              const addressBook = allAddressBooks.getNext().QueryInterface(Ci.nsIAbDirectory);
              if (addressBook instanceof Ci.nsIAbDirectory /* && !addressBook.isRemote */) {
                try {
                  card = addressBook.cardForEmailAddress(email); // case-insensitive && sync, only return 1st one if multiple match, but it search on all email addresses
                } catch (err) {}
                if (card) {
                  const PreferDisplayName = Number(card.getProperty('PreferDisplayName', 1));
                  if (PreferDisplayName) displayName = card.displayName;
                }
              }
            }
          }
          if (!displayName) displayName = parsedMails.names[i] || parsedMails.fullAddresses[i];
          displayName = displayName.replace(/['"<>]/g, '');
          if (parsedMails.fullAddresses[i].indexOf(displayName) != -1) email = displayName;
        }
        const search = (autoArchivePref.options.generate_rule_use == 2) ? email : email.replace(/(.*@).*/, '$1');
        if (returnMails.indexOf(search) < 0) returnMails.push(search);
      }
      return returnMails.join(', ');
    } catch (err) { autoArchiveLog.logException(err); }
    return mails;
  },

  getOneRule: function (row) {
    const rule = {};
    for (const item of row.childNodes) {
      const key = item.getAttribute('rule');
      if (key) {
        let value = item.value || item.checked;
        if (item.getAttribute('type') == 'number' && typeof (item.valueNumber) !== 'undefined') value = item.valueNumber;
        if (key == 'sub') value = Number(value); // menulist.value is always 'string'
        rule[key] = value;
      }
    }
    return rule;
  },

  getRules: function () {
    const rules = [];
    try {
      const container = this._doc.getElementById('awsome_auto_archive-rules');
      if (!container) return rules;
      for (const row of container.childNodes) {
        if (row.classList.contains(ruleClass)) {
          const rule = this.getOneRule(row);
          if (Object.keys(rule).length > 0) rules.push(rule);
        }
      }
      // autoArchiveLog.logObject(rules,'got rules',1);
    } catch (err) { autoArchiveLog.logException(err); throw err; } // throw the error out so syncToPerf won't get an empty rules
    return rules;
  },

  acceptPerfWindow: function () {
    try {
      autoArchiveLog.info('acceptPerfWindow');
      if (!this.instantApply) autoArchivePref.setPerf('rules', this.syncToPerf());
    } catch (err) { autoArchiveLog.logException(err); }
    return true;
  },
  unLoadPerfWindow: function () {
    if (!autoArchiveService || !autoArchivePref || !autoArchiveLog || !autoArchiveUtil) return true;
    self.hookedFunctions.forEach(function (hooked) {
      hooked.unweave();
    });
    if (this._savedRules != autoArchivePref.options.rules) autoArchiveUtil.backupRules(autoArchivePref.options.rules, autoArchivePref.options.rules_to_keep);
    autoArchiveService.removeStatusListener(this.statusCallback);
    const tooltip = this._doc.getElementById(perfDialogTooltipID);
    if (tooltip) tooltip.removeEventListener('popupshowing', this.PopupShowing, true);
    if (this.instantApply) autoArchivePref.validateRules();
    delete this._doc;
    delete this._win;
    delete this.oldvalue;
    delete this.instantApply;
    delete self.hookedFunctions;
    autoArchiveLog.info('prefwindow unload');
    return true;
  },

  // https://github.com/protz/thunderbird-stdlib/blob/master/misc.js
  fillIdentities: function (aSkipNntp) {
    const doc = self._doc;
    const group = doc.getElementById('awsome_auto_archive-IDs');
    const pane = doc.getElementById('awsome_auto_archive-perfpane');
    const tabbox = doc.getElementById('awsome_auto_archive-tabbox');
    if (!group || !pane || !tabbox) return;
    let firstNonNull = null; const gIdentities = {}; const gAccounts = {};
    for (const account of fixIterator(MailServices.accounts.accounts, Ci.nsIMsgAccount)) {
      const server = account.incomingServer;
      if (aSkipNntp && (!server || server.type != 'pop3' && server.type != 'imap')) {
        continue;
      }
      for (const id of fixIterator(account.identities, Ci.nsIMsgIdentity)) {
        // We're only interested in identities that have a real email.
        if (id.email) {
          gIdentities[id.email.toLowerCase()] = id;
          gAccounts[id.email.toLowerCase()] = account;
          if (!firstNonNull) firstNonNull = id;
        }
      }
    }
    gIdentities.default = MailServices.accounts.defaultAccount.defaultIdentity || firstNonNull;
    gAccounts.default = MailServices.accounts.defaultAccount;
    Object.keys(gIdentities).sort().forEach(function (id) {
      const button = doc.createElementNS(XUL, 'button');
      button.setAttribute('label', id);
      button.addEventListener('command', function (aEvent) { self._win.openDialog('chrome://messenger/content/am-identity-edit.xul', 'dlg', '', { identity: gIdentities[id], account: gAccounts[id], result: false }); }, false);
      group.insertBefore(button, null);
    });
    pane.style.minHeight = pane.clientHeight + 10 + 'px'; // reset the pane height after fill Identities, to prevent vertical scrollbar

    try {
      const perfDialog = self._doc.getElementById('awsome_auto_archive-prefs');
      const buttonBox = perfDialog;
      let targetWinHeight = buttonBox.scrollHeight + pane.clientHeight;
      if (targetWinHeight > this._win.screen.availHeight) targetWinHeight = this._win.screen.availHeight;
      const currentWinHeight = perfDialog.height;
      if (currentWinHeight < targetWinHeight + 62) perfDialog.setAttribute('height', targetWinHeight + 62);
      const width = Number(perfDialog.width || perfDialog.getAttribute('width'));
      const targetWidth = Number(tabbox.clientWidth || tabbox.scrollWidth) + 36;
      if (width < targetWidth) perfDialog.setAttribute('width', targetWidth);
    } catch (err) { autoArchiveLog.logException(err); }
  },

  applyChanges: function () {
    Array.prototype.forEach.call(self._doc.getElementById('awsome_auto_archive-prefs').preferencePanes, function (pane) {
      pane.writePreferences(true);
    });
  }

};

const self = autoArchivePrefDialog;
