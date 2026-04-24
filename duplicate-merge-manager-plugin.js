DuplicateMergeManagerPlugin = {
  id: null,
  version: null,
  rootURI: null,
  initialized: false,
  windowStates: new WeakMap(),
  config: {
    menuID: "duplicate-merge-manager-itemmenu",
    menuLabel: "查找可能重复条目...",
    batchMenuLabel: "查找可能重复条目...",
    pathSeparator: " / ",
    unnamedCollectionLabel: "未命名文件夹",
    noCollectionsLabel: "未加入任何文件夹",
    minCandidateScore: 68,
    maxCandidates: 50,
    maxIndexedTitleTokens: 18,
    maxPrefilterCandidates: 500,
    titleTokenMinLength: 3,
    titleSimilarityFloor: 0.76,
    exactTitleScore: 92,
    mergeActionLabel: "合并",
    trashActionLabel: "回收"
  },

  init({ id, version, rootURI }) {
    if (this.initialized) {
      return;
    }
    this.id = id;
    this.version = version;
    this.rootURI = rootURI;
    this.initialized = true;
  },

  log(message) {
    Zotero.debug(`Duplicate Merge Manager: ${message}`);
  },

  addToAllWindows() {
    for (const window of Zotero.getMainWindows()) {
      if (!window.ZoteroPane) {
        continue;
      }
      this.addToWindow(window);
    }
  },

  removeFromAllWindows() {
    for (const window of Zotero.getMainWindows()) {
      if (!window.ZoteroPane) {
        continue;
      }
      this.removeFromWindow(window);
    }
  },

  addToWindow(window) {
    if (this.windowStates.has(window)) {
      return;
    }

    const itemMenuPopup = window.document.querySelector("#zotero-itemmenu");
    if (!itemMenuPopup) {
      this.log("Zotero item context menu was not found");
      return;
    }

    const menuItem = this.createXULElement(window.document, "menuitem");
    menuItem.setAttribute("id", this.config.menuID);
    menuItem.setAttribute("label", this.config.menuLabel);
    menuItem.hidden = true;

    const state = {
      itemMenuPopup,
      menuItem,
      onPopupShowing: null,
      onMenuCommand: null,
      panel: null
    };

    const plugin = this;
    state.onMenuCommand = function (commandEvent) {
      commandEvent.stopPropagation();
      plugin.openDuplicatePanel(window);
    };
    menuItem.addEventListener("command", state.onMenuCommand);

    state.onPopupShowing = function (event) {
      if (event.target !== itemMenuPopup) {
        return;
      }
      const items = plugin.getSelectedTopLevelRegularItems(window);
      menuItem.hidden = !items.length;
      menuItem.disabled = !items.length;
      menuItem.setAttribute("label", items.length > 1
        ? `${plugin.config.batchMenuLabel}（${items.length} 项）`
        : plugin.config.menuLabel
      );
    };

    itemMenuPopup.addEventListener("popupshowing", state.onPopupShowing);
    itemMenuPopup.appendChild(menuItem);
    this.windowStates.set(window, state);
  },

  removeFromWindow(window) {
    const state = this.windowStates.get(window);
    if (!state) {
      return;
    }

    if (state.itemMenuPopup && state.onPopupShowing) {
      state.itemMenuPopup.removeEventListener("popupshowing", state.onPopupShowing);
    }

    if (state.menuItem) {
      if (state.onMenuCommand) {
        state.menuItem.removeEventListener("command", state.onMenuCommand);
      }
      state.menuItem.remove();
    }

    if (state.panel && state.panel.overlay) {
      state.panel.overlay.remove();
    }

    this.windowStates.delete(window);
  },

  createXULElement(document, tagName) {
    if (typeof document.createXULElement === "function") {
      return document.createXULElement(tagName);
    }
    return document.createElement(tagName);
  },

  createHTMLElement(document, tagName) {
    return document.createElementNS("http://www.w3.org/1999/xhtml", tagName);
  },

  clearElement(element) {
    while (element.firstChild) {
      element.firstChild.remove();
    }
  },

  getSelectedTopLevelRegularItems(window) {
    const candidates = [];
    const pushItems = (values) => {
      if (!Array.isArray(values)) {
        return;
      }
      for (const value of values) {
        if (value && typeof value === "object") {
          candidates.push(value);
        }
      }
    };

    const readSelection = (label, callback) => {
      try {
        pushItems(callback());
      } catch (error) {
        this.log(`Unable to read selected items via ${label}: ${error}`);
      }
    };

    readSelection("ZoteroPane.getSelectedObjects", () => window.ZoteroPane.getSelectedObjects());
    readSelection("ZoteroPane.getSelectedItems", () => window.ZoteroPane.getSelectedItems());

    if (window.ZoteroPane.itemsView) {
      readSelection("itemsView.getSelectedObjects", () => window.ZoteroPane.itemsView.getSelectedObjects());
      readSelection("itemsView.getSelectedItems", () => window.ZoteroPane.itemsView.getSelectedItems());

      const selection = window.ZoteroPane.itemsView.selection;
      if (selection && selection.selected && typeof window.ZoteroPane.itemsView.getRow === "function") {
        readSelection("itemsView.selection.selected", () => {
          const values = [];
          for (const index of selection.selected) {
            const row = window.ZoteroPane.itemsView.getRow(index);
            if (row && row.ref) {
              values.push(row.ref);
            }
          }
          return values;
        });
      }
    }

    const selectedItems = [];
    const seenItemIDs = new Set();

    for (const item of candidates) {
      const resolvedItem = this.resolveTopLevelRegularItem(item);
      if (!resolvedItem) {
        continue;
      }
      if (resolvedItem.deleted) {
        continue;
      }
      if (seenItemIDs.has(resolvedItem.id)) {
        continue;
      }
      seenItemIDs.add(resolvedItem.id);
      selectedItems.push(resolvedItem);
    }

    return selectedItems;
  },

  resolveTopLevelRegularItem(item) {
    if (!item || typeof item !== "object") {
      return null;
    }
    if (typeof item.isRegularItem === "function"
      && typeof item.isTopLevelItem === "function"
      && item.isRegularItem()
      && item.isTopLevelItem()) {
      return item;
    }

    const parentID = item.parentItemID || item.parentID;
    if (!parentID) {
      return null;
    }

    const parentItem = Zotero.Items.get(parentID);
    if (parentItem
      && typeof parentItem.isRegularItem === "function"
      && typeof parentItem.isTopLevelItem === "function"
      && parentItem.isRegularItem()
      && parentItem.isTopLevelItem()) {
      return parentItem;
    }

    return null;
  },

  async openDuplicatePanel(window, itemIDs) {
    const state = this.windowStates.get(window);
    if (!state) {
      return;
    }

    if (state.panel && state.panel.overlay) {
      state.panel.overlay.remove();
    }

    const selectedItems = Array.isArray(itemIDs) && itemIDs.length
      ? itemIDs.map((itemID) => Zotero.Items.get(itemID)).filter(Boolean)
      : this.getSelectedTopLevelRegularItems(window);
    const selectedItemIDs = selectedItems.map((item) => item.id);

    const panel = {
      itemIDs: selectedItemIDs,
      overlay: this.createHTMLElement(window.document, "div"),
      statusText: "正在扫描当前文库..."
    };
    panel.overlay.setAttribute("class", "dmm-overlay");
    panel.overlay.addEventListener("keydown", (event) => {
      if (event.key === "Escape") {
        event.stopPropagation();
        this.closeDuplicatePanel(window);
      }
    });

    state.panel = panel;
    window.document.documentElement.appendChild(panel.overlay);
    this.renderDuplicatePanel(window, panel, null);

    try {
      const data = await this.buildDuplicatePanelData(selectedItemIDs);
      panel.statusText = "";
      this.renderDuplicatePanel(window, panel, data);
    } catch (error) {
      Zotero.logError(error);
      panel.statusText = error && error.message ? error.message : String(error);
      this.renderDuplicatePanel(window, panel, null);
    }
  },

  closeDuplicatePanel(window) {
    const state = this.windowStates.get(window);
    if (!state || !state.panel) {
      return;
    }
    if (state.panel.overlay) {
      state.panel.overlay.remove();
    }
    state.panel = null;
  },

  async buildDuplicatePanelData(itemIDs) {
    const targetItems = [];

    for (const itemID of itemIDs) {
      const item = Zotero.Items.get(itemID);
      if (item && !item.deleted && item.isRegularItem() && item.isTopLevelItem()) {
        targetItems.push(item);
      }
    }

    if (!targetItems.length) {
      throw new Error("当前选择不包含顶层文献条目");
    }

    const libraries = [...new Set(targetItems.map((item) => item.libraryID))];
    const indexesByLibrary = new Map();
    for (const libraryID of libraries) {
      indexesByLibrary.set(libraryID, await this.buildLibraryDuplicateIndex(libraryID));
    }

    const groups = [];
    for (const targetItem of targetItems) {
      const target = this.createItemSummary(targetItem);
      const candidates = [];
      const index = indexesByLibrary.get(targetItem.libraryID);
      const excludeIDs = new Set([targetItem.id]);
      const prefilteredCandidates = this.mergeCandidateLists([
        this.findExactTitleCandidates(target, index, excludeIDs),
        this.findSelectedExactTitleCandidates(target, targetItems, excludeIDs),
        this.findPrefilteredCandidates(target, index, excludeIDs)
      ]);

      for (const candidate of prefilteredCandidates) {
        const match = this.scoreDuplicateCandidate(target, candidate);
        if (match.score >= this.config.minCandidateScore) {
          candidates.push({
            item: candidate,
            score: match.score,
            reasons: match.reasons
          });
        }
      }

      candidates.sort((left, right) => {
        if (right.score !== left.score) {
          return right.score - left.score;
        }
        return left.item.title.localeCompare(right.item.title);
      });

      groups.push({
        target,
        candidates: candidates.slice(0, this.config.maxCandidates)
      });
    }

    return {
      targets: groups.map((group) => group.target),
      selectedExactTitleGroups: this.buildSelectedExactTitleGroups(targetItems),
      groups
    };
  },

  buildSelectedExactTitleGroups(items) {
    const byTitle = new Map();
    for (const item of items) {
      const summary = this.createItemSummary(item);
      const groupKey = summary.titleCompactKey || summary.titleKey;
      if (!groupKey) {
        continue;
      }
      if (!byTitle.has(groupKey)) {
        byTitle.set(groupKey, []);
      }
      byTitle.get(groupKey).push(summary);
    }

    return [...byTitle.values()]
      .filter((group) => group.length > 1)
      .map((itemsInGroup) => ({
        title: itemsInGroup[0].title || "(无标题)",
        items: itemsInGroup.sort((left, right) => left.id - right.id)
      }));
  },

  findExactTitleCandidates(target, index, excludeIDs) {
    if (!target.titleKey && !target.titleCompactKey) {
      return [];
    }
    const candidates = [];
    const seen = new Set();
    const add = (entries) => {
      for (const entry of entries || []) {
        if (!entry || excludeIDs.has(entry.id) || seen.has(entry.id)) {
          continue;
        }
        if (this.hasExactTitleMatch(target, entry)) {
          seen.add(entry.id);
          candidates.push(entry);
        }
      }
    };
    add(index.byTitleKey.get(target.titleKey));
    add(index.byTitleCompactKey.get(target.titleCompactKey));
    return candidates;
  },

  findSelectedExactTitleCandidates(target, selectedItems, excludeIDs) {
    if (!target.titleKey && !target.titleCompactKey) {
      return [];
    }
    const candidates = [];
    for (const item of selectedItems) {
      if (!item || excludeIDs.has(item.id) || item.libraryID !== target.libraryID) {
        continue;
      }
      const candidate = this.createItemSummary(item);
      if (this.hasExactTitleMatch(target, candidate)) {
        candidates.push(candidate);
      }
    }
    return candidates;
  },

  mergeCandidateLists(lists) {
    const merged = [];
    const seen = new Set();
    for (const list of lists) {
      for (const candidate of list || []) {
        if (!candidate || seen.has(candidate.id)) {
          continue;
        }
        seen.add(candidate.id);
        merged.push(candidate);
      }
    }
    return merged;
  },

  async buildLibraryDuplicateIndex(libraryID) {
    const items = (await Zotero.Items.getAll(libraryID, true, false)) || [];
    const index = {
      entries: [],
      byDOI: new Map(),
      byTitleKey: new Map(),
      byTitleCompactKey: new Map(),
      byYear: new Map(),
      byCreator: new Map(),
      byTitleToken: new Map()
    };

    const addToMap = (map, key, entry) => {
      if (!key) {
        return;
      }
      if (!map.has(key)) {
        map.set(key, []);
      }
      map.get(key).push(entry);
    };

    for (const item of items) {
      if (!item || item.deleted || !item.isRegularItem() || !item.isTopLevelItem()) {
        continue;
      }

      const entry = this.createItemSummary(item);
      index.entries.push(entry);

      for (const doi of entry.doiValues) {
        addToMap(index.byDOI, doi, entry);
      }

      addToMap(index.byTitleKey, entry.titleKey, entry);
      addToMap(index.byTitleCompactKey, entry.titleCompactKey, entry);
      addToMap(index.byYear, entry.year, entry);

      for (const creator of entry.creators.slice(0, 3)) {
        addToMap(index.byCreator, creator, entry);
      }

      for (const token of this.getIndexedTitleTokens(entry.titleKey)) {
        addToMap(index.byTitleToken, token, entry);
      }
    }

    return index;
  },

  findPrefilteredCandidates(target, index, excludeIDs) {
    const weighted = new Map();
    const addEntries = (entries, weight) => {
      for (const entry of entries || []) {
        if (!entry || excludeIDs.has(entry.id)) {
          continue;
        }
        weighted.set(entry.id, (weighted.get(entry.id) || 0) + weight);
      }
    };

    if (target.titleKey) {
      addEntries(index.byTitleKey.get(target.titleKey), 1200);
    }
    if (target.titleCompactKey && target.titleCompactKey !== target.titleKey) {
      addEntries(index.byTitleCompactKey.get(target.titleCompactKey), 1200);
    }
    for (const token of this.getIndexedTitleTokens(target.titleKey)) {
      addEntries(index.byTitleToken.get(token), 80);
    }
    if (target.doiKey) {
      addEntries(index.byDOI.get(target.doiKey), 160);
    }
    for (const creator of target.creators.slice(0, 3)) {
      addEntries(index.byCreator.get(creator), 12);
    }
    if (target.year) {
      addEntries(index.byYear.get(target.year), 4);
    }

    const ids = [...weighted.entries()]
      .sort((left, right) => right[1] - left[1])
      .slice(0, this.config.maxPrefilterCandidates)
      .map(([id]) => id);

    const byID = new Map(index.entries.map((entry) => [entry.id, entry]));
    return ids.map((id) => byID.get(id)).filter(Boolean);
  },

  getIndexedTitleTokens(titleKey) {
    const tokens = new Set();
    for (const token of String(titleKey || "").split(/\s+/).filter(Boolean)) {
      if (token.length >= this.config.titleTokenMinLength && !/^\d+$/.test(token)) {
        tokens.add(token);
      }
      for (const sequence of token.match(/[\u4e00-\u9fff]{3,}/g) || []) {
        for (const ngram of this.createCharacterNgrams(sequence, 3)) {
          tokens.add(ngram);
        }
      }
    }

    return [...tokens]
      .sort((left, right) => {
        const lengthDiff = right.length - left.length;
        if (lengthDiff) {
          return lengthDiff;
        }
        return left.localeCompare(right);
      })
      .slice(0, this.config.maxIndexedTitleTokens);
  },

  createItemSummary(item) {
    const title = item.getField("title") || "";
    const doiValues = this.getItemDOIs(item);
    return {
      id: item.id,
      title,
      titleKey: this.normalizeTitle(title),
      titleCompactKey: this.normalizeCompactTitle(title),
      doiValues,
      doiKey: doiValues[0] || "",
      year: this.extractYear(item.getField("date") || ""),
      creators: this.getItemCreatorKeys(item),
      itemType: item.itemType || "",
      dateAdded: item.dateAdded || "",
      collections: this.getItemCollectionLocations(item),
      attachmentCount: this.countChildIDs(item, "getAttachments"),
      noteCount: this.countChildIDs(item, "getNotes"),
      tagCount: (typeof item.getTags === "function" ? item.getTags() : []).length
    };
  },

  countChildIDs(item, methodName) {
    if (typeof item[methodName] !== "function") {
      return 0;
    }
    try {
      return item[methodName](true).length;
    } catch (error) {
      this.log(`Unable to count children for item ${item.id}: ${error}`);
      return 0;
    }
  },

  getItemDOIs(item) {
    const values = new Set();
    const push = (candidate) => {
      const direct = this.normalizeDOI(candidate);
      if (direct) {
        values.add(direct);
      }
      const text = String(candidate || "");
      const matches = text.match(/\b10\.\d{4,9}\/[-._;()/:A-Z0-9]+/ig) || [];
      for (const match of matches) {
        const doi = this.normalizeDOI(match);
        if (doi) {
          values.add(doi);
        }
      }
    };

    push(item.getField("DOI"));
    if (typeof item.getExtraField === "function") {
      push(item.getExtraField("DOI"));
    }
    push(item.getField("extra"));
    push(item.getField("url"));
    return [...values];
  },

  normalizeDOI(value) {
    const raw = String(value || "").trim();
    if (!raw) {
      return "";
    }
    const cleaned = raw
      .replace(/^https?:\/\/(?:dx\.)?doi\.org\//i, "")
      .replace(/^doi:\s*/i, "")
      .replace(/[)\].,;。\s]+$/g, "");
    return cleaned.toLowerCase();
  },

  normalizeTitle(value) {
    let text = String(value || "").normalize("NFKC").replace(/<[^>]+>/g, " ");
    if (Zotero.Utilities && typeof Zotero.Utilities.removeDiacritics === "function") {
      text = Zotero.Utilities.removeDiacritics(text);
    }
    return text
      .toLowerCase()
      .replace(/[^0-9a-z\u4e00-\u9fff]+/g, " ")
      .trim()
      .replace(/\s+/g, " ");
  },

  normalizeCompactTitle(value) {
    return this.normalizeTitle(value).replace(/\s+/g, "");
  },

  hasExactTitleMatch(left, right) {
    if (!left || !right) {
      return false;
    }
    if (left.titleKey && right.titleKey && left.titleKey === right.titleKey) {
      return true;
    }
    return Boolean(
      left.titleCompactKey
      && right.titleCompactKey
      && left.titleCompactKey === right.titleCompactKey
    );
  },

  extractYear(value) {
    const match = String(value || "").match(/\b(1[5-9]\d{2}|20\d{2}|21\d{2})\b/);
    return match ? match[1] : "";
  },

  getItemCreatorKeys(item) {
    const creators = typeof item.getCreators === "function" ? item.getCreators() : [];
    return creators
      .map((creator) => this.normalizeTitle(creator.lastName || creator.name || creator.firstName || ""))
      .filter(Boolean);
  },

  scoreDuplicateCandidate(target, candidate) {
    const reasons = [];
    let score = 0;
    const doiMatched = Boolean(target.doiKey && candidate.doiValues.includes(target.doiKey));
    const exactTitleMatched = this.hasExactTitleMatch(target, candidate);
    let titleSimilarity = 0;

    if (target.titleKey && candidate.titleKey) {
      if (exactTitleMatched) {
        titleSimilarity = 1;
        score += this.config.exactTitleScore;
        reasons.push("标题完全相同");
      } else {
        titleSimilarity = this.calculateTitleSimilarity(target.titleKey, candidate.titleKey);
        if (titleSimilarity >= 0.92) {
          score += 76;
          reasons.push(`标题高度相似 ${Math.round(titleSimilarity * 100)}%`);
        } else if (titleSimilarity >= 0.84) {
          score += 68;
          reasons.push(`标题相似 ${Math.round(titleSimilarity * 100)}%`);
        } else if (titleSimilarity >= this.config.titleSimilarityFloor) {
          score += 56;
          reasons.push(`标题部分相似 ${Math.round(titleSimilarity * 100)}%`);
        } else if (!doiMatched) {
          return { score: 0, reasons: [] };
        }
      }
    } else if (!doiMatched) {
      return { score: 0, reasons: [] };
    }

    if (doiMatched) {
      score += titleSimilarity >= this.config.titleSimilarityFloor || !target.titleKey || !candidate.titleKey ? 14 : 8;
      reasons.push("DOI 完全相同");
    }

    if (target.year && candidate.year) {
      if (target.year === candidate.year) {
        score += 8;
        reasons.push("年份相同");
      } else {
        score -= 8;
        reasons.push("年份不同");
      }
    }

    const creatorOverlap = this.countCreatorOverlap(target.creators, candidate.creators);
    if (creatorOverlap > 0) {
      score += target.creators[0] && target.creators[0] === candidate.creators[0] ? 8 : 5;
      reasons.push("作者重合");
    } else if (target.creators.length && candidate.creators.length) {
      score -= 4;
      reasons.push("作者未重合");
    }

    if (target.itemType && candidate.itemType && target.itemType === candidate.itemType) {
      score += 2;
    }

    return {
      score: Math.max(0, Math.min(100, score)),
      reasons
    };
  },

  calculateTitleSimilarity(left, right) {
    if (left === right) {
      return 1;
    }

    const leftTokens = new Set(left.split(/\s+/).filter(Boolean));
    const rightTokens = new Set(right.split(/\s+/).filter(Boolean));
    let tokenSimilarity = 0;

    if (leftTokens.size && rightTokens.size) {
      let overlap = 0;
      for (const token of leftTokens) {
        if (rightTokens.has(token)) {
          overlap += 1;
        }
      }
      tokenSimilarity = overlap / Math.max(leftTokens.size, rightTokens.size);
    }

    const leftCompact = left.replace(/\s+/g, "");
    const rightCompact = right.replace(/\s+/g, "");
    const ngramSize = Math.min(3, leftCompact.length, rightCompact.length);
    const ngramSimilarity = ngramSize ? this.calculateNgramSimilarity(leftCompact, rightCompact, ngramSize) : 0;

    return Math.max(tokenSimilarity, ngramSimilarity);
  },

  calculateNgramSimilarity(left, right, size) {
    const leftNgrams = new Set(this.createCharacterNgrams(left, size));
    const rightNgrams = new Set(this.createCharacterNgrams(right, size));
    if (!leftNgrams.size || !rightNgrams.size) {
      return 0;
    }
    let overlap = 0;
    for (const token of leftNgrams) {
      if (rightNgrams.has(token)) {
        overlap += 1;
      }
    }
    return overlap / Math.max(leftNgrams.size, rightNgrams.size);
  },

  createCharacterNgrams(value, size) {
    const text = String(value || "");
    if (!text) {
      return [];
    }
    if (text.length <= size) {
      return [text];
    }
    const ngrams = [];
    for (let index = 0; index <= text.length - size; index += 1) {
      ngrams.push(text.slice(index, index + size));
    }
    return ngrams;
  },

  countCreatorOverlap(left, right) {
    const rightSet = new Set(right);
    let count = 0;
    for (const creator of left) {
      if (rightSet.has(creator)) {
        count += 1;
      }
    }
    return count;
  },

  getItemCollectionLocations(item) {
    const collectionIDs = [...new Set(item.getCollections() || [])];
    const locations = [];
    for (const collectionID of collectionIDs) {
      const location = this.createCollectionLocation(collectionID);
      if (location) {
        locations.push(location);
      }
    }
    return locations.sort((left, right) => left.path.localeCompare(right.path));
  },

  createCollectionLocation(collectionID) {
    const collection = Zotero.Collections.get(collectionID);
    if (!collection || collection.deleted) {
      return null;
    }

    const names = [];
    const seen = new Set();
    let current = collection;

    while (current && !seen.has(current.id)) {
      seen.add(current.id);
      names.unshift(current.name || this.config.unnamedCollectionLabel);
      if (!current.parentID) {
        break;
      }
      current = Zotero.Collections.get(current.parentID);
      if (current && current.deleted) {
        break;
      }
    }

    return {
      collectionID: collection.id,
      path: names.join(this.config.pathSeparator)
    };
  },

  renderDuplicatePanel(window, panel, data) {
    const document = window.document;
    this.clearElement(panel.overlay);
    this.appendPanelStyles(document, panel.overlay);

    const box = this.createHTMLElement(document, "div");
    box.setAttribute("class", "dmm-panel");
    panel.overlay.appendChild(box);

    const header = this.createHTMLElement(document, "div");
    header.setAttribute("class", "dmm-header");
    box.appendChild(header);

    const titleBox = this.createHTMLElement(document, "div");
    header.appendChild(titleBox);

    const title = this.createHTMLElement(document, "div");
    title.setAttribute("class", "dmm-title");
    title.textContent = "可能重复条目";
    titleBox.appendChild(title);

    const subtitle = this.createHTMLElement(document, "div");
    subtitle.setAttribute("class", "dmm-muted");
    subtitle.textContent = `v${this.version}：以标题相似为主；可确认后合并或移入回收站。`;
    titleBox.appendChild(subtitle);

    const closeButton = this.createHTMLElement(document, "button");
    closeButton.setAttribute("type", "button");
    closeButton.textContent = "关闭";
    closeButton.addEventListener("click", () => this.closeDuplicatePanel(window));
    header.appendChild(closeButton);

    if (!data) {
      const status = this.createHTMLElement(document, "div");
      status.setAttribute("class", "dmm-status");
      status.textContent = panel.statusText || "正在扫描当前文库...";
      box.appendChild(status);
      return;
    }

    const content = this.createHTMLElement(document, "div");
    content.setAttribute("class", "dmm-content");
    box.appendChild(content);

    this.appendPanelSummary(document, content, data);
    this.appendSelectedExactTitleGroups(window, document, content, data.selectedExactTitleGroups || []);
    for (const group of data.groups) {
      this.appendTargetSummary(document, content, group.target);
      this.appendCandidateTable(window, document, content, group);
    }
  },

  appendPanelSummary(document, box, data) {
    const totalCandidates = data.groups.reduce((sum, group) => sum + group.candidates.length, 0);
    const exactTitleGroups = data.selectedExactTitleGroups || [];
    const exactTitleItems = exactTitleGroups.reduce((sum, group) => sum + group.items.length, 0);
    const summary = this.createHTMLElement(document, "div");
    summary.setAttribute("class", "dmm-summary");
    summary.textContent = `查询种子 ${data.groups.length} 项，候选重复 ${totalCandidates} 项，同题目组选中 ${exactTitleGroups.length} 组 / ${exactTitleItems} 项`;
    box.appendChild(summary);
  },

  appendSelectedExactTitleGroups(window, document, box, groups) {
    if (!groups.length) {
      return;
    }

    const heading = this.createHTMLElement(document, "div");
    heading.setAttribute("class", "dmm-summary dmm-selected-duplicates-heading");
    heading.textContent = `当前选中项中题目完全相同：${groups.length} 组`;
    box.appendChild(heading);

    for (const group of groups) {
      const section = this.createHTMLElement(document, "div");
      section.setAttribute("class", "dmm-target dmm-exact-group");
      box.appendChild(section);

      const title = this.createHTMLElement(document, "div");
      title.setAttribute("class", "dmm-target-title dmm-candidate-title");
      title.textContent = `${group.title}（${group.items.length} 项）`;
      section.appendChild(title);

      const recommendedMaster = this.chooseBestMasterSummary(group.items);
      const actionBar = this.createHTMLElement(document, "div");
      actionBar.setAttribute("class", "dmm-actions dmm-group-actions");
      section.appendChild(actionBar);

      const mergeGroupButton = this.createHTMLElement(document, "button");
      mergeGroupButton.setAttribute("type", "button");
      mergeGroupButton.setAttribute("class", "dmm-primary-action");
      mergeGroupButton.textContent = `合并本组（保留 ${recommendedMaster ? recommendedMaster.id : ""}）`;
      mergeGroupButton.disabled = !recommendedMaster || group.items.length < 2;
      mergeGroupButton.addEventListener("click", () => {
        if (!recommendedMaster) {
          return;
        }
        this.mergeSummaries(window, recommendedMaster, group.items.filter((item) => item.id !== recommendedMaster.id), "同题目组");
      });
      actionBar.appendChild(mergeGroupButton);

      const trashOthersButton = this.createHTMLElement(document, "button");
      trashOthersButton.setAttribute("type", "button");
      trashOthersButton.setAttribute("class", "dmm-danger-action");
      trashOthersButton.textContent = "回收非保留项";
      trashOthersButton.disabled = !recommendedMaster || group.items.length < 2;
      trashOthersButton.addEventListener("click", () => {
        if (!recommendedMaster) {
          return;
        }
        this.trashSummaries(window, group.items.filter((item) => item.id !== recommendedMaster.id), "同题目组非保留项");
      });
      actionBar.appendChild(trashOthersButton);

      const tableWrap = this.createHTMLElement(document, "div");
      tableWrap.setAttribute("class", "dmm-table-wrap");
      section.appendChild(tableWrap);

      const table = this.createHTMLElement(document, "table");
      table.setAttribute("class", "dmm-table dmm-exact-table");
      tableWrap.appendChild(table);

      const thead = this.createHTMLElement(document, "thead");
      table.appendChild(thead);
      const headerRow = this.createHTMLElement(document, "tr");
      thead.appendChild(headerRow);
      for (const label of ["条目", "DOI / 年份", "资料", "操作"]) {
        const th = this.createHTMLElement(document, "th");
        th.textContent = label;
        headerRow.appendChild(th);
      }

      const tbody = this.createHTMLElement(document, "tbody");
      table.appendChild(tbody);
      for (const item of group.items) {
        const row = this.createHTMLElement(document, "tr");
        row.setAttribute("class", "dmm-candidate-row");
        tbody.appendChild(row);

        const titleCell = this.createHTMLElement(document, "td");
        titleCell.setAttribute("class", "dmm-candidate-title");
        titleCell.textContent = item.title || "(无标题)";
        row.appendChild(titleCell);

        const metaCell = this.createHTMLElement(document, "td");
        metaCell.textContent = [
          item.doiKey ? `DOI: ${item.doiKey}` : "DOI: 无",
          item.year ? `年份: ${item.year}` : "年份: 无"
        ].join("  |  ");
        row.appendChild(metaCell);

        const countsCell = this.createHTMLElement(document, "td");
        countsCell.textContent = `附件 ${item.attachmentCount} / 笔记 ${item.noteCount} / 标签 ${item.tagCount}`;
        row.appendChild(countsCell);

        const actionsCell = this.createHTMLElement(document, "td");
        actionsCell.setAttribute("class", "dmm-actions");
        const revealButton = this.createHTMLElement(document, "button");
        revealButton.setAttribute("type", "button");
        revealButton.textContent = "跳转";
        revealButton.addEventListener("click", () => this.revealItem(window, item));
        actionsCell.appendChild(revealButton);

        const mergeAsMasterButton = this.createHTMLElement(document, "button");
        mergeAsMasterButton.setAttribute("type", "button");
        mergeAsMasterButton.setAttribute("class", "dmm-primary-action");
        mergeAsMasterButton.textContent = "保留并合并";
        mergeAsMasterButton.addEventListener("click", () => {
          this.mergeSummaries(window, item, group.items.filter((otherItem) => otherItem.id !== item.id), "同题目组");
        });
        actionsCell.appendChild(mergeAsMasterButton);

        const trashButton = this.createHTMLElement(document, "button");
        trashButton.setAttribute("type", "button");
        trashButton.setAttribute("class", "dmm-danger-action");
        trashButton.textContent = this.config.trashActionLabel;
        trashButton.addEventListener("click", () => this.trashSummaries(window, [item], "单个同题目条目"));
        actionsCell.appendChild(trashButton);

        row.appendChild(actionsCell);
      }
    }
  },

  appendTargetSummary(document, box, target) {
    const section = this.createHTMLElement(document, "div");
    section.setAttribute("class", "dmm-target");
    box.appendChild(section);

    const title = this.createHTMLElement(document, "div");
    title.setAttribute("class", "dmm-target-title");
    title.textContent = `查询条目：${target.title || "(无标题)"}`;
    section.appendChild(title);

    const meta = this.createHTMLElement(document, "div");
    meta.setAttribute("class", "dmm-muted");
    meta.textContent = [
      target.doiKey ? `DOI: ${target.doiKey}` : "DOI: 无",
      target.year ? `年份: ${target.year}` : "年份: 无",
      target.creators.length ? `作者: ${target.creators.slice(0, 3).join(", ")}` : "作者: 无",
      `附件: ${target.attachmentCount}`,
      `笔记: ${target.noteCount}`
    ].join("  |  ");
    section.appendChild(meta);
  },

  appendCandidateTable(window, document, box, data) {
    const summary = this.createHTMLElement(document, "div");
    summary.setAttribute("class", "dmm-summary");
    summary.textContent = data.candidates.length
      ? `找到 ${data.candidates.length} 个候选重复条目`
      : "未找到达到阈值的候选重复条目";
    box.appendChild(summary);

    const tableWrap = this.createHTMLElement(document, "div");
    tableWrap.setAttribute("class", "dmm-table-wrap");
    box.appendChild(tableWrap);

    const table = this.createHTMLElement(document, "table");
    table.setAttribute("class", "dmm-table");
    tableWrap.appendChild(table);

    const thead = this.createHTMLElement(document, "thead");
    table.appendChild(thead);
    const headerRow = this.createHTMLElement(document, "tr");
    thead.appendChild(headerRow);
    for (const label of ["分数", "候选条目", "匹配原因", "所在文件夹", "资料", "操作"]) {
      const th = this.createHTMLElement(document, "th");
      th.textContent = label;
      headerRow.appendChild(th);
    }

    const tbody = this.createHTMLElement(document, "tbody");
    table.appendChild(tbody);

    for (const candidate of data.candidates) {
      const row = this.createHTMLElement(document, "tr");
      row.setAttribute("class", "dmm-candidate-row");
      tbody.appendChild(row);

      const scoreCell = this.createHTMLElement(document, "td");
      scoreCell.setAttribute("class", "dmm-candidate-score");
      scoreCell.textContent = String(candidate.score);
      row.appendChild(scoreCell);

      const itemCell = this.createHTMLElement(document, "td");
      const itemTitle = this.createHTMLElement(document, "div");
      itemTitle.setAttribute("class", "dmm-candidate-title");
      itemTitle.textContent = candidate.item.title || "(无标题)";
      itemCell.appendChild(itemTitle);
      const itemMeta = this.createHTMLElement(document, "div");
      itemMeta.setAttribute("class", "dmm-muted");
      itemMeta.textContent = [
        candidate.item.doiKey ? `DOI: ${candidate.item.doiKey}` : "DOI: 无",
        candidate.item.year ? `年份: ${candidate.item.year}` : "年份: 无"
      ].join("  |  ");
      itemCell.appendChild(itemMeta);
      row.appendChild(itemCell);

      const reasonsCell = this.createHTMLElement(document, "td");
      reasonsCell.setAttribute("class", "dmm-candidate-reasons");
      reasonsCell.textContent = candidate.reasons.join("；") || "相似";
      row.appendChild(reasonsCell);

      const collectionsCell = this.createHTMLElement(document, "td");
      collectionsCell.textContent = candidate.item.collections.length
        ? candidate.item.collections.map((location) => location.path).join("；")
        : this.config.noCollectionsLabel;
      row.appendChild(collectionsCell);

      const countsCell = this.createHTMLElement(document, "td");
      countsCell.textContent = `附件 ${candidate.item.attachmentCount} / 笔记 ${candidate.item.noteCount} / 标签 ${candidate.item.tagCount}`;
      row.appendChild(countsCell);

      const actionsCell = this.createHTMLElement(document, "td");
      actionsCell.setAttribute("class", "dmm-actions");
      const revealButton = this.createHTMLElement(document, "button");
      revealButton.setAttribute("type", "button");
      revealButton.textContent = "跳转";
      revealButton.addEventListener("click", () => this.revealItem(window, candidate.item));
      actionsCell.appendChild(revealButton);

      const mergeButton = this.createHTMLElement(document, "button");
      mergeButton.setAttribute("type", "button");
      mergeButton.setAttribute("class", "dmm-primary-action");
      mergeButton.textContent = "合并到查询";
      mergeButton.addEventListener("click", () => this.mergeSummaries(window, data.target, [candidate.item], "候选重复条目"));
      actionsCell.appendChild(mergeButton);

      const trashButton = this.createHTMLElement(document, "button");
      trashButton.setAttribute("type", "button");
      trashButton.setAttribute("class", "dmm-danger-action");
      trashButton.textContent = this.config.trashActionLabel;
      trashButton.addEventListener("click", () => this.trashSummaries(window, [candidate.item], "候选重复条目"));
      actionsCell.appendChild(trashButton);

      row.appendChild(actionsCell);
    }
  },

  async revealItem(window, itemSummary) {
    try {
      if (itemSummary.collections.length) {
        const collectionID = itemSummary.collections[0].collectionID;
        await window.ZoteroPane.collectionsView.selectCollection(collectionID);
        if (window.ZoteroPane.itemsView && typeof window.ZoteroPane.itemsView.waitForLoad === "function") {
          await window.ZoteroPane.itemsView.waitForLoad();
        }
        await window.ZoteroPane.selectItem(itemSummary.id);
      } else {
        await window.ZoteroPane.selectItem(itemSummary.id, { inLibraryRoot: true });
      }
    } catch (error) {
      Zotero.logError(error);
      Zotero.alert(window, "无法跳转", error && error.message ? error.message : String(error));
    }
  },

  chooseBestMasterSummary(items) {
    const summaries = (items || []).filter(Boolean);
    if (!summaries.length) {
      return null;
    }
    return summaries.slice().sort((left, right) => {
      const rightScore = this.getMasterQualityScore(right);
      const leftScore = this.getMasterQualityScore(left);
      if (rightScore !== leftScore) {
        return rightScore - leftScore;
      }
      if ((left.dateAdded || "") !== (right.dateAdded || "")) {
        return String(left.dateAdded || "").localeCompare(String(right.dateAdded || ""));
      }
      return left.id - right.id;
    })[0];
  },

  getMasterQualityScore(summary) {
    return (summary.attachmentCount * 10)
      + (summary.noteCount * 6)
      + (summary.tagCount * 2)
      + (summary.collections.length * 3)
      + (summary.doiKey ? 4 : 0)
      + (summary.year ? 1 : 0);
  },

  getLiveEditableTopLevelItem(summary) {
    if (!summary || !summary.id) {
      return null;
    }
    const item = Zotero.Items.get(summary.id);
    if (!item || item.deleted) {
      return null;
    }
    if (!item.isRegularItem() || !item.isTopLevelItem()) {
      return null;
    }
    if (typeof item.isEditable === "function" && !item.isEditable()) {
      return null;
    }
    return item;
  },

  async mergeSummaries(window, masterSummary, otherSummaries, contextLabel) {
    try {
      if (!Zotero.Items || typeof Zotero.Items.merge !== "function") {
        Zotero.alert(window, "无法合并", "当前 Zotero 环境未提供 Zotero.Items.merge()。");
        return;
      }

      const masterItem = this.getLiveEditableTopLevelItem(masterSummary);
      const otherItems = (otherSummaries || [])
        .map((summary) => this.getLiveEditableTopLevelItem(summary))
        .filter(Boolean)
        .filter((item) => item.id !== (masterItem && masterItem.id));

      if (!masterItem || !otherItems.length) {
        Zotero.alert(window, "无法合并", "没有找到可合并的可编辑顶层条目。");
        return;
      }

      const validationError = this.validateMergeItems(masterItem, otherItems);
      if (validationError) {
        Zotero.alert(window, "无法合并", validationError);
        return;
      }

      const message = [
        `将 ${otherItems.length} 个${contextLabel || "候选"}合并到保留条目：`,
        "",
        masterItem.getField("title") || "(无标题)",
        "",
        "合并会把重复条目的附件、笔记、标签、文件夹位置等转移到保留条目，并把被合并条目移入 Zotero 回收站。",
        "是否继续？"
      ].join("\n");

      if (!this.confirm(window, "确认合并条目", message)) {
        return;
      }

      if (Zotero.CollectionTreeCache && typeof Zotero.CollectionTreeCache.clear === "function") {
        Zotero.CollectionTreeCache.clear();
      }
      await Zotero.Items.merge(masterItem, otherItems);
      await this.refreshDuplicatePanelAfterChange(window, masterItem.id);
    } catch (error) {
      Zotero.logError(error);
      Zotero.alert(window, "合并失败", error && error.message ? error.message : String(error));
    }
  },

  validateMergeItems(masterItem, otherItems) {
    for (const item of otherItems) {
      if (item.libraryID !== masterItem.libraryID) {
        return "只能合并同一个文库中的条目。";
      }
      if (item.itemTypeID !== masterItem.itemTypeID) {
        return "为避免字段损失，当前版本只允许合并相同条目类型的文献。";
      }
      if (typeof item.isEditable === "function" && !item.isEditable()) {
        return "候选条目不可编辑，无法合并。";
      }
    }
    return "";
  },

  async trashSummaries(window, summaries, contextLabel) {
    try {
      if (!Zotero.Items || typeof Zotero.Items.trashTx !== "function") {
        Zotero.alert(window, "无法回收", "当前 Zotero 环境未提供 Zotero.Items.trashTx()。");
        return;
      }

      const items = (summaries || [])
        .map((summary) => this.getLiveEditableTopLevelItem(summary))
        .filter(Boolean);
      const ids = [...new Set(items.map((item) => item.id))];

      if (!ids.length) {
        Zotero.alert(window, "无法回收", "没有找到可移入回收站的可编辑顶层条目。");
        return;
      }

      const message = [
        `将 ${ids.length} 个${contextLabel || "条目"}移入 Zotero 回收站。`,
        "",
        "这不会永久删除文件，可在 Zotero 回收站中恢复或清空。",
        "是否继续？"
      ].join("\n");

      if (!this.confirm(window, "确认移入回收站", message)) {
        return;
      }

      await Zotero.Items.trashTx(ids);
      await this.refreshDuplicatePanelAfterChange(window);
    } catch (error) {
      Zotero.logError(error);
      Zotero.alert(window, "回收失败", error && error.message ? error.message : String(error));
    }
  },

  confirm(window, title, message) {
    if (typeof Services !== "undefined" && Services.prompt && typeof Services.prompt.confirm === "function") {
      return Services.prompt.confirm(window, title, message);
    }
    return window.confirm(`${title}\n\n${message}`);
  },

  async refreshDuplicatePanelAfterChange(window, preferredItemID) {
    const state = this.windowStates.get(window);
    if (!state || !state.panel || !state.panel.overlay) {
      return;
    }

    const liveIDs = [];
    const seen = new Set();
    for (const id of state.panel.itemIDs || []) {
      const item = Zotero.Items.get(id);
      if (!item || item.deleted || seen.has(id)) {
        continue;
      }
      seen.add(id);
      liveIDs.push(id);
    }
    if (preferredItemID && !seen.has(preferredItemID)) {
      const item = Zotero.Items.get(preferredItemID);
      if (item && !item.deleted) {
        liveIDs.unshift(preferredItemID);
      }
    }

    if (!liveIDs.length) {
      this.closeDuplicatePanel(window);
      return;
    }

    state.panel.itemIDs = liveIDs;
    state.panel.statusText = "正在刷新结果...";
    this.renderDuplicatePanel(window, state.panel, null);

    const data = await this.buildDuplicatePanelData(liveIDs);
    state.panel.statusText = "";
    this.renderDuplicatePanel(window, state.panel, data);
  },

  appendPanelStyles(document, parent) {
    const style = this.createHTMLElement(document, "style");
    style.textContent = `
      .dmm-overlay {
        position: fixed;
        inset: 0;
        z-index: 2147483647;
        display: flex;
        align-items: center;
        justify-content: center;
        background: rgba(17, 24, 39, 0.42);
        font: menu;
      }
      .dmm-panel {
        width: min(1080px, calc(100vw - 64px));
        height: min(760px, calc(100vh - 64px));
        max-height: calc(100vh - 64px);
        display: flex;
        flex-direction: column;
        gap: 10px;
        box-sizing: border-box;
        padding: 16px;
        border: 1px solid rgba(0, 0, 0, 0.24);
        border-radius: 8px;
        background: -moz-dialog;
        color: -moz-dialogtext;
        box-shadow: 0 18px 50px rgba(0, 0, 0, 0.35);
        overflow: hidden;
      }
      .dmm-header {
        flex: 0 0 auto;
        display: flex;
        align-items: flex-start;
        justify-content: space-between;
        gap: 12px;
      }
      .dmm-content {
        flex: 1 1 auto;
        min-height: 0;
        overflow-x: hidden;
        overflow-y: scroll;
        display: flex;
        flex-direction: column;
        gap: 10px;
        padding-right: 10px;
        scrollbar-width: auto;
      }
      .dmm-title {
        font-size: 18px;
        font-weight: 600;
      }
      .dmm-muted,
      .dmm-status {
        color: GrayText;
      }
      .dmm-target {
        flex: 0 0 auto;
        padding: 10px;
        border: 1px solid rgba(0, 0, 0, 0.16);
        border-radius: 6px;
        background: Field;
        color: FieldText;
      }
      .dmm-exact-group {
        border-color: #c1121f;
      }
      .dmm-target-title {
        font-weight: 600;
        margin-bottom: 4px;
      }
      .dmm-selected-duplicates-heading {
        color: #c1121f;
      }
      .dmm-summary {
        flex: 0 0 auto;
        font-weight: 600;
      }
      .dmm-table-wrap {
        flex: 0 0 auto;
        max-height: none;
        overflow-x: auto;
        overflow-y: visible;
        border: 1px solid rgba(0, 0, 0, 0.18);
        border-radius: 6px;
        background: Field;
      }
      .dmm-table {
        width: 100%;
        border-collapse: collapse;
        color: FieldText;
      }
      .dmm-table th,
      .dmm-table td {
        padding: 7px 8px;
        border-bottom: 1px solid rgba(0, 0, 0, 0.12);
        text-align: left;
        vertical-align: top;
      }
      .dmm-table th:first-child,
      .dmm-table td:first-child {
        width: 56px;
        text-align: right;
        font-variant-numeric: tabular-nums;
      }
      .dmm-exact-table th:first-child,
      .dmm-exact-table td:first-child {
        width: auto;
        text-align: left;
        font-variant-numeric: normal;
      }
      .dmm-candidate-row {
        border-left: 4px solid #c1121f;
      }
      .dmm-candidate-score,
      .dmm-candidate-title,
      .dmm-candidate-reasons {
        color: #c1121f;
        font-weight: 600;
      }
      .dmm-candidate-row .dmm-muted {
        color: #9f1d2a;
      }
      .dmm-actions {
        display: flex;
        flex-wrap: wrap;
        gap: 6px;
        align-items: center;
      }
      .dmm-group-actions {
        margin: 8px 0;
      }
      .dmm-actions button {
        color: ButtonText;
      }
      .dmm-primary-action {
        font-weight: 600;
      }
      .dmm-danger-action {
        color: #9f1d2a !important;
        font-weight: 600;
      }
      .dmm-status {
        padding: 20px;
      }
    `;
    parent.appendChild(style);
  }
};
